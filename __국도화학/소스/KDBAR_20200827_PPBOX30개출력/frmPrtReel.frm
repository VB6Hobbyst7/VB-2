VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmPrtReel 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Reel 라벨 출력"
   ClientHeight    =   10635
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20790
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10635
   ScaleWidth      =   20790
   WindowState     =   2  '최대화
   Begin VB.TextBox txtComm 
      Appearance      =   0  '평면
      Height          =   5955
      Left            =   22110
      MultiLine       =   -1  'True
      TabIndex        =   61
      Top             =   2430
      Visible         =   0   'False
      Width           =   5475
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   10995
      Left            =   90
      TabIndex        =   12
      Top             =   1050
      Width           =   21735
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Height          =   9615
         Left            =   8160
         TabIndex        =   13
         Top             =   210
         Width           =   13305
         Begin VB.TextBox txtOrderMemo 
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
            Locked          =   -1  'True
            TabIndex        =   72
            Top             =   6540
            Width           =   5400
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
            Left            =   9750
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   70
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
            Left            =   9750
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   67
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
            TabIndex        =   62
            Top             =   360
            Width           =   1500
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00FFFFFF&
            Height          =   4455
            Left            =   7230
            ScaleHeight     =   4395
            ScaleWidth      =   5775
            TabIndex        =   27
            Top             =   1440
            Width           =   5835
            Begin VB.Image imgBar2 
               Height          =   465
               Left            =   330
               Picture         =   "frmPrtReel.frx":0000
               Stretch         =   -1  'True
               Top             =   2340
               Visible         =   0   'False
               Width           =   2685
            End
            Begin VB.Image imgQrBar 
               Height          =   750
               Left            =   660
               Picture         =   "frmPrtReel.frx":3FCD
               Stretch         =   -1  'True
               Top             =   2910
               Visible         =   0   'False
               Width           =   840
            End
            Begin VB.Image imgBar1 
               Height          =   465
               Left            =   330
               Picture         =   "frmPrtReel.frx":A605
               Stretch         =   -1  'True
               Top             =   1860
               Visible         =   0   'False
               Width           =   2685
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
               Index           =   14
               Left            =   870
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
               Index           =   13
               Left            =   690
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
               Index           =   12
               Left            =   390
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
               Index           =   11
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
               Index           =   10
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
               Index           =   9
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
               Index           =   8
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
               Index           =   7
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
               Index           =   6
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
               Index           =   5
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
               Index           =   4
               Left            =   0
               TabIndex        =   48
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
               TabIndex        =   47
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
               TabIndex        =   46
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
               Index           =   15
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
               Index           =   14
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
               Index           =   13
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
               Index           =   12
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
               Index           =   11
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
               Index           =   10
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
               Index           =   9
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
               Index           =   8
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
               Index           =   7
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
               Index           =   6
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
               Index           =   5
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
               Index           =   4
               Left            =   0
               TabIndex        =   33
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
               Index           =   2
               Left            =   0
               TabIndex        =   31
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
               TabIndex        =   30
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
               TabIndex        =   29
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
               TabIndex        =   28
               Top             =   0
               Visible         =   0   'False
               Width           =   570
            End
         End
         Begin VB.TextBox txtReelQTY 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00C0FFFF&
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
            Top             =   6090
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
            Left            =   8190
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   24
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
            TabIndex        =   23
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
            TabIndex        =   22
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
            TabIndex        =   21
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
            TabIndex        =   15
            Top             =   360
            Width           =   1485
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  '없음
            Height          =   735
            Left            =   7440
            TabIndex        =   14
            Top             =   7890
            Width           =   5805
            Begin VB.CommandButton cmdSamplePrint 
               Appearance      =   0  '평면
               BackColor       =   &H00E0E0E0&
               Caption         =   "샘플출력"
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
               Left            =   1140
               Style           =   1  '그래픽
               TabIndex        =   73
               Top             =   180
               Width           =   1095
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
               Left            =   4470
               Style           =   1  '그래픽
               TabIndex        =   66
               Top             =   180
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.CommandButton cmdAllPrint 
               Appearance      =   0  '평면
               BackColor       =   &H00E0E0E0&
               Caption         =   "일괄출력"
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
               Left            =   2250
               Style           =   1  '그래픽
               TabIndex        =   65
               Top             =   180
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
               Left            =   30
               Style           =   1  '그래픽
               TabIndex        =   8
               Top             =   180
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
               Left            =   3360
               Style           =   1  '그래픽
               TabIndex        =   9
               Top             =   180
               Width           =   1095
            End
         End
         Begin FPSpread.vaSpread spdPrtReelDetail 
            Height          =   4455
            Left            =   240
            TabIndex        =   6
            Top             =   1440
            Width           =   6915
            _Version        =   393216
            _ExtentX        =   12197
            _ExtentY        =   7858
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
            SpreadDesigner  =   "frmPrtReel.frx":E5D2
            ScrollBarTrack  =   3
            ShowScrollTips  =   3
         End
         Begin FPSpread.vaSpread spdRegOrderDetail 
            Height          =   2505
            Left            =   240
            TabIndex        =   64
            Top             =   6990
            Width           =   6945
            _Version        =   393216
            _ExtentX        =   12250
            _ExtentY        =   4419
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
            SpreadDesigner  =   "frmPrtReel.frx":F110
            ScrollBarTrack  =   3
            ShowScrollTips  =   3
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "전달 메모"
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
            TabIndex        =   71
            Top             =   6540
            Width           =   1500
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
            TabIndex        =   63
            Top             =   360
            Width           =   1800
         End
         Begin VB.Label Label3 
            Alignment       =   1  '오른쪽 맞춤
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
            Height          =   285
            Left            =   6690
            TabIndex        =   26
            Top             =   6180
            Width           =   495
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
            TabIndex        =   25
            Top             =   6090
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
            TabIndex        =   20
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
            TabIndex        =   19
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
            TabIndex        =   18
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
            TabIndex        =   17
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
            TabIndex        =   16
            Top             =   360
            Width           =   1500
         End
      End
      Begin FPSpread.vaSpread spdPrtReel 
         Height          =   9525
         Left            =   210
         TabIndex        =   5
         Top             =   300
         Width           =   7905
         _Version        =   393216
         _ExtentX        =   13944
         _ExtentY        =   16801
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
         RetainSelBlock  =   0   'False
         ScrollBarExtMode=   -1  'True
         ScrollBars      =   2
         ShadowColor     =   16774120
         SpreadDesigner  =   "frmPrtReel.frx":10091
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
      Width           =   21765
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
         TabIndex        =   68
         Top             =   360
         Width           =   1185
      End
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
         Format          =   128974849
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
         Format          =   128974849
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
               Picture         =   "frmPrtReel.frx":1134D
               Key             =   "RUN"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrtReel.frx":118E7
               Key             =   "NOT"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrtReel.frx":11E81
               Key             =   "STOP"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrtReel.frx":1241B
               Key             =   "LST"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrtReel.frx":12CAD
               Key             =   "ITM"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrtReel.frx":12E07
               Key             =   "ERR"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrtReel.frx":12F61
               Key             =   "NOF"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrtReel.frx":130BB
               Key             =   "ON"
               Object.Tag             =   "OFF"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPrtReel.frx":13995
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
         Left            =   16770
         TabIndex        =   69
         Top             =   360
         Width           =   2265
      End
      Begin VB.Image imgPort 
         Height          =   240
         Left            =   11340
         Picture         =   "frmPrtReel.frx":1426F
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
         TabIndex        =   60
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
         TabIndex        =   11
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
         TabIndex        =   10
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
Dim gAllPrt As Boolean

Private Sub cmdAllPrint_Click()
    Dim i As Integer
    
    If MsgBox("일괄출력을 하시겠습니까?", vbYesNo + vbCritical + vbDefaultButton1, Me.Caption) = vbYes Then
        gAllPrt = True
        For i = 1 To spdPrtReel.MaxRows
            If GetText(spdPrtReel, i, 1) = "1" Then
                Call spdPrtReel_Click(2, i)
                DoEvents
                Call cmdPrint_Click
                DoEvents
                Call SetText(spdPrtReel, "", i, 1)
            End If
        Next
        gAllPrt = False
    End If
    
End Sub

Private Sub cmdClear_Click()
    Dim i   As Integer
    
    spdPrtReel.MaxRows = 0
    spdPrtReelDetail.MaxRows = 0
    spdRegOrderDetail.MaxRows = 0
    
    dtpFromDate.Value = Now - 1
    dtpToDate.Value = Now

    txtProdNm.Text = ""
    txtProdOrderDt.Text = ""
    'txtProdPosNo.Text = ""
    txtPackNm.Text = ""
    txtCompNm.Text = ""
    txtSlittingNo.Text = ""
    txtReelQTY.Text = ""
    txtOrderMemo.Text = ""
    
    lblstrPrtLabelName.Caption = ""
    
    For i = 0 To 15
        'barReel.Visible = False
        'barPart.Visible = False
        lblTitle(i).Visible = False
    Next
    
    imgBar1.Visible = False
    imgBar2.Visible = False
    imgQrBar.Visible = False
    
    
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
Private Sub GetReelOrderList(ByVal pOrderFromDate As String, ByVal pOrderToDate As String, Optional ByVal pProdCd As String, Optional ByVal pOrderNo As String, Optional ByVal pLabelType As String, Optional ByVal pYN As String)

    Dim strLabelType    As String

    Set AdoRs = Get_OrderList(pOrderFromDate, pOrderToDate, pProdCd, pOrderNo, pLabelType, pYN)

    If AdoRs Is Nothing Then
        '등록된 정보 없음
    Else
        Do Until AdoRs.EOF
            With spdPrtReel
                .MaxRows = .MaxRows + 1

                Call SetText(spdPrtReel, "1", .MaxRows, 1)
                Call SetText(spdPrtReel, AdoRs.Fields("LOT_NO").Value & "", .MaxRows, 2)
                Call SetText(spdPrtReel, Format(AdoRs.Fields("PROD_ORDER_DT").Value & "", "####-##-##"), .MaxRows, 3)
                'Call SetText(spdPrtReel, AdoRs.Fields("PROD_POS_NO").Value & "", .MaxRows, 4)
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
    Dim strPrtData()   As String
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
    Dim strPlusXPos200DPI As String
    Dim strAFont    As String
    Dim strAFont200DPI    As String
    Dim i           As Integer
    Dim J           As Integer
    Dim k           As Integer
    Dim blnPrint    As Boolean
    
    '-- 트래킹용
    Dim intCnt          As Integer
    Dim strTrackBC()    As Variant
    '-- 재출력용
    Dim strPrintVal()   As Variant
    '-- 출력용
    Dim intPrtCnt   As Integer
    Dim strPrint()  As Variant
    
    Dim intPrintOddCnt      As Integer
    Dim strPrintOddVal()    As Variant   '홀수
    Dim intPrintEvenCnt     As Integer
    Dim strPrintEvenVal()   As Variant   '짝수
    
    Dim intPrint1Cnt        As Integer
    Dim strPrint1()         As Variant
    Dim intPrint2Cnt        As Integer
    Dim strPrint2()         As Variant
    Dim intPrint3Cnt        As Integer
    Dim strPrint3()         As Variant
    Dim intPrint4Cnt        As Integer
    Dim strPrint4()         As Variant
    
    Dim intPrint1MaxCnt        As Integer
    Dim intPrint2MaxCnt        As Integer
    Dim intPrint3MaxCnt        As Integer
    Dim intPrint4MaxCnt        As Integer
    
    Dim intP1Cnt        As Integer
    Dim intP2Cnt        As Integer
    Dim intP3Cnt        As Integer
    Dim intP4Cnt        As Integer
    
'    Dim intPrint1MaxCntB        As Integer
'    Dim intPrint2MaxCntB        As Integer
'    Dim intPrint3MaxCntB        As Integer
'    Dim intPrint4MaxCntB        As Integer
    
    Dim intCount    As Integer
    
    '-- 첫번째 P번호가 홀수인지 짝수인지 갖고 있는다.
    ' 홀수일 경우 blnOdd = True
'    Dim blnOdd      As Boolean
    
    intPrintOddCnt = 0
    intPrintEvenCnt = 0
    Erase strPrintOddVal
    Erase strPrintEvenVal
    
    intPrint1Cnt = 0
    intPrint2Cnt = 0
    intPrint3Cnt = 0
    intPrint4Cnt = 0
    
    intPrint1MaxCnt = 0
    intPrint2MaxCnt = 0
    intPrint3MaxCnt = 0
    intPrint4MaxCnt = 0

    intP1Cnt = 0
    intP2Cnt = 0
    intP3Cnt = 0
    intP4Cnt = 0

    Erase strPrint1
    Erase strPrint2
    Erase strPrint3
    Erase strPrint4
    
    
    blnPrint = False
    strBarcode = ""
    Erase strPrtData
    Erase strTrackBC
    Erase strPrintVal
    intCnt = 0
    intPrt = 0
    intPrtCnt = 0
'    blnOdd = False
    
    If txtProdNm.Text = "" Then
        MsgBox "출력할 Reel라벨을 선택하세요", vbOKOnly + vbCritical, Me.Caption
        spdPrtReel.SetFocus
        Exit Sub
    End If
    
    strPlusXPos = 630
    strPlusXPos200DPI = 430
    strAFont = "^A0N,35,30"
    strAFont200DPI = "^A0N,25,15"
    
    strHeader = "^XA" & vbLf
    strHeader = strHeader & "^SEE:UHANGUL.DAT^FS" & vbLf
    strHeader = strHeader & "^PON^FS" & vbLf
    strHeader = strHeader & "^CWJ,E:KFONT3.FNT^FS" & vbLf
    strHeader = strHeader & "^CI26"
'    strHeader = strHeader & "^PR12,12"  'speed
'    strHeader = strHeader & "^MD7"      'darkness
    
    Select Case lblstrPrtLabelName.Caption
        Case "P0011", "P0018", "P0019", "P0020"
            strPlusXPos = 500
            i = 0
            For intPRow = 1 To spdRegOrderDetail.MaxRows
                strPrtFNo = GetText(spdRegOrderDetail, intPRow, 7)
                strPrtTNo = GetText(spdRegOrderDetail, intPRow, 8)
                intPrtFNo = CCur(Mid(strPrtFNo, 3))
                intPrtTNo = CCur(Mid(strPrtTNo, 3))
                
                For intPrt = intPrtFNo To intPrtTNo
                    strOutput = ""
                    Erase strPrtData
                    'i = 0
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
                                strLotSub = "P" & Mid(strPrtFNo, 2, 1) & Format(intPrt, "00")
                                strData = Mid(strData, 1, 9) & strLotSub
                                '좌측 출력
                                ReDim Preserve strPrtData(i) As String
                                strPrtData(i) = ""
                                strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                                '-- 테스트 해야 함
                               ' strPrtData(i) = strPrtData(i) & "^BY" & "2,4,50"
                                strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
                                strPrtData(i) = strPrtData(i) & "^FD" & strData
                                strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                i = i + 1
                                '우측 출력
                                ReDim Preserve strPrtData(i) As String
                                strPrtData(i) = ""
                                strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) & "," & strYPos
                                '-- 테스트 해야 함
                                'strPrtData(i) = strPrtData(i) & "^BY" & "2,4,50"
                                strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
                                strPrtData(i) = strPrtData(i) & "^FD" & strData
                                strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                i = i + 1
                                strBarcode = strData
                            Else
                                If strType = "ACF Lot" Or strType = "LotNo" Or strType = "Lot No." Then
                                    strLot = mGetP(strData, 1, "(")
                                    strLot = strLot & "(" & strLotSub & ")"
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & strAFont
                                    strPrtData(i) = strPrtData(i) & "^FH"
                                    If strNamePrt = "Y" Then
                                        strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strLot
                                    Else
                                        strPrtData(i) = strPrtData(i) & "^FD" & strLot
                                    End If
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & strAFont
                                    strPrtData(i) = strPrtData(i) & "^FH"
                                    If strNamePrt = "Y" Then
                                        strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strLot
                                    Else
                                        strPrtData(i) = strPrtData(i) & "^FD" & strLot
                                    End If
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                ElseIf strType = "Storage" Then
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & strAFont
                                    strPrtData(i) = strPrtData(i) & "^FH"
                                    If strNamePrt = "Y" Then
                                        strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                                    Else
                                        strPrtData(i) = strPrtData(i) & "^FD" & strData
                                    End If
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos + 240 & "," & CCur(strYPos) - 5 & "^A0N,20,20^FDo^FS" & vbLf
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos + 250 & "," & strYPos & strAFont & "^FDC^FS" & vbLf
                                    i = i + 1
                                    
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & strAFont
                                    strPrtData(i) = strPrtData(i) & "^FH"
                                    If strNamePrt = "Y" Then
                                        strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                                    Else
                                        strPrtData(i) = strPrtData(i) & "^FD" & strData
                                    End If
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) + 240 & "," & CCur(strYPos) - 5 & "^A0N,20,20^FDo^FS" & vbLf
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) + 250 & "," & strYPos & strAFont & "^FDC^FS" & vbLf
                                    i = i + 1
                                Else
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & strAFont
                                    strPrtData(i) = strPrtData(i) & "^FH"
                                    If strNamePrt = "Y" Then
                                        strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                                    Else
                                        strPrtData(i) = strPrtData(i) & "^FD" & strData
                                    End If
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) & "," & strYPos
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
                    
                    '특수문자변환
                    strOutput = Replace(strOutput, "~", "_7E")  'asc("~") =  126
                    strOutput = Replace(strOutput, "℃", "") '°C
                    strOutput = strHeader & strOutput & "^XZ" & vbLf
                    If intPrtTNo > 30 Then
                        Select Case intPrt
                            Case 1, 5, 9, 13, 17, 21, 25, 29, 33, 37, 41, 45, 49, 53, 57, 61, 65, 69, 73, 77, 81
                                ReDim Preserve strPrint1(intPrint1Cnt) As Variant
                                strPrint1(intPrint1Cnt) = strOutput
                                intPrint1Cnt = intPrint1Cnt + 1
                            Case 2, 6, 10, 14, 18, 22, 26, 30, 34, 38, 42, 46, 50, 54, 58, 62, 66, 70, 74, 78, 82
                                ReDim Preserve strPrint2(intPrint2Cnt) As Variant
                                strPrint2(intPrint2Cnt) = strOutput
                                intPrint2Cnt = intPrint2Cnt + 1
                            Case 3, 7, 11, 15, 19, 23, 27, 31, 35, 39, 43, 47, 51, 55, 59, 63, 67, 71, 75, 79, 83
                                ReDim Preserve strPrint3(intPrint3Cnt) As Variant
                                strPrint3(intPrint3Cnt) = strOutput
                                intPrint3Cnt = intPrint3Cnt + 1
                            Case 4, 8, 12, 16, 20, 24, 28, 32, 36, 40, 44, 48, 52, 56, 60, 64, 68, 72, 76, 80, 84
                                ReDim Preserve strPrint4(intPrint4Cnt) As Variant
                                strPrint4(intPrint4Cnt) = strOutput
                                intPrint4Cnt = intPrint4Cnt + 1
                        End Select
                    Else
                        '짝수
                        If intPrt Mod 2 = 0 Then
                            '속도문제로 배열에 담았다가 한꺼번에  출력한다.
                            ReDim Preserve strPrintEvenVal(intPrintEvenCnt) As Variant
                            strPrintEvenVal(intPrintEvenCnt) = strOutput
                            intPrintEvenCnt = intPrintEvenCnt + 1
                        '홀수
                        Else
                            '속도문제로 배열에 담았다가 한꺼번에  출력한다.
                            ReDim Preserve strPrintOddVal(intPrintOddCnt) As Variant
                            strPrintOddVal(intPrintOddCnt) = strOutput
                            intPrintOddCnt = intPrintOddCnt + 1
                        End If
                    End If
                    '트래킹용
                    ReDim Preserve strTrackBC(intCnt)
                    strTrackBC(intCnt) = strBarcode
                    '재출력용
                    ReDim Preserve strPrintVal(intCnt)
                    strPrintVal(intCnt) = strOutput
                    intCnt = intCnt + 1
                    blnPrint = True
                    txtComm.Text = strOutput
                    strOutput = ""
                Next
                    
                If intPrtTNo > 30 Then
                    If intPRow = 1 Then
                        intP1Cnt = 0
                        intP2Cnt = 0
                        intP3Cnt = 0
                        intP4Cnt = 0
                    Else
                        intP1Cnt = intPrint1MaxCnt
                        intP2Cnt = intPrint2MaxCnt
                        intP3Cnt = intPrint3MaxCnt
                        intP4Cnt = intPrint4MaxCnt
                    End If
                    intPrint1MaxCnt = intPrint1Cnt
                    intPrint2MaxCnt = intPrint2Cnt
                    intPrint3MaxCnt = intPrint3Cnt
                    intPrint4MaxCnt = intPrint4Cnt
                    For i = intPrtFNo To intPrtTNo
                        Select Case i
                            Case 1, 5, 9, 13, 17, 21, 25, 29, 33, 37, 41, 45, 49, 53, 57, 61, 65, 69, 73, 77, 81
                                ReDim Preserve strPrint(intPrtCnt) As Variant
                                strPrint(intPrtCnt) = strPrint1(intP1Cnt)
                                intP1Cnt = intP1Cnt + 1
                                intPrtCnt = intPrtCnt + 1
                        End Select
                    Next
                    For i = intPrtFNo To intPrtTNo
                        Select Case i
                            Case 2, 6, 10, 14, 18, 22, 26, 30, 34, 38, 42, 46, 50, 54, 58, 62, 66, 70, 74, 78, 82
                                ReDim Preserve strPrint(intPrtCnt) As Variant
                                strPrint(intPrtCnt) = strPrint2(intP2Cnt)
                                intP2Cnt = intP2Cnt + 1
                                intPrtCnt = intPrtCnt + 1
                        End Select
                    Next
                    For i = intPrtFNo To intPrtTNo
                        Select Case i
                            Case 3, 7, 11, 15, 19, 23, 27, 31, 35, 39, 43, 47, 51, 55, 59, 63, 67, 71, 75, 79, 83
                                ReDim Preserve strPrint(intPrtCnt) As Variant
                                strPrint(intPrtCnt) = strPrint3(intP3Cnt)
                                intP3Cnt = intP3Cnt + 1
                                intPrtCnt = intPrtCnt + 1
                        End Select
                    Next
                    For i = intPrtFNo To intPrtTNo
                        Select Case i
                            Case 4, 8, 12, 16, 20, 24, 28, 32, 36, 40, 44, 48, 52, 56, 60, 64, 68, 72, 76, 80, 84
                                ReDim Preserve strPrint(intPrtCnt) As Variant
                                strPrint(intPrtCnt) = strPrint4(intP4Cnt)
                                intP4Cnt = intP4Cnt + 1
                                intPrtCnt = intPrtCnt + 1
                        End Select
                    Next
                Else
                    intPrintOddCnt = 0
                    intPrintEvenCnt = 0
                    For i = intPrtFNo To intPrtTNo Step 2
                        ReDim Preserve strPrint(intPrtCnt) As Variant
                        If i Mod 2 = 0 Then
                            strPrint(intPrtCnt) = strPrintEvenVal(intPrintOddCnt)
                            intPrintOddCnt = intPrintOddCnt + 1
                        Else
                            strPrint(intPrtCnt) = strPrintOddVal(intPrintEvenCnt)
                            intPrintEvenCnt = intPrintEvenCnt + 1
                        End If
                        'Debug.Print strPrint(intPrtCnt)
                        intPrtCnt = intPrtCnt + 1
                    Next
                    intPrintOddCnt = 0
                    intPrintEvenCnt = 0
                    For i = intPrtFNo + 1 To intPrtTNo Step 2
                        ReDim Preserve strPrint(intPrtCnt) As Variant
                        If i Mod 2 = 0 Then
                            strPrint(intPrtCnt) = strPrintEvenVal(intPrintOddCnt)
                            intPrintOddCnt = intPrintOddCnt + 1
                        Else
                            strPrint(intPrtCnt) = strPrintOddVal(intPrintEvenCnt)
                            intPrintEvenCnt = intPrintEvenCnt + 1
                        End If
                        intPrtCnt = intPrtCnt + 1
                    Next
                    Erase strPrintOddVal
                    Erase strPrintEvenVal
                    intPrintOddCnt = 0
                    intPrintEvenCnt = 0
                End If
            Next
        Case "P0001", "P0002"
            i = 0
            For intPRow = 1 To spdRegOrderDetail.MaxRows
                strPrtFNo = GetText(spdRegOrderDetail, intPRow, 7)
                strPrtTNo = GetText(spdRegOrderDetail, intPRow, 8)
                intPrtFNo = CCur(Mid(strPrtFNo, 3))
                intPrtTNo = CCur(Mid(strPrtTNo, 3))
                
                For intPrt = intPrtFNo To intPrtTNo
                    strOutput = ""
                    Erase strPrtData
                    'i = 0
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
                            '-- 2020.07.23 이노룩스만 P에서 K로 바뀜
                                'strLotSub = "P" & Mid(strPrtFNo, 2, 1) & Format(intPrt, "00")
                                strLotSub = "K" & Mid(strPrtFNo, 2, 1) & Format(intPrt, "00")
                                
                                strData = Mid(strData, 1, 13) & strLotSub & Mid(strData, 18, 5)
                                
                                If Mid(strBarType, 1, 1) = "1" Then
                                    '좌측 출력
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                                                                        
                                    '-- 테스트 해야 함
                                    'strPrtData(i) = strPrtData(i) & "^BY" & "2,4,50"

                                    strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
                                    strPrtData(i) = strPrtData(i) & "^FD" & strData
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                    '우측 출력
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) & "," & strYPos
                                    
                                    '-- 테스트 해야 함
                                    'strPrtData(i) = strPrtData(i) & "^BY" & "2,4,50"
                                    
                                    strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
                                    strPrtData(i) = strPrtData(i) & "^FD" & strData
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                Else
                                    '좌측 출력
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
'                                    strPrtData(i) = strPrtData(i) & "^CI26"
                                    strPrtData(i) = strPrtData(i) & "^BY" & "2,4,50"
                                    strPrtData(i) = strPrtData(i) & "^BQ"
                                    strPrtData(i) = strPrtData(i) & "^FD" & strData
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                    '우측 출력
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) & "," & strYPos
'                                    strPrtData(i) = strPrtData(i) & "^CI26"
                                    strPrtData(i) = strPrtData(i) & "^BY" & "2,4,50"
                                    strPrtData(i) = strPrtData(i) & "^BQ"
                                    strPrtData(i) = strPrtData(i) & "^FD" & strData
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                End If
                                strBarcode = strData
                            Else
                                'If strType = "LotNo" Then
                                If strType = "ACF Lot" Or strType = "LotNo" Then
                                    strLot = mGetP(strData, 1, "(")
                                    strLot = strLot & "(" & strLotSub & ")"
                                    
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
'                                    strPrtData(i) = strPrtData(i) & "^CI26"
                                    strPrtData(i) = strPrtData(i) & strAFont
                                    strPrtData(i) = strPrtData(i) & "^FH"
                                    If strNamePrt = "Y" Then
                                        strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strLot
                                    Else
                                        strPrtData(i) = strPrtData(i) & "^FD" & strLot
                                    End If
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & strAFont
                                    strPrtData(i) = strPrtData(i) & "^FH"
                                    If strNamePrt = "Y" Then
                                        strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strLot
                                    Else
                                        strPrtData(i) = strPrtData(i) & "^FD" & strLot
                                    End If
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                ElseIf strType = "Storage Temperature" Then
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & strAFont
                                    strPrtData(i) = strPrtData(i) & "^FH"
                                    If strNamePrt = "Y" Then
                                        strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                                    Else
                                        strPrtData(i) = strPrtData(i) & "^FD" & strData
                                    End If
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    'o추가
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos + 400 & "," & CCur(strYPos) - 5 & "^A0N,20,20^FDo^FS" & vbLf
                                    'C추가
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos + 410 & "," & strYPos & strAFont & "^FDC^FS" & vbLf
                                    i = i + 1
                                    
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & strAFont
                                    strPrtData(i) = strPrtData(i) & "^FH"
                                    If strNamePrt = "Y" Then
                                        strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                                    Else
                                        strPrtData(i) = strPrtData(i) & "^FD" & strData
                                    End If
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    'o추가
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) + 400 & "," & CCur(strYPos) - 5 & "^A0N,20,20^FDo^FS" & vbLf
                                    'C추가
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) + 410 & "," & strYPos & strAFont & "^FDC^FS" & vbLf
                                    i = i + 1
                                Else
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
'                                    strPrtData(i) = strPrtData(i) & "^CI26"
                                    strPrtData(i) = strPrtData(i) & strAFont
                                    strPrtData(i) = strPrtData(i) & "^FH"
                                    If strNamePrt = "Y" Then
                                        strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                                    Else
                                        strPrtData(i) = strPrtData(i) & "^FD" & strData
                                    End If
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                    
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) & "," & strYPos
'                                    strPrtData(i) = strPrtData(i) & "^CI26"
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
                    
                    '특수문자변환
                    strOutput = Replace(strOutput, "~", "_7E")  'asc("~") =  126
'                    strOutput = Replace(strOutput, "℃", "'C") '°C
                    strOutput = Replace(strOutput, "℃", "") '°C
                    
                    strOutput = strHeader & strOutput & "^XZ" & vbLf
                
                    '속도문제로 배열에 담았다가 한꺼번에  출력한다.
                    'ReDim Preserve strPrint(intPrtCnt) As Variant
                    'strPrint(intPrtCnt) = strOutput
                    'intPrtCnt = intPrtCnt + 1
                    
                    If intPrtTNo > 30 Then
                        Select Case intPrt
                            Case 1, 5, 9, 13, 17, 21, 25, 29, 33, 37, 41, 45, 49, 53, 57, 61, 65, 69, 73, 77, 81
                                ReDim Preserve strPrint1(intPrint1Cnt) As Variant
                                strPrint1(intPrint1Cnt) = strOutput
                                intPrint1Cnt = intPrint1Cnt + 1
                                
                            Case 2, 6, 10, 14, 18, 22, 26, 30, 34, 38, 42, 46, 50, 54, 58, 62, 66, 70, 74, 78, 82
                                ReDim Preserve strPrint2(intPrint2Cnt) As Variant
                                strPrint2(intPrint2Cnt) = strOutput
                                intPrint2Cnt = intPrint2Cnt + 1
                                
                            Case 3, 7, 11, 15, 19, 23, 27, 31, 35, 39, 43, 47, 51, 55, 59, 63, 67, 71, 75, 79, 83
                                ReDim Preserve strPrint3(intPrint3Cnt) As Variant
                                strPrint3(intPrint3Cnt) = strOutput
                                intPrint3Cnt = intPrint3Cnt + 1
                                
                            Case 4, 8, 12, 16, 20, 24, 28, 32, 36, 40, 44, 48, 52, 56, 60, 64, 68, 72, 76, 80, 84
                                ReDim Preserve strPrint4(intPrint4Cnt) As Variant
                                strPrint4(intPrint4Cnt) = strOutput
                                intPrint4Cnt = intPrint4Cnt + 1
                                
                        End Select
                    Else
                        '짝수
                        If intPrt Mod 2 = 0 Then
                            '속도문제로 배열에 담았다가 한꺼번에  출력한다.
                            ReDim Preserve strPrintEvenVal(intPrintEvenCnt) As Variant
                            strPrintEvenVal(intPrintEvenCnt) = strOutput
                            intPrintEvenCnt = intPrintEvenCnt + 1
                        '홀수
                        Else
                            '속도문제로 배열에 담았다가 한꺼번에  출력한다.
                            ReDim Preserve strPrintOddVal(intPrintOddCnt) As Variant
                            strPrintOddVal(intPrintOddCnt) = strOutput
                            intPrintOddCnt = intPrintOddCnt + 1
                        End If
                    End If
                    
                    '트래킹용
                    ReDim Preserve strTrackBC(intCnt)
                    strTrackBC(intCnt) = strBarcode
                    
                    '재출력용
                    ReDim Preserve strPrintVal(intCnt)
                    strPrintVal(intCnt) = strOutput
                    
                    intCnt = intCnt + 1
                    
                    blnPrint = True
                    txtComm.Text = strOutput
                    strOutput = ""
                Next
                    
                '속도문제로 배열에 담았다가 한꺼번에  출력한다.
                If intPrtTNo > 30 Then
                    If intPRow = 1 Then
                        intP1Cnt = 0
                        intP2Cnt = 0
                        intP3Cnt = 0
                        intP4Cnt = 0
                    Else
                        intP1Cnt = intPrint1MaxCnt
                        intP2Cnt = intPrint2MaxCnt
                        intP3Cnt = intPrint3MaxCnt
                        intP4Cnt = intPrint4MaxCnt
                    End If
                    
                    intPrint1MaxCnt = intPrint1Cnt
                    intPrint2MaxCnt = intPrint2Cnt
                    intPrint3MaxCnt = intPrint3Cnt
                    intPrint4MaxCnt = intPrint4Cnt
                                        
                    For i = intPrtFNo To intPrtTNo
                        Select Case i
                            Case 1, 5, 9, 13, 17, 21, 25, 29, 33, 37, 41, 45, 49, 53, 57, 61, 65, 69, 73, 77, 81
                                ReDim Preserve strPrint(intPrtCnt) As Variant
                                strPrint(intPrtCnt) = strPrint1(intP1Cnt)
                                intP1Cnt = intP1Cnt + 1
                                intPrtCnt = intPrtCnt + 1
                        End Select
                    Next
                    
                    For i = intPrtFNo To intPrtTNo
                        Select Case i
                            Case 2, 6, 10, 14, 18, 22, 26, 30, 34, 38, 42, 46, 50, 54, 58, 62, 66, 70, 74, 78, 82
                                ReDim Preserve strPrint(intPrtCnt) As Variant
                                strPrint(intPrtCnt) = strPrint2(intP2Cnt)
                                intP2Cnt = intP2Cnt + 1
                                intPrtCnt = intPrtCnt + 1
                        End Select
                    Next
                    
                    For i = intPrtFNo To intPrtTNo
                        Select Case i
                            Case 3, 7, 11, 15, 19, 23, 27, 31, 35, 39, 43, 47, 51, 55, 59, 63, 67, 71, 75, 79, 83
                                ReDim Preserve strPrint(intPrtCnt) As Variant
                                strPrint(intPrtCnt) = strPrint3(intP3Cnt)
                                intP3Cnt = intP3Cnt + 1
                                intPrtCnt = intPrtCnt + 1
                        End Select
                    Next
                    
                    For i = intPrtFNo To intPrtTNo
                        Select Case i
                            Case 4, 8, 12, 16, 20, 24, 28, 32, 36, 40, 44, 48, 52, 56, 60, 64, 68, 72, 76, 80, 84
                                ReDim Preserve strPrint(intPrtCnt) As Variant
                                strPrint(intPrtCnt) = strPrint4(intP4Cnt)
                                intP4Cnt = intP4Cnt + 1
                                intPrtCnt = intPrtCnt + 1
                        End Select
                    Next
                    
                Else
                    intPrintOddCnt = 0
                    intPrintEvenCnt = 0
                    For i = intPrtFNo To intPrtTNo Step 2
                        ReDim Preserve strPrint(intPrtCnt) As Variant
                        If i Mod 2 = 0 Then
                            strPrint(intPrtCnt) = strPrintEvenVal(intPrintOddCnt)
                            intPrintOddCnt = intPrintOddCnt + 1
                        Else
                            strPrint(intPrtCnt) = strPrintOddVal(intPrintEvenCnt)
                            intPrintEvenCnt = intPrintEvenCnt + 1
                        End If
                        intPrtCnt = intPrtCnt + 1
                    Next
                    
                    intPrintOddCnt = 0
                    intPrintEvenCnt = 0
                    For i = intPrtFNo + 1 To intPrtTNo Step 2
                        ReDim Preserve strPrint(intPrtCnt) As Variant
                        If i Mod 2 = 0 Then
                            strPrint(intPrtCnt) = strPrintEvenVal(intPrintOddCnt)
                            intPrintOddCnt = intPrintOddCnt + 1
                        Else
                            strPrint(intPrtCnt) = strPrintOddVal(intPrintEvenCnt)
                            intPrintEvenCnt = intPrintEvenCnt + 1
                        End If
                        intPrtCnt = intPrtCnt + 1
                    Next
                    
                    Erase strPrintOddVal
                    Erase strPrintEvenVal
                    intPrintOddCnt = 0
                    intPrintEvenCnt = 0
                End If
            Next
            
        Case "P0003"
            i = 0
'            strPlusXPos = 630
'            strAFont = "^A0N,35,25"
'
'            strHeader = "^XA" & vbLf
'            strHeader = strHeader & "^SEE:UHANGUL.DAT^FS" & vbLf
'            strHeader = strHeader & "^PON^FS" & vbLf
'            strHeader = strHeader & "^CWJ,E:KFONT3.FNT^FS" & vbLf
'            strHeader = strHeader & "^MD9"
            
            strHeader = strHeader & "^PR4"  'speed
            strHeader = strHeader & "^MD8"      'darkness
            
            For intPRow = 1 To spdRegOrderDetail.MaxRows
                strPrtFNo = GetText(spdRegOrderDetail, intPRow, 7)
                strPrtTNo = GetText(spdRegOrderDetail, intPRow, 8)
                intPrtFNo = CCur(Mid(strPrtFNo, 3))
                intPrtTNo = CCur(Mid(strPrtTNo, 3))
                For intPrt = intPrtFNo To intPrtTNo
                    strOutput = ""
                    Erase strPrtData
                    i = 0
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
                            
'203 DPI
'^XA
'^SEE:UHANGUL.DAT^FS
'^PON^FS
'^CWJ,E:KFONT3.FNT^FS
'^CI26
'^FO90,50^BY1,1,80^BCN,60,Y,N,N^FD32401002000001^FS
'^FO530,50^BY1,1,80^BCN,60,Y,N,N^FD32401002000001^FS
'^FO160,140^A0N,35,25^FH^FDKAF-TP408G^FS
'^FO600,140^A0N,35,25^FH^FDKAF-TP408G^FS
'^FO130,185^A0N,25,15^FH^FD1.5 x 200^FS
'^FO570,185^A0N,25,15^FH^FD1.5 x 200^FS
'^FO200,220^A0N,25,15^FH^FD2020.04.02^FS
'^FO640,220^A0N,25,15^FH^FD2020.04.02^FS
'^FO200,250^A0N,25,15^FH^FD2020.09.01^FS
'^FO640,250^A0N,25,15^FH^FD2020.09.01^FS
'^FO200,285^A0N,25,15^FH^FD-10 _7E 5'C^FS
'^FO640,285^A0N,25,15^FH^FD-10 _7E 5'C^FS
'^FO170,320^A0N,25,15^FH^FDK4F021Y2CS (P101)^FS
'^FO610,320^A0N,25,15^FH^FDK4F021Y2CS (P101)^FS
'^FO30,380^BY1,1,80^BCN,50,Y,N,N^FD32401002000001KK420101020000^FS
'^FO470,380^BY1,1,80^BCN,50,Y,N,N^FD32401002000001KK420101020000^FS
'^XZ
                            If strType = "바코드" Then
                                strLotSub = Format(Mid(strPrtFNo, 2, 1) & Format(intPrt, "00"), "0000")
                                strData = Mid(strData, 1, 18) & strLotSub & Mid(strData, 23, 6)
                                If Mid(strBarType, 1, 1) = "1" Then
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & "^BY" & "1,1,80"
                                    strPrtData(i) = strPrtData(i) & "^BC" & "N,50,Y,N,N"
                                    strPrtData(i) = strPrtData(i) & "^FD" & strData
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos200DPI) & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & "^BY" & "1,1,80"
                                    strPrtData(i) = strPrtData(i) & "^BC" & "N,50,Y,N,N"
                                    strPrtData(i) = strPrtData(i) & "^FD" & strData
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                Else
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & "^BQ" & "N,2,8"
                                    strPrtData(i) = strPrtData(i) & "^FD" & "MM,A" & strData
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos200DPI) & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & "^BQ" & "N,2,8"
                                    strPrtData(i) = strPrtData(i) & "^FD" & "MM,A" & strData
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                End If
                                strBarcode = strData
                            
                            ElseIf strType = "PartsID" Then
                                ReDim Preserve strPrtData(i) As String
                                strPrtData(i) = ""
                                strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
'                                strPrtData(i) = strPrtData(i) & "^CI26"
                                '^FO120,80^CI26^BY2,2,100^BCN,60,Y,N,N^FD32401002000001^FS
                                strPrtData(i) = strPrtData(i) & "^BY" & "1,1,80"
                                strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
                                strPrtData(i) = strPrtData(i) & "^FD" & strData
                                strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                i = i + 1
                            
                                ReDim Preserve strPrtData(i) As String
                                strPrtData(i) = ""
                                strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos200DPI) & "," & strYPos
'                                strPrtData(i) = strPrtData(i) & "^CI26"
                                strPrtData(i) = strPrtData(i) & "^BY" & "1,1,80"
                                strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
                                strPrtData(i) = strPrtData(i) & "^FD" & strData
                                strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                i = i + 1
                            ElseIf strType = "LotNo" Then
                                strLotSub = "P" & Format(Mid(strPrtFNo, 2, 1) & Format(intPrt, "00"), "000")
                                strLot = mGetP(strData, 1, "(")
                                strLot = strLot & "(" & strLotSub & ")"
                                
                            
                                ReDim Preserve strPrtData(i) As String
                                strPrtData(i) = ""
                                strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
'                                strPrtData(i) = strPrtData(i) & "^CI26"
                                strPrtData(i) = strPrtData(i) & strAFont200DPI
                                strPrtData(i) = strPrtData(i) & "^FH"
                                If strNamePrt = "Y" Then
                                    strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strLot
                                Else
                                    strPrtData(i) = strPrtData(i) & "^FD" & strLot
                                End If
                                strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                i = i + 1
                            
                                ReDim Preserve strPrtData(i) As String
                                strPrtData(i) = ""
                                strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos200DPI) & "," & strYPos
'                                strPrtData(i) = strPrtData(i) & "^CI26"
                                strPrtData(i) = strPrtData(i) & strAFont200DPI
                                strPrtData(i) = strPrtData(i) & "^FH"
                                If strNamePrt = "Y" Then
                                    strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strLot
                                Else
                                    strPrtData(i) = strPrtData(i) & "^FD" & strLot
                                End If
                                strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                i = i + 1
                            ElseIf strType = "보관온도" Then
                                
                                ReDim Preserve strPrtData(i) As String
                                strPrtData(i) = ""
                                strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                                If strType = "Name" Then
                                    strPrtData(i) = strPrtData(i) & "^A0N,35,25"
                                Else
                                    strPrtData(i) = strPrtData(i) & strAFont200DPI
                                End If
                                strPrtData(i) = strPrtData(i) & "^FH"
                                If strNamePrt = "Y" Then
                                    strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                                Else
                                    strPrtData(i) = strPrtData(i) & "^FD" & strData
                                End If
                                strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                'o 추가
                                strPrtData(i) = strPrtData(i) & "^FO" & strXPos + 65 & "," & strYPos & "^A0N,10,10^FH^FDo^FS"
                                'C 추가
                                strPrtData(i) = strPrtData(i) & "^FO" & strXPos + 70 & "," & strYPos & strAFont200DPI & "^FH^FDC^FS"
                                
                                i = i + 1
                            
                                ReDim Preserve strPrtData(i) As String
                                strPrtData(i) = ""
                                strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos200DPI) & "," & strYPos
                                If strType = "Name" Then
                                    strPrtData(i) = strPrtData(i) & "^A0N,35,25"
                                Else
                                    strPrtData(i) = strPrtData(i) & strAFont200DPI
                                End If
                                strPrtData(i) = strPrtData(i) & "^FH"
                                If strNamePrt = "Y" Then
                                    strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                                Else
                                    strPrtData(i) = strPrtData(i) & "^FD" & strData
                                End If
                                strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                'o 추가
                                strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos200DPI) + 65 & "," & strYPos & "^A0N,10,10^FH^FDo^FS"
                                'C 추가
                                strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos200DPI) + 70 & "," & strYPos & strAFont200DPI & "^FH^FDC^FS"
                                i = i + 1
                            Else
                                
                                ReDim Preserve strPrtData(i) As String
                                strPrtData(i) = ""
                                strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
'                                strPrtData(i) = strPrtData(i) & "^CI26"
                                If strType = "Name" Then
                                    strPrtData(i) = strPrtData(i) & "^A0N,35,25"
                                Else
                                    strPrtData(i) = strPrtData(i) & strAFont200DPI
                                End If
                                strPrtData(i) = strPrtData(i) & "^FH"
                                If strNamePrt = "Y" Then
                                    strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                                Else
                                    strPrtData(i) = strPrtData(i) & "^FD" & strData
                                End If
                                strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                i = i + 1
                            
                                ReDim Preserve strPrtData(i) As String
                                strPrtData(i) = ""
                                strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos200DPI) & "," & strYPos
'                                strPrtData(i) = strPrtData(i) & "^CI26"
                                If strType = "Name" Then
                                    strPrtData(i) = strPrtData(i) & "^A0N,35,25"
                                Else
                                    strPrtData(i) = strPrtData(i) & strAFont200DPI
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
                        Next
                    End With
                
                    strOutput = ""
                    For J = 0 To UBound(strPrtData)
                        strOutput = strOutput & strPrtData(J)
                    Next
                    strOutput = strHeader & strOutput & "^XZ" & vbLf
                    strOutput = Replace(strOutput, "~", "_7E")
                    'strOutput = Replace(strOutput, "℃", "'C")
                    strOutput = Replace(strOutput, "℃", "")
                    
                    '속도문제로 배열에 담았다가 한꺼번에  출력한다.
'                    ReDim Preserve strPrint(intPrtCnt) As Variant
'                    strPrint(intPrtCnt) = strOutput
'                    intPrtCnt = intPrtCnt + 1
                    
                    If intPrtTNo > 30 Then
                        Select Case intPrt
                            Case 1, 5, 9, 13, 17, 21, 25, 29, 33, 37, 41, 45, 49, 53, 57, 61, 65, 69, 73, 77, 81
                                ReDim Preserve strPrint1(intPrint1Cnt) As Variant
                                strPrint1(intPrint1Cnt) = strOutput
                                intPrint1Cnt = intPrint1Cnt + 1
                            Case 2, 6, 10, 14, 18, 22, 26, 30, 34, 38, 42, 46, 50, 54, 58, 62, 66, 70, 74, 78, 82
                                ReDim Preserve strPrint2(intPrint2Cnt) As Variant
                                strPrint2(intPrint2Cnt) = strOutput
                                intPrint2Cnt = intPrint2Cnt + 1
                            Case 3, 7, 11, 15, 19, 23, 27, 31, 35, 39, 43, 47, 51, 55, 59, 63, 67, 71, 75, 79, 83
                                ReDim Preserve strPrint3(intPrint3Cnt) As Variant
                                strPrint3(intPrint3Cnt) = strOutput
                                intPrint3Cnt = intPrint3Cnt + 1
                            Case 4, 8, 12, 16, 20, 24, 28, 32, 36, 40, 44, 48, 52, 56, 60, 64, 68, 72, 76, 80, 84
                                ReDim Preserve strPrint4(intPrint4Cnt) As Variant
                                strPrint4(intPrint4Cnt) = strOutput
                                intPrint4Cnt = intPrint4Cnt + 1
                        End Select
                    Else
                        '짝수
                        If intPrt Mod 2 = 0 Then
                            '속도문제로 배열에 담았다가 한꺼번에  출력한다.
                            ReDim Preserve strPrintEvenVal(intPrintEvenCnt) As Variant
                            strPrintEvenVal(intPrintEvenCnt) = strOutput
                            intPrintEvenCnt = intPrintEvenCnt + 1
                        '홀수
                        Else
                            '속도문제로 배열에 담았다가 한꺼번에  출력한다.
                            ReDim Preserve strPrintOddVal(intPrintOddCnt) As Variant
                            strPrintOddVal(intPrintOddCnt) = strOutput
                            intPrintOddCnt = intPrintOddCnt + 1
                        End If
                    End If
                    
                    '트래킹용
                    ReDim Preserve strTrackBC(intCnt)
                    strTrackBC(intCnt) = strBarcode
                    
                    '재출력용
                    ReDim Preserve strPrintVal(intCnt)
                    strPrintVal(intCnt) = strOutput
                    
                    intCnt = intCnt + 1
                    
                    blnPrint = True
                    txtComm.Text = strOutput
                    strOutput = ""
                Next
                
                If intPrtTNo > 30 Then
'''                    intPrint1Cnt = 0
'''                    intPrint2Cnt = 0
'''                    intPrint3Cnt = 0
'''                    intPrint4Cnt = 0
'''
'''                    For i = intPrtFNo To intPrtTNo
'''                        Select Case i
'''                            Case 1, 5, 9, 13, 17, 21, 25, 29, 33, 37, 41, 45, 49, 53, 57, 61, 65, 69, 73, 77, 81
'''                                ReDim Preserve strPrint(intPrtCnt) As Variant
'''                                strPrint(intPrtCnt) = strPrint1(intPrint1Cnt)
'''                                intPrint1Cnt = intPrint1Cnt + 1
'''                                intPrtCnt = intPrtCnt + 1
'''                        End Select
'''                    Next
'''
'''                    For i = intPrtFNo To intPrtTNo
'''                        Select Case i
'''                            Case 2, 6, 10, 14, 18, 22, 26, 30, 34, 38, 42, 46, 50, 54, 58, 62, 66, 70, 74, 78, 82
'''                                ReDim Preserve strPrint(intPrtCnt) As Variant
'''                                strPrint(intPrtCnt) = strPrint2(intPrint2Cnt)
'''                                intPrint2Cnt = intPrint2Cnt + 1
'''                                intPrtCnt = intPrtCnt + 1
'''                        End Select
'''                    Next
'''
'''                    For i = intPrtFNo To intPrtTNo
'''                        Select Case i
'''                            Case 3, 7, 11, 15, 19, 23, 27, 31, 35, 39, 43, 47, 51, 55, 59, 63, 67, 71, 75, 79, 83
'''                                ReDim Preserve strPrint(intPrtCnt) As Variant
'''                                strPrint(intPrtCnt) = strPrint3(intPrint3Cnt)
'''                                intPrint3Cnt = intPrint3Cnt + 1
'''                                intPrtCnt = intPrtCnt + 1
'''                        End Select
'''                    Next
'''
'''                    For i = intPrtFNo To intPrtTNo
'''                        Select Case i
'''                            Case 4, 8, 12, 16, 20, 24, 28, 32, 36, 40, 44, 48, 52, 56, 60, 64, 68, 72, 76, 80, 84
'''                                ReDim Preserve strPrint(intPrtCnt) As Variant
'''                                strPrint(intPrtCnt) = strPrint4(intPrint4Cnt)
'''                                intPrint4Cnt = intPrint4Cnt + 1
'''                                intPrtCnt = intPrtCnt + 1
'''                        End Select
'''                    Next
                    If intPRow = 1 Then
                        intP1Cnt = 0
                        intP2Cnt = 0
                        intP3Cnt = 0
                        intP4Cnt = 0
                    Else
                        intP1Cnt = intPrint1MaxCnt
                        intP2Cnt = intPrint2MaxCnt
                        intP3Cnt = intPrint3MaxCnt
                        intP4Cnt = intPrint4MaxCnt
                    End If
                    
                    intPrint1MaxCnt = intPrint1Cnt
                    intPrint2MaxCnt = intPrint2Cnt
                    intPrint3MaxCnt = intPrint3Cnt
                    intPrint4MaxCnt = intPrint4Cnt
                                        
                    For i = intPrtFNo To intPrtTNo
                        Select Case i
                            Case 1, 5, 9, 13, 17, 21, 25, 29, 33, 37, 41, 45, 49, 53, 57, 61, 65, 69, 73, 77, 81
                                ReDim Preserve strPrint(intPrtCnt) As Variant
                                strPrint(intPrtCnt) = strPrint1(intP1Cnt)
                                intP1Cnt = intP1Cnt + 1
                                intPrtCnt = intPrtCnt + 1
                        End Select
                    Next
                    
                    For i = intPrtFNo To intPrtTNo
                        Select Case i
                            Case 2, 6, 10, 14, 18, 22, 26, 30, 34, 38, 42, 46, 50, 54, 58, 62, 66, 70, 74, 78, 82
                                ReDim Preserve strPrint(intPrtCnt) As Variant
                                strPrint(intPrtCnt) = strPrint2(intP2Cnt)
                                intP2Cnt = intP2Cnt + 1
                                intPrtCnt = intPrtCnt + 1
                        End Select
                    Next
                    
                    For i = intPrtFNo To intPrtTNo
                        Select Case i
                            Case 3, 7, 11, 15, 19, 23, 27, 31, 35, 39, 43, 47, 51, 55, 59, 63, 67, 71, 75, 79, 83
                                ReDim Preserve strPrint(intPrtCnt) As Variant
                                strPrint(intPrtCnt) = strPrint3(intP3Cnt)
                                intP3Cnt = intP3Cnt + 1
                                intPrtCnt = intPrtCnt + 1
                        End Select
                    Next
                    
                    For i = intPrtFNo To intPrtTNo
                        Select Case i
                            Case 4, 8, 12, 16, 20, 24, 28, 32, 36, 40, 44, 48, 52, 56, 60, 64, 68, 72, 76, 80, 84
                                ReDim Preserve strPrint(intPrtCnt) As Variant
                                strPrint(intPrtCnt) = strPrint4(intP4Cnt)
                                intP4Cnt = intP4Cnt + 1
                                intPrtCnt = intPrtCnt + 1
                        End Select
                    Next
                Else
                    '속도문제로 배열에 담았다가 한꺼번에  출력한다.
                    intPrintOddCnt = 0
                    intPrintEvenCnt = 0
                    For i = intPrtFNo To intPrtTNo Step 2
                        ReDim Preserve strPrint(intPrtCnt) As Variant
                        If i Mod 2 = 0 Then
                            strPrint(intPrtCnt) = strPrintEvenVal(intPrintOddCnt)
                            intPrintOddCnt = intPrintOddCnt + 1
                        Else
                            strPrint(intPrtCnt) = strPrintOddVal(intPrintEvenCnt)
                            intPrintEvenCnt = intPrintEvenCnt + 1
                        End If
                        intPrtCnt = intPrtCnt + 1
                    Next
                    
                    intPrintOddCnt = 0
                    intPrintEvenCnt = 0
                    For i = intPrtFNo + 1 To intPrtTNo Step 2
                        ReDim Preserve strPrint(intPrtCnt) As Variant
                        If i Mod 2 = 0 Then
                            strPrint(intPrtCnt) = strPrintEvenVal(intPrintOddCnt)
                            intPrintOddCnt = intPrintOddCnt + 1
                        Else
                            strPrint(intPrtCnt) = strPrintOddVal(intPrintEvenCnt)
                            intPrintEvenCnt = intPrintEvenCnt + 1
                        End If
                        intPrtCnt = intPrtCnt + 1
                    Next
                    
                    Erase strPrintOddVal
                    Erase strPrintEvenVal
                    intPrintOddCnt = 0
                    intPrintEvenCnt = 0
                End If
            Next
            
        Case "P0004", "P0005", "P0008", "P0009"
            i = 0
'            strPlusXPos = 630
'            strAFont = "^A0N,35,25"
'
'            strHeader = "^XA" & vbLf
'            strHeader = strHeader & "^SEE:UHANGUL.DAT^FS" & vbLf
'            strHeader = strHeader & "^PON^FS" & vbLf
'            strHeader = strHeader & "^CWJ,E:KFONT3.FNT^FS" & vbLf
'            strHeader = strHeader & "^MD9"
            
            strHeader = strHeader & "^PR4"  'speed
            strHeader = strHeader & "^MD8"      'darkness
            
            For intPRow = 1 To spdRegOrderDetail.MaxRows
                strPrtFNo = GetText(spdRegOrderDetail, intPRow, 7)
                strPrtTNo = GetText(spdRegOrderDetail, intPRow, 8)
                intPrtFNo = CCur(Mid(strPrtFNo, 3))
                intPrtTNo = CCur(Mid(strPrtTNo, 3))
                For intPrt = intPrtFNo To intPrtTNo
                    strOutput = ""
                    Erase strPrtData
                    i = 0
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
                                strLotSub = Mid(strPrtFNo, 2, 1) & Format(intPrt, "00")
                                strData = Mid(strData, 1, 21) & strSlt & Mid(strData, 23, 2) & strLotSub & Mid(strData, 28, 5)
                                If Mid(strBarType, 1, 1) = "1" Then
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
'                                    strPrtData(i) = strPrtData(i) & "^CI26"
                                    strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
                                    strPrtData(i) = strPrtData(i) & "^FD" & strData
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) & "," & strYPos
'                                    strPrtData(i) = strPrtData(i) & "^CI26"
                                    strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
                                    strPrtData(i) = strPrtData(i) & "^FD" & strData
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                Else
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
'                                    strPrtData(i) = strPrtData(i) & "^CI26"
                                    'strPrtData(i) = strPrtData(i) & "^BQ" & "N,2,8"
                                    strPrtData(i) = strPrtData(i) & "^BQ" & ",2,8"
                                    'strPrtData(i) = strPrtData(i) & "^FD" & strData
                                    strPrtData(i) = strPrtData(i) & "^FD" & "MM,A" & strData
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) & "," & strYPos
'                                    strPrtData(i) = strPrtData(i) & "^CI26"
                                    'strPrtData(i) = strPrtData(i) & "^BQ" & "N,2,8"
                                    strPrtData(i) = strPrtData(i) & "^BQ" & ",2,8"
                                    'strPrtData(i) = strPrtData(i) & "^FD" & strData
                                    strPrtData(i) = strPrtData(i) & "^FD" & "MM,A" & strData
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                End If
                                strBarcode = strData
                            Else
                                If strType = "LotNo" Then
                                    strLotSub = Mid(strPrtFNo, 2, 1) & Format(intPrt, "00")
                                    strLot = mGetP(strData, 1, "(")
                                    strLot = strLot & "(P" & strLotSub & ")"
                                    
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
'                                    strPrtData(i) = strPrtData(i) & "^CI26"
                                    strPrtData(i) = strPrtData(i) & strAFont
                                    strPrtData(i) = strPrtData(i) & "^FH"
                                    If strNamePrt = "Y" Then
                                        strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strLot
                                    Else
                                        strPrtData(i) = strPrtData(i) & "^FD" & strLot
                                    End If
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) & "," & strYPos
'                                    strPrtData(i) = strPrtData(i) & "^CI26"
                                    strPrtData(i) = strPrtData(i) & strAFont
                                    strPrtData(i) = strPrtData(i) & "^FH"
                                    If strNamePrt = "Y" Then
                                        strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strLot
                                    Else
                                        strPrtData(i) = strPrtData(i) & "^FD" & strLot
                                    End If
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                
                                ElseIf strType = "바코드값" Then
                                    strLotSub = Mid(strPrtFNo, 2, 1) & Format(intPrt, "00")
                                    strData = Mid(strData, 1, 21) & strSlt & Mid(strData, 23, 2) & strLotSub & Mid(strData, 28, 5)
                                    
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
'                                    strPrtData(i) = strPrtData(i) & "^CI26"
                                    If strType = "Name" Then
                                        strPrtData(i) = strPrtData(i) & "^A0N,55,45"
                                    ElseIf strType = "바코드값" Then
                                        strPrtData(i) = strPrtData(i) & "^A0N,45,35"
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
                                
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) & "," & strYPos
'                                    strPrtData(i) = strPrtData(i) & "^CI26"
                                    If strType = "Name" Then
                                        strPrtData(i) = strPrtData(i) & "^A0N,55,45"
                                    ElseIf strType = "바코드값" Then
                                        strPrtData(i) = strPrtData(i) & "^A0N,45,35"
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
                                ElseIf strType = "보관온도" Then
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & "^CI26"
                                    If strType = "Name" Then
                                        strPrtData(i) = strPrtData(i) & "^A0N,55,45"
                                    ElseIf strType = "바코드값" Then
                                        strPrtData(i) = strPrtData(i) & "^A0N,45,35"
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
                                    'o추가
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos + 130 & "," & CCur(strYPos) - 5 & "^A0N,20,20^FDo^FS" & vbLf
                                    'C추가
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos + 140 & "," & strYPos & strAFont & "^FDC^FS" & vbLf
                                    
                                    i = i + 1
                                
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & "^CI26"
                                    If strType = "Name" Then
                                        strPrtData(i) = strPrtData(i) & "^A0N,55,45"
                                    ElseIf strType = "바코드값" Then
                                        strPrtData(i) = strPrtData(i) & "^A0N,45,35"
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
                                    'o추가
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) + 130 & "," & CCur(strYPos) - 5 & "^A0N,20,20^FDo^FS" & vbLf
                                    'C추가
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) + 140 & "," & strYPos & strAFont & "^FDC^FS" & vbLf

                                    i = i + 1
                                Else
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & "^CI26"
                                    If strType = "Name" Then
                                        strPrtData(i) = strPrtData(i) & "^A0N,55,45"
                                    ElseIf strType = "바코드값" Then
                                        strPrtData(i) = strPrtData(i) & "^A0N,45,35"
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
                                
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & "^CI26"
                                    If strType = "Name" Then
                                        strPrtData(i) = strPrtData(i) & "^A0N,55,45"
                                    ElseIf strType = "바코드값" Then
                                        strPrtData(i) = strPrtData(i) & "^A0N,45,35"
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
                    
                    strOutput = Replace(strOutput, "~", "_7E")
                    'strOutput = Replace(strOutput, "℃", "'C")
                    strOutput = Replace(strOutput, "℃", "")
                    
                    '속도문제로 배열에 담았다가 한꺼번에  출력한다.
                    'comEqp.Output = strOutput
                    'ReDim Preserve strPrint(intPrtCnt) As Variant
                    'strPrint(intPrtCnt) = strOutput
                    'intPrtCnt = intPrtCnt + 1
                    
                    If intPrtTNo > 30 Then
                        Select Case intPrt
                            Case 1, 5, 9, 13, 17, 21, 25, 29, 33, 37, 41, 45, 49, 53, 57, 61, 65, 69, 73, 77, 81
                                ReDim Preserve strPrint1(intPrint1Cnt) As Variant
                                strPrint1(intPrint1Cnt) = strOutput
                                intPrint1Cnt = intPrint1Cnt + 1
                            Case 2, 6, 10, 14, 18, 22, 26, 30, 34, 38, 42, 46, 50, 54, 58, 62, 66, 70, 74, 78, 82
                                ReDim Preserve strPrint2(intPrint2Cnt) As Variant
                                strPrint2(intPrint2Cnt) = strOutput
                                intPrint2Cnt = intPrint2Cnt + 1
                            Case 3, 7, 11, 15, 19, 23, 27, 31, 35, 39, 43, 47, 51, 55, 59, 63, 67, 71, 75, 79, 83
                                ReDim Preserve strPrint3(intPrint3Cnt) As Variant
                                strPrint3(intPrint3Cnt) = strOutput
                                intPrint3Cnt = intPrint3Cnt + 1
                            Case 4, 8, 12, 16, 20, 24, 28, 32, 36, 40, 44, 48, 52, 56, 60, 64, 68, 72, 76, 80, 84
                                ReDim Preserve strPrint4(intPrint4Cnt) As Variant
                                strPrint4(intPrint4Cnt) = strOutput
                                intPrint4Cnt = intPrint4Cnt + 1
                        End Select
                    Else
                        '짝수
                        If intPrt Mod 2 = 0 Then
                            '속도문제로 배열에 담았다가 한꺼번에  출력한다.
                            ReDim Preserve strPrintEvenVal(intPrintEvenCnt) As Variant
                            strPrintEvenVal(intPrintEvenCnt) = strOutput
                            intPrintEvenCnt = intPrintEvenCnt + 1
                        '홀수
                        Else
                            '속도문제로 배열에 담았다가 한꺼번에  출력한다.
                            ReDim Preserve strPrintOddVal(intPrintOddCnt) As Variant
                            strPrintOddVal(intPrintOddCnt) = strOutput
                            intPrintOddCnt = intPrintOddCnt + 1
                        End If
                    End If
                    
                    '트래킹용
                    ReDim Preserve strTrackBC(intCnt)
                    strTrackBC(intCnt) = strBarcode
                    '재출력용
                    ReDim Preserve strPrintVal(intCnt)
                    strPrintVal(intCnt) = strOutput
                    intCnt = intCnt + 1
                    
                    blnPrint = True
                    txtComm.Text = strOutput
                    strOutput = ""
                Next
                                
                If intPrtTNo > 30 Then
                    If intPRow = 1 Then
                        intP1Cnt = 0
                        intP2Cnt = 0
                        intP3Cnt = 0
                        intP4Cnt = 0
                    Else
                        intP1Cnt = intPrint1MaxCnt
                        intP2Cnt = intPrint2MaxCnt
                        intP3Cnt = intPrint3MaxCnt
                        intP4Cnt = intPrint4MaxCnt
                    End If
                    
                    intPrint1MaxCnt = intPrint1Cnt
                    intPrint2MaxCnt = intPrint2Cnt
                    intPrint3MaxCnt = intPrint3Cnt
                    intPrint4MaxCnt = intPrint4Cnt
                                        
                    For i = intPrtFNo To intPrtTNo
                        Select Case i
                            Case 1, 5, 9, 13, 17, 21, 25, 29, 33, 37, 41, 45, 49, 53, 57, 61, 65, 69, 73, 77, 81
                                ReDim Preserve strPrint(intPrtCnt) As Variant
                                strPrint(intPrtCnt) = strPrint1(intP1Cnt)
                                intP1Cnt = intP1Cnt + 1
                                intPrtCnt = intPrtCnt + 1
                        End Select
                    Next
                    
                    For i = intPrtFNo To intPrtTNo
                        Select Case i
                            Case 2, 6, 10, 14, 18, 22, 26, 30, 34, 38, 42, 46, 50, 54, 58, 62, 66, 70, 74, 78, 82
                                ReDim Preserve strPrint(intPrtCnt) As Variant
                                strPrint(intPrtCnt) = strPrint2(intP2Cnt)
                                intP2Cnt = intP2Cnt + 1
                                intPrtCnt = intPrtCnt + 1
                        End Select
                    Next
                    
                    For i = intPrtFNo To intPrtTNo
                        Select Case i
                            Case 3, 7, 11, 15, 19, 23, 27, 31, 35, 39, 43, 47, 51, 55, 59, 63, 67, 71, 75, 79, 83
                                ReDim Preserve strPrint(intPrtCnt) As Variant
                                strPrint(intPrtCnt) = strPrint3(intP3Cnt)
                                intP3Cnt = intP3Cnt + 1
                                intPrtCnt = intPrtCnt + 1
                        End Select
                    Next
                    
                    For i = intPrtFNo To intPrtTNo
                        Select Case i
                            Case 4, 8, 12, 16, 20, 24, 28, 32, 36, 40, 44, 48, 52, 56, 60, 64, 68, 72, 76, 80, 84
                                ReDim Preserve strPrint(intPrtCnt) As Variant
                                strPrint(intPrtCnt) = strPrint4(intP4Cnt)
                                intP4Cnt = intP4Cnt + 1
                                intPrtCnt = intPrtCnt + 1
                        End Select
                    Next
                Else
                    '속도문제로 배열에 담았다가 한꺼번에  출력한다.
                    intPrintOddCnt = 0
                    intPrintEvenCnt = 0
                    For i = intPrtFNo To intPrtTNo Step 2
                        ReDim Preserve strPrint(intPrtCnt) As Variant
                        If i Mod 2 = 0 Then
                            strPrint(intPrtCnt) = strPrintEvenVal(intPrintOddCnt)
                            intPrintOddCnt = intPrintOddCnt + 1
                        Else
                            strPrint(intPrtCnt) = strPrintOddVal(intPrintEvenCnt)
                            intPrintEvenCnt = intPrintEvenCnt + 1
                        End If
                        intPrtCnt = intPrtCnt + 1
                    Next
                    
                    intPrintOddCnt = 0
                    intPrintEvenCnt = 0
                    For i = intPrtFNo + 1 To intPrtTNo Step 2
                        ReDim Preserve strPrint(intPrtCnt) As Variant
                        If i Mod 2 = 0 Then
                            strPrint(intPrtCnt) = strPrintEvenVal(intPrintOddCnt)
                            intPrintOddCnt = intPrintOddCnt + 1
                        Else
                            strPrint(intPrtCnt) = strPrintOddVal(intPrintEvenCnt)
                            intPrintEvenCnt = intPrintEvenCnt + 1
                        End If
                        intPrtCnt = intPrtCnt + 1
                    Next
                    
                    Erase strPrintOddVal
                    Erase strPrintEvenVal
                    intPrintOddCnt = 0
                    intPrintEvenCnt = 0
                End If
            Next
            
        Case "P0006", "P0007", "P0010"
            i = 0
'            strPlusXPos = 630
'            strAFont = "^A0N,35,25"
'
'            strHeader = "^XA" & vbLf
'            strHeader = strHeader & "^SEE:UHANGUL.DAT^FS" & vbLf
'            strHeader = strHeader & "^PON^FS" & vbLf
'            strHeader = strHeader & "^CWJ,E:KFONT3.FNT^FS" & vbLf
'            strHeader = strHeader & "^MD9"
            
            strHeader = strHeader & "^PR4"  'speed
            strHeader = strHeader & "^MD8"      'darkness
            
            For intPRow = 1 To spdRegOrderDetail.MaxRows
                strPrtFNo = GetText(spdRegOrderDetail, intPRow, 7)
                strPrtTNo = GetText(spdRegOrderDetail, intPRow, 8)
                intPrtFNo = CCur(Mid(strPrtFNo, 3))
                intPrtTNo = CCur(Mid(strPrtTNo, 3))
                For intPrt = intPrtFNo To intPrtTNo
                    strOutput = ""
                    Erase strPrtData
                    i = 0
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
                                strLotSub = Mid(strPrtFNo, 2, 1) & Format(intPrt, "00")
                                strData = Mid(strData, 1, 12) & strLotSub & Mid(strData, 16, 8)
                                
                                If Mid(strBarType, 1, 1) = "1" Then
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
'                                    strPrtData(i) = strPrtData(i) & "^CI26"
                                    strPrtData(i) = strPrtData(i) & "^BC" & "N,100,Y,N,N"
                                    strPrtData(i) = strPrtData(i) & "^FD" & strData
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & "^CI26"
                                    strPrtData(i) = strPrtData(i) & "^BC" & "N,100,Y,N,N"
                                    strPrtData(i) = strPrtData(i) & "^FD" & strData
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                Else
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
'                                    strPrtData(i) = strPrtData(i) & "^CI26"
                                    'strPrtData(i) = strPrtData(i) & "^BQ" & "N,2,8"
                                    strPrtData(i) = strPrtData(i) & "^BQ" & ",2,8"
                                    'strPrtData(i) = strPrtData(i) & "^FD" & strData
                                    strPrtData(i) = strPrtData(i) & "^FD" & "MM,A" & strData
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & "^CI26"
                                    'strPrtData(i) = strPrtData(i) & "^BQ" & "N,2,8"
                                    strPrtData(i) = strPrtData(i) & "^BQ" & ",2,8"
                                    'strPrtData(i) = strPrtData(i) & "^FD" & strData
                                    strPrtData(i) = strPrtData(i) & "^FD" & "MM,A" & strData
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                End If
                                strBarcode = strData
                            ElseIf strType = "PartsID" Then
                                'strLotSub = strSlt & Format(intPrt, "00")
                                If Mid(strBarType, 1, 1) = "1" Or strBarType = "" Then
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
'                                    strPrtData(i) = strPrtData(i) & "^CI26"
                                    strPrtData(i) = strPrtData(i) & "^BC" & "N,100,Y,N,N"
                                    strPrtData(i) = strPrtData(i) & "^FD" & strData
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) & "," & strYPos
'                                    strPrtData(i) = strPrtData(i) & "^CI26"
                                    strPrtData(i) = strPrtData(i) & "^BC" & "N,100,Y,N,N"
                                    strPrtData(i) = strPrtData(i) & "^FD" & strData
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                Else
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
'                                    strPrtData(i) = strPrtData(i) & "^CI26"
                                    strPrtData(i) = strPrtData(i) & "^BQ" & "N,2,8"
                                    strPrtData(i) = strPrtData(i) & "^FD" & strData
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) & "," & strYPos
'                                    strPrtData(i) = strPrtData(i) & "^CI26"
                                    strPrtData(i) = strPrtData(i) & "^BQ" & "N,2,8"
                                    strPrtData(i) = strPrtData(i) & "^FD" & strData
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                End If
                            Else
                                If strType = "LotNo" Then
                                    'strLot = mGetP(strData, 1, "(")
                                    strLotSub = Mid(strPrtFNo, 2, 1) & Format(intPrt, "00")

                                    strLot = mGetP(strData, 1, "(") & "(P" & strLotSub & ")"
                                    
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
'                                    strPrtData(i) = strPrtData(i) & "^CI26"
                                    strPrtData(i) = strPrtData(i) & strAFont
                                    strPrtData(i) = strPrtData(i) & "^FH"
                                    If strNamePrt = "Y" Then
                                        strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strLot
                                    Else
                                        strPrtData(i) = strPrtData(i) & "^FD" & strLot
                                    End If
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) & "," & strYPos
'                                    strPrtData(i) = strPrtData(i) & "^CI26"
                                    strPrtData(i) = strPrtData(i) & strAFont
                                    strPrtData(i) = strPrtData(i) & "^FH"
                                    If strNamePrt = "Y" Then
                                        strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strLot
                                    Else
                                        strPrtData(i) = strPrtData(i) & "^FD" & strLot
                                    End If
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                ElseIf strType = "보관온도" Then
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                                    If strType = "Name" Then
                                        strPrtData(i) = strPrtData(i) & "^A0N,55,45"
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
                                    'o추가
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos + 130 & "," & CCur(strYPos) - 5 & "^A0N,20,20^FDo^FS" & vbLf
                                    'C추가
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos + 140 & "," & strYPos & strAFont & "^FDC^FS" & vbLf
                                    
                                    i = i + 1
                                
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) & "," & strYPos
                                    If strType = "Name" Then
                                        strPrtData(i) = strPrtData(i) & "^A0N,55,45"
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
                                    'o추가
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) + 130 & "," & CCur(strYPos) - 5 & "^A0N,20,20^FDo^FS" & vbLf
                                    'C추가
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) + 140 & "," & strYPos & strAFont & "^FDC^FS" & vbLf
                                    
                                    i = i + 1
                                Else
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
'                                    strPrtData(i) = strPrtData(i) & "^CI26"
                                    'strPrtData(i) = strPrtData(i) & strAFont
                                    If strType = "Name" Then
                                        strPrtData(i) = strPrtData(i) & "^A0N,55,45"
'                                    ElseIf strType = "바코드값" Then
'                                        strPrtData(i) = strPrtData(i) & "^A0N,45,35"
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
                                
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) & "," & strYPos
'                                    strPrtData(i) = strPrtData(i) & "^CI26"
                                    'strPrtData(i) = strPrtData(i) & strAFont
                                    If strType = "Name" Then
                                        strPrtData(i) = strPrtData(i) & "^A0N,55,45"
'                                    ElseIf strType = "바코드값" Then
'                                        strPrtData(i) = strPrtData(i) & "^A0N,45,35"
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
                    
                    strOutput = Replace(strOutput, "~", "_7E")
                    'strOutput = Replace(strOutput, "℃", "'C")
                    strOutput = Replace(strOutput, "℃", "")
                    
                    '속도문제로 배열에 담았다가 한꺼번에  출력한다.
'                    ReDim Preserve strPrint(intPrtCnt) As Variant
'                    strPrint(intPrtCnt) = strOutput
'                    intPrtCnt = intPrtCnt + 1
                    
                    If intPrtTNo > 30 Then
                        Select Case intPrt
                            Case 1, 5, 9, 13, 17, 21, 25, 29, 33, 37, 41, 45, 49, 53, 57, 61, 65, 69, 73, 77, 81
                                ReDim Preserve strPrint1(intPrint1Cnt) As Variant
                                strPrint1(intPrint1Cnt) = strOutput
                                intPrint1Cnt = intPrint1Cnt + 1
                            Case 2, 6, 10, 14, 18, 22, 26, 30, 34, 38, 42, 46, 50, 54, 58, 62, 66, 70, 74, 78, 82
                                ReDim Preserve strPrint2(intPrint2Cnt) As Variant
                                strPrint2(intPrint2Cnt) = strOutput
                                intPrint2Cnt = intPrint2Cnt + 1
                            Case 3, 7, 11, 15, 19, 23, 27, 31, 35, 39, 43, 47, 51, 55, 59, 63, 67, 71, 75, 79, 83
                                ReDim Preserve strPrint3(intPrint3Cnt) As Variant
                                strPrint3(intPrint3Cnt) = strOutput
                                intPrint3Cnt = intPrint3Cnt + 1
                            Case 4, 8, 12, 16, 20, 24, 28, 32, 36, 40, 44, 48, 52, 56, 60, 64, 68, 72, 76, 80, 84
                                ReDim Preserve strPrint4(intPrint4Cnt) As Variant
                                strPrint4(intPrint4Cnt) = strOutput
                                intPrint4Cnt = intPrint4Cnt + 1
                        End Select
                    Else
                        '짝수
                        If intPrt Mod 2 = 0 Then
                            '속도문제로 배열에 담았다가 한꺼번에  출력한다.
                            ReDim Preserve strPrintEvenVal(intPrintEvenCnt) As Variant
                            strPrintEvenVal(intPrintEvenCnt) = strOutput
                            intPrintEvenCnt = intPrintEvenCnt + 1
                        '홀수
                        Else
                            '속도문제로 배열에 담았다가 한꺼번에  출력한다.
                            ReDim Preserve strPrintOddVal(intPrintOddCnt) As Variant
                            strPrintOddVal(intPrintOddCnt) = strOutput
                            intPrintOddCnt = intPrintOddCnt + 1
                        End If
                    End If
                    
                    '트래킹용
                    ReDim Preserve strTrackBC(intCnt)
                    strTrackBC(intCnt) = strBarcode
                    
                    '재출력용
                    ReDim Preserve strPrintVal(intCnt)
                    strPrintVal(intCnt) = strOutput
                    
                    intCnt = intCnt + 1
                    
                    blnPrint = True
                    txtComm.Text = strOutput
                    strOutput = ""
                Next
                
                If intPrtTNo > 30 Then
                    If intPRow = 1 Then
                        intP1Cnt = 0
                        intP2Cnt = 0
                        intP3Cnt = 0
                        intP4Cnt = 0
                    Else
                        intP1Cnt = intPrint1MaxCnt
                        intP2Cnt = intPrint2MaxCnt
                        intP3Cnt = intPrint3MaxCnt
                        intP4Cnt = intPrint4MaxCnt
                    End If
                    
                    intPrint1MaxCnt = intPrint1Cnt
                    intPrint2MaxCnt = intPrint2Cnt
                    intPrint3MaxCnt = intPrint3Cnt
                    intPrint4MaxCnt = intPrint4Cnt
                                        
                    For i = intPrtFNo To intPrtTNo
                        Select Case i
                            Case 1, 5, 9, 13, 17, 21, 25, 29, 33, 37, 41, 45, 49, 53, 57, 61, 65, 69, 73, 77, 81
                                ReDim Preserve strPrint(intPrtCnt) As Variant
                                strPrint(intPrtCnt) = strPrint1(intP1Cnt)
                                intP1Cnt = intP1Cnt + 1
                                intPrtCnt = intPrtCnt + 1
                        End Select
                    Next
                    
                    For i = intPrtFNo To intPrtTNo
                        Select Case i
                            Case 2, 6, 10, 14, 18, 22, 26, 30, 34, 38, 42, 46, 50, 54, 58, 62, 66, 70, 74, 78, 82
                                ReDim Preserve strPrint(intPrtCnt) As Variant
                                strPrint(intPrtCnt) = strPrint2(intP2Cnt)
                                intP2Cnt = intP2Cnt + 1
                                intPrtCnt = intPrtCnt + 1
                        End Select
                    Next
                    
                    For i = intPrtFNo To intPrtTNo
                        Select Case i
                            Case 3, 7, 11, 15, 19, 23, 27, 31, 35, 39, 43, 47, 51, 55, 59, 63, 67, 71, 75, 79, 83
                                ReDim Preserve strPrint(intPrtCnt) As Variant
                                strPrint(intPrtCnt) = strPrint3(intP3Cnt)
                                intP3Cnt = intP3Cnt + 1
                                intPrtCnt = intPrtCnt + 1
                        End Select
                    Next
                    
                    For i = intPrtFNo To intPrtTNo
                        Select Case i
                            Case 4, 8, 12, 16, 20, 24, 28, 32, 36, 40, 44, 48, 52, 56, 60, 64, 68, 72, 76, 80, 84
                                ReDim Preserve strPrint(intPrtCnt) As Variant
                                strPrint(intPrtCnt) = strPrint4(intP4Cnt)
                                intP4Cnt = intP4Cnt + 1
                                intPrtCnt = intPrtCnt + 1
                        End Select
                    Next
                Else
                    '속도문제로 배열에 담았다가 한꺼번에  출력한다.
                    intPrintOddCnt = 0
                    intPrintEvenCnt = 0
                    For i = intPrtFNo To intPrtTNo Step 2
                        ReDim Preserve strPrint(intPrtCnt) As Variant
                        If i Mod 2 = 0 Then
                            strPrint(intPrtCnt) = strPrintEvenVal(intPrintOddCnt)
                            intPrintOddCnt = intPrintOddCnt + 1
                        Else
                            strPrint(intPrtCnt) = strPrintOddVal(intPrintEvenCnt)
                            intPrintEvenCnt = intPrintEvenCnt + 1
                        End If
                        intPrtCnt = intPrtCnt + 1
                    Next
                    
                    intPrintOddCnt = 0
                    intPrintEvenCnt = 0
                    For i = intPrtFNo + 1 To intPrtTNo Step 2
                        ReDim Preserve strPrint(intPrtCnt) As Variant
                        If i Mod 2 = 0 Then
                            strPrint(intPrtCnt) = strPrintEvenVal(intPrintOddCnt)
                            intPrintOddCnt = intPrintOddCnt + 1
                        Else
                            strPrint(intPrtCnt) = strPrintOddVal(intPrintEvenCnt)
                            intPrintEvenCnt = intPrintEvenCnt + 1
                        End If
                        intPrtCnt = intPrtCnt + 1
                    Next
                    
                    Erase strPrintOddVal
                    Erase strPrintEvenVal
                    intPrintOddCnt = 0
                    intPrintEvenCnt = 0
                End If
            Next
    End Select
   
    If UBound(strPrint) > 0 Then
        '-- 프로그레스바 열기
        frmProgress.Show
        frmProgress.ZOrder 0
        frmProgress.Xprog.Min = 0
        frmProgress.Xprog.Max = IIf(CCur(txtReelQTY.Text) = 0, 100, CCur(txtReelQTY.Text))
        frmProgress.lblProgress.Caption = ""
        
        For intCount = 0 To UBound(strPrint)
            'If intCount + 1 > CCur(txtReelQTY.Text) Then
            '    Exit For
            'End If
            Debug.Print strPrint(intCount)
            comEqp.Output = strPrint(intCount)
'
'Printer.Print strPrint(intCount) '<--요넘이 ZPL코드
'Printer.EndDoc

            Call SetPrtData("REEL" & "_" & txtProdNm.Text & "_" & txtProdLen.Text, "[출력시간 : " & Format(Now, "yyyy-mm-dd hh:mm:ss.nnn") & "]" & vbCrLf & strPrint(intCount), "A")

            '-- 프로그레스바 진행
            frmProgress.Xprog.Value = intCount
            frmProgress.lblProgress.Caption = " 출력중 [전체수량 / 현재수량]  " & CCur(txtReelQTY.Text) & " / " & intCount
            DoEvents

        Next
        
        '-- 프로그레스바 닫기
        Unload frmProgress
        
    End If
   
    If blnPrint = True Then
        Call SetPackTrack(strTrackBC, strPrintVal)
    End If
    
End Sub




Private Sub SetPackTrack(ByVal pBarcode As Variant, ByVal pPrint As Variant)
    Dim intCnt      As Integer
    
    
    For intCnt = 0 To UBound(pBarcode)
        gPackTrack.ORDERDT = Format(txtProdOrderDt.Text, "yyyymmdd")     'Key
        gPackTrack.PRODCD = txtProdCd.Text                               'Key
        gPackTrack.REELBAR = pBarcode(intCnt)
        gPackTrack.PPBAR = ""
        gPackTrack.ICEBAR = ""
        gPackTrack.PPBARIN = ""
        gPackTrack.ICEBARIN = ""
        gPackTrack.LOTNO = txtLotNo.Text
        gPackTrack.REELPRTID = gKUKDO.USERID
        gPackTrack.REELPRTDT = ""
        gPackTrack.PPPRTID = ""
        gPackTrack.PPPRTDT = ""
        gPackTrack.ICEPRTID = ""
        gPackTrack.ICEPRTDT = ""
        '재출력용
        gPackTrack.REELVAL = pPrint(intCnt)
        gPackTrack.PPVAL = ""
        gPackTrack.ICEVAL = ""
        
        '트래킹 저장
        '-- Insert / Update 찾아오기
        Set AdoRs = Get_Pack_Track(gPackTrack.ORDERDT, gPackTrack.PRODCD, gPackTrack.REELBAR, "", "")
    
        If AdoRs.RecordCount = 0 Then
            'INSERT
            If Set_Pack_Track("IN", "R") Then
            End If
        Else
            'UPDATE
            If Set_Pack_Track("UP", "R") Then
            End If
        End If
    Next
    
    gOrder.ORDDATE = Format(txtProdOrderDt.Text, "yyyymmdd")      'Key
'    gOrder.PRODPOSNO = txtProdPosNo.Text                          'Key
    gOrder.PRODCD = txtProdCd.Text                                'Key
    gOrder.SLITINGNO = txtSlittingNo.Text                         'Key


    '촐력구분 UPDATE : LBL_PROD_ORDER..CLOSE_YN
    If Set_Order_CloseYN("UP") Then
        If gAllPrt = False Then
            Call cmdSearch_Click
        End If
    End If


End Sub


Private Sub cmdSamplePrint_Click()
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
    Dim strPrtData()   As String
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
    Dim strPlusXPos200DPI As String
    Dim strAFont    As String
    Dim strAFont200DPI    As String
    Dim i           As Integer
    Dim J           As Integer
    Dim k           As Integer
    Dim blnPrint    As Boolean
    
    '-- 트래킹용
    Dim intCnt          As Integer
    Dim strTrackBC()    As Variant
    '-- 재출력용
    Dim strPrintVal()   As Variant
    '-- 출력용
    Dim intPrtCnt   As Integer
    Dim strPrint()  As Variant
    
    Dim intPrintOddCnt      As Integer
    Dim strPrintOddVal()    As Variant   '홀수
    Dim intPrintEvenCnt     As Integer
    Dim strPrintEvenVal()   As Variant   '짝수
    
    Dim intPrint1Cnt        As Integer
    Dim strPrint1()         As Variant
    Dim intPrint2Cnt        As Integer
    Dim strPrint2()         As Variant
    Dim intPrint3Cnt        As Integer
    Dim strPrint3()         As Variant
    Dim intPrint4Cnt        As Integer
    Dim strPrint4()         As Variant
    
    Dim intPrint1MaxCnt        As Integer
    Dim intPrint2MaxCnt        As Integer
    Dim intPrint3MaxCnt        As Integer
    Dim intPrint4MaxCnt        As Integer
    
    Dim intP1Cnt        As Integer
    Dim intP2Cnt        As Integer
    Dim intP3Cnt        As Integer
    Dim intP4Cnt        As Integer
    
'    Dim intPrint1MaxCntB        As Integer
'    Dim intPrint2MaxCntB        As Integer
'    Dim intPrint3MaxCntB        As Integer
'    Dim intPrint4MaxCntB        As Integer
    
    Dim intCount    As Integer
    
    '-- 첫번째 P번호가 홀수인지 짝수인지 갖고 있는다.
    ' 홀수일 경우 blnOdd = True
'    Dim blnOdd      As Boolean
    
    intPrintOddCnt = 0
    intPrintEvenCnt = 0
    Erase strPrintOddVal
    Erase strPrintEvenVal
    
    intPrint1Cnt = 0
    intPrint2Cnt = 0
    intPrint3Cnt = 0
    intPrint4Cnt = 0
    
    intPrint1MaxCnt = 0
    intPrint2MaxCnt = 0
    intPrint3MaxCnt = 0
    intPrint4MaxCnt = 0

    intP1Cnt = 0
    intP2Cnt = 0
    intP3Cnt = 0
    intP4Cnt = 0

    Erase strPrint1
    Erase strPrint2
    Erase strPrint3
    Erase strPrint4
    
    blnPrint = False
    strBarcode = ""
    Erase strPrtData
    Erase strTrackBC
    Erase strPrintVal
    intCnt = 0
    intPrt = 0
    intPrtCnt = 0
    
    If txtProdNm.Text = "" Then
        MsgBox "출력할 Reel라벨을 선택하세요", vbOKOnly + vbCritical, Me.Caption
        spdPrtReel.SetFocus
        Exit Sub
    End If
    
    strPlusXPos = 630
    strPlusXPos200DPI = 430
    strAFont = "^A0N,35,30"
    strAFont200DPI = "^A0N,25,15"
    
    strHeader = "^XA" & vbLf
    strHeader = strHeader & "^SEE:UHANGUL.DAT^FS" & vbLf
    strHeader = strHeader & "^PON^FS" & vbLf
    strHeader = strHeader & "^CWJ,E:KFONT3.FNT^FS" & vbLf
    strHeader = strHeader & "^CI26"
    
    Select Case lblstrPrtLabelName.Caption
        Case "P0011", "P0018", "P0019", "P0020"
            strPlusXPos = 500
            i = 0
            For intPRow = 1 To spdRegOrderDetail.MaxRows
                strPrtFNo = GetText(spdRegOrderDetail, intPRow, 7)
                strPrtTNo = strPrtFNo
                intPrtFNo = CCur(Mid(strPrtFNo, 3))
                intPrtTNo = intPrtFNo
                For intPrt = intPrtFNo To intPrtTNo
                    strOutput = ""
                    Erase strPrtData
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
                                strLotSub = "P" & Mid(strPrtFNo, 2, 1) & Format(intPrt, "00")
                                strData = Mid(strData, 1, 9) & strLotSub
                                '좌측 출력
                                ReDim Preserve strPrtData(i) As String
                                strPrtData(i) = ""
                                strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                                strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
                                strPrtData(i) = strPrtData(i) & "^FD" & strData
                                strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                i = i + 1
                                '우측 출력
                                ReDim Preserve strPrtData(i) As String
                                strPrtData(i) = ""
                                strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) & "," & strYPos
                                strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
                                strPrtData(i) = strPrtData(i) & "^FD" & strData
                                strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                i = i + 1
                                strBarcode = strData
                            Else
                                If strType = "ACF Lot" Or strType = "LotNo" Or strType = "Lot No." Then
                                    strLot = mGetP(strData, 1, "(")
                                    strLot = strLot & "(" & strLotSub & ")"
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & strAFont
                                    strPrtData(i) = strPrtData(i) & "^FH"
                                    If strNamePrt = "Y" Then
                                        strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strLot
                                    Else
                                        strPrtData(i) = strPrtData(i) & "^FD" & strLot
                                    End If
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & strAFont
                                    strPrtData(i) = strPrtData(i) & "^FH"
                                    If strNamePrt = "Y" Then
                                        strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strLot
                                    Else
                                        strPrtData(i) = strPrtData(i) & "^FD" & strLot
                                    End If
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                ElseIf strType = "Storage" Then
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & strAFont
                                    strPrtData(i) = strPrtData(i) & "^FH"
                                    If strNamePrt = "Y" Then
                                        strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                                    Else
                                        strPrtData(i) = strPrtData(i) & "^FD" & strData
                                    End If
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos + 240 & "," & CCur(strYPos) - 5 & "^A0N,20,20^FDo^FS" & vbLf
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos + 250 & "," & strYPos & strAFont & "^FDC^FS" & vbLf
                                    i = i + 1
                                    
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & strAFont
                                    strPrtData(i) = strPrtData(i) & "^FH"
                                    If strNamePrt = "Y" Then
                                        strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                                    Else
                                        strPrtData(i) = strPrtData(i) & "^FD" & strData
                                    End If
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) + 240 & "," & CCur(strYPos) - 5 & "^A0N,20,20^FDo^FS" & vbLf
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) + 250 & "," & strYPos & strAFont & "^FDC^FS" & vbLf
                                    i = i + 1
                                Else
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & strAFont
                                    strPrtData(i) = strPrtData(i) & "^FH"
                                    If strNamePrt = "Y" Then
                                        strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                                    Else
                                        strPrtData(i) = strPrtData(i) & "^FD" & strData
                                    End If
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) & "," & strYPos
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
                    '특수문자변환
                    strOutput = Replace(strOutput, "~", "_7E")  'asc("~") =  126
                    strOutput = Replace(strOutput, "℃", "") '°C
                    strOutput = strHeader & strOutput & "^XZ" & vbLf
                    ReDim Preserve strPrintOddVal(intPrintOddCnt) As Variant
                    strPrintOddVal(intPrintOddCnt) = strOutput
                    intPrintOddCnt = intPrintOddCnt + 1
                    blnPrint = True
                    strOutput = ""
                Next
                    
                '속도문제로 배열에 담았다가 한꺼번에  출력한다.
                intPrintOddCnt = 0
                intPrintEvenCnt = 0
                For i = intPrtFNo To intPrtTNo Step 2
                    ReDim Preserve strPrint(intPrtCnt) As Variant
                    If i Mod 2 = 0 Then
                        strPrint(intPrtCnt) = strPrintEvenVal(intPrintOddCnt)
                        intPrintOddCnt = intPrintOddCnt + 1
                    Else
                        strPrint(intPrtCnt) = strPrintOddVal(intPrintEvenCnt)
                        intPrintEvenCnt = intPrintEvenCnt + 1
                    End If
                    intPrtCnt = intPrtCnt + 1
                Next
            Next
        
        Case "P0001", "P0002"
            i = 0
            For intPRow = 1 To spdRegOrderDetail.MaxRows
                strPrtFNo = GetText(spdRegOrderDetail, intPRow, 7)
                strPrtTNo = strPrtFNo
                intPrtFNo = CCur(Mid(strPrtFNo, 3))
                intPrtTNo = intPrtFNo
                For intPrt = intPrtFNo To intPrtTNo
                    strOutput = ""
                    Erase strPrtData
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
                                strLotSub = "K" & Mid(strPrtFNo, 2, 1) & Format(intPrt, "00")
                                strData = Mid(strData, 1, 13) & strLotSub & Mid(strData, 18, 5)
                                If Mid(strBarType, 1, 1) = "1" Then
                                    '좌측 출력
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
                                    strPrtData(i) = strPrtData(i) & "^FD" & strData
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                    '우측 출력
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
                                    strPrtData(i) = strPrtData(i) & "^FD" & strData
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                Else
                                    '좌측 출력
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & "^BY" & "2,4,50"
                                    strPrtData(i) = strPrtData(i) & "^BQ"
                                    strPrtData(i) = strPrtData(i) & "^FD" & strData
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                    '우측 출력
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & "^BY" & "2,4,50"
                                    strPrtData(i) = strPrtData(i) & "^BQ"
                                    strPrtData(i) = strPrtData(i) & "^FD" & strData
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                End If
                                strBarcode = strData
                            Else
                                If strType = "ACF Lot" Or strType = "LotNo" Then
                                    strLot = mGetP(strData, 1, "(")
                                    strLot = strLot & "(" & strLotSub & ")"
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & strAFont
                                    strPrtData(i) = strPrtData(i) & "^FH"
                                    If strNamePrt = "Y" Then
                                        strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strLot
                                    Else
                                        strPrtData(i) = strPrtData(i) & "^FD" & strLot
                                    End If
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & strAFont
                                    strPrtData(i) = strPrtData(i) & "^FH"
                                    If strNamePrt = "Y" Then
                                        strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strLot
                                    Else
                                        strPrtData(i) = strPrtData(i) & "^FD" & strLot
                                    End If
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                ElseIf strType = "Storage Temperature" Then
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & strAFont
                                    strPrtData(i) = strPrtData(i) & "^FH"
                                    If strNamePrt = "Y" Then
                                        strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                                    Else
                                        strPrtData(i) = strPrtData(i) & "^FD" & strData
                                    End If
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    'o추가
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos + 400 & "," & CCur(strYPos) - 5 & "^A0N,20,20^FDo^FS" & vbLf
                                    'C추가
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos + 410 & "," & strYPos & strAFont & "^FDC^FS" & vbLf
                                    i = i + 1
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & strAFont
                                    strPrtData(i) = strPrtData(i) & "^FH"
                                    If strNamePrt = "Y" Then
                                        strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                                    Else
                                        strPrtData(i) = strPrtData(i) & "^FD" & strData
                                    End If
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    'o추가
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) + 400 & "," & CCur(strYPos) - 5 & "^A0N,20,20^FDo^FS" & vbLf
                                    'C추가
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) + 410 & "," & strYPos & strAFont & "^FDC^FS" & vbLf
                                    i = i + 1
                                Else
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & strAFont
                                    strPrtData(i) = strPrtData(i) & "^FH"
                                    If strNamePrt = "Y" Then
                                        strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                                    Else
                                        strPrtData(i) = strPrtData(i) & "^FD" & strData
                                    End If
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                    
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) & "," & strYPos
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
                    '특수문자변환
                    strOutput = Replace(strOutput, "~", "_7E")  'asc("~") =  126
                    strOutput = Replace(strOutput, "℃", "") '°C
                    strOutput = strHeader & strOutput & "^XZ" & vbLf
                    ReDim Preserve strPrintOddVal(intPrintOddCnt) As Variant
                    strPrintOddVal(intPrintOddCnt) = strOutput
                    intPrintOddCnt = intPrintOddCnt + 1
                    blnPrint = True
                    strOutput = ""
                Next
                    
                '속도문제로 배열에 담았다가 한꺼번에  출력한다.
                intPrintOddCnt = 0
                intPrintEvenCnt = 0
                For i = intPrtFNo To intPrtTNo Step 2
                    ReDim Preserve strPrint(intPrtCnt) As Variant
                    If i Mod 2 = 0 Then
                        strPrint(intPrtCnt) = strPrintEvenVal(intPrintOddCnt)
                        intPrintOddCnt = intPrintOddCnt + 1
                    Else
                        strPrint(intPrtCnt) = strPrintOddVal(intPrintEvenCnt)
                        intPrintEvenCnt = intPrintEvenCnt + 1
                    End If
                    intPrtCnt = intPrtCnt + 1
                Next
            Next
            
        Case "P0003"
            i = 0
            strHeader = strHeader & "^PR4"  'speed
            strHeader = strHeader & "^MD8"      'darkness
            For intPRow = 1 To spdRegOrderDetail.MaxRows
                strPrtFNo = GetText(spdRegOrderDetail, intPRow, 7)
                strPrtTNo = strPrtTNo
                intPrtFNo = CCur(Mid(strPrtFNo, 3))
                intPrtTNo = intPrtFNo
                For intPrt = intPrtFNo To intPrtTNo
                    strOutput = ""
                    Erase strPrtData
                    i = 0
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
                                strLotSub = Format(Mid(strPrtFNo, 2, 1) & Format(intPrt, "00"), "0000")
                                strData = Mid(strData, 1, 18) & strLotSub & Mid(strData, 23, 6)
                                If Mid(strBarType, 1, 1) = "1" Then
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & "^BY" & "1,1,80"
                                    strPrtData(i) = strPrtData(i) & "^BC" & "N,50,Y,N,N"
                                    strPrtData(i) = strPrtData(i) & "^FD" & strData
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos200DPI) & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & "^BY" & "1,1,80"
                                    strPrtData(i) = strPrtData(i) & "^BC" & "N,50,Y,N,N"
                                    strPrtData(i) = strPrtData(i) & "^FD" & strData
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                Else
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & "^BQ" & "N,2,8"
                                    strPrtData(i) = strPrtData(i) & "^FD" & "MM,A" & strData
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos200DPI) & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & "^BQ" & "N,2,8"
                                    strPrtData(i) = strPrtData(i) & "^FD" & "MM,A" & strData
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                End If
                                strBarcode = strData
                            ElseIf strType = "PartsID" Then
                                ReDim Preserve strPrtData(i) As String
                                strPrtData(i) = ""
                                strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                                strPrtData(i) = strPrtData(i) & "^BY" & "1,1,80"
                                strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
                                strPrtData(i) = strPrtData(i) & "^FD" & strData
                                strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                i = i + 1
                                ReDim Preserve strPrtData(i) As String
                                strPrtData(i) = ""
                                strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos200DPI) & "," & strYPos
                                strPrtData(i) = strPrtData(i) & "^BY" & "1,1,80"
                                strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
                                strPrtData(i) = strPrtData(i) & "^FD" & strData
                                strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                i = i + 1
                            ElseIf strType = "LotNo" Then
                                strLotSub = "P" & Format(Mid(strPrtFNo, 2, 1) & Format(intPrt, "00"), "000")
                                strLot = mGetP(strData, 1, "(")
                                strLot = strLot & "(" & strLotSub & ")"
                                ReDim Preserve strPrtData(i) As String
                                strPrtData(i) = ""
                                strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                                strPrtData(i) = strPrtData(i) & strAFont200DPI
                                strPrtData(i) = strPrtData(i) & "^FH"
                                If strNamePrt = "Y" Then
                                    strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strLot
                                Else
                                    strPrtData(i) = strPrtData(i) & "^FD" & strLot
                                End If
                                strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                i = i + 1
                                ReDim Preserve strPrtData(i) As String
                                strPrtData(i) = ""
                                strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos200DPI) & "," & strYPos
                                strPrtData(i) = strPrtData(i) & strAFont200DPI
                                strPrtData(i) = strPrtData(i) & "^FH"
                                If strNamePrt = "Y" Then
                                    strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strLot
                                Else
                                    strPrtData(i) = strPrtData(i) & "^FD" & strLot
                                End If
                                strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                i = i + 1
                            ElseIf strType = "보관온도" Then
                                ReDim Preserve strPrtData(i) As String
                                strPrtData(i) = ""
                                strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                                If strType = "Name" Then
                                    strPrtData(i) = strPrtData(i) & "^A0N,35,25"
                                Else
                                    strPrtData(i) = strPrtData(i) & strAFont200DPI
                                End If
                                strPrtData(i) = strPrtData(i) & "^FH"
                                If strNamePrt = "Y" Then
                                    strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                                Else
                                    strPrtData(i) = strPrtData(i) & "^FD" & strData
                                End If
                                strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                'o 추가
                                strPrtData(i) = strPrtData(i) & "^FO" & strXPos + 65 & "," & strYPos & "^A0N,10,10^FH^FDo^FS"
                                'C 추가
                                strPrtData(i) = strPrtData(i) & "^FO" & strXPos + 70 & "," & strYPos & strAFont200DPI & "^FH^FDC^FS"
                                i = i + 1
                                ReDim Preserve strPrtData(i) As String
                                strPrtData(i) = ""
                                strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos200DPI) & "," & strYPos
                                If strType = "Name" Then
                                    strPrtData(i) = strPrtData(i) & "^A0N,35,25"
                                Else
                                    strPrtData(i) = strPrtData(i) & strAFont200DPI
                                End If
                                strPrtData(i) = strPrtData(i) & "^FH"
                                If strNamePrt = "Y" Then
                                    strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                                Else
                                    strPrtData(i) = strPrtData(i) & "^FD" & strData
                                End If
                                strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                'o 추가
                                strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos200DPI) + 65 & "," & strYPos & "^A0N,10,10^FH^FDo^FS"
                                'C 추가
                                strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos200DPI) + 70 & "," & strYPos & strAFont200DPI & "^FH^FDC^FS"
                                i = i + 1
                            Else
                                
                                ReDim Preserve strPrtData(i) As String
                                strPrtData(i) = ""
                                strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                                If strType = "Name" Then
                                    strPrtData(i) = strPrtData(i) & "^A0N,35,25"
                                Else
                                    strPrtData(i) = strPrtData(i) & strAFont200DPI
                                End If
                                strPrtData(i) = strPrtData(i) & "^FH"
                                If strNamePrt = "Y" Then
                                    strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strData
                                Else
                                    strPrtData(i) = strPrtData(i) & "^FD" & strData
                                End If
                                strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                i = i + 1
                                ReDim Preserve strPrtData(i) As String
                                strPrtData(i) = ""
                                strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos200DPI) & "," & strYPos
                                If strType = "Name" Then
                                    strPrtData(i) = strPrtData(i) & "^A0N,35,25"
                                Else
                                    strPrtData(i) = strPrtData(i) & strAFont200DPI
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
                        Next
                    End With
                    strOutput = ""
                    For J = 0 To UBound(strPrtData)
                        strOutput = strOutput & strPrtData(J)
                    Next
                    '특수문자변환
                    strOutput = Replace(strOutput, "~", "_7E")  'asc("~") =  126
                    strOutput = Replace(strOutput, "℃", "") '°C
                    strOutput = strHeader & strOutput & "^XZ" & vbLf
                    ReDim Preserve strPrintOddVal(intPrintOddCnt) As Variant
                    strPrintOddVal(intPrintOddCnt) = strOutput
                    intPrintOddCnt = intPrintOddCnt + 1
                    blnPrint = True
                    strOutput = ""
                Next
                    
                '속도문제로 배열에 담았다가 한꺼번에  출력한다.
                intPrintOddCnt = 0
                intPrintEvenCnt = 0
                For i = intPrtFNo To intPrtTNo Step 2
                    ReDim Preserve strPrint(intPrtCnt) As Variant
                    If i Mod 2 = 0 Then
                        strPrint(intPrtCnt) = strPrintEvenVal(intPrintOddCnt)
                        intPrintOddCnt = intPrintOddCnt + 1
                    Else
                        strPrint(intPrtCnt) = strPrintOddVal(intPrintEvenCnt)
                        intPrintEvenCnt = intPrintEvenCnt + 1
                    End If
                    intPrtCnt = intPrtCnt + 1
                Next
            Next
        Case "P0004", "P0005", "P0008", "P0009"
            i = 0
            strHeader = strHeader & "^PR4"  'speed
            strHeader = strHeader & "^MD8"      'darkness
            For intPRow = 1 To spdRegOrderDetail.MaxRows
                strPrtFNo = GetText(spdRegOrderDetail, intPRow, 7)
                strPrtTNo = strPrtFNo
                intPrtFNo = CCur(Mid(strPrtFNo, 3))
                intPrtTNo = intPrtFNo
                For intPrt = intPrtFNo To intPrtTNo
                    strOutput = ""
                    Erase strPrtData
                    i = 0
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
                                strLotSub = Mid(strPrtFNo, 2, 1) & Format(intPrt, "00")
                                strData = Mid(strData, 1, 21) & strSlt & Mid(strData, 23, 2) & strLotSub & Mid(strData, 28, 5)
                                If Mid(strBarType, 1, 1) = "1" Then
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
                                    strPrtData(i) = strPrtData(i) & "^FD" & strData
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & "^BC" & "N,60,Y,N,N"
                                    strPrtData(i) = strPrtData(i) & "^FD" & strData
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                Else
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & "^BQ" & ",2,8"
                                    strPrtData(i) = strPrtData(i) & "^FD" & "MM,A" & strData
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & "^BQ" & ",2,8"
                                    strPrtData(i) = strPrtData(i) & "^FD" & "MM,A" & strData
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                End If
                                strBarcode = strData
                            Else
                                If strType = "LotNo" Then
                                    strLotSub = Mid(strPrtFNo, 2, 1) & Format(intPrt, "00")
                                    strLot = mGetP(strData, 1, "(")
                                    strLot = strLot & "(P" & strLotSub & ")"
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & strAFont
                                    strPrtData(i) = strPrtData(i) & "^FH"
                                    If strNamePrt = "Y" Then
                                        strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strLot
                                    Else
                                        strPrtData(i) = strPrtData(i) & "^FD" & strLot
                                    End If
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & strAFont
                                    strPrtData(i) = strPrtData(i) & "^FH"
                                    If strNamePrt = "Y" Then
                                        strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strLot
                                    Else
                                        strPrtData(i) = strPrtData(i) & "^FD" & strLot
                                    End If
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                
                                ElseIf strType = "바코드값" Then
                                    strLotSub = Mid(strPrtFNo, 2, 1) & Format(intPrt, "00")
                                    strData = Mid(strData, 1, 21) & strSlt & Mid(strData, 23, 2) & strLotSub & Mid(strData, 28, 5)
                                    
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                                    If strType = "Name" Then
                                        strPrtData(i) = strPrtData(i) & "^A0N,55,45"
                                    ElseIf strType = "바코드값" Then
                                        strPrtData(i) = strPrtData(i) & "^A0N,45,35"
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
                                
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) & "," & strYPos
                                    If strType = "Name" Then
                                        strPrtData(i) = strPrtData(i) & "^A0N,55,45"
                                    ElseIf strType = "바코드값" Then
                                        strPrtData(i) = strPrtData(i) & "^A0N,45,35"
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
                                ElseIf strType = "보관온도" Then
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & "^CI26"
                                    If strType = "Name" Then
                                        strPrtData(i) = strPrtData(i) & "^A0N,55,45"
                                    ElseIf strType = "바코드값" Then
                                        strPrtData(i) = strPrtData(i) & "^A0N,45,35"
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
                                    'o추가
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos + 130 & "," & CCur(strYPos) - 5 & "^A0N,20,20^FDo^FS" & vbLf
                                    'C추가
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos + 140 & "," & strYPos & strAFont & "^FDC^FS" & vbLf
                                    i = i + 1
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & "^CI26"
                                    If strType = "Name" Then
                                        strPrtData(i) = strPrtData(i) & "^A0N,55,45"
                                    ElseIf strType = "바코드값" Then
                                        strPrtData(i) = strPrtData(i) & "^A0N,45,35"
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
                                    'o추가
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) + 130 & "," & CCur(strYPos) - 5 & "^A0N,20,20^FDo^FS" & vbLf
                                    'C추가
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) + 140 & "," & strYPos & strAFont & "^FDC^FS" & vbLf
                                    i = i + 1
                                Else
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & "^CI26"
                                    If strType = "Name" Then
                                        strPrtData(i) = strPrtData(i) & "^A0N,55,45"
                                    ElseIf strType = "바코드값" Then
                                        strPrtData(i) = strPrtData(i) & "^A0N,45,35"
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
                                
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & "^CI26"
                                    If strType = "Name" Then
                                        strPrtData(i) = strPrtData(i) & "^A0N,55,45"
                                    ElseIf strType = "바코드값" Then
                                        strPrtData(i) = strPrtData(i) & "^A0N,45,35"
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
                    '특수문자변환
                    strOutput = Replace(strOutput, "~", "_7E")  'asc("~") =  126
                    strOutput = Replace(strOutput, "℃", "") '°C
                    strOutput = strHeader & strOutput & "^XZ" & vbLf
                    ReDim Preserve strPrintOddVal(intPrintOddCnt) As Variant
                    strPrintOddVal(intPrintOddCnt) = strOutput
                    intPrintOddCnt = intPrintOddCnt + 1
                    blnPrint = True
                    strOutput = ""
                Next
                    
                '속도문제로 배열에 담았다가 한꺼번에  출력한다.
                intPrintOddCnt = 0
                intPrintEvenCnt = 0
                For i = intPrtFNo To intPrtTNo Step 2
                    ReDim Preserve strPrint(intPrtCnt) As Variant
                    If i Mod 2 = 0 Then
                        strPrint(intPrtCnt) = strPrintEvenVal(intPrintOddCnt)
                        intPrintOddCnt = intPrintOddCnt + 1
                    Else
                        strPrint(intPrtCnt) = strPrintOddVal(intPrintEvenCnt)
                        intPrintEvenCnt = intPrintEvenCnt + 1
                    End If
                    intPrtCnt = intPrtCnt + 1
                Next
            Next
            
        Case "P0006", "P0007", "P0010"
            i = 0
            strHeader = strHeader & "^PR4"  'speed
            strHeader = strHeader & "^MD8"      'darkness
            For intPRow = 1 To spdRegOrderDetail.MaxRows
                strPrtFNo = GetText(spdRegOrderDetail, intPRow, 7)
                strPrtTNo = strPrtFNo
                intPrtFNo = CCur(Mid(strPrtFNo, 3))
                intPrtTNo = intPrtFNo
                For intPrt = intPrtFNo To intPrtTNo
                    strOutput = ""
                    Erase strPrtData
                    i = 0
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
                                strLotSub = Mid(strPrtFNo, 2, 1) & Format(intPrt, "00")
                                strData = Mid(strData, 1, 12) & strLotSub & Mid(strData, 16, 8)
                                If Mid(strBarType, 1, 1) = "1" Then
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & "^BC" & "N,100,Y,N,N"
                                    strPrtData(i) = strPrtData(i) & "^FD" & strData
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & "^CI26"
                                    strPrtData(i) = strPrtData(i) & "^BC" & "N,100,Y,N,N"
                                    strPrtData(i) = strPrtData(i) & "^FD" & strData
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                Else
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & "^BQ" & ",2,8"
                                    strPrtData(i) = strPrtData(i) & "^FD" & "MM,A" & strData
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & "^CI26"
                                    strPrtData(i) = strPrtData(i) & "^BQ" & ",2,8"
                                    strPrtData(i) = strPrtData(i) & "^FD" & "MM,A" & strData
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                End If
                                strBarcode = strData
                            ElseIf strType = "PartsID" Then
                                If Mid(strBarType, 1, 1) = "1" Or strBarType = "" Then
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & "^BC" & "N,100,Y,N,N"
                                    strPrtData(i) = strPrtData(i) & "^FD" & strData
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & "^BC" & "N,100,Y,N,N"
                                    strPrtData(i) = strPrtData(i) & "^FD" & strData
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                Else
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & "^BQ" & "N,2,8"
                                    strPrtData(i) = strPrtData(i) & "^FD" & strData
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & "^BQ" & "N,2,8"
                                    strPrtData(i) = strPrtData(i) & "^FD" & strData
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                End If
                            Else
                                If strType = "LotNo" Then
                                    strLotSub = Mid(strPrtFNo, 2, 1) & Format(intPrt, "00")
                                    strLot = mGetP(strData, 1, "(") & "(P" & strLotSub & ")"
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & strAFont
                                    strPrtData(i) = strPrtData(i) & "^FH"
                                    If strNamePrt = "Y" Then
                                        strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strLot
                                    Else
                                        strPrtData(i) = strPrtData(i) & "^FD" & strLot
                                    End If
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) & "," & strYPos
                                    strPrtData(i) = strPrtData(i) & strAFont
                                    strPrtData(i) = strPrtData(i) & "^FH"
                                    If strNamePrt = "Y" Then
                                        strPrtData(i) = strPrtData(i) & "^FD" & strType & " : " & strLot
                                    Else
                                        strPrtData(i) = strPrtData(i) & "^FD" & strLot
                                    End If
                                    strPrtData(i) = strPrtData(i) & "^FS" & vbLf
                                    i = i + 1
                                ElseIf strType = "보관온도" Then
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                                    If strType = "Name" Then
                                        strPrtData(i) = strPrtData(i) & "^A0N,55,45"
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
                                    'o추가
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos + 130 & "," & CCur(strYPos) - 5 & "^A0N,20,20^FDo^FS" & vbLf
                                    'C추가
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos + 140 & "," & strYPos & strAFont & "^FDC^FS" & vbLf
                                    i = i + 1
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) & "," & strYPos
                                    If strType = "Name" Then
                                        strPrtData(i) = strPrtData(i) & "^A0N,55,45"
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
                                    'o추가
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) + 130 & "," & CCur(strYPos) - 5 & "^A0N,20,20^FDo^FS" & vbLf
                                    'C추가
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) + 140 & "," & strYPos & strAFont & "^FDC^FS" & vbLf
                                    i = i + 1
                                Else
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & strXPos & "," & strYPos
                                    If strType = "Name" Then
                                        strPrtData(i) = strPrtData(i) & "^A0N,55,45"
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
                                    ReDim Preserve strPrtData(i) As String
                                    strPrtData(i) = ""
                                    strPrtData(i) = strPrtData(i) & "^FO" & CCur(strXPos) + CCur(strPlusXPos) & "," & strYPos
                                    If strType = "Name" Then
                                        strPrtData(i) = strPrtData(i) & "^A0N,55,45"
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
                    '특수문자변환
                    strOutput = Replace(strOutput, "~", "_7E")  'asc("~") =  126
                    strOutput = Replace(strOutput, "℃", "") '°C
                    strOutput = strHeader & strOutput & "^XZ" & vbLf
                    ReDim Preserve strPrintOddVal(intPrintOddCnt) As Variant
                    strPrintOddVal(intPrintOddCnt) = strOutput
                    intPrintOddCnt = intPrintOddCnt + 1
                    blnPrint = True
                    strOutput = ""
                Next
                    
                '속도문제로 배열에 담았다가 한꺼번에  출력한다.
                intPrintOddCnt = 0
                intPrintEvenCnt = 0
                For i = intPrtFNo To intPrtTNo Step 2
                    ReDim Preserve strPrint(intPrtCnt) As Variant
                    If i Mod 2 = 0 Then
                        strPrint(intPrtCnt) = strPrintEvenVal(intPrintOddCnt)
                        intPrintOddCnt = intPrintOddCnt + 1
                    Else
                        strPrint(intPrtCnt) = strPrintOddVal(intPrintEvenCnt)
                        intPrintEvenCnt = intPrintEvenCnt + 1
                    End If
                    intPrtCnt = intPrtCnt + 1
                Next
            Next


    End Select
   
    If UBound(strPrint) >= 0 Then
        For intCount = 0 To UBound(strPrint)
            comEqp.Output = strPrint(intCount)
            DoEvents
        Next
    End If
   
End Sub

Private Sub cmdSearch_Click()
    Dim strFromDt    As String
    Dim strToDt      As String
    Dim strYN        As String
    
    strFromDt = Format(dtpFromDate, "yyyymmdd")
    strToDt = Format(dtpToDate, "yyyymmdd")
    
    Call cmdClear_Click
    
    Call GetReelOrderList(strFromDt, strToDt, "", "", "R", IIf(chkYN.Value = "0", "N", ""))

End Sub


Private Sub cmdView_Click()
    
    If txtComm.Visible = False Then
        txtComm.Visible = True
    Else
        txtComm.Visible = False
    End If
    
End Sub



Private Sub Form_Load()
    
    Unload frmPrtPP
    Unload frmPrtICE
    Unload frmPrtICE
    Unload frmPrtReprint
    
    gAllPrt = False
    
    Call CtlInitializing
    
    '-- 통신열기
    Call OpenCommunication
    
    'Call Ws_PrtLabel2("")
    
End Sub

'
'Private Sub Ws_PrtLabel2(Ps_Param As String) '오른쪽 USB타입 제브라프린터 추가
'
'    Dim prt As Printer
'    Dim prtname As String
'    Dim currentPrt As String
'    Dim i   As Integer
'
'    For i = 0 To Printers.Count - 1
'        prtname = Printers(i).DeviceName
'        Debug.Print prtname
'            'If InStr(1, prtname, "Zebra") Then  '제브라 프린터 찾음
'            If InStr(1, prtname, "ZPL") Then  '제브라 프린터 찾음
'            currentPrt = Printers(i).DeviceName
'            MsgBox currentPrt
'            Exit For
'        End If
'    Next i
'
'    For Each prt In Printers
'        If prt.DeviceName = currentPrt Then
'            Set Printer = prt
'                Exit For
'        End If
'    Next
'
'
'
'End Sub


Private Sub Ws_PrtLabel2(Ps_Param As String) '오른쪽 USB타입 제브라프린터 추가

    Dim prt As Printer
    Dim PRTNAME As String
    Dim currentPrt As String
    Dim i   As Integer

    Dim filehandle      As Integer

    filehandle = FreeFile

    For i = 0 To Printers.Count - 1
        PRTNAME = Printers(i).DeviceName
        'If InStr(1, prtname, "Zebra") Then
        If InStr(1, PRTNAME, "ZPL") Then
            currentPrt = Printers(i).DeviceName
            Exit For
        End If
    Next i

    For Each prt In Printers
        If prt.DeviceName = currentPrt Then
            Set Printer = prt
            Exit For
        End If
    Next



    Open "USB002:" For Output As #filehandle  'LPT포트를 알수가 없어요 ㅠㅠ
    Print #filehandle, Ps_Param '<--요넘이 ZPL코드


    Close #filehandle



End Sub



Private Sub OpenCommunication()

On Error GoTo ErrHandle

'    frmPrtPP.comEqp.PortOpen = False
'    frmPrtICE.comEqp.PortOpen = False
    
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
        'comEqp.PortOpen = False
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
        Call SetText(spdPrtReel, "SLT No", 0, 11):           .ColWidth(11) = 6
        Call SetText(spdPrtReel, "수량", 0, 12):             .ColWidth(12) = 4
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
        Call SetText(spdPrtReelDetail, "항목", 0, 1):           .ColWidth(1) = 20
        Call SetText(spdPrtReelDetail, "순서", 0, 2):           .ColWidth(2) = 4
        Call SetText(spdPrtReelDetail, "내용", 0, 3):           .ColWidth(3) = 30
        Call SetText(spdPrtReelDetail, "X", 0, 4):              .ColWidth(4) = 0
        Call SetText(spdPrtReelDetail, "Y", 0, 5):              .ColWidth(5) = 0
        Call SetText(spdPrtReelDetail, "", 0, 6):               .ColWidth(6) = 0
        Call SetText(spdPrtReelDetail, "", 0, 7):               .ColWidth(7) = 0
        Call SetText(spdPrtReelDetail, "", 0, 8):               .ColWidth(8) = 0
        Call SetText(spdPrtReelDetail, "", 0, 9):               .ColWidth(9) = 0
        Call SetText(spdPrtReelDetail, "", 0, 10):              .ColWidth(10) = 0
        Call SetText(spdPrtReelDetail, "", 0, 11):              .ColWidth(11) = 0
        Call SetText(spdPrtReelDetail, "", 0, 12):              .ColWidth(12) = 0
        Call SetText(spdPrtReelDetail, "", 0, 13):              .ColWidth(13) = 0
        Call SetText(spdPrtReelDetail, "", 0, 14):              .ColWidth(14) = 0
        Call SetText(spdPrtReelDetail, "", 0, 15):              .ColWidth(15) = 0
        Call SetText(spdPrtReelDetail, "", 0, 16):              .ColWidth(16) = 0
        Call SetText(spdPrtReelDetail, "", 0, 17):              .ColWidth(17) = 0
        Call SetText(spdPrtReelDetail, "", 0, 18):              .ColWidth(18) = 0
        Call SetText(spdPrtReelDetail, "", 0, 19):              .ColWidth(19) = 0
    
        .MaxRows = 0
    End With
    
    With spdRegOrderDetail
        Call SetText(spdRegOrderDetail, "제조일자", 0, 1):        .ColWidth(1) = 0
        Call SetText(spdRegOrderDetail, "순번", 0, 2):            .ColWidth(2) = 0
        Call SetText(spdRegOrderDetail, "제품코드", 0, 3):        .ColWidth(3) = 0
        Call SetText(spdRegOrderDetail, "SLT No", 0, 4):          .ColWidth(4) = 0
        Call SetText(spdRegOrderDetail, "No", 0, 5):              .ColWidth(5) = 4
        Call SetText(spdRegOrderDetail, "SLT내용", 0, 6):         .ColWidth(6) = 28
        Call SetText(spdRegOrderDetail, "시작번호", 0, 7):        .ColWidth(7) = 10
        Call SetText(spdRegOrderDetail, "끝번호", 0, 8):          .ColWidth(8) = 10
        Call SetText(spdRegOrderDetail, "", 0, 9):                .ColWidth(9) = 0
        Call SetText(spdRegOrderDetail, "No", 0, 10):             .ColWidth(10) = 0
        Call SetText(spdRegOrderDetail, "", 0, 11):               .ColWidth(11) = 0
        Call SetText(spdRegOrderDetail, "", 0, 12):               .ColWidth(12) = 0
        Call SetText(spdRegOrderDetail, "", 0, 13):               .ColWidth(13) = 0
        Call SetText(spdRegOrderDetail, "", 0, 14):               .ColWidth(14) = 0
        Call SetText(spdRegOrderDetail, "사용여부", 0, 15):       .ColWidth(15) = 0
        Call SetText(spdRegOrderDetail, "입력자", 0, 16):         .ColWidth(16) = 0
        Call SetText(spdRegOrderDetail, "입력일시", 0, 17):       .ColWidth(17) = 0
        Call SetText(spdRegOrderDetail, "수정자", 0, 18):         .ColWidth(18) = 0
        Call SetText(spdRegOrderDetail, "수정일시", 0, 19):       .ColWidth(19) = 0
    
        .MaxRows = 0
    End With
    
    
    dtpFromDate.Value = Now - 1
    dtpToDate.Value = Now

    txtLotNo.Text = ""
    txtProdNm.Text = ""
    txtProdOrderDt.Text = ""
'    txtProdPosNo.Text = ""
    txtPackNm.Text = ""
    txtCompNm.Text = ""
    txtSlittingNo.Text = ""
    txtReelQTY.Text = ""
    txtOrderMemo.Text = ""

    gSORT = 0

End Sub

' 작업지시서 리스트 가져옴 'strDate, cboProdPosNo.Text, cboProdCd.Text, cboSlittingNo.Text
Private Sub GetOrderDetail(ByVal pDate As String, ByVal pProCd As String, ByVal pSltNo As String)
    
    Set AdoRs = Get_OrderDetail(pDate, pProCd, pSltNo)
            
    If AdoRs Is Nothing Then
        '등록된 정보 없음
    Else
        spdRegOrderDetail.MaxRows = 0
        Do Until AdoRs.EOF
            With spdRegOrderDetail
                .MaxRows = .MaxRows + 1
                Call SetText(spdRegOrderDetail, AdoRs.Fields("PROD_ORDER_DT").Value & "", .MaxRows, 1)
                'Call SetText(spdRegOrderDetail, AdoRs.Fields("PROD_POS_NO").Value & "", .MaxRows, 2)
                Call SetText(spdRegOrderDetail, AdoRs.Fields("PROD_CD").Value & "", .MaxRows, 3)
                Call SetText(spdRegOrderDetail, AdoRs.Fields("SLITING_NO").Value & "", .MaxRows, 4)
                Call SetText(spdRegOrderDetail, AdoRs.Fields("SEQ_NO").Value & "", .MaxRows, 5)
                Call SetText(spdRegOrderDetail, AdoRs.Fields("SLITING_INFO").Value & "", .MaxRows, 6)
                Call SetText(spdRegOrderDetail, AdoRs.Fields("P_NO_F").Value & "", .MaxRows, 7)
                Call SetText(spdRegOrderDetail, AdoRs.Fields("P_NO_T").Value & "", .MaxRows, 8)
  
            End With
            AdoRs.MoveNext
        Loop
        AdoRs.Close
    End If

End Sub



Private Sub spdPrtReel_Click(ByVal Col As Long, ByVal Row As Long)
    Dim pAdoRS1      As ADODB.Recordset
    Dim pAdoRS2      As ADODB.Recordset
    Dim i               As Integer
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
    Dim strExDate       As String
    Dim strProdTemp     As String
    Dim strPrtLabelName As String
    Dim strBNum         As String
    
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
    txtOrderMemo.Text = GetText(spdPrtReel, Row, 8)
    strProdLen = GetText(spdPrtReel, Row, 10)
    strProdLen = strProdLen * 100 '미터를 cm으로 변환
    txtProdLen.Text = strProdLen
    'txtProdPosNo.Text = GetText(spdPrtReel, Row, 4)
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
              
    Call GetOrderDetail(Format(strDate, "yyyymmdd"), strProdCd, txtSlittingNo.Text)

    With spdPrtReelDetail
        .MaxRows = 0
    End With

    Set AdoRs = Get_LabelDetail(strProdLabelCd, "R")
            
    If AdoRs Is Nothing Then
        '등록된 정보 없음
    Else
        '-- 제품정보 찾아오기
        SQL = ""
        SQL = SQL & "SELECT             " & vbCrLf
        SQL = SQL & "  PROD_NAME        " & vbCrLf
        SQL = SQL & ", PROD_LENGTH      " & vbCrLf
        SQL = SQL & ", PROD_MATERIAL_CD " & vbCrLf
        SQL = SQL & ", EXPIR_MONTH      " & vbCrLf
        SQL = SQL & ", PROD_STOR_TEMP   " & vbCrLf
        SQL = SQL & ", PROD_SIZE        " & vbCrLf
        SQL = SQL & ", PROD_CHIMEI_PN   " & vbCrLf
        SQL = SQL & ", VENDER_CD        " & vbCrLf
        SQL = SQL & ", PROD_LINE_FA     " & vbCrLf
        SQL = SQL & ", PROD_SLIT_FA     " & vbCrLf
        SQL = SQL & ", PROD_CONTROL_YN  " & vbCrLf
        SQL = SQL & ", PROD_PCN_NO                          " & vbCrLf
        SQL = SQL & "  FROM LBL_M_PROD                      " & vbCrLf
        SQL = SQL & " WHERE PROD_CD  = '" & strProdCd & "'  " & vbCrLf
        SQL = SQL & "   AND COMP_CD  = '" & strCompCd & "'  " & vbCrLf
        SQL = SQL & "   AND USED_YN  = 'Y'                  " & vbCrLf
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
                    Case "P0011", "P0018", "P0019", "P0020"
                        '-- 바코드 등록 정보 찾아오기
                        If AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "바코드" Then
                            strBarData = ""
                            '============================== 바코드 정보 찾아오기 ==============================
                            Set pAdoRS1 = Get_BarDetail_Prt(strProdCd, strCompCd, "R")
                            If pAdoRS1 Is Nothing Then
                                '등록된 정보 없음
                            Else
                                'strBarData = ""
                                Do Until pAdoRS1.EOF
                                    strTmp = pAdoRS1.Fields("BAR_ITEM_NAME").Value & ""
                                    Select Case strTmp
                                        'Case "자재코드":       ' strBarData = strBarData
                                        Case "유효기간_년":     strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", Year(strExDate))
                                        Case "유효기간_월":     strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", MONTH(strExDate))
                                        Case "유효기간_일":     strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", Day(strExDate))
                                        Case "제조일_년":       strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", Year(strDate))
                                        Case "제조일_월":       strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", MONTH(strDate))
                                        Case "제조일_일":       strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", Day(strDate))
                                        Case "Slitting순번":    strBarData = strBarData & Format(txtSlittingNo.Text, "00")
                                        '-- 2020.07.23 이노룩스만 P에서 K로 바뀜
                                        Case "Product No":      strBarData = strBarData & "P" & Format(txtSlittingNo.Text, "0") & "01"
                                        Case "제품길이":        strBarData = strBarData & (strProdLen * 100)
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
                        'ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Name" Then
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Product" Then
                            Call SetText(spdPrtReelDetail, txtProdNm.Text, .MaxRows, 3)
                            strContents = txtProdNm.Text
                            Call SetText(spdPrtReelDetail, txtProdNm.Text, .MaxRows, 3)
                            strContents = txtProdNm.Text
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Size" Then
                            Call SetText(spdPrtReelDetail, strProdSize, .MaxRows, 3)
                            strContents = strProdSize
                        'ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Storage Temperature" Then
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Storage" Then
                            Call SetText(spdPrtReelDetail, strProdTemp, .MaxRows, 3)
                            strContents = strProdTemp
                        'ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Production Date" Then
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Prod. Date" Then
                            Call SetText(spdPrtReelDetail, Format(strDate, "yyyy" & strGu & "mm" & strGu & "dd"), .MaxRows, 3)
                            strContents = Format(strDate, "yyyy" & strGu & "mm" & strGu & "dd")
                        'ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Expiration Date" Then
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Exp. Date" Then
                            Call SetText(spdPrtReelDetail, Format(strExDate, "yyyy" & strGu & "mm" & strGu & "dd"), .MaxRows, 3)
                            strContents = Format(strExDate, "yyyy" & strGu & "mm" & strGu & "dd")
                        'ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "LotNo" Then
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "LotNo" Or AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "ACF Lot" Or AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Lot No." Then
                            If strLotNo = "" Then
                                strLotNo = GetLotNo(strDate, txtSlittingNo.Text, txtPackNm.Text, txtCompNm.Text)
                            End If
                            
                            'strLotNoFull = Space(1) & "(" & "P" & txtSlittingNo.Text & Format(txtP1From.Text, "00") & ")"
                            
                            '-- 2020.06.04 수정
                            'strLotNoFull = Space(1) & "(" & "P" & txtSlittingNo.Text & "01" & ")"
                            strLotNoFull = Space(1) & "(" & GetText(spdRegOrderDetail, 1, 7) & ")"
                            
                            strLotNo = strLotNo & strLotNoFull
                            
                            
                            Call SetText(spdPrtReelDetail, strLotNo, .MaxRows, 3)
                            strContents = strLotNo
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Material Code" Then
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
                    
                    '============== TP203C(ACF) ======================================================================
                    Case "P0001", "P0002"
                        '-- 바코드 등록 정보 찾아오기
                        If AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "바코드" Then
                            'strBarData = ""
                            '============================== 바코드 정보 찾아오기 ==============================
                            Set pAdoRS1 = Get_BarDetail_Prt(strProdCd, strCompCd, "R")
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
                                        Case "제조일_년":       strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", Year(strDate))
                                        Case "제조일_월":       strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", MONTH(strDate))
                                        Case "제조일_일":       strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", Day(strDate))
                                        Case "Slitting순번":    strBarData = strBarData & Format(txtSlittingNo.Text, "00")
                                        '-- 2020.07.23 이노룩스만 P에서 K로 바뀜
                                        'Case "Product No":      strBarData = strBarData & "P" & Format(txtSlittingNo.Text, "0") & "01"
                                        Case "Product No":      strBarData = strBarData & "K" & Format(txtSlittingNo.Text, "0") & "01"
                                        Case "제품길이":        strBarData = strBarData & (strProdLen * 100)
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
                        'ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "LotNo" Then
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "LotNo" Or AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "ACF Lot" Then
                            If strLotNo = "" Then
                                strLotNo = GetLotNo(strDate, txtSlittingNo.Text, txtPackNm.Text, txtCompNm.Text)
                            End If
                            
                            'strLotNoFull = Space(1) & "(" & "P" & txtSlittingNo.Text & Format(txtP1From.Text, "00") & ")"
                            
                            '-- 2020.06.04 수정
                            'strLotNoFull = Space(1) & "(" & "P" & txtSlittingNo.Text & "01" & ")"
                            strLotNoFull = Space(1) & "(" & GetText(spdRegOrderDetail, 1, 7) & ")"
                            
                            strLotNo = strLotNo & strLotNoFull
                            Call SetText(spdPrtReelDetail, strLotNo, .MaxRows, 3)
                            strContents = strLotNo
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "Material Code" Then
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
                                        

                    '============== CF-TP408A ======================================================================
                    Case "P0003"

                        '-- 바코드 등록 정보 찾아오기
                        If AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "바코드" Then
                            'strBarData = ""
                            '============================== 바코드 정보 찾아오기 ==============================
                            Set pAdoRS1 = Get_BarDetail_Prt(strProdCd, strCompCd, "R")
                        
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
                            'strLotNoFull = Space(1) & "(" & "P" & txtSlittingNo.Text & Format(txtP1From.Text, "00") & ")"
                            strLotNoFull = Space(1) & "(" & "P" & txtSlittingNo.Text & "01" & ")"
                            strLotNo = strLotNo & strLotNoFull
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
                            Set pAdoRS1 = Get_BarDetail_Prt(strProdCd, strCompCd, "R")
                        
                            If pAdoRS1 Is Nothing Then
                                '등록된 정보 없음
                            Else
                                'strBarData = ""
                                Do Until pAdoRS1.EOF
                                    strTmp = pAdoRS1.Fields("BAR_ITEM_NAME").Value & ""
                                    Select Case strTmp
                                        Case "자재코드"
                                        Case "Vendor코드":      strBarData = strBarData & strVendorCd
                                        Case "제조라인공장":    strBarData = strBarData & strProdLineFA
                                        Case "Sliting공장":     strBarData = strBarData & strProdSlitFA
                                        Case "관리선이탈여부":  strBarData = strBarData & strContYN
                                        Case "PCN차수":         strBarData = strBarData & strPcnNO
                                        Case "제조일_년":       strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", Year(strDate))
                                        Case "제조일_월":       strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", MONTH(strDate))
                                        Case "제조일_일":       strBarData = strBarData & Get_YMD(pAdoRS1.Fields("LABEL_ITEM_TYPE").Value & "", Day(strDate))
                                        Case "생산LOT":         strBarData = strBarData & Format(txtSlittingNo.Text, "0") & txtCompNm.Text & Format(txtSlittingNo.Text, "0") & "01"
                                        Case "REEL단위":        strBarData = strBarData & strProdLen
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
                            strLotNoFull = Space(1) & "(" & "P" & txtSlittingNo.Text & "01" & ")"
                            strLotNo = strLotNo & strLotNoFull
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
                            Set pAdoRS1 = Get_BarDetail_Prt(strProdCd, strCompCd, "R")
                        
                            If pAdoRS1 Is Nothing Then
                                '등록된 정보 없음
                            Else
                                'strBarData = ""
                                Do Until pAdoRS1.EOF
                                    strTmp = pAdoRS1.Fields("BAR_ITEM_NAME").Value & ""
                                    Select Case strTmp
                                        Case "ACF":         strBarData = "C"
                                        '생산LOT 는 삭제예정
                                        Case "생산LOT":     strBarData = strBarData & strLotNo
                                        Case "발행LOT":     strBarData = strBarData & strLotNo
                                        Case "_":           strBarData = strBarData & "_"
                                        Case "P/N":         strBarData = strBarData & "101"
                                        Case "유효기간":    strBarData = strBarData & Format(strExDate, "yyyymmdd")
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
                        ElseIf AdoRs.Fields("LABEL_ITEM_NAME").Value & "" = "PartsID" Then
                            strBarData = Mid(strBarData, 1, 4) & "-" & Mid(strBarData, 5)
                            Call SetText(spdPrtReelDetail, strBarData, .MaxRows, 3)
                            strContents = strBarData
                            
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
                            strLotNoFull = Space(1) & "(" & "P" & txtSlittingNo.Text & "01" & ")"
                            'strLotNo = strLotNo & strLotNoFull
                            Call SetText(spdPrtReelDetail, strLotNo & strLotNoFull, .MaxRows, 3)
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
                                                        
                                                        
                End Select
                
                AdoRs.MoveNext
            End With
        Loop
    End If
    
    AdoRs.Close

'    Call GetReelOrderList(strFromDt, strToDt, "", "", "R")

End Sub


