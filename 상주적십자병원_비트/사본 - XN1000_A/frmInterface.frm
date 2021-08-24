VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInterface 
   BorderStyle     =   1  '´ÜÀÏ °íÁ¤
   Caption         =   "XN1000_A"
   ClientHeight    =   11790
   ClientLeft      =   330
   ClientTop       =   825
   ClientWidth     =   20805
   BeginProperty Font 
      Name            =   "±¼¸²Ã¼"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInterface.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmInterface.frx":1272
   ScaleHeight     =   11790
   ScaleWidth      =   20805
   StartUpPosition =   1  '¼ÒÀ¯ÀÚ °¡¿îµ¥
   Begin VB.PictureBox Picture2 
      Align           =   1  'À§ ¸ÂÃã
      Height          =   810
      Left            =   0
      ScaleHeight     =   750
      ScaleWidth      =   20745
      TabIndex        =   37
      Top             =   525
      Width           =   20805
      Begin VB.OptionButton optSaveResult 
         Caption         =   "¼öÁ¤"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   17550
         TabIndex        =   54
         Top             =   180
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optSaveResult 
         Caption         =   "Àåºñ"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   16770
         TabIndex        =   53
         Top             =   180
         Width           =   735
      End
      Begin VB.Frame fraWork 
         Appearance      =   0  'Æò¸é
         BorderStyle     =   0  '¾øÀ½
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   60
         TabIndex        =   38
         Top             =   60
         Width           =   8055
         Begin VB.CheckBox chkSaveAll 
            Caption         =   "ÀúÀåÆ÷ÇÔ"
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   7350
            TabIndex        =   40
            Top             =   -30
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.ComboBox cboChk 
            Height          =   315
            ItemData        =   "frmInterface.frx":14F5
            Left            =   7290
            List            =   "frmInterface.frx":1502
            TabIndex        =   39
            Top             =   390
            Visible         =   0   'False
            Width           =   825
         End
         Begin MSComCtl2.DTPicker dtpStopDt 
            Height          =   345
            Left            =   2670
            TabIndex        =   41
            Top             =   150
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±¼¸²Ã¼"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   21430273
            CurrentDate     =   40248
         End
         Begin MSComCtl2.DTPicker dtpStartDt 
            Height          =   345
            Left            =   930
            TabIndex        =   42
            Top             =   150
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±¼¸²Ã¼"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   21430273
            CurrentDate     =   40248
         End
         Begin Threed.SSCommand cmdSearch 
            Height          =   525
            Left            =   4320
            TabIndex        =   52
            Top             =   60
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   926
            _Version        =   131072
            Font3D          =   3
            ForeColor       =   192
            PictureFrames   =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Picture         =   "frmInterface.frx":1518
            Caption         =   "È¯ÀÚÁ¶È¸    "
            ButtonStyle     =   2
            PictureAlignment=   4
         End
         Begin Threed.SSCommand cmdPatDelete 
            Height          =   525
            Left            =   5790
            TabIndex        =   51
            Top             =   60
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   926
            _Version        =   131072
            Font3D          =   3
            ForeColor       =   16512
            PictureFrames   =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Picture         =   "frmInterface.frx":1832
            Caption         =   "¼±ÅÃÁ¦¿Ü    "
            ButtonStyle     =   2
            PictureAlignment=   4
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Æò¸é
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Åõ¸í
            Caption         =   "Ã³¹æÀÏÀÚ"
            BeginProperty Font 
               Name            =   "±¼¸²"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   3
            Left            =   60
            TabIndex        =   44
            Top             =   240
            Width           =   780
         End
         Begin VB.Label Label12 
            Appearance      =   0  'Æò¸é
            BackColor       =   &H80000005&
            BackStyle       =   0  'Åõ¸í
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "±¼¸²"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2550
            TabIndex        =   43
            Top             =   240
            Width           =   105
         End
      End
      Begin VB.Label lblExamDate 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         Caption         =   "20160202"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   165
         Left            =   19260
         TabIndex        =   65
         Top             =   270
         Width           =   1005
      End
      Begin VB.Label lblSaveSeq 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         Caption         =   "99999"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   165
         Left            =   18450
         TabIndex        =   56
         Top             =   270
         Width           =   615
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Æò¸é
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Åõ¸í
         Caption         =   "°á°úÀû¿ë"
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   15900
         TabIndex        =   55
         Top             =   270
         Width           =   780
      End
      Begin Threed.SSCommand cmdIFTrans 
         Height          =   615
         Left            =   13350
         TabIndex        =   50
         Top             =   60
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   1085
         _Version        =   131072
         Font3D          =   1
         ForeColor       =   12583104
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmInterface.frx":1B4C
         Caption         =   "¼±ÅÃÀü¼Û    "
         PictureAlignment=   4
      End
      Begin Threed.SSCommand cmdIFClear 
         Height          =   615
         Left            =   11670
         TabIndex        =   49
         Top             =   60
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   1085
         _Version        =   131072
         Font3D          =   1
         ForeColor       =   16512
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmInterface.frx":1E66
         Caption         =   "È­¸éÁö¿ò    "
         PictureAlignment=   4
      End
      Begin Threed.SSCommand cmdExcelExport 
         Height          =   615
         Left            =   9990
         TabIndex        =   48
         Top             =   60
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   1085
         _Version        =   131072
         Font3D          =   3
         ForeColor       =   12582912
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmInterface.frx":2180
         Caption         =   "¿¢¼¿ÀúÀå    "
         PictureAlignment=   4
      End
      Begin Threed.SSCommand cmdRsltSearch 
         Height          =   615
         Left            =   8310
         TabIndex        =   47
         Top             =   60
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   1085
         _Version        =   131072
         Font3D          =   3
         ForeColor       =   192
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmInterface.frx":249A
         Caption         =   "°á°úÁ¶È¸    "
         PictureAlignment=   4
      End
   End
   Begin VB.Frame Frame1 
      Height          =   9945
      Left            =   90
      TabIndex        =   26
      Top             =   1380
      Width           =   15105
      Begin VB.CommandButton cmdSL 
         Caption         =   "¢º"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   495
      End
      Begin VB.CheckBox chkWAll 
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   690
         TabIndex        =   28
         Top             =   270
         Width           =   225
      End
      Begin FPSpread.vaSpread vasID 
         Height          =   9645
         Left            =   90
         TabIndex        =   30
         Top             =   210
         Width           =   8085
         _Version        =   393216
         _ExtentX        =   14261
         _ExtentY        =   17013
         _StockProps     =   64
         ColHeaderDisplay=   0
         ColsFrozen      =   1
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   17
         MaxRows         =   20
         MoveActiveOnFocus=   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "frmInterface.frx":27B4
         UserResize      =   2
      End
      Begin VB.Frame Frame6 
         Height          =   585
         Left            =   8280
         TabIndex        =   31
         Top             =   120
         Width           =   6675
         Begin VB.TextBox txtBarNum 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   12
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   1890
            TabIndex        =   46
            Top             =   150
            Visible         =   0   'False
            Width           =   1875
         End
         Begin VB.CommandButton cmdBarInput 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   45
            Top             =   180
            Width           =   285
         End
         Begin VB.Label Label3 
            BackColor       =   &H80000008&
            ForeColor       =   &H8000000E&
            Height          =   315
            Left            =   180
            TabIndex        =   36
            Top             =   720
            Width           =   1155
         End
         Begin VB.Label lblPname 
            Caption         =   "1234567890ab"
            Height          =   225
            Index           =   0
            Left            =   5130
            TabIndex        =   35
            Top             =   240
            Width           =   1305
         End
         Begin VB.Label Label6 
            Caption         =   "È¯ÀÚ¸í :"
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4080
            TabIndex        =   34
            Top             =   240
            Width           =   945
         End
         Begin VB.Label lblBarcode 
            Caption         =   "12345"
            Height          =   165
            Index           =   0
            Left            =   1905
            TabIndex        =   33
            Top             =   240
            Width           =   1485
         End
         Begin VB.Label Label8 
            Caption         =   "¹ÙÄÚµå¹øÈ£ :"
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   480
            TabIndex        =   32
            Top             =   240
            Width           =   1380
         End
      End
      Begin FPSpread.vaSpread vasRes 
         Height          =   9090
         Left            =   8280
         TabIndex        =   29
         Top             =   750
         Width           =   6675
         _Version        =   393216
         _ExtentX        =   11774
         _ExtentY        =   16034
         _StockProps     =   64
         BackColorStyle  =   1
         ColHeaderDisplay=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   7
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "frmInterface.frx":350D
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'À§ ¸ÂÃã
      BackColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   20745
      TabIndex        =   18
      Top             =   0
      Width           =   20805
      Begin MSComCtl2.DTPicker dtpToday 
         Height          =   315
         Left            =   1050
         TabIndex        =   24
         Top             =   90
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   21430272
         CurrentDate     =   40457
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Æò¸é
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Åõ¸í
         Caption         =   "Sysmex XN1000_A"
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   285
         Index           =   4
         Left            =   3780
         TabIndex        =   64
         Top             =   90
         Width           =   2565
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Æò¸é
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Åõ¸í
         Caption         =   "°Ë»çÀÏÀÚ"
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   25
         Top             =   150
         Width           =   780
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Æò¸é
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Åõ¸í
         Caption         =   "Sysmex XN1000_A"
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   285
         Index           =   2
         Left            =   3810
         TabIndex        =   22
         Top             =   120
         Width           =   2565
      End
      Begin VB.Image imgPort 
         Height          =   240
         Left            =   12180
         Picture         =   "frmInterface.frx":7217
         Top             =   90
         Width           =   240
      End
      Begin VB.Image imgSend 
         Height          =   240
         Left            =   13380
         Picture         =   "frmInterface.frx":77A1
         Top             =   90
         Width           =   240
      End
      Begin VB.Image imgReceive 
         Height          =   240
         Left            =   14670
         Picture         =   "frmInterface.frx":7D2B
         Top             =   90
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Åõ¸í
         Caption         =   "Port"
         Height          =   195
         Index           =   0
         Left            =   11640
         TabIndex        =   21
         Top             =   120
         Width           =   420
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Åõ¸í
         Caption         =   "Send"
         Height          =   195
         Left            =   12765
         TabIndex        =   20
         Top             =   120
         Width           =   420
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Åõ¸í
         Caption         =   "Receive"
         Height          =   195
         Left            =   13800
         TabIndex        =   19
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame FrmHideControl 
      Caption         =   "HideControl"
      Height          =   9795
      Left            =   15780
      TabIndex        =   1
      Top             =   1500
      Visible         =   0   'False
      Width           =   4815
      Begin VB.TextBox txtTest 
         Height          =   375
         Left            =   240
         TabIndex        =   63
         Top             =   9270
         Width           =   3225
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Àü¼ÛÅ×½ºÆ®"
         Height          =   435
         Left            =   3510
         TabIndex        =   62
         Top             =   9240
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         Caption         =   "Error Log"
         Height          =   1365
         Left            =   180
         TabIndex        =   60
         Top             =   7740
         Width           =   4530
         Begin VB.TextBox txtErrLog 
            Appearance      =   0  'Æò¸é
            Height          =   1035
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  '¼öÁ÷
            TabIndex        =   61
            Top             =   240
            Width           =   4275
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Print"
         Height          =   2415
         Left            =   180
         TabIndex        =   57
         Top             =   5250
         Width           =   3045
         Begin FPSpread.vaSpread vasPrint 
            Height          =   1035
            Left            =   120
            TabIndex        =   58
            Top             =   1290
            Width           =   2760
            _Version        =   393216
            _ExtentX        =   4868
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
            SpreadDesigner  =   "frmInterface.frx":82B5
         End
         Begin FPSpread.vaSpread vasPrintBuf 
            Height          =   975
            Left            =   120
            TabIndex        =   59
            Top             =   240
            Width           =   2715
            _Version        =   393216
            _ExtentX        =   4789
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
            SpreadDesigner  =   "frmInterface.frx":9D2E
         End
      End
      Begin VB.CheckBox chkBar 
         Caption         =   "BARCODE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   465
         Left            =   3090
         Style           =   1  '±×·¡ÇÈ
         TabIndex        =   23
         Top             =   3210
         Value           =   1  'È®ÀÎ
         Width           =   1065
      End
      Begin FPSpread.vaSpread vasCode 
         Height          =   945
         Left            =   120
         TabIndex        =   15
         Top             =   2250
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
         SpreadDesigner  =   "frmInterface.frx":9F46
      End
      Begin FPSpread.vaSpread vasTemp1 
         Height          =   945
         Left            =   1860
         TabIndex        =   2
         Top             =   1290
         Width           =   2535
         _Version        =   393216
         _ExtentX        =   4471
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
         SpreadDesigner  =   "frmInterface.frx":A15E
      End
      Begin VB.PictureBox picLogin 
         Appearance      =   0  'Æò¸é
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1530
         Picture         =   "frmInterface.frx":A376
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   16
         Top             =   4710
         Width           =   285
      End
      Begin VB.CommandButton lblclear 
         Caption         =   "lblclear"
         Height          =   495
         Left            =   180
         TabIndex        =   14
         Top             =   4560
         Width           =   1215
      End
      Begin VB.TextBox txtData 
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  '¼öÁ÷
         TabIndex        =   8
         Top             =   3240
         Width           =   1665
      End
      Begin VB.TextBox txtTemp 
         Height          =   435
         Left            =   2730
         TabIndex        =   7
         Top             =   3690
         Width           =   645
      End
      Begin VB.TextBox Text_ini 
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2070
         TabIndex        =   6
         Top             =   3705
         Width           =   645
      End
      Begin VB.TextBox txtErr 
         ForeColor       =   &H000000FF&
         Height          =   585
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  '¾ç¹æÇâ
         TabIndex        =   5
         Top             =   3840
         Width           =   1635
      End
      Begin VB.CheckBox chkMode 
         Caption         =   "AUTO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   465
         Left            =   1980
         Style           =   1  '±×·¡ÇÈ
         TabIndex        =   4
         Top             =   3180
         Value           =   1  'È®ÀÎ
         Width           =   1065
      End
      Begin VB.Frame FrmUseControl 
         Caption         =   "UseControl"
         Height          =   870
         Left            =   1860
         TabIndex        =   3
         Top             =   2310
         Width           =   2835
         Begin VB.Timer tmrSend 
            Enabled         =   0   'False
            Interval        =   100
            Left            =   2220
            Top             =   300
         End
         Begin VB.Timer tmrReceive 
            Enabled         =   0   'False
            Interval        =   100
            Left            =   1740
            Top             =   300
         End
         Begin MSCommLib.MSComm comEqp 
            Left            =   90
            Top             =   210
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            DTREnable       =   -1  'True
            RThreshold      =   1
            RTSEnable       =   -1  'True
            EOFEnable       =   -1  'True
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   720
            Top             =   270
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
                  Picture         =   "frmInterface.frx":A900
                  Key             =   "RUN"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":AE9A
                  Key             =   "NOT"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":B434
                  Key             =   "STOP"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":B9CE
                  Key             =   "LST"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":C260
                  Key             =   "ITM"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":C3BA
                  Key             =   "ERR"
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":C514
                  Key             =   "NOF"
               EndProperty
            EndProperty
         End
      End
      Begin FPSpread.vaSpread vasList 
         Height          =   975
         Left            =   120
         TabIndex        =   9
         Top             =   270
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
         SpreadDesigner  =   "frmInterface.frx":C66E
      End
      Begin FPSpread.vaSpread vasResTemp 
         Height          =   1035
         Left            =   1860
         TabIndex        =   10
         Top             =   240
         Width           =   2505
         _Version        =   393216
         _ExtentX        =   4419
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
         SpreadDesigner  =   "frmInterface.frx":C886
      End
      Begin FPSpread.vaSpread vasTemp 
         Height          =   975
         Left            =   120
         TabIndex        =   11
         Top             =   1260
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
         SpreadDesigner  =   "frmInterface.frx":CA9E
      End
      Begin VB.Label lblUser 
         BackStyle       =   0  'Åõ¸í
         BorderStyle     =   1  '´ÜÀÏ °íÁ¤
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1980
         TabIndex        =   17
         Top             =   4680
         Width           =   825
      End
      Begin VB.Label lblChangeBar 
         BackColor       =   &H000000FF&
         Height          =   405
         Left            =   2880
         TabIndex        =   13
         Top             =   4650
         Width           =   465
      End
      Begin VB.Label lblChangePID 
         BackColor       =   &H000000FF&
         Height          =   405
         Left            =   3390
         TabIndex        =   12
         Top             =   4650
         Width           =   435
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '¾Æ·¡ ¸ÂÃã
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   11385
      Width           =   20805
      _ExtentX        =   36698
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   8202
            MinWidth        =   8202
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   8114
            MinWidth        =   8114
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   6068
            MinWidth        =   6068
            TextSave        =   "2016-02-05"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   4304
            MinWidth        =   4304
            TextSave        =   "¿ÀÀü 2:34"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²Ã¼"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu MnMain 
      Caption         =   "Main"
      Begin VB.Menu MnExit 
         Caption         =   "Á¾·á"
      End
   End
   Begin VB.Menu MnConfig 
      Caption         =   "Setting"
      Begin VB.Menu MnTConfig 
         Caption         =   "Åë½Å¼³Á¤"
      End
      Begin VB.Menu MnExamConfig 
         Caption         =   "ÄÚµå¼³Á¤"
      End
   End
   Begin VB.Menu MnTrans 
      Caption         =   "Send"
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

Const colSpecNo = 0     '¹Ì»ç¿ë
Const colCheckBox = 1
Const colSAVESEQ = 2    'ÀúÀå¼ø¹ø(³¯Â¥º°)
Const colEXAMDATE = 3   '°Ë»çÀÏÀÚ
Const colHOSPDATE = 4   'º´¿øÁ¢¼öÀÏÀÚ
Const colBARCODE = 5
Const colCHARTNO = 6
Const colPID = 7        'º´·Ï¹øÈ£(³»¿ø¹øÈ£)
Const colINOUT = 8      'ÀÔ¿ø/¿Ü·¡
Const colDISKNO = 9
Const colPOSNO = 10
Const colPNAME = 11
Const colPSEX = 12
Const colPAGE = 13
Const colOCNT = 14
Const colRCNT = 15
Const colState = 16

'sendflag
'0: Order
'1: Result
'2: Trans
'vasres, vasrres colum
Const colEQUIPCODE = 1
Const colEXAMCODE = 2
Const colEXAMNAME = 3
Const colMachResult = 4
Const colRESULT = 5
Const colSeq = 6
Const colFLAG = 7
Const colSubCode = 8

Dim gRow As Long

Dim gsBarCode       As String
Dim gsSampleType    As String
Dim gsPID           As String
Dim gsRackNo        As String
Dim gsPosNo         As String
Dim gsResDateTime   As String
Dim gsSeqNo         As String
Dim gsExamCode      As String
Dim gsExamName      As String
Dim gsOrder         As String
Dim gsResult        As String
Dim gsFlag          As String

Dim gMT             As String
Dim gComState       As Long
Dim gErrState       As Long

Dim strBuffer       As String
Dim strORQN         As String


'===============================
Const SPCLEN As Integer = 10

Const STX As String = ""
Const ETX As String = ""
Const ENQ As String = ""
Const ACK As String = ""
Const NAK As String = ""
Const EOT As String = ""
Const ETB As String = ""
Const FS  As String = ""
Const RS  As String = ""
Const GS  As String = ""


Dim strRecvData()   As String
Dim intPhase        As Integer
Dim strState        As String
Dim intBufCnt       As Integer
Dim blnIsETB        As Boolean
Dim intSndPhase     As Integer
Dim intFrameNo      As Integer
'===============================


Private Sub chkMode_Click()
    If chkMode.Value = 1 Then
        chkMode.Caption = "Auto"
    Else
        chkMode.Caption = "Manual"
    End If
End Sub

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

Private Sub cmdBarInput_Click()
    If cmdBarInput.Caption = "+" Then
        cmdBarInput.Caption = "-"
        txtBarNum.Visible = True
        txtBarNum.SetFocus
    Else
        cmdBarInput.Caption = "+"
        txtBarNum.Visible = False
    End If
End Sub


Sub SaveExcel(Filename As String, argSpread As vaSpread)

On Error Resume Next

' Excel Object Library ¿Í ¿¬°áÇÕ´Ï´Ù.
Dim xlapp As Excel.Application
Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet

Dim iRow As Integer
Dim iCol As Integer
Dim i As Integer

    Set xlapp = CreateObject("Excel.Application")
    
    xlapp.DisplayAlerts = False
    
    Set xlBook = xlapp.Workbooks.Add
    
    Set xlSheet = xlBook.Worksheets(1)
     
    For iRow = 0 To argSpread.DataRowCnt
        For iCol = 1 To argSpread.DataColCnt
            argSpread.Row = iRow
            argSpread.Col = iCol
            xlSheet.Cells(iRow + 1, iCol) = argSpread.Text
        Next iCol
    Next iRow
    
    xlBook.SaveAs (Filename)
    xlapp.Quit


End Sub

Private Sub cmdExcelExport_Click()
    Dim iRow As Integer
    Dim j As Integer
    
    Dim sCurDate As String
    Dim sSerDate As String
    Dim sHead As String
    Dim sFoot As String
    Dim sFileName As String
    
    Dim sA1c As String
    Dim sIFCC As String
    Dim seAg As String
    
    ClearSpread vasPrint

    j = 1

    For iRow = 1 To vasID.DataRowCnt
        vasID.Row = iRow
        vasID.Col = 1

        If vasID.Value = 1 Then
            SetText vasPrint, Trim(GetText(vasID, iRow, colBARCODE)), j, 1
            SetText vasPrint, Trim(GetText(vasID, iRow, colPNAME)), j, 3
            
                  SQL = " SELECT RESULT " & vbCrLf
            SQL = SQL & "   FROM PATRESULT " & vbCrLf
            SQL = SQL & "  WHERE MID(EXAMDATE,1,8) Between '" & Format(dtpStartDt, "YYYYMMDD") & "' AND '" & Format(dtpStopDt, "YYYYMMDD") & "'" & vbCrLf
            SQL = SQL & "    AND EQUIPNO = '" & gEquip & "' " & vbCrLf
            SQL = SQL & "    AND SENDFLAG IN ('0','1','2') " & vbCrLf
            SQL = SQL & "  ORDER BY SEQNO * 10 "
            
            
           ' SQL = " SELECT RESULT " & vbCrLf & _
                  "   FROM PATRESULT " & vbCrLf & _
                  "  WHERE EXAMDATE = '" & Format(dtpExamDate, "YYYYMMDD") & "' " & vbCrLf & _
                  "    AND EQUIPNO  = '" & gEquip & "' " & vbCrLf & _
                  "    AND BARCODE  = '" & Trim(GetText(vasID, iRow, colBARCODE)) & "' " & vbCrLf & _
                  "    AND PID      = '" & Trim(GetText(vasPrint, iRow, colBARCODE)) & "' " & vbCrLf & _
                  "  ORDER BY SEQNO"
            Res = GetDBSelectVas(gLocal, SQL, vasPrintBuf)
            
            ClearSpread vasPrintBuf, 1, 1

            j = j + 1
        End If
    Next iRow
    
    If vasPrint.DataRowCnt < 1 Then
        MsgBox "ÀúÀåÇÒ ÀÚ·á°¡ ¾ø½À´Ï´Ù.", , "¾Ë ¸²"
        Exit Sub
    Else
        CommonDialog1.Filter = "Excel Files (*.xls)|*.xls|All Files (*.*)|*.*"
        CommonDialog1.ShowSave
        sFileName = CommonDialog1.Filename
        SaveExcel sFileName, vasPrint
        
    End If
End Sub

Private Sub cmdIFClear_Click()
    Dim i As Integer

    Var_Clear
    
    txtData.Text = ""
    txtErr.Text = ""
    
    SetForeColor vasID, 1, vasID.MaxRows, 1, vasID.MaxCols, 0, 0, 0
    SetForeColor vasRes, 1, vasRes.MaxRows, 1, vasRes.MaxCols, 0, 0, 0
    
    vasID.MaxRows = 0
    vasRes.MaxRows = 0
    
    gRow = 0
    
End Sub

Private Sub cmdIFTrans_Click()
    Dim lRow As Long
    
    For lRow = 1 To vasID.DataRowCnt
        vasID.Row = lRow
        vasID.Col = 1
        If vasID.Value = 1 Then
            
            Res = SaveTransDataW(lRow)
        
            If Res = -1 Then
                SetForeColor vasID, lRow, lRow, 1, colState, 255, 0, 0
                SetText vasID, "Failed", lRow, colState
            Else
                vasID.Row = lRow
                vasID.Col = 1
                vasID.Value = 1
                
                SetBackColor vasID, lRow, lRow, 1, colState, 202, 255, 112
                SetText vasID, "Trans", lRow, colState
                
                SQL = " UPDATE PATRESULT SET " & vbCrLf & _
                      " SENDFLAG = '2' " & vbCrLf & _
                      " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf & _
                      " AND BARCODE = '" & Trim(GetText(vasID, lRow, colBARCODE)) & "' "
                Res = SendQuery(gLocal, SQL)
                If Res = -1 Then
                    SaveQuery SQL
                    Exit Sub
                End If
                
            End If
            vasID.Row = lRow
            vasID.Col = 1
            vasID.Value = 0
        End If
    Next lRow
End Sub

Private Sub cmdPatDelete_Click()
    Dim i As Integer
    Dim j As Integer
    
    j = 0
    With vasID
        For i = .DataRowCnt To 1 Step -1
            .Row = i
            .Col = colCheckBox
            If .Value = "1" Then
                .Action = ActionDeleteRow
                .MaxRows = .MaxRows - 1
                j = j + 1
            End If
        Next
    End With
    
End Sub


Private Sub cmdRsltSearch_Click()
    Dim iRow As Long
    Dim strDate As String
    Dim strSaveSeq As String
    Dim strChart As String
    Dim RS          As ADODB.Recordset
    Dim i As Integer
    Dim blnSame As Boolean
    Dim intCol As Integer
    
    
    ClearSpread vasID
    ClearSpread vasRes

    vasID.MaxRows = 0
          
          SQL = " SELECT '', SAVESEQ, MID(EXAMDATE,1,8) AS EXAMDATE, HOSPDATE AS Á¢¼öÀÏÀÚ, BARCODE AS ¹ÙÄÚµå¹øÈ£, CHARTNO AS Â÷Æ®¹øÈ£, PID AS ³»¿ø¹øÈ£, PNAME AS ÀÌ¸§,PSEX AS ¼ºº°, PAGE AS ³ªÀÌ, DISKNO, POSNO, EXAMCODE, RESULT, SENDFLAG " & vbCrLf
    SQL = SQL & "   FROM PATRESULT " & vbCrLf
    SQL = SQL & "  WHERE MID(EXAMDATE,1,8) = '" & Format(dtpToday, "YYYYMMDD") & "'" & vbCrLf
    SQL = SQL & "    AND EQUIPNO = '" & gEquip & "' " & vbCrLf
    SQL = SQL & "    AND SENDFLAG IN ('0','1','2') " & vbCrLf
    SQL = SQL & " ORDER BY SAVESEQ,BARCODE "
    
    Set RS = cn.Execute(SQL, , 1)

    If Not RS.EOF = True And Not RS.BOF = True Then
        Do Until RS.EOF
            With vasID
                For i = 1 To .DataRowCnt
                    strDate = GetText(vasID, i, colHOSPDATE)
                    strChart = GetText(vasID, i, colBARCODE)
                    strSaveSeq = GetText(vasID, i, colSAVESEQ)
                    
                    If Trim(RS("Á¢¼öÀÏÀÚ")) = strDate And Trim(RS("SAVESEQ")) = strSaveSeq And Trim(RS("¹ÙÄÚµå¹øÈ£")) = strChart Then
                        blnSame = True
                    End If
                    
                    For intCol = colState + 1 To vasID.MaxCols
                        If Trim(RS.Fields("EXAMCODE")) = gArrEquip(intCol - colState, 3) Then
                            'vasID.Row = .MaxRows
                            'vasID.Col = intCol
                            'vasID.BackColor = vbYellow
                            SetText vasID, Trim(RS.Fields("RESULT")) & "", .MaxRows, intCol
                            Exit For
                        End If
                    Next
                Next

                If blnSame = False Then
                    .MaxRows = .MaxRows + 1

                    SetText vasID, "1", .MaxRows, colCheckBox
                    SetText vasID, Trim(RS.Fields("SAVESEQ")) & "", .MaxRows, colSAVESEQ
                    SetText vasID, Trim(RS.Fields("EXAMDATE")) & "", .MaxRows, colEXAMDATE
                    SetText vasID, Trim(RS.Fields("Á¢¼öÀÏÀÚ")) & "", .MaxRows, colHOSPDATE
                    SetText vasID, Trim(RS.Fields("¹ÙÄÚµå¹øÈ£")) & "", .MaxRows, colBARCODE
                    SetText vasID, Trim(RS.Fields("Â÷Æ®¹øÈ£")) & "", .MaxRows, colCHARTNO
                    SetText vasID, Trim(RS.Fields("³»¿ø¹øÈ£")) & "", .MaxRows, colPID
                    SetText vasID, Trim(RS.Fields("ÀÌ¸§")) & "", .MaxRows, colPNAME
                    SetText vasID, Trim(RS.Fields("¼ºº°")) & "", .MaxRows, colPSEX
                    SetText vasID, Trim(RS.Fields("³ªÀÌ")) & "", .MaxRows, colPAGE
                    SetText vasID, Trim(RS.Fields("DISKNO")) & "", .MaxRows, colDISKNO
                    SetText vasID, Trim(RS.Fields("POSNO")) & "", .MaxRows, colPOSNO
                    
                    For intCol = colState + 1 To vasID.MaxCols
                        If Trim(RS.Fields("EXAMCODE")) = gArrEquip(intCol - colState, 3) Then
                            'vasID.Row = .MaxRows
                            'vasID.Col = intCol
                            'vasID.BackColor = vbYellow
                            SetText vasID, Trim(RS.Fields("RESULT")) & "", .MaxRows, intCol
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
    
    vasID.RowHeight(-1) = 12
    
End Sub

Private Sub GetWorkList(ByVal pFrDt As String, ByVal pToDt As String, Optional pBarNo As String)
    Dim i           As Integer
    Dim j           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strExamCode As String
    Dim RS          As ADODB.Recordset
    Dim sSpecNo     As String
    Dim buff        As String
    Dim strTestNm   As String
    Dim strDate     As String
    Dim strChart    As String
    Dim blnSame     As Boolean
    Dim intCol      As Integer
    
    If pBarNo = "" Then
        vasID.MaxRows = 0
        intRow = 0
    End If
    
    blnSame = False
    
          SQL = " SELECT DISTINCT '1'"
    SQL = SQL & " ,'' AS SN"
    SQL = SQL & " ,'' AS °á°úÀÏ½Ã"
    SQL = SQL & ", REQ_DT AS Á¢¼öÀÏÀÚ"
    SQL = SQL & ", QC_BAR_NO AS ¹ÙÄÚµå¹øÈ£"
    SQL = SQL & ", LOT_NO AS Â÷Æ®¹øÈ£"
    SQL = SQL & ", REQ_SEQ AS ³»¿ø¹øÈ£"
    SQL = SQL & ", 'ÀÔ¿ø' AS ÀÔ¿Ü"
    SQL = SQL & ", '' AS R"
    SQL = SQL & ", '' AS P"
    SQL = SQL & ", 'È«±æµ¿' AS ÀÌ¸§"
    SQL = SQL & ", '³²ÀÚ' AS ¼ºº°"
    SQL = SQL & ", REQ_SEQ AS ³ªÀÌ" & vbCrLf
    SQL = SQL & ", ITEM_CD " & vbCrLf
    SQL = SQL & "  FROM S2QCS101 " & vbCrLf
    SQL = SQL & " WHERE 1=1 " & vbCrLf
    If pBarNo <> "" Then
        SQL = SQL & "   AND QC_BAR_NO = '" & pBarNo & "'" & vbCrLf
    Else
        SQL = SQL & "   AND REQ_DT BETWEEN '" & pFrDt & "' AND '" & pToDt & "'" & vbCrLf
    End If
    'SQL = SQL & "   AND ITEM_CD IN (" & gAllExam & ")" & vbCrLf
    SQL = SQL & " ORDER BY Á¢¼öÀÏÀÚ, ¹ÙÄÚµå¹øÈ£, Â÷Æ®¹øÈ£, ³»¿ø¹øÈ£"
    
'    If pBarNo <> "" Then
'        Res = GetDBSelectVas(gServer, SQL, vasID, vasID.MaxRows + 1)
'    Else
'        Res = GetDBSelectVas(gServer, SQL, vasID)
'    End If
    
    If Res > 0 Then
        chkWAll.Value = "1"
    Else
        chkWAll.Value = "0"
    End If

    Set RS = cn_Ser.Execute(SQL, , 1)

    If Not RS.EOF = True And Not RS.BOF = True Then
        Do Until RS.EOF
            With vasID
                For i = 1 To .DataRowCnt
                    strDate = GetText(vasID, i, colHOSPDATE)
                    strChart = GetText(vasID, i, colBARCODE)

                    If Trim(RS("Á¢¼öÀÏÀÚ")) = strDate And Trim(RS("¹ÙÄÚµå¹øÈ£")) = strChart Then
                        blnSame = True
                    End If
                    
                    For intCol = colState + 1 To vasID.MaxCols
                        If Trim(RS.Fields("ITEM_CD")) = gArrEquip(intCol - colState, 3) Then
                            vasID.Row = .MaxRows
                            vasID.Col = intCol
                            vasID.BackColor = vbYellow
                            Exit For
                        End If
                    Next
                Next

                If blnSame = False Then
                    .MaxRows = .MaxRows + 1

                    SetText vasID, "1", .MaxRows, colCheckBox
                    SetText vasID, Trim(RS.Fields("Á¢¼öÀÏÀÚ")) & "", .MaxRows, colHOSPDATE
                    SetText vasID, Trim(RS.Fields("¹ÙÄÚµå¹øÈ£")) & "", .MaxRows, colBARCODE
                    SetText vasID, Trim(RS.Fields("Â÷Æ®¹øÈ£")) & "", .MaxRows, colCHARTNO
                    SetText vasID, Trim(RS.Fields("³»¿ø¹øÈ£")) & "", .MaxRows, colPID
                    SetText vasID, Trim(RS.Fields("ÀÌ¸§")) & "", .MaxRows, colPNAME
                    SetText vasID, Trim(RS.Fields("¼ºº°")) & "", .MaxRows, colPSEX
                    SetText vasID, Trim(RS.Fields("³ªÀÌ")) & "", .MaxRows, colPAGE
                    
                    For intCol = colState + 1 To vasID.MaxCols
                        If Trim(RS.Fields("ITEM_CD")) = gArrEquip(intCol - colState, 3) Then
                            vasID.Row = .MaxRows
                            vasID.Col = intCol
                            vasID.BackColor = vbYellow
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
    
    vasID.RowHeight(-1) = 12
    
End Sub

Private Sub cmdSearch_Click()
                
    Call GetWorkList(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))
    
    vasID.RowHeight(-1) = 12

End Sub

Private Sub cmdWorkDelete_Click()
    Dim i As Integer
    Dim j As Integer
    
    j = 0
    With vasID
        For i = .DataRowCnt To 1 Step -1
            .Row = i
            .Col = colCheckBox
            If .Value = "1" Then
                .Action = ActionDeleteRow
                .MaxRows = .MaxRows - 1
                j = j + 1
            End If
        Next
    End With
End Sub

Private Sub cmdSL_Click()
    If cmdSL.Caption = "¢º" Then
        cmdSL.Caption = "¢¸"
        vasID.Width = 14865
    Else
        cmdSL.Caption = "¢º"
        vasID.Width = 8085
    End If

End Sub

Private Sub imgPort_DblClick()
    
    '-- °³¹ß½Ã¿¡¸¸ Remark Ç®¾î¼­ Å×½ºÆ®ÁøÇà
    If FrmHideControl.Visible = True Then
        Me.Width = 15435
        FrmHideControl.Visible = False
    Else
        Me.Width = 25000
        FrmHideControl.Visible = True
    End If

End Sub



Private Sub lblclear_Click()
    lblChangePID.Caption = ""
    lblChangeBar.Caption = ""
    lblBarcode(0).Caption = ""
    lblPname(0).Caption = ""
    lblSaveSeq.Caption = ""
    lblExamDate.Caption = ""
End Sub

Private Sub Command16_Click()
    Dim i As Long
    Dim lsChar As String
        
    strBuffer = strBuffer & "00762" & vbCr
    strBuffer = strBuffer & "ÿ RESULT  " & vbCr
    strBuffer = strBuffer & "p 75" & vbCr
    strBuffer = strBuffer & "q 07/09/07 10h16mn07s" & vbCr
    strBuffer = strBuffer & "u 0000000000000001" & vbCr
    strBuffer = strBuffer & "s 0002" & vbCr
    strBuffer = strBuffer & "v                               " & vbCr
    strBuffer = strBuffer & "t R" & vbCr
    strBuffer = strBuffer & "€ D" & vbCr
    strBuffer = strBuffer & "! 007.9  " & vbCr
    strBuffer = strBuffer & "2 04.76  " & vbCr
    strBuffer = strBuffer & "3 014.0  " & vbCr
    strBuffer = strBuffer & "4 040.4  " & vbCr
    strBuffer = strBuffer & "5 084.9  " & vbCr
    strBuffer = strBuffer & "6 029.5  " & vbCr
    strBuffer = strBuffer & "7 034.7  " & vbCr
    strBuffer = strBuffer & "8 014.5  " & vbCr
    strBuffer = strBuffer & "@ 00263  " & vbCr
    strBuffer = strBuffer & "A 006.7  " & vbCr
    strBuffer = strBuffer & "B  .175  " & vbCr
    strBuffer = strBuffer & "C 014.1  " & vbCr
    strBuffer = strBuffer & "# 024.6  " & vbCr
    strBuffer = strBuffer & "% 009.7  " & vbCr
    strBuffer = strBuffer & "' 065.7  " & vbCr
    strBuffer = strBuffer & Chr(22) & " 001.9" & vbCr
    strBuffer = strBuffer & "$ 000.7  " & vbCr
    strBuffer = strBuffer & "& 005.3  " & vbCr
'    strBuffer = strBuffer & "W           *.4027Faˆ²ÔðÿûêË°–†yiZQJFBB@DHKMJHBB?=94555222.****)))))))))*.000002799;@JOUUX\ejrwy{€†ˆ†Š“‘”“‘‘‘‹ˆ†ˆ?tpig_\XZ\ "
'    strBuffer=strBuffer & "X             !! !!!!!!""%&+-7>Pbw‘§ÄÌâõÿóñêÎÉ¶¦š?yl_WTICA=732.,+(++('''*%'''))''&&&%$&&#$$#"""""!"!!!       !                  '
'    strBuffer=strBuffer & "Y           #);KVW^aqs€|†…?†‚?uufciaaPLDGDC;;6610,-.210,,))&'''&&#$"$&&&$$""&&$! !!#"" !!! !""! !"!!!!""#!!!! """   !""""""$""!!
    strBuffer = strBuffer & "S       " & vbCr
    strBuffer = strBuffer & "_ 105" & vbCr
    strBuffer = strBuffer & "P           G3" & vbCr
    strBuffer = strBuffer & "] 000 000 000 027 039" & vbCr
    strBuffer = strBuffer & "?LC 250  " & vbCr
    strBuffer = strBuffer & "?V2.8 " & vbCr
    strBuffer = strBuffer & "?A75C" & vbCr
    strBuffer = strBuffer & "   " & vbCr

    strBuffer = STX & ";" & ETX & vbCr
    
    strBuffer = ":N1    80 81                 00620141422      15 1   7.0  2   4.1  3   0.5  4   4.5  5    34  6    20  7   417  8   239  9    97 14    85 15    14 16   0.7 18    93 19      T54     1 "
    
    Call comEqp_OnComm
        

End Sub

Private Sub Form_Load()
    Dim sDate As String
    Dim i As Integer
    
    If App.PrevInstance Then
        End
    End If
    
    Me.Left = 0
    Me.Top = 0
    
    'Me.Height = 11520
    Me.Width = 15435
    
    imgPort.Picture = imlStatus.ListImages("NOT").ExtractIcon
    imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
    imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
    
    cmdIFClear_Click
    lblclear_Click
    
    GetSetup
    
    If gSave = "True" Then
        chkMode.Caption = "Auto"
        MnTransAuto.Checked = True
        MnTransManual.Checked = False
        chkMode.Value = 1
    Else
        chkMode.Caption = "Manual"
        MnTransAuto.Checked = False
        MnTransManual.Checked = True
        chkMode.Value = 0
    End If
    
    If gIFMode = "Barcode" Then
        'fraBar.Visible = True
        fraWork.Visible = False
    
        chkMode.Caption = "Barcode"
        MnModeBarcode.Checked = True
        MnModeWorkList.Checked = False
        chkBar.Value = 1
    Else
        'fraBar.Visible = False
        fraWork.Visible = True
    
        chkMode.Caption = "WorkList"
        MnModeBarcode.Checked = False
        MnModeWorkList.Checked = True
        chkBar.Value = 0
    End If
    
'    If gScreen = "ÅëÇÕ" Then
'        cmdSL.Caption = "¢¸"
'        vasID.Width = 14595
'    Else
'        cmdSL.Caption = "¢º"
'        vasID.Width = 7725
'    End If
    
    frmInterface.StatusBar1.Panels(1).Text = gUserID
        
    cboChk.ListIndex = 0
    
    comEqp.CommPort = gSetup.gPort
    comEqp.RTSEnable = gSetup.gRTSEnable
    comEqp.DTREnable = gSetup.gDTREnable
    comEqp.Settings = gSetup.gSpeed & "," & gSetup.gParity & "," & gSetup.gDataBit & "," & gSetup.gStopBit

    If comEqp.PortOpen = False Then
        comEqp.PortOpen = True
    End If
    
    If comEqp.PortOpen Then
        frmInterface.StatusBar1.Panels(2).Text = "¿¬°á µÇ¾ú½À´Ï´Ù"
        imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
        imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
        imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
'        lblStatus = "ÀÛ¾÷Áß.."
    Else
        frmInterface.StatusBar1.Panels(2).Text = "¿¬°á µÇÁö ¾Ê¾Ò½À´Ï´Ù"
        imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
        imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
        imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
'        lblStatus = "ÀÛ¾÷ ´ë±âÁß.."
    End If

    If Not Connect_Local Then
        MsgBox "¿¬°áµÇÁö ¾Ê¾Ò½À´Ï´Ù."
        cn_Local_Flag = False
        Exit Sub
    Else
        cn_Local_Flag = True
    End If
    
    '-- osw Ãß°¡
    For i = 1 To 1
        If Not Connect_PRServer Then
            MsgBox "¿¬°áµÇÁö ¾Ê¾Ò½À´Ï´Ù."
            cn_Server_Flag = False
            Exit Sub
        Else
            cn_Server_Flag = True
        End If
    Next
    
    '-- osw Ãß°¡
'    For i = 1 To 1
'        If Not Connect_DRServer Then
'            MsgBox "¿¬°áµÇÁö ¾Ê¾Ò½À´Ï´Ù."
'            cn_Server_Flag = False
'            Exit Sub
'        Else
'            cn_Server_Flag = True
'        End If
'    Next
    
    GetExamCode
    
    SetExamCode
    
    dtpToday = Date
    dtpStartDt = Date
    dtpStopDt = Date
    
    sDate = Format(DateAdd("y", CDate(dtpToday.Value), -30), "yyyymmdd")
    
    SQL = "delete from PATRESULT where examdate < '" & sDate & "'"
    Res = SendQuery(gLocal, SQL)
    
    lblUser.Caption = gUserID
    
    If lblUser.Caption = "" Then
        Call picLogin_Click
    End If
    
'    stInterface.Tab = 0

    '==============================
    intPhase = 1
    strState = ""
    intBufCnt = 0
    blnIsETB = False
    intSndPhase = 0
    intFrameNo = 1
    '==============================
    
    '-- test
'    vasID.MaxRows = 10
    
End Sub

Private Sub SetExamCode()
    Dim i As Integer
    
    
    With vasID
        .MaxCols = colState + UBound(gArrEquip)
        For i = 0 To UBound(gArrEquip) - 1
            Call SetText(vasID, gArrEquip(i + 1, 2), 0, colState + (i + 1))
        Next
    End With
    
End Sub


Function GetExamCode() As Integer
    Dim i, j As Long
    
    ClearSpread vasTemp
    GetExamCode = -1
    gAllExam = ""
    SQL = "Select equipcode, examcode, examname, resprec, seqno " & vbCrLf & _
          "  From EQPMASTER " & vbCrLf & _
          " Where equipno = '" & gEquip & "' " & vbCrLf & _
          " Order by  seqno * 10 "
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

'    Call dce_close_env      ' Server¿Í ¿¬°áÀ» ²÷´Â °÷
    DisConnect_Server
    DisConnect_Local
    Unload Me
    End
End Sub

Private Sub MnExamConfig_Click()
    'frmTestSet.Show
    frmTestSet.Show
    GetExamCode
End Sub

Private Sub MnExit_Click()
    Unload Me
End Sub

Private Sub MnModeBarcode_Click()
    chkMode.Caption = "Barcode"
    MnModeBarcode.Checked = True
    MnModeWorkList.Checked = False
    chkBar.Value = 1
    
    gIFMode = "Barcode"
    Call WritePrivateProfileString("config", "IFMode", gIFMode, App.Path & "\Interface.ini")
 
End Sub

Private Sub MnModeWorkList_Click()
    chkMode.Caption = "WorkList"
    MnModeBarcode.Checked = False
    MnModeWorkList.Checked = True
    chkBar.Value = 0

    gIFMode = "WorkList"
    Call WritePrivateProfileString("config", "IFMode", gIFMode, App.Path & "\Interface.ini")

End Sub

'Private Sub MnScr1_Click()
'    MnScr1.Checked = True
'    MnScr2.Checked = False
'
'    gScreen = "ºÐ¸®"
'    Call WritePrivateProfileString("config", "IFScreen", gScreen, App.Path & "\Interface.ini")
'
'End Sub
'
'Private Sub MnScr2_Click()
'    MnScr1.Checked = False
'    MnScr2.Checked = True
'
'    gScreen = "ÅëÇÕ"
'    Call WritePrivateProfileString("config", "IFScreen", gScreen, App.Path & "\Interface.ini")
'
'End Sub

Private Sub MnTConfig_Click()
    frmConfig.Show
End Sub

Private Sub MnTransAuto_Click()
    chkMode.Caption = "Auto"
    MnTransAuto.Checked = True
    MnTransManual.Checked = False
    chkMode.Value = 1

    gSave = "True"
    Call WritePrivateProfileString("config", "AutoSave", gSave, App.Path & "\Interface.ini")

End Sub

Private Sub MnTransManual_Click()
    chkMode.Caption = "Manual"
    MnTransAuto.Checked = False
    MnTransManual.Checked = True
    chkMode.Value = 0
    
    gSave = "False"
    Call WritePrivateProfileString("config", "AutoSave", gSave, App.Path & "\Interface.ini")

End Sub

'-----------------------------------------------------------------------------'
'   ±â´É : ¿À´õÁ¤º¸ Àü¼Û
'-----------------------------------------------------------------------------'
Private Sub SendOrder()
    Dim strOutput As String     '¼Û½ÅÇÒ µ¥ÀÌÅÍ
    
    '-- ASTM TYPEº° Define ÇØ¾ßÇÔ.
    '-- ASTM TYPE = Standard
    Select Case intSndPhase
        Case 1  '## Header
            'strOutput = intFrameNo & "H|\^&||| XN-10^00-14^15097^^^^AP795756||||||||E1394-97" & vbCr & ETX
            strOutput = intFrameNo & "H|\^&||||||||||P|1" & vbCr & ETX
            intSndPhase = 2
            intFrameNo = intFrameNo + 1
        Case 2  '## Patient
            'strOutput = intFrameNo & "P|1||||^^|||U|||||^||||||||||||^^^" & vbCr & ETX
            strOutput = intFrameNo & "P|1" & vbCr & ETX
            
            intSndPhase = 4
            intFrameNo = intFrameNo + 1
            
        Case 3  '## No Order
            
        Case 4  '## Order
            If mOrder.NoOrder = True Then
                    
                strOutput = intFrameNo & "O|1|" & mOrder.RackNo & "^" & mOrder.TubePos & "^" & Right(Space(15) & mOrder.BarNo, 15) & "^B||" & mOrder.Order & "|||||||N||||||||||||||Q"
                intSndPhase = 5
            
            Else
                If mOrder.IsSending = False Then   '## ÃÖÃÊ º¸³¾¶§
                    strOutput = "O|1|" & mOrder.RackNo & "^" & mOrder.TubePos & "^" & Right(Space(15) & mOrder.BarNo, 15) & "^B||" & mOrder.Order & "|||||||N||||||||||||||Q"
                    
                    If Len(strOutput) > 230 Then
                        mOrder.IsSending = True
                        mOrder.Order = Mid$(strOutput, 231)
                        strOutput = intFrameNo & Left(strOutput, 230) & vbCr & ETB
                        intSndPhase = 4
                    Else
                        strOutput = intFrameNo & strOutput & vbCr & ETX
                        intSndPhase = 5
                    End If
                Else                        '## ³²Àº ¹®ÀÚ¿­ÀÌ ÀÖÀ»¶§
                    strOutput = mOrder.Order
                    If Len(strOutput) > 230 Then
                        mOrder.Order = Mid$(strOutput, 231)
                        strOutput = intFrameNo & Left(strOutput, 230) & vbCr & ETB
                        intSndPhase = 4
                    Else
                        mOrder.IsSending = False
                        strOutput = intFrameNo & strOutput & vbCr & ETX
                        intSndPhase = 5
                    End If
                End If
                
            End If
            
            intFrameNo = intFrameNo + 1
            
        Case 5  '## Termianator
            strOutput = intFrameNo & "L|1|N" & vbCr & ETX
            intSndPhase = 6
            intFrameNo = intFrameNo + 1
            
        Case 6  '## EOT
            strState = ""
            comEqp.Output = EOT
            SetRawData "[Tx]" & EOT
            intFrameNo = 1
            
            Exit Sub
    End Select
    
    strOutput = STX & strOutput & GetChkSum(strOutput) & vbCrLf
    comEqp.Output = strOutput
    Debug.Print strOutput
    SetRawData "[Tx]" & strOutput
    
End Sub

'-----------------------------------------------------------------------------'
'   ±â´É : ÇØ´ç ¹®ÀÚ¿­ÀÇ CheckSumÀ» ±¸ÇÔ
'   ÀÎ¼ö :
'       - pMsg : ¹®ÀÚ¿­
'   ¹ÝÈ¯ : CheckSum
'-----------------------------------------------------------------------------'
Public Function GetChkSum(ByVal pMsg As String) As String
    Dim lngChkSum   As Long
    Dim i           As Long

    For i = 1 To Len(pMsg)
        lngChkSum = (lngChkSum + Asc(Mid(pMsg, i, 1))) Mod 256
    Next

    If lngChkSum = 0 Then
        GetChkSum = "00"
    Else
        GetChkSum = Mid("0" & Hex(lngChkSum), Len(Hex(lngChkSum)), 2)
    End If
End Function

'-- Áö±Ý³¯Â¥¿Í °Ë»çÀÏÀÚ ºñ±³ÇÑ´Ù
Function DateCompare(ByVal FDate As String) As String
    
    DateCompare = FDate
    If FDate <> Format(Now, "yyyymmdd") Then
        DateCompare = Format(Now, "yyyymmdd")
    End If
    
End Function

Private Sub tmrReceive_Timer()
    
    imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
    tmrReceive.Enabled = False

End Sub

Private Sub tmrSend_Timer()
    
    imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
    tmrSend.Enabled = False

End Sub

Public Sub SndMore()
    Dim strSndMsg As String
    
    'Call Sleep(1000)
    
    strSndMsg = ">"
    strSndMsg = Chr(2) & strSndMsg & Chr(3) ' & GetChkSum(strSndMsg) & vbCr
    comEqp.Output = strSndMsg & vbCrLf
    
    'SetRawData "[Tx]" & strSndMsg & vbCrLf
    Debug.Print "[SndMore]" & strSndMsg
    
End Sub

Public Sub SndRec()
    Dim strSndMsg As String
    
    strSndMsg = "A"
    strSndMsg = Chr(2) & strSndMsg & Chr(3) '& GetChkSum(strSndMsg)
    comEqp.Output = strSndMsg & vbCrLf
    
End Sub

Private Sub comEqp_OnComm()
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

            Buffer = comEqp.Input
            'Buffer = strBuffer
            
            
            SetRawData "[Rx]" & Buffer
            lngBufLen = Len(Buffer)
            
            Debug.Print Buffer

            For i = 1 To lngBufLen
                BufChar = Mid$(Buffer, i, 1)

                Select Case intPhase
                    Case 1      '## Estabilshment Phase
                        Select Case BufChar
                            Case ENQ
                                intBufCnt = 1
                                Erase strRecvData
                                ReDim Preserve strRecvData(intBufCnt)
                                intPhase = 2
                                comEqp.Output = ACK
                                SetRawData "[Tx]" & ACK
                            Case ACK
                                '-- Àåºñ¿¡¼­ ³Ñ¾î¿Â ½Ã°£ÀÌ ¿ì¿¬È÷ 11:59:59ÃÊ³ª ÀÍÀÏ¿¡ °¡±î¿î ½Ã°£ÀÏ °æ¿ì
                                '-- °á°ú ÀúÀå½Ã ÀÌÀüÀÏÀ» °¡Á®¿Ã ¼ö ÀÖÀ¸¹Ç·Î ³¯Â¥¸¦ ½Ç½Ã°£ ¾÷µ¥ÀÌÆ® ÇÑ´Ù.
                                strDate = DateCompare(Format(CDate(dtpToday.Value), "yyyymmdd"))
                                dtpToday.Value = Format(strDate, "####-##-##")
                                
                                DoEvents
                                
                                If strState = "Q" Then Call SendOrder
                        
                        End Select
                    Case 2      '## Transfer Phase
                        Select Case BufChar
                            Case ENQ
                                Erase strRecvData
                                comEqp.Output = ACK
                                SetRawData "[Tx]" & ACK
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
                                comEqp.Output = ACK
                                SetRawData "[Tx]" & ACK
                        End Select
                    Case 4      '## Termination Phase
                        Select Case BufChar
                            Case STX
                                intPhase = 2
                            Case EOT
                                '-- Àåºñ¿¡¼­ ³Ñ¾î¿Â ½Ã°£ÀÌ ¿ì¿¬È÷ 11:59:59ÃÊ³ª ÀÍÀÏ¿¡ °¡±î¿î ½Ã°£ÀÏ °æ¿ì
                                '-- °á°ú ÀúÀå½Ã ÀÌÀüÀÏÀ» °¡Á®¿Ã ¼ö ÀÖÀ¸¹Ç·Î ³¯Â¥¸¦ ½Ç½Ã°£ ¾÷µ¥ÀÌÆ® ÇÑ´Ù.
                                strDate = DateCompare(Format(CDate(dtpToday.Value), "yyyymmdd"))
                                dtpToday.Value = Format(strDate, "####-##-##")

                                DoEvents
                                
                                Call EditRcvDataASTM
                                
                                SetRawData "[strState]" & strState

                                If strState = "Q" Then
                                    intSndPhase = 1
                                    intFrameNo = 1
                                    comEqp.Output = ENQ
                                    SetRawData "[Tx]" & ENQ
                                End If
                                
                                intPhase = 1
                        End Select
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
            EVMsg$ = "CTS º¯°æ °¨Áö"
        Case comEvDSR
            EVMsg$ = "DSR º¯°æ °¨Áö"
        Case comEvCD
            EVMsg$ = "CD º¯°æ °¨Áö"
        Case comEvRing
            EVMsg$ = "ÀüÈ­ º§ÀÌ ¿ï¸®´Â Áß"
        Case comEvEOF
            EVMsg$ = "EOF °¨Áö"

        '¿À·ù ¸Þ½ÃÁö
        Case comBreak
            ERMsg$ = "Áß´Ü ½ÅÈ£ ¼ö½Å"
        Case comCDTO
            ERMsg$ = "¹Ý¼ÛÆÄ °ËÃâ ½Ã°£ ÃÊ°ú"
        Case comCTSTO
            ERMsg$ = "CTS ½Ã°£ ÃÊ°ú"
        Case comDCB
            ERMsg$ = "DCB °Ë»ö ¿À·ù"
        Case comDSRTO
            ERMsg$ = "DSR ½Ã°£ ÃÊ°ú"
        Case comFrame
            ERMsg$ = "ÇÁ·¹ÀÌ¹Ö ¿À·ù"
        Case comOverrun
            ERMsg$ = "ÆÐ¸®Æ¼ ¿À·ù"
        Case comRxOver
            ERMsg$ = "¼ö½Å ¹öÆÛ ÃÊ°ú"
        Case comRxParity
            ERMsg$ = "ÆÐ¸®Æ¼ ¿À·ù"
        Case comTxFull
            ERMsg$ = "Àü¼Û ¹öÆÛ¿¡ ¿©À¯°¡ ¾øÀ½"
        Case Else
            ERMsg$ = "¾Ë ¼ö ¾ø´Â ¿À·ù ¶Ç´Â ÀÌº¥Æ®"
    End Select


End Sub

'-----------------------------------------------------------------------------'
'   ±â´É : ÇØ´ç ¹ÙÄÚµå¹øÈ£¿¡ ´ëÇÑ Á¢¼öÁ¤º¸ Á¶È¸, Ç¥½Ã, °Ë»ç¿À´õ¸¸µé±â
'   ÀÎ¼ö :
'       - pBarNo : ¹ÙÄÚµå¹øÈ£
'-----------------------------------------------------------------------------'
Private Sub GetOrder(ByVal pBarNo As String)

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String
    
    intRow = -1
    
    For i = 1 To vasID.DataRowCnt
        If Trim(GetText(vasID, i, colBARCODE)) = pBarNo Then
            intRow = i
            Exit For
        End If
    Next i
    
    If intRow < 0 Then
        intRow = vasID.DataRowCnt + 1
        If vasID.MaxRows < intRow Then
            vasID.MaxRows = intRow
        End If
    End If
    
    '-- Àåºñ¼ö½ÅÁ¤º¸ Ç¥½Ã
    Call SetText(vasID, pBarNo, intRow, colBARCODE)         '-- ¹ÙÄÚµå
    'Call SetText(vasID, mOrder.Seq, intRow, colSeq)         '-- Seq
    Call SetText(vasID, mOrder.RackNo, intRow, colDISKNO)     '-- Rack
    Call SetText(vasID, mOrder.TubePos, intRow, colPOSNO)     '-- Pos
    
    '-- È¯ÀÚÁ¤º¸ Ç¥½Ã
    Call vasActiveCell(vasID, intRow, colBARCODE)
    Call ClearSpread(vasRes)
    
    '-- °Ë»çÀÚ Á¤º¸ ¼­¹öÅ×ÀÌºí¿¡¼­ °¡Á®¿Í Ç¥½Ã(for ¿öÅ©¸®½ºÆ®)  '6,7,8,9
    Call GetSampleInfoW(intRow)
    
    '-- ¹ÙÄÚµå¹øÈ£¿¡ Á¸ÀçÇÏ´Â °Ë»çÄÚµå °¡Á®¿À±â(ÀÎ¼ö : ÀåºñÄÚµå,Ã­Æ®¹øÈ£,Á¢¼öÀÏ,³»¿ø¹øÈ£,°ËÁø¹øÈ£)
    gOrderExam = GetOrderExamCode(gEquip, pBarNo)

    '-- ·ÎÄÃÅ×ÀÌºí¿¡¼­ °Ë»çÇ×¸ñ¿¡ ÇØ´çÇÏ´Â °Ë»çÃ¤³Î Ã£¾Æ¿À±â (intRow = ±âÁ¸ °Ë»çÇß´ø ¹ÙÄÚµå°¡ ´Ù½Ã ¿Ã¶ó¿Ã °æ¿ì À§Ä¡¸¦ ¸øÃ£´Â´Ù.)
    strItems = GetGetEquipExamCode_XN1000(gEquip, pBarNo, intRow)

    '-- °Ë»çÃ¤³Î·Î Àåºñ¿À´õ ¸¸µé±â
    If Trim(strItems) = "" Then
        mOrder.NoOrder = True
        mOrder.Order = strItems
    Else
        mOrder.NoOrder = False
        mOrder.Order = strItems
    End If
    
    Call SetText(vasID, "Order", intRow, colState)         '12 ÁøÇà»óÅÂ

End Sub

'-- ·ÎÄÃÅ×ÀÌºí¿¡¼­ °Ë»çÇ×¸ñ¿¡ ÇØ´çÇÏ´Â °Ë»çÃ¤³Î Ã£¾Æ¿À±â
Function GetGetEquipExamCode_XN1000(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim i As Integer
    Dim strExamCode As String
    Dim sBarcode     As String
    Dim strCBC As String
    Dim strDiff As String
    
    GetGetEquipExamCode_XN1000 = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
    sBarcode = Trim(GetText(frmInterface.vasID, intRow, colBARCODE))   '2 »ùÇÃ ¹ÙÄÚµå ¹øÈ£
    SetRawData "[sBarcode]" & sBarcode
    
    If sBarcode = "" Then
        Exit Function
    End If
    
    ClearSpread frmInterface.vasTemp1
    
    '-- °¡Á®¿Â °Ë»çÄÚµåÀÇ Ã¤³Î Ã£±â
    SQL = ""
    SQL = SQL & "SELECT Distinct EQUIPCODE "
    SQL = SQL & "  FROM EQPMASTER "
    SQL = SQL & " WHERE EQUIPNO  = '" & Trim(gEquip) & "' "
    SQL = SQL & "   AND EXAMCODE in (" & Trim(gOrderExam) & ")"
    
    Res = GetDBSelectRow(gLocal, SQL)
    strExamCode = ""

    strCBC = ""
    strDiff = ""
    
    For i = 0 To UBound(gReadBuf)
        If gReadBuf(i) <> "" Then
            'NRBC%´Â ¿À´õ¸¦ ¾ÈÁØ´Ù
'            If Trim(gReadBuf(i)) <> "NRBC%" Then
'                strExamCode = strExamCode & "^^^^" & Trim(gReadBuf(i)) & "\"
'            End If
            
            
            If Trim(gReadBuf(i)) = "WBC" Or Trim(gReadBuf(i)) = "RBC" Or Trim(gReadBuf(i)) = "HGB" Or _
                Trim(gReadBuf(i)) = "HCT" Or Trim(gReadBuf(i)) = "MCV" Or Trim(gReadBuf(i)) = "MCH" Or Trim(gReadBuf(i)) = "MCHC" Or _
                Trim(gReadBuf(i)) = "PLT" Or Trim(gReadBuf(i)) = "RDW-SD" Or Trim(gReadBuf(i)) = "RDW-CV" Or Trim(gReadBuf(i)) = "PDW" Or _
                Trim(gReadBuf(i)) = "MPV" Or Trim(gReadBuf(i)) = "P-LCR" Or Trim(gReadBuf(i)) = "PCT" Or Trim(gReadBuf(i)) = "NRBC#" Or Trim(gReadBuf(i)) = "NRBC%" Then
                
                strCBC = "^^^^WBC\^^^^RBC\^^^^HGB\^^^^HCT\^^^^MCV\^^^^MCH\^^^^MCHC\^^^^PLT\^^^^RDW-SD\^^^^RDW-CV\^^^^PDW\^^^^MPV\^^^^P-LCR\^^^^PCT\^^^^NRBC#\^^^^NRBC%\"
                
            End If

            If Trim(gReadBuf(i)) = "NEUT#" Or Trim(gReadBuf(i)) = "LYMPH#" Or Trim(gReadBuf(i)) = "MONO#" Or Trim(gReadBuf(i)) = "EO#" Or Trim(gReadBuf(i)) = "BASO#" Or _
                Trim(gReadBuf(i)) = "NEUT%" Or Trim(gReadBuf(i)) = "LYMPH%" Or Trim(gReadBuf(i)) = "MONO%" Or Trim(gReadBuf(i)) = "EO%" Or Trim(gReadBuf(i)) = "BASO%" Or _
                Trim(gReadBuf(i)) = "IG#" Or Trim(gReadBuf(i)) = "IG%" Then
               
                '-- ^^^^LYMPH#\°¡ µÎ°³ÀÎ ÀÌÀ¯´Â ETB ¸¦ Àåºñ¿¡¼­ ÀÎ½ÄÇÏÁö ¸øÇÏ±â ‹š¹®..(±× ÀÚ¸®°¡ 230)
                strDiff = "^^^^NEUT#\^^^^LYMPH%\^^^^MONO#\^^^^EO#\^^^^BASO#\^^^^NEUT%\^^^^LYMPH#\^^^^LYMPH#\^^^^MONO%\^^^^EO%\^^^^BASO%\^^^^IG#\^^^^IG%\"
                
            End If
        Else
            Exit For
        End If
    Next

    strExamCode = strCBC & strDiff
    
    '-- ¿À´õ°¡ ¾øÀ» °æ¿ì CBC¸¸ °Ë»çÇÏµµ·Ï ÇÑ´Ù.
    If strExamCode = "" Then
        strExamCode = "^^^^WBC\^^^^RBC\^^^^HGB\^^^^HCT\^^^^MCV\^^^^MCH\^^^^MCHC\^^^^PLT\^^^^RDW-SD\^^^^RDW-CV\^^^^PDW\^^^^MPV\^^^^P-LCR\^^^^PCT\^^^^NRBC#\^^^^NRBC%\"
        strExamCode = strExamCode & "^^^^NEUT#\^^^^LYMPH%\^^^^MONO#\^^^^EO#\^^^^BASO#\^^^^NEUT%\^^^^LYMPH#\^^^^LYMPH#\^^^^MONO%\^^^^EO%\^^^^BASO%\^^^^IG#\^^^^IG%\"
    End If
    
    If strExamCode <> "" Then
        strExamCode = Mid(strExamCode, 1, Len(strExamCode) - 1)
    End If
    
    GetGetEquipExamCode_XN1000 = strExamCode
    
End Function

'-----------------------------------------------------------------------------'
'   ±â´É :
'   ÀÎ¼ö :
'       - pBarNo : ¹ÙÄÚµå¹øÈ£
'-----------------------------------------------------------------------------'
Private Sub SetPatInfo(ByVal pBarNo As String)
    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strTestDt   As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String
    
    intRow = -1
    
    For i = 1 To vasID.DataRowCnt
        If Trim(GetText(vasID, i, colBARCODE)) = pBarNo Then
            intRow = i
            Exit For
        End If
    Next i
    
    If intRow < 0 Then
        intRow = vasID.DataRowCnt + 1
        If vasID.MaxRows < intRow Then
            vasID.MaxRows = intRow
        End If
    End If
    
    '-- Àåºñ¼ö½ÅÁ¤º¸ Ç¥½Ã
    Call SetText(vasID, "1", intRow, colCheckBox)
    Call SetText(vasID, pBarNo, intRow, colBARCODE)
    Call SetText(vasID, mResult.RackNo, intRow, colDISKNO)
    Call SetText(vasID, mResult.TubePos, intRow, colPOSNO)
    Call SetText(vasID, mResult.RsltDate, intRow, colEXAMDATE)
    Call SetText(vasID, mResult.RsltSeq, intRow, colSAVESEQ)
    
    Call vasActiveCell(vasID, intRow, colBARCODE)
    
    '-- °á°ú½ºÇÁ·¹µå Áö¿ì±â
    Call ClearSpread(vasRes)
    
    '-- °Ë»çÀÚ Á¤º¸ ¼­¹öÅ×ÀÌºí¿¡¼­ °¡Á®¿Í Ç¥½Ã(for ¿öÅ©¸®½ºÆ®)  '6,7,8,9
    Call GetSampleInfoW(intRow)
    
    '-- ¹ÙÄÚµå¹øÈ£¿¡ Á¸ÀçÇÏ´Â °Ë»çÄÚµå °¡Á®¿À±â(ÀÎ¼ö : ÀåºñÄÚµå,Ã­Æ®¹øÈ£,Á¢¼öÀÏ,³»¿ø¹øÈ£,°ËÁø¹øÈ£)
    gOrderExam = GetOrderExamCode(gEquip, pBarNo)
    
    '-- ÇöÀç Row
    gRow = intRow
    
End Sub


'-----------------------------------------------------------------------------'
'   ±â´É : Àåºñ·ÎºÎ ¼ö½ÅÇÑ µ¥ÀÌÅÍ ÆíÁý
'-----------------------------------------------------------------------------'
Private Sub EditRcvDataASTM()
    Dim strRcvBuf    As String   '¼ö½ÅÇÑ Data
    Dim strType      As String   '¼ö½ÅÇÑ Record Type
    Dim strBarNo     As String   '¼ö½ÅÇÑ ¹ÙÄÚµå¹øÈ£
    Dim strSeq       As String   '¼ö½ÅÇÑ Sequence
    Dim strRackNo    As String   '¼ö½ÅÇÑ Rack Or Disk No
    Dim strTubePos   As String   '¼ö½ÅÇÑ Tube Position
    Dim strIntBase   As String   '¼ö½ÅÇÑ Àåºñ±âÁØ °Ë»ç¸í
    Dim strResult    As String   '¼ö½ÅÇÑ °á°ú(Á¤¼º)
    Dim strIntResult As String   '¼ö½ÅÇÑ °á°ú(Á¤·®)
    Dim strQCResult  As String   '¼ö½ÅÇÑ °á°ú(QC)
    Dim strFlag      As String   '¼ö½ÅÇÑ Abnormal Flag
    Dim strComm      As String   '¼ö½ÅÇÑ Comment
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
    Dim intIdx      As Integer
    Dim varRcvBuf   As Variant
    Dim intRow      As Integer
    Dim i As Integer
    Dim intCol As Integer
        
    For intCnt = 1 To UBound(strRecvData)
        strRcvBuf = strRecvData(intCnt)
        
        strType = Mid$(strRcvBuf, 1, 1)
        If IsNumeric(strType) Then
            strType = Mid$(strRcvBuf, 2, 1)
        End If
        
        Select Case strType
            Case "H"    '## Header
            Case "Q"    '## Request Information
                '## ¹ÙÄÚµå¹øÈ£, SEQ, Disk No, Tube Position Á¶È¸
                If mGetP(strRcvBuf, 13, "|") = "A" Then Exit Sub
                strTemp1 = mGetP(strRcvBuf, 3, "|")
                strBarNo = Trim$(mGetP(strTemp1, 3, "^"))
                '2Q|1|7^1^      201405080002^M||||20140508120418||||||N0B
                With mOrder
                    .BarNo = strBarNo
                    .RackNo = mGetP(strTemp1, 1, "^")
                    .TubePos = mGetP(strTemp1, 2, "^")
                    .Seq = mGetP(strTemp1, 4, "^")
                End With
                
                If strBarNo = "" Then Exit Sub
                
                Call GetOrder(strBarNo)
                
                strState = "Q"
                
            Case "O"    '## Order
                strBarNo = Trim(mGetP(mGetP(strRcvBuf, 4, "|"), 3, "^"))
                strTemp1 = mGetP(strRcvBuf, 4, "|")
                strRackNo = mGetP(strTemp1, 1, "^")
                strTubePos = mGetP(strTemp1, 2, "^")
                
                If strBarNo = "" Then Exit Sub
                
                With mResult
                    .BarNo = strBarNo
                    .RackNo = strRackNo
                    .TubePos = strTubePos
                    .RsltDate = Format(Now, "yyyymmddhhmmss")
                    .RsltSeq = getMaxTestNum(Format(dtpToday, "yyyymmdd"))
                End With
                                
                Call SetPatInfo(strBarNo)
                
                If gRow <= 0 Then
                    Exit Sub
                End If
                
                strState = "O"
                
                '-- ¿À¸¥ÂÊ °á°úÈ­¸é ÃÊ±âÈ­
                vasRes.MaxRows = 0
                
            Case "R"    '## Result
                '## Àåºñ±âÁØ °Ë»ç¸í, °á°ú, Abnormal Flag
                strTemp1 = mGetP(mGetP(strRcvBuf, 3, "|"), 5, "^")
                strTemp2 = mGetP(strRcvBuf, 4, "|")
                strFlag = mGetP(strRcvBuf, 7, "|")
                strIntBase = strTemp1
                strResult = ""
                If InStr(strTemp2, "^") > 0 Then
                    '## Á¤¼º°á°ú ÀúÀå
                    strResult = mGetP(strTemp2, 2, "^")
                Else
                    '## Á¤·®°á°ú ÀúÀå
                    strResult = strTemp2
                End If
                
                If strResult <> "" And Len(strIntBase) <= 6 Then
                    SQL = ""
                    SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
                    SQL = SQL & "  FROM EQPMASTER"
                    SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
                    SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
                    SQL = SQL & "   AND EXAMCODE in (" & gOrderExam & ") "
                    
                    Res = GetDBSelectColumn(gLocal, SQL)
                    
                    '-- ¿À´õ ÀÖÀ» °æ¿ì
                    If Res > 0 Then
                        lsExamCode = Trim(gReadBuf(0))
                        lsExamName = Trim(gReadBuf(1))
                        lsSeqNo = Trim(gReadBuf(2))
                        
                        lsResRow = vasRes.DataRowCnt + 1
                        If vasRes.MaxRows < lsResRow Then
                            vasRes.MaxRows = lsResRow
                        End If
                        
                        '¼Ò¼öÁ¡ Ã³¸®, °á°ú ÇüÅÂ Ã³¸®
                        lsEquipRes = strResult
                        strResult = SetResult(strResult, strIntBase)
                        lsResult_Buff = strResult
                        
                        '-- Work List
                        SetText vasID, "Result", gRow, colState                 '11 ÁøÇà»óÅÂ
                        

                        '-- °á°ú List
                        SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      'ÀåºñÄÚµå
                        SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '°Ë»çÄÚµå
                        SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '°Ë»ç¸í
                        SetText vasRes, lsEquipRes, lsResRow, colMachResult     'Àåºñ°á°ú
                        SetText vasRes, strResult, lsResRow, colRESULT          '°á°ú
                        SetText vasRes, lsSeqNo, lsResRow, colSeq               '¼ø¹ø
                        SetText vasRes, strComm, lsResRow, 7                    'Flag
                        '-- ·ÎÄÃ ÀúÀå
                        SetLocalDB gRow, lsResRow, "1", lsEquipRes
                        
                        For intCol = colState + 1 To vasID.MaxCols
                            If lsExamCode = gArrEquip(intCol - colState, 3) Then
                                SetText vasID, strResult, gRow, intCol
                                Exit For
                            End If
                        Next
                        
                        lsResult_Buff = ""
                        strState = "R"
                        
                    '-- ¿À´õ ¾øÀ» °æ¿ì
                    Else
                    
                              SQL = "Select examcode, examname, seqno "
                        SQL = SQL & "  From EQPMASTER"
                        SQL = SQL & " Where equipno = '" & gEquip & "' "
                        SQL = SQL & "   and equipcode = '" & strIntBase & "' "
                        Res = GetDBSelectColumn(gLocal, SQL)
                        
                        If Res > 0 Then
                            lsExamCode = Trim(gReadBuf(0))
                            lsExamName = Trim(gReadBuf(1))
                            lsSeqNo = Trim(gReadBuf(2))
                            
                            lsResRow = vasRes.DataRowCnt + 1
                            If vasRes.MaxRows < lsResRow Then
                                vasRes.MaxRows = lsResRow
                            End If
                            
                            '¼Ò¼öÁ¡ Ã³¸®, °á°ú ÇüÅÂ Ã³¸®
                            lsEquipRes = strResult
                            'strResult = SetResult(strResult, strIntBase)
                            lsResult_Buff = strResult
                            
                            '-- Work List
                            SetText vasID, "Result", gRow, colState                 'ÁøÇà»óÅÂ
                            
                            '-- °á°ú List
                            SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      'ÀåºñÄÚµå
                            SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '°Ë»çÄÚµå
                            SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '°Ë»ç¸í
                            SetText vasRes, lsEquipRes, lsResRow, colMachResult     'Àåºñ°á°ú
                            SetText vasRes, strResult, lsResRow, colRESULT          '°á°ú
                            SetText vasRes, lsSeqNo, lsResRow, colSeq               '¼ø¹ø
                            SetText vasRes, strComm, lsResRow, colFLAG              'Flag
                            '-- ·ÎÄÃ ÀúÀå
                            SetLocalDB gRow, lsResRow, "1", lsEquipRes
                            
                            For intCol = colState + 1 To vasID.MaxCols
                                If lsExamCode = gArrEquip(intCol - colState, 3) Then
                                    SetText vasID, strResult, gRow, intCol
                                    Exit For
                                End If
                            Next
                            
                            lsResult_Buff = ""
                            strState = "R"
                        End If
                    End If
                End If
                vasRes.RowHeight(-1) = 14
                
            Case "C"    '## Comment
                '## Abnormal °á°úÀÏ¶§ Comment ÀúÀå
'                If strFlag <> "N" Then
'                    strTemp1 = mGetP(strRcvBuf, 4, "|")
'                    strComm = "[Flag]: " & mGetP(strTemp1, 1, "^") & ", " & mGetP(strTemp1, 2, "^")
'                End If
                
            Case "L"    '## Terminator
                '## DB¿¡ °á°úÀúÀå
                If MnTransAuto.Checked = True And strState = "R" Then
                    Res = SaveTransDataW(gRow)
                    
                    If Res = -1 Then
                        '-- ÀúÀå ½ÇÆÐ
                        SetForeColor vasID, gRow, gRow, 1, colState, 255, 0, 0
                        SetText vasID, "Failed", gRow, colState
                    Else
                        '-- ÀúÀå ¼º°ø
                        SetBackColor vasID, gRow, gRow, 1, colState, 202, 255, 112
                        SetText vasID, "Trans", gRow, colState
                        SetText vasID, "0", gRow, colCheckBox
                        
                              SQL = "Update PATRESULT Set " & vbCrLf
                        SQL = SQL & " sendflag = '2' " & vbCrLf
                        SQL = SQL & " Where equipno = '" & gEquip & "' " & vbCrLf
                        'SQL = SQL & "   And examdate = '" & Trim(GetText(vasID, gRow, colOrdDate)) & "' "
                        'SQL = SQL & "   And barcode = '" & Trim(GetText(vasID, gRow, colBARCODE)) & "' "
                        
                        Res = SendQuery(gLocal, SQL)
                        If Res = -1 Then
                            SaveQuery SQL
                            Exit Sub
                        End If
                    End If
                    strState = ""
                End If
            
                'SetText vasID, "Result", gRow, colState
                'strState = ""
        
        End Select
    Next


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
' asRow2 = °á°ú List
Function SetLocalDB(ByVal asRow1 As Long, ByVal asRow2 As Long, asSend As String, Optional asEquipResult As String = "")
    Dim sCnt As String
    Dim sExamDate As String
    Dim strSaveSeq As String
    
    sExamDate = Format(dtpToday, "yyyymmddhhmmss")
    'sExamDate = Trim(GetText(vasID, asRow1, colOrdDate))
    
    SQL = ""
    SQL = "DELETE FROM PATRESULT " & vbCrLf & _
          " WHERE EXAMDATE = '" & sExamDate & "' " & vbCrLf & _
          "   AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
          "   AND SAVESEQ = " & Trim(GetText(vasID, asRow1, colSAVESEQ)) & vbCrLf & _
          "   AND BARCODE = '" & Trim(GetText(vasID, asRow1, colBARCODE)) & "' " & vbCrLf & _
          "   AND EQUIPCODE = '" & Trim(GetText(vasRes, asRow2, colEQUIPCODE)) & "'" & vbCrLf & _
          "   AND EXAMCODE = '" & Trim(GetText(vasRes, asRow2, colEXAMCODE)) & "'"
          
    Res = SendQuery(gLocal, SQL)
    
    If Res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
    'strSaveSeq = getMaxTestNum(Mid(sExamDate, 1, 8))

    SQL = ""
    SQL = SQL & "INSERT INTO PATRESULT (" & vbCrLf
    SQL = SQL & "SAVESEQ"                           'ÀúÀå¼ø¹ø(³¯Â¥º°)
    SQL = SQL & ", EXAMDATE"                        '°Ë»çÀÏÀÚ"
    SQL = SQL & ", HOSPDATE"                        'º´¿øÁ¢¼öÀÏÀÚ"
    SQL = SQL & ", EQUIPNO"                         'ÀåºñÄÚµå"
    SQL = SQL & ", BARCODE" & vbCrLf
    SQL = SQL & ", EQUIPCODE"                       '°Ë»çÃ¤³Î"
    SQL = SQL & ", EXAMCODE"                        'º´¿ø°Ë»çÄÚµå"
    SQL = SQL & ", EXAMSUBCODE"                     'º´¿ø°Ë»çÄÚµå(SUB)"
    SQL = SQL & ", EXAMNAME"
    SQL = SQL & ", SEQNO" & vbCrLf                  '°Ë»çÀÏ·Ã¹øÈ£"
    SQL = SQL & ", SAMPLETYPE"                      '°ËÃ¼À¯Çü"
    SQL = SQL & ", DISKNO"
    SQL = SQL & ", POSNO"
    SQL = SQL & ", EQUIPRESULT"                     'Àåºñ°á°ú"
    SQL = SQL & ", RESULT" & vbCrLf                 '¼Ò¼öÁ¡Àû¿ë°á°ú"
    SQL = SQL & ", REFFLAG"
    SQL = SQL & ", REFVALUE"
    SQL = SQL & ", CHARTNO"
    SQL = SQL & ", PID"                             'º´·Ï¹øÈ£(³»¿ø¹øÈ£)"
    SQL = SQL & ", PNAME" & vbCrLf
    SQL = SQL & ", PSEX"
    SQL = SQL & ", PAGE"
    SQL = SQL & ", PJUMIN"
    SQL = SQL & ", PANICVALUE"
    SQL = SQL & ", DELTAVALUE" & vbCrLf
    SQL = SQL & ", SENDFLAG"                        'Àü¼Û±¸ºÐ(0:¹ÌÀü¼Û,1:Àü¼Û)"
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
    SQL = SQL & "','" & Trim(GetText(vasRes, asRow2, colSubCode))
    SQL = SQL & "','" & Trim(GetText(vasRes, asRow2, colEXAMNAME))
    SQL = SQL & "','" & Trim(GetText(vasRes, asRow2, colSeq))
    SQL = SQL & "',''"
    SQL = SQL & ",'" & Trim(GetText(vasID, asRow1, colDISKNO))
    SQL = SQL & "','" & Trim(GetText(vasID, asRow1, colPOSNO))
    SQL = SQL & "','" & Trim(GetText(vasRes, asRow2, colMachResult))
    SQL = SQL & "','" & Trim(GetText(vasRes, asRow2, colRESULT))
    SQL = SQL & "',''"
    SQL = SQL & ",''"
    SQL = SQL & ",'" & Trim(GetText(vasID, asRow1, colCHARTNO))
    SQL = SQL & "','" & Trim(GetText(vasID, asRow1, colPID))
    SQL = SQL & "','" & Trim(GetText(vasID, asRow1, colPNAME))
    SQL = SQL & "','" & Trim(GetText(vasID, asRow1, colPSEX))
    SQL = SQL & "','" & Trim(GetText(vasID, asRow1, colPAGE))
    SQL = SQL & "',''"
    SQL = SQL & ",''"
    SQL = SQL & ",''"
    SQL = SQL & ",'0'"
    SQL = SQL & ",''"
    SQL = SQL & ",'" & gIFUser
    SQL = SQL & "','')"
    
    Res = SendQuery(gLocal, SQL)
    
    If Res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
End Function


'-- ¿À´Ã °Ë»çÇÑ ³¯Â¥ÀÇ Max + 1 ¹øÈ£¸¦ °¡Á®¿Â´Ù
Private Function getMaxTestNum(ByVal strDate As String) As Long

    getMaxTestNum = 1
    
    '-- °á°ú¾÷µ¥ÀÌÆ®
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

Private Sub Var_Clear()
    
    gsBarCode = ""
    gsPID = ""
    gsRackNo = ""
    gsPosNo = ""
    gsResDateTime = ""
    gsSeqNo = ""
    gsExamCode = ""
    gsExamName = ""
    gsOrder = ""
    gsResult = ""

End Sub



Private Sub picLogin_Click()

    Dim sMsg As String
    sMsg = "°Ë»çÀÚ¸¦ ÀÔ·ÂÇØÁÖ¼¼¿ä."
    lblUser.Caption = InputBox(sMsg, "°Ë»çÀÚ ÀÔ·Â")

End Sub

Private Sub txtBarNum_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Trim(txtBarNum) <> "" Then
        'If Len(txtBarNum) = 12 And IsNumeric(txtBarNum) Then
            Call GetWorkList(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"), Trim(txtBarNum))
        End If
        vasID.RowHeight(-1) = 12
        txtBarNum.Text = ""
    End If
    
End Sub


Private Sub vasID_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    Dim i As Integer
    
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
    Dim lsID As String
    
    If Row < 1 Or Row > vasID.DataRowCnt Then
        Exit Sub
    End If
    
'    lblDate.Caption = Trim(GetText(vasID, Row, colHOSPDATE))
    lsID = Trim(GetText(vasID, Row, colBARCODE))
    lblChangeBar.Caption = lsID
    lblBarcode(0).Caption = lsID
    lblPname(0).Caption = Trim(GetText(vasID, Row, colPNAME))
    lblSaveSeq.Caption = Trim(GetText(vasID, Row, colSAVESEQ))
    lblExamDate.Caption = Trim(GetText(vasID, Row, colEXAMDATE))
    
    'Local¿¡¼­ ºÒ·¯¿À±â
    ClearSpread vasRes
    
    'ÀåºñÄÚµå, °Ë»çÄÚµå, °Ë»ç¸í, °á°ú, ¼ø¹ø
          SQL = "SELECT EQUIPCODE, EXAMCODE, EXAMNAME, EQUIPRESULT, RESULT, SEQNO, SENDFLAG " & vbCrLf
    SQL = SQL & "  FROM PATRESULT " & vbCrLf
    SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "'" & vbCrLf
    SQL = SQL & "   AND SAVESEQ = " & lblSaveSeq.Caption & vbCrLf
    SQL = SQL & "   AND BARCODE = '" & lsID & "' " & vbCrLf
    'SQL = SQL & "   AND EXAMDATE = '" & Mid(Trim(GetText(vasID, Row, colOrdDate)), 1, 8) & "' " & vbCrLf
    SQL = SQL & " GROUP BY SEQNO, EQUIPCODE, EXAMCODE, EXAMNAME, EQUIPRESULT, RESULT, SENDFLAG "
    SQL = SQL & " ORDER BY SEQNO * 10"
    
    Res = GetDBSelectVas(gLocal, SQL, vasRes)
    If Res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If

    vasRes.MaxRows = vasRes.DataRowCnt
    vasRes.RowHeight(-1) = 12
    
End Sub

Private Sub vasID_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim iRow    As Long
    Dim lsID    As String
    Dim lsTime  As String
    Dim lsPid   As String
    Dim lsSeq   As String
    Dim i       As Integer

    iRow = vasID.ActiveRow
    If KeyCode = vbKeyDelete Then
        If iRow < 1 Or iRow > vasID.DataRowCnt Then
            Exit Sub
        End If

        lsID = Trim(GetText(vasID, iRow, colBARCODE))
        lsPid = Trim(GetText(vasID, iRow, colPID))
        lsSeq = Trim(GetText(vasID, iRow, colSAVESEQ))
        
        If lsID = "" Or lsPid = "" Or lsSeq = "" Then
            Exit Sub
        End If
        
        If MsgBox(Trim(GetText(vasID, iRow, colPNAME)) & " È¯ÀÚÀÇ °á°ú¸¦ »èÁ¦ÇÏ½Ã°Ú½À´Ï±î?", vbInformation + vbYesNo, "¾Ë¸²") = vbNo Then
            Exit Sub
        End If

              SQL = "DELETE FROM PATRESULT " & vbCrLf
        SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf
        SQL = SQL & "   AND BARCODE = '" & lsID & "' " & vbCrLf
        SQL = SQL & "   AND PID = '" & lsPid & "' " & vbCrLf
        SQL = SQL & "   AND SAVESEQ = " & lsSeq & vbCrLf
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Format(dtpToday.Value, "yyyymmdd") & "' "
        Res = SendQuery(gLocal, SQL)

        If Res = -1 Then
            SaveQuery SQL
            Exit Sub
        End If

        DeleteRow vasID, iRow, iRow
        vasRes.MaxRows = 0
        
    ElseIf KeyCode = vbKeyReturn Then

        Call GetSampleInfoW(iRow)

        lsID = Trim(GetText(vasID, iRow, colBARCODE))

        If lsID <> lblChangeBar.Caption Then
                  
                  SQL = "UPDATE PATRESULT SET"
            SQL = SQL & " HOSPDATE = '" & Trim(GetText(vasID, iRow, colHOSPDATE)) & "' " & vbCrLf
            SQL = SQL & ",BARCODE = '" & lsID & "' " & vbCrLf
            SQL = SQL & ",CHARTNO = '" & Trim(GetText(vasID, iRow, colCHARTNO)) & "' " & vbCrLf
            SQL = SQL & ",PID = '" & Trim(GetText(vasID, iRow, colPID)) & "' " & vbCrLf
            SQL = SQL & ",PNAME = '" & Trim(GetText(vasID, iRow, colPNAME)) & "' " & vbCrLf
            SQL = SQL & ",PSEX = '" & Trim(GetText(vasID, iRow, colPSEX)) & "' " & vbCrLf
            SQL = SQL & ",PAGE = '" & Trim(GetText(vasID, iRow, colPAGE)) & "' " & vbCrLf
            SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf
            SQL = SQL & "   AND SAVESEQ = " & Trim(GetText(vasID, iRow, colSAVESEQ)) & vbCrLf
            SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Trim(GetText(vasID, iRow, colEXAMDATE)) & "' " & vbCrLf
            SQL = SQL & "   AND BARCODE = '" & lblChangeBar.Caption & "' "
        
            Res = SendQuery(gLocal, SQL)
        End If
        SetText vasID, "¼öÁ¤", iRow, colState

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


Private Sub vasRes_KeyPress(KeyAscii As Integer)

    With vasRes
        If KeyAscii = 13 And .ActiveCol = colRESULT Then
            SQL = ""
            SQL = SQL & "UPDATE PATRESULT " & vbCrLf
            SQL = SQL & "   SET RESULT    ='" & Trim(GetText(vasRes, .ActiveRow, colRESULT)) & "' " & vbCrLf
            SQL = SQL & " WHERE BARCODE   = '" & Trim(lblBarcode(0).Caption) & "' " & vbCrLf
            SQL = SQL & "   AND MID(EXAMDATE,1,8)  = '" & Trim(lblExamDate.Caption) & "' " & vbCrLf
            SQL = SQL & "   AND SAVESEQ   = " & lblSaveSeq.Caption & vbCrLf
            SQL = SQL & "   AND EQUIPNO   = '" & gEquip & "' " & vbCrLf
            SQL = SQL & "   AND EXAMCODE  = '" & Trim(GetText(vasRes, .ActiveRow, colEXAMCODE)) & "' " & vbCrLf
            SQL = SQL & "   AND EQUIPCODE = '" & Trim(GetText(vasRes, .ActiveRow, colEQUIPCODE)) & "' " & vbCrLf

            Res = SendQuery(gLocal, SQL)

            If Res = -1 Then
                SaveQuery SQL
                Exit Sub
            End If

        End If
    End With

End Sub

