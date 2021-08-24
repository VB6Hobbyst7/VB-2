VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInterface 
   BorderStyle     =   1  '´ÜÀÏ °íÁ¤
   Caption         =   " AU5800 Interface Program"
   ClientHeight    =   10635
   ClientLeft      =   765
   ClientTop       =   840
   ClientWidth     =   15075
   FillColor       =   &H0000FFFF&
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10635
   ScaleWidth      =   15075
   Begin VB.CommandButton cmdEnd 
      Caption         =   "´Ý±â"
      BeginProperty Font 
         Name            =   "±¼¸²Ã¼"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13905
      TabIndex        =   55
      Top             =   45
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   585
      Left            =   7545
      TabIndex        =   54
      Top             =   5040
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.TextBox Text1 
      Height          =   1665
      Left            =   2175
      TabIndex        =   53
      Text            =   "Text1"
      Top             =   4710
      Visible         =   0   'False
      Width           =   4245
   End
   Begin VB.CommandButton cmd°á°ú»èÁ¦ 
      Caption         =   "°á°ú»èÁ¦"
      BeginProperty Font 
         Name            =   "±¼¸²Ã¼"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10500
      TabIndex        =   3
      Top             =   45
      Width           =   1095
   End
   Begin VB.CommandButton cmdCall 
      Caption         =   "°á°úÁ¶È¸"
      BeginProperty Font 
         Name            =   "±¼¸²Ã¼"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8220
      TabIndex        =   1
      Top             =   45
      Width           =   1095
   End
   Begin VB.CommandButton txtResPrint 
      Appearance      =   0  'Æò¸é
      Caption         =   "°á°úÃâ·Â"
      BeginProperty Font 
         Name            =   "±¼¸²Ã¼"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9360
      TabIndex        =   2
      Top             =   45
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   22200
      TabIndex        =   51
      Top             =   10740
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "[Á¶È¸ÀÏÀÚ]"
      BeginProperty Font 
         Name            =   "±¼¸²Ã¼"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   4140
      TabIndex        =   49
      Top             =   60
      Width           =   1755
      Begin MSComCtl2.DTPicker dtpExamDate 
         Height          =   315
         Left            =   60
         TabIndex        =   0
         Top             =   240
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
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
         Format          =   104595457
         CurrentDate     =   39699
      End
   End
   Begin FPSpread.vaSpread vasOrder 
      Height          =   1815
      Left            =   16500
      TabIndex        =   40
      Top             =   420
      Visible         =   0   'False
      Width           =   7425
      _Version        =   393216
      _ExtentX        =   13097
      _ExtentY        =   3201
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
      SpreadDesigner  =   "frmInterface.frx":0442
   End
   Begin FPSpread.vaSpread vasOrder1 
      Height          =   2505
      Left            =   16800
      TabIndex        =   29
      Top             =   2640
      Visible         =   0   'False
      Width           =   7005
      _Version        =   393216
      _ExtentX        =   12356
      _ExtentY        =   4419
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
      SpreadDesigner  =   "frmInterface.frx":0671
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   3210
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      InputLen        =   1
      RThreshold      =   1
      RTSEnable       =   -1  'True
      EOFEnable       =   -1  'True
   End
   Begin VB.TextBox txtBuff 
      Height          =   1995
      Left            =   17760
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   10620
      Visible         =   0   'False
      Width           =   4185
   End
   Begin FPSpread.vaSpread vasRes 
      Height          =   9195
      Left            =   6840
      TabIndex        =   11
      Top             =   900
      Width           =   8115
      _Version        =   393216
      _ExtentX        =   14314
      _ExtentY        =   16219
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      ColHeaderDisplay=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²Ã¼"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GridColor       =   16777215
      MaxCols         =   13
      Protect         =   0   'False
      SpreadDesigner  =   "frmInterface.frx":08A0
   End
   Begin VB.CheckBox ChkAll 
      Height          =   255
      Left            =   705
      TabIndex        =   20
      Top             =   990
      Width           =   165
   End
   Begin VB.CommandButton cmd_Trans 
      Caption         =   "¼±ÅÃÀü¼Û"
      BeginProperty Font 
         Name            =   "±¼¸²Ã¼"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11640
      Picture         =   "frmInterface.frx":47FA
      TabIndex        =   4
      Top             =   45
      Width           =   1095
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "±¼¸²Ã¼"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12780
      TabIndex        =   5
      Top             =   45
      Width           =   1095
   End
   Begin FPSpread.vaSpread vasID 
      Height          =   9195
      Left            =   180
      TabIndex        =   10
      Top             =   915
      Width           =   6585
      _Version        =   393216
      _ExtentX        =   11615
      _ExtentY        =   16219
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      ColHeaderDisplay=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GridColor       =   16777215
      MaxCols         =   21
      Protect         =   0   'False
      SpreadDesigner  =   "frmInterface.frx":54C4
   End
   Begin VB.Frame Frame1 
      Height          =   9495
      Index           =   0
      Left            =   60
      TabIndex        =   19
      Top             =   720
      Width           =   14955
      Begin VB.TextBox txtRowF 
         Height          =   375
         Left            =   600
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   300
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.TextBox txtRowT 
         Height          =   375
         Left            =   1620
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   300
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.CommandButton cmdRowSet 
         Caption         =   "Àû¿ë"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2340
         TabIndex        =   8
         Top             =   300
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Frame Frame3 
         Height          =   555
         Left            =   4200
         TabIndex        =   41
         Top             =   600
         Visible         =   0   'False
         Width           =   7635
         Begin VB.TextBox txtPS 
            Appearance      =   0  'Æò¸é
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3540
            TabIndex        =   44
            Top             =   150
            Width           =   705
         End
         Begin VB.TextBox txtPE 
            Appearance      =   0  'Æò¸é
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4680
            TabIndex        =   43
            Top             =   150
            Width           =   675
         End
         Begin VB.TextBox txtPrint 
            Appearance      =   0  'Æò¸é
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   180
            TabIndex        =   42
            Top             =   150
            Width           =   1515
         End
         Begin VB.Label Label8 
            Caption         =   "°Ë»ç¹øÈ£ :"
            Height          =   285
            Left            =   2400
            TabIndex        =   46
            Top             =   210
            Width           =   1065
         End
         Begin VB.Label Label7 
            Caption         =   "~"
            Height          =   195
            Left            =   4410
            TabIndex        =   45
            Top             =   210
            Width           =   255
         End
      End
      Begin VB.Frame Frame2 
         Height          =   555
         Left            =   7050
         TabIndex        =   32
         Top             =   120
         Visible         =   0   'False
         Width           =   7635
         Begin VB.CommandButton cmd°ËÃ¼¹øÈ£»ý¼º 
            Caption         =   "°ËÃ¼¹øÈ£»ý¼º"
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1800
            TabIndex        =   48
            Top             =   90
            Width           =   1515
         End
         Begin VB.TextBox txtReceHead 
            Appearance      =   0  'Æò¸é
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   180
            TabIndex        =   39
            Top             =   150
            Width           =   1515
         End
         Begin VB.TextBox txtResN 
            Appearance      =   0  'Æò¸é
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   6810
            TabIndex        =   35
            Top             =   150
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.TextBox txtStartS 
            Appearance      =   0  'Æò¸é
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   5670
            TabIndex        =   34
            Top             =   150
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.TextBox txtStartR 
            Appearance      =   0  'Æò¸é
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3510
            TabIndex        =   33
            Top             =   135
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.Label Label6 
            Caption         =   "~"
            Height          =   195
            Left            =   6540
            TabIndex        =   38
            Top             =   210
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label5 
            Caption         =   "°Ë»ç¹øÈ£ :"
            Height          =   285
            Left            =   4500
            TabIndex        =   37
            Top             =   210
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.Label Label4 
            Caption         =   "Start Row : "
            Height          =   315
            Left            =   2310
            TabIndex        =   36
            Top             =   210
            Visible         =   0   'False
            Width           =   1215
         End
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Àü¼Û¿Ï·áÈ¯ÀÚ"
         Height          =   345
         Left            =   4890
         TabIndex        =   31
         Top             =   660
         Visible         =   0   'False
         Width           =   1845
      End
      Begin MSComCtl2.DTPicker dtpÁ¢¼öÀÏÀÚ 
         Height          =   315
         Left            =   5340
         TabIndex        =   9
         Top             =   300
         Visible         =   0   'False
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         Format          =   104595457
         CurrentDate     =   39699
      End
      Begin VB.Timer Timer1 
         Interval        =   60000
         Left            =   13020
         Top             =   810
      End
      Begin VB.CommandButton Command2 
         Caption         =   "order test"
         Height          =   465
         Left            =   11940
         TabIndex        =   28
         Top             =   840
         Visible         =   0   'False
         Width           =   2565
      End
      Begin VB.CheckBox Check1 
         Caption         =   "°á°ú ÀÚµ¿ Àü¼Û"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   12810
         TabIndex        =   27
         Top             =   630
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.TextBox txtBarcode 
         Appearance      =   0  'Æò¸é
         BackColor       =   &H00C4FCFD&
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   13140
         TabIndex        =   25
         Top             =   690
         Visible         =   0   'False
         Width           =   2115
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   435
         Left            =   11610
         TabIndex        =   24
         Top             =   660
         Visible         =   0   'False
         Width           =   1485
         _Version        =   65536
         _ExtentX        =   2619
         _ExtentY        =   767
         _StockProps     =   15
         Caption         =   "¹ÙÄÚµå¹øÈ£"
         ForeColor       =   12582912
         BackColor       =   13160660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²Ã¼"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Row:          ~"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   52
         Top             =   360
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.Label Label3 
         Caption         =   "Á¢¼öÀÏÀÚ"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4230
         TabIndex        =   30
         Top             =   330
         Visible         =   0   'False
         Width           =   1035
      End
   End
   Begin VB.TextBox txtTemp 
      Height          =   375
      Left            =   1800
      TabIndex        =   12
      Top             =   1500
      Width           =   2055
   End
   Begin VB.TextBox txtAll 
      Height          =   375
      Left            =   1800
      MultiLine       =   -1  'True
      TabIndex        =   14
      Top             =   2610
      Width           =   2055
   End
   Begin FPSpread.vaSpread vasTemp1 
      Height          =   2355
      Left            =   1650
      TabIndex        =   22
      Top             =   2550
      Visible         =   0   'False
      Width           =   2985
      _Version        =   393216
      _ExtentX        =   5265
      _ExtentY        =   4154
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
      SpreadDesigner  =   "frmInterface.frx":986A
   End
   Begin VB.TextBox txtDate 
      Height          =   405
      Left            =   1140
      TabIndex        =   16
      Top             =   2340
      Width           =   2325
   End
   Begin VB.TextBox txtMsg 
      ForeColor       =   &H000000C0&
      Height          =   735
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   2  '¼öÁ÷
      TabIndex        =   15
      Top             =   8070
      Visible         =   0   'False
      Width           =   5985
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '¾Æ·¡ ¸ÂÃã
      Height          =   375
      Left            =   0
      TabIndex        =   23
      Top             =   10260
      Width           =   15075
      _ExtentX        =   26591
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   6165
            MinWidth        =   6174
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   5467
            MinWidth        =   5467
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "2012-07-10"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "¿ÀÈÄ 3:04"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8820
            MinWidth        =   8820
            Text            =   "¸Þµð¸ÞÀÌÆ® ¢Ï(051)462-1751"
            TextSave        =   "¸Þµð¸ÞÀÌÆ® ¢Ï(051)462-1751"
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
   Begin FPSpread.vaSpread vasRece 
      Height          =   1875
      Left            =   8220
      TabIndex        =   26
      Top             =   5610
      Visible         =   0   'False
      Width           =   5115
      _Version        =   393216
      _ExtentX        =   9022
      _ExtentY        =   3307
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
      SpreadDesigner  =   "frmInterface.frx":9A99
   End
   Begin FPSpread.vaSpread vasTemp 
      Height          =   2565
      Left            =   1650
      TabIndex        =   17
      Top             =   2340
      Visible         =   0   'False
      Width           =   4095
      _Version        =   393216
      _ExtentX        =   7223
      _ExtentY        =   4524
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "frmInterface.frx":9CC8
   End
   Begin FPSpread.vaSpread vasResTemp 
      Height          =   5505
      Left            =   900
      TabIndex        =   21
      Top             =   3030
      Width           =   10875
      _Version        =   393216
      _ExtentX        =   19182
      _ExtentY        =   9710
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
      SpreadDesigner  =   "frmInterface.frx":E1DD
   End
   Begin FPSpread.vaSpread vasPrint 
      Height          =   2295
      Left            =   16560
      TabIndex        =   47
      Top             =   5520
      Visible         =   0   'False
      Width           =   9615
      _Version        =   393216
      _ExtentX        =   16960
      _ExtentY        =   4048
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   16
      MaxRows         =   30
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "frmInterface.frx":E40C
   End
   Begin Threed.SSPanel sspMode 
      Height          =   630
      Left            =   7320
      TabIndex        =   56
      Top             =   45
      Width           =   825
      _Version        =   65536
      _ExtentX        =   1455
      _ExtentY        =   1111
      _StockProps     =   15
      Caption         =   "Àü¼Û¸ðµå"
      ForeColor       =   16777215
      BackColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²Ã¼"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
      BorderWidth     =   5
   End
   Begin MSComCtl2.DTPicker dtpTestDate 
      Height          =   315
      Left            =   5940
      TabIndex        =   57
      Top             =   330
      Width           =   1335
      _ExtentX        =   2355
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
      Format          =   104595457
      CurrentDate     =   39699
   End
   Begin VB.Label Label9 
      Caption         =   "°Ë»çÀÏ"
      BeginProperty Font 
         Name            =   "±¼¸²Ã¼"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5970
      TabIndex        =   58
      Top             =   60
      Width           =   915
   End
   Begin VB.Label lblÀåºñ¸í 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Åõ¸í
      Caption         =   "AU5800 INTERFACE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   420
      TabIndex        =   50
      Top             =   180
      Width           =   2910
   End
   Begin VB.Shape shpCon 
      BackStyle       =   1  'Åõ¸íÇÏÁö ¾ÊÀ½
      FillColor       =   &H00FF0000&
      FillStyle       =   0  '´Ü»ö
      Height          =   435
      Index           =   1
      Left            =   3960
      Top             =   240
      Width           =   135
   End
   Begin VB.Shape shpCon 
      BackStyle       =   1  'Åõ¸íÇÏÁö ¾ÊÀ½
      FillColor       =   &H000000FF&
      FillStyle       =   0  '´Ü»ö
      Height          =   435
      Index           =   0
      Left            =   60
      Top             =   60
      Width           =   135
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Åõ¸í
      Caption         =   "°Ë »ç ÀÚ"
      BeginProperty Font 
         Name            =   "±¼¸²Ã¼"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6450
      TabIndex        =   18
      Top             =   900
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      FillColor       =   &H00FF8080&
      FillStyle       =   0  '´Ü»ö
      Height          =   615
      Index           =   2
      Left            =   60
      Shape           =   4  'µÕ±Ù »ç°¢Çü
      Top             =   60
      Width           =   4035
   End
   Begin VB.Menu mnuFile 
      Caption         =   "ÆÄÀÏ(&F)    "
      Begin VB.Menu mnuExit 
         Caption         =   "Á¾·á(&X)"
      End
   End
   Begin VB.Menu mnuResult 
      Caption         =   "°á°úÀÛ¾÷    "
      Visible         =   0   'False
   End
   Begin VB.Menu mnuWorkList 
      Caption         =   "WorkList    "
   End
   Begin VB.Menu mnuSet 
      Caption         =   "È¯°æ¼³Á¤    "
      Begin VB.Menu mnuSetSub 
         Caption         =   "ÇÁ·ÎÆÄÀÏ¼³Á¤"
         Index           =   0
      End
      Begin VB.Menu mnuSetSub 
         Caption         =   "°Ë»çÄÚµå¼³Á¤"
         Index           =   1
      End
      Begin VB.Menu mnuSetSub 
         Caption         =   "Åë½Å¼³Á¤"
         Index           =   2
      End
   End
   Begin VB.Menu mnuPop 
      Caption         =   "pp"
      Visible         =   0   'False
      Begin VB.Menu subUp 
         Caption         =   "°ËÃ¼¹øÈ£ º¯°æ"
      End
      Begin VB.Menu subDel 
         Caption         =   "°ËÃ¼¹øÈ£ »èÁ¦"
      End
   End
End
Attribute VB_Name = "frmInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const colCheckBox = 1
Const colBARCODE = 2
Const colSampleNo = 3
Const colRack = 4
Const colPos = 5
Const colPID = 6
Const colPName = 7
Const colJumin = 8
Const colPSex = 9
Const colPAge = 10
Const colState = 11
Const colEXAMDATE = 12
Const colSlipNo1 = 13
Const colSlipNo2 = 14
Const colReqDate = 15

Const colEQUIPEXAM = 3
Const colExamCode = 4
Const colExamName = 5
Const colResult = 6
Const colRCheck = 7
Const colPCheck = 8
Const colDCheck = 9
Const colUnit = 10
Const colRef = 11
Const colPanic = 12
Const colResult1 = 13

Dim gRType As String
Dim ConfirmData As String
Dim sBarcode As String
Dim llRow As Long
Dim gRefFlag As String
Dim gPanicFlag As String
Dim SysDateTime As String
Dim TimerFlag As Integer
Dim SubStr(1 To 80) As String


Private Sub chkAll_Click()
    Dim iRow As Integer
    
    If ChkAll.Value = 1 Then
        For iRow = 1 To vasID.DataRowCnt
            '''If Trim(GetText(vasID, iRow, colState)) = "Result" And Trim(GetText(vasID, iRow, colBARCODE)) <> "" Then
            '''If Left(Trim(GetText(vasID, iRow, colBARCODE)), 2) = "09" Then '/¿ï»ê»ê¾÷º¸°Ç¼¾ÅÍ(09)
                vasID.Row = iRow
                vasID.Col = 1
                vasID.Value = 1
            '''End If
        Next iRow
    ElseIf ChkAll.Value = 0 Then
        For iRow = 1 To vasID.DataRowCnt
            vasID.Row = iRow
            vasID.Col = 1
            
            vasID.Value = 0
        Next iRow
    End If
End Sub

Private Sub cmd_Trans_Click()
    '¼±ÅÃÀü¼Û
    Dim vasIDRow As Integer
    Dim vasResRow As Integer
    Dim iRow As Integer
    Dim liRet As Integer
    Dim strSaveYN   As String
    
    If MsgBox(" " & vbCrLf & "¼±ÅÃÀü¼ÛÀ» ÇÏ½Ã°Ú½À´Ï±î?" & vbCrLf & " ", vbInformation + vbOKCancel, "¾Ë¸²:¼±ÅÃÀü¼Û") = vbCancel Then
        Exit Sub
    End If

    For intX = 1 To vasID.MaxRows
        If GET_CELL(vasID, 1, intX) = "1" Then strSaveYN = "Y": Exit For
    Next intX
    
    If strSaveYN <> "Y" Then
        MsgBox "Àü¼ÛÇÒ ÀÚ·á°¡ ¾ø½À´Ï´Ù." & vbCrLf & "¼±ÅÃÀü¼ÛÇÒ ÀÚ·á¸¦ ¼±ÅÃÇÏ½Ê½Ã¿À!", vbCritical, "¼±ÅÃÀü¼Û½ÇÆÐ"
        Exit Sub
    End If
    
'    db_BeginTran gServer
'    Connect_Server
    For vasIDRow = 1 To vasID.DataRowCnt
        vasID.Col = 1
        vasID.Row = vasIDRow
        
        If vasID.Value = 1 Then
            liRet = -1

            If Trim(GetText(vasID, vasIDRow, colBARCODE)) <> "" Then
                liRet = Insert_Data(vasIDRow)
            End If
            
            If liRet = 1 Then
'                db_Commit gServer
                
                SetBackColor vasID, vasIDRow, vasIDRow, colCheckBox, colCheckBox, 202, 255, 112
                SetText vasID, "¿Ï·á", vasIDRow, colState
'                DeleteRow vasID, vasIDRow, vasIDRow
                
            Else
                SetBackColor vasID, vasIDRow, vasIDRow, colCheckBox, colCheckBox, 255, 0, 0
                SetText vasID, "½ÇÆÐ", vasIDRow, colState
            End If
            
'            vasID.Row = vasIDRow
'            vasID.Col = 1
'            vasID.Value = 0
        Else
        
        End If
    Next vasIDRow
    
'    cmdClear_Click
    'db_Commit gServer
End Sub

Function Insert_Data(argSpcRow As Integer) As Integer
    '¼­¹öÀÇ µ¥ÀÌÅ¸ º£ÀÌ½º¿¡ ÀúÀå
    Dim sBarcode As String
    Dim sExamCode As String
    Dim sResult As String
    Dim sEXAMDATE As String
    
    Insert_Data = -1
    
    sBarcode = Trim(GetText(vasID, argSpcRow, colBARCODE))
    
    'Local¿¡¼­ È¯ÀÚº°·Î °á°ú°ª °¡Á®¿À±â
    ClearSpread vasResTemp
    
    SQL = "SELECT EQUIPCODE, examcode, result, EXAMDATE "
    SQL = SQL & "  From PAT_RES "
    SQL = SQL & " WHERE EQUIPNO   = '" & gtypREG_INFO.EQUIPCD & "' "
    SQL = SQL & "   AND EXAMDATE  = '" & Format(dtpExamDate.Value, "yyyymmdd") & "' "
    SQL = SQL & "   AND BARCODE   = '" & sBarcode & "' "
    SQL = SQL & "   AND SEQNO     = '" & Trim(GetText(vasID, argSpcRow, colSampleNo)) & "' "
    SQL = SQL & "   AND RESULT    <> '' "
    res = db_SELECT_Vas(gLocal, SQL, vasResTemp)
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
    vasResTemp.MaxRows = vasResTemp.DataRowCnt + 1
    
    '¼­¹ö·Î °á°ú°ª ÀúÀåÇÏ±â
    For intX = 1 To vasResTemp.DataRowCnt
        sExamCode = Trim(GetText(vasResTemp, intX, 2))
        sResult = Trim(GetText(vasResTemp, intX, 3))
        sEXAMDATE = Trim(GetText(vasResTemp, intX, 4))
        
        '-- 2012.05.09 osw ¼öÁ¤ [ÀåºñÄÚµå 11 : AU640]
        If Mid(sBarcode, 1, 2) <> gID_Par.BARCID Then
            'Save_Result_Data Mid(sBarcode, 1, 2) & "20" & Mid(sBarcode, 3, 8) & sEXAMDATE & Mid(sBarcode, 11, 4) & sExamCode & sResult
            Save_Result_Data gID_Par.BARCID & gID_Par.MACHID & Mid(sBarcode, 3, 8) & Mid(sBarcode, 3, 8) & Mid(sBarcode, 11, 4) & sExamCode & sResult
        Else
            Save_Result_Data Mid(sBarcode, 1, 2) & gID_Par.MACHID & Mid(sBarcode, 3, 8) & Mid(sBarcode, 3, 8) & Mid(sBarcode, 11, 4) & sExamCode & sResult
        End If
    Next intX
    
    SQL = "UPDATE PAT_RES SET "
    SQL = SQL & "       sendflag = '2' "
    SQL = SQL & " WHERE EQUIPNO   = '" & gtypREG_INFO.EQUIPCD & "' "
    SQL = SQL & "   AND EXAMDATE  = '" & Format(dtpExamDate.Value, "yyyymmdd") & "' "
    SQL = SQL & "   AND BARCODE   = '" & sBarcode & "' "
    SQL = SQL & "   AND SEQNO     = '" & Trim(GetText(vasID, argSpcRow, colSampleNo)) & "' "
    res = SendQuery(gLocal, SQL)
    If res = -1 Then
'        db_RollBack gServer
        Exit Function
    End If
         
    Insert_Data = 1
End Function

'''Function Insert_QC_Data(Optional argSqcRow As Integer) As Integer
'''    Dim sBarcode    As String
'''    Dim sEXAMDATE   As String
'''    Dim sEXAMTIME   As String
'''
'''    Insert_QC_Data = -1
'''
'''    If OpenDB(gtypREG_INFO.DB_CONSTR_LOCAL) = False Then Exit Sub   '/Local
'''    If OpenDB2(gtypREG_INFO.DB_CONSTR_QC) = False Then Exit Sub   '/QC
'''
'''    sBarcode = Trim(GET_CELL(vasID, colBARCODE, argSqcRow))
'''
'''    SQL = "SELECT BARCODE, EQUIPCODE, RESULT, RESDATE, EXAMDATE, EXAMTIME "
'''    SQL = SQL & vbCrLf & "  FROM PAT_RES "
'''    SQL = SQL & vbCrLf & " WHERE equipno   = '" & gtypREG_INFO.EQUIPCD & "' "
'''    SQL = SQL & vbCrLf & "   AND BARCODE   = '" & Trim(sBarcode) & "' "
'''    SQL = SQL & vbCrLf & "   AND EXAMDATE  = '" & Trim(Format(dtpExamDate.Value, "yyyymmdd")) & "' "
'''    SQL = SQL & vbCrLf & "   AND PID       = 'QC' "
'''    If ReadSQL(gstrQuy, ADR) = False Then
'''        Call CloseDB
'''        Call CloseDB2
'''        Exit Function
'''    End If
'''
'''    If Not ADR Is Nothing Then
'''        ADC2.BeginTrans
'''
'''        Do Until ADR.EOF
'''            sEXAMDATE = Trim(ADR!EXAMDATE & "")
'''            sEXAMTIME = Trim(ADR!EXAMTIME & "")
'''
'''            If sEXAMDATE = "" Then
'''                sEXAMDATE = Format(dtpExamDate.Value, "YYYYMMDD")
'''            End If
'''
'''            If sEXAMTIME = "" Then
'''                sEXAMTIME = Format(Time, "HHMMSS")
'''            End If
'''
'''            SQL = "Delete From qc_res "
'''            SQL = SQL & vbCrLf & " WHERE equipno   = '" & gtypREG_INFO.EQUIPCD & "' "
'''            SQL = SQL & vbCrLf & "   And EXAMDATE  = '" & sEXAMDATE & "' "
'''            SQL = SQL & vbCrLf & "   And EXAMTIME  = '" & sEXAMTIME & "' "
'''            SQL = SQL & vbCrLf & "   And levelname = '" & sBarcode & "' "
'''            SQL = SQL & vbCrLf & "   And equipcode = '" & Trim(ADR!equipcode & "") & "' "
'''            If RunSQL2(gstrQuy) = False Then
'''                ADC2.RollbackTrans
'''                Call CloseDB
'''                Call CloseDB2
'''                Call ErrQuery(gstrQuy, 0)
'''                Exit Function
'''            End If
'''
'''            SQL = "Insert into qc_res "
'''            SQL = SQL & vbCrLf & " (equipno,    EXAMDATE,   EXAMTIME,   levelname,  equipcode, "
'''            SQL = SQL & vbCrLf & "  result,     resflag,    examuid,    examuname ) "
'''            SQL = SQL & vbCrLf & " values "
'''            SQL = SQL & vbCrLf & " ('" & gtypREG_INFO.EQUIPCD & "', "
'''            SQL = SQL & vbCrLf & "  '" & sEXAMDATE & "', "
'''            SQL = SQL & vbCrLf & "  '" & sEXAMTIME & "', "
'''            SQL = SQL & vbCrLf & "  '" & sBarcode & "', "
'''            SQL = SQL & vbCrLf & "  '" & Trim(ADR!equipcode & "") & "', "
'''            SQL = SQL & vbCrLf & "  '" & Trim(ADR!Result & "") & "', "
'''            SQL = SQL & vbCrLf & "  '', "
'''            SQL = SQL & vbCrLf & "  'AUTO', "
'''            SQL = SQL & vbCrLf & "  'AUTO' )"
'''            If RunSQL2(gstrQuy) = False Then
'''                ADC2.RollbackTrans
'''                Call CloseDB
'''                Call CloseDB2
'''                Call ErrQuery(gstrQuy, 0)
'''                Exit Function
'''            End If
'''
'''            ADR.MoveNext
'''        Loop
'''
'''        ADC2.CommitTrans
'''
'''    End If
'''
'''    Call CloseDB
'''    Call CloseDB2
'''
'''    Insert_QC_Data = 1
'''End Function

Function CheckValue(asResult As String, asExamCode As String, asAge As String, asSex As String, asRegion As String, asDate As String)
    Dim sRefHigh As String
    Dim sRefLow As String
    Dim sPanicHigh As String
    Dim sPanicLow As String
    Dim i As Integer
    SQL = "SELECT sclvalue, schvalue, plvalue, phvalue from tl_standard " & vbCrLf & _
          "WHERE workname = '" & asExamCode & "' and region = '" & asRegion & "' " & vbCrLf & _
          "and f_age <= '" & asAge & "' and t_age >= '" & asAge & "'" & vbCrLf & _
          "and  f_date <= '" & asDate & "' and t_date >= '" & asDate & "' "
    res = db_SELECT_Col(gServer, SQL)
    
    If gReadBuf(0) = "" And gReadBuf(1) = "" Then
        If asSex = "M" Then
            SQL = "SELECT smlvalue, smhvalue, plvalue, phvalue from tl_standard " & vbCrLf & _
                  "WHERE workname = '" & asExamCode & "' and region = '" & asRegion & "' " & vbCrLf & _
                  "and f_age <= '" & asAge & "' and t_age >= '" & asAge & "' " & vbCrLf & _
                  "and  f_date <= '" & asDate & "' and t_date >= '" & asDate & "' "
            res = db_SELECT_Col(gServer, SQL)
        Else
            SQL = "SELECT sflvalue, sfhvalue, plvalue, phvalue from tl_standard " & vbCrLf & _
                  "WHERE workname = '" & asExamCode & "' and region = '" & asRegion & "' " & vbCrLf & _
                  "and f_age <= '" & asAge & "' and t_age >= '" & asAge & "' " & vbCrLf & _
                  "and  f_date <= '" & asDate & "' and t_date >= '" & asDate & "' "
            res = db_SELECT_Col(gServer, SQL)
        End If
        For i = 0 To 3
            If gReadBuf(i) = "" Then
                gReadBuf(i) = "0"
            End If
        Next
        sRefHigh = gReadBuf(1)
        sRefLow = gReadBuf(0)
        sPanicLow = gReadBuf(2)
        sPanicHigh = gReadBuf(3)
    Else
        For i = 0 To 3
            If gReadBuf(i) = "" Then
                gReadBuf(i) = "0"
            End If
        Next
        sRefHigh = gReadBuf(1)
        sRefLow = gReadBuf(0)
        sPanicLow = gReadBuf(2)
        sPanicHigh = gReadBuf(3)
    End If
    gRefFlag = ""
    gPanicFlag = ""
    If IsNumeric(asResult) = False Or IsNumeric(sRefLow) = False Or IsNumeric(sRefHigh) = False Or IsNumeric(sPanicLow) = False Or IsNumeric(sPanicHigh) = False Then
        Exit Function
    End If
    
    If CCur(asResult) < CCur(sRefLow) Then
        gRefFlag = "L"
    End If
    If CCur(asResult) > CCur(sRefHigh) Then
        gRefFlag = "H"
    End If
    If CCur(asResult) < CCur(sPanicLow) Then
        gPanicFlag = "L"
    End If
    If CCur(asResult) > CCur(sPanicHigh) Then
        gPanicFlag = "H"
    End If
    
End Function

Private Sub cmd°ËÃ¼¹øÈ£»ý¼º_Click()
    Dim intX        As Integer
    Dim lsBARCODE   As String
    
    For intX = 1 To vasID.MaxRows
        If Trim(GetText(vasID, intX, colSampleNo)) <> "" And Trim(GetText(vasID, intX, colBARCODE)) = "" Then
            lsBARCODE = txtReceHead & Trim(GetText(vasID, intX, colSampleNo))
            
            SetText vasID, lsBARCODE, intX, colBARCODE
    
            SQL = "UPDATE PAT_RES SET "
            SQL = SQL & vbCrLf & "       BARCODE   = '" & lsBARCODE & "' "
            SQL = SQL & vbCrLf & " WHERE EXAMDATE  = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' "
            SQL = SQL & vbCrLf & "   AND EQUIPNO   = '" & gEquip & "' "
            SQL = SQL & vbCrLf & "   AND SEQNO     = '" & Trim(GetText(vasID, intX, colSampleNo)) & "' "
            res = SendQuery(gLocal, SQL)
        End If
    Next intX
End Sub

Private Sub cmd°á°ú»èÁ¦_Click()
    Dim strSaveYN   As String
    
    For intX = 1 To vasID.MaxRows
        If GET_CELL(vasID, 1, intX) = "1" Then strSaveYN = "Y": Exit For
    Next intX
    
    If strSaveYN <> "Y" Then
        MsgBox "»èÁ¦ÇÒ ÀÚ·á°¡ ¾ø½À´Ï´Ù." & vbCrLf & "»èÁ¦ÇÒ ÀÚ·á¸¦ ¼±ÅÃÇÏ½Ê½Ã¿À!", vbCritical, "°á°ú»èÁ¦½ÇÆÐ"
        Exit Sub
    End If
    
    If MsgBox("°á°ú»èÁ¦¸¦ ÇÏ°Ú½À´Ï±î?", vbQuestion + vbOKCancel, "°á°ú»èÁ¦È®ÀÎ") = vbCancel Then Exit Sub

    For intX = 1 To vasID.MaxRows
        If GET_CELL(vasID, 1, intX) = "1" Then
            SQL = "DELETE FROM PAT_RES "
            SQL = SQL & " WHERE EQUIPNO   = '" & gtypREG_INFO.EQUIPCD & "' "
            SQL = SQL & "   AND EXAMDATE  = '" & Format(dtpExamDate.Value, "yyyymmdd") & "' "
            SQL = SQL & "   AND BARCODE   = '" & Trim(GetText(vasID, intX, colBARCODE)) & "' "
            SQL = SQL & "   AND SEQNO     = '" & Trim(GetText(vasID, intX, colSampleNo)) & "' "
            res = SendQuery(gLocal, SQL)
        End If
    Next intX
    
    Call cmdCall_Click
    
    MsgBox "»èÁ¦ ¿Ï·á", vbInformation, "È®ÀÎ"
End Sub

Private Sub cmdCall_Click()
    Dim i As Integer
    
    '''ClearSpread vasID
    If vasID.MaxRows > 0 Then vasID.MaxRows = 0
    
    SQL = "SELECT '', BARCODE, SEQNO, MAX(DISKNO), MAX(POSNO), MAX(pid), MAX(pname), MAX(pjumin), MAX(psex), MAX(page), MAX(sendflag) "
    SQL = SQL & "  FROM PAT_RES "
    SQL = SQL & " WHERE EQUIPNO  = '" & gtypREG_INFO.EQUIPCD & "' "
    SQL = SQL & "   AND EXAMDATE = '" & Format(dtpExamDate.Value, "yyyymmdd") & "' "
    SQL = SQL & " GROUP BY BARCODE, SEQNO "
    SQL = SQL & " ORDER BY BARCODE, SEQNO "
    res = db_SELECT_Vas(gLocal, SQL, vasID)

    vasID.MaxRows = vasID.DataRowCnt
    For i = 1 To vasID.DataRowCnt
'        vasID.RowHeight(i) = 13
'        gReadBuf(0) = ""

'        SQL = "SELECT refflag from PAT_RES WHERE BARCODE = '" & Trim(GetText(vasID, i, colBARCODE)) & "' and refflag in ('S', 'B')"
'        res = db_SELECT_Col(gLocal, SQL)
'
'        If gReadBuf(0) = "S" Or gReadBuf(0) = "B" Then
'            SetForeColor vasID, i, i, 0, 0, 255
'        End If

        gReadBuf(0) = ""
        SQL = "SELECT refflag from PAT_RES WHERE BARCODE = '" & Trim(GetText(vasID, i, colBARCODE)) & "' and refflag = 'R'"
        res = db_SELECT_Col(gLocal, SQL)

        If gReadBuf(0) = "R" Then
            SetForeColor vasID, i, i, 255, 0, 0
        End If

        gReadBuf(0) = ""
        SQL = "SELECT MAX(panicflag) "
        SQL = SQL & "  from PAT_RES "
        SQL = SQL & " WHERE EXAMDATE  = '" & Format(dtpExamDate.Value, "yyyymmdd") & "' "
        SQL = SQL & "   AND BARCODE   = '" & Trim(GetText(vasID, i, colBARCODE)) & "' "
        SQL = SQL & "   AND SEQNO     = '" & Trim(GetText(vasID, i, colSampleNo)) & "' "
        res = db_SELECT_Col(gLocal, SQL)

        If gReadBuf(0) = "L" Or gReadBuf(0) = "H" Then
            vasID.Row = i
            vasID.Col = -1
            vasID.ForeColor = RGB(255, 0, 0)
            vasID.FontBold = True
        End If

        Select Case GetText(vasID, i, colState)
            Case "0"
                SetText vasID, "W/S", i, colState
                SetBackColor vasID, i, i, 1, 1, 255, 250, 205
            Case "1"
                SetText vasID, "Result", i, colState
                SetBackColor vasID, i, i, 1, 1, 255, 250, 205
            Case "2"
                SetText vasID, "¿Ï·á", i, colState
                SetBackColor vasID, i, i, colCheckBox, colCheckBox, 202, 255, 112
        End Select
    Next
End Sub

Private Sub cmdClear_Click()
    Dim iRow As Integer

    txtMsg.Text = ""
    
''''    ClearSpread vasID, 1, 1
''''    vasID.MaxRows = 0
'''
'''    For iRow = 1 To vasID.DataRowCnt
'''        vasID.Row = iRow
'''        vasID.Col = 1
'''
'''        If vasID.Value = 1 Then
'''            vasDeleteRow vasID, iRow
'''
'''            iRow = iRow - 1
'''        End If
'''    Next iRow
'''
'''    ClearSpread vasRes, 1, 1
    vasRes.MaxRows = 0
    vasID.MaxRows = 0
End Sub

Private Sub cmdEnd_Click()
    Unload Me
End Sub

Private Sub cmdRowSet_Click()
    If Trim(txtRowF) = "" Then MsgBox "Row(From)´Â ¼ýÀÚÇüÀ¸·Î ÀÔ·ÂÇØ¾ß ÇÕ´Ï´Ù!", vbCritical, "Àû¿ë½ÇÆÐ": txtRowF.SetFocus: Exit Sub
    If Trim(txtRowT) = "" Then MsgBox "Row(To)´Â ¼ýÀÚÇüÀ¸·Î ÀÔ·ÂÇØ¾ß ÇÕ´Ï´Ù!", vbCritical, "Àû¿ë½ÇÆÐ": txtRowT.SetFocus: Exit Sub
    If IsNumeric(txtRowF) = False Then MsgBox "Row(From)´Â ¼ýÀÚÇüÀ¸·Î ÀÔ·ÂÇØ¾ß ÇÕ´Ï´Ù!", vbCritical, "Àû¿ë½ÇÆÐ": txtRowF.SetFocus: Exit Sub
    If IsNumeric(txtRowT) = False Then MsgBox "Row(To)´Â ¼ýÀÚÇüÀ¸·Î ÀÔ·ÂÇØ¾ß ÇÕ´Ï´Ù!", vbCritical, "Àû¿ë½ÇÆÐ": txtRowT.SetFocus: Exit Sub
    
    If Val(txtRowF) < 1 Then MsgBox "Row(From)´Â 1 Row ÀÌ»óÀÌ¾î¾ß ÇÕ´Ï´Ù!", vbCritical, "Àû¿ë½ÇÆÐ": txtRowF.SetFocus: Exit Sub
    If Val(txtRowT) < 1 Then MsgBox "Row(To)´Â 1 Row ÀÌ»óÀÌ¾î¾ß ÇÕ´Ï´Ù!", vbCritical, "Àû¿ë½ÇÆÐ": txtRowT.SetFocus: Exit Sub
    
    If Val(txtRowF) > Val(txtRowT) Then
        MsgBox "Row ¹üÀ§¸¦ (Àç)ÀÔ·ÂÇÏ½Ê½Ã¿À", vbCritical, "Àû¿ë½ÇÆÐ": txtRowF.SetFocus: Exit Sub
    End If
        
    For intX = Val(txtRowF) To Val(txtRowT)
        If intX > vasID.MaxRows Then Exit For
        Call SET_CELL(vasID, 1, intX, "1")
    Next intX
End Sub

Private Sub Command1_Click()
    
    Call AU5800("DQ", "       Q002                      01E1002       35  04       97  06      199  08      1.9  10      4.1  11      1.1  12      0.3  16       54  20      3.3  22       99  24      7.0  ")
End Sub

Private Sub Command2_Click()
    Dim i As Long
    Dim PJumin As String
    Dim SubDateTime As String
    Dim EndRow As Integer
    Dim subDate As String
    Dim j As Integer
    
    'Àåºñ¹øÈ£,¹ÙÄÚµå¹øÈ£,È¯ÀÚÀÌ¸§,³ªÀÌ,¼ºº°,¿À´õ,È¯ÀÚ¹øÈ£
    ClearSpread vasOrder1
    Connect_Server
    SubDateTime = Format(dtpÁ¢¼öÀÏÀÚ, "yyyy/mm/dd") & " " & Format(Time, "hh:mm:ss")
    subDate = Format(dtpÁ¢¼öÀÏÀÚ, "yyyy/mm/dd") & " " & "23:59:59"
    SQL = "SELECT '', a.sample, b.name, '', b.sex, '', a.hospno, b.jumin " & vbCrLf & _
          "from tl_workhead a, tb_idmast b, tl_workvalue c " & vbCrLf & _
          "WHERE a.hospno = b.hospno and a.sample = c.sample " & vbCrLf & _
          "and c.workname in (" & gAllExam & ") " & vbCrLf & _
          "and c.repdate between to_date('" & SysDateTime & "', 'yyyy/mm/dd hh24:mi:ss') and to_date('" & subDate & "', 'yyyy/mm/dd hh24:mi:ss') GROUP BY a.sample, b.name,  b.sex, a.hospno, b.jumin"
'
    res = db_SELECT_Vas(gServer, SQL, vasOrder1)
    If res = -1 Then
        Save_Raw_Data res & ">" & SQL
        Exit Sub
    End If
    
    SysDateTime = SubDateTime
    j = vasOrder1.DataRowCnt
    For i = 1 To vasOrder1.DataRowCnt
        CalAgeSex Trim(GetText(vasOrder1, i, 8)), dtpExamDate.Value
        SetText vasOrder1, gPatGen.Age, i, 4
        j = j - 1
        If Check2.Value = 0 Then
            SQL = "SELECT levelname from qc_res WHERE levelname = '" & Trim(GetText(vasOrder1, i, 2)) & "'"
            res = db_SELECT_Col(gLocal, SQL)

            If res > 0 Then
                DeleteRow vasOrder1, i, i
                i = i - 1

            ElseIf res = 0 Then
    '            SQL = "SELECT max(resflag) from qc_res WHERE EXAMDATE = '" & Format(dtpEXAMDATE.Value, "yyyymmdd") & "'"
    '            res = db_SELECT_Col(gLocal, SQL)
    '
    '            If Trim(gReadBuf(0)) = "" Then
    '                gReadBuf(0) = 1
    '            Else
    '                gReadBuf(0) = CInt(gReadBuf(0)) + 1
    '            End If

                SQL = "insert into qc_res(levelname, EXAMDATE) " & vbCrLf & _
                      "values('" & Trim(GetText(vasOrder1, i, 2)) & "', '" & Format(dtpExamDate.Value, "yyyymmdd") & "')"
                res = SendQuery(gLocal, SQL)


                SQL = "SELECT workname from tl_workvalue WHERE sample = '" & Trim(GetText(vasOrder1, i, 2)) & "' and workname in ('Basohpil','Lym','Monocyte','Seg/Neu','Eosinophil')"
                res = db_SELECT_Col(gServer, SQL)
                If res > 0 Then
                    SetText vasOrder1, "B", i, 6
                Else
                    SetText vasOrder1, "A", i, 6
                End If
                
            ElseIf res = -1 Then
                Save_Raw_Data "QueryErr>" & SQL
                ClearSpread vasOrder1
                Exit Sub
            
            End If
            If j = 0 Then
                Exit For
            End If
        Else
            SQL = "SELECT workname from tl_workvalue WHERE sample = '" & Trim(GetText(vasOrder1, i, 2)) & "' and workname in ('Basohpil','Lym','Monocyte','Seg/Neu','Eosinophil')"
            res = db_SELECT_Col(gServer, SQL)
            If res > 0 Then
                SetText vasOrder1, "B", i, 6
            Else
                SetText vasOrder1, "A", i, 6
            End If
        End If
    Next

    If vasOrder1.DataRowCnt > 0 Then
        MSComm1.Output = Chr(1)
        Save_Raw_Data "[TX]" & Chr(1)
        EndRow = vasOrder1.DataRowCnt + 1
        SetText vasOrder1, "ÿ END     " & Chr(13) & "", EndRow, 1
        SetText vasOrder1, "END", EndRow, 2
    End If
End Sub


Private Sub Command3_Click()
'''Dim sGubun string
'''Dim LineData string
Dim sGubun As String
Dim LineData As String

    sGubun = Mid(Text1.Text, 1, 2)     '/¼ö½ÅÀÚ·á ¼º°Ý±¸ºÐÀÚ(2 Byte)
    LineData = Mid(Text1.Text, 3) '/¼ö½ÅÀÚ·á º»¹®³»¿ª
    Call AU5800(sGubun, LineData)
    Text1.Text = ""
End Sub

'Private Sub dtpÁ¢¼öÀÏÀÚ_Change()
'Dim i As Integer
'    SysDateTime = Format(dtpÁ¢¼öÀÏÀÚ, "yyyy/mm/dd") & " 00:00:00"
'    Command2_Click
'End Sub

Private Sub dtpÁ¢¼öÀÏÀÚ_CloseUp()
Dim i As Integer
    SysDateTime = Format(dtpÁ¢¼öÀÏÀÚ, "yyyy/mm/dd") & " 00:00:00"
    txtReceHead.Text = "09" & Format(dtpÁ¢¼öÀÏÀÚ, "yyyymmdd")
    txtPrint = "09" & Format(dtpÁ¢¼öÀÏÀÚ, "yyyymmdd")
'    Command2_Click
End Sub

Private Sub Form_Activate()
    'txtBARCODE.SetFocus
End Sub

Private Sub Form_Load()
    Dim sDate As String
    Dim FindFile As String
    
    '1. È­¸é ¹× º¯¼ö ÃÊ±âÈ­
    '2. µ¥ÀÌÅ¸º£ÀÌ½º¿¡ Connect ÇÏ±â - Local - Server
    '3. Ini ³»¿ë ºÒ·¯¿À±â    GetSetup
    '4. Comport Open

On Error GoTo errFind
    

    Me.Left = 0
    Me.Top = 0
    
    gAllExam = ""
    
    'cmdClear_Click
    
    ClearSpread vasID, 1, 1
    'vasID.MaxRows = 1
    
    GetSetup    'ini¿¡¼­ DBÁ¤º¸ ºÒ·¯¿À±â
    
    gtypREG_INFO.EQUIPCD = "AU5800"
    gtypREG_INFO.DB_CONSTR_LOCAL = "Provider=Microsoft.Jet.OLEDB.4.0;User ID=;Password=;" & _
                                   "Data Source=" & App.Path & "\Interface.mdb;" & _
                                   "Persist Security Info=True"
    
    'gtypREG_INFO.DB_CONSTR_QC = "driver={SQL Server};" & _
    '                            "server=" & "»ýÈ­ÇÐÀÎÅÍÆäÀÌ\SQLEXPRESS" & ";" & _
    '                            "uid=sa;" & _
    '                            "pwd=mate;" & _
    '                            "database=MMQC_mini"
'''    gtypREG_INFO.DB_CONSTR_QC = "driver={SQL Server};" & _
                                "server=" & "Localhost\SQLEXPRESS" & ";" & _
                                "uid=sa;" & _
                                "pwd=mate;" & _
                                "database=MMQC_mini"
                            
    '¼­¹ö¿¡ Á¢¼Ó
'    If Not Connect_Server Then
'        MsgBox "¿¬°áµÇÁö ¾Ê¾Ò½À´Ï´Ù."
'        Exit Sub
'    End If
    
    '·ÎÄÃ¿¡ Á¢¼Ó
    If Not Connect_Local Then
        MsgBox "¿¬°áµÇÁö ¾Ê¾Ò½À´Ï´Ù."
        Exit Sub
    End If
    
    'Comport Setting
    MSComm1.CommPort = gSetup.gPort
    MSComm1.RTSEnable = gSetup.gRTSEnable
    MSComm1.DTREnable = gSetup.gDTREnable
    MSComm1.Settings = gSetup.gSpeed & "," & gSetup.gParity & "," & gSetup.gDataBit & "," & gSetup.gStopBit
    
    If MSComm1.PortOpen = False Then
         MSComm1.PortOpen = True
    End If
    
    StatusBar1.Panels(1).Text = "[ ÄÄÆ÷Æ® ¼³Á¤ ] " & MSComm1.Settings
    
    raw_data = ""
    dtpÁ¢¼öÀÏÀÚ = Format(Date, "yyyy-mm-dd")
    dtpExamDate = Date
    dtpTestDate = Date
    
    '====================·ÎÄÃ DBÁö¿ì±â - 10ÀÏ º¸°ü======================
    sDate = Format(DateAdd("y", CDate(dtpExamDate.Value), -10), "yyyymmdd")
    
    SQL = "Delete from PAT_RES WHERE EXAMDATE < '" & sDate & "' "
    SendQuery gLocal, SQL
    
    SQL = "Delete from qc_res WHERE EXAMDATE < '" & sDate & "' "
    SendQuery gLocal, SQL
    '===================================================================
    
    '°Ë»çÄÚµå °¡Á®¿À±â
    GetExamCode
      
    '2006/11/20 ÀÌ»óÀº
    SQL = " Alter Table PAT_RES " & CR & _
          " Alter Column result text(50) "
    res = SendQuery(gLocal, SQL)
    
    
    'MultiSELECT Mode
    vasRes.OperationMode = 1
    SysDateTime = Format(Date, "yyyy/mm/dd") & " 00:00:00"
    TimerFlag = 1
    txtReceHead.Text = "09" & Format(Date, "yyyymmdd")
    txtPrint.Text = "09" & Format(Date, "yyyymmdd")
    txtStartR.Text = "1"
    
    txtRowF = ""
    txtRowT = ""
    
errFind:
'2005/06/16 ÀÌ»óÀº
    If Err = 8002 Then      'Port
        MsgBox "Åë½Å Æ÷Æ®¸¦ È®ÀÎÇÏ¼¼¿ä!", vbExclamation, "¾Ë¸²"
        Call mnuSetSub_Click(2) '/Åë½Å¼³Á¤
    Else
        Resume Next
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
    
    DisConnect_Server
    DisConnect_Local
End Sub

Sub GetExamCode()
'°Ë»çÄÚµå¸¦ array¿¡ ÀúÀå
    Dim i As Integer
    Dim j As Integer
    
    gAllExam = ""
    
    ClearSpread vasTemp
    
    SQL = "SELECT EQUIPCODE, examcode, examname From EQUIPEXAM " & CR & _
          " WHERE EQUIPNO = '" & gEquip & "' " & CR & _
          " ORDER BY SEQNO"
          
    res = db_SELECT_Vas(gLocal, SQL, vasTemp)

    If res > 0 Then
        ReDim gArr_ExamCode(1 To vasTemp.DataRowCnt, 1 To 3)
    Else
        SaveQuery SQL
        Exit Sub
    End If
        
    For i = 1 To vasTemp.DataRowCnt
        gArr_ExamCode(i, 1) = i
        For j = 1 To 2
            gArr_ExamCode(i, j + 1) = Trim(GetText(vasTemp, i, j))
        Next j
        
        If gAllExam = "" Then
            gAllExam = "'" & Trim(GetText(vasTemp, i, 2)) & "'"
        Else
            gAllExam = gAllExam & ", '" & Trim(GetText(vasTemp, i, 2)) & "'"
        End If
    Next i
    
End Sub


Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuSetSub_Click(Index As Integer)
    Select Case Index
        Case 0: '/ÇÁ·ÎÆÄÀÏ¼³Á¤
            frmÇÁ·ÎÆÄÀÏ.Show 1
        Case 1: '/°Ë»çÄÚµå¼³Á¤
            frmEquipExam.SSPanel1.Caption = "  AU5800 Àåºñ ÄÚµå ¼³Á¤"
            frmEquipExam.Show 1
            GetExamCode
        
        Case 2: '/Åë½Å¼³Á¤
            frmConfig.SSPanel_machine.Caption = "AU5800"
            frmConfig.Show 1

    End Select
End Sub

Private Sub mnuWorkList_Click()
    frmWorkList.Show vbModal
End Sub

Private Sub MSComm1_OnComm()
    Dim S           As String
    Dim i           As Integer
    Dim sGubun      As String
    Dim sMode       As Integer
    Dim LineData    As String
    
    S = MSComm1.Input
    Debug.Print S
    Select Case S
        Case chrSTX 'Chr(2)
            If Right(txtBuff, 1) = "" Then
                sMode = 1
                i = InStr(1, txtBuff, "")
                txtBuff = Mid(txtBuff, 1, i - 1)
            Else
                sMode = 0
                txtBuff = ""
            End If
        
        Case chrETX
            Save_Raw_Data "[RX" & CDate(Time) & "]" & txtBuff.Text & S
            
            sGubun = Mid(txtBuff, 1, 2)     '/¼ö½ÅÀÚ·á ¼º°Ý±¸ºÐÀÚ(2 Byte)
            LineData = Mid(txtBuff.Text, 3) '/¼ö½ÅÀÚ·á º»¹®³»¿ª
            Call AU5800(sGubun, LineData)
    
        Case Else
            If sMode = 1 Then
                If S = "E" Then
                    sMode = 0
                End If
            Else
                txtBuff = txtBuff & S
            End If
    End Select
End Sub

Sub SendOrder()
    If gOrderMessage <> "" Then
        gPreData = gOrderMessage
        gOrderMessage = ""
        
        MSComm1.Output = gPreData
        Debug.Print gPreData
        Save_Raw_Data "[TX" & CDate(Time) & "]" & gPreData
    End If
End Sub

Private Sub AU5800(asGubun As String, asData As String)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
    Dim iRow As Integer
    Dim llRow As Integer
    Dim liRet As Integer
    
    Dim lResRow As Long         '°á°ú°ü·Ã Row
    
    Dim lsRackNo As String
    Dim lsPos As String
    
    Dim lsSampleType As String
    Dim lsSampleNo As String
    Dim lsSampleID As String
    Dim lsID As String
    Dim lsPID As String
    
    Dim sExamCode As String
    Dim sSubCode As String
    Dim sExamName As String
    
    Dim lsCode As String
    Dim lsRt As String
    Dim lsFlag As String
    
    Dim lsSEQNO As String
    
    Dim sEXAMDATE As String
    Dim sEXAMTIME As String
    Dim sDate As String

    Dim lsData As String
    
    Dim iCnt As Integer
    Dim iExamCnt As Integer
    Dim sAllResult As String
    
    Dim iLen As String
    
    Dim lsControlNo As String
    Dim lsLotNo As String
    Dim lsLevel As String
    Dim lsLevelName As String
    
    sEXAMDATE = Format(dtpÁ¢¼öÀÏÀÚ.Value, "yyyymmdd")
    sEXAMTIME = Format(Time, "hhmmss")
    
    sDate = sEXAMDATE & " " & sEXAMTIME
    
    lsSampleNo = Trim(Mid(asData, 8, 4))
    
    Select Case asGubun
        Case "R ": GoSub RTN_PROCESS_R  '/Inquery Order
        Case "D ": GoSub RTN_PROCESS_D  '/Result
'''        Case "DQ": GoSub RTN_PROCESS_DQ '/QC Result
        Case "DA": GoSub RTN_PROCESS_DA '/
        Case "DB"   'Result Start
        Case "DE"   'Result End
    End Select
Exit Sub

'/========================================================================================================================================================================================================/

RTN_PROCESS_R:
    lsSampleID = ""
    lsData = ""
    
    If OpenDB(gtypREG_INFO.DB_CONSTR_LOCAL) = False Then Exit Sub

    gstrQuy = "SELECT A.BARCODE, A.EQUIPCODE, B.ORDERENABLE "
    gstrQuy = gstrQuy & vbCrLf & "  FROM PAT_RES A, EQUIPEXAM B "
    gstrQuy = gstrQuy & vbCrLf & " WHERE A.EQUIPNO       = '" & gtypREG_INFO.EQUIPCD & "'"
    gstrQuy = gstrQuy & vbCrLf & "   AND A.EXAMDATE      = '" & Format(CDate(dtpTestDate.Value), "yyyymmdd") & "'"
    gstrQuy = gstrQuy & vbCrLf & "   AND A.SEQNO         = '" & Trim(lsSampleNo) & "'"
    gstrQuy = gstrQuy & vbCrLf & "   AND A.EQUIPCODE   = B.EQUIPCODE "
    gstrQuy = gstrQuy & vbCrLf & "   AND A.EQUIPNO    = B.EQUIPNO "
    
    '°Ë»çÄÚµå¸¦ ¿Ö ÀúÀå¾ÈÇß´ÂÁö ¸ð¸£°ÚÀ½.. È®ÀÎÈÄ Ã³¸®
    'gstrQuy = gstrQuy & vbCrLf & "   AND A.EXAMCODE    = B.EXAMCODE "
    '¿À´õÁÖ±â Ã³¸®ÇßÀ½
    'gstrQuy = gstrQuy & vbCrLf & "   AND A.EQUIPCODE <= '30' " '/ÀåºñÄÚµå 30 ÀÌ»óÀÌ¸é °è»êµÇ¾îÁö´Â °Ë»çÇ×¸ñÀÌ¹Ç·Î ¿À´õ »ý¼º½Ã Á¦¿Ü.
    If ReadSQL(gstrQuy, ADR) = False Then
        Call CloseDB
        Exit Sub
    End If
    
    If Not ADR Is Nothing Then
        lsSampleID = Trim(ADR!barcode & "")
        Do Until ADR.EOF
            If Trim(ADR!ORDERENABLE) = "1" Then
                lsData = lsData & Format(Trim(ADR!EQUIPCODE & ""), "00")  '& " "
            End If
            ADR.MoveNext
        Loop
        ADR.Close: Set ADR = Nothing
    End If
    Call CloseDB
    
'    lsData = "030"
    
'    gOrderMessage = chrSTX & _
                    "S " & _
                    Left(asData, 11) & TEXT_RSET(lsSampleID, 20) & Space(4) & "E" & "000000" & _
                    lsData & _
                    chrETX
    
    '-- AU5800
    gOrderMessage = chrSTX & _
                    "S " & _
                    Left(asData, 11) & TEXT_RSET(lsSampleID, 20) & Space(4) & "E" & _
                    lsData & _
                    chrETX
    
'    GetOrder = STX & "S " & mRackNo & mCupPos & mBarNo & mSampleNo & strOrder & ETX
    
    '-- AU5800
'    gOrderMessage = chrSTX & _
                    "S " & _
                    Left(asData, 11) & TEXT_RSET(lsSampleID, 20) & Space(4) & "E" & _
                    lsData & _
                    chrETX
    
    Call SendOrder
    
    'gOrderMessage = chrSTX & "SE" & chrETX
    'Call SendOrder
Return

'/========================================================================================================================================================================================================/

RTN_PROCESS_D:
    lsRackNo = Mid(asData, 1, 4)
    lsPos = Mid(asData, 5, 2)
    
    lsSampleType = Mid(asData, 7, 1)
    
    lsSampleID = Trim(Mid(asData, 15, 20))
    lsID = CStr(lsSampleID)
    
    '°°Àº ¹ÙÄÚµå¹øÈ£ÀÇ °ËÃ¼´Â µð½ºÇÃ·¹ÀÌµÇÁö ¾ÊÀ½
    glRow = -1
    For iRow = 1 To vasID.DataRowCnt
        If Trim(GetText(vasID, iRow, colBARCODE)) = lsID And Trim(GetText(vasID, iRow, colSampleNo)) = lsSampleNo Then
            glRow = iRow
            Exit For
        End If
    Next iRow

    If glRow = -1 Then      'vasID¿¡ ¾ø´Â °ËÃ¼ÀÇ °á°ú°¡ ³ª¿Ã ¶§ µ¥ÀÌÅÍ Ãß°¡
        glRow = vasID.DataRowCnt + 1
        If glRow > vasID.MaxRows Then
            vasID.MaxRows = glRow + 1
        End If
        vasActiveCell vasID, glRow, colBARCODE
    End If
    
    SetText vasID, lsID, glRow, colBARCODE
     
    'È¯ÀÚÁ¤º¸
'''    If Trim(GetText(vasID, glRow, colPID)) = "" Then
'''        Get_Sample_Info lsID, glRow
'''    End If
    
    '°á°ú µð½ºÇÃ·¹ÀÌ
    vasActiveCell vasID, glRow, colBARCODE
    
    ClearSpread vasRes, 1, 1
    
    SetText vasID, lsSampleID, glRow, colBARCODE
    SetText vasID, lsSampleNo, glRow, colSampleNo
    SetText vasID, lsRackNo, glRow, colRack
    SetText vasID, lsPos, glRow, colPos
    
    '¼ö½ÅÁß========================================================
    SetText vasID, "¼ö½ÅÁß", glRow, colState
    SetBackColor vasID, glRow, glRow, 1, 1, 255, 250, 205
    '==============================================================
    
    vasRes.MaxRows = 0
    
    '°á°ú Àß¶ó ³Ö±â
    j = 0
                                
'''    If Trim(Mid(asData, 36, 2)) = "E0" Or Trim(Mid(asData, 36, 2)) = "00" Then
'''        lsData = Trim(Mid(asData, 43)) '/°Ë»ç°á°úÇ×¸ñÀÌ ±æ¸é µÎ¹øÂ° °á°ú ¹ÞÀ»¶§ E·Î Ç¥±âµÊ.
'''    Else
'''        lsData = Trim(Mid(asData, 37))
'''    End If
    If Trim(Mid(asData, 36, 1)) = "E" And Trim(Mid(asData, 37, 2)) <> "0" Then
'        lsData = Trim(Mid(asData, 37))
        lsData = Trim(Mid(asData, 38))
    Else
'        lsData = Trim(Mid(asData, 43))
        lsData = Trim(Mid(asData, 44))
    End If
'''    lsData = Trim(Mid(asData, 37))
    
    
'''    If Trim(Mid(asData, 36, 2)) = "E" Or Trim(Mid(asData, 36, 2)) = "0" Then
'''        lsData = Trim(Mid(asData, 37))
'''    Else
'''        lsData = Trim(Mid(asData, 37))
'''    End If
    
    'D 001101 0001      09201105300001    E01       45H 02       23  03      278H 04       96  05      184  06      126  07       44  08      0.9  41    114.8  
    'D 001001 0001                        E0     01     13.0r 02      2.5Gr03     22.0r 04-  0.8r 05 219.8Hr06  69.1Hr07-113.9Lr08   3.0r 09  10.2r 10   5.1@r11   5.1@r12   1.0@r14   0.1r 18  15.7r 24   9.4r 25 100.1Hr
    
    Do While Len(lsData) >= 5
        lsCode = Trim(Left(lsData, 2))
        lsRt = Trim(Mid(lsData, 3, 1)) & Trim(Mid(lsData, 4, 7)) '/Parameter->Online->Protocal Tab(Text Format->Data Format->9ÀÚ¸®)

        lsRt = Replace(lsRt, "D", "")
        lsRt = Replace(lsRt, "@", "")
        lsRt = Replace(lsRt, "r", "")
        lsRt = Replace(lsRt, "*", "")
        lsRt = Replace(lsRt, "e", "")
        lsRt = Replace(lsRt, "?", "")
        
        'Dim lngVal As Long
        Dim strVal As String
        
        strVal = ""
        
        For k = 1 To Len(lsRt)
            Select Case Mid(lsRt, k, 1)
            Case ".", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
                strVal = strVal & Mid(lsRt, k, 1)
            Case Else
            End Select
        Next
        
        lsRt = strVal
        
        Select Case Trim(lsCode)
        'T.P,ALB,D-Bil,BUN,Crea,UA,Ca,IP,T-Bil,CRP,RA
        Case "04", "05", "06", "12", "13", "14", "18", "19", "20", "21", "22", "24", "25", "26", "27"
            lsRt = Format(lsRt, "##0.0")
            'lsRt = Round(lsRt, 1)
        Case Else
            lsRt = Format(lsRt, "##0")
            'lsRt = Round(lsRt, 0)
        End Select
        
        
        
'''        lsFlag = Trim(Mid(lsData, 12, 2))
'''        If lsFlag = "%?" Then
'''            lsRt = ""
'''            lsFlag = ""
'''        Else
'''            i = InStr(1, lsFlag, "r")
'''            If i = 0 Then
'''
'''            Else
'''                lsFlag = Left(lsFlag, i - 1)
'''            End If
'''
'''        End If
        
        '°á°ú µð½ºÇÃ·¹ÀÌ
        SQL = "SELECT examcode, examname, SEQNO, paniclow, panichigh from EQUIPEXAM WHERE EQUIPCODE = '" & Trim(lsCode) & "'"
        res = db_SELECT_Col(gLocal, SQL)
        If res > 0 Then
            vasRes.MaxRows = vasRes.MaxRows + 1
            lResRow = vasRes.MaxRows
            
            SetText vasRes, lsID, lResRow, colBARCODE
            SetText vasRes, lsCode, lResRow, colEQUIPEXAM           'ÀåºñÄÚµå
            SetText vasRes, Trim(gReadBuf(0)), lResRow, colExamCode '°Ë»çÄÚµå
            SetText vasRes, Trim(gReadBuf(1)), lResRow, colExamName '°Ë»ç¸í
            SetText vasRes, lsRt, lResRow, colResult                '°Ë»ç°á°ú
            SetText vasRes, lsFlag, lResRow, colRCheck              'ÆÇÁ¤

            If gReadBuf(3) <> "" Then
                If Val(lsRt) <= Val(gReadBuf(3)) Then
                    SetText vasRes, "L", lResRow, colPCheck              'Panic(Low)
                
                    vasRes.Row = lResRow
                    vasRes.Col = 8
                    vasRes.ForeColor = RGB(65, 105, 225)
                
                    vasID.Row = glRow
                    vasID.Col = -1
                    vasID.ForeColor = RGB(255, 0, 0)
                    vasID.FontBold = True
                End If
            End If
            
            If gReadBuf(4) <> "" Then
                If Val(lsRt) >= Val(gReadBuf(4)) Then
                    SetText vasRes, "H", lResRow, colPCheck              'Panic(High)
                
                    vasRes.Row = lResRow
                    vasRes.Col = 8
                    vasRes.ForeColor = RGB(205, 55, 0)
                
                    vasID.Row = glRow
                    vasID.Col = -1
                    vasID.ForeColor = RGB(255, 0, 0)
                    vasID.FontBold = True
                End If
            End If
            
            Save_Local_One glRow, lResRow, "1"
            
            j = j + 1
            

            SetText vasID, "Result", glRow, colState
            SetBackColor vasID, glRow, glRow, 1, 1, 0, 128, 64
        End If
               
'        lsData = Mid(lsData, 14)
        lsData = Mid(lsData, 12)
        
        If Mid(lsData, 1, 1) = "D" Then     'ETB
            'lsData = Mid(lsData, 37)
            lsData = Mid(lsData, 39)
        End If
    Loop
Return

'/========================================================================================================================================================================================================/

RTN_PROCESS_DQ:
    lsControlNo = Trim(Mid(asData, 34, 2))
    
    lsSampleID = lsSampleNo & lsControlNo
    lsID = lsSampleID
    
    '°°Àº ¹ÙÄÚµå¹øÈ£ÀÇ °ËÃ¼´Â µð½ºÇÃ·¹ÀÌµÇÁö ¾ÊÀ½
    glRow = -1
    For iRow = 1 To vasID.DataRowCnt
        If Trim(GetText(vasID, iRow, colBARCODE)) = lsID And Trim(GetText(vasID, iRow, colSampleNo)) = lsSampleNo Then
            glRow = iRow
            Exit For
        End If
    Next iRow

    If glRow = -1 Then      'vasID¿¡ ¾ø´Â °ËÃ¼ÀÇ °á°ú°¡ ³ª¿Ã ¶§ µ¥ÀÌÅÍ Ãß°¡
        glRow = vasID.DataRowCnt + 1
        If glRow > vasID.MaxRows Then
            vasID.MaxRows = glRow + 1
        End If
        vasActiveCell vasID, glRow, colBARCODE
    End If
    
    SetText vasID, lsID, glRow, colBARCODE
    
    '°á°ú µð½ºÇÃ·¹ÀÌ
    vasActiveCell vasID, glRow, colBARCODE
    
    ClearSpread vasRes, 1, 1
    
    SetText vasID, lsSampleID, glRow, colBARCODE
    SetText vasID, lsSampleNo, glRow, colSampleNo
    SetText vasID, "QC", glRow, colPID
'''    SetText vasID, lsRackNo, glRow, colRack
'''    SetText vasID, lsPos, glRow, colPos
    
    '¼ö½ÅÁß========================================================
    SetText vasID, "¼ö½ÅÁß", glRow, colState
    SetBackColor vasID, glRow, glRow, 1, 1, 255, 250, 205
    '==============================================================
    
    vasRes.MaxRows = 0
    
    '°á°ú Àß¶ó ³Ö±â
    j = 0
                                
    lsData = Trim(Mid(asData, 39))
    
    sEXAMDATE = Format(dtpTestDate.Value, "YYYYMMDD")
    sEXAMTIME = Format(Time, "HHMMSS")
    
    If OpenDB2(gtypREG_INFO.DB_CONSTR_QC) = False Then Exit Sub   '/QC
    ADC2.BeginTrans
    
    Do While Len(lsData) >= 5
        lsCode = Trim(Left(lsData, 2))
        lsRt = Trim(Mid(lsData, 3, 9)) '/Parameter->Online->Protocal Tab(Text Format->Data Format->9ÀÚ¸®)
        
        '°á°ú µð½ºÇÃ·¹ÀÌ
        SQL = "SELECT examcode, examname, SEQNO from EQUIPEXAM WHERE EQUIPCODE = '" & Trim(lsCode) & "'"
        res = db_SELECT_Col(gLocal, SQL)
        If res > 0 Then
'''            vasRes.MaxRows = vasRes.MaxRows + 1
'''            lResRow = vasRes.MaxRows
'''
'''            SetText vasRes, lsID, lResRow, colBARCODE
'''            SetText vasRes, lsCode, lResRow, colEQUIPEXAM           'ÀåºñÄÚµå
'''            SetText vasRes, Trim(gReadBuf(0)), lResRow, colExamCode '°Ë»çÄÚµå
'''            SetText vasRes, Trim(gReadBuf(1)), lResRow, colExamName '°Ë»ç¸í
'''            SetText vasRes, lsRt, lResRow, colResult                '°Ë»ç°á°ú
'''            SetText vasRes, lsFlag, lResRow, colRCheck              'ÆÇÁ¤
'''            Save_Local_One glRow, lResRow, "1"
            
            SQL = "Delete From qc_res "
            SQL = SQL & vbCrLf & " WHERE equipno   = '" & gtypREG_INFO.EQUIPCD & "' "
            SQL = SQL & vbCrLf & "   And EXAMDATE  = '" & sEXAMDATE & "' "
            SQL = SQL & vbCrLf & "   And EXAMTIME  = '" & sEXAMTIME & "' "
            SQL = SQL & vbCrLf & "   And levelname = '" & lsSampleID & "' "
            SQL = SQL & vbCrLf & "   And equipcode = '" & lsCode & "' "
            If RunSQL2(SQL) = False Then
                ADC2.RollbackTrans
                Call CloseDB2
                Call ErrQuery(gstrQuy, 0)
                Exit Sub
            End If

            SQL = "Insert into qc_res "
            SQL = SQL & vbCrLf & " (equipno,    EXAMDATE,   EXAMTIME,   levelname,  equipcode, "
            SQL = SQL & vbCrLf & "  result,     resflag,    examuid,    examuname ) "
            SQL = SQL & vbCrLf & " values "
            SQL = SQL & vbCrLf & " ('" & gtypREG_INFO.EQUIPCD & "', "
            SQL = SQL & vbCrLf & "  '" & sEXAMDATE & "', "
            SQL = SQL & vbCrLf & "  '" & sEXAMTIME & "', "
            SQL = SQL & vbCrLf & "  '" & lsSampleID & "', "
            SQL = SQL & vbCrLf & "  '" & lsCode & "', "
            SQL = SQL & vbCrLf & "  '" & lsRt & "', "
            SQL = SQL & vbCrLf & "  '', "
            SQL = SQL & vbCrLf & "  'AUTO', "
            SQL = SQL & vbCrLf & "  'AUTO')"
            If RunSQL2(SQL) = False Then
                ADC2.RollbackTrans
                Call CloseDB2
                Call ErrQuery(gstrQuy, 0)
                Exit Sub
            End If
            
'''            Call Insert_QC_Data(lResRow)
'''
'''            j = j + 1
'''
'''
'''            SetText vasID, "¼ö½Å¿Ï·á", glRow, colState
'''            SetBackColor vasID, glRow, glRow, 1, 1, 0, 128, 64
        End If
               
        lsData = Mid(lsData, 14)
    Loop
    
    ADC2.CommitTrans
    Call CloseDB2
    
    SetText vasID, "¼ö½Å¿Ï·á", glRow, colState
    SetBackColor vasID, glRow, glRow, 1, 1, 0, 128, 64
Return

'/========================================================================================================================================================================================================/

RTN_PROCESS_DA:
    lsControlNo = Trim(Mid(asData, 34, 2))
    
    lsSampleID = lsSampleNo & lsControlNo
    lsID = lsSampleID
    
    '°°Àº ¹ÙÄÚµå¹øÈ£ÀÇ °ËÃ¼´Â µð½ºÇÃ·¹ÀÌµÇÁö ¾ÊÀ½
    glRow = -1
    For iRow = 1 To vasID.DataRowCnt
        If Trim(GetText(vasID, iRow, colBARCODE)) = lsID And Trim(GetText(vasID, iRow, colSampleNo)) = lsSampleNo Then
            glRow = iRow
            Exit For
        End If
    Next iRow

    If glRow = -1 Then      'vasID¿¡ ¾ø´Â °ËÃ¼ÀÇ °á°ú°¡ ³ª¿Ã ¶§ µ¥ÀÌÅÍ Ãß°¡
        glRow = vasID.DataRowCnt + 1
        If glRow > vasID.MaxRows Then
            vasID.MaxRows = glRow + 1
        End If
        vasActiveCell vasID, glRow, colBARCODE
    End If
    
    SetText vasID, lsID, glRow, colBARCODE
    
    '°á°ú µð½ºÇÃ·¹ÀÌ
    vasActiveCell vasID, glRow, colBARCODE
    
    ClearSpread vasRes, 1, 1
    
    SetText vasID, lsSampleID, glRow, colBARCODE
    SetText vasID, lsSampleNo, glRow, colSampleNo
    SetText vasID, "DA", glRow, colPID
'''    SetText vasID, lsRackNo, glRow, colRack
'''    SetText vasID, lsPos, glRow, colPos
    
    '¼ö½ÅÁß========================================================
    SetText vasID, "¼ö½ÅÁß", glRow, colState
    SetBackColor vasID, glRow, glRow, 1, 1, 255, 250, 205
    '==============================================================
    
    vasRes.MaxRows = 0
    
    '°á°ú Àß¶ó ³Ö±â
    j = 0
'''
'''    lsData = Trim(Mid(asData, 39))
    
    lsData = Trim(Mid(asData, 43))
    
    Do While Len(lsData) >= 5
        lsCode = Trim(Left(lsData, 2))
        lsRt = Trim(Mid(lsData, 3, 1)) & Trim(Mid(lsData, 4, 8)) '/Parameter->Online->Protocal Tab(Text Format->Data Format->9ÀÚ¸®)
        
        '°á°ú µð½ºÇÃ·¹ÀÌ
        SQL = "SELECT examcode, examname, SEQNO from EQUIPEXAM WHERE EQUIPCODE = '" & Trim(lsCode) & "'"
        res = db_SELECT_Col(gLocal, SQL)
        If res > 0 Then
            vasRes.MaxRows = vasRes.MaxRows + 1
            lResRow = vasRes.MaxRows

            SetText vasRes, lsID, lResRow, colBARCODE
            SetText vasRes, lsCode, lResRow, colEQUIPEXAM           'ÀåºñÄÚµå
            SetText vasRes, Trim(gReadBuf(0)), lResRow, colExamCode '°Ë»çÄÚµå
            SetText vasRes, Trim(gReadBuf(1)), lResRow, colExamName '°Ë»ç¸í
            SetText vasRes, lsRt, lResRow, colResult                '°Ë»ç°á°ú
            SetText vasRes, lsFlag, lResRow, colRCheck              'ÆÇÁ¤

            Save_Local_One glRow, lResRow, "1"
            
            j = j + 1

            SetText vasID, "¼ö½Å¿Ï·á", glRow, colState
            SetBackColor vasID, glRow, glRow, 1, 1, 0, 128, 64
        End If
               
        lsData = Mid(lsData, 14)
    Loop
Return
End Sub

Sub Save_Result_Data(argSQL As String)
'argSQLÀÇ ³»¿ëÀ» ÆÄÀÏ·Î ÀúÀå
    Dim FilNum
    Dim sFileName As String
    
    FilNum = FreeFile
    
    '-- °á°úÀúÀåÆú´õ º¯°æ[¼­¿ïÁöºÎ]
'    If Dir("C:\AU5800Result", vbDirectory) <> "AU5800Result" Then
'        MkDir ("C:\AU5800Result")
'    End If
    
    If Dir("C:\Kiha2002\AU5800", vbDirectory) <> "AU5800" Then
        MkDir ("C:\Kiha2002\AU5800")
    End If
    
    sFileName = Format(CDate(dtpExamDate.Value), "yyyymmdd")
    
'    Open App.Path & "\Result\" & sFileName & ".txt" For Output As FilNum
'    Open "C:\AU5800Result\" & sFileName & ".txt" For Append As FilNum
    Open "C:\Kiha2002\AU5800\" & sFileName & ".txt" For Append As FilNum
    Print #FilNum, argSQL
    Close FilNum
End Sub

Sub Save_Raw_Data(argSQL As String)
'argSQLÀÇ ³»¿ëÀ» ÆÄÀÏ·Î ÀúÀå
    Dim FilNum
    Dim sFileName As String
    
    FilNum = FreeFile
    
    If Dir(App.Path & "\Result", vbDirectory) <> "Result" Then
        MkDir (App.Path & "\Result")
    End If
    
    sFileName = Format(CDate(dtpExamDate.Value), "yyyymmdd")
    
'    Open App.Path & "\Result\" & sFileName & ".txt" For Output As FilNum
    Open App.Path & "\Result\" & sFileName & ".txt" For Append As FilNum
    Print #FilNum, argSQL
    Close FilNum
End Sub

Function Get_Sample_Info(ByVal asRow As Long) As Integer
    Dim lsBARCODE As String
    Dim lsSEQNO As String
    Dim lsDate As String
    
    'Á¢¼öÀÏÀÚ,Á¢¼ö¹øÈ£·Î »ùÇÃ È¯ÀÚ Á¤º¸ °¡Á®¿À±â
    lsBARCODE = Trim(GetText(vasID, asRow, colBARCODE))   '»ùÇÃ ¹ÙÄÚµå ¹øÈ£
    lsDate = ""
    lsDate = Format(Trim(dtpExamDate.Value), "yyyymmdd")
    
    lsSEQNO = ""
    lsSEQNO = Trim(GetText(vasID, asRow, colRack))
    

'    SQL = "SELECT a.ptno, b.sname, a.sex, a.ageyy, a.JEOBSUDT, a.slipno1, a.slipno2 from twexam_general_sub a, tw_mis_pmpa.twbas_patient b " & vbCrLf & _
'          "WHERE a.ptno = b.ptno and a.ptno = '" & lsBARCODE & "' and jeobsudt = to_date('" & Format(Date, "yyyymmdd") & "', 'yyyy/mm/dd hh24/mi/ss') and itemcd in (" & gAllExam & ") ORDER BY slipno2"
          
          
    SQL = "SELECT a.hospno, b.name, b.sex, a.requestdate, b.jumin" & vbCrLf & _
          "from tl_workhead a, tb_idmast b " & vbCrLf & _
          "WHERE a.sample = '" & lsBARCODE & "' and a.hospno = b.hospno"
    res = db_SELECT_Col(gServer, SQL)
    
    If res = 1 Then
        SetText vasID, Trim(gReadBuf(0)), asRow, colPID
        
        SetText vasID, Trim(gReadBuf(1)), asRow, colPName
'        SetText vasID, Trim(gReadBuf(4)), asRow, colJumin
        
        CalAgeSex Trim(gReadBuf(4)), dtpExamDate.Value
        SetText vasID, gPatGen.Age, asRow, colPAge
        SetText vasID, Trim(gReadBuf(2)), asRow, colPSex
        
        
        SetText vasID, Format(Trim(gReadBuf(3)), "yyyymmdd"), asRow, colReqDate
'        SetText vasID, Trim(gReadBuf(5)), asRow, colSlipNo1
'        SetText vasID, Trim(gReadBuf(6)), asRow, colSlipNo2
'
'        SetText vasID, Trim(gReadBuf(6)), asRow, colBARCODE
    Else
        SetText vasID, "", asRow, colPID
        SetText vasID, "", asRow, colPName
'        SetText vasID, "", asRow, colJumin
        SetText vasID, "", asRow, colPSex
        SetText vasID, "0", asRow, colPAge
        SetText vasID, "", asRow, colReqDate
'        SetText vasID, "", asRow, colBARCODE
    End If
    
    gReadBuf(0) = ""
    gReadBuf(1) = ""
    gReadBuf(2) = ""
    gReadBuf(3) = ""
    
End Function

Function SetResult(asResult As String, aiItem As Integer) As String
'DB¿¡¼­ ºÒ·¯¿À±â
    Dim iFloat As Integer
    
    If Not IsNumeric(asResult) Then
        Exit Function
    End If

    Select Case aiItem
    Case 7, 16
        iFloat = 2
    Case 14
        iFloat = 0
    Case Else
        iFloat = 1
    End Select

    If iFloat = 0 Then
        SetResult = CStr(CCur(asResult))
    Else
        SetResult = CStr(CCur(Left(asResult, 5 - iFloat)) & "." & Right(asResult, iFloat))
    End If
 
End Function

Private Sub sspMode_Click()
    If sspMode.Caption = "¼öÁ¤¸ðµå" Then
        sspMode.Caption = "Àü¼Û¸ðµå"
        sspMode.BackColor = &HFF0000
        sspMode.ForeColor = &HFFFFFF
        vasRes.OperationMode = 1
        
    ElseIf sspMode.Caption = "Àü¼Û¸ðµå" Then
        sspMode.Caption = "¼öÁ¤¸ðµå"
        sspMode.BackColor = &H8000&
        sspMode.ForeColor = &HFFFFFF
        vasRes.OperationMode = 0
        
        vasActiveCell vasRes, 1, colResult
        vasRes.SetFocus
    End If
End Sub

Private Sub subDel_Click()
    Dim i As Long
    
    i = vasID.ActiveRow
    
    vasID.DeleteRows i, 1
    If i > vasID.DataRowCnt Then
        i = vasID.DataRowCnt
    End If
    vasID.MaxRows = vasID.DataRowCnt
    vasActiveCell vasID, i, colBARCODE
    vasID.SetFocus
End Sub

Private Sub subUp_Click()
Dim sValue As String
Dim sTmp As String
Dim i As Integer
Dim j As Integer

    sTmp = ""
    
    vasID.Row = vasID.ActiveRow
    vasID.Col = vasID.ActiveCol
    
    sTmp = vasID.Text
    
    sValue = InputBox("º¯°æÇÒ °ËÃ¼¹øÈ£¸¦ ÀÔ·ÂÇÏ¼¼¿ä")
        
    If Trim(sValue) <> "" Then
        If MsgBox("" & sTmp & "¸¦ " & sValue & "·Î ¼öÁ¤ÇÏ½Ã°Ú½À´Ï±î?", vbYesNo, "È®ÀÎ") = vbYes Then
            SetText vasID, sValue, vasID.Row, vasID.Col
            
            If Trim(GetText(vasID, vasID.Row, colBARCODE)) <> "" Then
                '''Get_Sample_Info vasID.Row
                            
                For i = 1 To vasRes.DataRowCnt
                    Save_Local_One vasID.Row, i, "A"
                Next
            End If
        End If
    End If

End Sub

Private Sub txtPS_GotFocus()
    SELECTFocus txtPS
End Sub

Private Sub txtPS_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txtPS = "" Then
            txtPS.SetFocus
            Exit Sub
        End If
        If IsNumeric(txtPS) = False Then
            txtPS.SetFocus
            Exit Sub
        End If
        
        txtPS.Text = Format(Trim(txtPS.Text), "000#")
        
        txtPE.SetFocus
    End If
End Sub

Private Sub txtPE_GotFocus()
    SELECTFocus txtPE
End Sub

Private Sub txtPE_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim lsBARCODE As String
    Dim lRow As Long
    
    If KeyCode = vbKeyReturn Then
        If txtPE = "" Then
            txtPE.SetFocus
            Exit Sub
        End If
        If IsNumeric(txtPE) = False Then
            txtPE.SetFocus
            Exit Sub
        End If
        txtPE.Text = Format(Trim(txtPE.Text), "000#")
        txtResPrint.SetFocus
    End If
End Sub

Private Sub txtResPrint_Click()

frmPrint.Show 1
    
End Sub

Private Sub txtRowF_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub txtRowF_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtRowT_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub txtRowT_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call cmdRowSet_Click
End Sub

Private Sub txtStartR_GotFocus()
    SELECTFocus txtStartR
End Sub

Private Sub txtStartR_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txtStartR = "" Then
            txtStartR.SetFocus
            Exit Sub
        End If
        If IsNumeric(txtStartR) = False Then
            txtStartR.SetFocus
            Exit Sub
        End If
        txtStartS.SetFocus
    End If
End Sub

Private Sub txtStartS_GotFocus()
    SELECTFocus txtStartS
End Sub

Private Sub txtStartS_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txtStartS = "" Then
            txtStartS.SetFocus
            Exit Sub
        End If
        If IsNumeric(txtStartS) = False Then
            txtStartS.SetFocus
            Exit Sub
        End If
        
        txtResN.SetFocus
    End If
End Sub

Private Sub txtResN_GotFocus()
    SELECTFocus txtResN
End Sub

Private Sub txtResN_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim lsBARCODE As String
    Dim lRow As Long
    
    If KeyCode = vbKeyReturn Then
        If txtResN = "" Then
            txtResN.SetFocus
            Exit Sub
        End If
        If IsNumeric(txtResN) = False Then
            txtResN.SetFocus
            Exit Sub
        End If
    'vasid ¿¡ Á¢¼ö¹øÈ£ ÀÔ·ÂÈÄ ÀúÀå
    
        For i = 1 To CInt(txtResN.Text) - CInt(txtStartS.Text) + 1
            lsBARCODE = txtReceHead & Format(CInt(txtStartS) + i - 1, "000#")
            lRow = CInt(txtStartR.Text) + i - 1
'            txtStartS.Text = CInt(txtStartS.Text) + 1
            
            If Trim(GetText(vasID, lRow, colRack)) = "" Then
            Else
                 SetText vasID, lsBARCODE, lRow, colBARCODE
    
    '            Get_Sample_Info llRow
    
                SQL = "UPDATE PAT_RES SET "
                SQL = SQL & vbCrLf & "       BARCODE   = '" & lsBARCODE & "' "
                SQL = SQL & vbCrLf & " WHERE EXAMDATE  = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' "
                SQL = SQL & vbCrLf & "   AND EQUIPNO   = '" & gEquip & "' "
                SQL = SQL & vbCrLf & "   AND SEQNO     = '" & Trim(GetText(vasID, lRow, colSampleNo)) & "' "
                res = SendQuery(gLocal, SQL)
            End If
           
            
        Next
        
        txtStartR.SetFocus
    End If
End Sub

'Private Sub Timer1_Timer()
'    If TimerFlag < 5 Then
'        TimerFlag = TimerFlag + 1
'        Exit Sub
'    ElseIf TimerFlag >= 5 Then
'        TimerFlag = 1
'        Command2_Click
'    End If
'End Sub

Private Sub txtBARCODE_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim iRow As Integer
    Dim lRow As Integer
    
    If KeyCode = vbKeyReturn Then
        If txtBarcode.Text = "" Then
            Exit Sub
        End If
            
        lRow = -1
        For iRow = 1 To vasID.DataRowCnt
            If txtBarcode.Text = Trim(GetText(vasID, iRow, colBARCODE)) Then
                lRow = iRow
                Exit For
            End If
        Next iRow
        

        If lRow = -1 Then
            lRow = vasID.DataRowCnt + 1
            If lRow > vasID.MaxRows Then
                vasID.MaxRows = lRow
            End If
        End If
            
        SetText vasID, Trim(txtBarcode.Text), lRow, colBARCODE
        
        'È¯ÀÚ¹øÈ£, Å¸±ÞÁ¾°ü·Ã È¯ÀÚÁ¤º¸, È¯ÀÚÀÌ¸§, ¼ºº°, ³ªÀÌ
        SQL = " SELECT distinct scp41idnoa, scp41idnob, scp41name, scp41sex, scp41birth, scp41jdate " & vbCrLf & _
              " From scprst41 " & vbCrLf & _
              " WHERE scp41pcode = '75' " & vbCrLf & _
              " And SCP41SPMNO2 = '" & Trim(txtBarcode.Text) & "'"
        res = db_SELECT_Col(gServer, SQL)
        
        If res = 1 Then
            SetText vasID, Trim(gReadBuf(0)), lRow, colPID
            
            SetText vasID, Trim(gReadBuf(2)), lRow, colPName
            SetText vasID, Trim(gReadBuf(4)), lRow, colJumin
            
            CalAgeSex Trim(gReadBuf(4)), dtpExamDate.Value
            SetText vasID, Trim(gReadBuf(3)), lRow, colPSex
            SetText vasID, gPatGen.Age, lRow, colPAge
            
            SetText vasID, Trim(gReadBuf(5)), lRow, colReqDate
            
        Else
            SetText vasID, "", lRow, colPName
            SetText vasID, "", lRow, colJumin
            SetText vasID, "", lRow, colPSex
            SetText vasID, "", lRow, colPAge
            SetText vasID, "", lRow, colReqDate
        End If
        
        txtBarcode.Text = ""
        txtBarcode.SetFocus
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
    Dim lsTmpID As String
    
    Dim i As Integer
    
    If Row < 1 Or Row > vasID.DataRowCnt Then
        Exit Sub
    End If
    
    lsID = Trim(GetText(vasID, Row, colBARCODE))

    ClearSpread vasRes
    vasRes.MaxRows = 0
    
    SQL = "SELECT '', a.BARCODE, a.EQUIPCODE,  a.examcode, a.examname, a.result, a.refflag, a.panicflag, a.deltaflag, a.unit, a.refvalue, a.panicvalue, a.result "
    SQL = SQL & "  FROM PAT_RES a, EQUIPEXAM b"
'''    SQL = SQL & " WHERE A.examcode  = B.examcode "
    SQL = SQL & " WHERE A.EQUIPCODE = B.EQUIPCODE "
    SQL = SQL & "   AND A.EQUIPNO   = '" & gtypREG_INFO.EQUIPCD & "' "
    SQL = SQL & "   AND A.EXAMDATE  = '" & Format(dtpExamDate.Value, "yyyymmdd") & "' "
    SQL = SQL & "   AND A.BARCODE   = '" & Trim(GetText(vasID, vasID.Row, colBARCODE)) & "' "
    SQL = SQL & "   AND A.SEQNO     = '" & Trim(GetText(vasID, vasID.Row, colSampleNo)) & "' "
    SQL = SQL & " ORDER BY B.SEQNO " '/°Ë»çÇ×¸ñ ¼ø¼­
          
    res = db_SELECT_Vas(gLocal, SQL, vasRes)
    If res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If
'    For i = 1 To vasRes.DataRowCnt
'        vasRes.RowHeight(i) = 13
'    Next
    
    For i = 1 To vasRes.DataRowCnt
'        'ÂüÁ¶Ä¡
'        SELECT Case Trim(GetText(vasRes, i, colRCheck))
'        Case "H"
'            vasRes.Row = i
'            vasRes.Col = 7
'            vasRes.ForeColor = RGB(205, 55, 0)
'        Case "L"
'            vasRes.Row = i
'            vasRes.Col = 7
'            vasRes.ForeColor = RGB(65, 105, 225)
'        Case ""
'             vasRes.Row = i
'            vasRes.Col = 7
'            vasRes.ForeColor = RGB(255, 255, 255)
'        End SELECT
'
        'Panic
        Select Case Trim(GetText(vasRes, i, 8))
        Case "H"
            vasRes.Row = i
            vasRes.Col = 8
            vasRes.ForeColor = RGB(205, 55, 0)
        Case "L"
            vasRes.Row = i
            vasRes.Col = 8
            vasRes.ForeColor = RGB(65, 105, 225)
        Case ""
             vasRes.Row = i
            vasRes.Col = 8
            vasRes.ForeColor = RGB(255, 255, 255)
        End Select
'
'        'Delta
'        SELECT Case Trim(GetText(vasRes, i, 9))
'        Case "D"
'            vasRes.Row = i
'            vasRes.Col = 9
'            vasRes.ForeColor = RGB(205, 55, 0)
'        Case "L"
'            vasRes.Row = i
'            vasRes.Col = 9
'            vasRes.ForeColor = RGB(65, 105, 225)
'        Case ""
'             vasRes.Row = i
'            vasRes.Col = 9
'            vasRes.ForeColor = RGB(255, 255, 255)
'        End SELECT
    Next i

End Sub

Function Save_Local_One(ByVal asRow1 As Long, ByVal asRow2 As Long, asSend As String)
    Dim sCnt As String
    Dim sEXAMDATE As String
    
'    If Trim(GetText(vasID, asRow1, colSEQNO)) = "QC" Then
'        sEXAMDATE = Trim(GetText(vasID, asRow1, colEXAMDATE))
'
'        '2004/05/28 ÀÌ»óÀº
'        'sEXAMDATE = Left(sEXAMDATE, 4) & "-" & Mid(sEXAMDATE, 5, 2) & "-" & Mid(sEXAMDATE, 7, 2) & " " & Mid(sEXAMDATE, 9, 2) & ":" & Mid(sEXAMDATE, 11, 2) & ":00"
'    Else
'        sEXAMDATE = GetDateFull
'    End If
    
    sCnt = ""
    SQL = "DELETE FROM PAT_RES "
    SQL = SQL & vbCrLf & " WHERE EQUIPNO   = '" & gtypREG_INFO.EQUIPCD & "' "
    SQL = SQL & vbCrLf & "   AND EXAMDATE  = '" & Format(CDate(dtpTestDate.Value), "yyyymmdd") & "' "
    SQL = SQL & vbCrLf & "   AND BARCODE   = '" & Trim(GetText(vasID, asRow1, colBARCODE)) & "' "
    SQL = SQL & vbCrLf & "   AND SEQNO     = '" & Trim(GetText(vasID, asRow1, colSampleNo)) & "' "
    SQL = SQL & vbCrLf & "   AND EQUIPCODE = '" & Trim(GetText(vasRes, asRow2, colEQUIPEXAM)) & "' "
'    SaveQuery SQL
    res = SendQuery(gLocal, SQL)
'    If res = -1 Then
'        SaveQuery SQL
'        Exit Function
'    End If
    
    If Not IsNumeric(GetText(vasID, asRow1, colPAge)) Then
        SetText vasID, "0", asRow1, colPAge
    End If
'    If Not IsDate(Trim(GetText(vasExam, asRow, colEXAMDATE))) Then
'        SetText vasExam, "1900-01-01", asRow, colEXAMDATE
'    End If
    
    SQL = "INSERT INTO PAT_RES "
    SQL = SQL & vbCrLf & " (EXAMDATE,   EQUIPNO,    BARCODE,    receno,     pid, "
    SQL = SQL & vbCrLf & "  pname,      pjumin,     page,       psex,       resdate, "
    SQL = SQL & vbCrLf & "  EQUIPCODE,  examcode,   examtype,   result,     sendflag, "
    SQL = SQL & vbCrLf & "  examname,   refflag,    panicflag,  deltaflag,  unit, "
    SQL = SQL & vbCrLf & "  refvalue,   panicvalue, SEQNO,      diskno,     posno) "
    SQL = SQL & vbCrLf & " VALUES "
    SQL = SQL & vbCrLf & " ('" & Format(CDate(dtpTestDate.Value), "yyyymmdd") & "', " '/EXAMDATE
    SQL = SQL & vbCrLf & "  '" & Trim(gtypREG_INFO.EQUIPCD) & "', " '/EQUIPNO
    SQL = SQL & vbCrLf & "  '" & Trim(GetText(vasID, asRow1, colBARCODE)) & "', " '/BARCODE
    SQL = SQL & vbCrLf & "  '', " '/receno
    SQL = SQL & vbCrLf & "  '" & Trim(GetText(vasID, asRow1, colPID)) & "', " '/pid
    SQL = SQL & vbCrLf & "  '" & Trim(GetText(vasID, asRow1, colPName)) & "', " '/pname
    SQL = SQL & vbCrLf & "  '" & Trim(GetText(vasID, asRow1, colJumin)) & "', " '/pjumin
    SQL = SQL & vbCrLf & "  '" & Trim(GetText(vasID, asRow1, colPAge)) & "', " '/page
    SQL = SQL & vbCrLf & "  '" & Trim(GetText(vasID, asRow1, colPSex)) & "', " '/psex
    SQL = SQL & vbCrLf & "  '" & sEXAMDATE & "', " '/resdate
    SQL = SQL & vbCrLf & "  '" & Trim(GetText(vasRes, asRow2, colEQUIPEXAM)) & "', " '/EQUIPCODE
    SQL = SQL & vbCrLf & "  '" & Trim(GetText(vasRes, asRow2, colExamCode)) & "', " '/examcode
    SQL = SQL & vbCrLf & "  '', " '/examtype
    SQL = SQL & vbCrLf & "  '" & Trim(GetText(vasRes, asRow2, colResult)) & "', " '/result
    SQL = SQL & vbCrLf & "  '" & asSend & "', " '/sendflag
    SQL = SQL & vbCrLf & "  '" & Trim(GetText(vasRes, asRow2, colExamName)) & "', " '/examname
    SQL = SQL & vbCrLf & "  '" & Trim(GetText(vasRes, asRow2, colRCheck)) & "', " '/refflag
    SQL = SQL & vbCrLf & "  '" & Trim(GetText(vasRes, asRow2, colPCheck)) & "', " '/panicflag
    SQL = SQL & vbCrLf & "  '" & Trim(GetText(vasRes, asRow2, colDCheck)) & "', " '/deltaflag
    SQL = SQL & vbCrLf & "  '" & Trim(GetText(vasRes, asRow2, colUnit)) & "', "     '/unit
    SQL = SQL & vbCrLf & "  '" & Trim(GetText(vasRes, asRow2, colRef)) & "', "      '/refvalue
    SQL = SQL & vbCrLf & "  '" & Trim(GetText(vasRes, asRow2, colPanic)) & "', "    '/panicvalue
    SQL = SQL & vbCrLf & "  '" & Trim(GetText(vasID, asRow1, colSampleNo)) & "', "      '/SEQNO
    SQL = SQL & vbCrLf & "  '" & Trim(GetText(vasID, asRow1, colRack)) & "', "       '/diskno
    SQL = SQL & vbCrLf & "  '" & Trim(GetText(vasID, asRow1, colPos)) & "') "       '/posno
    
'    SaveQuery SQL
    res = SendQuery(gLocal, SQL)
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
End Function


'Private Sub vasID_KeyDown(KeyCode As Integer, Shift As Integer)
Private Sub vasID_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer
Dim j As Integer
Dim iStrNum As Integer

Dim iRow As Integer
Dim lRow As Long

    iRow = vasID.ActiveRow
    lRow = iRow
    
    If KeyCode = vbKeyReturn Then
'        If Trim(GetText(vasID, iRow, colSEQNO)) <> "" Then
'                iStrNum = Trim(GetText(vasID, iRow, colSEQNO))
'            For j = iRow To vasID.DataRowCnt
'                If j = iRow Then
'                Else
'                    iStrNum = iStrNum + 1
'                    SetText vasID, iStrNum, j, colSEQNO
'                End If
'            Next j
'        ElseIf Trim(GetText(vasID, iRow, colBARCODE)) <> "" Then
            SetText vasID, Trim(GetText(vasID, lRow, colBARCODE)), lRow, colBARCODE
            Get_Sample_Info lRow

            '2004/03/10 ÀÌ»óÀº
            For i = 1 To vasRes.DataRowCnt
                Save_Local_One lRow, i, "1"
            Next
            SQL = "delete from PAT_RES WHERE BARCODE = '' and SEQNO = '" & Trim(GetText(vasID, lRow, colRack)) & "'"
            res = SendQuery(gLocal, SQL)
'        End If
    End If
End Sub
'End Sub

Private Sub vasID_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    If Row < 1 Or Row > vasID.DataRowCnt Then
        Exit Sub
    End If
    
    PopupMenu mnuPop
End Sub

Private Sub vasRes_Click(ByVal Col As Long, ByVal Row As Long)
   vasRes.Row = vasRes.ActiveRow
   vasRes.Col = vasRes.ActiveCol
   ConfirmData = vasRes.Value
    
End Sub

Private Sub vasRes_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim Response, Help
    Dim vasResRow As Long
    Dim vasResCol As Long
    Dim vasIDRow As Long
        
    vasResRow = vasRes.ActiveRow
    vasResCol = vasRes.ActiveCol
    If KeyCode = vbKeyReturn Then
        vasIDRow = vasID.ActiveRow
        If vasResCol = colResult And _
           Trim(GetText(vasRes, vasResRow, colResult)) <> Trim(GetText(vasRes, vasResRow, colResult1)) Then
            
            Response = MsgBox("ÀúÀåÇÏ½Ã°Ú½À´Ï±î?", vbYesNo + vbCritical + vbDefaultButton2, "ÁÖÀÇ!!!  È®ÀÎ!!!", Help, 100)
            If Response = vbYes Then
                'ÆÇÁ¤, µ¨Å¸, ÆÐ´Ð ¼öÁ¤
'                Check_Result Trim(GetText(vasID, vasIDRow, colBARCODE)), _
'                             Trim(GetText(vasID, vasIDRow, colPID)), _
'                             Trim(GetText(vasRes, vasResRow, colExamCode)), _
'                             Trim(GetText(vasRes, vasResRow, colResult)), _
'                             vasResRow, Trim(GetText(vasID, vasIDRow, colPSex))

                SQL = " UPDATE PAT_RES " & vbCrLf & _
                      " Set result = '" & Trim(GetText(vasRes, vasResRow, colResult)) & "', " & vbCrLf & _
                      " refFlag = '" & Trim(GetText(vasRes, vasResRow, colRCheck)) & "', " & vbCrLf & _
                      " panicFlag = '" & Trim(GetText(vasRes, vasResRow, colPCheck)) & "', " & vbCrLf & _
                      " deltaFlag = '" & Trim(GetText(vasRes, vasResRow, colDCheck)) & "' " & vbCrLf & _
                      " WHERE EXAMDATE = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
                      "  AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
                      "  AND EQUIPCODE = '" & Trim(GetText(vasRes, vasResRow, colEQUIPEXAM)) & "'" & vbCrLf & _
                      "  AND BARCODE = '" & Trim(GetText(vasID, vasIDRow, colBARCODE)) & "' "
                res = SendQuery(gLocal, SQL)
                
                SetText vasRes, Trim(GetText(vasRes, vasResRow, colResult)), vasResRow, colResult1
                
            End If
        End If
        
    End If
End Sub

'Public Function QC_Result(argBARCODE As String, argExamCode As String, _
'                            argResult As String, ByVal argRow As Integer) As Integer
'    Dim sDiffRet, sDiffRet1 As String
'    Dim PreResult   As String
'
'    Dim sResClassCode As String     '°á°úÁ¾·ù
'    Dim sLow        As String       'ÂüÁ¶Ä¡
'    Dim sHigh       As String
'    Dim RefRet      As String
'
'    Dim sPart       As String
'    Dim sEquip      As String
'    Dim sLevel      As String
'    Dim sLotNo      As String
'
'    Dim sTmpRece1, sTmpRet1 As String
'    Dim sTmpRece2, sTmpRet2 As String
'    Dim i           As Integer
'    Dim sReceNo     As String
'    Dim sPID        As String
'
'    Dim sTmpStr As String
'
'    QC_Result = -1
'
'    If argBARCODE = "" Then
'        Exit Function
'    End If
'
'    If argExamCode = "" Then
'        Exit Function
'    End If
'
'
'    RefRet = ""
'
'    sDiffRet = argResult
'    If sDiffRet = "" Then
'        QC_Result = -1
'        Exit Function
'    End If
'    sPart = Trim(GetText(vasID, argRow, colJumin))
'    sEquip = gEquip
'    sLevel = Trim(GetText(vasID, argRow, colPName))
'    sLotNo = Trim(GetText(vasID, argRow, colPID))
'
'    SQL = "SELECT Max(q.AppDate), e.ResClassCode, e.Point, q.LimitLow, q.LimitHigh   " & vbCrLf & _
'          "From QCInItem q, ExamMaster e " & vbCrLf & _
'          "WHERE q.LabCode = '" & sPart & "' " & vbCrLf & _
'          "  and q.EQUIPCODE = '" & sEquip & "' " & vbCrLf & _
'          "  and q.QCInLevel = '" & sLevel & "' " & vbCrLf & _
'          "  and q.LotNo = '" & sLotNo & "' " & vbCrLf & _
'          "  and q.QCBARCODE = '" & argBARCODE & "' " & vbCrLf & _
'          "  and q.ExamCode = '" & argExamCode & "' " & vbCrLf & _
'          "  and q.AppDate >= '1900-01-01' " & vbCrLf & _
'          "  and e.AppDate = (SELECT Max(c.AppDate) from ExamMaster c WHERE c.AppDate >= '1900-01-01' and c.ExamCode = q.ExamCode)" & vbCrLf & _
'          "  and e.ExamCode = q.ExamCode " & vbCrLf & _
'          "GROUP BY e.ResClassCode, e.Point, q.LimitLow, q.LimitHigh"
'    res = db_SELECT_Col(gServer, SQL)
'    sResClassCode = Trim(gReadBuf(1))
'
'    If sResClassCode = "1" Then '¼ýÀÚ
''ÂüÁ¶Ä¡ Ã¼Å©
'        sLow = ""
'        sHigh = ""
'
'        '¼ýÀÚÀÎÁö ¾Æ´ÑÁö È®ÀÎ
'        If IsNumeric(sDiffRet) = False Then
'           MsgBox "°á°úÇü½ÄÀÌ ÀÏÄ¡ÇÏÁö ¾Ê½À´Ï´Ù.", vbInformation, "¾Ë¸²"
'           QC_Result = -1
'           Exit Function
'        End If
'
'        If IsNumeric(gReadBuf(2)) Then
'            If CInt(gReadBuf(2)) > 0 Then
'                sTmpStr = "#0."
'                For i = 1 To CInt(gReadBuf(2))
'                    sTmpStr = sTmpStr & "0"
'                Next i
'            Else
'                sTmpStr = "#0"
'            End If
'            sDiffRet = Format(sDiffRet, sTmpStr)
'            SetText vasRes, sDiffRet, argRow, colResult
'            SetText vasRes, sDiffRet, argRow, colResult1
'        End If
'
'        sLow = Trim(gReadBuf(3))
'        sHigh = Trim(gReadBuf(4))
'
'        If sLow = "" And sHigh = "" Then
'            RefRet = ""
'        ElseIf sLow = "" And sHigh <> "" Then   'ÀÌ»ó
'            If CCur(sHigh) < CCur(sDiffRet) Then
'                RefRet = "H"
'            End If
'        ElseIf sLow <> "" And sHigh = "" Then   'ÀÌÇÏ
'            If CCur(sLow) > CCur(sDiffRet) Then
'                RefRet = "L"
'            End If
'        Else
'            If CCur(sLow) > CCur(sDiffRet) Then
'                RefRet = "L"
'            ElseIf CCur(sHigh) < CCur(sDiffRet) Then
'                RefRet = "H"
'            ElseIf CCur(sLow) <= CCur(sDiffRet) And CCur(sHigh) <= CCur(sDiffRet) Then
'                RefRet = ""
'            End If
'        End If
'
'
'
'    ElseIf sResClassCode = "2" Then '¹®ÀÚ
''        Dim sRefValue As String
''        Dim sPanicValue As String
''        Dim sResult As String
''
''        sLow = ""
''        sLow = UCase(Trim(GetText(argTable, argRow, iresRefValue)))
''
''        '2003/03/17 ÀÌ»óÀº ¼öÁ¤
''        '°Ë»ç Ç×¸ñ °á°ú ÂüÁ¶ ÄÚµå Ã¼Å©¿¡¼­ 1 ÀÌ»óÀÏ °æ¿ì¸¸ ÆÇÁ¤µÇ°Ô
''        If Trim(GetText(argTable, argRow, iresDecision)) <> "" Then
''            Exit Function
''        End If
''
''        '2002³â 3¿ù 12ÀÏ +-¿¡¼­ +/-·Î ¼öÁ¤
''        '2002³â 5¿ù 13ÀÏ NON-REACTIVE ÆÇÁ¤ ¾ÈµÅ¼­ Ãß°¡
''        '2003³â 2¿ù 4ÀÏ ÀÌ»óÀº ¼öÁ¤ - 0-1·Î ÂüÁ¶Ä¡´Â 1ÀÌ³ª ÆÇÁ¤µÊ
''        '=================================================================================
''        '2002³â 5¿ù 13ÀÏ 1 : 40 ¹Ì¸¸ ÆÇÁ¤ ¾ÈµÊ
''        '2002³â 6¿ù 11ÀÏ (°á°úÂüÁ¶°¡ 1:·Î ½ÃÀÛÇÏ¸é ÆÇÁ¤Ã¼Å© ¾ÈÇÏ°Ô ¼öÁ¤)
''        If Trim(Left(sDiffRet, 3)) = "1 :" Or Trim(Left(sDiffRet, 3)) = "1:" Then
''            Exit Function
''        End If
''        '=================================================================================
''
''        SELECT Case UCase(sDiffRet)
''        Case "-", "NEGATIVE", "À½¼º", "1", "NON-REACTIVE", "0-1"
''            sResult = 1
''        Case "+/-", "2", "+-", "2-5"
''            sResult = 2
''        Case "+", "POSITIVE", "¾ç¼º", "3", "6-10"
''            sResult = 3
''        Case "++", "4", "11-20"
''            sResult = 4
''        Case "+++", "5", "21-30"
''            sResult = 5
''        Case "++++", "6"
''            sResult = 6
''        Case "+++++", "7"
''            sResult = 7
''        Case "++++++", "8"
''            sResult = 8
''        Case Else
''            sResult = sDiffRet
''        End SELECT
''        'sLow = "0-2"
''        If Trim(sLow) <> "" Then
''            SELECT Case UCase(Trim(sLow))
''            Case "-", "NEGATIVE", "À½¼º", "1", "NON-REACTIVE", "0-2"
''                sRefValue = 1
''            Case "+/-", "2", "+-"
''                sRefValue = 2
''            Case "+", "POSITIVE", "¾ç¼º", "3"
''                sRefValue = 3
''            Case "++", "4"
''                sRefValue = 4
''            Case "+++", "5"
''                sRefValue = 5
''            Case "++++", "6"
''                sRefValue = 6
''            Case "+++++", "7"
''                sRefValue = 7
''            Case "++++++", "8"
''                sRefValue = 8
''            Case Else
''                If Trim(GetText(argTable, argRow, iresDecision)) <> "" Then
''                    RefRet = Trim(GetText(argTable, argRow, iresDecision))
''                ElseIf UCase(sDiffRet) <> UCase(sLow) Then
''                    RefRet = sDiffRet
''                End If
''            End SELECT
''            If Trim(GetText(argTable, argRow, iresDecision)) <> "" Then
''
''            ElseIf sRefValue < sResult Then
'''                RefRet = "H"
''                RefRet = sDiffRet
''
'''                argTable.Row = argRow
'''                argTable.Col = iresDecision
'''                argTable.ForeColor = RGB(205, 55, 0)
''
''
''            End If
''        End If
''        If Trim(GetText(argTable, argRow, iresDecision)) <> "" Then
''            RefRet = Trim(GetText(argTable, argRow, iresDecision))
''        End If
'    End If
'
'    SetText vasRes, RefRet, argRow, colRCheck
'
'    If RefRet = "L" Then
'        vasRes.Row = argRow
'        vasRes.Col = colRCheck
'        vasRes.ForeColor = RGB(65, 105, 225)
'    Else
'        vasRes.Row = argRow
'        vasRes.Col = colRCheck
'        vasRes.ForeColor = RGB(205, 55, 0)
'    End If
'
'    QC_Result = 1
'
'End Function

Public Function Check_Result(argBARCODE As String, argPID As String, argExamCode As String, _
                            argResult As String, ByVal argRow As Integer, asSex As String) As Integer
    Dim sDiffRet, sDiffRet1 As String
    Dim PreResult   As String
    
    Dim sResClassCode As String     '°á°úÁ¾·ù
    Dim sLow        As String       'ÂüÁ¶Ä¡
    Dim sHigh       As String
    Dim RefRet      As String
    Dim sPanicGubun As String
    Dim sPanicLow   As String       'Panic
    Dim sPanicHigh  As String
    Dim PanicRet    As String
    Dim sDeltaGubun As String
    Dim sDeltaLow   As String       'Delta
    Dim sDeltaHigh  As String
    Dim DeltaRet    As String
    
    Dim sTmpRece1, sTmpRet1 As String
    Dim sTmpRece2, sTmpRet2 As String
    Dim sMax_ReceNo As String
    Dim i           As Integer
    Dim sReceNo     As String
    Dim sPID        As String
    
    Dim sTmpStr As String
    
    Check_Result = -1
    
    If argBARCODE = "" Then
        Exit Function
    End If
    
    If argExamCode = "" Then
        Exit Function
    End If
    

    RefRet = ""
    PanicRet = ""
    DeltaRet = ""
    
    sDiffRet = argResult
    If sDiffRet = "" Then
        Check_Result = -1
        Exit Function
    End If
    
    SQL = " SELECT ResClassCode, Res_M_Low, Res_M_High, Res_F_Low, Res_F_High, " & CR & _
          "        PanicValueGubun, Panic_M_Low, Panic_M_High, Panic_F_Low, Panic_F_High, " & CR & _
          "        DeltaValueGubun, DeltaLow, DeltaHigh, Point " & CR & _
          "From ExamMaster " & CR & _
          " WHERE HID = '117' " & CR & _
          " And ExamCode = '" & Trim(argExamCode) & "' "
    res = db_SELECT_Col(gServer, SQL)
    
    sResClassCode = Trim(gReadBuf(0))
    
    If sResClassCode = "1" Then '¼ýÀÚ
'ÂüÁ¶Ä¡ Ã¼Å©
        sLow = ""
        sHigh = ""
        
        '¼ýÀÚÀÎÁö ¾Æ´ÑÁö È®ÀÎ
        If IsNumeric(sDiffRet) = False Then
           'MsgBox "°á°úÇü½ÄÀÌ ÀÏÄ¡ÇÏÁö ¾Ê½À´Ï´Ù.", vbInformation, "¾Ë¸²"
           Check_Result = -1
           Exit Function
        End If
        
        If IsNumeric(gReadBuf(13)) Then
            If CInt(gReadBuf(13)) > 0 Then
                sTmpStr = "#0."
                For i = 1 To CInt(gReadBuf(13))
                    sTmpStr = sTmpStr & "0"
                Next i
            Else
                sTmpStr = "#0"
            End If
            sDiffRet = Format(sDiffRet, sTmpStr)
            SetText vasRes, sDiffRet, argRow, colResult
            SetText vasRes, sDiffRet, argRow, colResult1
        End If
        
        Select Case asSex
        Case "M", ""
            sLow = Trim(gReadBuf(1))
            sHigh = Trim(gReadBuf(2))
        Case "F"
            sLow = Trim(gReadBuf(3))
            sHigh = Trim(gReadBuf(4))
        End Select
        
        If sLow = "" And sHigh = "" Then
            RefRet = ""
        ElseIf sLow = "" And sHigh <> "" Then   'ÀÌ»ó
            If CCur(sHigh) < CCur(sDiffRet) Then
                RefRet = "H"
            End If
        ElseIf sLow <> "" And sHigh = "" Then   'ÀÌÇÏ
            If CCur(sLow) > CCur(sDiffRet) Then
                RefRet = "L"
            End If
        Else
            If CCur(sLow) > CCur(sDiffRet) Then
                RefRet = "L"
            ElseIf CCur(sHigh) < CCur(sDiffRet) Then
                RefRet = "H"
            ElseIf CCur(sLow) <= CCur(sDiffRet) And CCur(sHigh) <= CCur(sDiffRet) Then
                RefRet = ""
            End If
        End If


'Panic Ã¼Å©
        sPanicLow = ""
        sPanicHigh = ""
        
        sPanicGubun = Trim(gReadBuf(5))
        
        Select Case asSex
        Case "M", ""
            sPanicLow = Trim(gReadBuf(6))
            sPanicHigh = Trim(gReadBuf(7))
        Case "F"
            sPanicLow = Trim(gReadBuf(8))
            sPanicHigh = Trim(gReadBuf(9))
        End Select
        
        If sPanicGubun = "0" Then '»óÇÑ/ÇÏÇÑ
            If sPanicLow = "" Or sPanicHigh = "" Then
                PanicRet = ""
            Else
                If CCur(sPanicLow) > CCur(sDiffRet) Then
                    PanicRet = "L"
                ElseIf CCur(sPanicHigh) < CCur(sDiffRet) Then
                    PanicRet = "H"
                ElseIf CCur(sPanicLow) <= CCur(sDiffRet) And CCur(sPanicHigh) <= CCur(sDiffRet) Then
                    PanicRet = ""
                End If
            End If
        ElseIf sPanicGubun = "1" Then 'percent
            If sPanicLow = "" Then
                PanicRet = ""
            Else
                If CCur(sPanicLow) - CCur(sDiffRet) > 0 Then
                    If ((CCur(sPanicLow) - CCur(sDiffRet)) / CCur(sDiffRet)) * 100 >= CCur(sPanicHigh) Then
                        PanicRet = "L"
                    Else
                        PanicRet = ""
                    End If
                ElseIf CCur(sPanicHigh) - CCur(sDiffRet) < 0 Then
                    If ((CCur(sDiffRet) - CCur(sPanicLow)) / CCur(sDiffRet)) * 100 >= CCur(sPanicHigh) Then
                        PanicRet = "H"
                    Else
                        PanicRet = ""
                    End If
                Else
                    PanicRet = ""
                End If
            End If
        End If
        

'Delta Ã¼Å©
        sDeltaLow = ""
        sDeltaHigh = ""
                
        sTmpRece1 = ""
        sTmpRet1 = ""
        sTmpRece2 = ""
        sTmpRet2 = ""
        PreResult = ""
        
        sMax_ReceNo = ""
'        sTmpRece1 = Trim(argForm.dtpReceDate.Value)
        sReceNo = argBARCODE
        
'        SQL = "SELECT Result,Max(ReceNo) From ExamRes " & CR & _
'              " WHERE HID = '117' " & CR & _
'              " And PID = '" & Trim(argPID) & "' " & CR & _
'              " And ReceNo < '" & argBARCODE & "' " & CR & _
'              " And ExamCode = '" & Trim(argExamCode) & "' " & CR & _
'              " GROUP BY Result"
              
'2004/12/30 ÀÌ»óÀº - Á¤·ÄºÎºÐ Ãß°¡
        SQL = "SELECT Result,Max(ReceNo) From ExamRes " & CR & _
              " WHERE HID = '117' " & CR & _
              " And PID = '" & Trim(argPID) & "' " & CR & _
              " And ReceNo < '" & argBARCODE & "' " & CR & _
              " And ExamCode = '" & Trim(argExamCode) & "' " & CR & _
              " GROUP BY Result" & CR & _
              " ORDER BY 2 desc "
        res = db_SELECT_Col(gServer, SQL)
              
        If res > 0 And gReadBuf(0) <> "" Then
            PreResult = gReadBuf(0)
        Else
            PreResult = ""
        End If
        
        'ÀÌÀü°á°ú°¡ °ø¹éÀÌ ¾Æ´Ï°í, ¼ýÀÚÀÎ °æ¿ì¸¸
        If PreResult <> "" And IsNumeric(PreResult) Then
          'PreResult = Trim(gReadBuf(0))
          sDeltaGubun = Trim(gReadBuf(10))
          
          sDeltaLow = Trim(gReadBuf(11))
          sDeltaHigh = Trim(gReadBuf(12))
          
            'ÀÌÀü°á°ú¿¡¼­ ÇöÀç°á°ú »«°ªÀÌ sDiffRetÀÓ (2002³â 3¿ù 15ÀÏ ¼öÁ¤)
'            sDiffRet = PreResult - sDiffRet
            sDiffRet1 = sDiffRet - PreResult
            If sDeltaGubun = "0" Then '»óÇÑ/ÇÏÇÑ
                If sDeltaLow = "" Or sDeltaHigh = "" Then
                    DeltaRet = ""
                Else
                    If CCur(sDeltaLow) > CCur(sDiffRet1) Then
                        DeltaRet = "L"
                    ElseIf CCur(sDeltaHigh) < CCur(sDiffRet1) Then
                        DeltaRet = "H"
                    ElseIf CCur(sDeltaLow) <= CCur(sDiffRet1) And CCur(sDeltaHigh) <= CCur(sDiffRet1) Then
                        DeltaRet = ""
                    End If
                End If
              
            ElseIf sDeltaGubun = "1" Then 'percent
               If CInt(PreResult) = 0 Or CInt(sDiffRet) = 0 Then
                  DeltaRet = ""
               Else
                   If sDeltaLow = "" Then
                        DeltaRet = ""
                    Else
                        If (Abs(CCur(PreResult) - CCur(sDiffRet)) / CCur(PreResult)) * 100 >= CCur(sDeltaLow) Then
                            DeltaRet = "D"
                        Else
                            DeltaRet = ""
                        End If
                    End If
               End If
            End If
        End If
        
    ElseIf sResClassCode = "2" Then '¹®ÀÚ
'        Dim sRefValue As String
'        Dim sPanicValue As String
'        Dim sResult As String
'
'        sLow = ""
'        sLow = UCase(Trim(GetText(argTable, argRow, iresRefValue)))
'
'        '2003/03/17 ÀÌ»óÀº ¼öÁ¤
'        '°Ë»ç Ç×¸ñ °á°ú ÂüÁ¶ ÄÚµå Ã¼Å©¿¡¼­ 1 ÀÌ»óÀÏ °æ¿ì¸¸ ÆÇÁ¤µÇ°Ô
'        If Trim(GetText(argTable, argRow, iresDecision)) <> "" Then
'            Exit Function
'        End If
'
'        '2002³â 3¿ù 12ÀÏ +-¿¡¼­ +/-·Î ¼öÁ¤
'        '2002³â 5¿ù 13ÀÏ NON-REACTIVE ÆÇÁ¤ ¾ÈµÅ¼­ Ãß°¡
'        '2003³â 2¿ù 4ÀÏ ÀÌ»óÀº ¼öÁ¤ - 0-1·Î ÂüÁ¶Ä¡´Â 1ÀÌ³ª ÆÇÁ¤µÊ
'        '=================================================================================
'        '2002³â 5¿ù 13ÀÏ 1 : 40 ¹Ì¸¸ ÆÇÁ¤ ¾ÈµÊ
'        '2002³â 6¿ù 11ÀÏ (°á°úÂüÁ¶°¡ 1:·Î ½ÃÀÛÇÏ¸é ÆÇÁ¤Ã¼Å© ¾ÈÇÏ°Ô ¼öÁ¤)
'        If Trim(Left(sDiffRet, 3)) = "1 :" Or Trim(Left(sDiffRet, 3)) = "1:" Then
'            Exit Function
'        End If
'        '=================================================================================
'
'        SELECT Case UCase(sDiffRet)
'        Case "-", "NEGATIVE", "À½¼º", "1", "NON-REACTIVE", "0-1"
'            sResult = 1
'        Case "+/-", "2", "+-", "2-5"
'            sResult = 2
'        Case "+", "POSITIVE", "¾ç¼º", "3", "6-10"
'            sResult = 3
'        Case "++", "4", "11-20"
'            sResult = 4
'        Case "+++", "5", "21-30"
'            sResult = 5
'        Case "++++", "6"
'            sResult = 6
'        Case "+++++", "7"
'            sResult = 7
'        Case "++++++", "8"
'            sResult = 8
'        Case Else
'            sResult = sDiffRet
'        End SELECT
'        'sLow = "0-2"
'        If Trim(sLow) <> "" Then
'            SELECT Case UCase(Trim(sLow))
'            Case "-", "NEGATIVE", "À½¼º", "1", "NON-REACTIVE", "0-2"
'                sRefValue = 1
'            Case "+/-", "2", "+-"
'                sRefValue = 2
'            Case "+", "POSITIVE", "¾ç¼º", "3"
'                sRefValue = 3
'            Case "++", "4"
'                sRefValue = 4
'            Case "+++", "5"
'                sRefValue = 5
'            Case "++++", "6"
'                sRefValue = 6
'            Case "+++++", "7"
'                sRefValue = 7
'            Case "++++++", "8"
'                sRefValue = 8
'            Case Else
'                If Trim(GetText(argTable, argRow, iresDecision)) <> "" Then
'                    RefRet = Trim(GetText(argTable, argRow, iresDecision))
'                ElseIf UCase(sDiffRet) <> UCase(sLow) Then
'                    RefRet = sDiffRet
'                End If
'            End SELECT
'            If Trim(GetText(argTable, argRow, iresDecision)) <> "" Then
'
'            ElseIf sRefValue < sResult Then
''                RefRet = "H"
'                RefRet = sDiffRet
'
''                argTable.Row = argRow
''                argTable.Col = iresDecision
''                argTable.ForeColor = RGB(205, 55, 0)
'
'
'            End If
'        End If
'        If Trim(GetText(argTable, argRow, iresDecision)) <> "" Then
'            RefRet = Trim(GetText(argTable, argRow, iresDecision))
'        End If
'        sLow = ""
'        sLow = Trim(GetText(argTable, argRow, iresPanicValue))
'        If Trim(sLow) <> "" Then
'            SELECT Case UCase(Trim(sLow))
'            Case "-", "NEGATIVE", "À½¼º"
'                sPanicValue = 1
'            Case "+/-"
'                sPanicValue = 2
'            Case "+", "POSITIVE", "¾ç¼º"
'                sPanicValue = 3
'            Case "++"
'                sPanicValue = 4
'            Case "+++"
'                sPanicValue = 5
'            Case "++++"
'                sPanicValue = 6
'            Case "+++++"
'                sPanicValue = 7
'            Case "++++++"
'                sPanicValue = 8
'            Case Else
'                If UCase(sDiffRet) > UCase(sLow) Then
'                    PanicRet = sDiffRet
'                End If
'            End SELECT
'            If sPanicValue < sResult Then
'                'PanicRet = "H"
'                PanicRet = sDiffRet
'            End If
'        End If
'
'        'Delta Check
'        sMax_ReceNo = ""
'        DeltaRet = ""
'        sReceNo = Trim(GetText(argForm.vasPatient, 1, 1))
'        sPID = Trim(GetText(argForm.vasPatient, 1, 3))
'
'        SQL = "SELECT Result,Max(ReceNo) From ExamRes " & CR & _
'              " WHERE PID = '" & sPID & "' " & CR & _
'              " And ReceNo < '" & sReceNo & "' " & CR & _
'              " And ExamCode = '" & Trim(GetText(argTable, argRow, iresExamCode)) & "' " & CR & _
'              " GROUP BY Result"
'
'        res = db_SELECT_Col(SQL)
'
'        If res > 0 And gReadBuf(0) <> "" Then
'               If sDiffRet <> gReadBuf(0) Then
'                  DeltaRet = "D"
'               End If
'        Else
'            DeltaRet = ""
'        End If
    End If
    
    SetText vasRes, RefRet, argRow, colRCheck
    SetText vasRes, PanicRet, argRow, colPCheck
    SetText vasRes, DeltaRet, argRow, colDCheck
    

    '2002³â 2¿ù 15ÀÏ ¼öÁ¤ (ÆÇÁ¤½Ã H, L ÀÏ¶§ ±ÛÀÚ »ö±ò º¯È­)
    '2002³â 3¿ù 14ÀÏ ¼öÁ¤ (ÆÇÁ¤½Ã LÀÏ¶§´Â ÆÄ¶õ»ö ±× ¿Ü´Â »¡°£»ö)
    If RefRet = "L" Then
        vasRes.Row = argRow
        vasRes.Col = colRCheck
        vasRes.ForeColor = RGB(65, 105, 225)
    Else
        vasRes.Row = argRow
        vasRes.Col = colRCheck
        vasRes.ForeColor = RGB(205, 55, 0)
    End If
    
    If PanicRet = "L" Then
        vasRes.Row = argRow
        vasRes.Col = colPCheck
        vasRes.ForeColor = RGB(65, 105, 225)
    Else
        vasRes.Row = argRow
        vasRes.Col = colPCheck
        vasRes.ForeColor = RGB(205, 55, 0)
    End If
    
    If DeltaRet = "L" Then
        vasRes.Row = argRow
        vasRes.Col = colDCheck
        vasRes.ForeColor = RGB(65, 105, 225)
    ElseIf DeltaRet = "D" Then
        vasRes.Row = argRow
        vasRes.Col = colDCheck
        vasRes.ForeColor = RGB(65, 105, 225)
    Else
        vasRes.Row = argRow
        vasRes.Col = colDCheck
        vasRes.ForeColor = RGB(205, 55, 0)
    End If
    
    Check_Result = 1

End Function

