VERSION 5.00
Object = "{8996B0A4-D7BE-101B-8650-00AA003A5593}#4.0#0"; "Cfx4032.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm302QCReview_N_ALL 
   BackColor       =   &H00DBE6E6&
   ClientHeight    =   9240
   ClientLeft      =   90
   ClientTop       =   600
   ClientWidth     =   14640
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   9240
   ScaleWidth      =   14640
   Begin MSComDlg.CommonDialog DlgSave 
      Left            =   7065
      Top             =   4365
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00E0E0E0&
      Caption         =   "È­¸éÁö¿ò(&C)"
      Height          =   510
      Left            =   11700
      Style           =   1  '±×·¡ÇÈ
      TabIndex        =   68
      Tag             =   "124"
      ToolTipText     =   "È­¸éÀ» ÃÊ±â»óÅÂ·Î Áö¿ò´Ï´Ù."
      Top             =   150
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Á¾     ·á(&X)"
      Height          =   510
      Left            =   13080
      Style           =   1  '±×·¡ÇÈ
      TabIndex        =   67
      Tag             =   "128"
      ToolTipText     =   "ÀÌ È­¸éÀ» Á¾·áÇÕ´Ï´Ù."
      Top             =   150
      Width           =   1320
   End
   Begin VB.Frame fraQcParameter 
      BackColor       =   &H00DBE6E6&
      Height          =   1410
      Left            =   75
      TabIndex        =   0
      Top             =   -45
      Width           =   14385
      Begin VB.ComboBox cboWorkarea 
         Height          =   300
         Left            =   1470
         Sorted          =   -1  'True
         Style           =   2  'µå·Ó´Ù¿î ¸ñ·Ï
         TabIndex        =   75
         Top             =   600
         Width           =   2805
      End
      Begin VB.CommandButton cmdControl 
         BackColor       =   &H00F4F0F2&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7545
         Picture         =   "frm302QCReview_N_ALL.frx":0000
         Style           =   1  '±×·¡ÇÈ
         TabIndex        =   70
         Top             =   180
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.TextBox txtCtrlCd 
         Appearance      =   0  'Æò¸é
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6420
         TabIndex        =   69
         Top             =   180
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.ComboBox cboSection 
         Height          =   300
         Left            =   1470
         Style           =   2  'µå·Ó´Ù¿î ¸ñ·Ï
         TabIndex        =   62
         Top             =   600
         Visible         =   0   'False
         Width           =   2805
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Excel ÀúÀå(&E)"
         Height          =   510
         Left            =   13000
         Style           =   1  '±×·¡ÇÈ
         TabIndex        =   15
         ToolTipText     =   "Á¶È¸µÈ ÄÁÆ®·ÑÀÇ ¸ðµç °Ë»çÇ×¸ñ¿¡ ´ëÇÑ °á°ú¸®½ºÆ®¸¦ Ãâ·ÂÇÕ´Ï´Ù."
         Top             =   825
         Width           =   1320
      End
      Begin VB.CommandButton cmdQuery 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Á¶      È¸(&Q)"
         Height          =   510
         Left            =   11640
         Style           =   1  '±×·¡ÇÈ
         TabIndex        =   14
         Tag             =   "133"
         ToolTipText     =   "¼³Á¤µÈ °Ë»öÁ¶°Ç¿¡ ÀÇÇØ µ¥ÀÌÅÍ¸¦ Á¶È¸ÇÕ´Ï´Ù."
         Top             =   825
         Width           =   1320
      End
      Begin MSComCtl2.DTPicker dtpFdate 
         Height          =   375
         Left            =   1485
         TabIndex        =   1
         Top             =   165
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   85393411
         CurrentDate     =   36850
      End
      Begin MSComCtl2.DTPicker dtpTdate 
         Height          =   375
         Left            =   2970
         TabIndex        =   7
         Top             =   165
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   85393411
         CurrentDate     =   36850
      End
      Begin MedControls1.LisLabel lblSectiona 
         Height          =   300
         Left            =   1485
         TabIndex        =   12
         Top             =   600
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   529
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Appearance      =   0
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00DBE6E6&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   6435
         TabIndex        =   2
         Top             =   870
         Visible         =   0   'False
         Width           =   4575
         Begin VB.OptionButton optLevel 
            BackColor       =   &H00DBE6E6&
            Caption         =   "Low"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   3
            Left            =   3135
            TabIndex        =   6
            Tag             =   "'L'"
            Top             =   165
            Width           =   645
         End
         Begin VB.OptionButton optLevel 
            BackColor       =   &H00DBE6E6&
            Caption         =   "High"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   2
            Left            =   2145
            TabIndex        =   5
            Tag             =   "'H'"
            Top             =   165
            Width           =   690
         End
         Begin VB.OptionButton optLevel 
            BackColor       =   &H00DBE6E6&
            Caption         =   "Normal"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   930
            TabIndex        =   4
            Tag             =   "'N'"
            Top             =   165
            Width           =   915
         End
         Begin VB.OptionButton optLevel 
            BackColor       =   &H00DBE6E6&
            Caption         =   "All"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   135
            TabIndex        =   3
            Tag             =   "'N','H','L'"
            Top             =   165
            Value           =   -1  'True
            Width           =   495
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00DBE6E6&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1485
         TabIndex        =   8
         Top             =   1260
         Visible         =   0   'False
         Width           =   2805
         Begin VB.OptionButton optAR 
            BackColor       =   &H00DBE6E6&
            Caption         =   "Accept"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   220
            Index           =   1
            Left            =   795
            TabIndex        =   11
            Tag             =   "'A'"
            Top             =   165
            Width           =   840
         End
         Begin VB.OptionButton optAR 
            BackColor       =   &H00DBE6E6&
            Caption         =   "Reject"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   220
            Index           =   2
            Left            =   1815
            TabIndex        =   10
            Tag             =   "'R'"
            Top             =   165
            Width           =   825
         End
         Begin VB.OptionButton optAR 
            BackColor       =   &H00DBE6E6&
            Caption         =   "All"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   220
            Index           =   0
            Left            =   120
            TabIndex        =   9
            Tag             =   "'A','R'"
            Top             =   165
            Value           =   -1  'True
            Width           =   495
         End
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   4
         Left            =   5055
         TabIndex        =   66
         Top             =   975
         Visible         =   0   'False
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "Control Level"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   11
         Left            =   105
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   165
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "°Ë»ç±â°£"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   1
         Left            =   105
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   1365
         Visible         =   0   'False
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "ÀûÇÕ¿©ºÎ"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   3
         Left            =   105
         TabIndex        =   65
         Top             =   570
         Visible         =   0   'False
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "Section"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblCtrlNm 
         Height          =   360
         Left            =   7875
         TabIndex        =   71
         Top             =   180
         Visible         =   0   'False
         Width           =   3105
         _ExtentX        =   5477
         _ExtentY        =   635
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblEqp 
         Height          =   360
         Left            =   6420
         TabIndex        =   72
         Top             =   585
         Visible         =   0   'False
         Width           =   4560
         _ExtentX        =   8043
         _ExtentY        =   635
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   2
         Left            =   5040
         TabIndex        =   73
         Top             =   585
         Visible         =   0   'False
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "°Ë»çÀåºñ"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   0
         Left            =   5040
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   180
         Visible         =   0   'False
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "Control Info."
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   5
         Left            =   105
         TabIndex        =   76
         Top             =   570
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "Workarea"
         Appearance      =   0
      End
      Begin VB.Line Line1 
         X1              =   2820
         X2              =   2930
         Y1              =   330
         Y2              =   330
      End
   End
   Begin Crystal.CrystalReport crtRpt 
      Left            =   6450
      Top             =   4365
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MedControls1.LisLabel lblMsg 
      Height          =   495
      Left            =   5160
      TabIndex        =   18
      Top             =   5040
      Visible         =   0   'False
      Width           =   7800
      _ExtentX        =   13758
      _ExtentY        =   873
      BackColor       =   16252927
      ForeColor       =   14641726
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "QC µ¥ÀÌÅ¸¸¦ °Ë»öÁßÀÔ´Ï´Ù. Àá½Ã¸¸ ±â´Ù¸®¼¼¿ä...."
      Appearance      =   0
      LeftGab         =   0
   End
   Begin VB.Frame fraResult 
      BackColor       =   &H00DBE6E6&
      Height          =   7080
      Left            =   90
      TabIndex        =   16
      Top             =   1260
      Width           =   14400
      Begin FPSpread.vaSpread tblResult 
         Height          =   6840
         Left            =   90
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   210
         Width           =   14250
         _Version        =   196608
         _ExtentX        =   25135
         _ExtentY        =   12065
         _StockProps     =   64
         AllowCellOverflow=   -1  'True
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "µ¸¿ò"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16515071
         MaxCols         =   20
         MaxRows         =   24
         OperationMode   =   1
         Protect         =   0   'False
         ShadowColor     =   15461355
         ShadowDark      =   14737632
         ShadowText      =   4210752
         SpreadDesigner  =   "frm302QCReview_N_ALL.frx":00B2
         UserResize      =   1
         VisibleCols     =   14
         TextTip         =   2
      End
   End
   Begin VB.Frame fraGrpAll 
      BackColor       =   &H00DBE6E6&
      Height          =   7080
      Left            =   75
      TabIndex        =   58
      Top             =   1275
      Visible         =   0   'False
      Width           =   14385
      Begin VB.CommandButton cmdQuitAll 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   14070
         TabIndex        =   60
         Top             =   105
         Width           =   285
      End
      Begin VB.CommandButton cmdGrpPrtAll 
         Height          =   255
         Left            =   13785
         Picture         =   "frm302QCReview_N_ALL.frx":0AF1
         Style           =   1  '±×·¡ÇÈ
         TabIndex        =   59
         ToolTipText     =   "±×·¡ÇÁ¸¦ Ãâ·ÂÇÕ´Ï´Ù."
         Top             =   105
         Width           =   285
      End
      Begin ChartfxLibCtl.ChartFX cfxRst 
         Height          =   6555
         Index           =   3
         Left            =   135
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   375
         Width           =   13980
         _cx             =   2037211219
         _cy             =   2037198122
         Build           =   7
         TypeMask        =   -718749695
         Style           =   -106979329
         RightGap        =   10
         TopGap          =   10
         BottomGap       =   25
         WallWidth       =   8
         View3DDepth     =   60
         TypeEx          =   0
         StyleEx         =   1
         CylSides        =   2
         MarkerShape     =   2
         MarkerSize      =   2
         Volume          =   10
         BorderColor     =   16777263
         Axis(0).MinorStep=   -0.2
         Axis(0).Min     =   4
         Axis(0).Max     =   8
         Axis(0).Style   =   10280
         Axis(0).GridColor=   0
         Axis(1).Min     =   0
         Axis(1).Max     =   100
         Axis(1).Decimals=   0
         Axis(1).Style   =   10344
         Axis(1).GridColor=   0
         Axis(2).Step    =   1
         Axis(2).MinorStep=   1
         Axis(2).Style   =   14368
         Axis(2).PixPerUnit=   19
         Axis(2).GridColor=   0
         RGBBk           =   14737632
         RGB2DBk         =   15461355
         RGB3DBk         =   16777215
         nColors         =   5
         Pallete         =   "frm302QCReview_N_ALL.frx":0E33
         Colors          =   "frm302QCReview_N_ALL.frx":0F17
         LeftFontMask    =   805306368
         BeginProperty LeftFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²Ã¼"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RightFontMask   =   1879048192
         BeginProperty RightFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²Ã¼"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TopFontMask     =   268435456
         BeginProperty TopFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²Ã¼"
            Size            =   6.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BottomFontMask  =   268435456
         BeginProperty BottomFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²Ã¼"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Axis(2).FontMask=   268435456
         BeginProperty Axis(2).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²Ã¼"
            Size            =   6.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Axis(0).FontMask=   268435456
         BeginProperty Axis(0).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²Ã¼"
            Size            =   6.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FixedFontMask   =   268435456
         BeginProperty FixedFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²Ã¼"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LegendFontMask  =   268435456
         BeginProperty LegendFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²Ã¼"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PointLabelsFontMask=   268435456
         BeginProperty PointLabelsFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²Ã¼"
            Size            =   6.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PointFontMask   =   268435456
         BeginProperty PointFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²Ã¼"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         nPts            =   20
         nSer            =   5
         NumPoint        =   20
         NumSer          =   5
         Stripes         =   "frm302QCReview_N_ALL.frx":0F5F
         Fixed           =   "frm302QCReview_N_ALL.frx":0F9F
         Multi           =   "frm302QCReview_N_ALL.frx":0FDF
         MMask           =   4
         BorderS         =   4
         Enabled         =   -1
         _Data_          =   "frm302QCReview_N_ALL.frx":10B3
      End
   End
   Begin VB.ListBox lstControl 
      BeginProperty Font 
         Name            =   "±¼¸²Ã¼"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      Left            =   6510
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   495
      Visible         =   0   'False
      Width           =   4515
   End
   Begin VB.Frame fraGraph 
      BackColor       =   &H00DBE6E6&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   7080
      Left            =   75
      TabIndex        =   19
      Tag             =   "30218"
      Top             =   1365
      Width           =   14385
      Begin VB.CommandButton cmdQuit 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   14055
         TabIndex        =   56
         ToolTipText     =   "±×·¡ÇÁ È­¸éÀ» ´Ý½À´Ï´Ù."
         Top             =   150
         Width           =   285
      End
      Begin VB.CommandButton cmdGrpPrt 
         Height          =   255
         Left            =   13770
         Picture         =   "frm302QCReview_N_ALL.frx":1395
         Style           =   1  '±×·¡ÇÈ
         TabIndex        =   55
         ToolTipText     =   "±×·¡ÇÁ¸¦ Ãâ·ÂÇÕ´Ï´Ù."
         Top             =   150
         Width           =   285
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Æò¸é
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   75
         ScaleHeight     =   225
         ScaleWidth      =   13875
         TabIndex        =   43
         Top             =   390
         Width           =   13905
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Åõ¸í
            Caption         =   "Lot No. : "
            BeginProperty Font 
               Name            =   "±¼¸²"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Left            =   240
            TabIndex        =   49
            Top             =   30
            Width           =   960
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Åõ¸í
            Caption         =   "Open Date : "
            BeginProperty Font 
               Name            =   "±¼¸²"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Left            =   3525
            TabIndex        =   48
            Top             =   15
            Width           =   1260
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Åõ¸í
            Caption         =   "Expire Date : "
            BeginProperty Font 
               Name            =   "±¼¸²"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Left            =   6480
            TabIndex        =   47
            Top             =   15
            Width           =   1380
         End
         Begin VB.Label lblLotNo 
            BackStyle       =   0  'Åõ¸í
            Caption         =   "30030"
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Index           =   0
            Left            =   1200
            TabIndex        =   46
            Top             =   30
            Width           =   975
         End
         Begin VB.Label lblOpenDt 
            BackStyle       =   0  'Åõ¸í
            Caption         =   "30030"
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Index           =   0
            Left            =   4800
            TabIndex        =   45
            Top             =   30
            Width           =   975
         End
         Begin VB.Label lblExpDt 
            BackStyle       =   0  'Åõ¸í
            Caption         =   "30030"
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Index           =   0
            Left            =   7845
            TabIndex        =   44
            Top             =   30
            Width           =   975
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Æò¸é
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   75
         ScaleHeight     =   225
         ScaleWidth      =   13875
         TabIndex        =   36
         Top             =   2595
         Width           =   13905
         Begin VB.Label lblExpDt 
            BackStyle       =   0  'Åõ¸í
            Caption         =   "30030"
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Index           =   1
            Left            =   7845
            TabIndex        =   42
            Top             =   30
            Width           =   1005
         End
         Begin VB.Label lblOpenDt 
            BackStyle       =   0  'Åõ¸í
            Caption         =   "30030"
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Index           =   1
            Left            =   4815
            TabIndex        =   41
            Top             =   30
            Width           =   1005
         End
         Begin VB.Label lblLotNo 
            BackStyle       =   0  'Åõ¸í
            Caption         =   "30030"
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Index           =   1
            Left            =   1200
            TabIndex        =   40
            Top             =   30
            Width           =   1005
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Åõ¸í
            Caption         =   "Expire Date : "
            BeginProperty Font 
               Name            =   "±¼¸²"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Left            =   6480
            TabIndex        =   39
            Top             =   15
            Width           =   1380
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Åõ¸í
            Caption         =   "Open Date : "
            BeginProperty Font 
               Name            =   "±¼¸²"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Left            =   3525
            TabIndex        =   38
            Top             =   15
            Width           =   1260
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Åõ¸í
            Caption         =   "Lot No. : "
            BeginProperty Font 
               Name            =   "±¼¸²"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Left            =   240
            TabIndex        =   37
            Top             =   30
            Width           =   960
         End
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Æò¸é
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   75
         ScaleHeight     =   225
         ScaleWidth      =   13875
         TabIndex        =   29
         Top             =   4815
         Width           =   13905
         Begin VB.Label lblExpDt 
            BackStyle       =   0  'Åõ¸í
            Caption         =   "30030"
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Index           =   2
            Left            =   7875
            TabIndex        =   35
            Top             =   30
            Width           =   1050
         End
         Begin VB.Label lblOpenDt 
            BackStyle       =   0  'Åõ¸í
            Caption         =   "30030"
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Index           =   2
            Left            =   4800
            TabIndex        =   34
            Top             =   30
            Width           =   1050
         End
         Begin VB.Label lblLotNo 
            BackStyle       =   0  'Åõ¸í
            Caption         =   "30030"
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Index           =   2
            Left            =   1200
            TabIndex        =   33
            Top             =   45
            Width           =   1050
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Åõ¸í
            Caption         =   "Expire Date : "
            BeginProperty Font 
               Name            =   "±¼¸²"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Left            =   6480
            TabIndex        =   32
            Top             =   15
            Width           =   1380
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Åõ¸í
            Caption         =   "Open Date : "
            BeginProperty Font 
               Name            =   "±¼¸²"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Left            =   3525
            TabIndex        =   31
            Top             =   15
            Width           =   1260
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Åõ¸í
            Caption         =   "Lot No. : "
            BeginProperty Font 
               Name            =   "±¼¸²"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Left            =   240
            TabIndex        =   30
            Top             =   30
            Width           =   960
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00E0E0E0&
         Height          =   1590
         Left            =   13605
         ScaleHeight     =   1530
         ScaleWidth      =   330
         TabIndex        =   26
         Top             =   1005
         Width           =   390
         Begin VB.CheckBox chkPrint 
            BackColor       =   &H00E0E0E0&
            Height          =   330
            Index           =   0
            Left            =   60
            TabIndex        =   27
            Top             =   810
            Width           =   270
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Åõ¸í
            Caption         =   "Ãâ·Â"
            Height          =   375
            Left            =   60
            TabIndex        =   28
            Top             =   1155
            Width           =   225
         End
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H00E0E0E0&
         Height          =   1590
         Left            =   13605
         ScaleHeight     =   1530
         ScaleWidth      =   330
         TabIndex        =   23
         Top             =   3210
         Width           =   390
         Begin VB.CheckBox chkPrint 
            BackColor       =   &H00E0E0E0&
            Height          =   330
            Index           =   1
            Left            =   60
            TabIndex        =   24
            Top             =   810
            Width           =   270
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Åõ¸í
            Caption         =   "Ãâ·Â"
            Height          =   375
            Left            =   60
            TabIndex        =   25
            Top             =   1155
            Width           =   225
         End
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H00E0E0E0&
         Height          =   1605
         Left            =   13605
         ScaleHeight     =   1545
         ScaleWidth      =   330
         TabIndex        =   20
         Top             =   5400
         Width           =   390
         Begin VB.CheckBox chkPrint 
            BackColor       =   &H00E0E0E0&
            Height          =   330
            Index           =   2
            Left            =   60
            TabIndex        =   21
            Top             =   810
            Width           =   270
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Åõ¸í
            Caption         =   "Ãâ·Â"
            Height          =   375
            Left            =   60
            TabIndex        =   22
            Top             =   1155
            Width           =   225
         End
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   360
         Index           =   2
         Left            =   13605
         TabIndex        =   50
         Top             =   5055
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   635
         _Version        =   393216
         BuddyControl    =   "cfxRst(2)"
         BuddyDispid     =   196662
         BuddyIndex      =   2
         OrigLeft        =   13605
         OrigTop         =   5475
         OrigRight       =   13980
         OrigBottom      =   5835
         Orientation     =   1
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   345
         Index           =   0
         Left            =   13605
         TabIndex        =   51
         Top             =   660
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   609
         _Version        =   393216
         BuddyControl    =   "cfxRst(0)"
         BuddyDispid     =   196662
         BuddyIndex      =   0
         OrigLeft        =   13605
         OrigTop         =   795
         OrigRight       =   13980
         OrigBottom      =   1155
         Orientation     =   1
         Enabled         =   -1  'True
      End
      Begin ChartfxLibCtl.ChartFX cfxRst 
         Height          =   1935
         Index           =   1
         Left            =   75
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   2850
         Width           =   13530
         _cx             =   2037210425
         _cy             =   2037189973
         Build           =   7
         TypeMask        =   -1794588671
         Style           =   -106979329
         LeftGap         =   71
         RightGap        =   20
         TopGap          =   9
         BottomGap       =   24
         WallWidth       =   8
         View3DDepth     =   60
         TypeEx          =   0
         StyleEx         =   1
         CylSides        =   2
         MarkerShape     =   2
         MarkerSize      =   2
         Volume          =   10
         BorderColor     =   16777263
         Axis(0).MinorStep=   -2
         Axis(0).Min     =   -1
         Axis(0).Max     =   11
         Axis(0).Style   =   10280
         Axis(0).GridColor=   0
         Axis(1).Min     =   0
         Axis(1).Max     =   100
         Axis(1).Decimals=   0
         Axis(1).Style   =   10344
         Axis(1).GridColor=   0
         Axis(2).Step    =   1
         Axis(2).MinorStep=   1
         Axis(2).Style   =   14368
         Axis(2).PixPerUnit=   19
         Axis(2).GridColor=   0
         RGBBk           =   15396338
         RGB2DBk         =   15988216
         RGB3DBk         =   14737632
         nColors         =   1
         Pallete         =   "frm302QCReview_N_ALL.frx":16D7
         Colors          =   "frm302QCReview_N_ALL.frx":17BB
         LeftFontMask    =   805306368
         BeginProperty LeftFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RightFontMask   =   1879048192
         BeginProperty RightFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TopFontMask     =   268435456
         BeginProperty TopFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BottomFontMask  =   268435456
         BeginProperty BottomFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Axis(2).FontMask=   268435456
         BeginProperty Axis(2).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Axis(0).FontMask=   268435456
         BeginProperty Axis(0).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FixedFontMask   =   268435456
         BeginProperty FixedFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LegendFontMask  =   268435456
         BeginProperty LegendFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PointLabelsFontMask=   268435456
         BeginProperty PointLabelsFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PointFontMask   =   268435456
         BeginProperty PointFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         nPts            =   20
         nSer            =   1
         NumPoint        =   20
         NumSer          =   1
         Stripes         =   "frm302QCReview_N_ALL.frx":17E3
         Fixed           =   "frm302QCReview_N_ALL.frx":1823
         BorderS         =   4
         _Data_          =   "frm302QCReview_N_ALL.frx":1863
      End
      Begin ChartfxLibCtl.ChartFX cfxRst 
         Height          =   1890
         Index           =   2
         Left            =   75
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   5085
         Width           =   13545
         _cx             =   2037210452
         _cy             =   2037189894
         Build           =   7
         TypeMask        =   -1794588671
         Style           =   -106979329
         LeftGap         =   71
         RightGap        =   20
         TopGap          =   9
         BottomGap       =   19
         WallWidth       =   8
         View3DDepth     =   60
         TypeEx          =   0
         StyleEx         =   1
         CylSides        =   2
         MarkerShape     =   2
         MarkerSize      =   2
         Volume          =   10
         BorderColor     =   16777263
         Axis(0).MinorStep=   -2
         Axis(0).Min     =   -1
         Axis(0).Max     =   11
         Axis(0).Style   =   10280
         Axis(0).GridColor=   0
         Axis(1).Min     =   0
         Axis(1).Max     =   100
         Axis(1).Decimals=   0
         Axis(1).Style   =   10344
         Axis(1).GridColor=   0
         Axis(2).Step    =   1
         Axis(2).MinorStep=   1
         Axis(2).Style   =   14368
         Axis(2).PixPerUnit=   19
         Axis(2).GridColor=   0
         RGBBk           =   15133416
         RGB2DBk         =   15133416
         RGB3DBk         =   14737632
         nColors         =   1
         Pallete         =   "frm302QCReview_N_ALL.frx":18D9
         Colors          =   "frm302QCReview_N_ALL.frx":19BD
         LeftFontMask    =   805306368
         BeginProperty LeftFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RightFontMask   =   1879048192
         BeginProperty RightFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TopFontMask     =   268435456
         BeginProperty TopFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BottomFontMask  =   268435456
         BeginProperty BottomFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Axis(2).FontMask=   268435456
         BeginProperty Axis(2).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Axis(0).FontMask=   268435456
         BeginProperty Axis(0).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FixedFontMask   =   268435456
         BeginProperty FixedFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LegendFontMask  =   268435456
         BeginProperty LegendFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PointLabelsFontMask=   268435456
         BeginProperty PointLabelsFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PointFontMask   =   268435456
         BeginProperty PointFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         nPts            =   20
         nSer            =   1
         NumPoint        =   20
         NumSer          =   1
         Stripes         =   "frm302QCReview_N_ALL.frx":19E5
         Fixed           =   "frm302QCReview_N_ALL.frx":1A25
         BorderS         =   4
         _Data_          =   "frm302QCReview_N_ALL.frx":1A65
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   360
         Index           =   1
         Left            =   13605
         TabIndex        =   54
         Top             =   2850
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   635
         _Version        =   393216
         BuddyControl    =   "cfxRst(1)"
         BuddyDispid     =   196662
         BuddyIndex      =   1
         OrigLeft        =   13605
         OrigTop         =   3135
         OrigRight       =   13980
         OrigBottom      =   3495
         Orientation     =   1
         Enabled         =   -1  'True
      End
      Begin ChartfxLibCtl.ChartFX cfxRst 
         Height          =   1935
         Index           =   0
         Left            =   75
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   645
         Width           =   13545
         _cx             =   2037210452
         _cy             =   2037189973
         Build           =   7
         TypeMask        =   -1794588671
         Style           =   -106979329
         LeftGap         =   71
         RightGap        =   20
         TopGap          =   9
         BottomGap       =   24
         WallWidth       =   8
         View3DDepth     =   60
         TypeEx          =   0
         StyleEx         =   1
         CylSides        =   2
         MarkerShape     =   2
         MarkerSize      =   2
         Volume          =   10
         BorderColor     =   16777263
         Axis(0).MinorStep=   -2
         Axis(0).Min     =   -1
         Axis(0).Max     =   11
         Axis(0).Style   =   10280
         Axis(0).GridColor=   0
         Axis(1).Min     =   0
         Axis(1).Max     =   100
         Axis(1).Decimals=   0
         Axis(1).Style   =   10344
         Axis(1).GridColor=   0
         Axis(2).Step    =   1
         Axis(2).MinorStep=   1
         Axis(2).Style   =   14368
         Axis(2).PixPerUnit=   19
         Axis(2).GridColor=   0
         RGBBk           =   14737632
         RGB2DBk         =   15461355
         RGB3DBk         =   16777215
         nColors         =   1
         Pallete         =   "frm302QCReview_N_ALL.frx":1ADB
         Colors          =   "frm302QCReview_N_ALL.frx":1BBF
         LeftFontMask    =   805306368
         BeginProperty LeftFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RightFontMask   =   1879048192
         BeginProperty RightFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TopFontMask     =   268435456
         BeginProperty TopFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BottomFontMask  =   268435456
         BeginProperty BottomFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Axis(2).FontMask=   268435456
         BeginProperty Axis(2).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Axis(0).FontMask=   268435456
         BeginProperty Axis(0).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FixedFontMask   =   268435456
         BeginProperty FixedFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LegendFontMask  =   268435456
         BeginProperty LegendFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PointLabelsFontMask=   268435456
         BeginProperty PointLabelsFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PointFontMask   =   268435456
         BeginProperty PointFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         nPts            =   20
         nSer            =   1
         NumPoint        =   20
         NumSer          =   1
         Stripes         =   "frm302QCReview_N_ALL.frx":1BE7
         Fixed           =   "frm302QCReview_N_ALL.frx":1C27
         BorderS         =   4
         _Data_          =   "frm302QCReview_N_ALL.frx":1C67
      End
   End
End
Attribute VB_Name = "frm302QCReview_N_ALL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Binary
'Coding By Legends
'QC °á°úÁ¶È¸

Private objCtrl As New clsQCTtest
Private objQcReview As clsQcReview

'Private WithEvents mnuPopup As Menu
'Private WithEvents mnuGraph As Menu
'Private WithEvents mnuGraphAll As Menu
'Private WithEvents mnuPrint As Menu
Private WithEvents objPop As clsPopupMenu
Attribute objPop.VB_VarHelpID = -1
Private Const MENU_ONE& = 1
Private Const MENU_ALL& = 2
Private Const MENU_PRT& = 3

Private ActControl As String

Private Type CalRef
    calVALUE() As Double
    calMEAN As Double
    calCV As Double
    calSD As Double
    calMIN As Double
    calMAX As Double
    calTOTAL As Double
    calCOUNT As Long
End Type

Private CalRefVal(0 To 2) As CalRef

Private Const Fld_Div = "" 'chr(15)
Private Const Rec_Div = "" 'chr(14)
Private Const CFX_NULL = 1E+308

Public Event LastFormUnload()

Private Sub cmdClear_Click()
'    txtCtrlCd.Text = ""
'    lblCtrlNm.Caption = ""
'    txtMeanL.Text = ""
'    txtMeanN.Text = ""
'    txtMeanH.Text = ""
'    txtSDL.Text = ""
'    txtSDN.Text = ""
'    txtSDH.Text = ""
'    txtCVL.Text = ""
'    txtCVN.Text = ""
'    txtCVH.Text = ""

    Call InitForm
    Call InitGraph
    
    Call LoadSection
End Sub

Private Sub cmdControl_Click()
'    ActControl = cmdControl.Name
'    Call GetControl
'    On Error Resume Next
'    If lstControl.Visible Then lstControl.SetFocus
    If lblCtrlNm.Caption <> "" Then
        DoEvents
        txtCtrlCd.Text = ""
        lblCtrlNm.Caption = ""
        
        Call InitForm
        Call InitGraph
    End If
    DoEvents
    Call LoadControlInfo
    DoEvents
    Call LoadTestItem
    DoEvents
    Call LoadData
End Sub

Private Sub LoadControlInfo(Optional ByVal pCtrlCd As String = "")
'ÄÁÆ®·ÑÀÇ ÀÏ¹Ý Á¤º¸¸¦ ºÒ·¯¿Â´Ù..
    Dim objPop As clsPopUpList
    Dim i As Long
    
    Set objPop = New clsPopUpList

    With objPop
        .Recordset = GetControlInfo(pCtrlCd)
        .FormCaption = "ÄÁÆ®·Ñ Ã£±â"
        .Delimiter = COL_DIV
        .FormWidth = 6470
        .ColumnHeaderText = "ÄÚµåÄÁÆ®·Ñ¸íLevelÀåºñÄÚµåÀåºñ¸í"
        .ColumnHeaderWidth = "854.92922475.213002475.213"
        '0 ¿ÞÂÊ, 1 ¿À¸¥ÂÊ, 2 °¡¿îµ¥
        
        Call .LoadPopUp
        
        DoEvents
        
        txtCtrlCd.Text = medGetP(.SelectedString, 1, .Delimiter)
        lblCtrlNm.Caption = medGetP(.SelectedString, 2, .Delimiter)
        
        lblEqp.Caption = Format(medGetP(.SelectedString, 4, .Delimiter), "!" & String(10, "@")) & medGetP(.SelectedString, 5, .Delimiter)
        lblEqp.ToolTipText = Format(medGetP(.SelectedString, 4, .Delimiter), "!" & String(10, "@")) & medGetP(.SelectedString, 5, .Delimiter)
    End With
    
    Set objPop = Nothing
End Sub

Private Function GetControlInfo(Optional ByVal pCtrlCd As String = "") As Recordset
    Dim strSql As String
    
    strSql = " select a.ctrlcd,a.ctrlnm,a.levelcd,a.eqpcd,b.eqpnm" & _
             " from " & T_LAB021 & " a, " & T_LAB006 & " b " & _
             " where " & DBJ("a.eqpcd*= b.eqpcd") & " and ( inusefg =  '1' or a.eqpcd is null)"
             
    If pCtrlCd <> "" Then
        strSql = strSql & " and " & DBW("a.ctrlcd=", pCtrlCd)
    End If
    
    If cboSection.ListIndex > 0 Then '¼½¼ÇÀ» ¼±ÅÃÇßÀ» °æ¿ì
        strSql = strSql & "and " & DBW("a.sectcd=", Trim(medGetP(cboSection.Text, 2, COL_DIV)))
    End If
    
    Set GetControlInfo = New Recordset
    GetControlInfo.Open strSql, DBConn
End Function

Private Sub GetControl()
    Dim Rs As Recordset
    Dim strSql As String
    Dim strList As String
    
    Select Case ActControl
        Case cmdControl.Name
            strSql = " select a.ctrlcd,a.ctrlnm,a.eqpcd,b.eqpnm,a.sectcd,c.field1 as sectnm " & _
                     " from " & T_LAB021 & " a, " & T_LAB006 & " b, " & T_LAB032 & " c " & _
                     " where " & DBJ("a.eqpcd *= b.eqpcd ") & _
                     " and a.sectcd=c.cdval1 " & _
                     " and " & DBW("c.cdindex=", LC3_Section)
                     
            If cboSection.ListIndex > 0 Then '¼½¼ÇÀ» ¼±ÅÃÇßÀ» °æ¿ì
                strSql = strSql & "and " & DBW("a.sectcd=", Trim(medGetP(cboSection.Text, 2, COL_DIV)))
            End If

        Case Else
            strSql = " select a.ctrlcd,a.ctrlnm,a.eqpcd,b.eqpnm,a.sectcd,c.field1 as sectnm " & _
                     " from " & T_LAB021 & " a, " & T_LAB006 & " b, " & T_LAB032 & " c " & _
                     " where " & DBJ("a.eqpcd *= b.eqpcd ") & _
                     " and a.sectcd=c.cdval1 " & _
                     " and " & DBW("c.cdindex=", LC3_Section) & _
                     " and " & DBW("ctrlcd=", Trim(txtCtrlCd.Text))
                     
            If cboSection.ListIndex > 0 Then '¼½¼ÇÀ» ¼±ÅÃÇßÀ» °æ¿ì
                strSql = strSql & "and " & DBW("a.sectcd=", Trim(medGetP(cboSection.Text, 2, COL_DIV)))
            End If
    End Select
    
    strSql = strSql & " order by ctrlcd,ctrlnm"
    
    Set Rs = New Recordset
    Rs.Open strSql, DBConn
    
    lstControl.Clear
    Do Until Rs.EOF
        strList = Format(Rs.Fields("ctrlcd").Value & "", "!" & String(12, "@")) & _
                  Format(Rs.Fields("ctrlnm").Value & "", "!" & String(35, "@")) & _
                  Format(Rs.Fields("eqpcd").Value & "", "!" & String(10, "@")) & _
                  Format(Rs.Fields("eqpnm").Value & "", "!" & String(32, "@")) & _
                  Format(Rs.Fields("sectcd").Value & "", "!" & String(4, "@")) & _
                  Format(Rs.Fields("sectnm").Value & "", "!" & String(50, "@"))
                
        lstControl.addItem strList
        Rs.MoveNext
    Loop
        
    If lstControl.ListCount > 0 Then
        lstControl.Visible = True
        lstControl.ZOrder 0
    Else
        txtCtrlCd.Text = ""
        lblCtrlNm.Caption = ""
        lblEqp.Caption = ""
'        lblSection.Caption = ""
        MsgBox "ÄÁÆ®·ÑÀÌ Á¸ÀçÇÏÁö ¾Ê½À´Ï´Ù.", vbExclamation
    End If
        
    Set Rs = Nothing
End Sub

Private Sub cmdExit_Click()
    Unload Me
'    Unload frmControls
    
'    Set frmControls = Nothing
    Set frm302QCReview_N = Nothing
    
    If IsLastForm Then RaiseEvent LastFormUnload
End Sub

Private Sub GraphPrint(ByVal lngGrp As Long)
''''    Dim vbTitle As String
''''    Dim vbBottom(5) As String
''''    Dim vbConLn As String
''''    Dim vbRefLn As String
''''    Dim vbCntLn As String
''''    Dim i As Long
''''    Dim lngIdx As Long
''''
''''    lngIdx = Val(cfxRst(lngGrp).Tag)
''''
''''    For i = optLevel.LBound To optLevel.UBound
''''        If optLevel(i).Value = True Then vbConLn = vbConLn & "Level(" & optLevel(i).Caption & ")   "
''''    Next i
''''
''''    For i = optAR.LBound To optAR.UBound
''''        If optAR(i).Value = True Then vbConLn = vbConLn & "ÆÇÁ¤(" & optAR(i).Caption & ")   "
''''    Next i
''''
''''    Call RefCalculate(lngGrp)
''''
''''    With objQcReview.Item(lngIdx)
''''        vbRefLn = "MEAN(" & .MeanVal & ")   SD(" & .SdVal & ")   CV(" & .CvVal & " )" '& Space(150)
''''        vbCntLn = "TOTAL(" & .TotCnt & ")   ACCEPT(" & .AccCnt & ")   REJECT(" & .RejCnt & ")" '& Space(150)
''''    End With
''''
''''    With CalRefVal(lngGrp)
''''        Dim strCV As String
''''
''''        strCV = .calCV * 100
''''
''''        vbRefLn = vbRefLn & " >> Calculated : MEAN(" & .calMEAN & ")   SD(" & .calSD & ")   CV(" & strCV & ")" '& Space(150)
''''
''''        vbTitle = "<< Quality Control : " & Trim(Mid(lstTestItem.Text, 1, 30)) & " >>" '& Space(150)
''''        vbBottom(1) = "* Á¶   °Ç : " & vbConLn '& Space(150)
''''        vbBottom(2) = "* ±â   °£ : " & Format(dtpFdate.Value, CS_DateLongFormat) & " ºÎÅÍ " & Format(dtpTdate.Value, CS_DateLongFormat) & " ±îÁö " '& Space(150)
''''        vbBottom(3) = "* Control : " & txtCtrlCd.Text & "      Lot Number : " & lblLotNo(lngGrp).Caption & _
''''                   "     Open Date : " & lblOpenDt(lngGrp).Caption & "     Expire Date : " & lblExpDt(lngGrp).Caption '& Space(150)
''''        vbBottom(4) = "* ÂüÁ¶Ä¡  : " & vbRefLn '& Space(150)
''''        vbBottom(5) = "* °Ç   ¼ö : " & vbCntLn '& Space(150)
''''    End With
''''
''''    cfxRst(lngGrp).Height = 4000
''''    cfxRst(lngGrp).FixedGap = 15
''''    cfxRst(lngGrp).BottomGap = 33
''''    cfxRst(lngGrp).LeftGap = 70
''''    cfxRst(lngGrp).RightGap = 50
''''
''''    cfxRst(lngGrp).Title(CHART_TOPTIT) = vbTitle & vbCrLf & vbCrLf & _
''''                                        vbBottom(1) & vbCrLf & _
''''                                        vbBottom(2) & vbCrLf & _
''''                                        vbBottom(3) & vbCrLf & _
''''                                        vbBottom(4) & vbCrLf & vbCrLf & _
''''                                        vbBottom(5) & vbCrLf & vbCrLf & vbCrLf
''''
''''    cfxRst(lngGrp).Printer.Orientation = ORIENTATION_LANDSCAPE
'''''Ãâ·Â¹°¿¡ À­ ¸¶Áø ³Ð°Ô º¸±â
''''    cfxRst(lngGrp).Printer.TopMargin = 3
''''    cfxRst(lngGrp).Printer.LeftMargin = 1.5
'''''Ãâ·Â¹° À­ ¸¶Áø Àû´çÈ÷ º¸±â(ÀÌ ¸®¸¶Å©´Â Áö¿ìÁö ¸»°Í)
'''''    cfxRst(lngGrp).Printer.TopMargin = 1 / 2.54
'''''    cfxRst(lngGrp).Printer.LeftMargin = 1 / 2.54
''''    cfxRst(lngGrp).Printer.RightMargin = 1 / 2.5
''''    cfxRst(lngGrp).Printer.BottomMargin = 1 / 2.54
''''    cfxRst(lngGrp).Printer.Compress = True
''''    cfxRst(lngGrp).PrintIt 0, 0
''''
''''    cfxRst(lngGrp).Height = 1995
''''    cfxRst(lngGrp).Title(CHART_TOPTIT) = ""
''''
''''    cfxRst(lngGrp).FixedGap = 5
''''    cfxRst(lngGrp).BottomGap = 33
''''    cfxRst(lngGrp).LeftGap = 60
''''    cfxRst(lngGrp).RightGap = 10
''''    cfxRst(lngGrp).TopGap = 5
''''
''''    cfxRst(lngGrp).Refresh
End Sub

Private Sub RefCalculate(ByVal lngIdx As Long)
    
    Dim dblTemp As Double
    Dim k As Long
    
    dblTemp = 0
    With CalRefVal(lngIdx)
        .calMEAN = .calTOTAL / .calCOUNT
        For k = 0 To .calCOUNT - 1
            dblTemp = dblTemp + ((.calVALUE(k) - .calMEAN) * (.calVALUE(k) - .calMEAN))
        Next k
        If .calCOUNT > 1 Then .calSD = Sqr(dblTemp / (.calCOUNT - 1))
        If .calMEAN <> 0 Then .calCV = (.calSD / .calMEAN)     ' * 100
        
        .calMEAN = CStr(ksR(.calMEAN, 3))
        .calCV = CStr(ksR(.calCV, 3))
        .calSD = CStr(ksR(.calSD, 3))
        .calMIN = CStr(ksR(.calMIN, 3))
        .calMAX = CStr(ksR(.calMAX, 3))
    End With

End Sub

Private Sub cmdGrpPrt_Click()
    Dim i As Long
    
    For i = chkPrint.LBound To chkPrint.UBound
        If chkPrint(i).Value = 1 Then
            Call GraphPrint(i)
        End If
    Next
End Sub

Private Sub cmdGrpPrtAll_Click()
    Dim strTmp As String
    Dim i As Long
    
    With cfxRst(3)
        strTmp = " << Á¤µµ°ü¸® ±×·¡ÇÁ >> " & vbNewLine & vbNewLine
        strTmp = strTmp & " * °Ë»ç   ±â°£ : " & Replace(Format(dtpFdate.Value, "yyyy/MM/dd") & "~" & Format(dtpTdate.Value, "yyyy/MM/dd"), "-", "/") & vbNewLine
        strTmp = strTmp & " * ÄÁÆ®·Ñ Á¤º¸ : " & txtCtrlCd.Text & String(5, " ") & lblCtrlNm.Caption & vbNewLine
        strTmp = strTmp & " * Section     : " & Trim(medGetP(cboSection.Text, 1, COL_DIV)) & vbNewLine
        strTmp = strTmp & " * °Ë»ç   Àåºñ : " & lblEqp.ToolTipText & vbNewLine
        
        For i = optAR.LBound To optAR.UBound
            If optAR(i).Value Then
                strTmp = strTmp & " * ÀûÇÕ   ¿©ºÎ : " & optAR(i).Caption & vbNewLine
                Exit For
            End If
        Next
            
        For i = optLevel.LBound To optLevel.UBound
            If optLevel(i).Value Then
                strTmp = strTmp & " * ÄÁÆ®·Ñ ·¹º§ : " & optLevel(i).Caption & vbNewLine
                Exit For
            End If
        Next
        
        .Title(CHART_TOPTIT) = strTmp
        
        .LineWidth = 4
        For i = 0 To .ConstantLine.Count - 1
            .ConstantLine(i).LineWidth = 2
        Next
        
        .Printer.Orientation = ORIENTATION_LANDSCAPE
        .Printer.BottomMargin = 1 / 2.54
        .Printer.TopMargin = 1 / 2.54
        .Printer.LeftMargin = 1 / 2.54
        .Printer.RightMargin = 1 / 2.54
        .Printer.Compress = True
        
        .PrintIt 0, 0
        
        .Title(CHART_TOPTIT) = ""
        .LineWidth = 1.5
        For i = 0 To .ConstantLine.Count - 1
            .ConstantLine(i).LineWidth = 1
        Next
        
        .LeftGap = 50
        .RightGap = 10
        .TopGap = 10
        .BottomGap = 25
        .Refresh
    End With
End Sub

Private Sub cmdPrint_Click()
''    Dim i As Long
''
''    If tblResult.DataRowCnt = 0 Then Exit Sub
''    cmdPrint.Tag = "cmdPrint"
''
''    Call PrintCrystalAllItem
    Dim OneRec As String '±×¸®µå ÇÑÇàÀÇ ³»¿ëÀ» °¡Áö°í ÀÖ´Â º¯¼ö
    Dim FileName As String
    Dim rCount As Integer
    Dim cCount As Integer
    Dim f As String
    Dim i As Double
    Dim j As Integer
    Dim varTmp
   
    'CancelÀ» True·Î ¼³Á¤ÇÑ´Ù.
    DlgSave.CancelError = True
On Error GoTo Err_Handler:
    OneRec = ""
    FileName = ""
    'Flags»ó¼ö : cdlOFNOverwritePrompt(ÀÌ¹Ì Á¸ÀçÇÏ´Â È­ÀÏÀ» ¼±ÅÃÇÑ °æ¿ì ¿¡·¯Ã³¸®)
    '           cdlOFNExplorer(Å½»ö±â¿Í °°Àº ÇüÅÂÀÇ ÆÄÀÏ¼±ÅÃ È­¸é (Win95, 32bit))
    '           cdlOFNLongNames(±ä ÆÄÀÏ ÀÌ¸§(Long File Name) Çã¿ë)
    DlgSave.Flags = cdlOFNOverwritePrompt Or cdlOFNExplorer Or cdlOFNLongNames
    DlgSave.Filter = "¿¢¼¿ÆÄÀÏ (*.xls) |*.xls|¸ðµçÆÄÀÏ (*.*)|*.*"
    DlgSave.DialogTitle = "¿¢¼¿È­ÀÏ ÇüÅÂ·Î¸¸ ÀúÀåµË´Ï´Ù"
    DlgSave.InitDir = App.Path
    DlgSave.FileName = FileName
    DlgSave.ShowSave
    
    If Len(DlgSave.FileName) = 0 Then Exit Sub
    
    f = FreeFile()
'''    rCount = lvwLabNum.ListItems.Count
'''    cCount = lvwLabNum.ListItems.Item(1).ListSubItems.Count
'''
    OneRec = ""
    Open DlgSave.FileName For Output As #f
        For i = 1 To tblResult.MaxRows
            'TITLE PRINT
            If i = 1 Then
               OneRec = "No" & vbTab & "Control Inf" & vbTab & "°Ë»çÀåºñ" & vbTab & "°Ë»çÇ×¸ñ" & vbTab & "ÀÏÀÚ/½Ã°£" & vbTab & "Level" & vbTab & "°á°ú" & vbTab & "AR" & vbTab & "Remark" & vbTab
               OneRec = OneRec & "ÀÏÀÚ/½Ã°£" & vbTab & "Level" & vbTab & "°á°ú" & vbTab & "AR" & vbTab & "Remark" & vbTab
               OneRec = OneRec & "ÀÏÀÚ/½Ã°£" & vbTab & "Level" & vbTab & "°á°ú" & vbTab & "AR" & vbTab & "Remark" & vbTab
               Print #f, OneRec
               OneRec = ""
            End If

            With tblResult
                .GetText 1, i, varTmp: OneRec = OneRec & varTmp & vbTab
                .GetText 2, i, varTmp: OneRec = OneRec & varTmp & vbTab
                .GetText 3, i, varTmp: OneRec = OneRec & varTmp & vbTab
                .GetText 4, i, varTmp: OneRec = OneRec & varTmp & vbTab
                .GetText 5, i, varTmp: OneRec = OneRec & varTmp & vbTab
                .GetText 6, i, varTmp: OneRec = OneRec & varTmp & vbTab
                .GetText 7, i, varTmp: OneRec = OneRec & varTmp & vbTab
                .GetText 8, i, varTmp: OneRec = OneRec & varTmp & vbTab
                .GetText 9, i, varTmp: OneRec = OneRec & varTmp & vbTab
                .GetText 10, i, varTmp: OneRec = OneRec & varTmp & vbTab
                .GetText 11, i, varTmp: OneRec = OneRec & varTmp & vbTab
                .GetText 12, i, varTmp: OneRec = OneRec & varTmp & vbTab
                .GetText 13, i, varTmp: OneRec = OneRec & varTmp & vbTab
                .GetText 14, i, varTmp: OneRec = OneRec & varTmp & vbTab
                .GetText 15, i, varTmp: OneRec = OneRec & varTmp & vbTab
                .GetText 16, i, varTmp: OneRec = OneRec & varTmp & vbTab
                .GetText 17, i, varTmp: OneRec = OneRec & varTmp & vbTab
                .GetText 18, i, varTmp: OneRec = OneRec & varTmp & vbTab
                .GetText 19, i, varTmp: OneRec = OneRec & varTmp & vbTab
            End With
'''            OneRec
'''            OneRec = lvwLabNum.ListItems.Item(i) & vbTab
'''            For j = 1 To cCount
'''                OneRec = OneRec & lvwLabNum.ListItems.Item(i).SubItems(j) & vbTab
'''            Next j
            Print #f, OneRec
            OneRec = ""
        Next i
    Close #f

    Exit Sub
    
Err_Handler:
    If Err.Number = cdlCancel Then
        'Ãë¼Ò´ÜÃß¸¦ ´­·¶½À´Ï´Ù.
    Else
        MsgBox Err.Number & ":" & Err.Description, vbQuestion
    End If
    Exit Sub
End Sub

Private Sub PrintCrystalAllItem()
    Dim Rs As Recordset
    Dim strSql As String
    Dim lngFNo As Long
    Dim strFNm As String
    Dim strRptNm As String
    Dim strTmp As String
    Dim i As Long
    Dim j As Long
    Dim varTmp As Variant
    Dim strKey As String
    Dim strAccNo As String
    
    If Dir(InstallDir & "lis\rpt\QCReviewAllItem.rpt") = "" Then
        MsgBox "Ãâ·ÂµµÁß ¿À·ù°¡ ¹ß»ýÇÏ¿´½À´Ï´Ù. QCReviewAllItem.rpt ÆÄÀÏÀÌ ¾ø½À´Ï´Ù.", vbExclamation
        Exit Sub
    Else
        strRptNm = InstallDir & "lis\rpt\QCReviewAllItem.rpt"
    End If
    
    If Dir(InstallDir & "lis\rpt\CrystalReport.txt") = "" Then
        MsgBox "Ãâ·ÂµµÁß ¿À·ù°¡ ¹ß»ýÇÏ¿´½À´Ï´Ù. CrystalReport.txt ÆÄÀÏÀÌ ¾ø½À´Ï´Ù.", vbExclamation
        Exit Sub
    Else
        strFNm = InstallDir & "lis\rpt\CrystalReport.txt"
    End If
   
    strSql = " select a.workarea,a.accdt,a.accseq,a.testcd,h.abbrnm10,a.rstcd,a.radiv,a.vfydt,a.vfytm,a.ctrlcd,a.levelcd,a.lotno,c.text" & _
             " from " & T_LAB026 & " a, " & T_LAB001 & " h, " & T_LAB028 & " c " & _
             " where " & DBW("a.ctrlcd=", txtCtrlCd.Text) & _
             " and   a.levelcd in (" & GetLevel & ") " & _
             " and  a.testcd in(" & GetTestCd & ") " & _
             " and  " & DBW("a.vfydt >=", Format(dtpFdate.Value, CS_DateDbFormat)) & _
             " and  " & DBW("a.vfydt <=", Format(dtpTdate.Value, CS_DateDbFormat)) & _
             " and   a.radiv in ( " & GetAR & ") " & _
             " and   h.testcd = a.testcd " & _
             " and   h.applydt = (select max(applydt) from s2lab001 where testcd = h.testcd) " & _
             " and " & DBJ("a.workarea*=c.workarea") & _
             " and " & DBJ("a.accdt*=c.accdt") & _
             " and " & DBJ("a.accseq*=c.accseq") & _
             " and " & DBJ("a.testcd*=c.testcd") & _
             " order by a.testcd,a.ctrlcd,a.levelcd,a.lotno,a.vfydt,a.vfytm"
    
    Set Rs = New Recordset
    Rs.Open strSql, DBConn
            
    strTmp = ""
    Do Until Rs.EOF
        strKey = ""
        
        strTmp = strTmp & Replace(Format(dtpFdate.Value, "yyyy/MM/dd") & "~" & Format(dtpTdate.Value, "yyyy/MM/dd"), "-", "/") & vbTab
        strTmp = strTmp & txtCtrlCd.Text & String(5, " ") & lblCtrlNm.Caption & vbTab
        strTmp = strTmp & Trim(medGetP(cboSection.Text, 1, COL_DIV)) & vbTab
        strTmp = strTmp & lblEqp.ToolTipText & vbTab
        For j = optAR.LBound To optAR.UBound
            If optAR(j).Value Then
                strTmp = strTmp & optAR(j).Caption & vbTab
                Exit For
            End If
        Next
        
        For j = optLevel.LBound To optLevel.UBound
            If optLevel(j).Value Then
                strTmp = strTmp & optLevel(j).Caption & vbTab
                Exit For
            End If
        Next
                
        strTmp = strTmp & Rs.Fields("abbrnm10").Value & "" & vbTab
        strTmp = strTmp & Rs.Fields("lotno").Value & "" & vbTab
        strTmp = strTmp & Rs.Fields("levelcd").Value & "" & vbTab
        strTmp = strTmp & Mid(Rs.Fields("vfydt").Value & "", 5) & "/" & Mid(Rs.Fields("vfytm").Value & "", 1, 4) & vbTab
        strTmp = strTmp & Rs.Fields("rstcd").Value & "" & vbTab
        strTmp = strTmp & Rs.Fields("radiv").Value & "" & vbTab
        
        If Rs.Fields("radiv").Value & "" = "R" Then
            strAccNo = Rs.Fields("workarea").Value & "" & Fld_Div & _
                       Rs.Fields("accdt").Value & "" & Fld_Div & _
                       Rs.Fields("accseq").Value & ""
            
            strKey = Rs.Fields("levelcd").Value & "" & Fld_Div & _
                     Rs.Fields("lotno").Value & "" & Fld_Div & _
                     Rs.Fields("testcd").Value
                     
            strTmp = strTmp & Replace(Replace(Rs.Fields("text").Value & "", vbTab, String(5, " ")) & GetRecentRstForPrt(strAccNo, strKey), vbNewLine, COL_DIV) & vbNewLine
        Else
            strTmp = strTmp & Replace(Replace(Rs.Fields("text").Value & "", vbTab, String(5, " ")), vbNewLine, COL_DIV) & vbNewLine
        End If
        Rs.MoveNext
    Loop
                
    Set Rs = Nothing
    
    strTmp = Mid(strTmp, 1, Len(strTmp) - 2)
    lngFNo = FreeFile
    
    Open strFNm For Output As #lngFNo
    Print #lngFNo, strTmp
    Close #lngFNo
    
    With crtRpt
        .ReportFileName = strRptNm
        .ParameterFields(0) = "HospNm;" & ObjSysInfo.Hospital & ";true"
        .RetrieveDataFiles
        .WindowState = 2
        .Destination = crptToWindow
        .Action = 1
        .Reset
    End With
End Sub

Private Function GetRecentRstForPrt(ByVal pAccNo As String, ByVal pKey As String) As String
    Dim Rs As Recordset
    Dim strSql As String
    Dim strWorkArea As String
    Dim strAccDt As String
    Dim strAccSeq As String

    Dim strPreWA As String  'ÀÌÀü°á°ú
    Dim strPreAD As String
    Dim strPreAS As String
    Dim strPreRst As String
    Dim strPreRA As String
    
    Dim strCurWA As String  'ÇöÀç°á°ú
    Dim strCurAD As String
    Dim strCurAS As String
    Dim strCurRst As String
    Dim strCurRA As String
    
    Dim strPostWA As String 'ÀÌÈÄ°á°ú
    Dim strPostAD As String
    Dim strPostAS As String
    Dim strPostRst As String
    Dim strPostRA As String
    
    Dim PreTmp As String
    Dim PostTmp As String
    
    strWorkArea = medGetP(pAccNo, 1, Fld_Div)
    strAccDt = medGetP(pAccNo, 2, Fld_Div)
    strAccSeq = medGetP(pAccNo, 3, Fld_Div)
    
    strSql = " select * from " & T_LAB026 & _
             " where " & DBW("ctrlcd=", txtCtrlCd.Text) & _
             " and " & DBW("levelcd=", medGetP(pKey, 1, Fld_Div)) & _
             " and " & DBW("lotno=", medGetP(pKey, 2, Fld_Div)) & _
             " and " & DBW("testcd=", medGetP(pKey, 3, Fld_Div)) & _
             " order by workarea,accdt,accseq "


    Set Rs = New Recordset
    Rs.Open strSql, DBConn
    
    Do Until Rs.EOF
        strCurWA = Rs.Fields("workarea").Value & ""
        strCurAD = Rs.Fields("accdt").Value & ""
        strCurAS = Rs.Fields("accseq").Value & ""
        strCurRst = Rs.Fields("rstcd").Value & ""
        strCurRA = Rs.Fields("radiv").Value & ""
        
        If (strCurAD = strAccDt) And (strCurAS = strAccSeq) Then
            Rs.MoveNext
            
            If Rs.EOF = False Then
                strPostWA = Rs.Fields("workarea").Value & ""
                strPostAD = Rs.Fields("accdt").Value & ""
                strPostAS = Rs.Fields("accseq").Value & ""
                strPostRst = Rs.Fields("rstcd").Value & ""
                strPostRA = Rs.Fields("radiv").Value & ""
            End If
            
            Exit Do
        End If
        
        strPreWA = Rs.Fields("workarea").Value & ""
        strPreAD = Rs.Fields("accdt").Value & ""
        strPreAS = Rs.Fields("accseq").Value & ""
        strPreRst = Rs.Fields("rstcd").Value & ""
        strPreRA = Rs.Fields("radiv").Value & ""
        
        Rs.MoveNext
    Loop


    If strPreWA <> "" Then
        PreTmp = "¢ÑÀÌÀü °á°ú(" & _
                     strPreRst & "," & _
                     strPreRA & ")"
    End If
'                     strPreWA & "-" & Mid(strPreAD, 3) & "-" & strPreAS & "," & _

    If strPostWA <> "" Then
        PostTmp = "¢ÑÀÌÈÄ °á°ú(" & _
                      strPostRst & "," & _
                      strPostRA & ")"
    End If
'                      strPostWA & "-" & Mid(strPostAD, 3) & "-" & strPostAS & "," & _

    GetRecentRstForPrt = PreTmp & PostTmp
    
    Set Rs = Nothing
End Function


Private Sub PrintCrystal()
    Dim lngFNo As Long
    Dim strTmp As String
    Dim strTmp2 As String
    Dim strRptNm As String
    Dim i As Long
    Dim j As Long
    Dim varTmp As Variant
    Dim strFNm As String
        
    If Dir(InstallDir & "lis\rpt\QCReview.rpt") = "" Then
        MsgBox "Ãâ·ÂµµÁß ¿À·ù°¡ ¹ß»ýÇÏ¿´½À´Ï´Ù. QCReview.rpt ÆÄÀÏÀÌ ¾ø½À´Ï´Ù.", vbExclamation
        Exit Sub
    Else
        strRptNm = InstallDir & "lis\rpt\QCReview.rpt"
    End If
    
    If Dir(InstallDir & "lis\rpt\CrystalReport.txt") = "" Then
        MsgBox "Ãâ·ÂµµÁß ¿À·ù°¡ ¹ß»ýÇÏ¿´½À´Ï´Ù. CrystalReport.txt ÆÄÀÏÀÌ ¾ø½À´Ï´Ù.", vbExclamation
        Exit Sub
    Else
        strFNm = InstallDir & "lis\rpt\CrystalReport.txt"
    End If
        
    strTmp = ""
    With tblResult
        For i = 1 To .DataRowCnt
            strTmp = strTmp & Replace(Format(dtpFdate.Value, "yyyy/MM/dd") & "~" & Format(dtpTdate.Value, "yyyy/MM/dd"), "-", "/") & vbTab
            strTmp = strTmp & txtCtrlCd.Text & String(5, " ") & lblCtrlNm.Caption & vbTab
            strTmp = strTmp & Trim(medGetP(cboSection.Text, 1, COL_DIV)) & vbTab
            strTmp = strTmp & lblEqp.ToolTipText & vbTab
            For j = optAR.LBound To optAR.UBound
                If optAR(j).Value Then
                    strTmp = strTmp & optAR(j).Caption & vbTab
                    Exit For
                End If
            Next
            
            For j = optLevel.LBound To optLevel.UBound
                If optLevel(j).Value Then
                    strTmp = strTmp & optLevel(j).Caption & vbTab
                    Exit For
                End If
            Next
            
            Call .GetText(2, i, varTmp)
            strTmp = strTmp & varTmp & vbTab
            Call .GetText(3, i, varTmp)
            strTmp = strTmp & IIf(varTmp = "H", "High", IIf(varTmp = "L", "Low", IIf(varTmp = "N", "Normal", "All"))) & vbTab
            Call .GetText(4, i, varTmp)
            strTmp = strTmp & varTmp & vbTab
            Call .GetText(5, i, varTmp)
            strTmp = strTmp & IIf(varTmp = "A", "Accept", "Reject") & vbTab
            Call .GetText(6, i, varTmp)
            varTmp = Replace(varTmp, vbTab, String(5, " "))
            varTmp = Replace(varTmp, vbNewLine, COL_DIV)
            strTmp = strTmp & varTmp & vbTab & vbNewLine
            
            Call .GetText(8, i, varTmp)
            If varTmp <> "" Then
                strTmp = strTmp & Replace(Format(dtpFdate.Value, "yyyy/MM/dd") & "~" & Format(dtpTdate.Value, "yyyy/MM/dd"), "-", "/") & vbTab
                strTmp = strTmp & txtCtrlCd.Text & String(5, " ") & lblCtrlNm.Caption & vbTab
                strTmp = strTmp & Trim(medGetP(cboSection.Text, 1, COL_DIV)) & vbTab
                strTmp = strTmp & lblEqp.ToolTipText & vbTab
                For j = optAR.LBound To optAR.UBound
                    If optAR(j).Value Then
                        strTmp = strTmp & optAR(j).Caption & vbTab
                        Exit For
                    End If
                Next
                
                For j = optLevel.LBound To optLevel.UBound
                    If optLevel(j).Value Then
                        strTmp = strTmp & optLevel(j).Caption & vbTab
                        Exit For
                    End If
                Next
                
                Call .GetText(8, i, varTmp)
                strTmp = strTmp & varTmp & vbTab
                Call .GetText(9, i, varTmp)
                strTmp = strTmp & IIf(varTmp = "H", "High", IIf(varTmp = "L", "Low", IIf(varTmp = "N", "Normal", "All"))) & vbTab
                Call .GetText(10, i, varTmp)
                strTmp = strTmp & varTmp & vbTab
                Call .GetText(11, i, varTmp)
                strTmp = strTmp & IIf(varTmp = "A", "Accept", "Reject") & vbTab
                Call .GetText(12, i, varTmp)
                varTmp = Replace(varTmp, vbTab, String(5, " "))
                varTmp = Replace(varTmp, vbNewLine, COL_DIV)
                strTmp = strTmp & varTmp & vbTab & vbNewLine
            End If
        Next
    End With
    
    strTmp = Mid(strTmp, 1, Len(strTmp) - 2)
    
    lngFNo = FreeFile
    
    Open strFNm For Output As #lngFNo
    Print #lngFNo, strTmp
    Close #lngFNo
    With crtRpt
        .ReportFileName = strRptNm

        .ParameterFields(0) = "HospNm;" & ObjSysInfo.Hospital & ";true"

        .RetrieveDataFiles
        .WindowState = 2 ' crptMaximized
        .Destination = crptToWindow
        .Action = 1
        .Reset
    End With
End Sub

Private Sub PrintTable()
'''    Dim vbConLn As String
'''    Dim vbRefLn As String
'''    Dim vbCntLn As String
'''    Dim font1 As String
'''    Dim i As Long
'''
'''    For i = optLevel.LBound To optLevel.UBound
'''        If optLevel(i).Value = True Then vbConLn = vbConLn & "Level(" & optLevel(i).Caption & ")   "
'''    Next i
'''
'''    For i = optAR.LBound To optAR.UBound
'''        If optAR(i).Value = True Then vbConLn = vbConLn & "ÆÇÁ¤(" & optAR(i).Caption & ")   "
'''    Next i
'''
'''    tblResult.Row = -1: tblResult.Col = -1
'''    tblResult.FontSize = 8
'''    tblResult.PrintAbortMsg = "Ãâ·ÂÁß - Ãë¼ÒÇÏ·Á¸é Cancel Key¸¦ ´©¸£¼¼¿ä"
'''    tblResult.PrintJobName = "Quality Control"
'''    font1 = "/fn""±¼¸²Ã¼""/fz""10""/fb1/fi0/fu0/fk0/fs1"
'''    tblResult.PrintHeader = font1 & "/f1/l<< Quality Control : " & Trim(Mid(lstTestItem.Text, 1, 30)) & " >>           /rPAGE : /p/n/n" & _
'''                            "/lControl : " & txtCtrlCd.Text & "   " & lblCtrlNm.Caption & "/n/n" & _
'''                            "/l±â°£ : " & Format(dtpFdate.Value, CS_DateLongFormat) & " ~ " & Format(dtpTdate.Value, CS_DateLongFormat) & "/n" & _
'''                            "/lÁ¶°Ç : " & vbConLn & "/n/n"    '      Lot Number : " & lblLotNo & "     Open Date : " & lblOpenDate & "     Expire Date : " & lblExpDate & "/n" & _
'''
'''    tblResult.PrintBorder = True
'''    tblResult.PrintColHeaders = True
'''    tblResult.PrintColor = True
'''    tblResult.PrintGrid = True
'''    tblResult.PrintMarginTop = 1440
'''    tblResult.PrintMarginBottom = 1220
'''    tblResult.PrintMarginLeft = 150
'''    tblResult.PrintMarginRight = 50
'''    tblResult.PrintType = PrintTypeAll
'''    tblResult.PrintRowHeaders = True
'''    tblResult.PrintShadows = True
'''    tblResult.PrintUseDataMax = False
'''    ' Perform the printing action
'''    tblResult.Action = 13  'ActionPrint
'''
'''    tblResult.Row = -1: tblResult.Col = -1
'''    tblResult.FontSize = 10
    
End Sub

Private Sub cmdQuery_Click()
'    Dim i As Integer
'    Dim strSQL As String
'    Dim itmX As ListItem
'
'    If lstTestItem.ListCount < 1 Then Exit Sub
'    If lstTestItem.SelCount < 1 Then Exit Sub
'
'    Screen.MousePointer = vbArrowHourglass
'    lblMsg.Visible = True
'    DoEvents
'
'    Set objQcReview = New clsQcReview
'    With objQcReview
'
'        .CtrlCd = txtCtrlCd.Text
        'Call .GetControlInform(txtCtrlCd.Text, GetLevel, Trim(Mid(lstTestItem.Text, 31)))
'
'        Call LoadResult
'        Call LoadLotNo
'
'    End With
'
'    lblMsg.Visible = False
'    Screen.MousePointer = vbDefault
    Debug.Print Right(cboSection, 2)
    Call LoadData
End Sub

Private Sub LoadData()
    Dim i As Integer
    Dim strSql As String
    Dim itmX As ListItem
    
    'If lstTestItem.ListCount < 1 Then Exit Sub
    'If lstTestItem.SelCount < 1 Then Exit Sub
    
    MousePointer = vbHourglass
    lblMsg.Visible = True
    DoEvents
    
    Set objQcReview = New clsQcReview
    With objQcReview
        
        '.CtrlCd = txtCtrlCd.Text
       'Call .GetControlInform(txtCtrlCd.Text, GetLevel, Trim(Mid(lstTestItem.Text, 31)))
        
        'Debug.Print txtCtrlCd.Text
        Call LoadResult
        'Call LoadLotNo
        
    End With
    
    lblMsg.Visible = False
    MousePointer = vbDefault
End Sub

Private Sub LoadLotNo()
'''    Dim i As Long
'''    Dim itmX As ListItem
'''
'''    lvwLotNo.ListItems.Clear
'''
'''    With objQcReview
'''        For i = 1 To .ItemCount
'''            Set itmX = lvwLotNo.ListItems.Add(, .Item(i).KeyString, .Item(i).Lotno)
'''            itmX.SubItems(1) = .Item(i).OpenDt
'''            itmX.SubItems(2) = .Item(i).LevelNm
'''            itmX.SubItems(3) = .Item(i).ExpDt
'''            itmX.SubItems(4) = .Item(i).MeanVal
'''            itmX.SubItems(5) = .Item(i).CvVal
'''            itmX.SubItems(6) = .Item(i).SdVal
'''            itmX.SubItems(7) = .Item(i).TotCnt
'''            itmX.SubItems(8) = .Item(i).AccCnt
'''            itmX.SubItems(9) = .Item(i).RejCnt
'''        Next
'''    End With
End Sub

Private Sub LoadResult()
    Dim i As Long, j As Long, k As Long, l As Long, m As Long
    Dim strLOTNO As String
    Dim itmX As ListItem
    Dim blnTF As Boolean
    Dim lngListCount As Long
    Dim strTmp  As String
    Dim M1, M2, M3, S1, S2, S3, C1, C2, C3 As Long
    Dim strSql As String
    Dim strKey      As String
    Dim strKeyFlag  As String
    Dim lngRowNo    As Double
    Dim strLow      As String
    Dim strHigh     As String
    Dim varTmp
    Dim Rs, Ds      As Recordset
    
    strSql = "         SELECT a.ctrlcd, k.ctrlnm, a.lotno, a.workarea,a.accdt,a.accseq,a.testcd,a.rstval, " & vbCrLf
    strSql = strSql & "       a.rstcd,a.rstunit, a.rsttype,a.rstdiv,a.radiv,a.detailfg,a.vfydt,a.vfytm,a.vfyid, " & vbCrLf
    strSql = strSql & "       a.mfyfg, a.txtfg,a.autofg,a.eqpcd,a.method,c.rcvdt,c.rcvtm,a.ctrlcd,a.levelcd,a.lotno, " & vbCrLf
    strSql = strSql & "       h.testnm, h.abbrnm5, h.abbrnm10, h.txttype,  i.avalval, i.meanval, i.sdval, i.refcd, " & vbCrLf
    strSql = strSql & "       i.cvval, i.minval, i.maxval, i.wmset,  j.field1 as methodnm  " & vbCrLf
    strSql = strSql & "FROM   s2lab026 a, s2lab201 c, s2lab001 h, s2lab024 i, s2lab032 j, S2LAB021 k " & vbCrLf
    strSql = strSql & "WHERE  1 = 1 " & vbCrLf
'    strSql = strSql & "AND    a.ctrlcd =  '02_ADA'"
'    strSQL = strSQL & "AND  a.levelcd = 'H'"
'    strSQL = strSQL & "AND  a.lotno = '14412'"
'    strSql = strSql & "AND  a.testcd = 'C2200' " & vbCrLf
'    strSql = strSql & "AND  k.CTRLCD = '42 Dxc800' " & vbCrLf
    strSql = strSql & "AND  a.ctrlcd = k.CTRLCD " & vbCrLf
    ' °­¼º¼ö»ù ¿ä±¸·Î ¼½¼¾¿¡¼­ ÇÐºÎ·Î º¯°æ
    'strSql = strSql & "AND  k.SECTCD = '" & Right(cboSection, 2) & "' " & vbCrLf
    strSql = strSql & "AND  k.WORKAREA = '" & Right(cboWorkarea, 2) & "' " & vbCrLf
    strSql = strSql & "AND  a.vfydt between '" & Format(dtpFdate.Value, "yyyymmdd") & "' AND '" & Format(dtpTdate.Value, "yyyymmdd") & "' " & vbCrLf
    strSql = strSql & "AND   a.radiv in ('A','R')  AND   h.testcd = a.testcd " & vbCrLf
    strSql = strSql & "AND   h.applydt = (select max(applydt) from s2lab001 where testcd = h.testcd) " & vbCrLf
    strSql = strSql & "AND   c.workarea = a.workarea  AND  c.accdt = a.accdt  AND c.accseq = a.accseq " & vbCrLf
    strSql = strSql & "AND   i.ctrlcd = a.ctrlcd  AND  i.levelcd = a.levelcd  AND i.lotno = a.lotno " & vbCrLf
'    strSql = strSql & "AND   k.levelcd = a.levelcd " & vbCrLf
    strSql = strSql & "AND   i.testcd = a.testcd " & vbCrLf
'    strSQL = strSQL & "AND   j.cdindex(+) = 'C240'
    strSql = strSql & "AND   j.cdval1(+) = a.method " & vbCrLf
    strSql = strSql & "ORDER BY  a.ctrlcd, k.ctrlnm, a.testcd, h.rptseq, a.accdt, a.accseq " & vbCrLf
    
    'Debug.Print strSql
    Set Rs = New Recordset
    Rs.Open strSql, DBConn
    strKey = ""
    strLow = "N"
    strHigh = "N"
    strKeyFlag = ""
    lngRowNo = 0
    tblResult.MaxRows = 0
    While (Not Rs.EOF)
        With tblResult
            If strKeyFlag <> Trim("" & Rs.Fields("ctrlcd").Value) & Trim("" & Rs.Fields("testcd").Value) & Trim("" & Rs.Fields("VfyDt").Value) & Left(Trim("" & Rs.Fields("VfyTm").Value), 4) Then
               strKeyFlag = Trim("" & Rs.Fields("ctrlcd").Value) & Trim("" & Rs.Fields("testcd").Value) & Trim("" & Rs.Fields("VfyDt").Value) & Left(Trim("" & Rs.Fields("VfyTm").Value), 4)
               .MaxRows = .MaxRows + 1
               lngRowNo = lngRowNo + 1
               .SetText 1, .MaxRows, Str(lngRowNo)
            End If
            
            If .MaxRows = 0 Then
               .MaxRows = .MaxRows + 1
               lngRowNo = lngRowNo + 1
               .SetText 1, .MaxRows, Str(lngRowNo)
            End If
            
            strKey = ""
            Select Case Trim("" & Rs.Fields("levelcd").Value)
             Case "L"
'''                 .GetText 5, .MaxRows, varTmp: strKey = varTmp
'''                 If Len(Trim(strKey & "")) > 0 Then
'''                    .MaxRows = .MaxRows + 1
'''                    lngRowNo = lngRowNo + 1
'''                    .SetText 1, .MaxRows, Str(lngRowNo)
'''                 End If
                 
                 .SetText 5, .MaxRows, Trim("" & Rs.Fields("VfyDt").Value) & "/" & Mid(Trim("" & Rs.Fields("VfyTm").Value), 1, 4) '°Ë»çÀÏ½Ã
                 .SetText 6, .MaxRows, Trim("" & Rs.Fields("levelcd").Value)   'Lovel
                 .SetText 7, .MaxRows, Trim("" & Rs.Fields("rstcd").Value)   '°á°ú
                 .SetText 8, .MaxRows, Trim("" & Rs.Fields("RaDiv").Value)    'AR
                 .SetText 9, .MaxRows, " "   'Remark
             Case "N"
'''                 .GetText 10, .MaxRows, varTmp: strKey = varTmp
'''                 If Len(Trim(strKey & "")) > 0 Then
'''                    .MaxRows = .MaxRows + 1
'''                    lngRowNo = lngRowNo + 1
'''                    .SetText 1, .MaxRows, Str(lngRowNo)
'''                 End If
                 
                 .SetText 10, .MaxRows, Trim("" & Rs.Fields("VfyDt").Value) & "/" & Mid(Trim("" & Rs.Fields("VfyTm").Value), 1, 4) '°Ë»çÀÏ½Ã
                 .SetText 11, .MaxRows, Trim("" & Rs.Fields("levelcd").Value)   'Lovel
                 .SetText 12, .MaxRows, Trim("" & Rs.Fields("rstcd").Value)   '°á°ú
                 .SetText 13, .MaxRows, Trim("" & Rs.Fields("RaDiv").Value)    'AR
                 .SetText 14, .MaxRows, " "   'Remark
             Case "H"
'''                 .GetText 15, .MaxRows, varTmp: strKey = varTmp
'''                 If Len(Trim(strKey & "")) > 0 Then
'''                    .MaxRows = .MaxRows + 1
'''                    lngRowNo = lngRowNo + 1
'''                    .SetText 1, .MaxRows, Str(lngRowNo)
'''                 End If
                 
                 .SetText 15, .MaxRows, Trim("" & Rs.Fields("VfyDt").Value) & "/" & Mid(Trim("" & Rs.Fields("VfyTm").Value), 1, 4) '°Ë»çÀÏ½Ã
                 .SetText 16, .MaxRows, Trim("" & Rs.Fields("levelcd").Value)   'Lovel
                 .SetText 17, .MaxRows, Trim("" & Rs.Fields("rstcd").Value)   '°á°ú
                 .SetText 18, .MaxRows, Trim("" & Rs.Fields("RaDiv").Value)    'AR
                 .SetText 19, .MaxRows, " "   'Remark
            End Select

            .SetText 2, .MaxRows, Trim("" & Rs.Fields("ctrlcd").Value)
            .SetText 3, .MaxRows, Trim("" & Rs.Fields("ctrlnm").Value)
            .SetText 4, .MaxRows, Trim("" & Rs.Fields("testnm").Value)
            .SetText 20, .MaxRows, Trim("" & Rs.Fields("VfyDt").Value) & "/" & Mid(Trim("" & Rs.Fields("VfyTm").Value), 1, 4) '°Ë»çÀÏ½Ã
        End With
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
'''    With objQcReview
'''        Call .GetQcResult_ALL(a, b, c)
'''    End With
End Sub

Private Function GetLevel() As String
    Dim i As Long
    
    For i = optLevel.LBound To optLevel.UBound
        If optLevel(i).Value Then
            GetLevel = optLevel(i).Tag
            Exit For
        End If
    Next
End Function

Private Function GetAR() As String
    Dim i As Long
    
    For i = optAR.LBound To optAR.UBound
        If optAR(i).Value Then
            GetAR = optAR(i).Tag
            Exit For
        End If
    Next
End Function

Private Function GetTestCd() As String
'''    Dim i As Long
'''    Dim strTmp As String
'''
'''    If cmdPrint.Tag = "cmdPrint" Then
'''        strTmp = ""
'''        For i = 0 To lstTestItem.ListCount - 1
'''            strTmp = strTmp & "'" & Mid(lstTestItem.List(i), 31) & "',"
'''        Next
'''
'''        GetTestCd = Mid(strTmp, 1, Len(strTmp) - 1)
'''    Else
'''        For i = 0 To lstTestItem.ListCount - 1
'''            If lstTestItem.Selected(i) Then
'''                strTmp = Mid(lstTestItem.List(i), 31)
'''                Exit For
'''            End If
'''        Next
'''
'''        GetTestCd = "'" & strTmp & "'"
'''    End If
End Function

Private Sub cmdQuit_Click()
    fraGraph.Visible = False
    Call InitGraph
End Sub

Private Sub cmdQuitAll_Click()
    fraGrpAll.Visible = False
    Call InitGraph
End Sub

Private Sub Form_Load()
        txtCtrlCd.Text = ""
    lblCtrlNm.Caption = ""
    
    dtpTdate.Value = GetSystemDate
    dtpFdate.Value = DateAdd("d", -7, GetSystemDate)

    Call InitForm
    Call InitGraph
    
'    Call LoadSection
    Call LoadWorkArea
End Sub

Private Sub InitForm()
    
    
    lblEqp.Caption = ""
    'lstTestItem.Clear
    'lvwLotNo.ListItems.Clear
    Call medClearTable(tblResult)
    tblResult.MaxRows = 0
    lblMsg.Visible = False
    'txtMeanL.Text = ""
'    txtMeanN.Text = ""
'    txtMeanH.Text = ""
'    txtSDL.Text = ""
'    txtSDN.Text = ""
'    txtSDH.Text = ""
'    txtCVL.Text = ""
'    txtCVN.Text = ""
'    txtCVH.Text = ""

End Sub


Private Sub InitGraph()
    Dim i As Long
    
    On Error Resume Next
    
    For i = 0 To 3
        Erase CalRefVal
        chkPrint(i).Value = 0
        With cfxRst(i)
            .ClearData CD_VALUES
            .ClearLegend CHART_LEGEND
            .ClearData CD_CONSTANTLINES
            .ClearData CD_STRIPES
            .CloseData COD_VALUES
        End With
    Next
End Sub

Private Sub LoadWorkArea()
    Dim objWA As clsLISSqlQc
    Dim Rs As Recordset
    
    Set objWA = New clsLISSqlQc
       
    Set Rs = New Recordset
    Rs.Open objWA.GetWorkArea, DBConn
       
    cboWorkarea.Clear
    cboWorkarea.addItem " Àü Ã¼ "
    
    Do Until Rs.EOF
        cboWorkarea.addItem Format(Rs.Fields("cdval1").Value & "", "!" & String(10, "@")) & Format(Rs.Fields("field1").Value & "", "!" & String(100, "@")) & COL_DIV & _
                            Rs.Fields("cdval1").Value & ""
        
        Rs.MoveNext
    Loop
    
'    If cboWorkArea.ListCount > 0 Then cboWorkArea.ListIndex = 0
    
    Set Rs = Nothing
    Set objWA = Nothing
End Sub

Private Sub LoadSection()
    
    
    Dim Rs As Recordset
    Dim strSql As String
    Dim strItem As String
    
    strSql = " select * from " & T_LAB032 & _
             " where " & DBW("cdindex=", LC3_Section)
    Set Rs = New Recordset
    Rs.Open strSql, DBConn
    
    cboSection.Clear
        
    cboSection.addItem " Àü Ã¼ "
    
    Do Until Rs.EOF
        cboSection.addItem Format(Rs.Fields("field1").Value & "", "!" & String(100, "@")) & COL_DIV & _
                           Rs.Fields("cdval1").Value & ""
    
        Rs.MoveNext
    Loop
    
    cboSection.ListIndex = 0
    
    Set Rs = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objCtrl = Nothing
    Set objQcReview = Nothing
End Sub

Private Sub lstControl_Click()
    Dim strControl As String
    
    strControl = lstControl.Text
       
    txtCtrlCd.Text = Trim(Mid(strControl, 1, 12))
    lblCtrlNm.Caption = Trim(Mid(strControl, 13, 35))
    lblCtrlNm.ToolTipText = lblCtrlNm.Caption
    lblEqp.Caption = Trim(Mid(strControl, 58, 32)) & Format(Trim(Mid(strControl, 48, 8)), String(100, "@"))
    lblEqp.ToolTipText = Trim(Mid(strControl, 58, 32))
'    lblSection.Caption = Trim(Mid(strControl, 94, 50)) & Format(Trim(Mid(strControl, 90, 4)), String(30, "@"))
'    lblSection.ToolTipText = Trim(Mid(strControl, 94, 50))
    
    If cboSection.ListIndex < 1 Then cboSection.ListIndex = medComboFind(cboSection, Format(Trim(Mid(strControl, 94, 50)), "!" & String(100, "@")) & COL_DIV & Trim(Mid(strControl, 90, 4)))
    
    
    lstControl.Clear
    lstControl.Visible = False
    
    Call LoadTestItem
End Sub

Private Sub LoadTestItem()
'''    Dim Rs As Recordset
'''    Dim strSql As String
'''
'''    strSql = " select distinct a.testcd, b.testnm from " & T_LAB022 & " a, " & T_LAB001 & " b  " & _
'''             " where  " & DBW("a.ctrlcd=", Trim(txtCtrlCd.Text)) & _
'''             " and    b.testcd = a.testcd " & _
'''             " and    b.applydt = (select max(applydt) from " & T_LAB001 & _
'''                                "  where testcd = b.testcd) "
'''
'''    Set Rs = New Recordset
'''    Rs.Open strSql, DBConn
'''
'''    lstTestItem.Clear
'''    Do Until Rs.EOF
'''        lstTestItem.addItem Format(Rs.Fields("testnm").Value & "", "!" & String(30, "@")) & _
'''                            Rs.Fields("testcd").Value & ""
'''
'''        Rs.MoveNext
'''    Loop
'''
'''    If lstTestItem.ListCount = 0 Then
'''        MsgBox "¼³Á¤µÈ °Ë»çÇ×¸ñÀÌ ¾ø½À´Ï´Ù.", vbExclamation
'''    Else
'''        lstTestItem.ListIndex = 0
'''    End If
'''
'''    Set Rs = Nothing
End Sub

Private Sub lstTestItem_Click()
'''    If Screen.ActiveControl.Name <> lstTestItem.Name Then Exit Sub
'''    If lstTestItem.ListCount = 0 Then Exit Sub
'''    If lstTestItem.SelCount = 0 Then Exit Sub
    
    Call LoadData
End Sub

'''Private Sub lvwLotNo_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'''    Static lngOrder(2) As Integer
'''
'''    With lvwLotNo
'''        Select Case ColumnHeader.Index - 1
'''            Case 0
'''                .SortKey = 0
'''                .SortOrder = IIf(lngOrder(0) = 0, lvwAscending, lvwDescending)
'''                .Sorted = True
'''                lngOrder(0) = (lngOrder(0) + 1) Mod 2
'''            Case 1
'''                .SortKey = 1
'''                .SortOrder = IIf(lngOrder(1) = 0, lvwAscending, lvwDescending)
'''                .Sorted = True
'''                lngOrder(1) = (lngOrder(1) + 1) Mod 2
'''            Case 2
'''                .SortKey = 2
'''                .SortOrder = IIf(lngOrder(1) = 0, lvwAscending, lvwDescending)
'''                .Sorted = True
'''                lngOrder(2) = (lngOrder(2) + 1) Mod 2
'''        End Select
'''    End With
'''End Sub

Private Sub lvwLotNo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'''    If lvwLotNo.ListItems.Count = 0 Then Exit Sub
'''    If lvwLotNo.SelectedItem.ListSubItems.Count = 0 Then Exit Sub
'''
'''    If Button = 2 Then
'''        Set objPop = New clsPopupMenu
'''        With objPop
'''            .AddMenu MENU_ONE, "ºÐÇÒ GRAPH"
'''            .AddMenu MENU_ALL, "ÅëÇÕ GRAPH"
'''
'''            .PopupMenus Me.hwnd
'''        End With
'''        Set objPop = Nothing
''''        Set mnuPopup = frmControls.mnuPopup
''''        Set mnuGraph = frmControls.mnuSub1
''''        Set mnuGraphAll = frmControls.mnuSub2
''''
''''        mnuGraph.Caption = "ºÐÇÒ Graph"
''''        mnuGraphAll.Caption = "ÅëÇÕ Graph"
''''
''''        Me.PopupMenu mnuPopup
''''
''''        Set mnuGraph = Nothing
''''        Set mnuGraphAll = Nothing
''''        Set mnuPopup = Nothing
'''    End If
End Sub

'Private Sub mnuGraph_Click()
''µû·Îµû·Î º¸¿©ÁÖ´Â ±×·¡ÇÁ
'    Dim lngCnt As Long
'    Dim i As Long
'
'    For i = 1 To lvwLotNo.ListItems.Count
'        If lvwLotNo.ListItems(i).Checked Then
'            lngCnt = lngCnt + 1
'        End If
'    Next
'
'    If lngCnt = 0 Then Exit Sub
'
'    Call InitGraph
'
'    lngCnt = 0
'    For i = 1 To lvwLotNo.ListItems.Count
'        If lvwLotNo.ListItems(i).Checked Then
'            Call ShowGraph(lngCnt, objQcReview.Item(i).KeyString)
'            Call SetRefRange(lngCnt, i)
'            lngCnt = lngCnt + 1
'            If lngCnt >= 3 Then Exit For
'        End If
'    Next
'
'    fraGraph.Visible = True
'    fraGraph.ZOrder 0
'End Sub

'Private Sub mnuGraphAll_Click()
'    Dim lngCnt As Long
'    Dim i As Long
'
'    For i = 1 To lvwLotNo.ListItems.Count
'        If lvwLotNo.ListItems(i).Checked Then
'            lngCnt = lngCnt + 1
'        End If
'    Next
'
'    If lngCnt = 0 Then Exit Sub
'
'    Call InitGraph
'    Call ShowGraphAll
''    Call ShowAdvancedGraph
'
'    fraGrpAll.Visible = True
'    fraGrpAll.ZOrder 0
'End Sub

Private Sub ShowGraphAll()
'ÇÏ³ªÀÇ Ã­Æ®¿¡ ¼±ÅÃµÈ°Å ´Ùº¸¿©ÁÖ´Â ±×·¡ÇÁ
'''    Dim i As Long
'''    Dim j As Long
'''    Dim k As Long
'''    Dim strKey As String
'''    Dim aryKey() As String
'''    Dim Min As Double       'YÃà ÃÖ¼Ò°ª
'''    Dim Max As Double       'YÃà ÃÖ´ë°ª
'''    Dim dblMean As Double   'Mean °ª
'''    Dim dblSd As Double     'Sd°ª
'''    Dim lngTCnt As Long     'Total °¹¼ö
'''    Dim avgSd As Double     'Sd Æò±Õ°ª
'''    Dim SCnt As Long        '¾¾¸®Áî °¹¼ö
'''    Dim CCnt As Long        'ÄÃ·³°¹¼ö
'''    Dim Index As Long       'Index
'''    'Stripe¿Í Line¸¦ ±×¸®±â À§ÇÑ °è»êÄ¡
'''    Dim N_1SdVal As Double, N_2SdVal As Double, N_3SdVal As Double, N_4SdVal As Double
'''    Dim P_1SdVal As Double, P_2SdVal As Double, P_3SdVal As Double, p_4SdVal As Double
'''
'''    With lvwLotNo
'''        For i = 1 To .ListItems.Count
'''            If .ListItems(i).Checked Then
'''
'''                Min = .ListItems(i).SubItems(4)
'''                Exit For
'''            End If
'''        Next
'''
'''        For i = 1 To .ListItems.Count
'''            If .ListItems(i).Checked Then
'''                dblMean = .ListItems(i).SubItems(4)
'''                dblSd = .ListItems(i).SubItems(6)
'''                lngTCnt = .ListItems(i).SubItems(7)
'''
'''                avgSd = avgSd + dblSd
'''
'''                If CCnt >= lngTCnt Then
'''                    CCnt = CCnt
'''                Else
'''                    CCnt = lngTCnt
'''                End If
'''
'''                If Min <= (dblMean - (dblSd * 5)) Then
'''                    Min = Min
'''                Else
'''                    Min = dblMean - (dblSd * 5)
'''                End If
'''
'''                If Max >= (dblMean + (dblSd * 5)) Then
'''                    Max = Max
'''                Else
'''                    Max = dblMean + (dblSd * 5)
'''                End If
'''
'''                strKey = strKey & objQcReview.Item(i).KeyString & Fld_Div & _
'''                         lngTCnt & Fld_Div & _
'''                         dblMean & Fld_Div & _
'''                         dblSd & Rec_Div
'''                SCnt = SCnt + 1
'''            End If
'''        Next i
'''    End With
'''
'''    aryKey = Split(strKey, Rec_Div)
'''
'''    With cfxRst(3)
'''        .ClearData CD_VALUES
'''        .ClearLegend CHART_LEGEND
'''        .ClearData CD_CONSTANTLINES
'''        .ClearData CD_STRIPES
'''
'''        .OpenDataEx COD_VALUES, SCnt, CCnt
'''        .OpenDataEx COD_CONSTANTS, SCnt * 9, 0
'''        .OpenDataEx COD_STRIPES, SCnt * 8, 0
'''
'''        .PointLabels = True
'''        .XLegFont.Size = 7
'''        .YLegFont.Size = 7
'''
'''        .LineWidth = 1.5
'''
'''        .Adm(CSA_MIN) = Min
'''        .Adm(CSA_MAX) = Max
'''        .Adm(CSA_GAP) = avgSd / SCnt
'''        .Axis(AXIS_Y).STEP = (Max - Min) / SCnt
'''
'''        For i = LBound(aryKey) To UBound(aryKey) - 1
'''            dblMean = medGetP(aryKey(i), 3, Fld_Div)
'''            dblSd = medGetP(aryKey(i), 4, Fld_Div)
'''
'''            Index = 0
'''            For j = 1 To objQcReview.TestCount
'''                If objQcReview.TestItem(j).KeyString = medGetP(aryKey(i), 1, Fld_Div) Then
'''                    .Axis(AXIS_Y).Decimals = objQcReview.TestItem(j).AvalVal
'''                    .Series(i).Legend = "LotNo : " & medGetP(aryKey(i), 1, Fld_Div)
'''                    Index = Index + 1
'''                    .Series(i).Yvalue(Index) = Val(objQcReview.TestItem(j).RstCd)
'''
'''
'''                    If Index <> CCnt Then
'''                        If medGetP(aryKey(i), 2, Fld_Div) = Index Then
'''                            For k = Index To CCnt - 1
'''                                .Series(i).Yvalue(k) = CFX_NULL
'''                            Next
'''                        End If
'''                    End If
'''                End If
'''            Next j
'''
'''            N_1SdVal = dblMean - dblSd
'''            N_2SdVal = dblMean - (dblSd * 2)
'''            N_3SdVal = dblMean - (dblSd * 3)
'''            N_4SdVal = dblMean - (dblSd * 4)
'''            P_1SdVal = dblMean + dblSd
'''            P_2SdVal = dblMean + (dblSd * 2)
'''            P_3SdVal = dblMean + (dblSd * 3)
'''            p_4SdVal = dblMean + (dblSd * 4)
'''
'''            'Stripe ±×¸®±â
'''            '4SD(-)~3SD(-)
'''            .Stripe(8 * i).Axis = AXIS_Y
'''            .Stripe(8 * i).Color = Choose(1, &HF7F0F0, &HFFF1DF, &HDBF2FD, &HC0FFFF, &HC0FFFF, &HDBF2FD, &HFFF1DF, &HF7F0F0)
'''            .Stripe(8 * i).From = N_4SdVal
'''            .Stripe(8 * i).To = N_3SdVal
'''
'''            '3SD(-)~2SD(-)
'''            .Stripe(8 * i + 1).Axis = AXIS_Y
'''            .Stripe(8 * i + 1).Color = Choose(2, &HF7F0F0, &HFFF1DF, &HDBF2FD, &HC0FFFF, &HC0FFFF, &HDBF2FD, &HFFF1DF, &HF7F0F0)
'''            .Stripe(8 * i + 1).From = N_3SdVal
'''            .Stripe(8 * i + 1).To = N_2SdVal
'''
'''            '2SD(-)~1SD(-)
'''            .Stripe(8 * i + 2).Axis = AXIS_Y
'''            .Stripe(8 * i + 2).Color = Choose(3, &HF7F0F0, &HFFF1DF, &HDBF2FD, &HC0FFFF, &HC0FFFF, &HDBF2FD, &HFFF1DF, &HF7F0F0)
'''            .Stripe(8 * i + 2).From = N_2SdVal
'''            .Stripe(8 * i + 2).To = N_1SdVal
'''
'''            '1SD(-)~Mean
'''            .Stripe(8 * i + 3).Axis = AXIS_Y
'''            .Stripe(8 * i + 3).Color = Choose(4, &HF7F0F0, &HFFF1DF, &HDBF2FD, &HC0FFFF, &HC0FFFF, &HDBF2FD, &HFFF1DF, &HF7F0F0)
'''            .Stripe(8 * i + 3).From = N_1SdVal
'''            .Stripe(8 * i + 3).To = dblMean
'''
'''            'Mean~1SD(+)
'''            .Stripe(8 * i + 4).Axis = AXIS_Y
'''            .Stripe(8 * i + 4).Color = Choose(5, &HF7F0F0, &HFFF1DF, &HDBF2FD, &HC0FFFF, &HC0FFFF, &HDBF2FD, &HFFF1DF, &HF7F0F0)
'''            .Stripe(8 * i + 4).From = dblMean
'''            .Stripe(8 * i + 4).To = P_1SdVal
'''
'''            '1SD(+)~2SD(+)
'''            .Stripe(8 * i + 5).Axis = AXIS_Y
'''            .Stripe(8 * i + 5).Color = Choose(6, &HF7F0F0, &HFFF1DF, &HDBF2FD, &HC0FFFF, &HC0FFFF, &HDBF2FD, &HFFF1DF, &HF7F0F0)
'''            .Stripe(8 * i + 5).From = P_1SdVal
'''            .Stripe(8 * i + 5).To = P_2SdVal
'''
'''            '2SD(+)~3SD(+)
'''            .Stripe(8 * i + 6).Axis = AXIS_Y
'''            .Stripe(8 * i + 6).Color = Choose(7, &HF7F0F0, &HFFF1DF, &HDBF2FD, &HC0FFFF, &HC0FFFF, &HDBF2FD, &HFFF1DF, &HF7F0F0)
'''            .Stripe(8 * i + 6).From = P_2SdVal
'''            .Stripe(8 * i + 6).To = P_3SdVal
'''
'''            '3SD(+)~4SD(+)
'''            .Stripe(8 * i + 7).Axis = AXIS_Y
'''            .Stripe(8 * i + 7).Color = Choose(8, &HF7F0F0, &HFFF1DF, &HDBF2FD, &HC0FFFF, &HC0FFFF, &HDBF2FD, &HFFF1DF, &HF7F0F0)
'''            .Stripe(8 * i + 7).From = P_3SdVal
'''            .Stripe(8 * i + 7).To = p_4SdVal
'''
'''            'Line ±×¸®±â
'''            '4SD(+)
'''            .ConstantLine(9 * i).Value = p_4SdVal
'''            .ConstantLine(9 * i).Axis = AXIS_Y
'''            .ConstantLine(9 * i).Label = Choose(1, "", "3SD(+)", "2SD(+)", "1SD(+)", "Mean", "1SD(-)", "2SD(-)", "3SD(-)", "")
'''            .ConstantLine(9 * i).LineWidth = 1
'''            .ConstantLine(9 * i).LineStyle = Choose(1, 0, 4, 1, 2, 3, 2, 1, 4, 0)
'''            .ConstantLine(9 * i).Style = CC_COLORTEXT
'''            Select Case .ConstantLine(9 * i).LineStyle
'''               Case 0:
'''                  .ConstantLine(9 * i).LineColor = &H0&
'''               Case 3:
'''                  .ConstantLine(9 * i).LineColor = &H80&
'''               Case Else
'''                  .ConstantLine(9 * i).LineColor = &H808080
'''            End Select
'''
'''            '3SD(+)
'''            .ConstantLine(9 * i + 1).Value = P_3SdVal
'''            .ConstantLine(9 * i + 1).Axis = AXIS_Y
'''            .ConstantLine(9 * i + 1).Label = Choose(2, "", "3SD(+)", "2SD(+)", "1SD(+)", "Mean", "1SD(-)", "2SD(-)", "3SD(-)", "")
'''            .ConstantLine(9 * i + 1).LineWidth = 1
'''            .ConstantLine(9 * i + 1).LineStyle = Choose(2, 0, 4, 1, 2, 3, 2, 1, 4, 0)
'''            .ConstantLine(9 * i + 1).Style = CC_COLORTEXT
'''            Select Case .ConstantLine(9 * i + 1).LineStyle
'''               Case 0:
'''                  .ConstantLine(9 * i + 1).LineColor = &H0&
'''               Case 3:
'''                  .ConstantLine(9 * i + 1).LineColor = &H80&
'''               Case Else
'''                  .ConstantLine(9 * i + 1).LineColor = &H808080
'''            End Select
'''
'''            '2SD(+)
'''            .ConstantLine(9 * i + 2).Value = P_2SdVal
'''            .ConstantLine(9 * i + 2).Axis = AXIS_Y
'''            .ConstantLine(9 * i + 2).Label = Choose(3, "", "3SD(+)", "2SD(+)", "1SD(+)", "Mean", "1SD(-)", "2SD(-)", "3SD(-)", "")
'''            .ConstantLine(9 * i + 2).LineWidth = 1
'''            .ConstantLine(9 * i + 2).LineStyle = Choose(3, 0, 4, 1, 2, 3, 2, 1, 4, 0)
'''            .ConstantLine(9 * i + 2).Style = CC_COLORTEXT
'''            Select Case .ConstantLine(9 * i + 2).LineStyle
'''               Case 0:
'''                  .ConstantLine(9 * i + 2).LineColor = &H0&
'''               Case 3:
'''                  .ConstantLine(9 * i + 2).LineColor = &H80&
'''               Case Else
'''                  .ConstantLine(9 * i + 2).LineColor = &H808080
'''            End Select
'''
'''            '1SD(+)
'''            .ConstantLine(9 * i + 3).Value = P_1SdVal
'''            .ConstantLine(9 * i + 3).Axis = AXIS_Y
'''            .ConstantLine(9 * i + 3).Label = Choose(4, "", "3SD(+)", "2SD(+)", "1SD(+)", "Mean", "1SD(-)", "2SD(-)", "3SD(-)", "")
'''            .ConstantLine(9 * i + 3).LineWidth = 1
'''            .ConstantLine(9 * i + 3).LineStyle = Choose(4, 0, 4, 1, 2, 3, 2, 1, 4, 0)
'''            .ConstantLine(9 * i + 3).Style = CC_COLORTEXT
'''            Select Case .ConstantLine(9 * i + 3).LineStyle
'''               Case 0:
'''                  .ConstantLine(9 * i + 3).LineColor = &H0&
'''               Case 3:
'''                  .ConstantLine(9 * i + 3).LineColor = &H80&
'''               Case Else
'''                  .ConstantLine(9 * i + 3).LineColor = &H808080
'''            End Select
'''
'''            'Mean
'''            .ConstantLine(9 * i + 4).Value = dblMean
'''            .ConstantLine(9 * i + 4).Axis = AXIS_Y
'''            .ConstantLine(9 * i + 4).Label = Choose(5, "", "3SD(+)", "2SD(+)", "1SD(+)", "Mean", "1SD(-)", "2SD(-)", "3SD(-)", "")
'''            .ConstantLine(9 * i + 4).LineWidth = 1.5
'''            .ConstantLine(9 * i + 4).LineStyle = Choose(5, 0, 4, 1, 2, 3, 2, 1, 4, 0)
'''            .ConstantLine(9 * i + 4).Style = CC_COLORTEXT
'''            Select Case .ConstantLine(9 * i + 4).LineStyle
'''               Case 0:
'''                  .ConstantLine(9 * i + 4).LineColor = &H0&
'''               Case 3:
'''                  .ConstantLine(9 * i + 4).LineColor = &H80&
'''               Case Else
'''                  .ConstantLine(9 * i + 4).LineColor = &H808080
'''            End Select
'''
'''            '1SD(-)
'''            .ConstantLine(9 * i + 5).Value = N_1SdVal
'''            .ConstantLine(9 * i + 5).Axis = AXIS_Y
'''            .ConstantLine(9 * i + 5).Label = Choose(6, "", "3SD(+)", "2SD(+)", "1SD(+)", "Mean", "1SD(-)", "2SD(-)", "3SD(-)", "")
'''            .ConstantLine(9 * i + 5).LineWidth = 1
'''            .ConstantLine(9 * i + 5).LineStyle = Choose(6, 0, 4, 1, 2, 3, 2, 1, 4, 0)
'''            .ConstantLine(9 * i + 5).Style = CC_COLORTEXT
'''            Select Case .ConstantLine(9 * i + 5).LineStyle
'''               Case 0:
'''                  .ConstantLine(9 * i + 5).LineColor = &H0&
'''               Case 3:
'''                  .ConstantLine(9 * i + 5).LineColor = &H80&
'''               Case Else
'''                  .ConstantLine(9 * i + 5).LineColor = &H808080
'''            End Select
'''
'''            '2SD(-)
'''            .ConstantLine(9 * i + 6).Value = N_2SdVal
'''            .ConstantLine(9 * i + 6).Axis = AXIS_Y
'''            .ConstantLine(9 * i + 6).Label = Choose(7, "", "3SD(+)", "2SD(+)", "1SD(+)", "Mean", "1SD(-)", "2SD(-)", "3SD(-)", "")
'''            .ConstantLine(9 * i + 6).LineWidth = 1
'''            .ConstantLine(9 * i + 6).LineStyle = Choose(7, 0, 4, 1, 2, 3, 2, 1, 4, 0)
'''            .ConstantLine(9 * i + 6).Style = CC_COLORTEXT
'''            Select Case .ConstantLine(9 * i + 6).LineStyle
'''               Case 0:
'''                  .ConstantLine(9 * i + 6).LineColor = &H0&
'''               Case 3:
'''                  .ConstantLine(9 * i + 6).LineColor = &H80&
'''               Case Else
'''                  .ConstantLine(9 * i + 6).LineColor = &H808080
'''            End Select
'''
'''            '3SD(-)
'''            .ConstantLine(9 * i + 7).Value = N_3SdVal
'''            .ConstantLine(9 * i + 7).Axis = AXIS_Y
'''            .ConstantLine(9 * i + 7).Label = Choose(8, "", "3SD(+)", "2SD(+)", "1SD(+)", "Mean", "1SD(-)", "2SD(-)", "3SD(-)", "")
'''            .ConstantLine(9 * i + 7).LineWidth = 1
'''            .ConstantLine(9 * i + 7).LineStyle = Choose(8, 0, 4, 1, 2, 3, 2, 1, 4, 0)
'''            .ConstantLine(9 * i + 7).Style = CC_COLORTEXT
'''            Select Case .ConstantLine(9 * i + 7).LineStyle
'''               Case 0:
'''                  .ConstantLine(9 * i + 7).LineColor = &H0&
'''               Case 3:
'''                  .ConstantLine(9 * i + 7).LineColor = &H80&
'''               Case Else
'''                  .ConstantLine(9 * i + 7).LineColor = &H808080
'''            End Select
'''
'''            '4SD(-)
'''            .ConstantLine(9 * i + 8).Value = N_4SdVal
'''            .ConstantLine(9 * i + 8).Axis = AXIS_Y
'''            .ConstantLine(9 * i + 8).Label = Choose(9, "", "3SD(+)", "2SD(+)", "1SD(+)", "Mean", "1SD(-)", "2SD(-)", "3SD(-)", "")
'''            .ConstantLine(9 * i + 8).LineWidth = 1
'''            .ConstantLine(9 * i + 8).LineStyle = Choose(9, 0, 4, 1, 2, 3, 2, 1, 4, 0)
'''            .ConstantLine(9 * i + 8).Style = CC_COLORTEXT
'''            Select Case .ConstantLine(9 * i + 8).LineStyle
'''               Case 0:
'''                  .ConstantLine(9 * i + 8).LineColor = &H0&
'''               Case 3:
'''                  .ConstantLine(9 * i + 8).LineColor = &H80&
'''               Case Else
'''                  .ConstantLine(9 * i + 8).LineColor = &H808080
'''            End Select
'''        Next
'''
'''        .CloseData COD_VALUES
'''        .CloseData COD_CONSTANTS
'''        .CloseData COD_STRIPES
'''    End With
End Sub

Private Sub ShowGraph(ByVal lngGrp As Long, ByVal strKeyString As String)
    Dim i As Long, j As Long, k As Long
    Dim vbIndex As Long
    Dim strKey As String
    Dim dblValue As Double

    With objQcReview
        vbIndex = -1
        
        ReDim CalRefVal(lngGrp).calVALUE(0)
        CalRefVal(lngGrp).calMEAN = 0
        CalRefVal(lngGrp).calCV = 0
        CalRefVal(lngGrp).calSD = 0
        CalRefVal(lngGrp).calMIN = 999999999: CalRefVal(lngGrp).calMAX = -999999999:
        CalRefVal(lngGrp).calCOUNT = 0:  CalRefVal(lngGrp).calTOTAL = 0
        
        cfxRst(lngGrp).ClearData CD_VALUES
        cfxRst(lngGrp).ClearLegend CHART_LEGEND
        
        cfxRst(lngGrp).OpenDataEx COD_VALUES, 1, .TestCount
        cfxRst(lngGrp).PointLabels = True
        cfxRst(lngGrp).XLegFont.Size = 7
        cfxRst(lngGrp).YLegFont.Size = 7
               
        For j = 1 To .TestCount
        
            strKey = .TestItem(j).KeyString
            
            If Trim(strKey) = Trim(strKeyString) Then
                If .TestItem(j).VfyDt <> "" Then
                    vbIndex = vbIndex + 1
                    cfxRst(lngGrp).Axis(AXIS_Y).Decimals = .TestItem(j).AvalVal
                    cfxRst(lngGrp).Legend(vbIndex) = Format(Mid(.TestItem(j).VfyDt, 5), "0#/##") & vbCrLf & Format(Mid(.TestItem(j).VfyTm, 1, 4), "0#:##")
                    dblValue = Val(.TestItem(j).RstCd)
                    cfxRst(lngGrp).Value(vbIndex) = dblValue
                    
                    ReDim Preserve CalRefVal(lngGrp).calVALUE(vbIndex)

                    CalRefVal(lngGrp).calVALUE(vbIndex) = dblValue
                    CalRefVal(lngGrp).calCOUNT = CalRefVal(lngGrp).calCOUNT + 1
                    CalRefVal(lngGrp).calTOTAL = CalRefVal(lngGrp).calTOTAL + dblValue
            
                    If (dblValue < CalRefVal(lngGrp).calMIN) Then CalRefVal(lngGrp).calMIN = dblValue
                    If (dblValue > CalRefVal(lngGrp).calMAX) Then CalRefVal(lngGrp).calMAX = dblValue
                
                End If
            End If
            
        Next j
        
        On Error Resume Next
        chkPrint(lngGrp).Value = 1
        chkPrint(lngGrp).Enabled = True
        cfxRst(lngGrp).CloseData COD_VALUES
        cfxRst(lngGrp).OpenDataEx COD_VALUES, 1, vbIndex + 1
        If vbIndex < 0 Then
            chkPrint(lngGrp).Value = 0
            chkPrint(lngGrp).Enabled = False
            cfxRst(lngGrp).ClearData CD_ALLDATA
        End If
        cfxRst(lngGrp).CloseData COD_VALUES
        
    End With
    
End Sub

Private Sub ShowAdvancedGraph()
'''    Dim i As Long
'''    Dim strLotNo As String
'''    Dim strTmp As String
'''    Dim blnFirst As Boolean
'''    Dim strMin As String
'''    Dim strMax As String
'''    Dim strVfyDt As String
'''    Dim aryLotNo() As String
'''    Dim aryVfyDt() As String
'''    Dim j As Long
'''
'''    Dim strMean As String
'''
'''    For i = 1 To lvwLotNo.ListItems.Count
'''        '¸®½ºÆ®¿¡¼­ ¼±ÅÃÀÌ µÇ¾ß ÇÏ°í ¼±ÅÃµÈ ³ÑÁß °á°úº¸°í µÈ ³ÑÀÌ ÇÏ³ª ÀÌ»óÀÌ¾î¾ß ÇÑ´Ù..
'''        If lvwLotNo.ListItems(i).Checked And lvwLotNo.ListItems(i).SubItems(7) <> 0 Then
'''            strLotNo = strLotNo & lvwLotNo.ListItems(i).Text & Fld_Div
'''
'''            strMean = Val(strMean) + lvwLotNo.ListItems(i).SubItems(4)
'''        End If
'''    Next
'''
'''    For i = 1 To objQcReview.TestCount
'''        If InStr(strLotNo, objQcReview.TestItem(i).Lotno) > 0 Then
'''            strTmp = objQcReview.TestItem(i).RstCd
'''
'''            If blnFirst = False Then
'''                strMin = objQcReview.TestItem(i).RstCd
'''                blnFirst = True
'''            End If
'''
'''            If strMin > strTmp Then
'''                strMin = strTmp
'''            Else
'''                strMin = strMin 'Min °ª
'''            End If
'''
'''            If strMax > strTmp Then
'''                strMax = strMax 'Max °ª
'''            Else
'''                strMax = strTmp
'''            End If
'''
''''            If InStr(strVfyDt, Format(Mid(objQcReview.TestItem(i).VfyDt, 5), "0#/0#")) = 0 Then
''''                strVfyDt = strVfyDt & Format(Mid(objQcReview.TestItem(i).VfyDt, 5), "0#/0#") & Fld_Div
''''            End If
'''            If InStr(strVfyDt, Format(Mid(objQcReview.TestItem(i).VfyDt, 5), "0#/##") & vbNewLine & Format(Mid(objQcReview.TestItem(i).VfyTm, 1, 4), "0#:##")) = 0 Then
'''                strVfyDt = strVfyDt & Format(Mid(objQcReview.TestItem(i).VfyDt, 5), "0#/##") & vbNewLine & Format(Mid(objQcReview.TestItem(i).VfyTm, 1, 4), "0#:##") & Fld_Div
'''            End If
'''        End If
'''    Next
'''
'''
'''
'''    aryLotNo = Split(strLotNo, Fld_Div) 'ÀÌ³ÑÀÇ Ubound°¡ ¾¾¸®Áî °¹¼ö
'''    aryVfyDt = Split(strVfyDt, Fld_Div) 'ÀÌ³ÑÀÇ Ubound°¡ XÃà °¹¼ö
'''
'''    strMean = Val(strMean) / UBound(aryLotNo)
'''    '½îÆÃ
'''    For i = LBound(aryVfyDt) To UBound(aryVfyDt) - 1
'''        For j = i + 1 To UBound(aryVfyDt)
'''            If aryVfyDt(i) > aryVfyDt(j) Then
'''                strTmp = aryVfyDt(i)
'''                aryVfyDt(i) = aryVfyDt(j)
'''                aryVfyDt(j) = strTmp
'''            End If
'''        Next j
'''    Next i
'''
'''    strVfyDt = ""
'''    For i = LBound(aryVfyDt) To UBound(aryVfyDt)
'''        If aryVfyDt(i) <> "" Then
'''            strVfyDt = strVfyDt & aryVfyDt(i) & Fld_Div
'''        End If
'''    Next
'''
'''    aryVfyDt = Split(strVfyDt, Fld_Div)
'''
'''    fraGrpAll.Visible = True
'''    fraGrpAll.ZOrder 0
'''
'''    cfxRst(3).ClearData CD_VALUES
''''    cfxrst(3).ClearData CD_ALLDATA
'''
'''    cfxRst(3).OpenDataEx COD_VALUES, UBound(aryLotNo), UBound(aryVfyDt)
'''    cfxRst(3).OpenDataEx COD_CONSTANTS, UBound(aryLotNo), 0
'''
'''    cfxRst(3).Axis(AXIS_Y).STEP = Val(strMax) - Val(strMin)
'''    cfxRst(3).Axis(AXIS_Y).Min = strMin
'''    cfxRst(3).Axis(AXIS_Y).Max = strMax
'''    cfxRst(3).Axis(AXIS_Y).Decimals = 3
'''    cfxRst(3).PointLabels = False
'''
'''
'''    For i = 0 To UBound(aryVfyDt) - 1  'XÃà ±×¸®±â
'''        cfxRst(3).Legend(i) = aryVfyDt(i)
'''    Next
'''
'''    Dim objData As clsDictionary
'''    Dim aryData() As String
'''    Dim k As Long
'''
'''    Set objData = New clsDictionary
'''
'''    objData.Clear
'''    objData.FieldInialize "lotno", "data"   'vfydt,rstcd
'''    objData.DeleteAll
'''
'''    For i = 1 To objQcReview.TestCount
''''        If InStr(strLotNo, objQcReview.TestItem(i).Lotno) > 0 Then
'''        For j = LBound(aryLotNo) To UBound(aryLotNo)
'''            If aryLotNo(j) = objQcReview.TestItem(i).Lotno Then
''''            If objData.Exists(objQcReview.TestItem(i).Lotno) Then
''''                objData.KeyChange objQcReview.TestItem(i).Lotno
''''                objData.Fields("data") = objData.Fields("data") & Rec_Div & Format(Mid(objQcReview.TestItem(i).VfyDt, 5), "0#/0#") & Fld_Div & objQcReview.TestItem(i).RstCd & Fld_Div & objQcReview.TestItem(i).MeanVal
''''            Else
''''                objData.AddNew objQcReview.TestItem(i).Lotno, Format(Mid(objQcReview.TestItem(i).VfyDt, 5), "0#/0#") & Fld_Div & objQcReview.TestItem(i).RstCd & Fld_Div & objQcReview.TestItem(i).MeanVal
''''            End If
'''            If objData.Exists(objQcReview.TestItem(i).Lotno) Then
'''                objData.KeyChange objQcReview.TestItem(i).Lotno
'''                objData.Fields("data") = objData.Fields("data") & Rec_Div & Format(Mid(objQcReview.TestItem(i).VfyDt, 5), "0#/##") & vbNewLine & Format(Mid(objQcReview.TestItem(i).VfyTm, 1, 4), "0#:##") & Fld_Div & objQcReview.TestItem(i).RstCd & Fld_Div & objQcReview.TestItem(i).MeanVal
'''            Else
'''                objData.AddNew objQcReview.TestItem(i).Lotno, Format(Mid(objQcReview.TestItem(i).VfyDt, 5), "0#/##") & vbNewLine & Format(Mid(objQcReview.TestItem(i).VfyTm, 1, 4), "0#:##") & Fld_Div & objQcReview.TestItem(i).RstCd & Fld_Div & objQcReview.TestItem(i).MeanVal
'''            End If
'''            End If
'''        Next
''''        End If
'''    Next
'''
'''    i = 0
'''    objData.MoveFirst
'''    Do Until objData.EOF
'''
''''        Debug.Print objData.GetLine
'''
'''        aryData = Split(objData.Fields("data"), Rec_Div)
''''
'''        For j = LBound(aryData) To UBound(aryData)
'''            For k = 0 To UBound(aryVfyDt) - 1
''''                Debug.Print aryVfyDt(k)
'''                If medGetP(aryData(j), 1, Fld_Div) = aryVfyDt(k) Then
'''
'''                    cfxRst(3).Series(i).Yvalue(k) = Val(medGetP(aryData(j), 2, Fld_Div)) * (Val(strMean) / Val(medGetP(aryData(j), 3, Fld_Div)))  'medGetP(aryData(j), 2, Fld_Div)
'''                Else
'''                    If j = 0 Then
'''                        cfxRst(3).Series(i).Yvalue(k) = CFX_NULL
'''                    End If
'''                End If
'''            Next
'''        Next
'''
''''        Call DrawLine(3, i, medGetP(aryData(0), 3, Fld_Div))
'''
'''        i = i + 1 '¾¾¸®Áî °ª
'''        objData.MoveNext
'''    Loop
'''
'''    cfxRst(3).ConstantLine(0).Axis = AXIS_Y
'''    cfxRst(3).ConstantLine(0).Value = strMean
'''    cfxRst(3).ConstantLine(0).LineStyle = CHART_DOT
'''    cfxRst(3).ConstantLine(0).Style = CC_COLORTEXT
'''    cfxRst(3).ConstantLine(0).LineWidth = 1
'''    cfxRst(3).ConstantLine(0).LineColor = &H808080
'''
'''    cfxRst(3).CloseData COD_VALUES
'''    cfxRst(3).CloseData COD_CONSTANTS
'''
'''    Set objData = Nothing
End Sub

Private Sub SetRefRange(ByVal lngGrp As Long, ByVal lngIdx As Long)
'''    Dim iMinValue As Double
'''    Dim iMaxValue As Double
'''    Dim N_1SdVal As Double, N_2SdVal As Double, N_3SdVal As Double, N_4SdVal As Double
'''    Dim P_1SdVal As Double, P_2SdVal As Double, P_3SdVal As Double, p_4SdVal As Double
'''
'''    With objQcReview
'''
'''        If lngIdx = 0 Then Exit Sub
'''
'''        cfxRst(lngGrp).Tag = CStr(lngIdx)    'LotTableÀÇ ListIndex
'''
'''        lblLotNo(lngGrp).Caption = .Item(lngIdx).Lotno
'''        lblOpenDt(lngGrp).Caption = .Item(lngIdx).OpenDt
'''        lblExpDt(lngGrp).Caption = .Item(lngIdx).ExpDt
'''
'''        iMinValue = .Item(lngIdx).MeanVal - (4 * .Item(lngIdx).SdVal)
'''        iMaxValue = .Item(lngIdx).MeanVal + (4 * .Item(lngIdx).SdVal)
'''        cfxRst(lngGrp).Adm(CSA_MIN) = iMinValue
'''        cfxRst(lngGrp).Adm(CSA_MAX) = iMaxValue
'''        cfxRst(lngGrp).Adm(CSA_GAP) = .Item(lngIdx).SdVal
'''        cfxRst(lngGrp).Axis(AXIS_Y).STEP = (iMaxValue - iMinValue) / 4
'''
'''        cfxRst(lngGrp).OpenDataEx COD_STRIPES, 8, 0
'''
'''        N_1SdVal = .Item(lngIdx).MeanVal - .Item(lngIdx).SdVal
'''        N_2SdVal = .Item(lngIdx).MeanVal - (.Item(lngIdx).SdVal * 2)
'''        N_3SdVal = .Item(lngIdx).MeanVal - (.Item(lngIdx).SdVal * 3)
'''        N_4SdVal = .Item(lngIdx).MeanVal - (.Item(lngIdx).SdVal * 4)
'''        P_1SdVal = .Item(lngIdx).MeanVal + .Item(lngIdx).SdVal
'''        P_2SdVal = .Item(lngIdx).MeanVal + (.Item(lngIdx).SdVal * 2)
'''        P_3SdVal = .Item(lngIdx).MeanVal + (.Item(lngIdx).SdVal * 3)
'''        p_4SdVal = .Item(lngIdx).MeanVal + (.Item(lngIdx).SdVal * 4)
'''        Call DrawStripe(lngGrp, 0, N_4SdVal, N_3SdVal)
'''        Call DrawStripe(lngGrp, 1, N_3SdVal, N_2SdVal)
'''        Call DrawStripe(lngGrp, 2, N_2SdVal, N_1SdVal)
'''        Call DrawStripe(lngGrp, 3, N_1SdVal, .Item(lngIdx).MeanVal)
'''        Call DrawStripe(lngGrp, 4, .Item(lngIdx).MeanVal, P_1SdVal)
'''        Call DrawStripe(lngGrp, 5, P_1SdVal, P_2SdVal)
'''        Call DrawStripe(lngGrp, 6, P_2SdVal, P_3SdVal)
'''        Call DrawStripe(lngGrp, 7, P_3SdVal, p_4SdVal)
'''
'''        cfxRst(lngGrp).CloseData COD_STRIPES
'''
'''        ' »ó¼ö ¶óÀÎ
'''        cfxRst(lngGrp).OpenDataEx COD_CONSTANTS, 9, 0
'''
'''        Call DrawLine(lngGrp, 0, p_4SdVal)
'''        Call DrawLine(lngGrp, 1, P_3SdVal)
'''        Call DrawLine(lngGrp, 2, P_2SdVal)
'''        Call DrawLine(lngGrp, 3, P_1SdVal)
'''        Call DrawLine(lngGrp, 4, .Item(lngIdx).MeanVal, 0)
'''        Call DrawLine(lngGrp, 5, N_1SdVal)
'''        Call DrawLine(lngGrp, 6, N_2SdVal)
'''        Call DrawLine(lngGrp, 7, N_3SdVal)
'''        Call DrawLine(lngGrp, 8, N_4SdVal)
'''
'''        cfxRst(lngGrp).CloseData COD_CONSTANTS
'''
'''        cfxRst(lngGrp).LegStyle = 2
'''        cfxRst(lngGrp).FixedGap = 5
'''        cfxRst(lngGrp).BottomGap = 33
'''        cfxRst(lngGrp).LeftGap = 60
'''        cfxRst(lngGrp).RightGap = 10
'''        cfxRst(lngGrp).TopGap = 5
'''        cfxRst(lngGrp).Title(CHART_LEFTTIT) = .Item(lngIdx).LevelNm
'''        cfxRst(lngGrp).Scrollable = True
'''
'''    End With
End Sub

Private Sub DrawStripe(ByVal lngGrp As Long, ByVal iCnt As Integer, ByVal iFromVal As Double, ByVal iToVal As Double)

'''    With cfxRst(lngGrp).Stripe(iCnt)
'''        .Axis = AXIS_Y
'''        .Color = Choose(iCnt + 1, &HF7F0F0, &HFFF1DF, &HDBF2FD, &HC0FFFF, &HC0FFFF, &HDBF2FD, &HFFF1DF, &HF7F0F0)
'''        .From = iFromVal
'''        .To = iToVal
'''    End With
End Sub

Private Sub DrawLine(ByVal lngGrp As Long, ByVal iCnt As Integer, ByVal iValue As Double, Optional ByVal iStyle As Integer = CHART_DOT)

    Dim aryLineStyle()
    aryLineStyle = Array(0, 4, 1, 2, 3, 2, 1, 4, 0)
    With cfxRst(lngGrp).ConstantLine(iCnt)
        .Value = iValue
        '.LineColor = &H808080  '&H80&
        .Axis = AXIS_Y
        .Label = Choose(iCnt + 1, "", "3SD(+)", "2SD(+)", "1SD(+)", "Mean", "1SD(-)", "2SD(-)", "3SD(-)", "")
        .LineWidth = 1
        .LineStyle = aryLineStyle(iCnt)
        .Style = CC_COLORTEXT
        '.LineStyle = iStyle
        Select Case .LineStyle
           Case 0:
              .LineColor = &H0&
           Case 3:
              .LineColor = &H80&
           Case Else
              .LineColor = &H808080
        End Select
    End With
        
End Sub

'Private Sub mnuPrint_Click()
'    If tblResult.DataRowCnt = 0 Then Exit Sub
''    Call PrintTable
'
'    cmdPrint.Tag = "mnuPrint"
'
'    Call PrintCrystalAllItem
'End Sub

Private Sub objPop_Click(ByVal vMenuID As Long)
'''    Select Case vMenuID
'''        Case MENU_ONE
'''        'µû·Îµû·Î º¸¿©ÁÖ´Â ±×·¡ÇÁ
'''            Dim lngCnt As Long
'''            Dim i As Long
'''
'''            For i = 1 To lvwLotNo.ListItems.Count
'''                If lvwLotNo.ListItems(i).Checked Then
'''                    lngCnt = lngCnt + 1
'''                End If
'''            Next
'''
'''            If lngCnt = 0 Then Exit Sub
'''
'''            Call InitGraph
'''
'''            lngCnt = 0
'''            For i = 1 To lvwLotNo.ListItems.Count
'''                If lvwLotNo.ListItems(i).Checked Then
'''                    Call ShowGraph(lngCnt, objQcReview.Item(i).KeyString)
'''                    Call SetRefRange(lngCnt, i)
'''                    lngCnt = lngCnt + 1
'''                    If lngCnt >= 3 Then Exit For
'''                End If
'''            Next
'''
'''            fraGraph.Visible = True
'''            fraGraph.ZOrder 0
'''        Case MENU_ALL
''''            Dim lngCnt As Long
''''            Dim i As Long
'''
'''            For i = 1 To lvwLotNo.ListItems.Count
'''                If lvwLotNo.ListItems(i).Checked Then
'''                    lngCnt = lngCnt + 1
'''                End If
'''            Next
'''
'''            If lngCnt = 0 Then Exit Sub
'''
'''            Call InitGraph
'''            Call ShowGraphAll
'''        '    Call ShowAdvancedGraph
'''
'''            fraGrpAll.Visible = True
'''            fraGrpAll.ZOrder 0
'''        Case MENU_PRT
'''            If tblResult.DataRowCnt = 0 Then Exit Sub
'''        '    Call PrintTable
'''
'''            cmdPrint.Tag = "mnuPrint"
'''
'''            Call PrintCrystalAllItem
'''    End Select
End Sub

Private Sub optAR_Click(Index As Integer)
    On Error Resume Next
'''    If Screen.ActiveControl.Name <> optAR(Index).Name Then Exit Sub
'''
'''    If lstTestItem.ListCount <> 0 Then
'''        lstTestItem.ListIndex = 0
'''        Call LoadData
'''    End If
End Sub

Private Sub optLevel_Click(Index As Integer)
'''    On Error Resume Next
'''    If Screen.ActiveControl.Name <> optLevel(Index).Name Then Exit Sub
'''
'''    If lstTestItem.ListCount <> 0 Then
'''        lstTestItem.ListIndex = 0
'''        Call LoadData
'''    End If
End Sub

Private Sub tblResult_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
'''    If tblResult.DataRowCnt = 0 Then Exit Sub
'''    Set objPop = New clsPopupMenu
'''    With objPop
'''        .AddMenu MENU_PRT, "PRINT"
'''        .PopupMenus Me.hwnd
'''    End With
'''    Set objPop = Nothing
''''    Set mnuPopup = frmControls.mnuPopup
''''    Set mnuPrint = frmControls.mnuSub1
''''    Set mnuGraphAll = frmControls.mnuSub2
''''
''''    mnuPrint.Caption = "Print"
''''    mnuGraphAll.Visible = False
''''
''''    Me.PopupMenu mnuPopup
''''
''''    mnuGraphAll.Visible = True
''''    Set mnuPrint = Nothing
''''    Set mnuGraphAll = Nothing
''''    Set mnuPopup = Nothing
End Sub
Private Sub tblResult_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
'''
'''    txtMeanL.Text = ""
'''    txtMeanN.Text = ""
'''    txtMeanH.Text = ""
'''    txtSDL.Text = ""
'''    txtSDN.Text = ""
'''    txtSDH.Text = ""
'''    txtCVL.Text = ""
'''    txtCVN.Text = ""
'''    txtCVH.Text = ""
'''    Cancel = True
End Sub

Private Sub tblResult_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
'''    Dim varRst As Variant
'''    Dim strPreRst As String
'''    Dim strPostRst As String
'''    Dim varRow As Variant
'''    Dim varRet  As Boolean
'''
'''
'''    If tblResult.DataRowCnt = 0 Then Exit Sub
'''    If Row = 0 Then Exit Sub
'''
'''    If Col >= 1 And Col <= 6 Then Call tblResult.GetText(2, Row, varRow)
'''    If Col >= 7 And Col <= 11 Then Call tblResult.GetText(7, Row, varRow)
'''    If Col >= 12 And Col <= 16 Then Call tblResult.GetText(12, Row, varRow)
'''
''''    If varRow = "" Then Exit Sub
'''    varRet = True
'''    Call GetRecentRst(Row, varRow, strPreRst, strPostRst, varRet)
'''    If Not varRet Then Exit Sub
'''    MultiLine = 1
'''
'''    If strPreRst <> "" Then
'''        TipText = vbNewLine & strPreRst
'''    Else
'''        TipText = ""
'''    End If
'''
'''    If strPostRst <> "" Then
'''        If strPreRst = "" Then
'''            TipText = vbNewLine & strPostRst & vbNewLine
'''        Else
'''            TipText = TipText & vbNewLine & vbNewLine & strPostRst & vbNewLine
'''        End If
'''    Else
'''        TipText = TipText & vbNewLine
'''    End If
'''
'''    TipWidth = 3500
''''    tblResult.TextTipDelay = 1000
'''    Call tblResult.SetTextTipAppearance("±¼¸²Ã¼", 9, False, False, &HEEFDF2, &H996666)
'''    ShowTip = True
End Sub

Private Sub GetRecentRst(ByVal pRow As Long, ByVal sDate As String, ByRef pPreRst As String, ByRef pPostRst As String, ByRef ShowTip As Boolean)
'''    Dim Rs As Recordset
'''    Dim strSQL As String
'''    Dim strWorkArea As String
'''    Dim strAccDt As String
'''    Dim strAccSeq As String
''''---------------------------
'''' 2009.02.19 ¾ç¼ºÇö Ãß°¡
'''    Dim strMSDCV As String
'''    Dim strMean  As String
'''    Dim strSD    As String
'''    Dim strCV    As String
'''
''''---------------------------
'''    Dim strPreWA    As String  'ÀÌÀü°á°ú
'''    Dim strPreAD    As String
'''    Dim strPreAS    As String
'''    Dim strPreRst   As String
'''    Dim strPreVfyNm As String
'''    Dim strPreRA    As String
'''
'''    Dim strCurWA    As String  'ÇöÀç°á°ú
'''    Dim strCurAD    As String
'''    Dim strCurAS    As String
'''    Dim strCurRst   As String
'''    Dim strCurVfyNm As String
'''    Dim strCurRA    As String
'''
'''    Dim strPostWA       As String 'ÀÌÈÄ°á°ú
'''    Dim strPostAD       As String
'''    Dim strPostAS       As String
'''    Dim strPostRst      As String
'''    Dim strPostVfyNm    As String
'''    Dim strPostRA       As String
'''
'''
''''---------------------------
'''' 2009.10.06 ¾ç¼ºÇö Ãß°¡
'''    Dim dtpFrdate    As String
'''    Dim dtpToDate    As String
'''    Dim dtpModate    As String
'''
'''    Dim strCuH(5)    As String
'''    Dim strCuN(5)    As String
'''    Dim strCuL(5)    As String
'''    Dim strPrH(5)    As String
'''    Dim strPrN(5)    As String
'''    Dim strPrL(5)    As String
'''    Dim strPoH(5)    As String
'''    Dim strPoN(5)    As String
'''    Dim strPoL(5)    As String
'''    Dim IntTemp     As Integer
''''---------------------------
'''    '** °á°úÈ®ÀÎÀÚ Ãß°¡ By M.G.Choi 2006.09.01
'''
'''    strWorkArea = objQcReview.TestItem(pRow).WorkArea
'''    strAccDt = objQcReview.TestItem(pRow).AccDt
'''    strAccSeq = objQcReview.TestItem(pRow).AccSeq
'''
'''     dtpFrdate = Format(dtpFdate.Value, CS_DateDbFormat)
'''     dtpToDate = Format(dtpTdate.Value, CS_DateDbFormat)
'''     dtpModate = Format(Now, "YYYY")
'''     dtpModate = dtpModate & medGetP(sDate, 1, "/")
'''
'''    '** º¯°æ By M.G.Choi 2006.09.05
'''    strSQL = " select workarea, accdt, accseq, rstcd, vfyid, radiv " & _
'''             "   from " & T_LAB026 & _
'''             " where " & DBW("ctrlcd=", objQcReview.CtrlCd) & _
'''             " and " & DBW("levelcd=", objQcReview.TestItem(pRow).LevelCd) & _
'''             " and " & DBW("lotno=", objQcReview.TestItem(pRow).Lotno) & _
'''             " and " & DBW("testcd=", objQcReview.TestItem(pRow).TestCd) & _
'''             " order by workarea,accdt,accseq "
'''
'''    '** ¿øº» =================================================================
''''    strSQL = " select * " & _
''''             "   from " & T_LAB026 & _
''''             " where " & DBW("ctrlcd=", objQcReview.CtrlCd) & _
''''             " and " & DBW("levelcd=", objQcReview.TestItem(pRow).LevelCd) & _
''''             " and " & DBW("lotno=", objQcReview.TestItem(pRow).Lotno) & _
''''             " and " & DBW("testcd=", objQcReview.TestItem(pRow).TestCd) & _
''''             " order by workarea,accdt,accseq "
'''    '==========================================================================
'''
''''---------------------------
'''' 2009.10.06 ¾ç¼ºÇö ¼öÁ¤
'''
'''    strSQL = " select a.*,b.MeanVal, b.SDVal, b.CVVal from " & T_LAB026 & " a, " & T_LAB024 & " b " & _
'''             " where " & DBW("a.ctrlcd=", objQcReview.CtrlCd) & _
'''             " and a.accdt >= (select max(c.accdt) from " & T_LAB026 & " c where  " & DBW("c.ctrlcd=", objQcReview.CtrlCd) & " and c.accdt < '" & dtpModate & "' and " & DBW("c.testcd=", objQcReview.TestItem(pRow).TestCd) & ")" & _
'''             " and a.accdt <= (select max(c.accdt) from " & T_LAB026 & " c where  " & DBW("c.ctrlcd=", objQcReview.CtrlCd) & " and c.accdt > '" & dtpModate & "' and " & DBW("c.testcd=", objQcReview.TestItem(pRow).TestCd) & " and rownum < 2 ) " & _
'''             " and " & DBW("a.testcd=", objQcReview.TestItem(pRow).TestCd) & _
'''             " and a.ctrlcd=b.ctrlcd and a.levelcd=b.levelcd and a.lotno=b.lotno and a.testcd=b.testcd" & _
'''             " order by a.workarea, a.accdt, a.accseq "
''''---------------------------
'''
'''    Set Rs = New Recordset
'''    Rs.Open strSQL, DBConn
'''    If Rs.EOF Then
'''        ShowTip = False
'''        Set Rs = Nothing
'''        Exit Sub
'''    Else
'''        ShowTip = True
'''    End If
'''
'''    Do Until Rs.EOF
'''        strPreRst = Rs.Fields("workarea").Value & "" & Rs.Fields("accdt").Value & ""
'''        For IntTemp = 1 To 3
'''            If Rs.EOF Then Exit For
'''            If strPreRst <> Rs.Fields("workarea").Value & "" & Rs.Fields("accdt").Value & "" Then Exit For
'''            strPreWA = Rs.Fields("workarea").Value & ""
'''            strPreAD = Rs.Fields("accdt").Value & ""
'''
'''            strPreAS = Rs.Fields("accseq").Value & "" & "," & strPreAS
'''            strPreVfyNm = GetQCEmpName(strPreWA, strPreAD, Rs.Fields("accseq").Value & "")
'''            strPreRA = Rs.Fields("radiv").Value & ""
'''
'''            Select Case Rs.Fields("levelcd").Value & ""
'''                Case "H": strPrH(0) = Rs.Fields("rstcd").Value & "": strPrH(1) = Rs.Fields("meanval").Value & "": strPrH(2) = Rs.Fields("sdval").Value & "": strPrH(3) = Rs.Fields("cvval").Value & "": strPrH(4) = Rs.Fields("accseq").Value & ""
'''                Case "N": strPrN(0) = Rs.Fields("rstcd").Value & "": strPrN(1) = Rs.Fields("meanval").Value & "": strPrN(2) = Rs.Fields("sdval").Value & "": strPrN(3) = Rs.Fields("cvval").Value & "": strPrN(4) = Rs.Fields("accseq").Value & ""
'''                Case "L": strPrL(0) = Rs.Fields("rstcd").Value & "": strPrL(1) = Rs.Fields("meanval").Value & "": strPrL(2) = Rs.Fields("sdval").Value & "": strPrL(3) = Rs.Fields("cvval").Value & "": strPrL(4) = Rs.Fields("accseq").Value & ""
'''            End Select
'''            Rs.MoveNext
'''        Next IntTemp
'''        If Rs.EOF Then Exit Do
'''        strPreRst = Rs.Fields("workarea").Value & "" & Rs.Fields("accdt").Value & ""
'''        For IntTemp = 1 To 3
'''            If Rs.EOF Then Exit For
'''            If strPreRst <> Rs.Fields("workarea").Value & "" & Rs.Fields("accdt").Value & "" Then Exit For
'''            strCurWA = Rs.Fields("workarea").Value & ""
'''            strCurAD = Rs.Fields("accdt").Value & ""
'''            strCurAS = Rs.Fields("accseq").Value & "" & "," & strCurAS
'''            strCurRst = Rs.Fields("rstcd").Value & ""
'''            strCurVfyNm = GetQCEmpName(strCurWA, strCurAD, Rs.Fields("accseq").Value & "")
'''            strCurRA = Rs.Fields("radiv").Value & ""
'''            Select Case Rs.Fields("levelcd").Value & ""
'''                Case "H": strCuH(0) = Rs.Fields("rstcd").Value & "": strCuH(1) = Rs.Fields("meanval").Value & "": strCuH(2) = Rs.Fields("sdval").Value & "": strCuH(3) = Rs.Fields("cvval").Value & "": strCuH(4) = Rs.Fields("accseq").Value & ""
'''                Case "N": strCuN(0) = Rs.Fields("rstcd").Value & "": strCuN(1) = Rs.Fields("meanval").Value & "": strCuN(2) = Rs.Fields("sdval").Value & "": strCuN(3) = Rs.Fields("cvval").Value & "": strCuN(4) = Rs.Fields("accseq").Value & ""
'''                Case "L": strCuL(0) = Rs.Fields("rstcd").Value & "": strCuL(1) = Rs.Fields("meanval").Value & "": strCuL(2) = Rs.Fields("sdval").Value & "": strCuL(3) = Rs.Fields("cvval").Value & "": strCuL(4) = Rs.Fields("accseq").Value & ""
'''            End Select
'''            Rs.MoveNext
'''        Next IntTemp
'''        If Rs.EOF Then Exit Do
'''        strPreRst = Rs.Fields("workarea").Value & "" & Rs.Fields("accdt").Value & ""
'''        For IntTemp = 1 To 3
'''            If strPreRst <> Rs.Fields("workarea").Value & "" & Rs.Fields("accdt").Value & "" Then Exit For
'''            strPostWA = Rs.Fields("workarea").Value & ""
'''            strPostAD = Rs.Fields("accdt").Value & ""
'''            strPostAS = Rs.Fields("accseq").Value & "" & "," & strPostAS
'''            strPostRst = Rs.Fields("rstcd").Value & ""
'''            strPostVfyNm = GetQCEmpName(strPostWA, strPostAD, Rs.Fields("accseq").Value & "")
'''            strPostRA = Rs.Fields("radiv").Value & ""
'''            Select Case Rs.Fields("levelcd").Value & ""
'''                Case "H": strPoH(0) = Rs.Fields("rstcd").Value & "": strPoH(1) = Rs.Fields("meanval").Value & "": strPoH(2) = Rs.Fields("sdval").Value & "": strPoH(3) = Rs.Fields("cvval").Value & "": strPoH(4) = Rs.Fields("accseq").Value & ""
'''                Case "N": strPoN(0) = Rs.Fields("rstcd").Value & "": strPoN(1) = Rs.Fields("meanval").Value & "": strPoN(2) = Rs.Fields("sdval").Value & "": strPoN(3) = Rs.Fields("cvval").Value & "": strPoN(4) = Rs.Fields("accseq").Value & ""
'''                Case "L": strPoL(0) = Rs.Fields("rstcd").Value & "": strPoL(1) = Rs.Fields("meanval").Value & "": strPoL(2) = Rs.Fields("sdval").Value & "": strPoL(3) = Rs.Fields("cvval").Value & "": strPoL(4) = Rs.Fields("accseq").Value & ""
'''            End Select
'''            Rs.MoveNext
'''            If Rs.EOF Then Exit For
'''        Next IntTemp
'''
'''    Loop
'''
''''---------------------------
'''' 2009.02.19 ¾ç¼ºÇö Ãß°¡
'''
''''        strMSDCV = " - ÇöÀç °á°ú - " & vbNewLine & _
'''
''''                    "   Mean Value : " & strMean & vbNewLine & _
''''                    "   SD   Value : " & strSD & vbNewLine & _
''''                    "   CV   Value : " & strCV & vbNewLine & _
'''
'''        strMSDCV = "   Á¢¼ö ¹ø È£ : " & strWorkArea & "-" & Mid(strAccDt, 3) & " " & strCuL(4) & "," & strCuN(4) & "," & strCuH(4) & vbNewLine
'''
''''---------------------------
'''
'''    If strPreWA <> "" Then
'''        pPreRst = strMSDCV & vbNewLine & _
'''                  " - ÀÌÀü °á °ú - " & vbNewLine & _
'''                  "   Á¢¼ö ¹ø È£ : " & strPreWA & "-" & Mid(strPreAD, 3) & " " & strPrL(4) & "," & strPrN(4) & "," & strPrH(4) & vbNewLine & _
'''                  "   °á      °ú : " & strPrL(0) & " / " & strPrN(0) & " / " & strPrH(0) & vbNewLine & _
'''                  "   ÀûÇÕ ¿© ºÎ : " & IIf(strPreRA = "A", "Accept", "Reject") & vbNewLine & _
'''                  "   È®  ÀÎ  ÀÚ : " & strPreVfyNm
'''    End If
'''
'''    If strPostWA <> "" Then
'''        pPostRst = " - ÀÌÈÄ °á°ú - " & vbNewLine & _
'''                   "   Á¢¼ö ¹ø È£ : " & strPostWA & "-" & Mid(strPostAD, 3) & " " & strPoL(4) & "," & strPoN(4) & "," & strPoH(4) & vbNewLine & _
'''                   "   °á      °ú : " & strPoL(0) & " / " & strPoN(0) & " / " & strPoH(0) & vbNewLine & _
'''                   "   ÀûÇÕ ¿© ºÎ : " & IIf(strPostRA = "A", "Accept", "Reject") & vbNewLine & _
'''                   "   È®  ÀÎ  ÀÚ : " & strPostVfyNm
'''    End If
'''
'''    If strPreWA = "" Then pPostRst = strMSDCV & vbNewLine & pPostRst
'''    If strPreWA = "" And strPostWA = "" Then pPreRst = strMSDCV
'''
'''    Set Rs = Nothing
End Sub

Private Sub txtCtrlCd_Change()
''    On Error Resume Next
''    If Screen.ActiveControl.Name <> txtCtrlCd.Name Then Exit Sub
''    If lblCtrlNm.Caption = "" Then Exit Sub
''
''    lblCtrlNm.Caption = ""
''
''    Call InitForm
''    Call InitGraph
End Sub

Private Sub txtCtrlCd_GotFocus()
'''    SendKeys "{Home}+{End}"
End Sub

Private Sub txtCtrlCd_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(txtCtrlCd.Text) = "" Then Exit Sub
    
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtCtrlCd_LostFocus()
'    If Trim(txtCtrlCd.Text) = "" Then Exit Sub
'    ActControl = txtCtrlCd.Name
'    Call GetControl
'    On Error Resume Next
'    If lstControl.Visible Then lstControl.SetFocus
    Dim Rs As Recordset
    
    If Trim(txtCtrlCd.Text) = "" Then Exit Sub
    If Trim(lblCtrlNm.Caption) <> "" Then Exit Sub
    
    DoEvents
    Set Rs = GetControlInfo(Trim(txtCtrlCd.Text))
    
    If Rs.EOF = False Then
        DoEvents
        Call LoadControlInfo(Trim(txtCtrlCd.Text))
        DoEvents
        Call LoadTestItem
        DoEvents
        Call LoadData
    End If
    
    Set Rs = Nothing
End Sub

Private Sub UpDown1_DownClick(Index As Integer)
    cfxRst(Index).FixedGap = cfxRst(Index).FixedGap - 1
End Sub

Private Sub UpDown1_UpClick(Index As Integer)
    cfxRst(Index).FixedGap = cfxRst(Index).FixedGap + 1
End Sub

Private Function GetQCEmpName(ByVal pWorkArea As String, ByVal pAccDt As String, _
                                ByVal pAccSeq As String) As String
    Dim Rs          As New ADODB.Recordset
    Dim strSql      As String
    
    strSql = " select b.empnm from " & T_LAB201 & " a, " & T_COM006 & " b " & _
             "  where a.workarea = " & DBS(pWorkArea) & _
             "    and a.accdt = " & DBS(pAccDt) & _
             "    and a.accseq = " & DBS(pAccSeq) & _
             "    and a.vfyid = b.empid "
             
    Rs.Open strSql, DBConn, adOpenForwardOnly, adLockReadOnly
    
    If Rs.EOF = False Then
        GetQCEmpName = Rs.Fields("empnm").Value & ""
    End If
    
    Rs.Close
    Set Rs = Nothing
    
End Function
