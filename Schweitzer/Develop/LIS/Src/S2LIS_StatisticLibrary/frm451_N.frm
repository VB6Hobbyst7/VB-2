VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{9167B9A7-D5FA-11D2-86CA-00104BD5476F}#5.0#0"; "DRctl1.ocx"
Begin VB.Form frm451_N 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   0  '����
   Caption         =   "��(��)�� �˻��׸� ���"
   ClientHeight    =   8970
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15090
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8970
   ScaleWidth      =   15090
   ShowInTaskbar   =   0   'False
   WindowState     =   2  '�ִ�ȭ
   Begin DRcontrol1.DrFrame fraPtCnt 
      Height          =   5505
      Left            =   7350
      TabIndex        =   29
      Top             =   1785
      Visible         =   0   'False
      Width           =   7125
      _ExtentX        =   12568
      _ExtentY        =   9710
      Title           =   "ȯ�ڰǼ�(�ܷ�/����)"
      TitlePos        =   1
      BackColor       =   14411494
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmdCExcel 
         BackColor       =   &H00F4F0F2&
         Caption         =   "Excel"
         Height          =   510
         Left            =   2895
         Style           =   1  '�׷���
         TabIndex        =   43
         Tag             =   "158"
         Top             =   4860
         Width           =   1320
      End
      Begin FPSpread.vaSpread tblCnt 
         Height          =   3660
         Left            =   165
         TabIndex        =   30
         Top             =   1095
         Width           =   3300
         _Version        =   196608
         _ExtentX        =   5821
         _ExtentY        =   6456
         _StockProps     =   64
         BackColorStyle  =   1
         ButtonDrawMode  =   1
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   14411494
         MaxCols         =   3
         MaxRows         =   20
         ScrollBars      =   2
         ShadowColor     =   14737632
         ShadowDark      =   14737632
         SpreadDesigner  =   "frm451_N.frx":0000
         UserResize      =   1
         TextTip         =   4
      End
      Begin VB.CommandButton cmdRefresh 
         BackColor       =   &H00F4F0F2&
         Caption         =   "Refresh(&R)"
         Height          =   510
         Left            =   4215
         Style           =   1  '�׷���
         TabIndex        =   32
         Tag             =   "158"
         Top             =   4860
         Width           =   1320
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00F4F0F2&
         Caption         =   "�ݱ�"
         Height          =   510
         Left            =   5535
         Style           =   1  '�׷���
         TabIndex        =   31
         Tag             =   "158"
         Top             =   4860
         Width           =   1320
      End
      Begin MSComCtl2.DTPicker dtpStart 
         Height          =   315
         Left            =   750
         TabIndex        =   38
         Top             =   465
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   83689472
         CurrentDate     =   36238
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   315
         Left            =   4200
         TabIndex        =   39
         Top             =   465
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   83689472
         CurrentDate     =   36391
      End
      Begin FPSpread.vaSpread tblWCnt 
         Height          =   3660
         Left            =   3555
         TabIndex        =   42
         Top             =   1110
         Width           =   3300
         _Version        =   196608
         _ExtentX        =   5821
         _ExtentY        =   6456
         _StockProps     =   64
         BackColorStyle  =   1
         ButtonDrawMode  =   1
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   14411494
         MaxCols         =   3
         MaxRows         =   20
         ScrollBars      =   2
         ShadowColor     =   14737632
         ShadowDark      =   14737632
         SpreadDesigner  =   "frm451_N.frx":05D1
         UserResize      =   1
         TextTip         =   4
      End
      Begin VB.Label Label3 
         Alignment       =   2  '��� ����
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   41
         Top             =   510
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3555
         TabIndex        =   40
         Top             =   510
         Width           =   315
      End
   End
   Begin FPSpread.vaSpread tblData 
      Height          =   6315
      Left            =   75
      TabIndex        =   0
      Top             =   2160
      Width           =   14370
      _Version        =   196608
      _ExtentX        =   25347
      _ExtentY        =   11139
      _StockProps     =   64
      BackColorStyle  =   1
      ButtonDrawMode  =   1
      ColsFrozen      =   1
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   14411494
      MaxCols         =   14
      MaxRows         =   20
      ScrollBars      =   2
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      SpreadDesigner  =   "frm451_N.frx":0BA2
      TextTip         =   4
   End
   Begin FPSpread.vaSpread tblD 
      Height          =   6315
      Left            =   75
      TabIndex        =   15
      Top             =   2160
      Width           =   14370
      _Version        =   196608
      _ExtentX        =   25347
      _ExtentY        =   11139
      _StockProps     =   64
      BackColorStyle  =   1
      ButtonDrawMode  =   1
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   14411494
      MaxCols         =   33
      MaxRows         =   20
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      SpreadDesigner  =   "frm451_N.frx":1422
      UserResize      =   1
      TextTip         =   4
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   1470
      Left            =   1365
      TabIndex        =   23
      Top             =   8235
      Visible         =   0   'False
      Width           =   4605
      Begin VB.CommandButton cmdTime 
         Appearance      =   0  '���
         Caption         =   "�ð�����"
         Height          =   315
         Left            =   2175
         TabIndex        =   24
         Top             =   30
         Visible         =   0   'False
         Width           =   915
      End
      Begin MedControls1.LisLabel LisLabel1 
         Height          =   360
         Left            =   0
         TabIndex        =   25
         Top             =   0
         Visible         =   0   'False
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   635
         BackColor       =   8421504
         ForeColor       =   -2147483634
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   "�۾��ð�"
         Appearance      =   0
      End
      Begin FPSpread.vaSpread tblTime 
         Height          =   1230
         Left            =   30
         TabIndex        =   26
         Tag             =   "10114"
         Top             =   390
         Visible         =   0   'False
         Width           =   3075
         _Version        =   196608
         _ExtentX        =   5424
         _ExtentY        =   2170
         _StockProps     =   64
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
         BackColorStyle  =   1
         DisplayColHeaders=   0   'False
         DisplayRowHeaders=   0   'False
         EditEnterAction =   5
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FormulaSync     =   0   'False
         MaxCols         =   4
         MaxRows         =   5
         MoveActiveOnFocus=   0   'False
         ProcessTab      =   -1  'True
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         ShadowText      =   0
         SpreadDesigner  =   "frm451_N.frx":2155
         StartingColNumber=   2
         VirtualRows     =   24
         VisibleCols     =   2
         VisibleRows     =   3
      End
   End
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "����(&X)"
      Height          =   510
      Left            =   13140
      Style           =   1  '�׷���
      TabIndex        =   37
      Tag             =   "158"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "ȭ������(&C)"
      Height          =   510
      Left            =   11820
      Style           =   1  '�׷���
      TabIndex        =   36
      Tag             =   "158"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdExcel 
      BackColor       =   &H00F4F0F2&
      Caption         =   "Excel(&E)"
      Height          =   510
      Left            =   11895
      Style           =   1  '�׷���
      TabIndex        =   35
      Tag             =   "158"
      Top             =   1200
      Width           =   1320
   End
   Begin VB.CommandButton cmdQuery 
      BackColor       =   &H00F4F0F2&
      Caption         =   "��ȸ(&Q)"
      Height          =   510
      Left            =   11895
      Style           =   1  '�׷���
      TabIndex        =   34
      Tag             =   "158"
      Top             =   510
      Width           =   1320
   End
   Begin VB.CommandButton cmdPtCnt 
      BackColor       =   &H00F4F0F2&
      Caption         =   "ȯ�ڰǼ�(&N)"
      Height          =   1200
      Left            =   13215
      Style           =   1  '�׷���
      TabIndex        =   33
      Tag             =   "158"
      Top             =   510
      Width           =   1230
   End
   Begin FPSpread.vaSpread tblexcel 
      Height          =   675
      Left            =   270
      TabIndex        =   28
      Top             =   3210
      Visible         =   0   'False
      Width           =   675
      _Version        =   196608
      _ExtentX        =   1191
      _ExtentY        =   1191
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
      SpreadDesigner  =   "frm451_N.frx":2615
   End
   Begin VB.CommandButton cmdHelp 
      BackColor       =   &H00DBE6E6&
      Caption         =   "..."
      Height          =   375
      Left            =   5865
      TabIndex        =   2
      Top             =   915
      Width           =   435
   End
   Begin MedControls1.LisLabel lblNm 
      Height          =   390
      Left            =   6315
      TabIndex        =   1
      Top             =   915
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   688
      BackColor       =   14411494
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin DRcontrol1.DrText txtCd 
      Height          =   375
      Left            =   4710
      TabIndex        =   3
      Top             =   915
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "Text1"
      Appearance      =   1
      Alignment       =   2
      BorderColor     =   4210752
   End
   Begin MedControls1.LisLabel lblCondition 
      Height          =   360
      Left            =   2955
      TabIndex        =   4
      Top             =   915
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   635
      BackColor       =   10392451
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "�˻�����"
      Appearance      =   0
   End
   Begin MSComctlLib.TreeView tvwMenu 
      Height          =   1365
      Left            =   75
      TabIndex        =   5
      Top             =   420
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   2408
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin MedControls1.LisLabel LisLabel8 
      Height          =   360
      Left            =   75
      TabIndex        =   6
      Top             =   45
      Width           =   14355
      _ExtentX        =   25321
      _ExtentY        =   635
      BackColor       =   8388608
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "�� ȸ �� ��"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel lblCaption 
      Height          =   360
      Left            =   75
      TabIndex        =   10
      Top             =   1785
      Width           =   14370
      _ExtentX        =   25347
      _ExtentY        =   635
      BackColor       =   8388608
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   ""
      Appearance      =   0
   End
   Begin MSComDlg.CommonDialog DlgSave 
      Left            =   1110
      Top             =   270
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   555
      Left            =   2940
      TabIndex        =   7
      Top             =   345
      Width           =   1770
      Begin VB.OptionButton optC 
         BackColor       =   &H00DBE6E6&
         Caption         =   "�Ϻ�"
         Height          =   285
         Index           =   1
         Left            =   75
         TabIndex        =   8
         Top             =   165
         Value           =   -1  'True
         Width           =   795
      End
      Begin VB.OptionButton optC 
         BackColor       =   &H00DBE6E6&
         Caption         =   "����"
         Height          =   300
         Index           =   0
         Left            =   870
         TabIndex        =   9
         Top             =   165
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DBE6E6&
      Height          =   495
      Left            =   4890
      TabIndex        =   16
      Top             =   7740
      Visible         =   0   'False
      Width           =   2190
      Begin VB.OptionButton optTime 
         BackColor       =   &H00DBE6E6&
         Caption         =   "��ü�ð�"
         Height          =   225
         Index           =   0
         Left            =   90
         TabIndex        =   18
         Top             =   180
         Value           =   -1  'True
         Width           =   1020
      End
      Begin VB.OptionButton optTime 
         BackColor       =   &H00DBE6E6&
         Caption         =   "�۾��ð�"
         Height          =   225
         Index           =   1
         Left            =   1125
         TabIndex        =   17
         Top             =   180
         Width           =   1020
      End
   End
   Begin MedControls1.LisLabel LisLabel3 
      Height          =   345
      Left            =   2955
      TabIndex        =   27
      Top             =   1350
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   609
      BackColor       =   10392451
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "ȯ������"
      Appearance      =   0
   End
   Begin VB.Frame FraM 
      BackColor       =   &H00DBE6E6&
      Height          =   555
      Left            =   4710
      TabIndex        =   13
      Top             =   345
      Width           =   7170
      Begin MSComCtl2.DTPicker dtpFM 
         Height          =   375
         Left            =   30
         TabIndex        =   14
         Top             =   150
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "yyyy"
         Format          =   83689475
         CurrentDate     =   37063
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   345
         Index           =   1
         Left            =   1590
         TabIndex        =   45
         Top             =   150
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   609
         BackColor       =   10392451
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "��ȸ�⵵ ����"
         Appearance      =   0
      End
   End
   Begin VB.Frame FraD 
      BackColor       =   &H00DBE6E6&
      Height          =   555
      Left            =   4710
      TabIndex        =   11
      Top             =   345
      Width           =   7170
      Begin MSComCtl2.DTPicker dtpFrDt 
         Height          =   345
         Left            =   30
         TabIndex        =   12
         Top             =   165
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
         _Version        =   393216
         CustomFormat    =   "yyyy-MM"
         Format          =   83689475
         CurrentDate     =   37063
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   345
         Index           =   0
         Left            =   1605
         TabIndex        =   44
         Top             =   165
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   609
         BackColor       =   10392451
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "��ȸ������"
         Appearance      =   0
      End
   End
   Begin VB.Frame fraInOut 
      BackColor       =   &H00DBE6E6&
      Height          =   480
      Left            =   4710
      TabIndex        =   19
      Top             =   1245
      Width           =   7185
      Begin VB.OptionButton optInOut 
         BackColor       =   &H00DBE6E6&
         Caption         =   "��ü"
         Height          =   225
         Index           =   2
         Left            =   4005
         TabIndex        =   22
         Top             =   180
         Value           =   -1  'True
         Width           =   765
      End
      Begin VB.OptionButton optInOut 
         BackColor       =   &H00DBE6E6&
         Caption         =   "�ܷ�"
         Height          =   225
         Index           =   1
         Left            =   2040
         TabIndex        =   21
         Top             =   180
         Width           =   765
      End
      Begin VB.OptionButton optInOut 
         BackColor       =   &H00DBE6E6&
         Caption         =   "�Կ�"
         Height          =   225
         Index           =   0
         Left            =   150
         TabIndex        =   20
         Top             =   165
         Width           =   765
      End
   End
End
Attribute VB_Name = "frm451_N"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Mode        As String
Private ModeString  As String

Private WithEvents objCodeList  As clsPopUpList
Attribute objCodeList.VB_VarHelpID = -1
Private objProBar               As New jProgressBar.clsProgress

Private MySqlStmt               As New clsLISSqlStatement ' SQL Ŭ����
Private objWork                 As New clsDictionary
Private objDic                  As New clsDictionary      '�������
Private objDicT                 As New clsDictionary
Private objdicD                 As New clsDictionary      '�Ϻ����
Private objDicDT                As New clsDictionary

Private strTable  As String
Private strTable1 As String
Private strTable2 As String

Public Event LastFormUnload()


Private Sub cmdClear_Click()
    dtpFM.Value = GetSystemDate
    dtpFrDt.Value = GetSystemDate
    optC(0).Value = True
    optTime(0).Value = True
    Call DictionaryClear
End Sub

Private Sub cmdClose_Click()
    fraPtCnt.Visible = False
End Sub

Private Sub cmdExcel_Click()
    Dim strTmp      As String
    Dim objTable    As Object
    
    Set objTable = tblData
    If optC(1).Value Then Set objTable = tblD
    If objTable.DataRowCnt = 0 Then Exit Sub
    With objTable
        .Row = 0: .Row2 = .MaxRows
        .Col = 1: .COL2 = .MaxCols
        .BlockMode = True
        strTmp = .Clip
        .BlockMode = False
        tblexcel.MaxRows = .MaxRows + 1
        tblexcel.MaxCols = .MaxCols
        tblexcel.Row = 1: tblexcel.Row2 = tblexcel.MaxRows
        tblexcel.Col = 1: tblexcel.COL2 = tblexcel.MaxCols
        tblexcel.BlockMode = True
        tblexcel.Clip = strTmp
        tblexcel.BlockMode = False
    End With
    
    DlgSave.InitDir = "C:\"
    DlgSave.Filter = "ExCelFile(*.XLS)|*.XLS"
    DlgSave.FileName = "AccCount"
    DlgSave.ShowSave

    tblexcel.SaveTabFile (DlgSave.FileName)
    
End Sub

Private Sub cmdCExcel_Click()
    Dim strTmp  As String
    Dim lngRows As Long
    
    If tblCnt.DataRowCnt = 0 And tblWCnt.DataRowCnt = 0 Then Exit Sub
    
    With tblCnt
        .Row = 0: .Row2 = .MaxRows
        .Col = 1: .COL2 = .MaxCols
        .BlockMode = True
        strTmp = .Clip
        .BlockMode = False
        lngRows = .MaxRows
    End With
    With tblWCnt
        .Row = 0: .Row2 = .MaxRows
        .Col = 1: .COL2 = .MaxCols
        .BlockMode = True
        strTmp = strTmp & .Clip
        .BlockMode = False
        lngRows = lngRows + .MaxRows
    End With
    With tblexcel
        .MaxRows = lngRows + 1
        .MaxCols = 3
        .Row = 1: .Row2 = .MaxRows
        .Col = 1: .COL2 = 3
        .BlockMode = True
        .Clip = strTmp
        .BlockMode = False
    End With
    
    DlgSave.InitDir = "C:\"
    DlgSave.Filter = "ExCelFile(*.XLS)|*.XLS"
    DlgSave.FileName = "ȯ�ڰǼ�"
    DlgSave.ShowSave

    tblexcel.SaveTabFile (DlgSave.FileName)
End Sub
Private Sub cmdHelp_Click()
    Dim tmpSql  As String
    Dim lngTop  As Long
    Dim lngLeft As Long
    
    If Mode = "" Then
        MsgBox "��ȸ�׸��� �����Ͻ��� ��ȸ�ϼ���.", vbInformation + vbOKOnly, "��ȸ�׸���"
        Exit Sub
    End If
    
    Call DictionaryClear
    Set objCodeList = New clsPopUpList
    objCodeList.Connection = DBConn
    Select Case Mode
        Case "C1"   '�˻��׸�
            objCodeList.Tag = "TESTCD"
            objCodeList.FormCaption = "�˻��׸� ����Ʈ"
            objCodeList.ColumnHeaderText = "�˻��ڵ�;�˻��"
            tmpSql = MySqlStmt.SqlLAB001CodeList
        Case "C2"   '�����
            objCodeList.Tag = "DeptCd"
            objCodeList.FormCaption = "�μ� ����Ʈ"
            objCodeList.ColumnHeaderText = "�μ��ڵ�;�μ���"
        Case "C3"   'WorkArea'
            Dim objSQL As New clsLisSqlResult
            
            objCodeList.Tag = "WorkArea"
            objCodeList.FormCaption = "WorkArea ����Ʈ"
            objCodeList.ColumnHeaderText = "WorkArea;WorkArea Name"
            tmpSql = objSQL.GetWorkArea
            Set objSQL = Nothing
        Case "C4"
            Dim objEQL As New clsLisSqlResult
            
            objCodeList.Tag = "Eqpcd"
            objCodeList.FormCaption = "��� ����Ʈ"
            objCodeList.ColumnHeaderText = "����ڵ�;����"
            tmpSql = objEQL.GetEqpList("10")
            Set objEQL = Nothing
        Case "C5"
            objCodeList.Tag = "Doct"
            objCodeList.FormCaption = "ó���� ����Ʈ"
            objCodeList.ColumnHeaderText = "�ڵ�;����"
        Case "C6"
        
    End Select
       
'    Dim objData As clsBasisData
    
'    Set objData = New clsBasisData
    
    With objCodeList
        lngTop = txtCd.Top + txtCd.Height + Me.Top
        lngLeft = Me.Left + txtCd.Left
        Select Case Mode
            Case "C2": Call .LoadPopUp(GetSQLDeptList) ', lngTop, lngLeft) ', ObjLISComCode.DeptCd)
            Case "C5": Call .LoadPopUp(GetSQLDoctList) ', lngTop, lngLeft)
            Case Else: Call .LoadPopUp(tmpSql) ', lngTop, lngLeft)
        End Select
        
        txtCd.Text = Trim(medGetP(.SelectedString, 1, ";"))
        lblNm.Caption = Trim(medGetP(.SelectedString, 2, ";"))
    End With
'    Set objData = Nothing
End Sub

Private Sub QueryTestM(ByVal Div As String, ByVal FrDt As String, ByVal ToDT As String, ByVal Condition As String)
    Dim RS          As Recordset
    Dim strTmp      As String
    Dim strMainKey  As String
    Dim strSQL      As String
    Dim strKEY      As String
    
    Dim ii          As Long
    Dim jj          As Long
    Dim RowCnt      As Long
    Dim ColCnt      As Long
    Dim lngTm       As Long
    
    Select Case Div
        Case "C1"
            strTmp = "testnm"
            strSQL = MySqlStmt.GetAccCntSQL(FrDt, ToDT, , Condition)
        Case "C2"
            strTmp = "deptnm"
            strSQL = MySqlStmt.GetAccCntSQL(FrDt, ToDT, Condition)
        Case "C3"
            strTmp = "workarea"
            strSQL = MySqlStmt.GetAccCntSQL(FrDt, ToDT, , , , Condition)
        Case "C4"
            strTmp = "eqpcd"
            strSQL = MySqlStmt.GetAccCntSQL(FrDt, ToDT, , , Condition)
        Case "C5"
            strTmp = "orddoct"
            strSQL = MySqlStmt.GetAccCntSQL(FrDt, ToDT, , , , , , Condition)
    End Select
    
    Set RS = New Recordset
    RS.Open strSQL, DBConn
    
    Call ProGressSetting
    objProBar.Message = "�����Ǽ��� �����ϱ� �ֽ��ϴ�."
    
    If Not RS.EOF Then
        objProBar.Max = RS.RecordCount
        objWork.Sort = False: objDicT.Sort = False: objDic.Sort = False
        Do Until RS.EOF
            lngTm = Mid(RS.Fields("rcvtm").Value & "", 1, 4)
            objWork.MoveFirst
            Do Until objWork.EOF
                If lngTm >= Val(objWork.Fields("time1")) And lngTm <= Val(objWork.Fields("time2")) Then
                    strKEY = objWork.Fields("type")
                    Exit Do
                End If
                objWork.MoveNext
            Loop
            strKEY = strKEY & "��" & RS.Fields("testnm").Value & ""
            
            '�ð��� �����ؼ� ��Ƴ���
            If objDicT.Exists(RS.Fields(strTmp).Value & "" & COL_DIV & strKEY) Then
                objDicT.KeyChange RS.Fields(strTmp).Value & "" & COL_DIV & strKEY
                objDicT.Fields("C" & Mid(RS.Fields("rcvdt").Value & "", 5, 2)) = Val(objDicT.Fields("C" & Mid(RS.Fields("rcvdt").Value & "", 5, 2))) + Val(RS.Fields("cnt").Value & "")
            Else
                objDicT.AddNew Join(Array(RS.Fields(strTmp).Value & "", strKEY), COL_DIV), _
                              Join(Array("", "", "", "", "", "", "", "", "", "", "", ""), COL_DIV)
                
                objDicT.KeyChange RS.Fields(strTmp).Value & "" & COL_DIV & strKEY
                
                objDicT.Fields("C" & Mid(RS.Fields("rcvdt").Value & "", 5, 2)) = Val(objDicT.Fields("C" & Mid(RS.Fields("rcvdt").Value & "", 5, 2))) + Val(RS.Fields("cnt").Value & "")
            End If
            '�ð��� ������� ��Ƴ���
            strMainKey = ""
            If Mode = "C4" Then
                If RS.Fields(strTmp).Value & "" = "" Then
                    strMainKey = "���۾�"
                Else
                    strMainKey = RS.Fields(strTmp).Value & ""
                End If
            Else
                strMainKey = RS.Fields(strTmp).Value & ""
            End If
            If Mode <> "C1" Then strMainKey = strMainKey & "��" & RS.Fields("testnm").Value & ""
            
            If objDic.Exists(strMainKey) Then
                objDic.KeyChange strMainKey
                objDic.Fields("C" & Mid(RS.Fields("rcvdt").Value & "", 5, 2)) = Val(objDic.Fields("C" & Mid(RS.Fields("rcvdt").Value & "", 5, 2))) + Val(RS.Fields("cnt").Value & "")
            Else
                objDic.AddNew strMainKey, Join(Array("", "", "", "", "", "", "", "", "", "", "", ""), COL_DIV)
                
                objDic.KeyChange strMainKey
                objDic.Fields("C" & Mid(RS.Fields("rcvdt").Value & "", 5, 2)) = Val(objDic.Fields("C" & Mid(RS.Fields("rcvdt").Value & "", 5, 2))) + Val(RS.Fields("cnt").Value & "")
                
            End If
            
            ii = ii + 1
            objProBar.Value = ii
            objProBar.Message = "�����Ǽ��� �����ϱ� �ֽ��ϴ�.(" & ii & "��)"
            RS.MoveNext
        Loop
        
        objWork.Sort = True: objDic.Sort = True: objDicT.Sort = True
        
        If optTime(0).Value Then
            If optC(0).Value Then
                Call DisPlayData(objDic, tblData, 2, 13)
            Else
                Call DisPlayData(objdicD, tblD, 2, 32)
            End If
        Else
            If optC(0).Value Then
                Call DisPlayTime(objDicT, tblData, 2, 13)
            Else
                Call DisPlayTime(objDicDT, tblD, 2, 32)
            End If
        End If
    End If
    Set RS = Nothing
    Set objProBar = Nothing
End Sub
Private Sub QueryTestD(ByVal Div As String, ByVal FrDt As String, ByVal ToDT As String, ByVal Condition As String)
    Dim RS          As Recordset
    Dim strTmp      As String
    Dim strSQL      As String
    Dim strKEY      As String
    Dim strMainKey  As String
    Dim ii          As Long
    Dim jj          As Long
    Dim RowCnt      As Long
    Dim ColCnt      As Long
    Dim lngTm       As Long
   
    Select Case Div
        Case "C1"
            strTmp = "testnm"
            strSQL = MySqlStmt.GetAccCntSQL(FrDt, ToDT, , Condition)
        Case "C2"
            strTmp = "deptnm"
            strSQL = MySqlStmt.GetAccCntSQL(FrDt, ToDT, Condition)
        Case "C3"
            strTmp = "workarea"
            strSQL = MySqlStmt.GetAccCntSQL(FrDt, ToDT, , , , Condition)
        Case "C4"
            strTmp = "eqpcd"
            strSQL = MySqlStmt.GetAccCntSQL(FrDt, ToDT, , , Condition)
        Case "C5"
            strTmp = "orddoct"
            strSQL = MySqlStmt.GetAccCntSQL(FrDt, ToDT, , , , , , Condition)
    End Select
    
    Set RS = New Recordset
    RS.Open strSQL, DBConn
    
    If Not RS.EOF Then
        Call ProGressSetting
        objProBar.Max = RS.RecordCount
        objProBar.Message = "�����Ǽ��� �����ϱ� �ֽ��ϴ�..."
        Do Until RS.EOF
            lngTm = Mid(RS.Fields("rcvtm").Value & "", 1, 4)
            objWork.MoveFirst
            Do Until objWork.EOF
                If lngTm >= objWork.Fields("time1") And lngTm <= objWork.Fields("time2") Then
                    strKEY = objWork.Fields("type")
                    Exit Do
                End If
                objWork.MoveNext
            Loop
            strKEY = strKEY & "��" & RS.Fields("testnm").Value & ""
            
            If objDicDT.Exists(RS.Fields(strTmp).Value & "" & COL_DIV & strKEY) Then
                objDicDT.KeyChange RS.Fields(strTmp).Value & "" & COL_DIV & strKEY
                objDicDT.Fields("C" & Mid(RS.Fields("rcvdt").Value & "", 5, 2)) = Val(objDicDT.Fields("C" & Mid(RS.Fields("rcvdt").Value & "", 5, 2))) + Val(RS.Fields("cnt").Value & "")
            Else
            
                objDicDT.AddNew RS.Fields(strTmp).Value & "" & COL_DIV & strKEY, _
                                Join(Array("", "", "", "", "", "", "", "", "", "", "", "", _
                                "", "", "", "", "", "", "", "", "", "", "", "", _
                                "", "", "", "", "", "", "", "", "", "", "", "", ""), COL_DIV)
                
                objDicDT.KeyChange RS.Fields(strTmp).Value & "" & COL_DIV & strKEY
                objDicDT.Fields("C" & Mid(RS.Fields("rcvdt").Value & "", 7)) = Val(objDicDT.Fields("C" & Mid(RS.Fields("rcvdt").Value & "", 7))) + Val(RS.Fields("cnt").Value & "")
            End If
            
            strMainKey = ""
            If Mode = "C4" Then
                If RS.Fields(strTmp).Value & "" = "" Then strMainKey = "���۾�"
            Else
                strMainKey = RS.Fields(strTmp).Value & ""
            End If
            
            If Mode = "C4" Then
                If RS.Fields(strTmp).Value & "" = "" Then strMainKey = "���۾�"
            Else
                strMainKey = RS.Fields(strTmp).Value & ""
            End If
            If Mode <> "C1" Then strMainKey = strMainKey & "��" & RS.Fields("testnm").Value & ""
            
            If objdicD.Exists(strMainKey) Then
                objdicD.KeyChange strMainKey
                objdicD.Fields("C" & Mid(RS.Fields("rcvdt").Value & "", 5, 2)) = Val(objdicD.Fields("C" & Mid(RS.Fields("rcvdt").Value & "", 5, 2))) + Val(RS.Fields("cnt").Value & "")
            Else
                objdicD.AddNew strMainKey, Join(Array("", "", "", "", "", "", "", "", "", "", "", "", _
                                                                       "", "", "", "", "", "", "", "", "", "", "", "", _
                                                                       "", "", "", "", "", "", "", "", "", "", "", "", ""), COL_DIV)
                
                objdicD.KeyChange strMainKey
                objdicD.Fields("C" & Mid(RS.Fields("rcvdt").Value & "", 7)) = Val(objdicD.Fields("C" & Mid(RS.Fields("rcvdt").Value & "", 7))) + Val(RS.Fields("cnt").Value & "")
            End If
            ii = ii + 1
            objProBar.Value = ii
            RS.MoveNext
        Loop
        
        If optTime(0).Value Then
            If optC(0).Value Then
                Call DisPlayData(objDic, tblData, 2, 13)
            Else
                Call DisPlayData(objdicD, tblD, 2, 32)
            End If
        Else
            If optC(0).Value Then
                Call DisPlayTime(objDicT, tblData, 2, 13)
            Else
                Call DisPlayTime(objDicDT, tblD, 2, 32)
            End If
        End If
    End If
    Set RS = Nothing
    Set objProBar = Nothing
End Sub
Private Sub DisPlayData(ByVal objDic As clsDictionary, ByVal tblData As Object, ByVal MinCol As Long, _
                        ByVal MaxCol As Long)
                        
    Dim strKEY As String
    Dim RowAdd As String
    Dim RowCnt As Long
    Dim ColCnt As Long
    Dim ii     As Long
    Dim jj     As Long
        
    If objDic.GetString = "" Then Exit Sub
    Call ProGressSetting
    objProBar.Max = objDic.RecordCount
    objProBar.Message = "�����Ǽ��� DisPlay �ϱ� �ֽ��ϴ�..."

    With tblData
        .ReDraw = False
        objDic.MoveFirst
        Do Until objDic.EOF
            
            .Row = .DataRowCnt + 1
            If .Row > .MaxRows Then
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
            End If
            If Mode = "C3" Or Mode = "C5" Or Mode = "C4" Then
                If Mode <> "C1" Then

                    .Col = 1: .Value = GetConditionNm(medGetP(objDic.Fields("key"), 1, "��")): strKEY = objDic.Fields("key"):
                    .Value = IIf(.Value = "", "���۾�", .Value)
                    
                Else
                    .Col = 1: .Value = GetConditionNm(objDic.Fields("key")): strKEY = objDic.Fields("key"):
                End If
            Else
                If Mode <> "C1" Then
                    .Col = 1: .Value = medGetP(objDic.Fields("key"), 1, "��"): strKEY = .Value
                    
                Else
                    .Col = 1: .Value = objDic.Fields("key"): strKEY = .Value
                End If
                If Mode = "C4" Then .Value = IIf(.Value = "", "���۾�", .Value)
            End If
            .FontBold = True: .ForeColor = vbRed
            If Mode <> "C1" Then
                If RowAdd <> medGetP(objDic.Fields("key"), 1, "��") Then
                    .Row = .DataRowCnt + 1
                    If .Row > .MaxRows Then
                        .MaxRows = .MaxRows + 1
                        .Row = .MaxRows
                    End If
                End If
                RowAdd = medGetP(objDic.Fields("key"), 1, "��")
                .Col = 1: .Value = "   " & medGetP(objDic.Fields("key"), 2, "��")
                .FontBold = False: .ForeColor = vbBlue
            End If
            
            For ii = MinCol To MaxCol
                .Col = ii: .Value = objDic.Fields("C" & Format((ii - 1), "00"))
                .FontBold = False: .ForeColor = vbBlack
            Next
            ii = ii + 1
            objProBar.Value = ii
            objDic.MoveNext
        Loop
        Dim RowPos As Long
        
        For ii = 1 To .DataRowCnt
            .Row = ii
            RowCnt = 0
            For jj = MinCol To MaxCol
                .Col = jj
                RowCnt = RowCnt + Val(.Value)
            Next
            .Col = MaxCol + 1: .Value = IIf(RowCnt = 0, "", RowCnt)
        Next
        .Row = .DataRowCnt + 2
        If .Row > .MaxRows Then
            .MaxRows = .MaxRows + 2
            .Row = .MaxRows
        End If
        RowPos = .Row
        For ii = MinCol To MaxCol + 1
            .Col = ii
            ColCnt = 0
            For jj = 1 To .DataRowCnt
                .Row = jj
                ColCnt = ColCnt + Val(.Value)
            Next
            .Row = RowPos
            .Value = IIf(ColCnt = 0, "", ColCnt)
            .Col = 1
            .Value = "�հ�"
            .FontBold = True: .ForeColor = vbRed
        Next
        .ReDraw = False
    End With
    Set objProBar = Nothing
End Sub

Private Sub DisPlayTime(ByVal objDic As clsDictionary, ByVal tblData As Object, ByVal MinCol As Long, _
                             ByVal MaxCol As Long)
    Dim RowAdd      As String
    Dim RowADD1     As String
    Dim Timechk     As String
    Dim strKEY      As String
    
    Dim RowCnt      As Long
    Dim ColCnt      As Long
    Dim ii          As Long
    Dim jj          As Long

    RowAdd = "": RowADD1 = ""
    If objDic.GetString = "" Then Exit Sub
    Call ProGressSetting
    objProBar.Max = objDic.RecordCount
    objProBar.Message = "�����Ǽ��� Display �ϱ� �ֽ��ϴ�..."
    
    With tblData
        objDic.MoveFirst
        .ReDraw = True
        Do Until objDic.EOF
            If Mode = "C4" Then
                If objDic.Fields("key") = "" Then objDic.Fields("key") = "���۾�"
            End If
            If strKEY <> objDic.Fields("key") Then
                RowADD1 = ""
                .Row = .DataRowCnt + 1
                If .Row > .MaxRows Then
                    .MaxRows = .MaxRows + 1
                    .Row = .MaxRows
                End If
                
                If Mode = "C3" Or Mode = "C5" Or Mode = "C4" Then
                '��Ī ���ؿ���(workarea,���)
                    .Col = 1: .Value = GetConditionNm(objDic.Fields("key")):
                    strKEY = objDic.Fields("key")
                Else
                '�˻��׸� �Ǽ��� ������ ������ ��ȸ���� �˻��׸� �Ǽ��� ������
                    .Col = 1: .Value = objDic.Fields("key"):
                    strKEY = objDic.Fields("key")
                End If
                '���(""): ���۾�
                If Mode = "C4" Then .Value = IIf(.Value = "", "���۾�", .Value)
                .FontBold = True: .ForeColor = vbRed
                
                If Mode <> "C1" Then
                    If RowAdd <> medGetP(objDic.Fields("key1"), 2, "��") Then
                        .Row = .DataRowCnt + 1
                        If .Row > .MaxRows Then
                            .MaxRows = .MaxRows + 1
                            .Row = .MaxRows
                        End If
                    End If
                     .Col = 1: .Value = "   " & medGetP(objDic.Fields("key1"), 2, "��")
                     .FontBold = False: .ForeColor = vbBlue
                    If RowAdd <> medGetP(objDic.Fields("key1"), 1, "��") Then
                        .Row = .DataRowCnt + 1
                        If .Row > .MaxRows Then
                            .MaxRows = .MaxRows + 1
                            .Row = .MaxRows
                        End If
                    End If
                    .Col = 1: .Value = medGetP(objDic.Fields("key1"), 1, "��")
                Else
                    If RowADD1 <> medGetP(objDic.Fields("key1"), 1, "��") Then
                        .Row = .DataRowCnt + 1
                        If .Row > .MaxRows Then
                            .MaxRows = .MaxRows + 1
                            .Row = .MaxRows
                        End If
                    End If
                End If
                
                '�˻��׸� üũ
                If objWork.Exists(medGetP(objDic.Fields("key1"), 1, "��")) Then
                    objWork.KeyChange medGetP(objDic.Fields("key1"), 1, "��")
                End If
                
                .Col = 1: .Value = "               " & Format(objWork.Fields("time1"), "00:00") & " ~ " & Format(objWork.Fields("time2"), "00:00")
                .FontBold = False: .ForeColor = vbBlack
                RowAdd = medGetP(objDic.Fields("key1"), 2, "��")
                RowADD1 = medGetP(objDic.Fields("key1"), 2, "��")
            Else
                .Row = .DataRowCnt + 1
                If .Row > .MaxRows Then
                    .MaxRows = .MaxRows + 1
                    .Row = .MaxRows
                End If
                If Mode <> "C1" Then
                    .Col = 1: .Value = "   " & medGetP(objDic.Fields("key1"), 2, "��")
                    .FontBold = False: .ForeColor = vbBlue
                    
                    If RowAdd <> medGetP(objDic.Fields("key1"), 2, "��") Then
                        .Row = .DataRowCnt + 1
                        If .Row > .MaxRows Then
                            .MaxRows = .MaxRows + 1
                            .Row = .MaxRows
                        End If
                    Else
                    
                    End If
                    RowAdd = medGetP(objDic.Fields("key1"), 2, "��")
                    RowADD1 = medGetP(objDic.Fields("key1"), 2, "��")
                End If
                If objWork.Exists(medGetP(objDic.Fields("key1"), 1, "��")) Then
                    objWork.KeyChange medGetP(objDic.Fields("key1"), 1, "��")
                End If
                
                .Col = 1: .Value = "               " & Format(objWork.Fields("time1"), "00:00") & " ~ " & Format(objWork.Fields("time2"), "00:00")
                .FontBold = False: .ForeColor = vbBlack
            End If
            
            For ii = MinCol To MaxCol
                .Col = ii: .Value = objDic.Fields("C" & Format((ii - 1), "00"))
            Next
            ii = ii + 1
            objProBar.Value = ii
            objDic.MoveNext
        Loop
        Dim RowPos As Long
        
        For ii = 1 To .DataRowCnt
            .Row = ii
            RowCnt = 0
            For jj = MinCol To MaxCol
                .Col = jj
                RowCnt = RowCnt + Val(.Value)
            Next
            .Col = MaxCol + 1: .Value = IIf(RowCnt = 0, "", RowCnt)
        Next
        .Row = .DataRowCnt + 2
        If .Row > .MaxRows Then
            .MaxRows = .MaxRows + 2
            .Row = .MaxRows
        End If
        RowPos = .Row
        For ii = MinCol To MaxCol + 1
            .Col = ii
            ColCnt = 0
            For jj = 1 To .DataRowCnt
                .Row = jj
                ColCnt = ColCnt + Val(.Value)
            Next
            .Row = RowPos
            .Value = IIf(ColCnt = 0, "", ColCnt)
            .Col = 1
            .Value = "�հ�"
            .FontBold = True: .ForeColor = vbRed
        Next
        .ReDraw = False
    End With
    Set objProBar = Nothing
End Sub

Private Sub ProGressSetting(Optional ByVal BlnChk As Boolean = False)
    Set objProBar = New jProgressBar.clsProgress
    With objProBar
        .Container = Me
        .Left = lblCaption.Left ' + 1700
        .Top = lblCaption.Top  ' + 30 '
        .Width = (lblCaption.Width)
        .Height = lblCaption.Height
        If BlnChk = True Then .Message = "�ڷḦ �б� ���� �غ����Դϴ�..."
        DoEvents
    End With
End Sub

Private Sub cmdPtCnt_Click()
    Me.MousePointer = 11
    dtpStart.Value = Format(GetSystemDate, "yyyy-mm-dd")
    dtpEnd.Value = Format(GetSystemDate, "yyyy-mm-dd")
    Call GetPtCnt
    fraPtCnt.Visible = True
    fraPtCnt.ZOrder 0
    Me.MousePointer = 0
End Sub

Private Sub GetPtCnt()
    Dim objDic      As New clsDictionary
    Dim objWdic     As New clsDictionary
    Dim objSQL      As New clsLISSqlStatistic
    Dim RS          As Recordset
    
    Dim SSQL        As String
    
    Dim sStart      As String
    Dim sEnd        As String
    
    objDic.Clear
    objDic.FieldInialize "deptcd,doct", "cnt"
    Call medClearTable(tblCnt)
    
    objWdic.Clear
    objWdic.FieldInialize "deptcd,doct", "cnt"
    Call medClearTable(tblWCnt)
    
    
    sStart = Format(dtpStart.Value, "YYYYMMDD")
    sEnd = Format(dtpEnd.Value, "YYYYMMDD")
    
    '�ܷ�ȯ�� �Ǽ�
    SSQL = objSQL.GetPatientCount(enBussDiv.BussDiv_OutPatient, sStart, sEnd)
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
'    Dim objData As clsBasisData
    Dim strData As String
    
    If Not RS.EOF Then
        With tblCnt
            Do Until RS.EOF
                If .DataRowCnt + 1 > .MaxRows Then .MaxRows = .MaxRows + 1
                .Row = .DataRowCnt + 1
                .Col = 1: .Value = RS.Fields("deptcd").Value & ""
'                          Set objData = Nothing
'                          Set objData = New clsBasisData
                          strData = GetDeptNm(.Value)
'                          Set objData = Nothing
                          
'                          ObjLISComCode.DeptCd.KeyChange .Value
                          .Value = strData 'ObjLISComCode.DeptCd.Fields("deptnm")
                          
'                          Set objData = Nothing
'                          Set objData = New clsBasisData
                          strData = GetEmpNm(RS.Fields("orddoct").Value & "")
'                          Set objData = Nothing
                
                .Col = 2: .Value = strData 'getempname(RS.Fields("orddoct").Value & "")
                .Col = 3: .Value = RS.Fields("cnt").Value & ""
                RS.MoveNext
            Loop
        End With
    End If
    Set RS = Nothing
    
    '����ȯ�� �Ǽ�
    SSQL = objSQL.GetPatientCount(enBussDiv.BussDiv_InPatient, sStart, sEnd)
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    If Not RS.EOF Then
        With tblWCnt
             Do Until RS.EOF
                If .DataRowCnt + 1 > .MaxRows Then .MaxRows = .MaxRows + 1
                .Row = .DataRowCnt + 1
                .Col = 1: .Value = RS.Fields("wardid").Value & ""
'                          Set objData = Nothing
'                          Set objData = New clsBasisData
                          strData = GetWardNm(.Value)
'                          Set objData = Nothing
                
'                          ObjLISComCode.WardId.KeyChange .Value
                          .Value = strData 'ObjLISComCode.WardId.Fields("wardnm")
                          
'                          Set objData = Nothing
'                          Set objData = New clsBasisData
                          strData = GetEmpNm(RS.Fields("orddoct").Value & "")
'                          Set objData = Nothing
                          
                .Col = 2: .Value = strData 'getempname(RS.Fields("orddoct").Value & "")
                .Col = 3: .Value = RS.Fields("cnt").Value & ""
                RS.MoveNext
            Loop
        End With

    End If
    
    Set RS = Nothing
    Set objSQL = Nothing
    Set objDic = Nothing
    Set objWdic = Nothing
End Sub

Private Sub cmdQuery_Click()
    Dim strTmp      As String
    Dim strFDT      As String
    Dim strTDT      As String
    Dim STRINOUT    As String
    
    If Mode = "" Then
        MsgBox "��ȸ�׸��� �����Ͻ��� ��ȸ�ϼ���.", vbInformation + vbOKOnly, "��ȸ�׸���"
        Exit Sub
    End If
    
    Me.MousePointer = vbHourglass
    
    strTmp = "": STRINOUT = ""
    If txtCd.Text <> "(��ü)" Then strTmp = txtCd.Text
    
    Call DictionaryClear
    Call ProGressSetting(True)
    
    If optInOut(0).Value Then STRINOUT = "2"
    If optInOut(1).Value Then STRINOUT = "1"
    
    If optC(0).Value = True Then    '## ����
        strFDT = Format(dtpFM.Value, "YYYY") & "0101"
        strTDT = Format(dtpFM.Value, "YYYY") & "1231"
        Call GetCntGarthing(QueryStaticsTestCount(Mode, STRINOUT, strTmp, strFDT, strTDT), "2")
        Call CntDisplayNew("2")
    Else                            '## �Ϻ�
        strFDT = Format(dtpFrDt.Value, "YYYYMM") & "01"
        strTDT = Format(dtpFrDt.Value, "YYYYMM") & "31"
        Call GetCntGarthing(QueryStaticsTestCount(Mode, STRINOUT, strTmp, strFDT, strTDT), "1")
        Call CntDisplayNew("1")
    End If
    Set objProBar = Nothing
    
    Me.MousePointer = vbDefault
End Sub

Private Function GetCntGarthing(ByVal SSQL As String, ByVal optCon As String)
    Dim RS      As Recordset
    Dim strKEY  As String
    Dim ii      As Long
    
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    objdicD.Sort = False
    objDic.Sort = False
    If Not RS.EOF Then
        Call ProGressSetting
        objProBar.Max = RS.RecordCount
        objProBar.Message = "�����Ǽ��� �����ϰ� �ֽ��ϴ�..."
        Do Until RS.EOF
            Select Case Mode
                Case "C1": strKEY = RS.Fields("testnm").Value & ""
                Case "C2": strKEY = RS.Fields("deptcd").Value & "" & "��" & RS.Fields("testnm").Value & ""
                Case "C3": strKEY = RS.Fields("workarea").Value & "" & "��" & RS.Fields("testnm").Value & ""
                Case "C4": strKEY = RS.Fields("eqpcd").Value & "" & "��" & RS.Fields("testnm").Value & ""
                Case "C5": strKEY = RS.Fields("orddoct").Value & "" & "��" & RS.Fields("testnm").Value & ""
            End Select
            
            '���ں� ����(�Ϻ����)
            If objdicD.Exists(strKEY) Then
                objdicD.KeyChange strKEY
                objdicD.Fields("C" & RS.Fields("rcvdt").Value & "") = Val(objdicD.Fields("C" & RS.Fields("rcvdt").Value & "")) + Val(RS.Fields("cnt").Value & "")
            Else
                objdicD.AddNew strKEY, _
                                Join(Array("", "", "", "", "", "", "", "", "", "", "", "", _
                                "", "", "", "", "", "", "", "", "", "", "", "", _
                                "", "", "", "", "", "", "", "", "", "", "", "", ""), COL_DIV)
                
                objdicD.KeyChange strKEY
                objdicD.Fields("C" & RS.Fields("rcvdt").Value & "") = Val(objdicD.Fields("C" & RS.Fields("rcvdt").Value & "")) + Val(RS.Fields("cnt").Value & "")
            End If
            
            '��������
            If objDic.Exists(strKEY) Then
                objDic.KeyChange strKEY
                objDic.Fields("C" & RS.Fields("rcvdt").Value & "") = Val(objDic.Fields("C" & RS.Fields("rcvdt").Value & "")) + Val(RS.Fields("cnt").Value & "")
            Else
                objDic.AddNew strKEY, _
                                Join(Array("", "", "", "", "", "", "", "", "", "", "", ""), COL_DIV)
                
                objDic.KeyChange strKEY
                objDic.Fields("C" & RS.Fields("rcvdt").Value & "") = Val(objDic.Fields("C" & RS.Fields("rcvdt").Value & "")) + Val(RS.Fields("cnt").Value & "")
            End If

            ii = ii + 1
            objProBar.Value = ii
            objProBar.Message = "�����Ǽ��� �����ϰ� �ֽ��ϴ�..." & ii
            RS.MoveNext
        Loop
        objdicD.Sort = True
        objDic.Sort = True
        
    End If
    
    Set RS = Nothing
End Function

Private Sub CntDisplayNew(ByVal optCon As String)
    Dim strTmp      As String
    Dim RowTotal    As Double
    Dim ColTotal    As Double
    Dim ii          As Long
    Dim jj          As Long
    Dim kk          As Long
    Dim strData As String
    
    If objDic.RecordCount < 1 Then Exit Sub
    If objdicD.RecordCount < 1 Then Exit Sub
    
    Call ProGressSetting(False)
    tblD.ReDraw = False
    tblData.ReDraw = False
    RowTotal = 0: ColTotal = 0
    
    If optCon = "1" Then
        '�Ϻ�
        objProBar.Max = objdicD.RecordCount
        objdicD.MoveFirst
        tblD.MaxRows = 0
        Do Until objdicD.EOF
            With tblD
                If .DataRowCnt >= .MaxRows Then
                    .MaxRows = .MaxRows + 1
                    .Row = .MaxRows
                Else
                    .Row = .DataRowCnt + 1
                End If
                Select Case Mode
                    Case "C1":  .Col = 1: .Value = objdicD.Fields("key"): .ForeColor = DCM_LightBlue: .FontBold = True
                    Case "C2":  .Col = 1:
                                .Value = medGetP(objdicD.Fields("key"), 1, "��"): .ForeColor = DCM_LightBlue: .FontBold = True
                                If lblNm.Caption <> "" Then
                                    .Value = lblNm.Caption
                                Else
                                    If .Value <> "" Then
                                        If optInOut(0).Value = True Then
                                        
'                                            Set objData = Nothing
'                                            Set objData = New clsBasisData
                                            strData = GetWardNm(.Value)
'                                            Set objData = Nothing
                                            
'                                            ObjLISComCode.WardId.KeyChange .Value
                                            .Value = strData 'ObjLISComCode.WardId.Fields("wardnm")
                                        Else
                                        
'                                            Set objData = Nothing
'                                            Set objData = New clsBasisData
                                            strData = GetDeptNm(.Value)
'                                            Set objData = Nothing
                                            
'                                            ObjLISComCode.DeptCd.KeyChange .Value
                                            .Value = strData 'ObjLISComCode.DeptCd.Fields("deptnm")
                                        End If
                                    End If
                                End If
                                
                                If strTmp <> .Value Then
                                    strTmp = .Value
                                    If .DataRowCnt >= .MaxRows Then
                                        .MaxRows = .MaxRows + 1
                                        .Row = .MaxRows
                                    Else
                                        .Row = .DataRowCnt + 1
                                    End If
                                End If
                               .Col = 1: .Value = medGetP(objdicD.Fields("key"), 2, "��"): .ForeColor = DCM_LightRed: .FontBold = True
                    Case "C3": .Col = 1: .Value = medGetP(objdicD.Fields("key"), 1, "��"): .ForeColor = DCM_LightBlue: .FontBold = True
                               If lblNm.Caption = "" Then
                                    If .Value <> "" Then .Value = GetConditionNm(.Value)
                               Else
                                    .Value = lblNm.Caption
                               End If
                               
                               If strTmp <> .Value Then
                                    strTmp = .Value
                                    If .DataRowCnt >= .MaxRows Then
                                        .MaxRows = .MaxRows + 1
                                        .Row = .MaxRows
                                    Else
                                        .Row = .DataRowCnt + 1
                                    End If
                                    
                                End If
                                .Col = 1: .Value = medGetP(objdicD.Fields("key"), 2, "��"): .ForeColor = DCM_LightRed: .FontBold = True
                    Case "C4":  .Col = 1:
                            If lblNm.Caption = "" Then
                                 .Value = medGetP(objdicD.Fields("key"), 1, "��"): .ForeColor = DCM_LightBlue: .FontBold = True
                                 If .Value = "" Then
                                     .Value = "���۾�": .ForeColor = DCM_LightBlue: .FontBold = True
                                 Else
                                     .Value = GetConditionNm(.Value)
                                 End If
                                If strTmp <> .Value Then
                                    strTmp = .Value
                                    If .DataRowCnt >= .MaxRows Then
                                        .MaxRows = .MaxRows + 1
                                        .Row = .MaxRows
                                    Else
                                        .Row = .DataRowCnt + 1
                                    End If
                                End If
                                .Col = 1: .Value = medGetP(objdicD.Fields("key"), 2, "��"): .ForeColor = DCM_LightRed: .FontBold = True
                            Else
                                If Trim(medGetP(objdicD.Fields("key"), 1, "��")) <> "" Then
                                    .Value = medGetP(objdicD.Fields("key"), 1, "��"): .ForeColor = DCM_LightBlue: .FontBold = True
                                    .Value = lblNm.Caption
                                    If strTmp <> .Value Then
                                    strTmp = .Value
                                        If .DataRowCnt >= .MaxRows Then
                                            .MaxRows = .MaxRows + 1
                                            .Row = .MaxRows
                                        Else
                                            .Row = .DataRowCnt + 1
                                        End If
                                    End If
                                    .Col = 1: .Value = medGetP(objdicD.Fields("key"), 2, "��"): .ForeColor = DCM_LightRed: .FontBold = True
                                Else
                                    GoTo Skip
                                End If
                            End If
                    Case "C5": .Col = 1:
                            If lblNm.Caption = "" Then
                                .Value = medGetP(objdicD.Fields("key"), 1, "��"): .ForeColor = DCM_LightBlue: .FontBold = True
                                If .Value = "" Then
                                    .Value = "��Ȯ��": .ForeColor = DCM_LightBlue: .FontBold = True
                                Else
'                                    Set objData = Nothing
'                                    Set objData = New clsBasisData
                                    strData = GetEmpNm(.Value)
'                                    Set objData = Nothing
                                    
                                    .Value = strData 'getempname(.Value)
                                End If
                                If strTmp <> .Value Then
                                      strTmp = .Value
                                     If .DataRowCnt >= .MaxRows Then
                                        
                                         .MaxRows = .MaxRows + 1
                                         .Row = .MaxRows
                                     Else
                                         .Row = .DataRowCnt + 1
                                     End If
                                 End If
                                .Col = 1: .Value = medGetP(objdicD.Fields("key"), 2, "��"): .ForeColor = DCM_LightRed: .FontBold = True
                            Else
                                If Trim(.Value = medGetP(objdicD.Fields("key"), 1, "��")) <> "" Then
                                    .Value = medGetP(objdicD.Fields("key"), 1, "��"): .ForeColor = DCM_LightBlue: .FontBold = True
                                    .Value = lblNm.Caption
                                    If strTmp <> .Value Then
                                          strTmp = .Value
                                         If .DataRowCnt >= .MaxRows Then
                                            
                                             .MaxRows = .MaxRows + 1
                                             .Row = .MaxRows
                                         Else
                                             .Row = .DataRowCnt + 1
                                         End If
                                     End If
                                    .Col = 1: .Value = medGetP(objdicD.Fields("key"), 2, "��"): .ForeColor = DCM_LightRed: .FontBold = True
                                Else
                                    GoTo Skip
                                End If
                            End If
                End Select
                
                For ii = 1 To 31
                    .Col = ii + 1
                    .Value = objdicD.Fields("C" & Format(ii, "00"))
                Next
            End With
Skip:
            kk = kk + 1
            objProBar.Value = kk
            objProBar.Message = "�ڷḦ �������Դϴ�....  " & kk
            objdicD.MoveNext
        Loop
        
        With tblD
            .MaxRows = .MaxRows + 2
            'Row�հ�
            For ii = 1 To .DataRowCnt
                .Row = ii
                For jj = 2 To .MaxCols - 1
                    .Col = jj
                    RowTotal = RowTotal + Val(.Value)
                Next
                .Col = .MaxCols: .Value = RowTotal
                If .Value = 0 Then .Value = ""
                RowTotal = 0
            Next
            For ii = 2 To .MaxCols
                .Col = ii
                For jj = 1 To .DataRowCnt
                    .Row = jj
                    ColTotal = ColTotal + Val(.Value)
                Next
                .Row = .MaxRows: .Value = ColTotal
                If .Value = 0 Then .Value = ""
                .Col = 1: .Value = "��  ��": .ForeColor = DCM_LightBlue: .FontBold = True
                ColTotal = 0
            Next
        End With
    Else
    '����
        objProBar.Max = objDic.RecordCount
        objDic.MoveFirst
        tblData.MaxRows = 0
        Do Until objDic.EOF
            With tblData
                If .DataRowCnt >= .MaxRows Then
                    .MaxRows = .MaxRows + 1
                    .Row = .MaxRows
                Else
                    .Row = .DataRowCnt + 1
                End If
                Select Case Mode
                    Case "C1":  .Col = 1: .Value = objDic.Fields("key"): .ForeColor = DCM_LightBlue: .FontBold = True
                    Case "C2":  .Col = 1:
                                .Value = medGetP(objDic.Fields("key"), 1, "��"): .ForeColor = DCM_LightBlue: .FontBold = True
                                If lblNm.Caption <> "" Then
                                    .Value = lblNm.Caption
                                Else
                                    If .Value <> "" Then
                                        
                                        If optInOut(0).Value = True Then
'                                            Set objData = Nothing
'                                            Set objData = New clsBasisData
                                            strData = GetWardNm(.Value)
'                                            Set objData = Nothing
                                        
'                                            ObjLISComCode.WardId.KeyChange .Value
                                            .Value = strData 'ObjLISComCode.WardId.Fields("wardnm")
                                        Else
'                                            Set objData = Nothing
'                                            Set objData = New clsBasisData
                                            strData = GetDeptNm(.Value)
'                                            Set objData = Nothing

'                                            ObjLISComCode.DeptCd.KeyChange .Value
                                            .Value = strData 'ObjLISComCode.DeptCd.Fields("deptnm")
                                        End If
                                    End If
                                End If
                                If strTmp <> .Value Then
                                    strTmp = .Value
                                    If .DataRowCnt >= .MaxRows Then
                                        .MaxRows = .MaxRows + 1
                                        .Row = .MaxRows
                                    Else
                                        .Row = .DataRowCnt + 1
                                    End If
                                End If
                               .Col = 1: .Value = medGetP(objDic.Fields("key"), 2, "��"): .ForeColor = DCM_LightRed: .FontBold = True
                    
                    Case "C3": .Col = 1: .Value = medGetP(objDic.Fields("key"), 1, "��"): .ForeColor = DCM_LightBlue: .FontBold = True
                                If lblNm.Caption <> "" Then
                                    .Value = lblNm.Caption
                                Else
                                    If .Value <> "" Then .Value = GetConditionNm(.Value)
                                End If
                                If strTmp <> .Value Then
                                    strTmp = .Value
                                    If .DataRowCnt >= .MaxRows Then
                                        .MaxRows = .MaxRows + 1
                                        .Row = .MaxRows
                                    Else
                                        .Row = .DataRowCnt + 1
                                    End If
                                    
                                End If
                                .Col = 1: .Value = medGetP(objDic.Fields("key"), 2, "��"): .ForeColor = DCM_LightRed: .FontBold = True
                    
                    Case "C4": .Col = 1:
                                If lblNm.Caption = "" Then
                                    If .Value = "" Then
                                        .Value = "���۾�": .ForeColor = DCM_LightBlue: .FontBold = True
                                    Else
                                        .Value = GetConditionNm(.Value)
                                    End If
                                    
                                    If strTmp <> .Value Then
                                        strTmp = .Value
                                        If .DataRowCnt >= .MaxRows Then
                                            .MaxRows = .MaxRows + 1
                                            .Row = .MaxRows
                                        Else
                                            .Row = .DataRowCnt + 1
                                        End If
                                    End If
                                    .Col = 1: .Value = medGetP(objDic.Fields("key"), 2, "��"): .ForeColor = DCM_LightRed: .FontBold = True
                                Else
                                    If Trim(medGetP(objDic.Fields("key"), 1, "��")) <> "" Then
                                        .Value = medGetP(objDic.Fields("key"), 1, "��"): .ForeColor = DCM_LightBlue: .FontBold = True
                                        .Value = lblNm.Caption
                                        
                                        If strTmp <> .Value Then
                                        strTmp = .Value
                                            If .DataRowCnt >= .MaxRows Then
                                                .MaxRows = .MaxRows + 1
                                                .Row = .MaxRows
                                            Else
                                                .Row = .DataRowCnt + 1
                                            End If
                                        End If
                                        .Col = 1: .Value = medGetP(objDic.Fields("key"), 2, "��"): .ForeColor = DCM_LightRed: .FontBold = True
                                    Else
                                        GoTo Skip1
                                    End If
                                End If
                    Case "C5": .Col = 1: .Value = medGetP(objDic.Fields("key"), 1, "��"): .ForeColor = DCM_LightBlue: .FontBold = True
                                If lblNm.Caption <> "" Then
                                    .Value = lblNm.Caption
                                Else
                                    If .Value = "" Then
                                        .Value = "��Ȯ��": .ForeColor = DCM_LightBlue: .FontBold = True
                                    Else
'                                        Set objData = Nothing
'                                        Set objData = New clsBasisData
                                        strData = GetEmpNm(.Value)
'                                        Set objData = Nothing

                                        .Value = strData 'getempname(.Value)
                                    End If
                                End If
                                If strTmp <> .Value Then
                                     strTmp = .Value
                                    If .DataRowCnt >= .MaxRows Then
                                       
                                        .MaxRows = .MaxRows + 1
                                        .Row = .MaxRows
                                    Else
                                        .Row = .DataRowCnt + 1
                                    End If
                                End If
                            .Col = 1: .Value = medGetP(objDic.Fields("key"), 2, "��"): .ForeColor = DCM_LightRed: .FontBold = True
                End Select
                
                For ii = 1 To 12
                    .Col = ii + 1
                    .Value = objDic.Fields("C" & Format(ii, "00"))
                Next
            End With
            
Skip1:
            kk = kk + 1
            objProBar.Value = kk
            objProBar.Message = "�ڷḦ �������Դϴ�....  " & kk
            objDic.MoveNext
        Loop
        
        With tblData
            .MaxRows = .MaxRows + 2
            'Row�հ�
            For ii = 1 To .DataRowCnt
                .Row = ii
                For jj = 2 To .MaxCols - 1
                    .Col = jj
                    RowTotal = RowTotal + Val(.Value)
                Next
                .Col = .MaxCols: .Value = RowTotal
                If .Value = 0 Then .Value = ""
                RowTotal = 0
            Next
            For ii = 2 To .MaxCols
                .Col = ii
                For jj = 1 To .DataRowCnt
                    .Row = jj
                    ColTotal = ColTotal + Val(.Value)
                Next
                .Row = .MaxRows:
                
                .Value = ColTotal
                If .Value = 0 Then .Value = ""
                .Col = 1: .Value = "��  ��": .ForeColor = DCM_LightBlue: .FontBold = True
                ColTotal = 0
            Next
        End With
    End If
    If tblD.MaxRows < 30 Then tblD.MaxRows = 30
    If tblData.MaxRows < 30 Then tblData.MaxRows = 30
    tblD.ReDraw = True
    tblData.ReDraw = True
End Sub

Private Sub cmdStart_Click()
    Call medClearTable(tblCnt)
    fraPtCnt.Visible = False
End Sub

Private Sub cmdRefresh_Click()
    Call GetPtCnt
End Sub

'[�ð��뺰�� ��踦 �ۼ��Ϸ��� �Ͽ�����,
' Running Time�� ������� ����]
Private Sub cmdTime_Click()
    Dim strTmp As String
    Dim strVal1 As String
    Dim strVal2 As String
    Dim strFil1 As String
    Dim strFil2 As String
    Dim SSQL    As String
    Dim ii      As Integer
    Dim jj      As Integer
    
    On Error GoTo SAVE_ERROR
    
    strTmp = MsgBox("���ο� ������� �����մϴ�." & vbCrLf & "�����Ͻðڽ��ϱ�?", vbYesNo, "�۾��ð�����")
    If strTmp = vbYes Then
        DBConn.BeginTrans
        SSQL = "delete " & T_LAB031 & " WHERE " & DBW("cdindex=", "C123")
        DBConn.Execute SSQL
        With tblTime
            For ii = 1 To .DataRowCnt
                .Row = ii
                .Col = 2: strVal1 = Trim(.Value)
                If strVal1 <> "" Then
                    .Col = 3: strFil1 = Format(Mid(Trim(Replace(.Value, ":", "")), 1, 4), "0000")
                    .Col = 4: strFil2 = Format(Mid(Trim(Replace(.Value, ":", "")), 1, 4), "0000")
                    SSQL = "insert into " & T_LAB031 & "(cdindex,cdval1,cdval2,field1,field2) values(" & _
                           DBV("cdindex", "C123", 1) & DBV("cdval1", strVal1, 1) & DBV("cdval2", ii, 1) & DBV("field1", strFil1, 1) & _
                           DBV("field2", strFil2) & ")"
                    DBConn.Execute SSQL
                End If
                strVal1 = ""
            Next
        End With
        DBConn.CommitTrans
        Dim RS  As Recordset
        
        Set RS = New Recordset
        RS.Open MySqlStmt.WorkTime, DBConn
        
        If Not RS.EOF Then
            objWork.DeleteAll
            Do Until RS.EOF
                If objWork.Exists(RS.Fields("cdval1").Value & "") Then
                    objWork.KeyChange RS.Fields("cdval1").Value & ""
                    objWork.Fields("time1") = RS.Fields("field1").Value & ""
                    objWork.Fields("time2") = RS.Fields("field2").Value & ""
                Else
                    objWork.AddNew RS.Fields("cdval1").Value & "", _
                                   Join(Array(RS.Fields("field1").Value & "", RS.Fields("field2").Value & ""), COL_DIV)
                End If
                RS.MoveNext
            Loop
        End If
        Set RS = Nothing
        cmdQuery.Enabled = True
        
        With tblTime
            objWork.MoveFirst
            Do Until objWork.EOF
                ii = ii + 1
                .Row = ii
                .Col = 2: .Value = objWork.Fields("cdval1")
                .Col = 3: .Value = objWork.Fields("field1")
                .Col = 4: .Value = objWork.Fields("filed2")
                objWork.MoveNext
            Loop
        End With
    End If
    Exit Sub
SAVE_ERROR:
    DBConn.RollbackTrans
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub cmdExit_Click()
    Unload Me
    If IsLastForm Then RaiseEvent LastFormUnload
End Sub

Private Sub Form_Activate()
    MainFrm.lblSubMenu.Caption = Me.Caption
End Sub

Private Sub Form_Initialize()
    objWork.Clear
    objWork.FieldInialize "type", "time1,time2"
    objDicDT.Clear
    objDicDT.FieldInialize "key,key1", "C01,C02,C03,C04,C05,C06,C07,C08,C09,C10," & _
                                "C11,C12,C13,C14,C15,C16,C17,C18,C19,C20," & _
                                "C21,C22,C23,C24,C25,C26,C27,C28,C29,C30,C31"
    objdicD.Clear
    objdicD.FieldInialize "key", "C01,C02,C03,C04,C05,C06,C07,C08,C09,C10," & _
                                "C11,C12,C13,C14,C15,C16,C17,C18,C19,C20," & _
                                "C21,C22,C23,C24,C25,C26,C27,C28,C29,C30,C31"
    objDicT.Clear
    objDicT.FieldInialize "key,key1", "C01,C02,C03,C04,C05,C06,C07,C08,C09,C10,C11,C12"
    objDic.Clear
    objDic.FieldInialize "key", "C01,C02,C03,C04,C05,C06,C07,C08,C09,C10,C11,C12"
    
End Sub

Private Sub Form_Load()
    Dim RS  As Recordset
    Dim ii  As Integer
    
    Call medClearTable(tblTime)
'    fraInOut.Visible = False
    Set RS = New Recordset
    RS.Open MySqlStmt.WorkTime, DBConn
    
    If Not RS.EOF Then
        Do Until RS.EOF
            If objWork.Exists(RS.Fields("cdval1").Value & "") Then
                objWork.Fields("time1") = RS.Fields("field1").Value & ""
                objWork.Fields("time2") = RS.Fields("field2").Value & ""
            Else
                objWork.AddNew RS.Fields("cdval1").Value & "", _
                               Join(Array(RS.Fields("field1").Value & "", RS.Fields("field2").Value & ""), COL_DIV)
            End If
            ii = ii + 1
            tblTime.Row = ii
            tblTime.Col = 2: tblTime.Value = RS.Fields("cdval1").Value & ""
            tblTime.Col = 3: tblTime.Value = RS.Fields("field1").Value & ""
            tblTime.Col = 4: tblTime.Value = RS.Fields("field2").Value & ""
            RS.MoveNext
        Loop
    End If
    Set RS = Nothing
    
    If objWork.RecordCount = 0 Then
        MsgBox "�۾��ð��� �ۼ����ּ���", vbInformation + vbOKOnly, "�۾��ð����"
        cmdQuery.Enabled = False
    End If
    
    With tvwMenu
        .Nodes.Clear
        Call .Nodes.Add(, , "R", "���Ǽ���")
        Call .Nodes.Add("R", tvwChild, "C1", "�˻��׸�")
        Call .Nodes.Add("R", tvwChild, "C2", "����μ���")
        Call .Nodes.Add("R", tvwChild, "C3", "������Ʈ��")
        Call .Nodes.Add("R", tvwChild, "C4", "�˻����")
        Call .Nodes.Add("R", tvwChild, "C5", "ó����")
        Call .Nodes(.Nodes.Count).EnsureVisible
    End With
    
    '�Ϲݰ˻�
    strTable = T_LAB201 & " a, " & T_LAB302 & " b, " & T_LAB001 & " d "
    '��Ÿ�˻�
    strTable1 = T_LAB201 & " a, " & T_LAB351 & " b, " & T_LAB001 & " d "
    '�̻����˻�
    strTable2 = T_LAB201 & " a, " & T_LAB404 & " b, " & T_LAB001 & " d "
    
    FraM.Visible = False
    FraD.Visible = True
    tblData.Visible = False
    tblD.Visible = True
    dtpFM.Value = GetSystemDate ' Format(GetSystemDate, "yyyy")
    dtpFrDt.Value = GetSystemDate 'Format(GetSystemDate, "yyyy-mm")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objDic = Nothing
    Set objdicD = Nothing
    Set objDicT = Nothing
    Set objDicDT = Nothing
    Set objWork = Nothing
    Set MySqlStmt = Nothing
    Set objProBar = Nothing
End Sub

Private Sub optC_Click(Index As Integer)
    Call DictionaryClear
    Select Case Index
        Case 0  '����
            FraM.Visible = True
            FraD.Visible = False
            tblData.Visible = True
            tblD.Visible = False
            tblData.Row = 0: tblData.Col = 0: tblData.Value = lblCaption.Caption
        Case 1  '�Ϻ�
            FraM.Visible = False
            FraD.Visible = True
            tblData.Visible = False
            tblD.Visible = True
            tblD.Row = 0: tblD.Col = 0: tblD.Value = lblCaption.Caption
    End Select
    txtCd.Text = "": lblNm.Caption = ""
End Sub

Private Sub optTime_Click(Index As Integer)
    Call medClearTable(tblData)
    Call medClearTable(tblD)
    If Index = 0 Then
        If optC(0).Value Then
            Call DisPlayData(objDic, tblData, 2, 13)
        Else
            Call DisPlayData(objdicD, tblD, 2, 32)
        End If
    Else
        If optC(0).Value Then
            Call DisPlayTime(objDicT, tblData, 2, 13)
        Else
            Call DisPlayTime(objDicDT, tblD, 2, 32)
        End If
    End If
End Sub

Private Sub tvwMenu_NodeClick(ByVal Node As MSComctlLib.Node)
    '����Ÿ �ʱ�ȭ
    Call DictionaryClear
    With tvwMenu
        Mode = Node.Key
        Select Case Mode
            Case "R"
                Mode = ""
                ModeString = ""
                lblCaption.Caption = "������ �����ϼ���"
                lblCondition.Caption = ""
                txtCd.Text = "": lblNm.Caption = ""
                Exit Sub
            Case "C1"
                lblCondition.Caption = "�˻��׸�"
            Case "C2"
                lblCondition.Caption = "����μ�"
            Case "C3"
                lblCondition.Caption = "WORKAREA"
            Case "C4"
                lblCondition.Caption = "�˻����"
            Case "C5"
                lblCondition.Caption = "��ġ��"
            
        End Select

        ModeString = Node.Text
        txtCd.Text = "": lblNm.Caption = "": lblCaption.Caption = ""
        If Mode <> "" Then
            lblCaption.Caption = ModeString & " �� �ۼ��մϴ�"
            
            txtCd.Text = "(��ü)": txtCd.Locked = True
        End If
    End With
End Sub

Private Function GetConditionNm(ByVal NameCd As String) As String
    If Mode <> "C1" Then
        GetConditionNm = MySqlStmt.ConditionNm(Mode, NameCd)
    End If
End Function

Private Sub DictionaryClear()
    objDic.DeleteAll
    objDicT.DeleteAll
    objdicD.DeleteAll
    objDicDT.DeleteAll
    medClearTable tblD
    medClearTable tblData
End Sub

'==================================================================================
'                           '�������� SQL ����                                    =
'==================================================================================
Private Function DetailSQLStmt(ByVal qCondition As String, ByVal Inout As String, Optional ByVal Sel As String = "") As String
    
    Select Case Mode
        Case "C1":
                If qCondition <> "" Then DetailSQLStmt = " AND " & DBW("b.testcd=", qCondition)
        Case "C2":
                If qCondition <> "" Then DetailSQLStmt = " AND " & DBW("a.deptcd=", qCondition)
        Case "C3":
'                If qCondition <> "" Then DetailSQLStmt = " AND " & DBW("a.workarea>=", qCondition) & " AND " & DBW("a.workarea<=", qCondition)
                If qCondition <> "" Then DetailSQLStmt = " AND " & DBW("a.workarea=", qCondition) & " "
        Case "C4":
                If Sel = "" Then
                    If qCondition <> "" Then DetailSQLStmt = " AND " & DBW("b.eqpcd=", qCondition)
                End If
        Case "C5":
                If qCondition <> "" Then DetailSQLStmt = " AND " & DBW("a.orddoct=", qCondition)
    End Select
    
    '�������� ����� ������ �ʴ� Workarea
    If P_NoWorkareaReport <> "" Then
        DetailSQLStmt = DetailSQLStmt & " AND a.workarea not in (" & P_NoWorkareaReport & ")"
    End If
    
    '��� ��ȸ����(True=itemseq,False=detailfg)
    If P_ItemSeqFG = True Then
        DetailSQLStmt = DetailSQLStmt & "AND (d.itemseq>0) "
    Else
        DetailSQLStmt = DetailSQLStmt & " AND (d.detailfg = ''  or  d.detailfg is null)"
    End If
    
    '2:�Կ�, 1:�ܷ� ����
    Select Case Inout
        Case "1"
            DetailSQLStmt = DetailSQLStmt & " AND (a.wardid IS NULL OR a.wardid='')"
        Case "2"
            DetailSQLStmt = DetailSQLStmt & " AND a.wardid <> '' "
    End Select
            
End Function

Private Function SQLStmtH(Optional ByVal FirstFg As Boolean = True, Optional ByVal Inout As String, Optional ByVal Sel As String = "") As String
    Dim strTmp As String
    
    If optC(0).Value = True Then
        strTmp = "substr(a.accdt,5,2) as rcvdt,"
    Else
        strTmp = "substr(a.accdt,7) as rcvdt,"
    End If
    
    '## 5.0.2: �̻��(2004-12-29)
    '   - ��Ȯ�� �ε����� Ÿ���ʾ� RULE Base ����
    '## �½�ȣ(2010.03.31)
    '�꺣�̽����� �ε����� ��ȯ
    If FirstFg = True Then
        Select Case Mode
            Case "C1":
                SQLStmtH = " SELECT /*+ S2LAB201_IDX2 */ " & strTmp & " b.testcd, d.testnm, count(*) as Cnt "
            Case "C2":
                If Inout = "2" Then
                    SQLStmtH = " SELECT /*+ S2LAB201_IDX2 */ " & strTmp & " b.testcd, a.wardid as deptcd,d.testnm, count(*) as Cnt "
                Else
                    SQLStmtH = " SELECT  /*+ S2LAB201_IDX2 */ " & strTmp & " b.testcd, a.deptcd as deptcd,d.testnm, count(*) as Cnt "
                End If
            Case "C3":
                SQLStmtH = " SELECT  " & strTmp & " b.testcd, a.workarea,d.testnm, count(*) as Cnt "
            Case "C4":
                If Sel = "" Then
                    SQLStmtH = " SELECT  /*+ S2LAB201_IDX2 */ " & strTmp & " b.testcd, b.eqpcd,d.testnm, count(*) as Cnt "
                Else
                    SQLStmtH = " SELECT  /*+ S2LAB201_IDX2 */ " & strTmp & " b.testcd, '' as eqpcd,d.testnm, count(*) as Cnt "
                End If
            Case "C5":
                SQLStmtH = " SELECT  /*+ S2LAB201_IDX2 */ " & strTmp & " b.testcd, a.orddoct,d.testnm, count(*) as Cnt "
        End Select
    Else
        Select Case Mode
            Case "C1":
                    SQLStmtH = " GROUP BY a.accdt, b.testcd, d.testnm "
            Case "C2":
                If Inout = "2" Then
                    SQLStmtH = " GROUP BY a.accdt,a.wardid, b.testcd, d.testnm "
                Else
                    SQLStmtH = " GROUP BY a.accdt,a.deptcd, b.testcd, d.testnm "
                End If
            Case "C3":
                SQLStmtH = " GROUP BY a.accdt,a.workarea, b.testcd, d.testnm "
            Case "C4":
                If Sel = "" Then
                    SQLStmtH = " GROUP BY a.accdt, b.eqpcd,b.testcd, d.testnm "
                Else
                    SQLStmtH = " GROUP BY a.accdt, b.testcd, d.testnm "
                End If
                
            Case "C5":
                SQLStmtH = " GROUP BY a.accdt,a.orddoct, b.testcd, d.testnm "
        End Select
    End If
End Function

'==================================================================================
'                           '�˻��׸� ���                                      =
'==================================================================================
Private Function QueryStaticsTestCount(ByVal qMODE As String, ByVal qinout As String, ByVal qCondition As String, ByVal pStart As String, ByVal PEND As String) As String
    Dim SqlStmt   As String
    Dim strDept     As String
    

    SqlStmt = SQLStmtH(True, qinout)
    
    SqlStmt = SqlStmt & " FROM  " & strTable & _
              " WHERE " & DBW("a.accdt>=", Format(pStart, "########")) & " " & _
              " AND   " & DBW("a.accdt<=", Format(PEND, "########")) & " " & _
              " AND b.workarea = a.workarea " & _
              " AND b.accdt    = a.accdt " & _
              " AND b.accseq   = a.accseq "
    
'    SqlStmt = SqlStmt & " AND b.vfydt is not null AND  b.vfydt <> ' ' "
    
    SqlStmt = SqlStmt & _
                        " AND d.testcd = b.testcd " & _
                        " AND d.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " " & _
                                          "WHERE  testcd = b.testcd) "
    
    SqlStmt = SqlStmt & DetailSQLStmt(qCondition, qinout)
    
    SqlStmt = SqlStmt & " AND a.qcfg<>'1'"
    
    SqlStmt = SqlStmt & SQLStmtH(False, qinout)
    
'��Ÿ�˻�
    SqlStmt = SqlStmt & "UNION ALL " & SQLStmtH(True, qinout, "Eqpcd")
    
    SqlStmt = SqlStmt & "FROM  " & strTable1
    SqlStmt = SqlStmt & "WHERE " & DBW("a.accdt>=", Format(pStart, "########")) & " "
    SqlStmt = SqlStmt & "AND   " & DBW("a.accdt<=", Format(PEND, "########")) & " "
    SqlStmt = SqlStmt & " AND b.workarea = a.workarea "
    SqlStmt = SqlStmt & " AND b.accdt    = a.accdt "
    SqlStmt = SqlStmt & " AND b.accseq   = a.accseq "
    
'    SqlStmt = SqlStmt & " AND b.vfydt is not null AND  b.vfydt <> ' ' "
    SqlStmt = SqlStmt & " AND d.testcd = b.testcd "
    SqlStmt = SqlStmt & " AND d.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " " & _
                                  "WHERE  testcd = b.testcd) "
    
    SqlStmt = SqlStmt & DetailSQLStmt(qCondition, qinout, "Eqpcd")
    
    SqlStmt = SqlStmt & " AND a.qcfg<>'1'"
    SqlStmt = SqlStmt & SQLStmtH(False, qinout, "Eqpcd")
    
'�̻����˻�
    SqlStmt = SqlStmt & "UNION ALL " & SQLStmtH(True, qinout, "Eqpcd")
    
    SqlStmt = SqlStmt & "FROM  " & strTable2
    SqlStmt = SqlStmt & "WHERE " & DBW("a.accdt>=", Format(pStart, "########")) & " "
    SqlStmt = SqlStmt & "AND   " & DBW("a.accdt<=", Format(PEND, "########")) & " "
    SqlStmt = SqlStmt & " AND b.workarea = a.workarea "
    SqlStmt = SqlStmt & " AND b.accdt    = a.accdt "
    SqlStmt = SqlStmt & " AND b.accseq   = a.accseq "
    
'    SqlStmt = SqlStmt & " AND b.vfydt is not null AND  b.vfydt <> ' ' "
    SqlStmt = SqlStmt & " AND d.testcd = b.testcd "
    SqlStmt = SqlStmt & " AND d.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " " & _
                                          "WHERE  testcd = b.testcd) "
    
    SqlStmt = SqlStmt & DetailSQLStmt(qCondition, qinout, "Eqpcd")
    
    SqlStmt = SqlStmt & " AND a.qcfg<>'1'"
    
    SqlStmt = SqlStmt & SQLStmtH(False, qinout, "Eqpcd")
    
    Debug.Print SqlStmt
    QueryStaticsTestCount = SqlStmt
End Function
