VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frm312QCSchedule_N 
   BackColor       =   &H00DBE6E6&
   ClientHeight    =   9795
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   15240
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9795
   ScaleWidth      =   15240
   WindowState     =   2  '�ִ�ȭ
   Begin FPSpread.vaSpread tblSchedule 
      Height          =   6465
      Left            =   5340
      TabIndex        =   40
      Top             =   1965
      Width           =   9120
      _Version        =   196608
      _ExtentX        =   16087
      _ExtentY        =   11404
      _StockProps     =   64
      EditModePermanent=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   14
      MaxRows         =   21
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   15263976
      SpreadDesigner  =   "frm312QCSchedule_N.frx":0000
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ȭ������(&C)"
      Height          =   510
      Left            =   11820
      Style           =   1  '�׷���
      TabIndex        =   16
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00E0E0E0&
      Caption         =   "����(&X)"
      Height          =   510
      Left            =   13140
      Style           =   1  '�׷���
      TabIndex        =   15
      Top             =   8535
      Width           =   1320
   End
   Begin MedControls1.LisLabel LisLabel6 
      Height          =   300
      Left            =   75
      TabIndex        =   0
      Top             =   45
      Width           =   14385
      _ExtentX        =   25374
      _ExtentY        =   529
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
      Caption         =   "�� ��Ʈ�� ����"
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   1380
      Left            =   75
      TabIndex        =   1
      Top             =   270
      Width           =   14400
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00E0E0E0&
         Caption         =   "��������(&S)"
         Height          =   510
         Left            =   12765
         Style           =   1  '�׷���
         TabIndex        =   71
         Top             =   465
         Width           =   1320
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00E0E0E0&
         Caption         =   "�������(&P)"
         Height          =   510
         Left            =   12765
         Style           =   1  '�׷���
         TabIndex        =   72
         Top             =   735
         Visible         =   0   'False
         Width           =   1320
      End
      Begin MedControls1.LisLabel lblCtrlNm 
         Height          =   360
         Left            =   4050
         TabIndex        =   4
         Top             =   135
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   635
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
      Begin VB.TextBox txtCtrlCd 
         Height          =   360
         Left            =   1425
         MaxLength       =   9
         TabIndex        =   3
         Text            =   "�ϵѼ³ݴٿ��Ͽ���"
         Top             =   150
         Width           =   2280
      End
      Begin VB.CommandButton cmdPopCtrl 
         BackColor       =   &H00F4F0F2&
         Height          =   360
         Left            =   3705
         Picture         =   "frm312QCSchedule_N.frx":08C4
         Style           =   1  '�׷���
         TabIndex        =   2
         Top             =   135
         Width           =   330
      End
      Begin MedControls1.LisLabel lblCtrlDiv 
         Height          =   360
         Left            =   5550
         TabIndex        =   5
         Top             =   555
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   635
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "������������"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblEqp 
         Height          =   360
         Left            =   9720
         TabIndex        =   6
         Top             =   555
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   635
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "C001 Coulter Stks"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblBuilding 
         Height          =   360
         Left            =   1425
         TabIndex        =   7
         Top             =   960
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   635
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "10 ����"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblSection 
         Height          =   360
         Left            =   5550
         TabIndex        =   8
         Top             =   960
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   635
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "HE Hematology"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblWorkarea 
         Height          =   360
         Left            =   9720
         TabIndex        =   9
         Top             =   960
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   635
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "03 Hematology"
         Appearance      =   0
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00DBE6E6&
         Height          =   465
         Left            =   1425
         TabIndex        =   10
         Top             =   465
         Width           =   2640
         Begin VB.OptionButton optLevelCd 
            BackColor       =   &H00DBE6E6&
            Caption         =   "High"
            Height          =   180
            Index           =   2
            Left            =   1800
            TabIndex        =   13
            Top             =   150
            Width           =   705
         End
         Begin VB.OptionButton optLevelCd 
            BackColor       =   &H00DBE6E6&
            Caption         =   "Normal"
            Height          =   180
            Index           =   1
            Left            =   810
            TabIndex        =   12
            Top             =   150
            Width           =   960
         End
         Begin VB.OptionButton optLevelCd 
            BackColor       =   &H00DBE6E6&
            Caption         =   "Low"
            Height          =   180
            Index           =   0
            Left            =   60
            TabIndex        =   11
            Top             =   150
            Value           =   -1  'True
            Width           =   705
         End
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   11
         Left            =   45
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   150
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "Control ����"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   1
         Left            =   45
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   960
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "�ǹ�����"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   3
         Left            =   45
         TabIndex        =   66
         Top             =   555
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "Level ����"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   4
         Left            =   8280
         TabIndex        =   67
         Top             =   555
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "�˻����"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   5
         Left            =   8280
         TabIndex        =   68
         Top             =   960
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   2
         Left            =   4170
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   555
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "������������"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   6
         Left            =   4170
         TabIndex        =   70
         Top             =   960
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "���Ǳ���"
         Appearance      =   0
      End
   End
   Begin MedControls1.LisLabel LisLabel3 
      Height          =   300
      Left            =   75
      TabIndex        =   26
      Top             =   3255
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   529
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
      Caption         =   "�� �˻����� ����"
      Appearance      =   0
   End
   Begin VB.Frame fraConfigDate 
      BackColor       =   &H00DBE6E6&
      Height          =   4995
      Left            =   75
      TabIndex        =   27
      Top             =   3465
      Width           =   3645
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   8
         Left            =   75
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   3375
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "�˻��ϼ�"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   9
         Left            =   75
         TabIndex        =   75
         Top             =   3780
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "�˻�Ƚ��"
         Appearance      =   0
      End
      Begin VB.CommandButton cmdSelDateClear 
         BackColor       =   &H00EFFCFC&
         Caption         =   "��������"
         Height          =   435
         Left            =   60
         Style           =   1  '�׷���
         TabIndex        =   61
         Top             =   4305
         Width           =   1080
      End
      Begin VB.Frame Frame9 
         Appearance      =   0  '���
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  '����
         Caption         =   "Frame9"
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   30
         TabIndex        =   53
         Top             =   555
         Width           =   3585
         Begin VB.CheckBox chkDay 
            BackColor       =   &H00E0E0E0&
            Caption         =   "��"
            Height          =   180
            Index           =   0
            Left            =   30
            TabIndex        =   60
            Tag             =   "Sun"
            Top             =   90
            Value           =   1  'Ȯ��
            Width           =   450
         End
         Begin VB.CheckBox chkDay 
            BackColor       =   &H00E0E0E0&
            Caption         =   "��"
            Height          =   180
            Index           =   1
            Left            =   480
            TabIndex        =   59
            Tag             =   "Mon"
            Top             =   90
            Value           =   1  'Ȯ��
            Width           =   450
         End
         Begin VB.CheckBox chkDay 
            BackColor       =   &H00E0E0E0&
            Caption         =   "ȭ"
            Height          =   180
            Index           =   2
            Left            =   990
            TabIndex        =   58
            Tag             =   "Tue"
            Top             =   90
            Value           =   1  'Ȯ��
            Width           =   480
         End
         Begin VB.CheckBox chkDay 
            BackColor       =   &H00E0E0E0&
            Caption         =   "��"
            Height          =   180
            Index           =   3
            Left            =   1530
            TabIndex        =   57
            Tag             =   "Wed"
            Top             =   90
            Value           =   1  'Ȯ��
            Width           =   465
         End
         Begin VB.CheckBox chkDay 
            BackColor       =   &H00E0E0E0&
            Caption         =   "��"
            Height          =   180
            Index           =   4
            Left            =   2055
            TabIndex        =   56
            Tag             =   "Thu"
            Top             =   90
            Value           =   1  'Ȯ��
            Width           =   465
         End
         Begin VB.CheckBox chkDay 
            BackColor       =   &H00E0E0E0&
            Caption         =   "��"
            Height          =   180
            Index           =   5
            Left            =   2580
            TabIndex        =   55
            Tag             =   "Fri"
            Top             =   90
            Value           =   1  'Ȯ��
            Width           =   465
         End
         Begin VB.CheckBox chkDay 
            BackColor       =   &H00E0E0E0&
            Caption         =   "��"
            Height          =   180
            Index           =   6
            Left            =   3105
            TabIndex        =   54
            Tag             =   "Sat"
            Top             =   90
            Value           =   1  'Ȯ��
            Width           =   465
         End
      End
      Begin VB.Frame Frame8 
         Appearance      =   0  '���
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  '����
         Caption         =   "Frame8"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   30
         TabIndex        =   48
         Top             =   120
         Width           =   3585
         Begin VB.CommandButton cmdDateAdd 
            BackColor       =   &H00EFFCFC&
            Caption         =   "��¥�߰�"
            Height          =   345
            Left            =   2730
            Style           =   1  '�׷���
            TabIndex        =   49
            Top             =   0
            Width           =   840
         End
         Begin MSComCtl2.DTPicker dtpFrConfig 
            Height          =   300
            Left            =   30
            TabIndex        =   50
            Top             =   45
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   393216
            Format          =   79233025
            CurrentDate     =   37935
         End
         Begin MSComCtl2.DTPicker dtpToConfig 
            Height          =   300
            Left            =   1455
            TabIndex        =   51
            Top             =   45
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   393216
            Format          =   79233025
            CurrentDate     =   37935
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "~"
            Height          =   180
            Left            =   1305
            TabIndex        =   52
            Top             =   105
            Width           =   135
         End
      End
      Begin VB.CommandButton cmdDateClear 
         BackColor       =   &H00EFFCFC&
         Caption         =   "��������"
         Height          =   435
         Left            =   1200
         Style           =   1  '�׷���
         TabIndex        =   30
         Top             =   4305
         Width           =   1080
      End
      Begin VB.ListBox lstDate 
         Height          =   3840
         ItemData        =   "frm312QCSchedule_N.frx":0976
         Left            =   2325
         List            =   "frm312QCSchedule_N.frx":098C
         MultiSelect     =   2  'Ȯ����
         Sorted          =   -1  'True
         TabIndex        =   29
         Top             =   960
         Width           =   1275
      End
      Begin MSComCtl2.MonthView mvDate 
         Height          =   2220
         Left            =   30
         TabIndex        =   28
         Top             =   960
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   3916
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   14411494
         Appearance      =   1
         StartOfWeek     =   79233025
         CurrentDate     =   37935
      End
      Begin MedControls1.LisLabel lblDayCnt 
         Height          =   345
         Left            =   1170
         TabIndex        =   62
         Top             =   3390
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   609
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "999"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblTestCnt 
         Height          =   345
         Left            =   1170
         TabIndex        =   63
         Top             =   3795
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   609
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "999"
         Appearance      =   0
      End
   End
   Begin VB.Frame fraConfigTime 
      BackColor       =   &H00DBE6E6&
      Height          =   3765
      Left            =   3720
      TabIndex        =   31
      Top             =   3465
      Width           =   1605
      Begin VB.TextBox txtCnt 
         Alignment       =   1  '������ ����
         Height          =   300
         Left            =   45
         Locked          =   -1  'True
         TabIndex        =   45
         Text            =   "1"
         Top             =   150
         Width           =   555
      End
      Begin VB.CommandButton cmdTimeClear 
         BackColor       =   &H00EFFCFC&
         Caption         =   "����"
         Height          =   330
         Left            =   885
         Style           =   1  '�׷���
         TabIndex        =   35
         Top             =   765
         Width           =   675
      End
      Begin VB.CommandButton cmdTimeAdd 
         BackColor       =   &H00EFFCFC&
         Caption         =   "�߰�"
         Height          =   330
         Left            =   885
         Style           =   1  '�׷���
         TabIndex        =   34
         Top             =   435
         Width           =   675
      End
      Begin VB.ListBox lstTime 
         Height          =   2940
         ItemData        =   "frm312QCSchedule_N.frx":09D8
         Left            =   45
         List            =   "frm312QCSchedule_N.frx":09F4
         Sorted          =   -1  'True
         TabIndex        =   33
         Top             =   780
         Width           =   840
      End
      Begin MSComCtl2.DTPicker dtpTime 
         Height          =   300
         Left            =   45
         TabIndex        =   32
         Top             =   465
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "HH:mm"
         Format          =   79233027
         UpDown          =   -1  'True
         CurrentDate     =   37935
      End
      Begin MSComCtl2.UpDown udCnt 
         Height          =   300
         Left            =   615
         TabIndex        =   46
         Top             =   150
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtCnt"
         BuddyDispid     =   196643
         OrigLeft        =   3705
         OrigTop         =   30
         OrigRight       =   3945
         OrigBottom      =   330
         Max             =   30
         Min             =   1
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "ȸ/��"
         Height          =   180
         Left            =   1005
         TabIndex        =   47
         Top             =   195
         Width           =   450
      End
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   300
      Left            =   75
      TabIndex        =   14
      Top             =   1665
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   529
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
      Caption         =   "�� �˻����� ��ȸ"
      Appearance      =   0
   End
   Begin VB.Frame fraReview 
      BackColor       =   &H00DBE6E6&
      Height          =   1365
      Left            =   75
      TabIndex        =   17
      Top             =   1875
      Width           =   5250
      Begin VB.Frame Frame4 
         Appearance      =   0  '���
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  '����
         Caption         =   "Frame5"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   75
         TabIndex        =   41
         Top             =   585
         Width           =   3750
         Begin VB.CheckBox chkExist 
            BackColor       =   &H00E0E0E0&
            Caption         =   "������ ���� ����"
            Height          =   180
            Index           =   1
            Left            =   1905
            TabIndex        =   43
            Top             =   75
            Width           =   1755
         End
         Begin VB.CheckBox chkExist 
            BackColor       =   &H00E0E0E0&
            Caption         =   "������ �ִ� ����"
            Height          =   180
            Index           =   0
            Left            =   75
            TabIndex        =   42
            Top             =   75
            Value           =   1  'Ȯ��
            Width           =   1770
         End
      End
      Begin VB.CommandButton cmdQReview 
         BackColor       =   &H00FFF2EE&
         Caption         =   "���� ��ȸ"
         Height          =   510
         Left            =   3885
         Style           =   1  '�׷���
         TabIndex        =   25
         Top             =   480
         Width           =   1320
      End
      Begin VB.Frame Frame5 
         Appearance      =   0  '���
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  '����
         Caption         =   "Frame5"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   75
         TabIndex        =   21
         Top             =   975
         Width           =   5115
         Begin VB.CheckBox chkStatus 
            BackColor       =   &H00E0E0E0&
            Caption         =   "ó�����"
            Height          =   180
            Index           =   0
            Left            =   75
            TabIndex        =   44
            Top             =   75
            Value           =   1  'Ȯ��
            Width           =   1050
         End
         Begin VB.CheckBox chkStatus 
            BackColor       =   &H00E0E0E0&
            Caption         =   "��������"
            Height          =   180
            Index           =   1
            Left            =   1395
            TabIndex        =   24
            Top             =   75
            Value           =   1  'Ȯ��
            Width           =   1050
         End
         Begin VB.CheckBox chkStatus 
            BackColor       =   &H00E0E0E0&
            Caption         =   "�κа��"
            Height          =   180
            Index           =   2
            Left            =   2730
            TabIndex        =   23
            Top             =   75
            Value           =   1  'Ȯ��
            Width           =   1050
         End
         Begin VB.CheckBox chkStatus 
            BackColor       =   &H00E0E0E0&
            Caption         =   "�������"
            Height          =   180
            Index           =   3
            Left            =   4005
            TabIndex        =   22
            Top             =   75
            Value           =   1  'Ȯ��
            Width           =   1050
         End
      End
      Begin MSComCtl2.DTPicker dtpFrReview 
         Height          =   375
         Left            =   1440
         TabIndex        =   18
         Top             =   135
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   661
         _Version        =   393216
         Format          =   79233025
         CurrentDate     =   37935
      End
      Begin MSComCtl2.DTPicker dtpToReview 
         Height          =   360
         Left            =   2925
         TabIndex        =   20
         Top             =   135
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   635
         _Version        =   393216
         Format          =   79233025
         CurrentDate     =   37935
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   7
         Left            =   45
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   135
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   635
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "�Ⱓ"
         Appearance      =   0
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "~"
         Height          =   180
         Left            =   2745
         TabIndex        =   19
         Top             =   240
         Width           =   135
      End
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   300
      Index           =   0
      Left            =   5340
      TabIndex        =   39
      Top             =   1665
      Width           =   9120
      _ExtentX        =   16087
      _ExtentY        =   529
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
      Caption         =   "�� �˻����� ����"
      Appearance      =   0
   End
   Begin VB.Frame fraConfigButton 
      BackColor       =   &H00DBE6E6&
      Height          =   1305
      Left            =   3720
      TabIndex        =   36
      Top             =   7155
      Width           =   1605
      Begin VB.CommandButton cmdAllClear 
         BackColor       =   &H00FFF2EE&
         Caption         =   "�ٽü���"
         Height          =   510
         Left            =   150
         Style           =   1  '�׷���
         TabIndex        =   38
         Top             =   135
         Width           =   1320
      End
      Begin VB.CommandButton cmdMake 
         BackColor       =   &H00FFF2EE&
         Caption         =   "�����ۼ�"
         Height          =   510
         Left            =   135
         Style           =   1  '�׷���
         TabIndex        =   37
         Top             =   660
         Width           =   1320
      End
   End
End
Attribute VB_Name = "frm312QCSchedule_N"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event LastFormUnload()

Private mvarParentHwnd As Long

Public Property Let ParentHwnd(ByVal vData As Long)
    mvarParentHwnd = vData
End Property

Public Property Get ParentHwnd() As Long
    ParentHwnd = mvarParentHwnd
End Property

Public Sub CallByExternal(ByVal pCtrlCd As String, ByVal pLevelCd As String)
    txtCtrlCd.Text = ""
    Call InitControl
    Call InitReview
    Call InitConfig
    Call InitConfigDate
    Call InitConfigTime
    
    With tblSchedule
        Call medClearTable(tblSchedule)
        
        .MaxRows = 22
        .RowHeight(-1) = 12
        .Col = 9
        .Row = -1
        .BlockMode = True
        .CellType = CellTypeStaticText
        .TypeHAlign = TypeHAlignCenter
        .TypeVAlign = TypeVAlignCenter
        .BlockMode = False
    End With

    Dim Rs As Recordset
    Dim strSQL As String
    
    Set Rs = GetControlInfo(pCtrlCd, pLevelCd)
            
    txtCtrlCd.Text = Rs.Fields("ctrlcd").Value & ""
    lblCtrlNm.Caption = Rs.Fields("ctrlnm").Value & ""
    
    If Rs.Fields("levelcd").Value & "" = "L" Then
        optLevelcd(0).Value = True
    ElseIf Rs.Fields("levelcd").Value & "" = "N" Then
        optLevelcd(1).Value = True
    ElseIf Rs.Fields("levelcd").Value & "" = "H" Then
        optLevelcd(2).Value = True
    End If
    
    lblCtrlDiv.Caption = IIf(Rs.Fields("ctrldiv").Value & "" = "I", "������������", "�ܺ���������")
    lblEqp.Caption = Format(Rs.Fields("eqpcd").Value & "", "!" & String(5, "@")) & Rs.Fields("eqpnm").Value & ""
    lblEqp.ToolTipText = Format(Rs.Fields("eqpcd").Value & "", "!" & String(5, "@")) & Rs.Fields("eqpnm").Value & ""
    lblBuilding.Caption = Format(Rs.Fields("buildcd").Value & "", "!" & String(5, "@")) & Rs.Fields("buildnm").Value & ""
    lblBuilding.ToolTipText = Format(Rs.Fields("buildcd").Value & "", "!" & String(10, "@")) & Rs.Fields("buildnm").Value & ""
    lblSection.Caption = Format(Rs.Fields("sectcd").Value & "", "!" & String(5, "@")) & Rs.Fields("sectnm").Value & ""
    lblSection.ToolTipText = Format(Rs.Fields("sectcd").Value & "", "!" & String(5, "@")) & Rs.Fields("sectnm").Value & ""
    lblWorkarea.Caption = Format(Rs.Fields("workarea").Value & "", "!" & String(5, "@")) & Rs.Fields("workareanm").Value & ""
    lblWorkarea.ToolTipText = Format(Rs.Fields("workarea").Value & "", "!" & String(5, "@")) & Rs.Fields("workareanm").Value & ""
    
    Set Rs = Nothing
End Sub

Private Sub chkDay_Click(Index As Integer)
    Dim i As Long, j As Long
    Dim strlstDay As String
    Dim dtDate As Date
    Dim strTmp As String
    Dim aryTmp() As String
    
    On Error Resume Next
    If Screen.ActiveControl.Name <> chkDay(Index).Name Then Exit Sub
    
    MousePointer = vbHourglass
    
    If chkDay(Index).Value = 0 Then '������ ���� ����
        For j = 0 To lstDate.ListCount - 1
            strlstDay = lstDate.List(j)
            
            If Weekday(CDate(strlstDay)) = Index + 1 Then
                strTmp = strTmp & j & COL_DIV
'                lstDate.Selected(j) = True
            End If
        Next
        
        aryTmp() = Split(strTmp, COL_DIV)

        For i = UBound(aryTmp) To LBound(aryTmp) Step -1
            If aryTmp(i) <> "" Then
                lstDate.RemoveItem Val(aryTmp(i))
            End If
        Next
    Else    '������ ���� �߰�
        dtDate = Format(dtpFrConfig.Value, "yyyy-MM-dd")
        Do Until dtDate = Format(DateAdd("d", 1, dtpToConfig.Value), "yyyy-MM-dd")
            If medListFind(lstDate, dtDate) < 0 Then
                If Weekday(dtDate) = Index + 1 Then
                    lstDate.addItem Format(dtDate, "yyyy-MM-dd")
                End If
            End If
            dtDate = Format(DateAdd("d", 1, dtDate), "yyyy-MM-dd")
        Loop
    End If
    
    lblDayCnt.Caption = lstDate.ListCount
    lblTestCnt.Caption = lstDate.ListCount * Val(txtCnt.Text)
    
    MousePointer = vbDefault
End Sub

Private Sub chkExist_Click(Index As Integer)
    If Screen.ActiveControl.Name <> chkExist(Index).Name Then Exit Sub
    
    If chkExist(0).Value = 0 And chkExist(1).Value = 0 Then
        chkExist(IIf(Index = 0, 1, 0)).Value = 1
    End If
End Sub

Private Sub chkStatus_Click(Index As Integer)
    If Screen.ActiveControl.Name <> chkStatus(Index).Name Then Exit Sub
    
    '��� �ϳ��� �����ؾ� ������ �����Լ��� ����Ͽ� �ƹ��ų� �����ϵ��� �Ѵ�.
    
    Randomize
    
    If chkStatus(0).Value = 0 And chkStatus(1).Value = 0 And chkStatus(2).Value = 0 And chkStatus(3).Value = 0 Then
'        If chkExist(0).Value = 1 And chkExist(1).Value = 1 Then
            chkStatus(Int((3 - 0 + 1) * Rnd + 0)).Value = 1
'        ElseIf chkExist(0).Value = 1 And chkExist(1).Value = 0 Then
'            chkStatus(Int((3 - 0 + 1) * Rnd + 0)).Value = 1
'        ElseIf chkExist(0).Value = 0 And chkExist(1).Value = 1 Then
'            chkStatus(Int((3 - 1 + 1) * Rnd + 1)).Value = 1
'        End If
    End If
End Sub

Private Sub cmdAllClear_Click()
    Dim i As Long
    
    If lstDate.ListCount = 0 Then Exit Sub
    If lstTime.ListCount = 0 Then Exit Sub
    
    For i = chkDay.LBound To chkDay.UBound
        chkDay(i).Value = 0
    Next
    
    lblDayCnt.Caption = ""
    lblTestCnt.Caption = ""
    
    txtCnt.Text = "1"

    Call InitConfigDate
    Call InitConfigTime
    Call DeleteStandBy
End Sub

Private Sub DeleteStandBy()
'��� ������ �ڷ� ����
    Dim i As Long

    With tblSchedule
        .ReDraw = False
        For i = .MaxRows To 1 Step -1
            .Row = i
            .Col = 13
            If .Value = "" Then
                .Action = ActionDeleteRow
            End If
        Next
        
        .MaxRows = .DataRowCnt
        If .MaxRows < 22 Then
            .MaxRows = 22
        End If
        .RowHeight(-1) = 12
        .ReDraw = True
    End With
End Sub

Private Sub cmdClear_Click()
    txtCtrlCd.Text = ""
    Call InitControl
    Call InitReview
    Call InitConfig
    Call InitConfigDate
    Call InitConfigTime
    
    With tblSchedule
        Call medClearTable(tblSchedule)
        
        .MaxRows = 22
        .RowHeight(-1) = 12
        .Col = 9
        .Row = -1
        .BlockMode = True
        .CellType = CellTypeStaticText
        .TypeHAlign = TypeHAlignCenter
        .TypeVAlign = TypeVAlignCenter
        .BlockMode = False
    End With
End Sub

Private Sub cmdDateAdd_Click()
    Dim dtDate As Date
    Dim i As Long
    
    MousePointer = vbHourglass
    
    For i = chkDay.LBound To chkDay.UBound
        chkDay(i).Value = 1
    Next
    
    DoEvents
    lstDate.Clear
    dtDate = Format(dtpFrConfig.Value, "yyyy-MM-dd")
    Do Until dtDate = Format(DateAdd("d", 1, dtpToConfig.Value), "yyyy-MM-dd")
        If medListFind(lstDate, dtDate) < 0 Then
            lstDate.addItem Format(dtDate, "yyyy-MM-dd")
        End If
        dtDate = Format(DateAdd("d", 1, dtDate), "yyyy-MM-dd")
    Loop
    
    lblDayCnt.Caption = lstDate.ListCount
    lblTestCnt.Caption = lstDate.ListCount * Val(txtCnt.Text)
    
    MousePointer = vbDefault
    
    If tblSchedule.DataRowCnt > 0 Then
        If Trim(txtCtrlCd.Text) = "" Then Exit Sub
        DoEvents
        Call LoadSchedule
    End If
End Sub

Private Sub cmdDateClear_Click()
    Dim i As Long
    
    lstDate.Clear
    
    For i = chkDay.LBound To chkDay.UBound
        chkDay(i).Value = 0
    Next
    
    lblDayCnt.Caption = lstDate.ListCount
    lblTestCnt.Caption = lstDate.ListCount * Val(txtCnt.Text)
End Sub

Private Sub cmdExit_Click()
    Unload Me
'    Unload frm311QCResultEntry_N
'    Unload frm309QCOrder_N
    If IsLastForm Then RaiseEvent LastFormUnload
    If IsLastForm Then Call UnloadForm(Me)
'    If IsLastForm Then
'        If mvarParentHwnd <> 0 Then
'            Call SendMessage(mvarParentHwnd, WM_CLOSE, 0&, 0&)
'        End If
'    End If
End Sub

Private Sub cmdMake_Click()
    Dim strMsg As VbMsgBoxResult
    
    If Trim(txtCtrlCd.Text) = "" Then
        MsgBox "��Ʈ���� �����ϰų� �Է��Ͻʽÿ�.", vbExclamation
        Exit Sub
    End If
    
    If lblCtrlNm.Caption = "" Then
        MsgBox "��Ʈ���� �����ϰų� �Է��Ͻʽÿ�.", vbExclamation
        Exit Sub
    End If
    
    If lstDate.ListCount = 0 Then
        MsgBox "������ ��¥�� �����ϴ�. ��¥�� �������ֽʽÿ�.", vbExclamation
        Exit Sub
    End If
    
    If lstTime.ListCount = 0 Then
        MsgBox "������ �ð��� �����ϴ�. �ð��� �������ֽʽÿ�.", vbExclamation
        Exit Sub
    End If
    
    If txtCnt.Text > lstTime.ListCount Then
        MsgBox "������ 1�ϴ� Ƚ���� ������ �ð�Ƚ���� �ٸ��ϴ�. �����ð��� �߰��Ͻʽÿ�.", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("���� �ڷḦ ��ȸ�Ͻðڽ��ϱ�?", vbExclamation + vbYesNo + vbDefaultButton2) = vbYes Then
        chkExist(0).Value = 1
        chkStatus(0).Value = 1
        chkStatus(1).Value = 1
        chkStatus(2).Value = 1
        chkStatus(3).Value = 1
        
        Call LoadSchedule
    End If
    
    Call CreateSchedule(IIf(optLevelcd(0).Value, "L", IIf(optLevelcd(1).Value, "N", "H")))
End Sub

Private Sub CreateSchedule(ByVal pLevel As String)
    Dim objPro As clsProgress
    Dim i As Long, j As Long, k As Long
    Dim strLevelcd As String, strDate As String, strTime As String
    
    Set objPro = Nothing
    Set objPro = New clsProgress
    
    With objPro
        .Appearance = ccFlat
        .Container = Me
        .Width = tblSchedule.Width
        .Left = tblSchedule.Left
        .Top = tblSchedule.Top
        .Height = 430
        .ForeColor = &H864B24
        .Message = "������ �ۼ��ϰ� �ֽ��ϴ�..."
        .Max = lstDate.ListCount * lstTime.ListCount
        .Value = 1
    End With
    
'������� �ڷ� ����
    Call DeleteStandBy
    
'��ȸ�� �ڷῡ ������ ��¥�� �ƴ� �ٸ� ��¥�� ���� �ִ� ��� ó��
    If ProcessNotExistDate = False Then GoTo Nodata
    
'��ȸ�� �ڷῡ ������ �ð��� �ƴ� �ٸ� �ð��� ���� �ִ� ��� ó��
    If ProcessNotExistTime = False Then GoTo Nodata
    
    tblSchedule.ReDraw = False
    For i = 0 To lstDate.ListCount - 1
        For j = 0 To lstTime.ListCount - 1
            k = k + 1
            
            With tblSchedule
                strLevelcd = IIf(pLevel = "L", "Low", IIf(pLevel = "N", "Normal", "High"))
                strDate = lstDate.List(i)
                strTime = lstTime.List(j)
                
                If CheckDup(strLevelcd & strDate & strTime) = False Then
                    If .DataRowCnt >= .MaxRows Then
                        .MaxRows = .MaxRows + 1
                        .RowHeight(-1) = 12
                    End If
                    .Row = .DataRowCnt + 1
                    
                    .Col = 1: .Value = strLevelcd
                    .Col = 2: .Value = strDate
                    .Col = 3: .Value = strTime
                    .Col = 7: .Value = "���": .ForeColor = vbBlue
                    .Col = 8: .Value = "Y"
'                    .Col = 10: .Value = "��": .ForeColor = DCM_LightRed
                    .Col = 12: .Value = strLevelcd & strDate & strTime
                End If
            End With
            
            objPro.Value = k
        Next
    Next
    
    tblSchedule.SortBy = SortByRow
    tblSchedule.SortKey(1) = 2
    tblSchedule.SortKey(2) = 3
    tblSchedule.SortKeyOrder(1) = SortKeyOrderAscending
    tblSchedule.SortKeyOrder(2) = SortKeyOrderAscending
    tblSchedule.Col = 1: tblSchedule.Col2 = tblSchedule.MaxCols
    tblSchedule.Row = 1: tblSchedule.Row2 = tblSchedule.MaxRows
    tblSchedule.Action = ActionSort
    
    tblSchedule.ReDraw = True
    
Nodata:
    Set objPro = Nothing
End Sub

Private Function ProcessNotExistDate() As Boolean
'��ȸ�� �ڷῡ ������ ��¥�� �ƴ� �ٸ� ��¥�� ���� �ִ� ��� ó��
    Dim strLstDate As String
    Dim strMsg As VbMsgBoxResult
    Dim blnExists As Boolean
    Dim i As Long
    
    ProcessNotExistDate = False
    
    blnExists = True
    
    For i = 0 To lstDate.ListCount - 1
        strLstDate = strLstDate & lstDate.List(i) & COL_DIV
    Next
    
    With tblSchedule
        For i = 1 To .DataRowCnt
            .Row = i
            .Col = 8
            If .Value = "Y" Then
                .Col = 2
                If InStr(strLstDate, .Value) = 0 Then '������ ��¥�� ���� ��쿩..
                    blnExists = False
                    Exit For
                End If
            End If
        Next
    End With
    
    If blnExists = False Then
        strMsg = MsgBox("��ȸ�� �ڷ��߿� ���� ������ ��¥�� ���� �ڷᰡ �����մϴ�." & vbNewLine & _
                        "�� �ڷḦ �״�� ����Ϸ��� ���� ������ ��¥�� �߰��ؾ� �մϴ�." & vbNewLine & vbNewLine & _
                        "��� �����Ͻðڽ��ϱ�?" & vbNewLine & _
                        "(��:������¥�� ������¥�� �߰�, �ƴϿ�:������¥�� ���� �ڷ����)", vbExclamation + vbYesNoCancel + vbDefaultButton2)
        
        If strMsg = vbCancel Then '�ƹ����� ���ϱ� ��������
            Exit Function
        End If
        
        If strMsg = vbYes Then  '������¥�� ������¥�� �߰�
            With tblSchedule
                For i = 1 To .DataRowCnt
                    .Row = i
                    .Col = 8
                    If .Value = "Y" Then
                        .Col = 2
                        If InStr(strLstDate, .Value) = 0 Then '������ ��¥�� ���� ��쿩..
                            
                            Call LoadDate(.Value)
                        End If
                    End If
                Next
            End With
        End If
        
        If strMsg = vbNo Then '������¥�� ���� �ڷ����
            With tblSchedule
                .ReDraw = False
                For i = .MaxRows To 1 Step -1
                    .Row = i
                    .Col = 8
                    If .Value = "Y" Then
                        .Col = 2
                        If InStr(strLstDate, .Value) = 0 Then '������ ��¥�� ���� ��쿩..
                            .Action = ActionDeleteRow
                        End If
                    End If
                Next
                
                .MaxRows = tblSchedule.DataRowCnt
                If .MaxRows < 22 Then
                    .MaxRows = 22
                End If
                .RowHeight(-1) = 12
                .ReDraw = True
            End With
        End If
    End If
    
    ProcessNotExistDate = True
End Function

Private Function ProcessNotExistTime() As Boolean
'��ȸ�� �ڷῡ ������ �ð��� �ƴ� �ٸ� �ð��� ���� �ִ� ��� ó��
    Dim strLstTime As String
    Dim strMsg As VbMsgBoxResult
    Dim blnExists As Boolean
    Dim i As Long
    Dim strLevelcd As String, strDate As String, strTime As String
    
    ProcessNotExistTime = False
    
    blnExists = True
    
    For i = 0 To lstTime.ListCount - 1
        strLstTime = strLstTime & lstTime.List(i) & COL_DIV
    Next
    
    With tblSchedule
        For i = 1 To .DataRowCnt
            .Row = i
            .Col = 8
            If .Value = "Y" Then
                .Col = 3
                If InStr(strLstTime, .Value) = 0 Then '������ �ð��� ���� ��쿩..
                    blnExists = False
                    Exit For
                End If
            End If
        Next
    End With
    
    If blnExists = False Then
        strMsg = MsgBox("��ȸ�� �ڷ��߿� ���� ������ �ð��� ���� �ڷᰡ �����մϴ�." & vbNewLine & _
                        "�� �ڷḦ �״�� ����Ϸ��� ���� ������ �ð����� �����ؾ� �մϴ�." & vbNewLine & vbNewLine & _
                        "��� �����Ͻðڽ��ϱ�?" & vbNewLine & _
                        "(��:�����ð��� ����ð����� ����, �ƴϿ�:�����ð��� ���� �ڷ����)", vbExclamation + vbYesNoCancel)
        
        If strMsg = vbCancel Then '�ƹ����� ���ϱ� ��������
            Exit Function
        End If
        
        If strMsg = vbYes Then  '�����ð������ϰ� ����ð����� ����
            With tblSchedule
                .ReDraw = False
                For i = 1 To .DataRowCnt
                    .Row = i
                    .Col = 8
                    If .Value = "Y" Then
                        .Col = 3
                        If InStr(strLstTime, .Value) = 0 Then '������ �ð��� ���� ��쿩..
                            .Value = lstTime.List(0)
                            
                            .Col = 1: strLevelcd = .Value
                            .Col = 2: strDate = .Value
                            .Col = 3: strTime = .Value
                            .Col = 12: .Value = strLevelcd & strDate & strTime
                        End If
                    End If
                Next
                .ReDraw = True
            End With
            
            Dim strRtnRow As String
            '�ߺ��� �ڷḦ ã�Ƽ� ����..
            With tblSchedule
                .ReDraw = False
                For i = .MaxRows To 1 Step -1
                    .Row = i
                    .Col = 12
                    strRtnRow = i
                    If CheckDup(.Value, strRtnRow) Then '�ߺ��� ���� ����
                        .Row = strRtnRow
                        .Action = ActionDeleteRow
                    End If
                Next

                .MaxRows = .DataRowCnt
                If .MaxRows < 22 Then
                    .MaxRows = 22
                End If
                .RowHeight(-1) = 12
                .ReDraw = True
            End With
            
        End If
        
        If strMsg = vbNo Then   '�����ð��� ���� �ڷ� ����
            With tblSchedule
                .ReDraw = False
                For i = .MaxRows To 1 Step -1
                    .Row = i
                    .Col = 8
                    If .Value = "Y" Then
                        .Col = 3
                        If InStr(strLstTime, .Value) = 0 Then '������ �ð��� ���� ��쿩..
                            .Action = ActionDeleteRow
                        End If
                    End If
                Next
                
                .MaxRows = tblSchedule.DataRowCnt
                If .MaxRows < 22 Then
                    .MaxRows = 22
                End If
                .RowHeight(-1) = 12
                .ReDraw = True
            End With
        End If
    End If
    
    ProcessNotExistTime = True
End Function

Private Function CheckDup(ByVal pKeyString As String, Optional ByRef pRow As String = "") As Boolean
    Dim i As Long
    
    CheckDup = False
    If pKeyString = "" Then Exit Function
    
'    DoEvents
    For i = tblSchedule.MaxRows To 1 Step -1
        tblSchedule.Row = i
        If pRow = "" Then
            tblSchedule.Col = 12
            If tblSchedule.Value = pKeyString Then
                CheckDup = True
                Exit For
            End If
        ElseIf pRow <> "" Then
            If tblSchedule.Row = pRow Then  '�������� �ǹ� ���� �ο� üũ, �ش� �ο�� �ߺ��˻縦 �ϸ� �ȵǹǷ�
            
            Else
                tblSchedule.Col = 12
                If tblSchedule.Value = pKeyString Then
                    CheckDup = True
                    pRow = i
                    Exit For
                End If
            End If
        End If
    Next
    
End Function

Private Sub cmdPopCtrl_Click()
    If lblCtrlNm.Caption <> "" Then
        DoEvents
        Call InitControl
        Call InitReview
        Call InitConfig
        Call InitConfigDate
        Call InitConfigTime
        
        With tblSchedule
            Call medClearTable(tblSchedule)
            
            .MaxRows = 22
            .RowHeight(-1) = 12
            .Col = 9
            .Row = -1
            .BlockMode = True
            .CellType = CellTypeStaticText
            .TypeHAlign = TypeHAlignCenter
            .TypeVAlign = TypeVAlignCenter
            .BlockMode = False
        End With
    End If
    
    DoEvents
    Call LoadControlInfo
'    DoEvents
'    Call LoadLotNo
'    DoEvents
'    Call LoadTestItem
End Sub

Private Sub LoadControlInfo(Optional ByVal pCtrlCd As String = "")
'��Ʈ���� �Ϲ� ������ �ҷ��´�..
    Dim objPop As clsPopUpList
    Dim i As Long
    
    Set objPop = New clsPopUpList

    With objPop
        .Recordset = GetControlInfo(pCtrlCd)
        .FormCaption = "��Ʈ�� ã��"
        .Delimiter = COL_DIV
        .FormWidth = 4470
        .ColumnHeaderText = "�ڵ���Ʈ�Ѹ�Level��������ڵ������ǹ��ڵ��ǹ��������ڵ����Ǹ���ũ�ַ��ڵ���ũ�ַ���"
        .ColumnHeaderWidth = "854.92922475.213629.8583000000000"
        .ColumnHeaderAlign = "002"
        '0 ����, 1 ������, 2 ���
        
        Call .LoadPopUp
        
        DoEvents
        
        txtCtrlCd.Text = medGetP(.SelectedString, 1, .Delimiter)
        lblCtrlNm.Caption = medGetP(.SelectedString, 2, .Delimiter)
        
        If medGetP(.SelectedString, 3, .Delimiter) = "L" Then
            optLevelcd(0).Value = True
        ElseIf medGetP(.SelectedString, 3, .Delimiter) = "N" Then
            optLevelcd(1).Value = True
        ElseIf medGetP(.SelectedString, 3, .Delimiter) = "H" Then
            optLevelcd(2).Value = True
        End If
        
        lblCtrlDiv.Caption = IIf(medGetP(.SelectedString, 4, .Delimiter) = "I", "������������", "�ܺ���������")
        lblEqp.Caption = Format(medGetP(.SelectedString, 5, .Delimiter), "!" & String(5, "@")) & medGetP(.SelectedString, 6, .Delimiter)
        lblEqp.ToolTipText = Format(medGetP(.SelectedString, 5, .Delimiter), "!" & String(10, "@")) & medGetP(.SelectedString, 6, .Delimiter)
        lblBuilding.Caption = Format(medGetP(.SelectedString, 7, .Delimiter), "!" & String(5, "@")) & medGetP(.SelectedString, 8, .Delimiter)
        lblBuilding.ToolTipText = Format(medGetP(.SelectedString, 7, .Delimiter), "!" & String(10, "@")) & medGetP(.SelectedString, 8, .Delimiter)
        lblSection.Caption = Format(medGetP(.SelectedString, 9, .Delimiter), "!" & String(5, "@")) & medGetP(.SelectedString, 10, .Delimiter)
        lblSection.ToolTipText = Format(medGetP(.SelectedString, 9, .Delimiter), "!" & String(10, "@")) & medGetP(.SelectedString, 10, .Delimiter)
        lblWorkarea.Caption = Format(medGetP(.SelectedString, 11, .Delimiter), "!" & String(5, "@")) & medGetP(.SelectedString, 12, .Delimiter)
        lblWorkarea.ToolTipText = Format(medGetP(.SelectedString, 11, .Delimiter), "!" & String(10, "@")) & medGetP(.SelectedString, 12, .Delimiter)
    End With
    
    Set objPop = Nothing
End Sub

Private Function GetControlInfo(Optional ByVal pCtrlCd As String = "", _
                                Optional ByVal pLevelCd As String = "") As Recordset
    Dim strSQL As String
    
    strSQL = " select a.ctrlcd,a.ctrlnm,a.levelcd,a.ctrldiv,a.eqpcd,c.eqpnm, a.buildcd,d.field1 as buildnm, " & _
            " a.sectcd,e.field1 as sectnm, a.workarea, f.field1 as workareanm " & _
            " from " & T_LAB021 & " a, " & T_LAB006 & " c, " & T_LAB032 & " d, " & T_LAB032 & " e, " & T_LAB032 & " f " & _
            " where " & DBJ("a.eqpcd*=c.eqpcd") & _
            " and " & DBW("d.cdindex=", LC3_Buildings) & _
            " and a.buildcd=d.cdval1 " & _
            " and " & DBW("e.cdindex=", LC3_Section) & _
            " and a.sectcd=e.cdval1 " & _
            " and " & DBW("f.cdindex=", LC3_WorkArea) & _
            " and a.workarea=f.cdval1 "

    If pCtrlCd <> "" Then
        strSQL = strSQL & " and " & DBW("a.ctrlcd=", pCtrlCd)
    End If
    
    If pLevelCd <> "" Then
        strSQL = strSQL & " and " & DBW("a.levelcd=", pLevelCd)
    End If
    
    strSQL = strSQL & " order by a.ctrlcd,ctrlnm,levelcd"
            
    Set GetControlInfo = New Recordset
    GetControlInfo.Open strSQL, DBConn
End Function

Private Sub cmdPrint_Click()
    MsgBox "���Ŀ� ������ ����Դϴ�.", vbInformation
End Sub

Private Sub cmdQReview_Click()
    Dim i As Long
    
    If Trim(txtCtrlCd.Text) = "" Then Exit Sub
    
    DoEvents
    For i = chkDay.LBound To chkDay.UBound
        chkDay(i).Value = 0
    Next
    
    lblDayCnt.Caption = ""
    lblTestCnt.Caption = ""
    
    txtCnt.Text = "1"

    Call InitConfigDate
    Call InitConfigTime
    
    DoEvents
    Call LoadSchedule
End Sub

Private Sub LoadSchedule()
    Dim objPro As clsProgress
    Dim Rs As Recordset
    Dim strLevelcd As String
    Dim strDate As String
    Dim strTime As String
    Dim i As Long
    
    MousePointer = vbHourglass
    
    Set objPro = Nothing
    Set objPro = New clsProgress
    
    With objPro
        .Container = Me
        .Width = tblSchedule.Width
        .Left = tblSchedule.Left
        .Top = tblSchedule.Top
        .Height = 430
'        .ForeColor = &H864B24
        .Message = "�ڷḦ �а� �ֽ��ϴ�..."
        
'        .Choice = True
'        .Appearance = aPlate
'        .SetMyForm Me
'        .XWidth = tblSchedule.Width
'        .XPos = tblSchedule.Left
'        .YPos = tblSchedule.Top
'        .YHeight = 430
'        .ForeColor = &H864B24
'        .Msg = "�ڷḦ �а� �ֽ��ϴ�..."
'        .Value = 1
    End With
    
    Set Rs = GetSQL
    
    objPro.Max = Rs.RecordCount
    
    With tblSchedule
        Call medClearTable(tblSchedule)
        
        .MaxRows = 22
        .RowHeight(-1) = 12
        .Col = 9
        .Row = -1
        .BlockMode = True
        .CellType = CellTypeStaticText
        .TypeHAlign = TypeHAlignCenter
        .TypeVAlign = TypeVAlignCenter
        .BlockMode = False
    End With
    
    With tblSchedule
        .ReDraw = False
        
        Do Until Rs.EOF
            i = i + 1
            
            If .DataRowCnt >= .MaxRows Then
                .MaxRows = .MaxRows + 1
                .RowHeight(-1) = 12
            End If
            
            .Row = .DataRowCnt + 1
            
            .Col = 1: .Value = IIf(Rs.Fields("levelcd").Value & "" = "L", "Low", IIf(Rs.Fields("levelcd").Value & "" = "N", "Normal", IIf(Rs.Fields("levelcd").Value & "" = "H", "High", ""))): strLevelcd = .Value
            .Col = 2: .Value = Format(Rs.Fields("dodt").Value & "", "0###-##-##"): strDate = .Value
            .Col = 3: .Value = Format(Mid(Rs.Fields("dotm").Value & "", 1, 4), "0#:##"): strTime = .Value
            .Col = 4: .Value = Format(Rs.Fields("rcvdt").Value & "", "0###-##-##")
            .Col = 5: .Value = Format(Mid(Rs.Fields("rcvtm").Value & "", 1, 4), "0#:##")
            .Col = 6: .Value = IIf(Rs.Fields("workarea").Value & "" = "", "", Rs.Fields("workarea").Value & "" & "-" & Mid(Rs.Fields("accdt").Value & "", 3) & "-" & Rs.Fields("accseq").Value & "")
            .Col = 7: .Value = GetStatus(Rs.Fields("stscd").Value & "")
            .Col = 8: .Value = Rs.Fields("flag").Value & ""
            .Col = 11: .Value = IIf(Rs.Fields("spcyy").Value & "" = "", "", Rs.Fields("spcyy").Value & "" & "-" & Rs.Fields("spcno").Value & "")
            .Col = 12: .Value = strLevelcd & strDate & strTime
            .Col = 13: .Value = "Y"
            
            .Col = 7
            If .Value = "ó��" Then
                .Col = 10
                .Value = "��"
                .ForeColor = DCM_LightRed
            End If
            .Col = 7
            If .Value = "����" Or .Value = "�κ�" Then
                .Col = 10
                .Value = "��"
                .ForeColor = DCM_LightBlue
            End If
            
            .Col = 8
            If .Value = "Y" Then
                .Col = 7
                If .Value = "ó��" Then
                    .Col = 9
                    .CellType = CellTypeStaticText
                    .TypeHAlign = TypeHAlignCenter
                    .TypeVAlign = TypeVAlignCenter
                Else
                    .Col = 9
                    .CellType = CellTypeCheckBox
                    .TypeHAlign = TypeHAlignCenter
                    .TypeVAlign = TypeVAlignCenter
                    .Value = 1
                End If
                
                .Col = 2
                Call LoadDate(.Value)
                .Col = 3
                Call LoadTime(.Value)
            End If
            
            objPro.Value = i
            Rs.MoveNext
        Loop
        
        .ReDraw = True
    End With
    
    MousePointer = vbDefault
    
    If Not ((Screen.ActiveControl.Name = cmdDateAdd.Name) Or (Screen.ActiveControl.Name = cmdMake.Name)) Then
        If tblSchedule.DataRowCnt = 0 Then
            MsgBox "�ڷᰡ �������� �ʽ��ϴ�.", vbExclamation
        End If
    End If
        
    Set Rs = Nothing
    Set objPro = Nothing
End Sub

Private Function GetSQL() As Recordset
    Dim strSQL As String
    Dim strSQL1 As String
    Dim strSQL2 As String
    Dim strSQL3 As String
    Dim strSQL4 As String
    
    'ó��,����,�κ�,����
    If chkStatus(1).Value = 1 And chkStatus(2).Value = 1 And chkStatus(3).Value = 1 Then
        strSQL4 = ""
    'ó��,����,�κ�
    ElseIf chkStatus(1).Value = 1 And chkStatus(2).Value = 1 And chkStatus(3).Value = 0 Then
        strSQL4 = " and c.stscd in ('2','3','4') "
    'ó��,����,����
    ElseIf chkStatus(1).Value = 1 And chkStatus(2).Value = 0 And chkStatus(3).Value = 1 Then
        strSQL4 = " and c.stscd in ('2','5') "
    'ó��,�κ�,����
    ElseIf chkStatus(1).Value = 0 And chkStatus(2).Value = 1 And chkStatus(3).Value = 1 Then
        strSQL4 = " and c.stscd in ('3','4','5') "
    'ó��,����
    ElseIf chkStatus(1).Value = 1 And chkStatus(2).Value = 0 And chkStatus(3).Value = 0 Then
        strSQL4 = " and c.stscd in ('2') "
    'ó��,�κ�
    ElseIf chkStatus(1).Value = 0 And chkStatus(2).Value = 1 And chkStatus(3).Value = 0 Then
        strSQL4 = " and c.stscd in ('3','4') "
    'ó��,����
    ElseIf chkStatus(1).Value = 0 And chkStatus(2).Value = 0 And chkStatus(3).Value = 1 Then
        strSQL4 = " and c.stscd in ('5') "
    'ó��
    ElseIf chkStatus(1).Value = 0 And chkStatus(2).Value = 0 And chkStatus(3).Value = 0 Then
        strSQL4 = " and c.stscd in ('1') "
    End If

    '�����ٿ� �ִ� ���
    'ó�� ����
    strSQL1 = " select a.ctrlcd,a.levelcd,a.dodt,a.dotm,'' as rcvdt, '' as rcvtm,'' as workarea, '' as accdt, 0 as accseq, " & _
            " '1' as stscd, '' as spcyy, 0 as spcno,'Y' as flag " & _
            " from " & T_LAB025 & " a, " & T_LAB021 & " b " & _
            " where " & DBW("a.dodt>=", Format(dtpFrReview.Value, "yyyyMMdd")) & _
            " and " & DBW("a.dodt<=", Format(dtpToReview.Value, "yyyyMMdd")) & _
            " and " & DBW("a.ctrlcd=", Trim(txtCtrlCd.Text)) & _
            " and a.levelcd in (" & IIf(optLevelcd(0).Value, "'L'", IIf(optLevelcd(1).Value, "'N'", "'H'")) & ") " & _
            " and (a.donefg ='' or a.donefg is null) " & _
            " and b.ctrlcd = a.ctrlcd " & _
            " and b.levelcd = a.levelcd "
    '�����ٿ� �ִ� ���, ���°� 2,3,4,5,6�� ���
    strSQL2 = " select a.ctrlcd,a.levelcd,a.dodt,a.dotm,c.rcvdt,c.rcvtm,c.workarea,c.accdt,c.accseq, " & _
            " c.stscd,c.spcyy,c.spcno,'Y' as flag " & _
            " from " & T_LAB025 & " a, " & T_LAB021 & " b, " & T_LAB201 & " c " & _
            " where " & DBW("a.dodt>=", Format(dtpFrReview.Value, "yyyyMMdd")) & _
            " and " & DBW("a.dodt<=", Format(dtpToReview.Value, "yyyyMMdd")) & _
            " and " & DBW("a.ctrlcd=", Trim(txtCtrlCd.Text)) & _
            " and a.levelcd in (" & IIf(optLevelcd(0).Value, "'L'", IIf(optLevelcd(1).Value, "'N'", "'H'")) & ") " & _
            " and (a.donefg <>'' or a.donefg is not null) " & _
            " and b.ctrlcd = a.ctrlcd " & _
            " and b.levelcd = a.levelcd " & _
            " and a.spcyy=c.spcyy " & _
            " and a.spcno=c.spcno " & strSQL4 & _
            " and not exists (select * from " & T_LAB025 & _
            "                 where dodt = a.dodt " & _
            "                 and ctrlcd =a.ctrlcd " & _
            "                 and levelcd=a.levelcd " & _
            "                 and (donefg ='' or donefg is null)) "
    '�����ٿ� ���� ���
    '���°� 2,3,4,5,6 �� ���
    strSQL3 = " select distinct a.ctrlcd,a.levelcd,c.rcvdt as dodt,c.rcvtm as dotm, " & _
            " c.rcvdt,c.rcvtm,c.workarea,c.accdt,c.accseq,c.stscd,c.spcyy,c.spcno, '' as flag " & _
            " from " & T_LAB026 & " a, " & T_LAB201 & " c " & _
            " where " & DBW("a.ctrlcd=", Trim(txtCtrlCd.Text)) & _
            " and a.levelcd in (" & IIf(optLevelcd(0).Value, "'L'", IIf(optLevelcd(1).Value, "'N'", "'H'")) & ") " & _
            " and " & DBW("c.rcvdt>=", Format(dtpFrReview.Value, "yyyyMMdd")) & _
            " and " & DBW("c.rcvdt<=", Format(dtpToReview.Value, "yyyyMMdd")) & _
            " and a.workarea=c.workarea " & _
            " and a.accdt=c.accdt " & _
            " and a.accseq=c.accseq " & strSQL4 & _
            " and not exists (select * from " & T_LAB025 & _
            "                 where " & DBW("dodt>=", Format(dtpFrReview.Value, "yyyyMMdd")) & _
            "                 and " & DBW("dodt<=", Format(dtpToReview.Value, "yyyyMMdd")) & _
            "                 and " & DBW("ctrlcd=", Trim(txtCtrlCd.Text)) & _
            "                 and levelcd in (" & IIf(optLevelcd(0).Value, "'L'", IIf(optLevelcd(1).Value, "'N'", "'H'")) & ") " & _
            "                 and spcyy=c.spcyy " & _
            "                 and spcno=c.spcno) "
    'order by dodt,dotm,workarea,accdt,accseq
    
    If chkExist(0).Value = 1 And chkExist(1).Value = 1 Then '�Ѵ� �����Ѱ��
        If chkStatus(0).Value = 1 And chkStatus(1).Value = 0 And chkStatus(2).Value = 0 And chkStatus(3).Value = 0 Then
        'ó���γѸ� ����, strsql1�� ���
            strSQL = strSQL1 & " order by levelcd,dodt,dotm "
        ElseIf chkStatus(0).Value = 1 Then  '�ٸ��� ������� ó���� ���õ� ���
            strSQL = strSQL1 & " union " & strSQL2 & " union " & strSQL3 & " order by levelcd,dodt,dotm "
        ElseIf chkStatus(0).Value = 0 Then '�ٸ��� ������� ó���� ���õ��� ���� ���
            strSQL = strSQL2 & " union " & strSQL3 & " order by levelcd,dodt,dotm "
        End If
        
'        ElseIf chkStatus(0).Value = 0 And chkStatus(1).Value = 1 And chkStatus(2).Value = 0 And chkStatus(3).Value = 0 Then
'        '�����γѸ� ����, strsql2, strsql3�� ���
'            strSQL = strSQL2 & " union " & strSQL3 & " order by levelcd,dodt,dotm "
'        ElseIf chkStatus(0).Value = 0 And chkStatus(1).Value = 0 And chkStatus(2).Value = 1 And chkStatus(3).Value = 0 Then
'        '�κ��γѸ� ����,strsql2,strsql3�� ���
'            strSQL = strSQL2 & " union " & strSQL3 & " order by levelcd,dodt,dotm "
'        ElseIf chkStatus(0).Value = 0 And chkStatus(1).Value = 0 And chkStatus(2).Value = 0 And chkStatus(3).Value = 1 Then
'        '�����γѸ� ����
'            strSQL = strSQL2 & " union " & strSQL3 & " order by levelcd,dodt,dotm "
'
'
'        ElseIf chkStatus(0).Value = 1 And chkStatus(1).Value = 1 And chkStatus(2).Value = 0 And chkStatus(3).Value = 0 Then
'        'ó��,�����γѸ� ����
'            strSQL = strSQL1 & " union " & strSQL2 & " union " & strSQL3 & " order by levelcd,dodt,dotm "
'        ElseIf chkStatus(0).Value = 1 And chkStatus(1).Value = 0 And chkStatus(2).Value = 1 And chkStatus(3).Value = 0 Then
'        'ó��,�κ��γѸ� ����
'            strSQL = strSQL1 & " union " & strSQL2 & " union " & strSQL3 & " order by levelcd,dodt,dotm "
'        ElseIf chkStatus(0).Value = 1 And chkStatus(1).Value = 0 And chkStatus(2).Value = 0 And chkStatus(3).Value = 1 Then
'        'ó��,�����γѸ� ����
'            strSQL = strSQL1 & " union " & strSQL2 & " union " & strSQL3 & " order by levelcd,dodt,dotm "
'        'ó��,����,�κе� ���� ����, ó��,����,������ ���� ����
        
    
    ElseIf chkExist(0).Value = 1 And chkExist(1).Value = 0 Then '������ �ִ� �Ѹ�
        If chkStatus(0).Value = 1 And chkStatus(1).Value = 0 And chkStatus(2).Value = 0 And chkStatus(3).Value = 0 Then
        'ó���γѸ� ����, strsql1�� ���
            strSQL = strSQL1 & " order by levelcd,dodt,dotm "
        ElseIf chkStatus(0).Value = 1 Then  '�ٸ��� ������� ó���� ���õ� ���
            strSQL = strSQL1 & " union " & strSQL2 & " order by levelcd,dodt,dotm "
        ElseIf chkStatus(0).Value = 0 Then '�ٸ��� ������� ó���� ���õ��� ���� ���
            strSQL = strSQL2 & " order by levelcd,dodt,dotm "
        End If
    ElseIf chkExist(0).Value = 0 And chkExist(1).Value = 1 Then '������ ���� �Ѹ�
        strSQL = strSQL3 & " order by levelcd,dodt,dotm "
    End If
    
    Set GetSQL = New Recordset
    GetSQL.Open strSQL, DBConn
End Function

Private Function GetStatus(ByVal pStsCd As String) As String
'����
    With tblSchedule
        .Row = .DataRowCnt
        .Col = 7
        If pStsCd = "1" Then
            GetStatus = "ó��"
            .ForeColor = DCM_LightRed
        ElseIf pStsCd = "2" Then
            GetStatus = "����"
            .ForeColor = vbBlack
        ElseIf pStsCd = "3" Or pStsCd = "4" Then
            GetStatus = "�κ�"
            .ForeColor = DCM_LightBlue
        ElseIf pStsCd = "5" Then
            GetStatus = "����"
            .ForeColor = DCM_Green
        ElseIf pStsCd = "6" Then
            GetStatus = "����"
            .ForeColor = vbRed
        End If
    End With
End Function

Private Sub LoadDate(ByVal pDate As String)
    Dim i As Long
    Dim dtDate As Date
    Dim aryDtpWeekCnt(6) As String
    Dim aryLstWeekCnt(6) As String
    
    If medListFind(lstDate, pDate) < 0 Then
        lstDate.addItem pDate
    End If
    
    lblDayCnt.Caption = lstDate.ListCount
    lblTestCnt.Caption = lstDate.ListCount * Val(txtCnt.Text)
    
    dtDate = Format(dtpFrConfig.Value, "yyyy-MM-dd")
    Do Until dtDate = Format(DateAdd("d", 1, dtpToConfig.Value), "yyyy-MM-dd")
        aryDtpWeekCnt(Weekday(dtDate) - 1) = Val(aryDtpWeekCnt(Weekday(dtDate) - 1)) + 1
        
        dtDate = Format(DateAdd("d", 1, dtDate), "yyyy-MM-dd")
    Loop
    
    For i = 0 To lstDate.ListCount - 1
        dtDate = CDate(lstDate.List(i))
        aryLstWeekCnt(Weekday(dtDate) - 1) = Val(aryLstWeekCnt(Weekday(dtDate) - 1)) + 1
    Next
    
    For i = 0 To 6
        If aryDtpWeekCnt(i) = aryLstWeekCnt(i) Then
            chkDay(i).Value = 1
        ElseIf aryLstWeekCnt(i) = "" Then
            chkDay(i).Value = 0
        Else
            chkDay(i).Value = 2
        End If
    Next
End Sub

Private Sub LoadTime(ByVal pTime As String)
    If medListFind(lstTime, pTime) < 0 Then
        lstTime.addItem pTime
    End If
    
    txtCnt.Text = lstTime.ListCount
End Sub

Private Sub cmdSave_Click()
    If Trim(txtCtrlCd.Text) = "" Then
        MsgBox "��Ʈ���� �����ϰų� �Է��Ͻʽÿ�.", vbExclamation
        Exit Sub
    End If
    
    If tblSchedule.DataRowCnt = 0 Then
        MsgBox "���� ������ �ۼ��Ͻʽÿ�.", vbExclamation
        Exit Sub
    End If
    
    Call SaveSchedule
End Sub

Private Sub SaveSchedule()
    Dim objPro As clsProgress
    Dim Rs As Recordset
    Dim strMsg As VbMsgBoxResult
    Dim strSQL As String
    Dim strLevel As String
    Dim aryLevel() As String
    Dim arySQL() As String
    Dim i As Long, j As Long
    
    Dim strLevelcd As String
    Dim strDoDt As String
    Dim strDoTm As String
    Dim strDoneFg As String
    Dim strSpcYY As String
    Dim strSpcNo As String
    Dim strSchedule As String
    Dim strDel As String
    
    strMsg = MsgBox("�ۼ��� ������ �ش� ��Ʈ���� �������� �����մϴ�." & vbNewLine & _
            "����� �ش� ��Ʈ���� ��� Level�� ���� ������ ������ �� �ֽ��ϴ�." & vbNewLine & _
            "��� Level�� ������ ��� ������ Level �̿��� Level�� ���� ���������� ���� ���� �ֽ��ϴ�." & vbNewLine & _
            """�ƴϿ�""�� �����Ͽ� �� Level���� ������ ���� �����մϴ�." & vbNewLine & vbNewLine & _
            "��� �����Ͻðڽ��ϱ�?" & vbNewLine & _
            "(��:��� Level�� ����, �ƴϿ�:������ Level�� ����)", vbExclamation + vbYesNoCancel + vbDefaultButton2)

    If strMsg = vbCancel Then Exit Sub
    
    '���α׷����� ó��
    Set objPro = Nothing
    Set objPro = New clsProgress
    
    With objPro
        .Container = Me
        .Width = tblSchedule.Width
        .Left = tblSchedule.Left
        .Top = tblSchedule.Top
        .Height = 430
        .ForeColor = &H864B24
        .Message = "������ �����ϱ� ���� �ڷḦ �а� �ֽ��ϴ�..."
        .Max = tblSchedule.DataRowCnt
        .Value = 1
    End With
    
    ReDim arySQL(0)
    
    If strMsg = vbYes Then
        strSQL = " select * from " & T_LAB021 & _
                 " where " & DBW("ctrlcd=", Trim(txtCtrlCd.Text))
        
        Set Rs = New Recordset
        Rs.Open strSQL, DBConn
        
        Do Until Rs.EOF
            strLevel = strLevel & Rs.Fields("levelcd").Value & "" & COL_DIV
            Rs.MoveNext
        Loop
        
        Set Rs = Nothing
        
        arySQL(0) = " delete from " & T_LAB025 & " " & _
                    " where " & DBW("dodt>=", Format(dtpFrReview.Value, "yyyyMMdd")) & _
                    " and " & DBW("dodt<=", Format(dtpToReview.Value, "yyyyMMdd")) & _
                    " and " & DBW("ctrlcd=", Trim(txtCtrlCd.Text))
    End If
    
    If strMsg = vbNo Then
        strLevel = IIf(optLevelcd(0).Value, "L", IIf(optLevelcd(1).Value, "N", "H")) & COL_DIV
    
        arySQL(0) = " delete from " & T_LAB025 & " " & _
                    " where " & DBW("dodt>=", Format(dtpFrReview.Value, "yyyyMMdd")) & _
                    " and " & DBW("dodt<=", Format(dtpToReview.Value, "yyyyMMdd")) & _
                    " and " & DBW("ctrlcd=", Trim(txtCtrlCd.Text)) & _
                    " and levelcd in (" & IIf(optLevelcd(0).Value, "'L'", IIf(optLevelcd(1).Value, "'N'", "'H'")) & ") "
    End If
    
    aryLevel = Split(strLevel, COL_DIV)
    
    With tblSchedule
        For i = 1 To .DataRowCnt
            .Row = i
            .Col = 1: strLevelcd = Mid(.Value, 1, 1)
            .Col = 2: strDoDt = Format(.Value, "yyyyMMdd")
            .Col = 3: strDoTm = Format(.Value, "hhMM") & "00"
            .Col = 7: strDoneFg = IIf(.Value = "���" Or .Value = "ó��", "", "1")
            .Col = 8: strSchedule = .Value
            .Col = 9: strDel = .Value
            .Col = 11: strSpcYY = medGetP(.Value, 1, "-")
                       strSpcNo = medGetP(.Value, 2, "-")
                                   
            If strSchedule = "Y" And (strDel = "" Or strDel = "0") Then
                For j = LBound(aryLevel) To UBound(aryLevel)
                    If aryLevel(j) <> "" Then
                        If aryLevel(j) = strLevelcd Then
                            ReDim Preserve arySQL(UBound(arySQL) + 1)
                            arySQL(UBound(arySQL)) = "insert into " & T_LAB025 & _
                            "(dodt,dotm,sectcd,ctrlcd,levelcd,donefg,spcyy,spcno) values (" & _
                            DBV("dodt", strDoDt, 1) & DBV("dotm", strDoTm, 1) & _
                            DBV("sectcd", Trim(Mid(lblSection.Caption, 1, 5)), 1) & DBV("ctrlcd", Trim(txtCtrlCd.Text), 1) & _
                            DBV("levelcd", aryLevel(j), 1) & DBV("donefg", strDoneFg, 1) & _
                            DBV("spcyy", strSpcYY, 1) & DBV("spcno", strSpcNo) & ")"
                        Else
                            ReDim Preserve arySQL(UBound(arySQL) + 1)
                            arySQL(UBound(arySQL)) = "insert into " & T_LAB025 & _
                            "(dodt,dotm,sectcd,ctrlcd,levelcd,donefg,spcyy,spcno) values (" & _
                            DBV("dodt", strDoDt, 1) & DBV("dotm", strDoTm, 1) & _
                            DBV("sectcd", Trim(Mid(lblSection.Caption, 1, 5)), 1) & DBV("ctrlcd", Trim(txtCtrlCd.Text), 1) & _
                            DBV("levelcd", aryLevel(j), 1) & DBV("donefg", "", 1) & _
                            DBV("spcyy", "", 1) & DBV("spcno", "") & ")"
                        End If
                    End If
                Next j
            End If
            
            objPro.Value = i
        Next i
    End With
    
    Set objPro = Nothing
    
On Error GoTo ErrTrap

    Set objPro = Nothing
    Set objPro = New clsProgress
    
    With objPro
        .Container = Me
        .Width = tblSchedule.Width
        .Left = tblSchedule.Left
        .Top = tblSchedule.Top
        .Height = 430
        .ForeColor = 430
        .Message = "������ �����ϰ� �ֽ��ϴ�..."
        .Max = UBound(arySQL) - 1
        .Value = 1
    End With
    
    DBConn.BeginTrans
    For i = LBound(arySQL) To UBound(arySQL)
        If arySQL(i) <> "" Then
'            Debug.Print arySQL(i)
            DBConn.Execute arySQL(i)
        End If
        
        objPro.Value = i + 1
    Next
    DBConn.CommitTrans
    
    Set objPro = Nothing
    
    MousePointer = vbDefault
    
    Call cmdClear_Click
    MsgBox "���������� ó���Ǿ����ϴ�.", vbInformation
    
    Exit Sub
    
ErrTrap:
    MousePointer = vbDefault
    
    DBConn.RollbackTrans
    MsgBox "ó������ ������ �߻��Ͽ����ϴ�." & vbNewLine & Err.Description, vbCritical
    
    Set objPro = Nothing
End Sub

Private Sub cmdSelDateClear_Click()
    Dim strTmp As String
    Dim aryTmp() As String
    Dim i As Long
    Dim dtDate As Date
    Dim aryDtpWeekCnt(6) As String
    Dim aryLstWeekCnt(6) As String
    
    MousePointer = vbHourglass
    
    For i = 0 To lstDate.ListCount - 1
        If lstDate.Selected(i) Then
            strTmp = strTmp & i & COL_DIV
        End If
    Next
    
    aryTmp = Split(strTmp, COL_DIV)
    
    For i = UBound(aryTmp) To LBound(aryTmp) Step -1
        If aryTmp(i) <> "" Then
            lstDate.RemoveItem Val(aryTmp(i))
        End If
    Next
    
    lblDayCnt.Caption = lstDate.ListCount
    lblTestCnt.Caption = lstDate.ListCount * Val(txtCnt.Text)
    
    '����ǥ�� üũ���� ������ ���Ͽ�...
    '�ش� ������ ���� ���õǾ� ������ 1 �Ϻμ��õǾ� ������ 2 ���� ���õ��� �ʾ����� 0
    
    dtDate = Format(dtpFrConfig.Value, "yyyy-MM-dd")
    Do Until dtDate = Format(DateAdd("d", 1, dtpToConfig.Value), "yyyy-MM-dd")
        aryDtpWeekCnt(Weekday(dtDate) - 1) = Val(aryDtpWeekCnt(Weekday(dtDate) - 1)) + 1
        
        dtDate = Format(DateAdd("d", 1, dtDate), "yyyy-MM-dd")
    Loop
    
    For i = 0 To lstDate.ListCount - 1
        dtDate = CDate(lstDate.List(i))
        aryLstWeekCnt(Weekday(dtDate) - 1) = Val(aryLstWeekCnt(Weekday(dtDate) - 1)) + 1
    Next
    
    For i = 0 To 6
        If aryDtpWeekCnt(i) = aryLstWeekCnt(i) Then
            chkDay(i).Value = 1
        ElseIf aryLstWeekCnt(i) = "" Then
            chkDay(i).Value = 0
        Else
            chkDay(i).Value = 2
        End If
    Next
    
    MousePointer = vbDefault
End Sub

Private Sub cmdTimeAdd_Click()
    If Val(txtCnt.Text) = lstTime.ListCount Then
        MsgBox "�� �̻� �˻�ð��� �߰��� �� �����ϴ�.", vbExclamation
        Exit Sub
    End If
    
    If medListFind(lstTime, Format(dtpTime.Value, "HH:mm")) >= 0 Then
        MsgBox "�̹� �߰��� �˻�Ⱓ�Դϴ�.", vbExclamation
        Exit Sub
    Else
        lstTime.addItem Format(dtpTime.Value, "HH:mm")
    End If
End Sub

Private Sub cmdTimeClear_Click()
    lstTime.Clear
End Sub

Private Sub dtpFrConfig_Change()
    On Error Resume Next
    If Screen.ActiveControl.Name <> dtpFrConfig.Name Then Exit Sub
    
    dtpFrReview.Value = dtpFrConfig.Value
End Sub

Private Sub dtpFrReview_Change()
    On Error Resume Next
    If Screen.ActiveControl.Name <> dtpFrReview.Name Then Exit Sub
    
    dtpFrConfig.Value = dtpFrReview.Value
End Sub

Private Sub dtpToConfig_Change()
    On Error Resume Next
    If Screen.ActiveControl.Name <> dtpToConfig.Name Then Exit Sub
    
    dtpToReview.Value = dtpToConfig.Value
End Sub

Private Sub dtpToReview_Change()
    On Error Resume Next
    If Screen.ActiveControl.Name <> dtpToReview.Name Then Exit Sub
    
    dtpToConfig.Value = dtpToReview.Value
End Sub

Private Sub Form_Load()
    
    txtCtrlCd.Text = ""
    Call InitControl
    Call InitReview
    Call InitConfig
    Call InitConfigDate
    Call InitConfigTime
    
    dtpToConfig.Value = GetSystemDate
    dtpFrConfig.Value = DateAdd("m", -1, dtpToConfig.Value)
    dtpToReview.Value = GetSystemDate
    dtpFrReview.Value = DateAdd("m", -1, dtpToReview.Value)
    mvDate.Value = GetSystemDate
    dtpTime.Value = GetSystemDate
    
    With tblSchedule
        Call medClearTable(tblSchedule)
        
        .MaxRows = 22
        .RowHeight(-1) = 12
        .Col = 9
        .Row = -1
        .BlockMode = True
        .CellType = CellTypeStaticText
        .TypeHAlign = TypeHAlignCenter
        .TypeVAlign = TypeVAlignCenter
        .BlockMode = False
    End With
End Sub

Private Sub InitControl()
    lblCtrlNm.Caption = ""
    lblCtrlDiv.Caption = ""
    lblEqp.Caption = ""
    lblBuilding.Caption = ""
    lblSection.Caption = ""
    lblWorkarea.Caption = ""
End Sub

Private Sub InitReview()
'    dtpToReview.Value = GetSystemDate
'    dtpFrReview.Value = DateAdd("m", -1, dtpToReview.Value)
End Sub

Private Sub InitConfig()
    Dim i As Long
    
    txtCnt.Text = "1"
'    dtpToConfig.Value = GetSystemDate
'    dtpFrConfig.Value = DateAdd("m", -1, dtpToConfig.Value)
    
    For i = chkDay.LBound To chkDay.UBound
        chkDay(i).Value = 0
    Next
    
    lblDayCnt.Caption = ""
    lblTestCnt.Caption = ""
End Sub

Private Sub InitConfigDate()
'    mvDate.Value = GetSystemDate
    lstDate.Clear
End Sub

Private Sub InitConfigTime()
'    dtpTime.Value = GetSystemDate
    lstTime.Clear
End Sub

Private Sub mvDate_DateClick(ByVal DateClicked As Date)
    Dim lngDtpWeekCnt As Long
    Dim lngLstWeekCnt As Long
    Dim dtDate As Date
    Dim i As Long
        
    If Format(DateClicked, "yyyy-MM-dd") < Format(dtpFrConfig.Value, "yyyy-MM-dd") Then
        MsgBox "��¥ ���� ������ ������ϴ�.", vbExclamation
        Exit Sub
    End If
    
    If Format(DateClicked, "yyyy-MM-dd") > Format(dtpToConfig.Value, "yyyy-MM-dd") Then
        MsgBox "��¥ ���� ������ ������ϴ�.", vbExclamation
        Exit Sub
    End If
    
    MousePointer = vbHourglass
    
    If medListFind(lstDate, Format(DateClicked, "yyyy-MM-dd")) < 0 Then
        lstDate.addItem Format(DateClicked, "yyyy-MM-dd")
        
        lblDayCnt.Caption = lstDate.ListCount
        lblTestCnt.Caption = lstDate.ListCount * Val(txtCnt.Text)
        
        dtDate = Format(dtpFrConfig.Value, "yyyy-MM-dd")
        Do Until dtDate = Format(DateAdd("d", 1, dtpToConfig.Value), "yyyy-MM-dd")
            If Weekday(dtDate) = Weekday(DateClicked) Then
                lngDtpWeekCnt = lngDtpWeekCnt + 1
            End If
            dtDate = Format(DateAdd("d", 1, dtDate), "yyyy-MM-dd")
        Loop
        
        For i = 0 To lstDate.ListCount - 1
            dtDate = CDate(lstDate.List(i))
            If Weekday(dtDate) = Weekday(DateClicked) Then
                lngLstWeekCnt = lngLstWeekCnt + 1
            End If
        Next
        
        If lngDtpWeekCnt = lngLstWeekCnt Then
            chkDay(Weekday(DateClicked) - 1).Value = 1
        Else
            chkDay(Weekday(DateClicked) - 1).Value = 2
        End If
        
    End If
    
    MousePointer = vbDefault
End Sub

Private Sub optLevelcd_Click(Index As Integer)
    On Error Resume Next
    If Screen.ActiveControl.Name <> optLevelcd(Index).Name Then Exit Sub
    If Trim(txtCtrlCd.Text) = "" Then Exit Sub
    
    Call InitReview
    Call InitConfig
    Call InitConfigDate
    Call InitConfigTime
    
    With tblSchedule
        Call medClearTable(tblSchedule)
        
        .MaxRows = 22
        .RowHeight(-1) = 12
        .Col = 9
        .Row = -1
        .BlockMode = True
        .CellType = CellTypeStaticText
        .TypeHAlign = TypeHAlignCenter
        .TypeVAlign = TypeVAlignCenter
        .BlockMode = False
    End With
    
    Call LoadLevel
End Sub

Private Sub LoadLevel()
    Dim Rs As Recordset
    Dim strSQL As String
    
    strSQL = " select * from " & T_LAB021 & _
             " where " & DBW("ctrlcd=", Trim(txtCtrlCd.Text)) & _
             " and levelcd in (" & IIf(optLevelcd(0).Value, "'L'", IIf(optLevelcd(1).Value, "'N'", IIf(optLevelcd(2).Value, "'H'", "'L','N','H'"))) & ") " & _
             " order by ctrlnm "
    
    Set Rs = New Recordset
    Rs.Open strSQL, DBConn
    
    If Rs.EOF Then
        MsgBox "�ش� ��Ʈ���� �������� �ʽ��ϴ�.", vbExclamation
        txtCtrlCd.Text = ""
        Call InitControl
    Else
        lblCtrlNm.Caption = ""
        Do Until Rs.EOF
            lblCtrlNm.Caption = lblCtrlNm.Caption & Rs.Fields("ctrlnm").Value & "" & ","
        
            Rs.MoveNext
        Loop
        
        lblCtrlNm.Caption = Mid(lblCtrlNm.Caption, 1, Len(lblCtrlNm.Caption) - 1)
    End If
    
    Set Rs = Nothing
End Sub

Private Sub tblSchedule_Click(ByVal Col As Long, ByVal Row As Long)
    Static blnToggle As Boolean
    Dim i As Long
    Dim frm As Form
    Dim frmExist As Boolean
    
    If tblSchedule.DataRowCnt = 0 Then Exit Sub
    
    If Col = 9 And Row = 0 Then
        blnToggle = IIf(blnToggle, False, True)
        
        With tblSchedule
            .Col = Col
            For i = 1 To .DataRowCnt
                .Row = i
                If .CellType = CellTypeCheckBox Then
                    .Value = IIf(blnToggle, 0, 1)
                End If
            Next
        End With
    End If
    
    If Col = 10 And Row <> 0 Then
        tblSchedule.Col = 10
        tblSchedule.Row = Row
        If tblSchedule.Value = "��" Then  '������������ ȭ���� ����ش�.
            Dim strWorkArea As String
            Dim strAccDt As String
            Dim strAccSeq As String
            
            tblSchedule.Col = 6
            strWorkArea = medGetP(tblSchedule.Value, 1, "-")
            strAccDt = medGetP(tblSchedule.Value, 2, "-")
            strAccSeq = medGetP(tblSchedule.Value, 3, "-")
            
            Call LoadForm(frm311QCResultEntry_N, Me)
            
'            frm311QCResultEntry_N.ParentHwnd = GetAncestor(Me.hwnd, 1)
'
'            frmExist = False
'            For Each frm In Forms
'                If frm.Name = frm311QCResultEntry_N.Name Then
'                    frmExist = True
'                End If
'            Next
''            Unload frm311QCResultEntry_N
'
'            '���� �����ϴ� �� ã� ������ zorder�� 0���� ���ְ� ������ ����..
'            DoEvents
'            If frmExist = False Then
'                Call SetParent(frm311QCResultEntry_N.hwnd, frm311QCResultEntry_N.ParentHwnd)
'                frm311QCResultEntry_N.WindowState = 2
'                frm311QCResultEntry_N.Show
'            End If
'            frm311QCResultEntry_N.ZOrder 0
'
            Call frm311QCResultEntry_N.CallByExternal(strWorkArea & "-" & strAccDt & "-" & strAccSeq)
        ElseIf tblSchedule.Value = "��" Then
            Call LoadForm(frm309QCOrder_N, Me)
            
'            frm309QCOrder_N.ParentHwnd = GetAncestor(Me.hwnd, 1)
'
'            frmExist = False
'            For Each frm In Forms
'                If frm.Name = frm309QCOrder_N.Name Then
'                    frmExist = True
'                End If
'            Next
''            Unload frm309QCOrder_N
'
'            '���� �����ϴ� �� ã� ������ zorder�� 0���� ���ְ� ������ ����..
'            DoEvents
'            If frmExist = False Then
'                Call SetParent(frm309QCOrder_N.hwnd, frm309QCOrder_N.ParentHwnd)
'                frm309QCOrder_N.WindowState = 2
'                frm309QCOrder_N.Show
'            End If
'
'            frm309QCOrder_N.ZOrder 0
            
            Call frm309QCOrder_N.CallByExternal(Trim(txtCtrlCd.Text), IIf(optLevelcd(0).Value, "L", IIf(optLevelcd(1).Value, "N", "H")))
        End If
    End If
End Sub

Private Sub txtCnt_Change()
    On Error Resume Next
    If Screen.ActiveControl.Name <> txtCnt.Name Then Exit Sub
    
    lstTime.Clear
    lblTestCnt.Caption = lstDate.ListCount * Val(txtCnt.Text)
End Sub

Private Sub txtCtrlCd_Change()
    On Error Resume Next
    If Screen.ActiveControl.Name <> txtCtrlCd.Name Then Exit Sub
    
    If lblCtrlNm.Caption <> "" Then
        Call InitControl
        Call InitReview
        Call InitConfig
        Call InitConfigDate
        Call InitConfigTime
        
        With tblSchedule
            Call medClearTable(tblSchedule)
            
            .MaxRows = 22
            .RowHeight(-1) = 12
            .Col = 9
            .Row = -1
            .BlockMode = True
            .CellType = CellTypeStaticText
            .TypeHAlign = TypeHAlignCenter
            .TypeVAlign = TypeVAlignCenter
            .BlockMode = False
        End With
    End If
End Sub

Private Sub txtCtrlCd_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtCtrlCd_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(txtCtrlCd.Text) = "" Then Exit Sub
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtCtrlCd_LostFocus()
    Dim Rs As Recordset
'�̵����� �ۿ� ���ұ�? ���߿� �ٸ� ������� ���ľ���...

    If Trim(txtCtrlCd.Text) = "" Then Exit Sub
    If Trim(lblCtrlNm.Caption) <> "" Then Exit Sub
    
    DoEvents
    Set Rs = GetControlInfo(Trim(txtCtrlCd.Text))
    
    If Rs.EOF = False Then
        DoEvents
        Call LoadControlInfo(Trim(txtCtrlCd.Text))
'        DoEvents
'        Call LoadLotNo
'        DoEvents
'        Call LoadTestItem
    End If
    
    Set Rs = Nothing
End Sub

Private Sub udCnt_Change()
    lblTestCnt.Caption = lstDate.ListCount * Val(txtCnt.Text)
End Sub


