VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmInterface 
   Caption         =   "epoc Enterprise Data Manager ����͸�"
   ClientHeight    =   12195
   ClientLeft      =   345
   ClientTop       =   840
   ClientWidth     =   18825
   BeginProperty Font 
      Name            =   "����ü"
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
   Picture         =   "frmInterface.frx":1272
   ScaleHeight     =   12195
   ScaleWidth      =   18825
   StartUpPosition =   1  '������ ���
   WindowState     =   2  '�ִ�ȭ
   Begin VB.PictureBox Picture2 
      Align           =   1  '�� ����
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   18765
      TabIndex        =   37
      Top             =   525
      Width           =   18825
      Begin VB.OptionButton optVer 
         Caption         =   "2015"
         Height          =   345
         Index           =   0
         Left            =   10650
         TabIndex        =   72
         Top             =   30
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.OptionButton optVer 
         Caption         =   "2016"
         Height          =   345
         Index           =   1
         Left            =   10650
         TabIndex        =   71
         Top             =   300
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton Command16 
         Caption         =   "�����׽�Ʈ"
         Height          =   435
         Left            =   21960
         TabIndex        =   68
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtTest 
         Height          =   435
         Left            =   23190
         MultiLine       =   -1  'True
         TabIndex        =   67
         Top             =   120
         Visible         =   0   'False
         Width           =   3225
      End
      Begin VB.CommandButton cmdResult 
         Appearance      =   0  '���
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   13590
         TabIndex        =   63
         Top             =   120
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.CommandButton cmdRsltSearch 
         Caption         =   "�����ȸ"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   14820
         TabIndex        =   62
         Top             =   120
         Width           =   1155
      End
      Begin VB.CommandButton cmdExcelExport 
         Caption         =   "Excel"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   20850
         TabIndex        =   61
         Top             =   90
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.CommandButton cmdIFClear 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   17280
         TabIndex        =   60
         Top             =   120
         Width           =   1155
      End
      Begin VB.CommandButton cmdIFTrans 
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   16050
         TabIndex        =   59
         Top             =   120
         Width           =   1155
      End
      Begin VB.Frame fraWork 
         Height          =   735
         Left            =   90
         TabIndex        =   51
         Top             =   -60
         Width           =   10125
         Begin VB.CheckBox chkSave 
            Caption         =   "��������"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   8430
            TabIndex        =   77
            Top             =   210
            Width           =   1245
         End
         Begin VB.CheckBox chkSearch 
            Caption         =   "�����ȸ"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   7050
            TabIndex        =   76
            Top             =   210
            Value           =   1  'Ȯ��
            Width           =   1245
         End
         Begin VB.CommandButton cmdReadEpoc 
            Appearance      =   0  '���
            Caption         =   "����б�"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   4620
            TabIndex        =   75
            Top             =   180
            Width           =   1155
         End
         Begin VB.TextBox txtTimer 
            Alignment       =   2  '��� ����
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   5850
            TabIndex        =   74
            Text            =   "60"
            Top             =   180
            Width           =   1035
         End
         Begin VB.CommandButton cmdOrder 
            Appearance      =   0  '���
            Caption         =   "��������"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   11160
            TabIndex        =   70
            Top             =   150
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.CheckBox chkSaveAll 
            Caption         =   "��������"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   13260
            TabIndex        =   69
            Top             =   180
            Value           =   1  'Ȯ��
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.TextBox txtPos 
            Alignment       =   2  '��� ����
            Appearance      =   0  '���
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   10650
            TabIndex        =   65
            Text            =   "A"
            Top             =   180
            Width           =   315
         End
         Begin VB.TextBox txtRack 
            Alignment       =   2  '��� ����
            Appearance      =   0  '���
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   10290
            TabIndex        =   64
            Text            =   "5"
            Top             =   180
            Width           =   315
         End
         Begin VB.CommandButton cmdPatDelete 
            Caption         =   "��������"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   10230
            TabIndex        =   54
            Top             =   150
            Width           =   1185
         End
         Begin VB.CommandButton cmdSearch 
            Caption         =   "��ũ��ȸ"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   10230
            TabIndex        =   53
            Top             =   150
            Width           =   1185
         End
         Begin VB.ComboBox cboChk 
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmInterface.frx":14F5
            Left            =   10230
            List            =   "frmInterface.frx":1502
            TabIndex        =   52
            Top             =   240
            Visible         =   0   'False
            Width           =   975
         End
         Begin MSComCtl2.DTPicker dtpStopDt 
            Height          =   345
            Left            =   2820
            TabIndex        =   55
            Top             =   240
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����ü"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   126615553
            CurrentDate     =   40248
         End
         Begin MSComCtl2.DTPicker dtpStartDt 
            Height          =   345
            Left            =   1170
            TabIndex        =   56
            Top             =   240
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����ü"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   126615553
            CurrentDate     =   40248
         End
         Begin VB.Label Label12 
            Appearance      =   0  '���
            BackColor       =   &H80000005&
            BackStyle       =   0  '����
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2640
            TabIndex        =   58
            Top             =   330
            Width           =   105
         End
         Begin VB.Label Label20 
            Caption         =   "��ȸ����"
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
            Left            =   180
            TabIndex        =   57
            Top             =   300
            Width           =   915
         End
      End
      Begin VB.Label Label2 
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   8940
         TabIndex        =   73
         Top             =   150
         Visible         =   0   'False
         Width           =   525
      End
   End
   Begin VB.Frame Frame1 
      Height          =   10245
      Left            =   90
      TabIndex        =   26
      Top             =   1290
      Width           =   18495
      Begin VB.CommandButton cmdSL 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����ü"
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
         Top             =   210
         Visible         =   0   'False
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
         Height          =   9975
         Left            =   90
         TabIndex        =   30
         Top             =   180
         Width           =   11535
         _Version        =   393216
         _ExtentX        =   20346
         _ExtentY        =   17595
         _StockProps     =   64
         ButtonDrawMode  =   4
         ColHeaderDisplay=   0
         ColsFrozen      =   16
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         MaxCols         =   17
         MaxRows         =   20
         MoveActiveOnFocus=   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "frmInterface.frx":151A
      End
      Begin VB.Frame Frame6 
         Height          =   585
         Left            =   11700
         TabIndex        =   31
         Top             =   120
         Width           =   6675
         Begin VB.TextBox txtBarNum 
            Alignment       =   2  '��� ����
            BeginProperty Font 
               Name            =   "����ü"
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
            TabIndex        =   39
            Top             =   150
            Visible         =   0   'False
            Width           =   1875
         End
         Begin VB.CommandButton cmdBarInput 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   38
            Top             =   180
            Visible         =   0   'False
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
            Caption         =   "ȯ�ڸ� :"
            BeginProperty Font 
               Name            =   "����ü"
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
            Caption         =   "���ڵ��ȣ :"
            BeginProperty Font 
               Name            =   "����ü"
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
         Height          =   9390
         Left            =   11700
         TabIndex        =   29
         Top             =   750
         Width           =   6645
         _Version        =   393216
         _ExtentX        =   11721
         _ExtentY        =   16563
         _StockProps     =   64
         BackColorStyle  =   1
         ColHeaderDisplay=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         MaxCols         =   8
         MaxRows         =   10
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "frmInterface.frx":21A7
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  '�� ����
      BackColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   18765
      TabIndex        =   18
      Top             =   0
      Width           =   18825
      Begin VB.Timer tmrResult 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   6180
         Top             =   30
      End
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
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   126615552
         CurrentDate     =   40457
      End
      Begin VB.Label Label1 
         Appearance      =   0  '���
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '����
         Caption         =   "epoc EDM"
         BeginProperty Font 
            Name            =   "����"
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
         TabIndex        =   45
         Top             =   90
         Width           =   1380
      End
      Begin VB.Label Label1 
         Appearance      =   0  '���
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '����
         Caption         =   "�˻�����"
         BeginProperty Font 
            Name            =   "����"
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
         Appearance      =   0  '���
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '����
         Caption         =   "epoc EDM"
         BeginProperty Font 
            Name            =   "����"
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
         Width           =   1380
      End
      Begin VB.Image imgPort 
         Height          =   240
         Left            =   13560
         Picture         =   "frmInterface.frx":2855
         Top             =   90
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgSend 
         Height          =   240
         Left            =   14715
         Picture         =   "frmInterface.frx":2DDF
         Top             =   90
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgReceive 
         Height          =   240
         Left            =   15870
         Picture         =   "frmInterface.frx":3369
         Top             =   90
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "��Ʈ"
         Height          =   195
         Index           =   0
         Left            =   13050
         TabIndex        =   21
         Top             =   120
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "�۽�"
         Height          =   195
         Left            =   14235
         TabIndex        =   20
         Top             =   120
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "����"
         Height          =   195
         Left            =   15360
         TabIndex        =   19
         Top             =   120
         Visible         =   0   'False
         Width           =   420
      End
   End
   Begin VB.Frame FrmHideControl 
      Caption         =   "HideControl"
      Height          =   9945
      Left            =   17610
      TabIndex        =   1
      Top             =   1350
      Visible         =   0   'False
      Width           =   4815
      Begin VB.OptionButton optSaveResult 
         Caption         =   "���"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   1110
         TabIndex        =   47
         Top             =   5160
         Width           =   735
      End
      Begin VB.OptionButton optSaveResult 
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   1890
         TabIndex        =   46
         Top             =   5160
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.Frame Frame2 
         Caption         =   "Error Log"
         Height          =   945
         Left            =   180
         TabIndex        =   43
         Top             =   8190
         Width           =   4530
         Begin VB.TextBox txtErrLog 
            Appearance      =   0  '���
            Height          =   615
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  '����
            TabIndex        =   44
            Top             =   240
            Width           =   4275
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Print"
         Height          =   2415
         Left            =   180
         TabIndex        =   40
         Top             =   5670
         Width           =   3045
         Begin FPSpread.vaSpread vasPrint 
            Height          =   1035
            Left            =   120
            TabIndex        =   41
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
            SpreadDesigner  =   "frmInterface.frx":38F3
         End
         Begin FPSpread.vaSpread vasPrintBuf 
            Height          =   975
            Left            =   120
            TabIndex        =   42
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
            SpreadDesigner  =   "frmInterface.frx":53B9
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
         Style           =   1  '�׷���
         TabIndex        =   23
         Top             =   3210
         Value           =   1  'Ȯ��
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
         SpreadDesigner  =   "frmInterface.frx":561E
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
         SpreadDesigner  =   "frmInterface.frx":5883
      End
      Begin VB.PictureBox picLogin 
         Appearance      =   0  '���
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1530
         Picture         =   "frmInterface.frx":5AE8
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
            Name            =   "����"
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
         ScrollBars      =   2  '����
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
            Name            =   "����"
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
         ScrollBars      =   3  '�����
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
         Style           =   1  '�׷���
         TabIndex        =   4
         Top             =   3180
         Value           =   1  'Ȯ��
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
                  Picture         =   "frmInterface.frx":6072
                  Key             =   "RUN"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":660C
                  Key             =   "NOT"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":6BA6
                  Key             =   "STOP"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":7140
                  Key             =   "LST"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":79D2
                  Key             =   "ITM"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":7B2C
                  Key             =   "ERR"
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":7C86
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
         SpreadDesigner  =   "frmInterface.frx":7DE0
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
         SpreadDesigner  =   "frmInterface.frx":8045
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
         SpreadDesigner  =   "frmInterface.frx":82AA
      End
      Begin FPSpread.vaSpread vasExcel 
         Height          =   2205
         Left            =   3390
         TabIndex        =   66
         Top             =   5790
         Visible         =   0   'False
         Width           =   1185
         _Version        =   393216
         _ExtentX        =   2090
         _ExtentY        =   3889
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
         SpreadDesigner  =   "frmInterface.frx":850F
      End
      Begin VB.Label Label5 
         Appearance      =   0  '���
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '����
         Caption         =   "�������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   240
         TabIndex        =   50
         Top             =   5250
         Width           =   780
      End
      Begin VB.Label lblSaveSeq 
         Alignment       =   2  '��� ����
         Caption         =   "99999"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   165
         Left            =   2790
         TabIndex        =   49
         Top             =   5250
         Width           =   615
      End
      Begin VB.Label lblExamDate 
         Alignment       =   2  '��� ����
         Caption         =   "20160202"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   165
         Left            =   3600
         TabIndex        =   48
         Top             =   5250
         Width           =   1005
      End
      Begin VB.Label lblUser 
         BackStyle       =   0  '����
         BorderStyle     =   1  '���� ����
         BeginProperty Font 
            Name            =   "����"
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
      Align           =   2  '�Ʒ� ����
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   11790
      Width           =   18825
      _ExtentX        =   33205
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   2646
            MinWidth        =   2646
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   17109
            MinWidth        =   17109
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   2646
            MinWidth        =   2646
            TextSave        =   "2019-07-09"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2646
            MinWidth        =   2646
            TextSave        =   "���� 1:42"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu MnMain 
      Caption         =   "Main"
      Begin VB.Menu MnPrint 
         Caption         =   "�μ�"
         Begin VB.Menu MnPrintLand 
            Caption         =   "�����μ�"
         End
         Begin VB.Menu MnPrintPort 
            Caption         =   "�����μ�"
         End
      End
      Begin VB.Menu MnExit 
         Caption         =   "����"
      End
   End
   Begin VB.Menu MnConfig 
      Caption         =   "Setting"
      Begin VB.Menu MnTConfig 
         Caption         =   "��ż���"
      End
      Begin VB.Menu MnExamConfig 
         Caption         =   "�ڵ弳��"
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
      Visible         =   0   'False
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

'Const colSpecNo = 0     '�̻��
'Const colCheckBox = 1
'Const colSAVESEQ = 2    '�������(��¥��)
'Const colEXAMDATE = 3   '�˻�����
'Const colHOSPDATE = 4   '������������
'Const colBARCODE = 5
'Const colCHARTNO = 6
'Const colPID = 7        '���Ϲ�ȣ(������ȣ)
'Const colINOUT = 8      '�Կ�/�ܷ�
'Const colDISKNO = 9
'Const colPOSNO = 10
'Const colPNAME = 11
'Const colPSEX = 12
'Const colPAGE = 13
'Const colOCNT = 14
'Const colRCNT = 15
'Const colState = 16

'sendflag
'0: Order
'1: Result
'2: Trans
'vasres, vasrres colum
'Const colEQUIPCODE = 1
'Const colEXAMCODE = 2
'Const colEXAMNAME = 3
'Const colMachResult = 4
'Const colRESULT = 5
'Const colSeq = 6
'Const colFLAG = 7
'Const colSubCode = 8

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
Const SB As String = ""  'Chr(11)
Const EB As String = ""   'Chr(28)


Dim strRecvData()   As String
Dim intPhase        As Integer
Dim strState        As String
Dim intBufCnt       As Integer
Dim blnIsETB        As Boolean
Dim intSndPhase     As Integer
Dim intFrameNo      As Integer
'===============================

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type


Dim OFName As OPENFILENAME

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long



'''��������
''''Provider=MSDAORA.1;Password=megadb1;User ID=megadb1;'+
''''Data Source=megadb1;Persist Security Info=True';
'''
'''������
'''                                         ���
'''' Update tl_workvalue Set rvalue = '''+TMaster.FResult+''' '+#13#10+
'                        ����ڵ�
'''' , autocode = '''+TGlobal.FICode+''' '+#13#10+
                            '���ڵ�
'''' WHERE SAMPLE = '''+TMaster.FBarCode+''' '+#13#10+
                            '�˻��ڵ�
'''' AND RTNAME = '''+TMaster.FExamCode+'''


'''                             ���     ��¥   ���ڵ�   ���    ä��    '1'
'''' INSERT INTO TL_ACSDATA (AUTOCODE, LABDATE, SERIAL, RVALUE, SOKET, EVALUECHK) '+#13#10+
'''' values ('''+TGlobal.FICode+''', sysdate, '''+TMaster.FBarCode+''', '''+TMaster.FResult+''', '''+TMaster.FIfCode+''', 1) ';
'''


'''---���� ����---
'''sSender + PadRightStr(IntToStr(i),' ', 5) + PadRightStr(Bcd,' ', 20)+ PadRightStr(IfCd,' ', 150);
'''���ϸ��� nimbus.lis �� ����

' "1    " & "BCD1               " & "RV16" & space(146) &
' "2    " & "BCD2               " & "RV16" & space(146) &

Private Sub chkMode_Click()
    If chkMode.Value = 1 Then
        chkMode.Caption = "Auto"
    Else
        chkMode.Caption = "Manual"
    End If
End Sub


Private Sub chkSearch_Click()
    
    If chkSearch.Value = "1" Then
        tmrResult.Enabled = True
    Else
        tmrResult.Enabled = False
    End If
    
End Sub

Private Sub chkWAll_Click()
    Dim iRow As Long
    
    With vasID
        If chkWAll.Value = 1 Then
            For iRow = 1 To .DataRowCnt
                .Row = iRow
                .Col = colCHECKBOX
                .Value = 1
            Next iRow
        ElseIf chkWAll.Value = 0 Then
            For iRow = 1 To .DataRowCnt
                .Row = iRow
                .Col = colCHECKBOX
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

' Excel Object Library �� �����մϴ�.
Dim xlApp As Excel.Application
Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet

Dim iRow As Integer
Dim iCol As Integer
Dim i As Integer

    Set xlApp = CreateObject("Excel.Application")
    
    xlApp.DisplayAlerts = False
    
    Set xlBook = xlApp.Workbooks.Add
    
    Set xlSheet = xlBook.Worksheets(1)
     
    For iRow = 0 To argSpread.DataRowCnt
        For iCol = 1 To argSpread.DataColCnt
            argSpread.Row = iRow
            argSpread.Col = iCol
            xlSheet.Cells(iRow + 1, iCol) = argSpread.Text
        Next iCol
    Next iRow
    
    xlBook.SaveAs (Filename)
    xlApp.Quit


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
    Dim blnWrite As Variant
    
    ClearSpread vasPrint

    blnWrite = False
    vasPrint.MaxRows = vasID.MaxRows
    vasPrint.MaxCols = vasID.MaxCols
    
    For iRow = 1 To vasID.DataRowCnt
        vasID.Row = iRow
        vasID.Col = 1
            
        If vasID.Value = 1 Then
            If blnWrite = False Then
                For j = 1 To vasID.MaxCols
                    SetText vasPrint, Trim(GetText(vasID, 0, j)), 0, j
                Next
            End If
            
            For j = 1 To vasID.MaxCols
                SetText vasPrint, Trim(GetText(vasID, iRow, j)), iRow, j
            Next
        End If
    Next iRow
    
    If vasPrint.DataRowCnt < 1 Then
        MsgBox "������ �ڷᰡ �����ϴ�.", , "�� ��"
        Exit Sub
    Else
        CommonDialog1.Filter = "Excel Files (*.xls)|*.xls|All Files (*.*)|*.*"
        CommonDialog1.ShowSave
        sFileName = CommonDialog1.Filename
        SaveExcel sFileName, vasPrint
        MsgBox "���� ����Ϸ�", vbOKOnly + vbInformation, Me.Caption
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
    
    txtRack.Text = "5"
    txtPos.Text = "A"
    
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
                
                      SQL = " UPDATE PATRESULT SET " & vbCrLf
                SQL = SQL & "  SENDFLAG = '2' " & vbCrLf
                SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf
                SQL = SQL & "   AND BARCODE = '" & Trim(GetText(vasID, lRow, colBARCODE)) & "' "
                
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

Private Sub cmdOrder_Click()
    Dim intRow      As Long
    Dim intCnt      As Integer
    Dim strBC       As String
    Dim FileNumber
    Dim strAllStr   As String

    strAllStr = ""
    intCnt = 0
    
    '''---���� ����---
    '''sSender + PadRightStr(IntToStr(i),' ', 5) + PadRightStr(Bcd,' ', 20)+ PadRightStr(IfCd,' ', 150);
    '''���ϸ��� nimbus.lis �� ����
    
    ' "1    " & "BCD1               " & "RV16" & space(146) &
    ' "2    " & "BCD2               " & "RV16" & space(146) &
    
    With vasID
        For intRow = 1 To .MaxRows
            .Row = intRow
            .Col = colCHECKBOX
            If .Value = "1" Then
                strBC = GetText(vasID, intRow, colBARCODE)
                If strBC <> "" Then
                    intCnt = intCnt + 1
                    strAllStr = strAllStr & CStr(intCnt) & Space(5 - Len(CStr(intCnt))) & _
                                            strBC & Space(20 - Len(CStr(intCnt))) & _
                                            "RV16" & Space(146)
                    '-- �������(Order) ǥ��
                    Call SetText(vasID, "Order", intRow, colState)
                    Call SetText(vasID, "0", intRow, colCHECKBOX)
                End If
            End If
        Next
    End With

    If strAllStr <> "" Then
        'nimbus.lis ���� ����� ======================================================
        '-- OPEN
        FileNumber = FreeFile
        Open gOrderPath & "\nimbus.lis" For Output As FileNumber
        Close FileNumber
        '-- WRITE
        Open gOrderPath & "\nimbus.lis" For Append As FileNumber
        Print #FileNumber, strAllStr
        Close FileNumber
        '=========================================================================
    Else
        MsgBox "������ ������ �����ϴ�", vbOKOnly + vbCritical, Me.Caption
    End If

End Sub

Private Sub cmdPatDelete_Click()
    Dim i As Integer
    Dim j As Integer
    
    j = 0
    With vasID
        For i = .DataRowCnt To 1 Step -1
            .Row = i
            .Col = colCHECKBOX
            If .Value = "1" Then
                .Action = ActionDeleteRow
                .MaxRows = .MaxRows - 1
                j = j + 1
            End If
        Next
    End With
    
End Sub


'���� ������ �׸��忡 �ֱ�
Private Sub Excel_Open()
    Dim xlApp   As New Excel.Application
    Dim XLappWS As Worksheet
    Dim lngSCnt As Long
    Dim lngSColCnt(100) As Long
    Dim dummy       As String
    Dim strRowData  As Variant
    Dim lngRowCnt   As Long
    Dim chk_str     As String
    Dim dummy_max   As Long
    Dim lngTotColCnt   As Long
    Dim lngTotRowCnt   As Long
    Dim i, j, k     As Long

    
    
'Dim xlapp As New Excel.Application
'Dim xlapp_worksheet As Worksheet
'Dim sheet_count As Long
'Dim sheet_col_count(100) As Long
'Dim i, j, k As Long
'Dim dummy As String
'Dim row_data As Variant
'Dim row_cnt As Long
'Dim chk_str As String
'Dim dummy_max As Long
'Dim tot_col_count As Long
'Dim tot_row_count As Long
    
    lngTotColCnt = 0
    lngTotRowCnt = 0
    
    
    '���� ����
    CommonDialog1.Filter = "Excel(*.xlsx)|*.xlsx|Excel(*.xls)|*.xls"
    CommonDialog1.Action = 1
    
    
    If CommonDialog1.FileTitle = "" Then
        Exit Sub
    End If
    
    xlApp.Workbooks.Open (Trim(CommonDialog1.Filename))
    
    lngSCnt = xlApp.Worksheets.Count
    
    '-- ��ü ��ũ��Ʈ �ҷ�����ͼ� '�ӽ�.txt' ���Ϸ� ����
    For i = 1 To lngSCnt
        Set XLappWS = xlApp.Worksheets(i)
        XLappWS.Activate
        lngSColCnt(i) = XLappWS.UsedRange.Columns.Count
        xlApp.DisplayAlerts = False
    
        '''xlApp.ActiveWorkbook.SaveAs App.Path & "\" & Trim(i) & ".txt", xlText, "", "", False, False '==>2000 + 2003 ����
        xlApp.ActiveWorkbook.SaveAs "C:\CFX_EXCEL\" & Trim(i) & ".txt", xlText, "", "", False, False '==>2000 + 2003 ����
        
        
        'XLappWS.SaveAs App.Path & "\temp\temp" & Trim(i) & ".txt", xlText, "", "", False, False ==>���� 2000��
        'ActiveWorkbook.SaveAs App.Path & "\temp\temp" & Trim(i) & ".txt", xlText, "", "", False, False  ===>���� 2003��
    Next i
    
    xlApp.Quit
    Set XLappWS = Nothing
    Set xlApp = Nothing
    
    '-- ��ü ������ MAX cols�� ����
    dummy_max = 0
    For i = 1 To lngSCnt
        If lngSColCnt(i) >= dummy_max Then
            dummy_max = lngSColCnt(i)
        End If
    Next i
    lngTotColCnt = dummy_max
    
    '��ü row�� ����
    For i = 1 To lngSCnt
'''        Open (App.Path & "\" & Trim(i) & ".txt") For Input As #1
        Open ("C:\CFX_EXCEL\" & Trim(i) & ".txt") For Input As #1
        While Not EOF(1)
            Line Input #1, dummy
            strRowData = Split(Trim(dummy), Chr(9))
            chk_str = ""
            For j = 0 To UBound(strRowData)
                chk_str = chk_str & strRowData(j)
            Next j
            If Len(Trim(dummy)) > 0 Then
                lngTotRowCnt = lngTotRowCnt + 1
            End If
        Wend
        Close #1
    Next i
    
    '-- �׸��� �ʱ�ȭ
    vasExcel.MaxRows = 0
    vasExcel.MaxRows = lngTotRowCnt
    vasExcel.MaxCols = lngTotColCnt
    
    '-- �׸��忡 ���
    For i = 1 To lngSCnt
        '''Open (App.Path & "\" & Trim(i) & ".txt") For Input As #1
        Open ("C:\CFX_EXCEL\" & Trim(i) & ".txt") For Input As #1
        While Not EOF(1)
            Line Input #1, dummy
            strRowData = Split(Trim(dummy), Chr(9))
            chk_str = ""
            For j = 0 To UBound(strRowData)
                chk_str = chk_str & strRowData(j)
            Next j
            If Len(chk_str) > 0 Then
                lngRowCnt = lngRowCnt + 1
                For j = 0 To UBound(strRowData)
                    Call vasExcel.SetText(j + 1, lngRowCnt, CStr(strRowData(j)))
                Next j
            End If
        Wend
        Close #1
    Next i

'    Call SpreadSheetSort(vasExcel, 6, 2)
    With vasExcel
        .Col = 1: .Col2 = .MaxCols
        .Row = 2: .Row2 = .DataRowCnt
        .SortBy = 0
        .SortKey(1) = 2       '����Ű ����ȣ
        .SortKey(2) = 6       '����Ű ����ȣ

        .SortKeyOrder(1) = SortKeyOrderAscending
        .SortKeyOrder(2) = SortKeyOrderAscending

        .Action = ActionSort
    End With


End Sub



Private Function SeqSearch(ByVal brspread As Object, ByVal brSeq As String, ByVal brCol As Integer) As Long
Dim sCnt As Long

    SeqSearch = 0
    If brspread.MaxRows <= 0 Then
        Exit Function
    End If
    
    With brspread
        For sCnt = 1 To .MaxRows
            .Row = sCnt
            .Col = brCol
            If Trim(.Text) = brSeq Then
                SeqSearch = sCnt 'brSeq
                .Action = ActionActiveCell
                .Refresh
                Exit For
            End If
        Next sCnt
    End With

End Function



'Private Sub cmdResult_Click()
'    Dim sSeq As String
'    Dim sBarcode As String
'    Dim strEqpResult As String
'    Dim strLisResult As String
'    Dim strIntBase As String
'    Dim lsExamCode As String
'    Dim lsExamName As String
'    Dim lsSeqNo As String
'    Dim lsResRow    As String
'    Dim lsEquipRes As String
'    Dim lsResult_Buff As String
'
'    Dim lRow As Integer
'    Dim lRow1 As Integer
'    Dim intRow As Integer
'    Dim sWellOld As String
'    Dim sWell As String
'    Dim sExamCode As String
'    Dim sExamName As String
'    Dim sEquipCode As String
'    Dim sItemCode As String
'    Dim strAge As String
'    Dim strSex As String
'    Dim strPtno As String
'    Dim strPtname As String
'    Dim varTmp As Variant
'    Dim intTstCnt As Integer
'    Dim intCol   As Integer
'
'    Screen.MousePointer = 11
'
'    vasExcel.MaxRows = 0
'
'    Call Excel_Open
'
'    intTstCnt = 0
'
'    With vasExcel
'        For intRow = 2 To .DataRowCnt
'
'            .GetText 6, intRow, varTmp: sSeq = varTmp
'            .GetText 2, intRow, varTmp: sWell = varTmp
'            If sSeq <> "" Then
'                With mResult
'                    .BarNo = sSeq
'                    .RsltDate = Format(Now, "yyyymmddhhmmss")
'                    .RsltSeq = getMaxTestNum(Format(dtpToday, "yyyymmdd"))
'                    .RackNo = Val(Mid(sWell, 2))
'                    .TubePos = Mid(sWell, 1, 1)
'                End With
'
'                .GetText 3, intRow, varTmp: strIntBase = varTmp
'                '.GetText 10, intRow, varTmp: strEqpResult = varTmp
'
'                If strIntBase = "FAM" Then
'                    Call SetPatInfo(sSeq)
'
'                    vasID.GetText colBARCODE, gRow, varTmp: sBarcode = varTmp
'                    SetText vasID, "Result", gRow, colState
'
'                    '-- ä��
'                    .GetText 3, intRow, varTmp: strIntBase = varTmp
'
'                    '-- ���
'                    .GetText 10, intRow, varTmp: strEqpResult = varTmp
'
'                    If Val(strEqpResult) = 0 Then
'                        strLisResult = "Not-Detected"
'                    Else
'                       ' strLisResult = CSng(strEqpResult)
'                        strLisResult = Convert2EXP(strEqpResult, "")
'                    End If
'
'
'                    If strLisResult <> "" Then
'                              SQL = "Select examcode, examname, seqno "
'                        SQL = SQL & "  From EQPMASTER"
'                        SQL = SQL & " Where equipno = '" & gEquip & "' "
'                        SQL = SQL & "   AND equipcode = '" & strIntBase & "' "
'                        SQL = SQL & "   and examcode in (" & gOrderExam & ") "      '"'36721','36722','36723','36724'"
'
'                        Res = GetDBSelectColumn(gLocal, SQL)
'
'                        If Res > 0 Then
'                            lsExamCode = Trim(gReadBuf(0))
'                            lsExamName = Trim(gReadBuf(1))
'                            lsSeqNo = Trim(gReadBuf(2))
'
'                            lsResRow = vasRes.DataRowCnt + 1
'                            If vasRes.MaxRows < lsResRow Then
'                                vasRes.MaxRows = lsResRow
'                            End If
'
'                            '�Ҽ��� ó��, ��� ���� ó��
'                            'lsEquipRes = strLisResult
'                            'strLisResult = SetResult(strLisResult, strIntBase)
'                            'lsResult_Buff = strLisResult
'
'                            For intCol = colState + 1 To vasID.MaxCols
'                                If lsExamCode = gArrEquip(intCol - colState, 3) Then
'                                    SetText vasID, strLisResult, gRow, intCol
'                                    Exit For
'                                End If
'                            Next
'
'                            SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      '����ڵ�
'                            SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '�˻��ڵ�
'                            SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '�˻��
'                            SetText vasRes, lsEquipRes, lsResRow, colMachResult     '�����
'                            SetText vasRes, strLisResult, lsResRow, colRESULT          '���
'                            SetText vasRes, lsSeqNo, lsResRow, colSeq               '����
''                            SetText vasRes, strComm, lsResRow, 7                    'Flag
'                            '-- ���� ����
'                            SetLocalDB gRow, lsResRow, "1", lsEquipRes
'
'                            lsResult_Buff = ""
'
'                        Else
'                            '-- ���� ���� ���
'                                  SQL = "Select examcode, examname, seqno "
'                            SQL = SQL & "  From EQPMASTER"
'                            SQL = SQL & " Where equipno = '" & gEquip & "' "
'                            SQL = SQL & "   AND equipcode = '" & strIntBase & "' "
'                            Res = GetDBSelectColumn(gLocal, SQL)
'
'                            If Res > 0 Then
'                                lsExamCode = Trim(gReadBuf(0))
'                                lsExamName = Trim(gReadBuf(1))
'                                lsSeqNo = Trim(gReadBuf(2))
'
'                                lsResRow = vasRes.DataRowCnt + 1
'                                If vasRes.MaxRows < lsResRow Then
'                                    vasRes.MaxRows = lsResRow
'                                End If
'
'                                '�Ҽ��� ó��, ��� ���� ó��
'                                'lsEquipRes = strLisResult
'                                'strLisResult = SetResult(strLisResult, strIntBase)
'                                'lsResult_Buff = strLisResult
'
'                                For intCol = colState + 1 To vasID.MaxCols
'                                    If lsExamCode = gArrEquip(intCol - colState, 3) Then
'                                        SetText vasID, strLisResult, gRow, intCol
'                                        Exit For
'                                    End If
'                                Next
''
'                                SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      '����ڵ�
'                                SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '�˻��ڵ�
'                                SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '�˻��
'                                SetText vasRes, lsEquipRes, lsResRow, colMachResult     '�����
'                                SetText vasRes, strLisResult, lsResRow, colRESULT          '���
'                                SetText vasRes, lsSeqNo, lsResRow, colSeq               '����
''                                SetText vasRes, strComm, lsResRow, 7                    'Flag
'                                '-- ���� ����
'                                SetLocalDB gRow, lsResRow, "1", lsEquipRes
'
'
'                                lsResult_Buff = ""
'                                strState = ""
'                            End If
'                        End If
'                    End If
'
'                    strState = "R"
'                End If
'            End If
'        Next
'    End With
'
'    Screen.MousePointer = 0
'
'End Sub


Function Convert2EXP(ByVal srcV#, Optional fmt$) As String
    Dim mul%, dat#, sign$
    
    If srcV# = 0 Then
        Convert2EXP = "0E+00"
        Exit Function
    End If
    
    If srcV# < 0 Then
        sign$ = "-"
        srcV# = Abs(srcV#)
    Else
        sign$ = ""
    End If
    
    mul% = Int(Log(srcV#) / Log(10))
    dat# = srcV# * 10 ^ (mul% * -1)
        
    If fmt$ = "" Then
        Convert2EXP = sign$ & dat# & "E" & Format$(mul%, "+00;-00")
    Else
        Convert2EXP = sign$ & Format$(dat#, fmt$) & "E" & Format$(mul%, "+00;-00")
    End If
    
    If Right(Convert2EXP, 2) = "01" Then    '9.704507E+01
        Convert2EXP = Mid(Convert2EXP, 1, InStr(Convert2EXP, "E") - 1)
        Convert2EXP = Convert2EXP * 10
        Convert2EXP = Format(Convert2EXP, "#.#0")
    ElseIf Right(Convert2EXP, 2) = "02" Then
        Convert2EXP = Mid(Convert2EXP, 1, InStr(Convert2EXP, "E") - 1)
        Convert2EXP = Convert2EXP * 100
        Convert2EXP = Format(Convert2EXP, "#.#0")
    Else
        Convert2EXP = Format(Mid(Convert2EXP, 1, InStr(Convert2EXP, "E") - 1), "#.#0") & "X10^" & Val(Mid(Convert2EXP, InStr(Convert2EXP, "E") + 2))
        'Convert2EXP = Convert2EXP * 100
        'Convert2EXP = Format(Convert2EXP, "#.#0")
    End If
    
End Function



Private Sub cmdReadEpoc_Click()
    txtTimer.Text = "1"
End Sub

Private Sub cmdResult_Click()

    If optVer(0).Value = True Then
        Call getCFXExlData(0)
    Else
        Call getCFXExlData(1)
    End If
    
End Sub

Private Sub getCFXExlData(ByVal intIDX As Integer)

    Dim x As Integer, y As Boolean, z As Boolean
    Dim ListCount   As Integer, handle As Integer
    Dim List(10)    As String
    Dim intRow, intCol As Long
    Dim varTmp      As Variant
    Dim strExcel    As String
    Dim rv          As Integer
    Dim lRow        As Integer
    Dim lRow1       As Integer
    Dim sExamCode   As String
    Dim sExamName   As String
    Dim sEquipCode  As String
    Dim sItemCode   As String
    Dim strAge      As String
    Dim strSex      As String
    Dim strPtno     As String
    Dim strPtname   As String
    Dim strTmp      As String
    Dim sFile       As String
'    Dim intSheet    As Integer

    Dim idates1$, idates2$, iexamcode$
    Dim PT_NO$(), patname$(), Sex$(), Age$()
    Dim SPC_NO$(), gnl_item_cd$(), bl_gth_dte$()
    Dim dept$(), wd_no$(), TST_CD$()
    Dim ispcno$

    Dim strResult As String
    Dim strIntBase As String
    Dim lsExamCode As String
    Dim lsExamName As String
    Dim lsSeqNo As String
    Dim lsResult_Buff As String
    Dim lsExamDate As String
    Dim lsEquipRes As String
    Dim lsResRow    As String
    Dim strType     As String
    
'On Error Resume Next

    '-- ��������
    If intIDX = 0 Then
        sFile = ShowOpen("Excel Files (*.xls)|*.xls|All Files (*.*)|*.*", gResultPath)
        If sFile <> "" Then
            vasID.MaxRows = 0
            strExcel = sFile
            vasExcel.ScriptEnhanced = True
            x = vasExcel.IsExcelFile(strExcel)
            If x = 1 Then
                y = vasExcel.GetExcelSheetList(strExcel, List, ListCount, "Report.txt", handle, True)
                If y = True Then
                    z = vasExcel.ImportExcelSheet(handle, 0)
                    If z = True Then
                        'MsgBox "�������� ����"
                    Else
                        'MsgBox "�������� ����"
                    End If
                End If
            End If
            
            With vasExcel
                For intRow = 2 To .DataRowCnt
                    For intCol = 1 To 5
                        .GetText intCol, intRow, varTmp
                        If varTmp <> "" Then
                            Select Case intCol
                            Case 1
                                ispcno$ = varTmp '"12020356983" '"12020152330" ' "12020356983" ''varTmp
'                                rv = sl_d_60_sel_spcno&(ispcno$, pt_no$(), patname$(), sex$(), age$(), gnl_item_cd$(), bl_gth_dte$(), dept$(), wd_no$(), tst_cd$())
                                'rv = 1
                                
                                With mResult
                                    .BarNo = ispcno$
                                    '.RackNo = strRackNo
                                    '.TubePos = strTubePos
                                    .RsltDate = Format(Now, "yyyymmddhhmmss")
                                    .RsltSeq = getMaxTestNum(Format(dtpToday, "yyyymmdd"))
                                    '.SpcPos = strSeq
                                End With
                                   
                                'If ispcno$ = "3116001867" Then Stop
                                Call SetPatInfo(ispcno$)
                                
                                If gRow <= 0 Then
                                    Exit Sub
                                End If
                                
                                strState = "O"
                                
                                '-- ������ ���ȭ�� �ʱ�ȭ
                                vasRes.MaxRows = 0


                                If gRow >= 1 Then
                                    '-- ä��
                                    .GetText 5, intRow, varTmp: sEquipCode = varTmp
                                    '-- ���
                                    .GetText 6, intRow, varTmp: strTmp = varTmp
                                    
                                    'sExamCode = Get_ExamCode(sEquipCode)
                                    'sExamName = Get_ExamName(sExamCode)
                                    If strTmp = "-" Then
                                        strTmp = "N"
                                    ElseIf strTmp = "+" Then
                                        strTmp = "Positive(1+)"
                                    ElseIf strTmp = "++" Then
                                        strTmp = "Positive(2+)"
                                    ElseIf strTmp = "+++" Then
                                        strTmp = "Positive(3+)"
                                    ElseIf strTmp = "++++" Then
                                        strTmp = "Positive(4+)"
                                    'Else
                                    '    strTmp = "Positive"
                                    End If
                                    
                                    strIntBase = sEquipCode
                                    strResult = strTmp
                                    If strResult <> "" And Len(strIntBase) < 7 Then
                                        SQL = ""
                                        SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
                                        SQL = SQL & "  FROM EQPMASTER"
                                        SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
                                        SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
                                        'SQL = SQL & "   AND EXAMCODE in (" & gOrderExam & ") "
                                        
                                        Res = GetDBSelectColumn(gLocal, SQL)
                                        '-- ���� ���� ���
                                        If Res > 0 Then
                                            lsExamCode = Trim(gReadBuf(0))
                                            lsExamName = Trim(gReadBuf(1))
                                            lsSeqNo = Trim(gReadBuf(2))
                                            
                                            lsResRow = vasRes.DataRowCnt + 1
                                            If vasRes.MaxRows < lsResRow Then
                                                vasRes.MaxRows = lsResRow
                                            End If
                                            
                                            '�Ҽ��� ó��, ��� ���� ó��
                                            lsEquipRes = strResult
                                            strResult = SetResult(strResult, strIntBase)
                                            lsResult_Buff = strResult
                                            
                                            '-- Work List
                                            SetText vasID, "Result", gRow, colState                 '11 �������
                                            
                        
                                            '-- ��� List
                                            SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      '����ڵ�
                                            SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '�˻��ڵ�
                                            SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '�˻��
                                            SetText vasRes, lsEquipRes, lsResRow, colMachResult     '�����
                                            SetText vasRes, strResult, lsResRow, colRESULT          '���
                                            SetText vasRes, lsSeqNo, lsResRow, colSeq               '����
                                            'SetText vasRes, strComm, lsResRow, 7                    'Flag
                                            '-- ���� ����
                                            SetLocalDB gRow, lsResRow, "1", lsEquipRes
                                                        
                                            lsResult_Buff = ""
                                            
                                            strState = "R"
                                        End If
                                    End If
                                        
                                    Exit For
                                End If
                            Case "3"
                                'If ispcno$ = "3116001867" Then Stop
                                .GetText 3, intRow, varTmp: strType = varTmp
                                
                            Case "5"
                                If strType = "" Then
                                    '-- ä��
                                    .GetText 5, intRow, varTmp: sEquipCode = varTmp
                                    '-- ���
                                    .GetText 6, intRow, varTmp: strTmp = varTmp
                                    'sExamCode = Get_ExamCode(sEquipCode)
                                    'sExamName = Get_ExamName(sExamCode)
    
                                    '      SQL = "SELECT EXAMCODE FROM PAT_RES WHERE BARCODE = '" & ispcno$ & "' "
                                    'SQL = SQL & "   AND EXAMCODE = '" & sExamCode & "' "
                                    'Res = db_select_Col(gLocal, SQL)
                                    If strTmp = "-" Then
                                        strTmp = "N"
                                    ElseIf strTmp = "+" Then
                                        strTmp = "Positive(1+)"
                                    ElseIf strTmp = "++" Then
                                        strTmp = "Positive(2+)"
                                    ElseIf strTmp = "+++" Then
                                        strTmp = "Positive(3+)"
                                    ElseIf strTmp = "++++" Then
                                        strTmp = "Positive(4+)"
                                    End If
                                    
                                    strIntBase = sEquipCode
                                    strResult = strTmp
                                    If strResult <> "" And Len(strIntBase) < 7 Then
                                        SQL = ""
                                        SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
                                        SQL = SQL & "  FROM EQPMASTER"
                                        SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
                                        SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
                                        'SQL = SQL & "   AND EXAMCODE in (" & gOrderExam & ") "
                                        
                                        Res = GetDBSelectColumn(gLocal, SQL)
                                        '-- ���� ���� ���
                                        If Res > 0 Then
                                            lsExamCode = Trim(gReadBuf(0))
                                            lsExamName = Trim(gReadBuf(1))
                                            lsSeqNo = Trim(gReadBuf(2))
                                            
                                            lsResRow = vasRes.DataRowCnt + 1
                                            If vasRes.MaxRows < lsResRow Then
                                                vasRes.MaxRows = lsResRow
                                            End If
                                            
                                            '�Ҽ��� ó��, ��� ���� ó��
                                            lsEquipRes = strResult
                                            strResult = SetResult(strResult, strIntBase)
                                            lsResult_Buff = strResult
                                            
                                            '-- Work List
                                            SetText vasID, "Result", gRow, colState                 '11 �������
                                            
                        
                                            '-- ��� List
                                            SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      '����ڵ�
                                            SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '�˻��ڵ�
                                            SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '�˻��
                                            SetText vasRes, lsEquipRes, lsResRow, colMachResult     '�����
                                            SetText vasRes, strResult, lsResRow, colRESULT          '���
                                            SetText vasRes, lsSeqNo, lsResRow, colSeq               '����
                                            'SetText vasRes, strComm, lsResRow, 7                    'Flag
                                            '-- ���� ����
                                            SetLocalDB gRow, lsResRow, "1", lsEquipRes
                                                        
                                            lsResult_Buff = ""
                                            
                                            strState = "R"
                                        End If
                                    End If
                                    Exit For
                                End If
                            End Select
                        End If
                    Next
                Next
            End With
        End If
    
    Else
    '-- �Ź���
        sFile = ShowOpen("Excel Files (*.xls)|*.xls|All Files (*.*)|*.*", App.Path)
        If sFile <> "" Then
            vasID.MaxRows = 0
            strExcel = sFile
            vasExcel.ScriptEnhanced = True
            x = vasExcel.IsExcelFile(strExcel)
            If x = 1 Then
                y = vasExcel.GetExcelSheetList(strExcel, List, ListCount, "Report.txt", handle, True)
                If y = True Then
                    '-- 2��°sheet
                    z = vasExcel.ImportExcelSheet(handle, 1)
                    If z = True Then
                        'MsgBox "�������� ����"
                        With vasExcel
                            For intRow = 2 To .DataRowCnt Step 4
                                .GetText 1, intRow, varTmp
                                ispcno$ = varTmp '"12020356983" '"12020152330" ' "12020356983" ''varTmp
                                
                                With mResult
                                    .BarNo = ispcno$
                                    '.RackNo = strRackNo
                                    '.TubePos = strTubePos
                                    .RsltDate = Format(Now, "yyyymmddhhmmss")
                                    .RsltSeq = getMaxTestNum(Format(dtpToday, "yyyymmdd"))
                                    '.SpcPos = strSeq
                                End With
                                                    
                                Call SetPatInfo(ispcno$)
                                
                                If gRow <= 0 Then
                                    Exit Sub
                                End If
                                
                                strState = "O"
                                
                                '-- ������ ���ȭ�� �ʱ�ȭ
                                vasRes.MaxRows = 0
                                
'                                rv = sl_d_60_sel_spcno&(ispcno$, pt_no$(), patname$(), sex$(), age$(), gnl_item_cd$(), bl_gth_dte$(), dept$(), wd_no$(), tst_cd$())
'                                'rv = 1
'                                If rv >= 1 Then
'                                    '-- ȯ������
'                                    vasID.MaxRows = vasID.MaxRows + 1
'                                    vasID.RowHeight(-1) = 12
'                                    lRow = vasID.MaxRows
'                                    sExamName = Get_ExamName(tst_cd(0))
'                                    sEquipCode = Get_EquipCode(tst_cd(0))
'                                    sItemCode = gnl_item_cd(0)
'                                    strPtno = patname(0)
'                                    strPtname = patname(0)
'                                    strSex = sex(0)
'                                    strAge = age(0)
'
'                                    SetText vasID, ispcno$, lRow, 2
'                                    SetText vasID, gnl_item_cd(0), lRow, 4         'ä������
'                                    SetText vasID, pt_no(0), lRow, 6           '���Ϲ�ȣ
'                                    SetText vasID, patname(0), lRow, 7         '�̸�
'                                    SetText vasID, sex(0), lRow, 8            '����
'                                    SetText vasID, age(0), lRow, 9            '����
'                                End If
                            
                                For intCol = 6 To 22 Step 2
                                    .GetText intCol, intRow, varTmp
                                    If varTmp <> "" Then
                                        '-- Row = 1,2 ä��
                                        .GetText intCol, intRow, varTmp: sEquipCode = varTmp
                                        sExamCode = Get_ExamCode(sEquipCode)
                                        sExamName = Get_ExamName(sExamCode)
                
                                              SQL = "SELECT EXAMCODE FROM PAT_RES WHERE BARCODE = '" & ispcno$ & "' "
                                        SQL = SQL & "   AND EXAMCODE = '" & sExamCode & "' "
                                        Res = db_select_Col(gLocal, SQL)
                                        '-- ���
                                        .GetText intCol, intRow + 1, varTmp: strTmp = varTmp
                                        If strTmp = "-" Then
                                            strTmp = "Negative"
                                        ElseIf strTmp = "+" Then
                                            strTmp = "Positive(1+)"
                                        ElseIf strTmp = "++" Then
                                            strTmp = "Positive(2+)"
                                        ElseIf strTmp = "+++" Then
                                            strTmp = "Positive(3+)"
                                        ElseIf strTmp = "++++" Then
                                            strTmp = "Positive(4+)"
                                        'Else
                                        '    strTmp = "Positive"
                                        End If
                                        
                                        If Res > 0 Then
                                            SQL = "DELETE FROM PAT_RES " & vbCrLf & _
                                                  "WHERE EXAMDATE = '" & Format(dtpToday, "YYYYMMDD") & "' " & vbCrLf & _
                                                  "  AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
                                                  "  AND BARCODE = '" & ispcno$ & "' " & vbCrLf & _
                                                  "  and equipcode = '" & sEquipCode & "'" & vbCrLf & _
                                                  "  and examcode= '" & sExamCode & "'"
                                            Res = SendQuery(gLocal, SQL)
                                        End If
                                        
                                        SQL = "INSERT INTO PAT_RES( EQUIPNO, BARCODE, RECENO, PID, PNAME, EXAMCODE, EXAMNAME, PSEX, PAGE, EXAMDATE, EQUIPCODE,RESULT) " & vbCrLf & _
                                              "VALUES('" & gEquip & "', '" & ispcno$ & "', '" & sItemCode & "', " & vbCrLf & _
                                              "'" & strPtno & "', '" & strPtname & "', '" & sExamCode & "', " & vbCrLf & _
                                              "'" & sExamName & "', '" & strSex & "', '" & strAge & "', '" & Format(dtpToday, "YYYYMMDD") & "', '" & sEquipCode & "','" & strTmp & "')"
                                        Res = SendQuery(gLocal, SQL)
                                        
                                        '-- Row = 3,4 ä��
                                        .GetText intCol, intRow + 2, varTmp: sEquipCode = varTmp
                                        sExamCode = Get_ExamCode(sEquipCode)
                                        sExamName = Get_ExamName(sExamCode)
                
                                              SQL = "SELECT EXAMCODE FROM PAT_RES WHERE BARCODE = '" & ispcno$ & "' "
                                        SQL = SQL & "   AND EXAMCODE = '" & sExamCode & "' "
                                        Res = db_select_Col(gLocal, SQL)
                                        '-- ���
                                        .GetText intCol, intRow + 3, varTmp: strTmp = varTmp
                                        If strTmp = "-" Then
                                            strTmp = "Negative"
                                        ElseIf strTmp = "+" Then
                                            strTmp = "Positive(1+)"
                                        ElseIf strTmp = "++" Then
                                            strTmp = "Positive(2+)"
                                        ElseIf strTmp = "+++" Then
                                            strTmp = "Positive(3+)"
                                        ElseIf strTmp = "++++" Then
                                            strTmp = "Positive(4+)"
                                        'Else
                                        '    strTmp = "Positive"
                                        End If
                                        
                                        If Res > 0 Then
                                            SQL = "DELETE FROM PAT_RES " & vbCrLf & _
                                                  "WHERE EXAMDATE = '" & Format(dtpToday, "YYYYMMDD") & "' " & vbCrLf & _
                                                  "  AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
                                                  "  AND BARCODE = '" & ispcno$ & "' " & vbCrLf & _
                                                  "  and equipcode = '" & sEquipCode & "'" & vbCrLf & _
                                                  "  and examcode= '" & sExamCode & "'"
                                            Res = SendQuery(gLocal, SQL)
                                        End If
                                        
                                        SQL = "INSERT INTO PAT_RES( EQUIPNO, BARCODE, RECENO, PID, PNAME, EXAMCODE, EXAMNAME, PSEX, PAGE, EXAMDATE, EQUIPCODE,RESULT) " & vbCrLf & _
                                              "VALUES('" & gEquip & "', '" & ispcno$ & "', '" & sItemCode & "', " & vbCrLf & _
                                              "'" & strPtno & "', '" & strPtname & "', '" & sExamCode & "', " & vbCrLf & _
                                              "'" & sExamName & "', '" & strSex & "', '" & strAge & "', '" & Format(dtpToday, "YYYYMMDD") & "', '" & sEquipCode & "','" & strTmp & "')"
                                        Res = SendQuery(gLocal, SQL)
                                    End If
                                Next
                            Next
                        End With
                    End If
                    '-- 3��°sheet
                    z = vasExcel.ImportExcelSheet(handle, 2)
                    If z = True Then
                        'MsgBox "�������� ����"
                        With vasExcel
                            For intRow = 3 To .DataRowCnt
                                .GetText 1, intRow, varTmp
                                ispcno$ = varTmp '"12020356983" '"12020152330" ' "12020356983" ''varTmp
'                                rv = sl_d_60_sel_spcno&(ispcno$, pt_no$(), patname$(), sex$(), age$(), gnl_item_cd$(), bl_gth_dte$(), dept$(), wd_no$(), tst_cd$())
                                'rv = 1
                                If rv >= 1 Then
                                    '-- ȯ������
                                    vasID.MaxRows = vasID.MaxRows + 1
                                    vasID.RowHeight(-1) = 12
                                    lRow = vasID.MaxRows
                                    sExamName = Get_ExamName(TST_CD(0))
                                    sEquipCode = Get_EquipCode(TST_CD(0))
                                    sItemCode = gnl_item_cd(0)
                                    strPtno = patname(0)
                                    strPtname = patname(0)
                                    strSex = Sex(0)
                                    strAge = Age(0)
                                    
                                    SetText vasID, ispcno$, lRow, 2
                                    SetText vasID, gnl_item_cd(0), lRow, 4         'ä������
                                    SetText vasID, PT_NO(0), lRow, 6           '���Ϲ�ȣ
                                    SetText vasID, patname(0), lRow, 7         '�̸�
                                    SetText vasID, Sex(0), lRow, 8            '����
                                    SetText vasID, Age(0), lRow, 9            '����
                                End If
                            
                                For intCol = 6 To 20 Step 2
                                    .GetText intCol, 2, varTmp
                                    If varTmp <> "" Then
                                        '-- Row = 1,2 ä��
                                        .GetText intCol, 2, varTmp: sEquipCode = varTmp
                                        sExamCode = Get_ExamCode(sEquipCode)
                                        sExamName = Get_ExamName(sExamCode)
                
                                              SQL = "SELECT EXAMCODE FROM PAT_RES WHERE BARCODE = '" & ispcno$ & "' "
                                        SQL = SQL & "   AND EXAMCODE = '" & sExamCode & "' "
                                        Res = db_select_Col(gLocal, SQL)
                                        '-- ���
                                        .GetText intCol, intRow, varTmp: strTmp = varTmp
                                        If strTmp = "-" Then
                                            strTmp = "Negative"
                                        ElseIf strTmp = "+" Then
                                            strTmp = "Positive(1+)"
                                        ElseIf strTmp = "++" Then
                                            strTmp = "Positive(2+)"
                                        ElseIf strTmp = "+++" Then
                                            strTmp = "Positive(3+)"
                                        ElseIf strTmp = "++++" Then
                                            strTmp = "Positive(4+)"
                                        'Else
                                        '    strTmp = "Positive"
                                        End If
                                        
                                        If Res > 0 Then
                                            SQL = "DELETE FROM PAT_RES " & vbCrLf & _
                                                  "WHERE EXAMDATE = '" & Format(dtpToday, "YYYYMMDD") & "' " & vbCrLf & _
                                                  "  AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
                                                  "  AND BARCODE = '" & ispcno$ & "' " & vbCrLf & _
                                                  "  and equipcode = '" & sEquipCode & "'" & vbCrLf & _
                                                  "  and examcode= '" & sExamCode & "'"
                                            Res = SendQuery(gLocal, SQL)
                                        End If
                                        
                                        SQL = "INSERT INTO PAT_RES( EQUIPNO, BARCODE, RECENO, PID, PNAME, EXAMCODE, EXAMNAME, PSEX, PAGE, EXAMDATE, EQUIPCODE,RESULT) " & vbCrLf & _
                                              "VALUES('" & gEquip & "', '" & ispcno$ & "', '" & sItemCode & "', " & vbCrLf & _
                                              "'" & strPtno & "', '" & strPtname & "', '" & sExamCode & "', " & vbCrLf & _
                                              "'" & sExamName & "', '" & strSex & "', '" & strAge & "', '" & Format(dtpToday, "YYYYMMDD") & "', '" & sEquipCode & "','" & strTmp & "')"
                                        Res = SendQuery(gLocal, SQL)
                                        
                                        
                                    End If
                                Next
                            Next
                        End With
                    End If
                    
                End If
            End If
        End If
    End If

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
    vasRes.MaxRows = 0
          
          'SQL = " SELECT '', SAVESEQ, MID(EXAMDATE,1,8) AS EXAMDATE, HOSPDATE AS ��������, BARCODE AS ���ڵ��ȣ, CHARTNO AS ��Ʈ��ȣ, PID AS ������ȣ, PNAME AS �̸�,PSEX AS ����, PAGE AS ����, DISKNO, POSNO, EXAMCODE, RESULT, REFFLAG, SENDFLAG,INOUT " & vbCrLf
          SQL = " SELECT '', SAVESEQ, EXAMDATE, HOSPDATE AS ��������, BARCODE AS ���ڵ��ȣ, CHARTNO AS ��Ʈ��ȣ, PID AS ������ȣ, PNAME AS �̸�,PSEX AS ����, PAGE AS ����, DISKNO, POSNO, EXAMCODE, RESULT, REFFLAG, SENDFLAG,INOUT " & vbCrLf
    SQL = SQL & "   FROM PATRESULT " & vbCrLf
    SQL = SQL & "  WHERE MID(EXAMDATE,1,8) Between '" & Format(dtpStartDt, "YYYYMMDD") & "' AND '" & Format(dtpStopDt, "YYYYMMDD") & "'" & vbCrLf
    SQL = SQL & "    AND EQUIPNO = '" & gEquip & "' " & vbCrLf
    'SQL = SQL & "    AND SAVESEQ > 0 " & vbCrLf
    SQL = SQL & " ORDER BY EXAMDATE,SAVESEQ,HOSPDATE,BARCODE,SENDFLAG DESC "
    
    Set RS = cn.Execute(SQL, , 1)

    If Not RS.EOF = True And Not RS.BOF = True Then
        Do Until RS.EOF
            With vasID
                For i = 1 To .DataRowCnt
                    strDate = GetText(vasID, i, colHOSPDATE)
                    strChart = GetText(vasID, i, colBARCODE)
                    strSaveSeq = GetText(vasID, i, colSAVESEQ)
                    
                    If Trim(RS("��������")) = strDate And Trim(RS("SAVESEQ")) = strSaveSeq And Trim(RS("���ڵ��ȣ")) = strChart Then
                        blnSame = True
                    End If
                    
                    If blnSame = True Then
                        For intCol = colState + 1 To vasID.MaxCols
                            If Trim(RS.Fields("EXAMCODE")) = gArrEquip(intCol - colState, 3) Then
                                SetText vasID, Trim(RS.Fields("RESULT")) & "", .MaxRows, intCol
                                If Trim(RS.Fields("REFFLAG")) = "H" Then
                                    .Row = .MaxRows
                                    .Col = intCol
                                    .ForeColor = vbRed
                                ElseIf Trim(RS.Fields("REFFLAG")) = "L" Then
                                    .Row = .MaxRows
                                    .Col = intCol
                                    .ForeColor = vbBlue
                                End If
                                Exit For
                            End If
                        Next
                    End If
                Next

                If blnSame = False Then
                    .MaxRows = .MaxRows + 1

                    SetText vasID, "0", .MaxRows, colCHECKBOX
                    SetText vasID, Trim(RS.Fields("SAVESEQ")) & "", .MaxRows, colSAVESEQ
                    SetText vasID, Trim(RS.Fields("EXAMDATE")) & "", .MaxRows, colEXAMDATE
                    SetText vasID, Trim(RS.Fields("��������")) & "", .MaxRows, colHOSPDATE
                    SetText vasID, Trim(RS.Fields("���ڵ��ȣ")) & "", .MaxRows, colBARCODE
                    SetText vasID, Trim(RS.Fields("��Ʈ��ȣ")) & "", .MaxRows, colCHARTNO
                    SetText vasID, Trim(RS.Fields("������ȣ")) & "", .MaxRows, colPID
                    SetText vasID, Trim(RS.Fields("�̸�")) & "", .MaxRows, colPNAME
                    SetText vasID, Trim(RS.Fields("����")) & "", .MaxRows, colPSEX
                    SetText vasID, Trim(RS.Fields("����")) & "", .MaxRows, colPAGE
                    SetText vasID, Trim(RS.Fields("INOUT")) & "", .MaxRows, colINOUT
                    SetText vasID, Trim(RS.Fields("DISKNO")) & "", .MaxRows, colDISKNO
                    SetText vasID, Trim(RS.Fields("POSNO")) & "", .MaxRows, colPOSNO
                    
                    Select Case Trim(RS.Fields("SENDFLAG")) & ""
                        Case "0": SetText vasID, "����", .MaxRows, colState
                                  SetBackColor vasID, .MaxRows, .MaxRows, 1, colState, 202, 201, 112
                        Case "1": SetText vasID, "���", .MaxRows, colState
                        Case "2": SetText vasID, "�Ϸ�", .MaxRows, colState
                                  SetBackColor vasID, .MaxRows, .MaxRows, 1, colState, 202, 255, 112
                        Case "3": SetText vasID, "����", .MaxRows, colState
                                  SetBackColor vasID, .MaxRows, .MaxRows, 1, colState, 202, 245, 112
                    End Select
                    
                    For intCol = colState + 1 To vasID.MaxCols
                        If Trim(RS.Fields("EXAMCODE")) = gArrEquip(intCol - colState, 3) Then
                            SetText vasID, Trim(RS.Fields("RESULT")) & "", .MaxRows, intCol
                            If Trim(RS.Fields("REFFLAG")) = "H" Then
                                .Row = .MaxRows
                                .Col = intCol
                                .ForeColor = vbRed
                            ElseIf Trim(RS.Fields("REFFLAG")) = "L" Then
                                .Row = .MaxRows
                                .Col = intCol
                                .ForeColor = vbBlue
                            End If
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
    Dim RS          As ADODB.Recordset
    Dim i           As Integer
    Dim iCnt        As Long
    Dim intRow      As Long
    Dim intCol      As Integer
    Dim strDate     As String
    Dim strChart    As String
    Dim blnSame     As Boolean
    
    If pBarNo = "" Then
        vasID.MaxRows = 0
        intRow = 0
    End If
    
    blnSame = False
    vasID.ReDraw = False
    
          SQL = " SELECT DISTINCT '1', '' AS SN ,'' AS ����Ͻ�, REQ_DT AS ��������" & vbCrLf
    SQL = SQL & ", QC_BAR_NO AS ���ڵ��ȣ, LOT_NO AS ��Ʈ��ȣ, REQ_SEQ AS ������ȣ, '�Կ�' AS �Կ�" & vbCrLf
    SQL = SQL & ", '' AS R, '' AS P, REQ_SEQ AS �̸�, '����' AS ����, REQ_SEQ AS ����, ITEM_CD AS ITEM " & vbCrLf
    SQL = SQL & "  FROM S2QCS101 " & vbCrLf
    SQL = SQL & " WHERE 1=1 " & vbCrLf
    If pBarNo <> "" Then
        SQL = SQL & "   AND QC_BAR_NO = '" & pBarNo & "'" & vbCrLf
    Else
        SQL = SQL & "   AND REQ_DT BETWEEN '" & pFrDt & "' AND '" & pToDt & "'" & vbCrLf
    End If
    'SQL = SQL & "   AND ITEM_CD IN (" & gAllExam & ")" & vbCrLf
    SQL = SQL & " ORDER BY ��������, ���ڵ��ȣ, ��Ʈ��ȣ, ������ȣ"
    
'    If pBarNo <> "" Then
'        Res = GetDBSelectVas(gServer, SQL, vasID, vasID.MaxRows + 1)
'    Else
'        Res = GetDBSelectVas(gServer, SQL, vasID)
'    End If
    
    '-- Record Count ������
    cn_Ser.CursorLocation = adUseClient
    Set RS = cn_Ser.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then
        frmProgress.Show
        frmProgress.ZOrder 0
        frmProgress.Xprog.Min = 1
        frmProgress.Xprog.Max = RS.RecordCount
                
        Do Until RS.EOF
            iCnt = iCnt + 1
            With vasID
                .ReDraw = False
                For i = 1 To .DataRowCnt
                    strDate = GetText(vasID, i, colHOSPDATE)
                    strChart = GetText(vasID, i, colBARCODE)
                    If Trim(RS("��������")) = strDate And Trim(RS("���ڵ��ȣ")) = strChart Then
                        blnSame = True
                    End If
                    For intCol = colState + 1 To vasID.MaxCols
                        If Trim(RS.Fields("ITEM")) = gArrEquip(intCol - colState, 3) Then
                            vasID.Row = .MaxRows
                            vasID.Col = intCol
                            vasID.BackColor = vbYellow
                            Exit For
                        End If
                    Next
                Next
                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    SetText vasID, "1", .MaxRows, colCHECKBOX
                    SetText vasID, Trim(RS.Fields("��������")) & "", .MaxRows, colHOSPDATE
                    SetText vasID, Trim(RS.Fields("���ڵ��ȣ")) & "", .MaxRows, colBARCODE
                    SetText vasID, Trim(RS.Fields("��Ʈ��ȣ")) & "", .MaxRows, colCHARTNO
                    SetText vasID, Trim(RS.Fields("������ȣ")) & "", .MaxRows, colPID
                    SetText vasID, Trim(RS.Fields("�̸�")) & "", .MaxRows, colPNAME
                    SetText vasID, Trim(RS.Fields("����")) & "", .MaxRows, colPSEX
                    SetText vasID, Trim(RS.Fields("����")) & "", .MaxRows, colPAGE
                    For intCol = colState + 1 To vasID.MaxCols
                        If Trim(RS.Fields("ITEM")) = gArrEquip(intCol - colState, 3) Then
                            vasID.Row = .MaxRows
                            vasID.Col = intCol
                            vasID.BackColor = vbYellow
                            Exit For
                        End If
                    Next
                End If
                blnSame = False
            End With
            '-- ���α׷����� ����
            frmProgress.Xprog.Value = iCnt
            DoEvents
            
            RS.MoveNext
        Loop
        chkWAll.Value = "1"
    Else
        StatusBar1.Panels(3).Text = "��ȸ ����ڰ� �����ϴ�."
        chkWAll.Value = "0"
    End If
    
    RS.Close
    '-- ���α׷����� �ݱ�
    Unload frmProgress
    
    vasID.RowHeight(-1) = 12
    vasID.ReDraw = True
    
End Sub

Private Sub GetWorkList_DADESOFT(ByVal pFrDt As String, ByVal pToDt As String, Optional pBarNo As String)
    Dim RS          As ADODB.Recordset
    Dim i           As Integer
    Dim iCnt        As Long
    Dim intRow      As Long
    Dim intCol      As Integer
    Dim strDate     As String
    Dim strChart    As String
    Dim blnSame     As Boolean
    
    If pBarNo = "" Then
        vasID.MaxRows = 0
        intRow = 0
    End If
    
    blnSame = False
    vasID.ReDraw = False
    
'''          SQL = " SELECT DISTINCT '1', '' AS SN ,'' AS ����Ͻ�, '' AS ��������" & vbCrLf
'''    SQL = SQL & ", '' AS ���ڵ��ȣ, '' AS ��Ʈ��ȣ, '' AS ������ȣ, '' AS �Կ�" & vbCrLf
'''    SQL = SQL & ", '' AS R, '' AS P, '' AS �̸�, '' AS ����, '' AS ����, '' AS ITEM " & vbCrLf
'''    SQL = SQL & "  FROM S2QCS101 " & vbCrLf
'''    SQL = SQL & " WHERE 1=1 " & vbCrLf
'''    If pBarNo <> "" Then
'''        SQL = SQL & "   AND QC_BAR_NO = '" & pBarNo & "'" & vbCrLf
'''    Else
'''        SQL = SQL & "   AND REQ_DT BETWEEN '" & pFrDt & "' AND '" & pToDt & "'" & vbCrLf
'''    End If
'''    'SQL = SQL & "   AND ITEM_CD IN (" & gAllExam & ")" & vbCrLf
'''    SQL = SQL & " ORDER BY ��������, ���ڵ��ȣ, ��Ʈ��ȣ, ������ȣ"
    
          SQL = " SELECT DISTINCT '1', '' AS SN, '' AS ����Ͻ�, J.�������� AS ��������," & vbCrLf
    SQL = SQL & "        L.��ü��ȣ AS ���ڵ��ȣ, A.íƮ��ȣ AS ��Ʈ��ȣ, J.������ȣ AS ������ȣ,'�Կ�' AS �Կ�, " & vbCrLf
    SQL = SQL & "        J.����˻�ID AS R, L.��������ID AS P,  A.ȯ���̸� AS �̸�, A.ȯ�ڼ��� AS ����, A.ȯ�ڳ���  AS ����, L.ó���ڵ� + L.�����ڵ� AS ITEM " & vbCrLf
    SQL = SQL & "   FROM TB_����˻� L " & vbCrLf
    SQL = SQL & "  INNER JOIN TB_�������� J ON (L.��������ID=J.��������ID) " & vbCrLf
    SQL = SQL & "  INNER JOIN TB_�����Ϲ� A ON (J.��������=A.�������� AND J.íƮ��ȣ=A.íƮ��ȣ AND J.�����ȣ=A.�����ȣ) " & vbCrLf
    SQL = SQL & "  Where 1 = 1 " & vbCrLf
    SQL = SQL & "    AND J.�������� Between '" & pFrDt & "' and '" & pToDt & "'" & vbCrLf
    SQL = SQL & "    AND L.�˻����� = '" & gDept_Code & "'" & vbCrLf
    SQL = SQL & "    AND L.�˻���� < 5 " & vbCrLf
    If chkSaveAll.Value = "0" Then
        SQL = SQL & "  AND (L.�˻��� = '' OR L.�˻��� IS NULL)"
    End If
    SQL = SQL & "  ORDER BY J.��������, J.������ȣ"
    
    
'          SQL = " SELECT DISTINCT '1', '' AS SN, '' AS ����Ͻ�, L.�������� AS ��������," & vbCrLf
'    SQL = SQL & "        L.��ü��ȣ AS ���ڵ��ȣ, L.íƮ��ȣ AS ��Ʈ��ȣ, '55555' AS ������ȣ,'�Կ�' AS �Կ�, " & vbCrLf
'    SQL = SQL & "        L.����˻�ID AS R, L.��������ID AS P,  'ȫ�浿' AS �̸�, '����' AS ����, '35'  AS ����, L.ó���ڵ� + L.�����ڵ� AS ITEM " & vbCrLf
'    SQL = SQL & "   FROM TB_����˻� L " & vbCrLf
'    SQL = SQL & "  Where 1 = 1 " & vbCrLf
'    SQL = SQL & "    AND L.�������� Between convert(datetime,'" & pFrDt & "') and convert(datetime,'" & pToDt & "')" & vbCrLf
'    SQL = SQL & "    AND L.�˻����� = '" & gDept_Code & "'" & vbCrLf
'    SQL = SQL & "    AND L.�˻���� < 5 " & vbCrLf
'    If chkSaveAll.Value = "0" Then
'        SQL = SQL & "  AND (�˻��� = '' OR �˻��� IS NULL)"
'    End If
'    SQL = SQL & "  ORDER BY L.��������"
    
    
    '-- Record Count ������
    cn_Ser.CursorLocation = adUseClient
    Set RS = cn_Ser.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then
        frmProgress.Show
        frmProgress.ZOrder 0
        frmProgress.Xprog.Min = 1
        frmProgress.Xprog.Max = RS.RecordCount + 1
                
        Do Until RS.EOF
            iCnt = iCnt + 1
            With vasID
                .ReDraw = False
                For i = 1 To .DataRowCnt
                    strDate = GetText(vasID, i, colHOSPDATE)
                    strChart = GetText(vasID, i, colBARCODE)
                    If Trim(RS("��������")) = strDate And Trim(RS("���ڵ��ȣ")) = strChart Then
                        blnSame = True
                    End If
                    For intCol = colState + 1 To vasID.MaxCols
                        If Trim(RS.Fields("ITEM")) = gArrEquip(intCol - colState, 3) Then
                            vasID.Row = .MaxRows
                            vasID.Col = intCol
                            vasID.BackColor = vbYellow
                            Exit For
                        End If
                    Next
                Next
                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    SetText vasID, "1", .MaxRows, colCHECKBOX
                    SetText vasID, Trim(RS.Fields("��������")) & "", .MaxRows, colHOSPDATE
                    SetText vasID, Trim(RS.Fields("���ڵ��ȣ")) & "", .MaxRows, colBARCODE
                    SetText vasID, Trim(RS.Fields("��Ʈ��ȣ")) & "", .MaxRows, colCHARTNO
                    SetText vasID, Trim(RS.Fields("������ȣ")) & "", .MaxRows, colPID
                    
                    SetText vasID, Trim(RS.Fields("R")) & "", .MaxRows, colDISKNO
                    SetText vasID, Trim(RS.Fields("P")) & "", .MaxRows, colPOSNO

                    SetText vasID, Trim(RS.Fields("�̸�")) & "", .MaxRows, colPNAME
                    SetText vasID, Trim(RS.Fields("����")) & "", .MaxRows, colPSEX
                    SetText vasID, Trim(RS.Fields("����")) & "", .MaxRows, colPAGE
                    For intCol = colState + 1 To vasID.MaxCols
                        If Trim(RS.Fields("ITEM")) = gArrEquip(intCol - colState, 3) Then
                            vasID.Row = .MaxRows
                            vasID.Col = intCol
                            vasID.BackColor = vbYellow
                            Exit For
                        End If
                    Next
                End If
                blnSame = False
            End With
            '-- ���α׷����� ����
            frmProgress.Xprog.Value = iCnt
            DoEvents
            
            RS.MoveNext
        Loop
        chkWAll.Value = "1"
    Else
        StatusBar1.Panels(3).Text = "��ȸ ����ڰ� �����ϴ�."
        chkWAll.Value = "0"
    End If
    
    RS.Close
    '-- ���α׷����� �ݱ�
    Unload frmProgress
    
    vasID.RowHeight(-1) = 12
    vasID.ReDraw = True
    
End Sub

Private Sub GetWorkList_TWIN(ByVal pFrDt As String, ByVal pToDt As String, Optional pBarNo As String)
    Dim RS          As ADODB.Recordset
    Dim i           As Integer
    Dim iCnt        As Long
    Dim intRow      As Long
    Dim intCol      As Integer
    Dim strDate     As String
    Dim strChart    As String
    Dim blnSame     As Boolean
    
    If pBarNo = "" Then
        vasID.MaxRows = 0
        intRow = 0
    End If
    
    blnSame = False
    vasID.ReDraw = False
    
'             SQL = "Select C.SPECNO , C.SNAME, C.DEPTCODE, DECODE(C.GBIO,'I','�� �� ','O','�� �� ') as GBIO, B.EXAMNAME,  B.MASTERCODE, B.CHANNEL "
          SQL = " SELECT DISTINCT '1', '' AS SN ,'' AS ����Ͻ�, B.JOBDATE AS ��������" & vbCrLf
    SQL = SQL & ",       C.SPECNO AS ���ڵ��ȣ, C.PTNO AS ��Ʈ��ȣ, C.JOBNO AS ������ȣ, DECODE(C.GBIO,'I','�Կ�','O','�ܷ�') AS �Կ�" & vbCrLf
    SQL = SQL & ", '' AS R, '' AS P, C.SNAME AS �̸�, C.SEX AS ����, C.AGE AS ����, A.MASTERCODE AS ITEM " & vbCrLf
    SQL = SQL & "  From TW_HSP_OCS.TWEXAM_RESULTC A,"
    SQL = SQL & "       TW_HSP_OCS.TWEXAM_MASTER  B,"
    SQL = SQL & "       TW_HSP_OCS.TWEXAM_SPECMST C"
    SQL = SQL & " Where B.JOBDATE BETWEEN '" & pFrDt & "' AND '" & pToDt & "'" & vbCrLf '�۾�����
    SQL = SQL & "   And B.EQUCODE1 = '" & gEquipCode & "'" & vbCrLf                     ' ����ڵ�
    SQL = SQL & "   AND C.STATUS   = '3' " & vbCrLf                                     ' �˻����
    SQL = SQL & "   And (C.SPECNO  = A.SPECNO) " & vbCrLf
    SQL = SQL & "   And (A.MASTERCODE = B.MASTERCODE)"
    SQL = SQL & " ORDER BY ��������, ���ڵ��ȣ, ��Ʈ��ȣ, ������ȣ"

    SetRawData "[Sql]" & SQL

    '-- Record Count ������
    cn_Ser.CursorLocation = adUseClient
    Set RS = cn_Ser.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then
        frmProgress.Show
        frmProgress.ZOrder 0
        frmProgress.Xprog.Min = 1
        frmProgress.Xprog.Max = RS.RecordCount + 1
                
        Do Until RS.EOF
            iCnt = iCnt + 1
            With vasID
                .ReDraw = False
                For i = 1 To .DataRowCnt
                    strDate = GetText(vasID, i, colHOSPDATE)
                    strChart = GetText(vasID, i, colBARCODE)
                    If Trim(RS("��������")) = strDate And Trim(RS("���ڵ��ȣ")) = strChart Then
                        blnSame = True
                    End If
                    For intCol = colState + 1 To vasID.MaxCols
                        If Trim(RS.Fields("ITEM")) = gArrEquip(intCol - colState, 3) Then
                            vasID.Row = .MaxRows
                            vasID.Col = intCol
                            vasID.BackColor = vbYellow
                            Exit For
                        End If
                    Next
                Next
                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    SetText vasID, "1", .MaxRows, colCHECKBOX
                    SetText vasID, Trim(RS.Fields("��������")) & "", .MaxRows, colHOSPDATE
                    SetText vasID, Trim(RS.Fields("���ڵ��ȣ")) & "", .MaxRows, colBARCODE
                    SetText vasID, Trim(RS.Fields("��Ʈ��ȣ")) & "", .MaxRows, colCHARTNO
                    SetText vasID, Trim(RS.Fields("������ȣ")) & "", .MaxRows, colPID
                    SetText vasID, Trim(RS.Fields("�̸�")) & "", .MaxRows, colPNAME
                    SetText vasID, Trim(RS.Fields("����")) & "", .MaxRows, colPSEX
                    SetText vasID, Trim(RS.Fields("����")) & "", .MaxRows, colPAGE
                    For intCol = colState + 1 To vasID.MaxCols
                        If Trim(RS.Fields("ITEM")) = gArrEquip(intCol - colState, 3) Then
                            vasID.Row = .MaxRows
                            vasID.Col = intCol
                            vasID.BackColor = vbYellow
                            Exit For
                        End If
                    Next
                End If
                blnSame = False
            End With
            '-- ���α׷����� ����
            frmProgress.Xprog.Value = iCnt
            DoEvents
            
            RS.MoveNext
        Loop
        chkWAll.Value = "1"
    Else
        StatusBar1.Panels(3).Text = "��ȸ ����ڰ� �����ϴ�."
        chkWAll.Value = "0"
    End If
    
    RS.Close
    '-- ���α׷����� �ݱ�
    Unload frmProgress
    
    vasID.RowHeight(-1) = 12
    vasID.ReDraw = True
    
End Sub


Private Sub GetWorkList_BIT(ByVal pFrDt As String, ByVal pToDt As String, Optional pBarNo As String)
    Dim RS          As ADODB.Recordset
    Dim i           As Integer
    Dim iCnt        As Long
    Dim intRow      As Long
    Dim intCol      As Integer
    Dim strDate     As String
    Dim strChart    As String
    Dim blnSame     As Boolean
    
    If pBarNo = "" Then
        vasID.MaxRows = 0
        intRow = 0
    End If
    
    blnSame = False
    vasID.ReDraw = False
    
    '-- BIT
          SQL = " SELECT DISTINCT '1', '' AS SN ,'' AS ����Ͻ�, SUBSTRING(O.OCMACPDTM,1,8) AS ��������," & vbCrLf
    SQL = SQL & "        R.RESSPMNUM AS ���ڵ��ȣ, O.OCMCHTNUM AS ��Ʈ��ȣ,R.RESOCMNUM AS ������ȣ, '' AS �Կ�," & vbCrLf
    SQL = SQL & "        '' AS R, '' AS P, P.PBSPATNAM AS �̸�, P.PBSSEXTYP AS ����,'' AS ����, '' AS ITEM" & vbCrLf
    SQL = SQL & "   FROM DRBITPACK..RESINF AS R, DRBITPACK..OCMINF AS O, DRBITPACK..PBSINF AS P, DRBITPACK..LABMST AS E, DRBITPACK..ODRINF AS W" & vbCrLf
    SQL = SQL & " WHERE O.OCMACPDTM BETWEEN '" & pFrDt & "000000" & "' AND '" & pToDt & "235959" & "'" & vbCrLf
    SQL = SQL & "   AND O.OCMCOMSTT NOT IN ('CN', 'CR', 'VC')" & vbCrLf
    SQL = SQL & "   AND R.RESLABCOD IN (" & gAllExam & ")" & vbCrLf
    SQL = SQL & "   AND R.RESOCMNUM = O.OCMNUM" & vbCrLf
    SQL = SQL & "   AND O.OCMCHTNUM = P.PBSCHTNUM" & vbCrLf
    SQL = SQL & "   AND R.RESOCMNUM = W.ODROCMNUM" & vbCrLf
    SQL = SQL & "   AND R.RESLABCOD = W.ODRCOD" & vbCrLf
    SQL = SQL & "   AND R.RESLABCOD = E.LABCOD" & vbCrLf
    '-- ���������
    If chkSaveAll.Value = "0" Then
        SQL = SQL & "   AND (R.RESREPTYP IS NULL OR R.RESREPTYP <> 'F') " & vbCrLf         '--  'I':�߰� 'F' �Ϸ�"
        SQL = SQL & "   AND W.ODRDELFLG = 'N'" & vbCrLf
        SQL = SQL & "   AND (R.RESRLTVAL = ''  OR R.RESRLTVAL IS NULL)" & vbCrLf
    End If
    SQL = SQL & " ORDER BY �����Ͻ�, ��Ʈ��ȣ, ������ȣ"


    '-- Record Count ������
    cn_Ser.CursorLocation = adUseClient
    Set RS = cn_Ser.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then
        frmProgress.Show
        frmProgress.ZOrder 0
        frmProgress.Xprog.Min = 1
        frmProgress.Xprog.Max = RS.RecordCount + 1
                
        Do Until RS.EOF
            iCnt = iCnt + 1
            With vasID
                .ReDraw = False
                For i = 1 To .DataRowCnt
                    strDate = GetText(vasID, i, colHOSPDATE)
                    strChart = GetText(vasID, i, colBARCODE)
                    If Trim(RS("��������")) = strDate And Trim(RS("���ڵ��ȣ")) = strChart Then
                        blnSame = True
                    End If
                    For intCol = colState + 1 To vasID.MaxCols
                        If Trim(RS.Fields("ITEM")) = gArrEquip(intCol - colState, 3) Then
                            vasID.Row = .MaxRows
                            vasID.Col = intCol
                            vasID.BackColor = vbYellow
                            Exit For
                        End If
                    Next
                Next
                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    SetText vasID, "1", .MaxRows, colCHECKBOX
                    SetText vasID, Trim(RS.Fields("��������")) & "", .MaxRows, colHOSPDATE
                    SetText vasID, Trim(RS.Fields("���ڵ��ȣ")) & "", .MaxRows, colBARCODE
                    SetText vasID, Trim(RS.Fields("��Ʈ��ȣ")) & "", .MaxRows, colCHARTNO
                    SetText vasID, Trim(RS.Fields("������ȣ")) & "", .MaxRows, colPID
                    SetText vasID, Trim(RS.Fields("�̸�")) & "", .MaxRows, colPNAME
                    SetText vasID, Trim(RS.Fields("����")) & "", .MaxRows, colPSEX
                    SetText vasID, Trim(RS.Fields("����")) & "", .MaxRows, colPAGE
                    For intCol = colState + 1 To vasID.MaxCols
                        If Trim(RS.Fields("ITEM")) = gArrEquip(intCol - colState, 3) Then
                            vasID.Row = .MaxRows
                            vasID.Col = intCol
                            vasID.BackColor = vbYellow
                            Exit For
                        End If
                    Next
                End If
                blnSame = False
            End With
            '-- ���α׷����� ����
            frmProgress.Xprog.Value = iCnt
            DoEvents
            
            RS.MoveNext
        Loop
        chkWAll.Value = "1"
    Else
        StatusBar1.Panels(3).Text = "��ȸ ����ڰ� �����ϴ�."
        chkWAll.Value = "0"
    End If
    
    RS.Close
    '-- ���α׷����� �ݱ�
    Unload frmProgress
    
    vasID.RowHeight(-1) = 12
    vasID.ReDraw = True
    
End Sub

Private Sub GetWorkList_GINUSDLL(ByVal pFrDt As String, ByVal pToDt As String, Optional pBarNo As String)
    Dim RS          As ADODB.Recordset
    Dim i           As Integer
    Dim iCnt        As Long
    Dim intRow      As Long
    Dim intCol      As Integer
    Dim strDate     As String
    Dim strChart    As String
    Dim blnSame     As Boolean
    
    '-- ������
    Dim strRequest  As String
    Dim strResponse As String
    Dim varResponse As Variant
    
    If pBarNo = "" Then
        vasID.MaxRows = 0
        intRow = 0
    End If
    
    blnSame = False
    vasID.ReDraw = False
    
    '-- �˻����� ��������
                 strRequest = "jobs" + vbTab + "L" + vbTab
    strRequest = strRequest & "hos_org_no" + vbTab + gGINUS_Parm.HCD + vbTab
    strRequest = strRequest & "fr_ymd" + vbTab + pFrDt + vbTab
    strRequest = strRequest & "to_ymd" + vbTab + pToDt + vbTab
    strRequest = strRequest & "mach_cd" + vbTab + gGINUS_Parm.HCD + vbTab
    strRequest = strRequest & "smp_no" + vbTab + "%" + vbTab + vbCr
    
    strResponse = W2ACALL2("SCC0191A", strRequest, gGINUS_Parm.URL) '-- ���ڵ�� �˻��� ��ȸ(https://211.172.17.66)
    
    strResponse = Mid(strResponse, 90)
    varResponse = Split(strResponse, vbLf)
    
    If UBound(varResponse) > 0 Then
        chkWAll.Value = 1
    Else
        chkWAll.Value = 0
    End If
    
    For i = 0 To UBound(varResponse) - 1
        frmProgress.Show
        frmProgress.ZOrder 0
        frmProgress.Xprog.Min = 1
        frmProgress.Xprog.Max = UBound(varResponse) - 1
        With vasID
            If .MaxRows = 0 Then
                .MaxRows = .MaxRows + 1
                intRow = .MaxRows
                
                SetText vasID, "1", intRow, colCHECKBOX
                SetText vasID, Mid(mGetP(varResponse(i), 5, vbTab), 1, 8), intRow, colHOSPDATE  '-- ��������
                SetText vasID, mGetP(varResponse(i), 2, vbTab), intRow, colBARCODE              '-- ���ڵ��ȣ
                SetText vasID, mGetP(varResponse(i), 6, vbTab), intRow, colPID                  '-- ������ȣ
                SetText vasID, mGetP(varResponse(i), 7, vbTab), intRow, colPNAME                '-- �̸�
                Select Case mGetP(varResponse(i), 13, vbTab)                                    '-- ��/��
                    Case "O": SetText vasID, "�ܷ�", intRow, colINOUT
                    Case "E": SetText vasID, "����", intRow, colINOUT
                    Case "I": SetText vasID, "�Կ�", intRow, colINOUT
                End Select
                Call SetOrderColor(mGetP(varResponse(i), 2, vbTab), intRow)
            Else
                '-- ���� ���ڵ� ��ȣ�� �ִ��� üũ..
                intRow = GetSameRowNum(Trim(mGetP(varResponse(i), 2, vbTab)))
                If intRow = 0 Then
                    .MaxRows = .MaxRows + 1
                    intRow = .MaxRows
                    
                    SetText vasID, "1", intRow, colCHECKBOX
                    SetText vasID, Mid(mGetP(varResponse(i), 5, vbTab), 1, 8), intRow, colHOSPDATE  '-- ��������
                    SetText vasID, mGetP(varResponse(i), 2, vbTab), intRow, colBARCODE              '-- ���ڵ��ȣ
                    SetText vasID, mGetP(varResponse(i), 6, vbTab), intRow, colPID                  '-- ������ȣ
                    SetText vasID, mGetP(varResponse(i), 7, vbTab), intRow, colPNAME                '-- �̸�
                    Select Case mGetP(varResponse(i), 13, vbTab)                                    '-- ��/��
                        Case "O": SetText vasID, "�ܷ�", intRow, colINOUT
                        Case "E": SetText vasID, "����", intRow, colINOUT
                        Case "I": SetText vasID, "�Կ�", intRow, colINOUT
                    End Select
                    Call SetOrderColor(mGetP(varResponse(i), 2, vbTab), intRow)
                End If
            End If
        End With
        
        '-- ���α׷����� ����
        frmProgress.Xprog.Value = i + 1
        DoEvents
        
    Next
    
    '-- ���α׷����� �ݱ�
    Unload frmProgress
    
    vasID.RowHeight(-1) = 12
    vasID.ReDraw = True
    
End Sub


Private Sub SetOrderColor(ByVal pBarNo As String, ByVal pRow As Integer)
    Dim i       As Integer
    Dim intCol  As Integer
    Dim strItem As String
    
    '-- ������
    Dim strRequest  As String
    Dim strResponse As String
    Dim varResponse As Variant
    
    
    '-- �˻�ITEM ��������
                 strRequest = "jobs" + vbTab + "Q" + vbTab
    strRequest = strRequest & "hos_org_no" + vbTab + gGINUS_Parm.HCD + vbTab
    strRequest = strRequest & "smp_no" + vbTab + pBarNo + vbTab
    strRequest = strRequest & "mach_cd" + vbTab + gGINUS_Parm.MCD + vbTab + vbCr
    
    strResponse = W2ACALL2("SCC0191A", strRequest, gGINUS_Parm.URL) '-- ���ڵ�� �˻��� ��ȸ(https://211.172.17.66)
    strResponse = Mid(strResponse, 90)
    varResponse = Split(strResponse, vbLf)
    
    If UBound(varResponse) > 0 Then
        For i = 0 To UBound(varResponse) - 1
            For intCol = colState + 1 To vasID.MaxCols
                If mGetP(varResponse(i), 6, vbTab) = gArrEquip(intCol - colState, 3) Then
                    vasID.Row = pRow
                    vasID.Col = intCol
                    vasID.BackColor = vbYellow
                    '-- �������� SEQ
                    gArrEquip(intCol - colState, 7) = mGetP(varResponse(i), 3, vbTab) & "|" & mGetP(varResponse(i), 4, vbTab) & "|" & mGetP(varResponse(i), 5, vbTab)
                    Exit For
                End If
            Next intCol
        Next i
    Else
        SetText vasID, "No Order", pRow, colState
    End If
    
End Sub

Private Sub cmdSearch_Click()
                
    Select Case gOCS
        Case "BIT":         Call GetWorkList_BIT(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))
        Case "TWIN":        Call GetWorkList_TWIN(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))
        Case "DADESOFT":    Call GetWorkList_DADESOFT(Format(dtpStartDt.Value, "yyyy-mm-dd"), Format(dtpStopDt.Value, "yyyy-mm-dd"))
        Case "GINUSDLL":    Call GetWorkList_GINUSDLL(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))
        Case "GINUSDB":     Call GetWorkList(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))
        Case "BITSMALL":    Call GetWorkList(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))
        Case "BITLARGE":    Call GetWorkList(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))
        Case "MEDICHART":   Call GetWorkList(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))
        Case "JBUNIV":      Call GetWorkList_JBUNIV(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))
        Case "SUNGMO":
                            'Call GetWorkList_UBCARE(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))
                            frmPatSearch.Show vbModal
    End Select
    
    vasID.RowHeight(-1) = 12
    vasRes.MaxRows = 0
    
End Sub


Private Sub GetWorkList_JBUNIV(ByVal pFrDt As String, ByVal pToDt As String, Optional pBarNo As String)
    Dim RS          As ADODB.Recordset
    Dim i           As Integer
    Dim iCnt        As Long
    Dim intRow      As Long
    Dim intCol      As Integer
    Dim strDate     As String
    Dim strChart    As String
    Dim blnSame     As Boolean
    
    If pBarNo = "" Then
        vasID.MaxRows = 0
        intRow = 0
    End If
    
    blnSame = False
    vasID.ReDraw = False
    
    '-- ���ϴ뺴��  r010m.SPCCD
    SQL = ""
    SQL = SQL & " SELECT '1', '' AS SN ,'' AS ����Ͻ�, j011m.colldt AS ��������, j011m.bcno AS ���ڵ��ȣ, j010m.bcprtno AS ��Ʈ��ȣ" & vbCr
    SQL = SQL & "       , r010m.WKYMD||r010m.WKGRPCD||r010m.WKNO FLWKNO " & vbCr
    SQL = SQL & "       , r010m.WKNO AS ������ȣ " & vbCr
    SQL = SQL & "       , j011m.regno AS ������ȣ " & vbCr
    SQL = SQL & "       , j010m.patnm AS �̸� " & vbCr
    SQL = SQL & "       , j010m.age AS ���� " & vbCr
    SQL = SQL & "       , j010m.sex AS ���� " & vbCr
    SQL = SQL & "       , j011m.IOGBN  " & vbCr
    SQL = SQL & "       , j010m.DEPTCD " & vbCr
    SQL = SQL & "       , j010m.WARDNO " & vbCr
    SQL = SQL & "       , j010m.ROOMNO " & vbCr
    SQL = SQL & "       , f72m.testcd AS ITEM " & vbCr
    SQL = SQL & "       , r010m.SPCCD AS SPCCD " & vbCr
    SQL = SQL & "  FROM LJ011M j011m                                     " & vbCr
    SQL = SQL & "       INNER JOIN LJ010M j010m                          " & vbCr
    SQL = SQL & "               ON j011m.bcno  = j010m.bcno              " & vbCr
    SQL = SQL & "              AND j011m.regno = j010m.regno             " & vbCr
    SQL = SQL & "       INNER JOIN LR010M r010m                          " & vbCr
    SQL = SQL & "               ON j011m.bcno   = r010m.bcno             " & vbCr
    SQL = SQL & "              AND j011m.regno  = r010m.regno            " & vbCr
    SQL = SQL & "              AND NVL(r010m.rstflg,'0') = '0'       " & vbCr
    SQL = SQL & "       INNER JOIN LF072M f72m                           " & vbCr
    SQL = SQL & "               ON f72m.eqcd    = '" & gEquipCode & "' " & vbCr
'    SQL = SQL & "              AND f72m.testcd  = '" & mGetP(cboTest.Text, 2, "|") & "'   " & vbCr
    SQL = SQL & "              AND r010m.testcd = f72m.testcd            " & vbCr
    SQL = SQL & " WHERE j011m.colldt BETWEEN '" & pFrDt & "000000" & "' AND '" & pToDt & "235959" & "'  " & vbCr
'    SQL = SQL & "   and r010m.wkno between '" & txtStartNum.Text & "' AND '" & txtStopNum.Text & "' " & vbCr
    SQL = SQL & "   AND j011m.spcflg  = '4'                        " & vbCr
    SQL = SQL & "   AND NVL(j011m.rstflg, '0')  = '0'            " & vbCr
    SQL = SQL & " UNION                                              " & vbCr
    SQL = SQL & " SELECT '1', '' AS SN ,'' AS ����Ͻ�, j011m.colldt AS ��������, j011m.bcno AS ���ڵ��ȣ, j010m.bcprtno AS ��Ʈ��ȣ " & vbCr
    SQL = SQL & "        , r010m.FLWKNO " & vbCr
    SQL = SQL & "        , r010m.WKNO AS ������ȣ " & vbCr
    SQL = SQL & "        , j011m.regno AS ������ȣ " & vbCr
    SQL = SQL & "        , j010m.patnm AS �̸� " & vbCr
    SQL = SQL & "        , j010m.age AS ���� " & vbCr
    SQL = SQL & "        , j010m.sex AS ���� " & vbCr
    SQL = SQL & "        , j011m.IOGBN " & vbCr
    SQL = SQL & "        , j010m.DEPTCD " & vbCr
    SQL = SQL & "        , j010m.WARDNO " & vbCr
    SQL = SQL & "        , j010m.ROOMNO " & vbCr
    SQL = SQL & "       , f72m.testcd AS ITEM " & vbCr
    SQL = SQL & "       , r010m.SPCCD AS SPCCD " & vbCr
    SQL = SQL & "   FROM LJ011M j011m                                " & vbCr
    SQL = SQL & "        INNER JOIN LJ010M j010m                     " & vbCr
    SQL = SQL & "                ON j011m.bcno  = j010m.bcno         " & vbCr
    SQL = SQL & "               AND j011m.regno = j010m.regno        " & vbCr
    SQL = SQL & "        INNER JOIN LM010M r010m                     " & vbCr
    SQL = SQL & "                ON j011m.bcno   = r010m.bcno        " & vbCr
    SQL = SQL & "               AND j011m.regno  = r010m.regno       " & vbCr
    SQL = SQL & "               AND NVL(r010m.rstflg,'0') = '0'  " & vbCr
    SQL = SQL & "        INNER JOIN LF072M f72m                      " & vbCr
    SQL = SQL & "                ON f72m.eqcd    = '" & gEquipCode & "' " & vbCr
'    SQL = SQL & "                AND f72m.testcd  = '" & mGetP(cboTest.Text, 2, "|") & "'  " & vbCr
    SQL = SQL & "               AND r010m.testcd = f72m.testcd       " & vbCr
'    SQL = SQL & "  WHERE j011m.colldt BETWEEN '" & pFrDt & "000000" & "' AND '" & pToDt & "235959" & "'  " & vbCr
'    SQL = SQL & "   and r010m.wkno BETWEEN '" & txtStartNum.Text & "' AND '" & txtStopNum.Text & "' " & vbCr
    SQL = SQL & "    AND j011m.spcflg  = '4'               " & vbCr
    SQL = SQL & "    AND NVL(j011m.rstflg, '0')  = '0'     " & vbCr
    SQL = SQL & "    ORDER BY FLWKNO  " & vbCr

 '   SetRawData "[SQL]" & SQL

    '-- Record Count ������
    cn_Ser.CursorLocation = adUseClient
    Set RS = cn_Ser.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then
        frmProgress.Show
        frmProgress.ZOrder 0
        frmProgress.Xprog.Min = 1
        frmProgress.Xprog.Max = RS.RecordCount + 1
        
        Do Until RS.EOF
            iCnt = iCnt + 1
            With vasID
                .ReDraw = False
                For i = 1 To .DataRowCnt
                    strDate = GetText(vasID, i, colHOSPDATE)
                    strChart = GetText(vasID, i, colBARCODE)
                    If Trim(RS("��������")) = strDate And Trim(RS("���ڵ��ȣ")) = strChart Then
                        blnSame = True
                    End If
                    For intCol = colState + 1 To vasID.MaxCols
                        If Trim(RS.Fields("ITEM")) = gArrEquip(intCol - colState, 3) Then
                            vasID.Row = .MaxRows
                            vasID.Col = intCol
                            vasID.BackColor = vbYellow
                            Exit For
                        End If
                    Next
                Next
                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    SetText vasID, "1", .MaxRows, colCHECKBOX
                    SetText vasID, Trim(RS.Fields("��������")) & "", .MaxRows, colHOSPDATE
                    SetText vasID, Trim(RS.Fields("���ڵ��ȣ")) & "", .MaxRows, colBARCODE
                    SetText vasID, Trim(RS.Fields("��Ʈ��ȣ")) & "", .MaxRows, colCHARTNO
                    SetText vasID, Trim(RS.Fields("������ȣ")) & "", .MaxRows, colPID
                    SetText vasID, Trim(RS.Fields("�̸�")) & "", .MaxRows, colPNAME
                    SetText vasID, Trim(RS.Fields("����")) & "", .MaxRows, colPSEX
                    SetText vasID, Trim(RS.Fields("����")) & "", .MaxRows, colPAGE
                    SetText vasID, Trim(RS.Fields("SPCCD")) & "", .MaxRows, colINOUT
                    
                    '.MaxRows = .MaxRows + 1
                    
                    SetText vasID, txtRack.Text, .MaxRows, colDISKNO
                    SetText vasID, txtPos.Text, .MaxRows, colPOSNO
                    
                    For intCol = colState + 1 To vasID.MaxCols
                        If Trim(RS.Fields("ITEM")) = gArrEquip(intCol - colState, 3) Then
                            vasID.Row = .MaxRows
                            vasID.Col = intCol
                            vasID.BackColor = vbYellow
                            Exit For
                        End If
                    Next
                End If
                blnSame = False
            End With
            '-- ���α׷����� ����
            frmProgress.Xprog.Value = iCnt
            DoEvents
            
            txtPos.Text = Chr(Asc(txtPos.Text) + 1)
            If txtPos.Text = "I" Then
                txtPos.Text = "A"
                txtRack.Text = txtRack.Text - 1
            End If
            
            If txtRack.Text = "1" And txtPos.Text = "H" Then
                txtRack.Text = "5"
                txtPos.Text = "A"
            End If
            
            
            RS.MoveNext
        Loop
        chkWAll.Value = "1"
    Else
        StatusBar1.Panels(3).Text = "��ȸ ����ڰ� �����ϴ�."
        chkWAll.Value = "0"
    End If
    
    RS.Close
    '-- ���α׷����� �ݱ�
    Unload frmProgress
    
    vasID.RowHeight(-1) = 12
    vasID.ReDraw = True
    
End Sub

Private Sub cmdSL_Click()
    If cmdSL.Caption = "��" Then
        cmdSL.Caption = "��"
        vasID.Width = 18285 '18075 '15225
    Else
        cmdSL.Caption = "��"
        vasID.Width = 11535 '11355 '8475
    End If

End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Resize()
    On Error Resume Next

    If frmInterface.ScaleHeight = 0 Then Exit Sub
    
        
    If cmdSL.Caption = "��" Then
        Frame1.Height = frmInterface.ScaleHeight - (Picture2.Top) - 1200
        vasID.Height = Frame1.Height - 300
        
        Frame1.Width = frmInterface.ScaleWidth - 200
        vasID.Width = frmInterface.ScaleWidth - 7300
        
    
        Frame6.Left = vasID.Width + 300
        vasRes.Height = vasID.Height - 550
        vasRes.Left = Frame6.Left
    Else
        Frame1.Height = frmInterface.ScaleHeight - (Picture2.Top) - 1200
        vasID.Height = Frame1.Height - 300
        
        Frame1.Width = frmInterface.ScaleWidth - 200
        vasID.Width = frmInterface.ScaleWidth - 300
    
        'Frame6.Left = frmInterface.ScaleWidth - vasID.Width
        'vasRes.Height = vasID.Height - 550
        'vasRes.Left = Frame6.Left
    
    End If
    
    Picture2.Width = Frame1.Width
    
    StatusBar1.Panels(3).Width = Frame1.Width - 8500
End Sub

'Private Sub cmdTest_Click()
'    Dim adoRS2      As ADODB.Recordset
'    Dim strKeyno    As String
'    Dim intRow      As Integer
'    Dim intCnt      As Integer
'
'
'    Set adoRS2 = New ADODB.Recordset
'    Set adoRS2 = adoExecQuery50P("SLRTRM50P", Format(dtpFromDt.Value, "yyyymmdd"), gEquipCode, Val(txtFromWN.Text), Val(txtToWN.Text), Space$(5), "0", "")
'
'    spdOrder.MaxRows = 0
'    spdTot.MaxRows = 0
'
'    If adoRS2.RecordCount <= 0 Then
'        adoRS2.Close: Set adoRS2 = Nothing
'        Exit Sub
'    End If
'
'    If Not adoRS2.EOF Then
'        Text1.Text = ""
'
'        Do While Not adoRS2.EOF
'            For intRow = 0 To adoRS2.Fields.Count
'                If intRow = 0 Then
'                    Text1.Text = Text1.Text & adoRS2.Fields(intRow).Name
'                End If
'                Text1.Text = Text1.Text & adoRS2.Fields(intRow).Value & "|"
'            Next
'            adoRS2.MoveNext
'        Loop
'
'    End If
'
'    adoRS2.Close: Set adoRS2 = Nothing
'
'End Sub

Private Sub imgPort_DblClick()
    
    '-- ���߽ÿ��� Remark Ǯ� �׽�Ʈ����
    If FrmHideControl.Visible = True Then
        Me.Width = 16545
        FrmHideControl.Visible = False
    Else
        Me.Width = 22000
        FrmHideControl.Visible = True
    End If

End Sub




Private Sub Label1_DblClick(Index As Integer)

    If chkSave.Enabled = True Then
        chkSave.Enabled = False
    Else
        chkSave.Enabled = True
    End If
    
End Sub

'Private Sub Label1_Click(Index As Integer)
'
'    Call comEqp_OnComm
'
'End Sub

Private Sub lblclear_Click()
    lblChangePID.Caption = ""
    lblChangeBar.Caption = ""
    lblBarcode(0).Caption = ""
    lblPname(0).Caption = ""
    lblSaveSeq.Caption = ""
    lblExamDate.Caption = ""
End Sub

Private Sub Command16_Click()
    
'    strBuffer = ":N1    80 81                 00620141422      15 1   7.0  2   4.1  3   0.5  4   4.5  5    34  6    20  7   417  8   239  9    97 14    85 15    14 16   0.7 18    93 19      T54     1 "
'
'    strBuffer = txtTest.Text
'
'    Call comEqp_OnComm
        
        
        
    
    strBuffer = ""
    strBuffer = strBuffer & "MSH|^~\&|Mindray|BS-400|||20100608151653||ORU^R01|1|P|2.3.1||||0||ASCII|||" & vbCr
    strBuffer = strBuffer & "PID|8||||||19000101000000||||||||||||||||||||||||" & vbCr
    strBuffer = strBuffer & "OBR|8||1|Mindray^BS-400|N||20100608120913|||||||20100608115230|Serum|||||||||||||||||||||||||||||||||" & vbCr
    strBuffer = strBuffer & "OBX|1|NM|4|GOT|77.495280|U/L|1.000000-40.000000|H|||F||77.495280|20100608120913||||" & vbCr
    strBuffer = strBuffer & "OBX|2|NM|2|GPT|153.026287|U/L|1.000000-42.000000|H|||F||153.026287|20100608120913||||" & vbCr
    strBuffer = strBuffer & "OBX|3|NM|18|TP|7.132263|mg/dL|6.300000-8.400000|N|||F||7.132263|20100608120913||||" & vbCr
    strBuffer = strBuffer & "OBX|4|NM|1|ALB|4.570356|mg/dL|3.800000-5.300000|N|||F||4.570356|20100608120913||||" & vbCr
    strBuffer = strBuffer & "OBX|5|NM|16|T_BIL|1.426674|mg/dL|0.100000-1.200000|H|||F||1.426674|20100608120913||||" & vbCr
    strBuffer = strBuffer & "OBX|6|NM|6|ALP|764.171220|U/L|100.000000-290.000000|H|||F||764.171220|20100608120913||||" & vbCr
    strBuffer = strBuffer & "OBX|7|NM|5|BUN|12.854416|mg/dL|9.000000-23.000000|N|||F||12.854416|20100608120913||||" & vbCr
    strBuffer = strBuffer & "OBX|8|NM|10|CREA|0.948916|mg/dL|0.700000-1.500000|N|||F||0.948916|20100608120913||||" & vbCr
    
    Call ComReceive(strBuffer)
                

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
    'Me.Width = 16545
    
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
    
    frmInterface.StatusBar1.Panels(1).Text = gUserID
        
    cboChk.ListIndex = 0
    
    If Not Connect_Local Then
        MsgBox "������� �ʾҽ��ϴ�."
        cn_Local_Flag = False
        Exit Sub
    Else
        cn_Local_Flag = True
    End If
    
    '-- ��������
'    For i = 1 To 1
'        If Not Connect_PRServer Then
'            MsgBox "���������� ������� �ʾҽ��ϴ�."
'            cn_Server_Flag = False
'            Exit Sub
'        Else
'            cn_Server_Flag = True
'        End If
'    Next
    
    '-- EDM
    For i = 1 To 1
        If Not Connect_DRServer Then
            MsgBox "EDM�� ������� �ʾҽ��ϴ�."
            cn_Server_Flag = False
            Exit Sub
        Else
            cn_Server_Flag = True
        End If
    Next
    
    
    
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
    
    tmrResult.Interval = 1000
    tmrResult.Enabled = True

    Call cmdSL_Click
'    Winsock1.LocalPort = CInt(5150)  '5150 5051
'    Winsock1.Listen

'    Call cmdSL_Click
    
'    StatusBar1.Panels(2).Text = Winsock1.LocalIP
    '-- test
'    vasID.MaxRows = 10
    
  '  dtpFromDt = Format(CDate(Date), "yyyy/mm/dd")
  '  txtFromWN = "1"
  '  txtToWN = "99999"
    
End Sub

Private Sub SetExamCode()
    Dim i As Integer
    
    
    With vasID
        .MaxCols = colState + UBound(gArrEquip)
        
        For i = 0 To UBound(gArrEquip) - 1
            .Col = colState + (i + 1)
            .Row = -1
            .CellType = CellTypeEdit
            '.TypeEditCharSet = TypeEditCharSetAlphanumeric
            '.TypeEditCharCase = TypeEditCharCaseSetUpper
            
            .TypeEditCharSet = TypeEditCharSetASCII
            .TypeEditCharCase = TypeEditCharCaseSetNone
            
            .TypeHAlign = TypeHAlignCenter
            .TypeVAlign = TypeVAlignCenter
            'Call SetText(vasID, gArrEquip(i + 1, 2), 0, colState + (i + 1))
            Call SetText(vasID, gArrEquip(i + 1, 4), 0, colState + (i + 1))
            .ColWidth(colState + (i + 1)) = 9
            
            'cboTest.AddItem gArrEquip(i + 1, 4) & Space(20) & "|" & gArrEquip(i + 1, 3)
        Next
        
'        cboTest.ListIndex = 0
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


Private Function ShowOpen(Ufilter As String, Upath As String) As String
    
    OFName.lStructSize = Len(OFName)
    OFName.hwndOwner = Me.hWnd
    OFName.hInstance = App.hInstance
    OFName.lpstrFilter = Ufilter
    OFName.lpstrFile = Space$(254)
    OFName.nMaxFile = 255
    OFName.lpstrFileTitle = Space$(254)
    OFName.nMaxFileTitle = 255
    OFName.lpstrInitialDir = Upath
    OFName.lpstrTitle = "Open File"
    OFName.flags = 0

    If GetOpenFileName(OFName) Then
        ShowOpen = Trim$(OFName.lpstrFile)
        'ShowOpen = Mid(ShowOpen, 1, Len(ShowOpen) - 1)
    Else
        ShowOpen = ""
    End If
    
End Function


Public Function Get_ExamCode(asExamName As String) As String
    Dim strExamName As String
    Get_ExamCode = ""
    strExamName = Trim(asExamName)
        
    SQL = "SELECT EXAMCODE FROM EQUIPEXAM WHERE EQUIPCODE = '" & asExamName & "'"
    Res = db_select_Col(gLocal, SQL)
    
    If Res < 0 Then
        SaveQuery SQL
        Exit Function
    End If
    
    Get_ExamCode = Trim(gReadBuf(0))
    
End Function

Public Function Get_ExamName(asExamCode As String) As String
    Dim strExamCode As String
    Get_ExamName = ""
    strExamCode = Trim(asExamCode)
        
    SQL = "SELECT EXAMNAME FROM EQUIPEXAM WHERE EXAMCODE = '" & strExamCode & "'"
    Res = db_select_Col(gLocal, SQL)
    
    If Res < 0 Then
        SaveQuery SQL
        Exit Function
    End If
    
    Get_ExamName = Trim(gReadBuf(0))
    
End Function




Public Function Get_EquipCode(asExamCode As String) As String
    Dim strExamCode As String
    Get_EquipCode = ""
    strExamCode = Trim(asExamCode)
        
    SQL = "SELECT EQUIPCODE FROM EQUIPEXAM WHERE EXAMCODE = '" & strExamCode & "'"
    Res = db_select_Col(gLocal, SQL)
    
    If Res < 0 Then
        SaveQuery SQL
        Exit Function
    End If
    
    Get_EquipCode = Trim(gReadBuf(0))
    
End Function

Private Sub Form_Unload(Cancel As Integer)
    If comEqp.PortOpen = True Then
        comEqp.PortOpen = False
    End If

'    Call dce_close_env      ' Server�� ������ ���� ��
'    DisConnect_Server
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

Private Sub MnPrintLand_Click()

    vasID.PrintOrientation = PrintOrientationLandscape '�������
    vasID.Action = 13

End Sub

Private Sub MnPrintPort_Click()

    vasID.PrintOrientation = PrintOrientationPortrait '�������
    vasID.Action = 13

End Sub

'Private Sub MnScr1_Click()
'    MnScr1.Checked = True
'    MnScr2.Checked = False
'
'    gScreen = "�и�"
'    Call WritePrivateProfileString("config", "IFScreen", gScreen, App.Path & "\Interface.ini")
'
'End Sub
'
'Private Sub MnScr2_Click()
'    MnScr1.Checked = False
'    MnScr2.Checked = True
'
'    gScreen = "����"
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
'   ��� : �������� ����
'-----------------------------------------------------------------------------'
Private Sub SendOrder()
    Dim strOutput As String     '�۽��� ������
    
    '-- ASTM TYPE�� Define �ؾ���.
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
                If mOrder.IsSending = False Then   '## ���� ������
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
                Else                        '## ���� ���ڿ��� ������
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
'   ��� : �ش� ���ڿ��� CheckSum�� ����
'   �μ� :
'       - pMsg : ���ڿ�
'   ��ȯ : CheckSum
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

'-- ���ݳ�¥�� �˻����� ���Ѵ�
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

Private Sub tmrResult_Timer()
    Dim RSH          As ADODB.Recordset
    Dim RSD          As ADODB.Recordset
    
    Dim strFDT As String
    Dim strTDT As String
    Dim pPtID       As String
    Dim pBarNo       As String
    Dim intRow      As Integer
    Dim strIntBase   As String   '������ ������ �˻��
    Dim strResult    As String   '������ ���(����)
    Dim strIntResult As String   '������ ���(����)
    Dim lsExamCode As String
    Dim lsExamName As String
    Dim lsSeqNo As String
    Dim lsResRow    As String
    Dim lsEquipRes As String
    Dim lsResult_Buff As String
    Dim strComm    As String
    Dim strEDTM    As String
    Dim strSpcCd   As String
    Dim strTemperature   As String
    Dim strFiO2     As String
    
    Dim intCol     As Integer
    Dim strChartID  As String
    Dim strMaxSeq   As String
    Dim strQCID      As String
    
    
    txtTimer.Text = txtTimer.Text - 1
    If txtTimer.Text = "0" Then
        
        txtTimer.Enabled = False
        txtTimer.Text = "60"
        
        dtpStopDt.Value = Now
        
        strFDT = Format(dtpStartDt.Value, "yyyy-MM-dd HH:mm:00")
        strTDT = Format(dtpStopDt.Value, "yyyy-MM-dd 23:59:59")
             
        SQL = ""
        SQL = SQL & "Select M.TEST_ID, M.PatientOrLotID, CONVERT(CHAR(19), M.TESTDT, 20) AS EDTM, M.EDMUploadDT, M.OperatorID, M.HostSerialNumber, M.HostAlias, S.LISStatus, S.SendResultMessage,M.Department  " & vbNewLine
        SQL = SQL & "  From Tests M, TestSendLog S " & vbNewLine
        SQL = SQL & " Where M.TestStatus in ('Success', 'iQC') " & vbNewLine ''Incomplete'
        SQL = SQL & "   And M.Test_ID = S.EPOCTestID   " & vbNewLine
        SQL = SQL & "   And M.TestDT between '" & strFDT & "' and '" & strTDT & "'" & vbNewLine
        If gDepartment <> "" Then
            SQL = SQL & "   And M.Department = '" & gDepartment & "'" & vbNewLine
        End If
        
        If gHostSN <> "" Then
            SQL = SQL & "   And M.HostSerialNumber = '" & gHostSN & "'" & vbNewLine
        End If
        
        If chkSave.Value = "0" Then
            SQL = SQL & "   And s.LISStatus < 8 " & vbNewLine
        End If
        SQL = SQL & " Order by EDTM                  "
        
        Call SetSQLData("EDM_LIST", SQL)

        Set RSH = cn_Ser_Census.Execute(SQL, , 1)
    
        If Not RSH.EOF = True And Not RSH.BOF = True Then
            Do Until RSH.EOF
                vasID.MaxRows = vasID.MaxRows + 1
                '-- ���� Row
                intRow = vasID.MaxRows
                
                '-- �׽�Ʈ��ȣ
                pPtID = Trim(RSH.Fields("TEST_ID")) & ""
                '-- ȯ�ڹ�ȣ
                strChartID = Trim(RSH.Fields("PatientOrLotID")) & ""
                strQCID = ""
                '-- �˻�ð�
                strEDTM = Trim(RSH.Fields("EDTM")) & ""
                
                '-- ��ü����
                SQL = "Select Value From TestAttributes  "
                SQL = SQL & " Where Test_Id = '" & pPtID & "'  "
                SQL = SQL & "   And TestAttrName = 'Sample type' "
                Res = GetDBSelectColumn(gServer_Census, SQL)
                strSpcCd = Trim(gReadBuf(0))
                If Trim(strSpcCd) = "Venous" Then
                    '������
                    strSpcCd = "VBGA(OR)"
                Else
                    '������
                    strSpcCd = "ABGA(OR)"
                End If

                '-- ü��
                strTemperature = ""
                SQL = "Select Value From TestAttributes  "
                SQL = SQL & " Where Test_Id = '" & pPtID & "'  "
                SQL = SQL & "   And TestAttrName = 'Patient temperature' "
                Res = GetDBSelectColumn(gServer_Census, SQL)
                'Call SetSQLData("ü��", SQL)
                strTemperature = Trim(gReadBuf(0))

                '-- FiO2
                strFiO2 = ""
                SQL = "Select Value From TestAttributes  "
                SQL = SQL & " Where Test_Id = '" & pPtID & "'  "
                SQL = SQL & "   And TestAttrName = 'FiO2' "
                Res = GetDBSelectColumn(gServer_Census, SQL)
                strFiO2 = Trim(gReadBuf(0))
                
                strMaxSeq = getMaxTestNum(Format(dtpToday, "yyyymmdd"))
                
                SetText vasID, Format(strEDTM, "yyyymmddhhmmss"), intRow, colEXAMDATE
                SetText vasID, strMaxSeq, intRow, colSAVESEQ
                SetText vasID, strChartID, intRow, colCHARTNO
                SetText vasID, Trim(RSH.Fields("DEPARTMENT")) & "", intRow, colPOSNO
                
                '-- �˻���ID
                mResult.OperatorID = Trim(RSH.Fields("OperatorID")) & ""
                gIFUser = mResult.OperatorID
                
                If strChartID <> "" Then
                    '-- �˻��� ���� �������̺��� ������ ǥ��
                    Call GetSampleInfoW_NCC(intRow)
                End If
                
                vasRes.MaxRows = 0
                
                '-- �˻���
                SQL = ""
                SQL = SQL & "SELECT Result_Type, Analyte, Value, Unit, InRange, ReturnCode "
                SQL = SQL & "  From Results "
                SQL = SQL & " Where TestID = '" & pPtID & "'"
                
                Call SetSQLData("EDM_RSLT", SQL)
                
                Set RSD = cn_Ser_Census.Execute(SQL, , 1)
            
               If Not RSD.EOF = True And Not RSD.BOF = True Then
                    Do Until RSD.EOF
                        strIntBase = Trim(RSD.Fields("Analyte")) & ""
                        strResult = Trim(RSD.Fields("Value")) & ""
                        strComm = Trim(RSD.Fields("InRange")) & ""
                        
                        If strResult <> "" Then
                            SQL = ""
                            SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
                            SQL = SQL & "  FROM EQPMASTER"
                            SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
                            SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
                            'SQL = SQL & "   AND EXAMCODE in (" & gOrderExam & ") "
                            
                            Res = GetDBSelectColumn(gLocal, SQL)
                            '-- ���� ���� ���
                            If Res > 0 Then
                                lsExamCode = Trim(gReadBuf(0))
                                lsExamName = Trim(gReadBuf(1))
                                lsSeqNo = Trim(gReadBuf(2))
                                
                                lsResRow = vasRes.DataRowCnt + 1
                                If vasRes.MaxRows < lsResRow Then
                                    vasRes.MaxRows = lsResRow
                                End If
                                
                                '�Ҽ��� ó��, ��� ���� ó��
                                lsEquipRes = strResult
                                strResult = SetResult(strResult, strIntBase)
                                lsResult_Buff = strResult
                                
                                
                                '-- �������� seq
                                For intCol = colState + 1 To vasID.MaxCols
                                    If lsExamCode = gArrEquip(intCol - colState, 3) Then
                                        SetText vasID, strResult, vasID.MaxRows, intCol
                                        'SetText vasRes, gArrEquip(intCol - colState, 7), lsResRow, colSUBCODE               'subcode
                                        Exit For
                                    End If
                                Next
                                
                                '-- Work List
                                SetText vasID, "Result", vasID.MaxRows, colState                 '11 �������
                                
            
                                '-- ��� List
                                SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      '����ڵ�
                                SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '�˻��ڵ�
                                SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '�˻��
                                SetText vasRes, lsEquipRes, lsResRow, colMachResult     '�����
                                SetText vasRes, strResult, lsResRow, colRESULT          '���
                                SetText vasRes, lsSeqNo, lsResRow, colSeq               '����
                                SetText vasRes, strComm, lsResRow, colFLAG              'Flag
                                '-- ���� ����
                                SetLocalDB vasID.MaxRows, lsResRow, "1", lsEquipRes
                                
                                SetText vasID, "1", vasID.MaxRows, colCHECKBOX
                                lsResult_Buff = ""
                                
                                strState = "R"
                            End If
                        End If
                        RSD.MoveNext
                        
                    Loop
                    
                    '-- ü������
                    If strTemperature <> "" Then
                        strIntBase = "Pt.temp"
                        strResult = Trim(strTemperature)
                        strComm = ""
                    
                        SQL = ""
                        SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
                        SQL = SQL & "  FROM EQPMASTER"
                        SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
                        SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
                        'SQL = SQL & "   AND EXAMCODE in (" & gOrderExam & ") "
                        
                        Res = GetDBSelectColumn(gLocal, SQL)
                        '-- ���� ���� ���
                        If Res > 0 Then
                            lsExamCode = Trim(gReadBuf(0))
                            lsExamName = Trim(gReadBuf(1))
                            lsSeqNo = Trim(gReadBuf(2))
                            
                            lsResRow = vasRes.DataRowCnt + 1
                            If vasRes.MaxRows < lsResRow Then
                                vasRes.MaxRows = lsResRow
                            End If
                            
                            '�Ҽ��� ó��, ��� ���� ó��
                            lsEquipRes = strResult
                            strResult = SetResult(strResult, strIntBase)
                            lsResult_Buff = strResult
                            
                            
                            '-- �������� seq
                            For intCol = colState + 1 To vasID.MaxCols
                                If lsExamCode = gArrEquip(intCol - colState, 3) Then
                                    SetText vasID, strResult, vasID.MaxRows, intCol
                                    'SetText vasRes, gArrEquip(intCol - colState, 7), lsResRow, colSUBCODE               'subcode
                                    Exit For
                                End If
                            Next
                            
                            '-- Work List
                            SetText vasID, "Result", vasID.MaxRows, colState                 '11 �������
                            
        
                            '-- ��� List
                            SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      '����ڵ�
                            SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '�˻��ڵ�
                            SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '�˻��
                            SetText vasRes, lsEquipRes, lsResRow, colMachResult     '�����
                            SetText vasRes, strResult, lsResRow, colRESULT          '���
                            SetText vasRes, lsSeqNo, lsResRow, colSeq               '����
                            SetText vasRes, strComm, lsResRow, colFLAG              'Flag
                            '-- ���� ����
                            SetLocalDB vasID.MaxRows, lsResRow, "1", lsEquipRes
                            
                            SetText vasID, "1", vasID.MaxRows, colCHECKBOX
                            lsResult_Buff = ""
                            
                            strState = "R"
                        End If
                    End If
                    
                    '-- Fio2����
                    If strFiO2 <> "" Then
                        strIntBase = "FiO2"
                        strResult = Trim(strFiO2)
                        strComm = ""
                    
                        SQL = ""
                        SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
                        SQL = SQL & "  FROM EQPMASTER"
                        SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
                        SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
                        'SQL = SQL & "   AND EXAMCODE in (" & gOrderExam & ") "
                        
                        Res = GetDBSelectColumn(gLocal, SQL)
                        '-- ���� ���� ���
                        If Res > 0 Then
                            lsExamCode = Trim(gReadBuf(0))
                            lsExamName = Trim(gReadBuf(1))
                            lsSeqNo = Trim(gReadBuf(2))
                            
                            lsResRow = vasRes.DataRowCnt + 1
                            If vasRes.MaxRows < lsResRow Then
                                vasRes.MaxRows = lsResRow
                            End If
                            
                            '�Ҽ��� ó��, ��� ���� ó��
                            lsEquipRes = strResult
                            strResult = SetResult(strResult, strIntBase)
                            lsResult_Buff = strResult
                            
                            
                            '-- �������� seq
                            For intCol = colState + 1 To vasID.MaxCols
                                If lsExamCode = gArrEquip(intCol - colState, 3) Then
                                    SetText vasID, strResult, vasID.MaxRows, intCol
                                    'SetText vasRes, gArrEquip(intCol - colState, 7), lsResRow, colSUBCODE               'subcode
                                    Exit For
                                End If
                            Next
                            
                            '-- Work List
                            SetText vasID, "Result", vasID.MaxRows, colState                 '11 �������
                            
        
                            '-- ��� List
                            SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      '����ڵ�
                            SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '�˻��ڵ�
                            SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '�˻��
                            SetText vasRes, lsEquipRes, lsResRow, colMachResult     '�����
                            SetText vasRes, strResult, lsResRow, colRESULT          '���
                            SetText vasRes, lsSeqNo, lsResRow, colSeq               '����
                            SetText vasRes, strComm, lsResRow, colFLAG              'Flag
                            '-- ���� ����
                            SetLocalDB vasID.MaxRows, lsResRow, "1", lsEquipRes
                            
                            SetText vasID, "1", vasID.MaxRows, colCHECKBOX
                            lsResult_Buff = ""
                            
                            strState = "R"
                        End If
                    End If
                    
                    ' EDM - LIS ���� ������Ʈ
                    SQL = ""
                    SQL = SQL & "UPDATE TestSendLog SET LISStatus = '8' "
                    SQL = SQL & " WHERE EPOCTestID = '" & pPtID & "'"
                    
                    'Call SetSQLData("EDM_UPDATE", SQL)
                    
                    Res = SendQuery(gServer_Census, SQL)
                    If Res = -1 Then
                        SaveQuery SQL
                        Exit Sub
                    End If
                    
                    '-- ��������
                    If MnTransAuto.Checked = True And strState = "R" Then
                        Call cmdIFTrans_Click
                    End If
                End If
                RSH.MoveNext
            Loop
        End If
    End If
    
    txtTimer.Enabled = True
    dtpToday.Value = Now
    
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
            'Buffer = ENQ & "DU   5003 51              0.145   106(+)    2015/01/061:08:06 00  1 10B99"
            
            SetRawData "[Rx]" & Buffer
            StatusBar1.Panels(3).Text = Buffer
            
            lngBufLen = Len(Buffer)
            
            For i = 1 To lngBufLen
                BufChar = Mid$(Buffer, i, 1)
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        comEqp.Output = ACK
                        SetRawData "[Tx]" & ACK
                    Case STX
                        intBufCnt = 1
                        Erase strRecvData
                        ReDim Preserve strRecvData(intBufCnt)
                    Case ETX
                        comEqp.Output = ACK
                        SetRawData "[Tx]" & ACK
                        strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                        Call EditRcvDataASTM
                        'Erase strRecvData
                    Case EOT
                        Erase strRecvData
                    Case Else
                        strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                End Select
            Next i


'            For i = 1 To lngBufLen
'                BufChar = Mid$(Buffer, i, 1)
'
'                Select Case intPhase
'                    Case 1      '## Estabilshment Phase
'                        Select Case BufChar
'                            Case ENQ
'                                intBufCnt = 1
'                                Erase strRecvData
'                                ReDim Preserve strRecvData(intBufCnt)
'                                intPhase = 2
'                                comEqp.Output = ACK
'                                SetRawData "[Tx]" & ACK
'                            Case ACK
'                                '-- ��񿡼� �Ѿ�� �ð��� �쿬�� 11:59:59�ʳ� ���Ͽ� ����� �ð��� ���
'                                '-- ��� ����� �������� ������ �� �����Ƿ� ��¥�� �ǽð� ������Ʈ �Ѵ�.
'                                strDate = DateCompare(Format(CDate(dtpToday.Value), "yyyymmdd"))
'                                dtpToday.Value = Format(strDate, "####-##-##")
'
'                                DoEvents
'
'                                If strState = "Q" Then Call SendOrder
'
'                        End Select
'                    Case 2      '## Transfer Phase
'                        Select Case BufChar
'                            Case ENQ
'                                Erase strRecvData
'                                comEqp.Output = ACK
'                                SetRawData "[Tx]" & ACK
'                            Case STX
'                                intBufCnt = 1
'                                Erase strRecvData
'                                ReDim Preserve strRecvData(intBufCnt)
'                            Case ETB
'                                blnIsETB = True
'                                intPhase = 3
'                            Case ETX
'                                intBufCnt = intBufCnt + 1
'                                ReDim Preserve strRecvData(intBufCnt)
'                                intPhase = 3
'                            Case vbCr, vbLf
'                            Case EOT
'                                intPhase = 1
'                            Case Else
'                                If blnIsETB = False Then
'                                    strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
'                                Else
'                                    blnIsETB = False
'                                End If
'                        End Select
'                    Case 3      '## Transfer Phase
'                        Select Case BufChar
'                            Case vbCr
'                            Case vbLf
'                                intPhase = 4
'                                comEqp.Output = ACK
'                                SetRawData "[Tx]" & ACK
'                        End Select
'                    Case 4      '## Termination Phase
'                        Select Case BufChar
'                            Case STX
'                                intPhase = 2
'                            Case EOT
'                                '-- ��񿡼� �Ѿ�� �ð��� �쿬�� 11:59:59�ʳ� ���Ͽ� ����� �ð��� ���
'                                '-- ��� ����� �������� ������ �� �����Ƿ� ��¥�� �ǽð� ������Ʈ �Ѵ�.
'                                strDate = DateCompare(Format(CDate(dtpToday.Value), "yyyymmdd"))
'                                dtpToday.Value = Format(strDate, "####-##-##")
'
'                                DoEvents
'
'                                Call EditRcvDataASTM
'
'                                If strState = "Q" Then
'                                    intSndPhase = 1
'                                    intFrameNo = 1
'                                    comEqp.Output = ENQ
'                                    SetRawData "[Tx]" & ENQ
'                                End If
'
'                                intPhase = 1
'                        End Select
'                End Select
'            Next i
            
        Case comEvSend
            imgSend.Picture = imlStatus.ListImages("RUN").ExtractIcon
            If tmrSend.Enabled = False Then
                tmrSend.Enabled = True
            Else
                tmrSend.Enabled = False
                tmrSend.Enabled = True
            End If
        
        Case comEvCTS
            EVMsg$ = "CTS ���� ����"
        Case comEvDSR
            EVMsg$ = "DSR ���� ����"
        Case comEvCD
            EVMsg$ = "CD ���� ����"
        Case comEvRing
            EVMsg$ = "��ȭ ���� �︮�� ��"
        Case comEvEOF
            EVMsg$ = "EOF ����"

        '���� �޽���
        Case comBreak
            ERMsg$ = "�ߴ� ��ȣ ����"
        Case comCDTO
            ERMsg$ = "�ݼ��� ���� �ð� �ʰ�"
        Case comCTSTO
            ERMsg$ = "CTS �ð� �ʰ�"
        Case comDCB
            ERMsg$ = "DCB �˻� ����"
        Case comDSRTO
            ERMsg$ = "DSR �ð� �ʰ�"
        Case comFrame
            ERMsg$ = "�����̹� ����"
        Case comOverrun
            ERMsg$ = "�и�Ƽ ����"
        Case comRxOver
            ERMsg$ = "���� ���� �ʰ�"
        Case comRxParity
            ERMsg$ = "�и�Ƽ ����"
        Case comTxFull
            ERMsg$ = "���� ���ۿ� ������ ����"
        Case Else
            ERMsg$ = "�� �� ���� ���� �Ǵ� �̺�Ʈ"
    End Select


End Sub

'-----------------------------------------------------------------------------'
'   ��� : �ش� ���ڵ��ȣ�� ���� �������� ��ȸ, ǥ��, �˻���������
'   �μ� :
'       - pBarNo : ���ڵ��ȣ
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
    
    '-- ���������� ǥ��
    Call SetText(vasID, pBarNo, intRow, colBARCODE)             '-- ���ڵ�
    Call SetText(vasID, mOrder.RackNo, intRow, colDISKNO)       '-- Rack
    Call SetText(vasID, mOrder.TubePos, intRow, colPOSNO)       '-- Pos
    
    '-- ȯ������ ǥ��
    Call vasActiveCell(vasID, intRow, colBARCODE)
    
    '-- ����������� �����
    Call ClearSpread(vasRes)
    
    '-- �˻��� ���� ��������
    Call GetSampleInfoW(intRow)
    
    '-- ���ڵ��ȣ�� �ش��ϴ� �˻��ڵ� ��������
    gOrderExam = GetOrderExamCode(gEquip, pBarNo)

    '-- �������̺��� �˻��׸� �ش��ϴ� �˻�ä�� ã�ƿ��� (intRow = ���� �˻��ߴ� ���ڵ尡 �ٽ� �ö�� ��� ��ġ�� ��ã�´�.)
    strItems = GetGetEquipExamCode_XN1000(gEquip, pBarNo, intRow)

    '-- �˻�ä�η� ������ �����
    If Trim(strItems) = "" Then
        mOrder.NoOrder = True
        mOrder.Order = strItems
    Else
        mOrder.NoOrder = False
        mOrder.Order = strItems
    End If
    
    '-- �������(Order) ǥ��
    Call SetText(vasID, "Order", intRow, colState)
    
    '-- ���� Row
    gRow = intRow

End Sub

'-- �������̺��� �˻��׸� �ش��ϴ� �˻�ä�� ã�ƿ���
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
    
    sBarcode = Trim(GetText(frmInterface.vasID, intRow, colBARCODE))   '2 ���� ���ڵ� ��ȣ
    SetRawData "[sBarcode]" & sBarcode
    
    If sBarcode = "" Then
        Exit Function
    End If
    
    ClearSpread frmInterface.vasTemp1
    
    '-- ������ �˻��ڵ��� ä�� ã��
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
            'NRBC%�� ������ ���ش�
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
               
                '-- ^^^^LYMPH#\�� �ΰ��� ������ ETB �� ��񿡼� �ν����� ���ϱ� ����..(�� �ڸ��� 230)
                strDiff = "^^^^NEUT#\^^^^LYMPH%\^^^^MONO#\^^^^EO#\^^^^BASO#\^^^^NEUT%\^^^^LYMPH#\^^^^LYMPH#\^^^^MONO%\^^^^EO%\^^^^BASO%\^^^^IG#\^^^^IG%\"
                
            End If
        Else
            Exit For
        End If
    Next

    strExamCode = strCBC & strDiff
    
    '-- ������ ���� ��� CBC�� �˻��ϵ��� �Ѵ�.
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
'   ��� :
'   �μ� :
'       - pBarNo : ���ڵ��ȣ
'-----------------------------------------------------------------------------'
Private Sub SetPatInfo(ByVal pBarNo As String, Optional ByVal pRno As String, Optional ByVal pPno As String)
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
    
    '-- ���������� ǥ��
    Call SetText(vasID, "1", intRow, colCHECKBOX)
    If pBarNo = "" Then
        Call SetText(vasID, mResult.SpcPos, intRow, colBARCODE)
    Else
        Call SetText(vasID, mResult.BarNo, intRow, colBARCODE)
    End If
    'Call SetText(vasID, mResult.RackNo, intRow, colDISKNO)
    'Call SetText(vasID, mResult.TubePos, intRow, colPOSNO)
    Call SetText(vasID, mResult.RsltDate, intRow, colEXAMDATE)
    Call SetText(vasID, mResult.RsltSeq, intRow, colSAVESEQ)
    
    
    Call vasActiveCell(vasID, intRow, colBARCODE)
    
    '-- ����������� �����
    Call ClearSpread(vasRes)
    
    If Len(pBarNo) >= 10 Then
        '-- �˻��� ���� �������̺��� ������ ǥ��(for ��ũ����Ʈ)  '6,7,8,9
        Call GetSampleInfoW_SUNGMO(intRow)
    End If
    
    '-- ���ڵ��ȣ�� �����ϴ� �˻��ڵ� ��������(�μ� : ����ڵ�,íƮ��ȣ,������,������ȣ,������ȣ)
    'gOrderExam = GetOrderExamCode_UBCARE(gEquip, pBarNo)
    
    '-- ���� Row
    gRow = intRow
    
End Sub

Public Sub ComReceive(ByRef RecData As String)

    strBuffer = RecData
    
    Call EditRcvDataASTM
    
End Sub

'-----------------------------------------------------------------------------'
'   ��� : ���κ� ������ ������ ����
'-----------------------------------------------------------------------------'
Private Sub EditRcvDataASTM()
    Dim strRcvBuf    As String   '������ Data
    Dim strType      As String   '������ Record Type
    Dim strBarNo     As String   '������ ���ڵ��ȣ
    Dim strSeq       As String   '������ Sequence
    Dim strRackNo    As String   '������ Rack Or Disk No
    Dim strTubePos   As String   '������ Tube Position
    Dim strIntBase   As String   '������ ������ �˻��
    Dim strResult    As String   '������ ���(����)
    Dim strIntResult As String   '������ ���(����)
    Dim strQCResult  As String   '������ ���(QC)
    Dim strFlag      As String   '������ Abnormal Flag
    Dim strComm      As String   '������ Comment
    Dim strTemp1     As String
    Dim strTemp2     As String
    Dim intCnt       As Integer
    
    Dim lsExamCode As String
    Dim lsExamName As String
    Dim lsSeqNo As String
    Dim lsSubCode As String
    
    Dim lsResult_Buff As String
    Dim lsExamDate As String
    Dim lsEquipRes As String
    Dim lsResRow    As String
    Dim ii As Integer
    Dim strTmp      As String
    Dim intIDX      As Integer
    Dim i As Integer
    Dim varORQN As Variant
    
    Dim fLCBS300E As Variant
    Dim strSndBuffer As String
    Dim strSampleID As String
    
   ' On Error Resume Next
    
    fLCBS300E = Split(strBuffer, vbCr)
 
    For intCnt = 0 To UBound(fLCBS300E)
        strTemp1 = fLCBS300E(intCnt)
        Debug.Print fLCBS300E(intCnt)
        
        Select Case Mid(strTemp1, 1, 3)
            Case "MSH"
            Case "PID"
                    strSampleID = mGetP(strTemp1, 2, "|")
            Case "OBR"
                '-- �������̽� ����
                               strSndBuffer = SB & "MSH|^~\&|Mindray|BS-220|||" & Format(Now, "yyyymmddhhmmss") & "||ACK^R01|" & strSampleID & "|P|2.3.1||||0||ASCII|||" & vbCr
                strSndBuffer = strSndBuffer & "MSA|AA|1|Message accepted|||0|" & vbCr
                strSndBuffer = strSndBuffer & EB & vbCr
                               
'                sck.ProcSendMessage strSndBuffer
                
                
                
                'strBarno = Val(mGetP(strTemp1, 3, "|"))
                strSeq = Val(mGetP(strTemp1, 4, "|"))
                
                For i = 1 To vasID.DataRowCnt
                    vasID.Row = i
                    vasID.Col = colDISKNO
                    If vasID.Text = strSeq Then
                        vasID.Col = colBARCODE
                        strBarNo = Trim(vasID.Text)
                        Exit For
                    End If
                Next
            
                With mResult
                    .BarNo = strBarNo
                    .RackNo = strRackNo
                    .TubePos = strTubePos
                    .RsltDate = Format(Now, "yyyymmddhhmmss")
                    .RsltSeq = getMaxTestNum(Format(dtpToday, "yyyymmdd"))
                    .SpcPos = strSeq
                End With
                                    
                Call SetPatInfo(strBarNo)
                
                If gRow <= 0 Then
                    Exit Sub
                End If
                
                strState = "O"
                
                '-- ������ ���ȭ�� �ʱ�ȭ
                vasRes.MaxRows = 0
            
            Case "OBX"   '## Result
                '## ������ �˻��, ���, Abnormal Flag
                'OBX|1|NM|CRP|CRP|0.010000|mg/dL|-|Normal|||F||0.007145|20140508172657||||
                
                strIntBase = mGetP(strTemp1, 4, "|")
                'strIntBase = mGetP(strTemp1, 5, "|")
                strIntBase = Trim(strIntBase)
                
                strResult = mGetP(strTemp1, 6, "|")
                strResult = Trim(strResult)
                
                If strIntBase = "RF" Then
                    If Val(strResult) = 0 Then
                        strResult = "<0.1"
                    End If
                End If
                
                If strIntBase = "CRP" Then
                    If Val(strResult) = 0 Then
                        strResult = "<0.01"
                    End If
                End If
                
                If strResult <> "" Then
                    SQL = ""
                    SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
                    SQL = SQL & "  FROM EQPMASTER"
                    SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
                    SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
                    SQL = SQL & "   AND EXAMCODE  in (" & gOrderExam & ") "
                    
                    Res = GetDBSelectColumn(gLocal, SQL)
                    
                    '-- ���� ���� ���
                    If Res > 0 Then
                        lsExamCode = Trim(gReadBuf(0))
                        lsExamName = Trim(gReadBuf(1))
                        lsSeqNo = Trim(gReadBuf(2))
                        'lsSubCode = Trim(gReadBuf(3))
                        
                        lsResRow = vasRes.DataRowCnt + 1
                        If vasRes.MaxRows < lsResRow Then
                            vasRes.MaxRows = lsResRow
                        End If
                        
                        '�Ҽ��� ó��, ��� ���� ó��
                        lsEquipRes = strResult
                        strResult = SetResult(strResult, strIntBase)
                        
                        If strIntBase = "RF" Then
                            If Val(strResult) = 0 Then
                                strResult = "<0.1"
                            End If
                        End If
                        
                        If strIntBase = "CRP" Then
                            If Val(strResult) = 0 Then
                                strResult = "<0.01"
                            End If
                        End If
                        
                        lsResult_Buff = strResult
                        
                        '-- Work List
                        SetText vasID, "Result", gRow, colState                 '11 �������
                        

                        '-- ��� List
                        SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      '����ڵ�
                        SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '�˻��ڵ�
                        SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '�˻��
                        SetText vasRes, lsEquipRes, lsResRow, colMachResult     '�����
                        SetText vasRes, strResult, lsResRow, colRESULT          '���
                        SetText vasRes, lsSeqNo, lsResRow, colSeq               '����
                        SetText vasRes, lsSubCode, lsResRow, colSUBCODE         'SUBCODE
                        SetText vasRes, strComm, lsResRow, colFLAG              'Flag
                        '-- ���� ����
                        SetLocalDB gRow, lsResRow, "1", lsEquipRes
                    
                        lsResult_Buff = ""
                        
                        strState = "R"
                        
                    '-- ���� ���� ���
                    Else
                    
                              SQL = "Select EXAMCODE,EXAMNAME,SEQNO "
                        SQL = SQL & "  From EQPMASTER"
                        SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
                        SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
                        Res = GetDBSelectColumn(gLocal, SQL)
                        
                        If Res > 0 Then
                            lsExamCode = Trim(gReadBuf(0))
                            lsExamName = Trim(gReadBuf(1))
                            lsSeqNo = Trim(gReadBuf(2))
                            'lsSubCode = Trim(gReadBuf(3))
                            
                            lsResRow = vasRes.DataRowCnt + 1
                            If vasRes.MaxRows < lsResRow Then
                                vasRes.MaxRows = lsResRow
                            End If
                            
                            '�Ҽ��� ó��, ��� ���� ó��
                            lsEquipRes = strResult
                            strResult = SetResult(strResult, strIntBase)
                            
                            If strIntBase = "RF" Then
                                If Val(strResult) = 0 Then
                                    strResult = "<0.1"
                                End If
                            End If
                            
                            If strIntBase = "CRP" Then
                                If Val(strResult) = 0 Then
                                    strResult = "<0.01"
                                End If
                            End If
                            
                            lsResult_Buff = strResult
                            
                            '-- Work List
                            SetText vasID, "Result", gRow, colState                 '�������
                            
                            '-- ��� List
                            SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      '����ڵ�
                            SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '�˻��ڵ�
                            SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '�˻��
                            SetText vasRes, lsEquipRes, lsResRow, colMachResult     '�����
                            SetText vasRes, strResult, lsResRow, colRESULT          '���
                            SetText vasRes, lsSeqNo, lsResRow, colSeq               '����
                            SetText vasRes, lsSubCode, lsResRow, colSUBCODE         'SUBCODE
                            SetText vasRes, strComm, lsResRow, colFLAG              'Flag
                            '-- ���� ����
                            SetLocalDB gRow, lsResRow, "1", lsEquipRes
                                        
                            lsResult_Buff = ""
                            
                        End If
                    End If
                End If
        End Select
    Next

    '## DB�� �������
    If MnTransAuto.Checked = True And strState = "R" Then
        
        Res = SaveTransDataW(gRow)
        
        If Res = -1 Then
            '-- ���� ����
            SetForeColor vasID, gRow, gRow, 1, colState, 255, 0, 0
            SetText vasID, "Failed", gRow, colState
        Else
            '-- ���� ����
            SetBackColor vasID, gRow, gRow, 1, colState, 202, 255, 112
            SetText vasID, "Trans", gRow, colState
            
            SQL = " Update PATRESULT Set " & vbCrLf & _
                  " sendflag = '2' " & vbCrLf & _
                  " Where equipno = '" & gEquip & "' " & vbCrLf & _
                  " And barcode = '" & Trim(GetText(vasID, gRow, colBARCODE)) & "' "
            Res = SendQuery(gLocal, SQL)
            If Res = -1 Then
                SaveQuery SQL
                Exit Sub
            End If
        End If
    End If

    'SetText vasID, "Result", gRow, colState
    strState = ""

End Sub

''''-----------------------------------------------------------------------------'
''''   ��� : ���κ� ������ ������ ����
''''-----------------------------------------------------------------------------'
'''Private Sub EditRcvDataASTM_NEW()
'''    Dim strRcvBuf    As String   '������ Data
'''    Dim strType      As String   '������ Record Type
'''    Dim strBarNo     As String   '������ ���ڵ��ȣ
'''    Dim strSeq       As String   '������ Sequence
'''    Dim strRackNo    As String   '������ Rack Or Disk No
'''    Dim strTubePos   As String   '������ Tube Position
'''    Dim strIntBase   As String   '������ ������ �˻��
'''    Dim strResult    As String   '������ ���(����)
'''    Dim strIntResult As String   '������ ���(����)
'''    Dim strQCResult  As String   '������ ���(QC)
'''    Dim strFlag      As String   '������ Abnormal Flag
'''    Dim strComm      As String   '������ Comment
'''    Dim strTemp1     As String
'''    Dim strTemp2     As String
'''    Dim intCnt       As Integer
'''
'''    Dim lsExamCode As String
'''    Dim lsExamName As String
'''    Dim lsSeqNo As String
'''    Dim lsResult_Buff As String
'''    Dim lsExamDate As String
'''    Dim lsEquipRes As String
'''    Dim lsResRow    As String
'''    Dim ii As Integer
'''    Dim strTmp      As String
'''    Dim intIdx      As Integer
'''    Dim varRcvBuf   As Variant
'''    Dim intRow      As Integer
'''    Dim i As Integer
'''    Dim intCol As Integer
'''    Dim varHoriba As Variant
'''    Dim Pos As Integer
'''    Dim strSeqNo As String
'''
'''
'''    'For intCnt = 1 To UBound(strRecvData)
'''        strRcvBuf = strRecvData(1)
'''
'''        'Pos = InStr(strRcvBuf, "DU")
'''
'''        Pos = InStr(strRcvBuf, "DU")
'''        If Pos > 0 Then
'''            strSeqNo = Trim(Mid(strRcvBuf, 3, 4))
'''            strRackNo = Trim(Mid(strRcvBuf, 8, 2))
'''            strTubePos = Trim(Mid(strRcvBuf, 10, 2))
'''            strBarNo = Trim(Mid(strRcvBuf, 13, 13))
'''
'''           ' If strBarNo = "" Then Exit Sub
'''
'''            With mResult
'''                .BarNo = strBarNo
'''                .RackNo = strRackNo
'''                .TubePos = strTubePos
'''                .RsltDate = Format(Now, "yyyymmddhhmmss")
'''                .RsltSeq = getMaxTestNum(Format(dtpToday, "yyyymmdd"))
'''                .SpcPos = strSeqNo
'''            End With
'''
'''            Call SetPatInfo(strBarNo)
'''
'''            If gRow <= 0 Then
'''                Exit Sub
'''            End If
'''
'''            strState = "O"
'''
'''            '-- ������ ���ȭ�� �ʱ�ȭ
'''            vasRes.MaxRows = 0
'''
'''            '## ������ �˻��, ���, Abnormal Flag
'''            strIntBase = "1"
'''            strIntResult = Trim$(Mid$(strRcvBuf, 32, 6))   '����
'''            strResult = Trim$(Mid$(strRcvBuf, 43, 1))   '����
''''            If strResult = "-" Then
''''                strResult = "Negative"
''''            Else
''''                strResult = "Positive"
''''            End If
'''
'''            If strIntResult <> "" Then
'''                strResult = strIntResult & "(" & strResult & ")"
'''            End If
'''
'''            If strResult <> "" And Len(strIntBase) <= 6 Then
'''                SQL = ""
'''                SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
'''                SQL = SQL & "  FROM EQPMASTER"
'''                SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
'''                SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
'''                SQL = SQL & "   AND EXAMCODE in (" & gOrderExam & ") "
'''
'''                Res = GetDBSelectColumn(gLocal, SQL)
'''
'''                '-- ���� ���� ���
'''                If Res > 0 Then
'''                    lsExamCode = Trim(gReadBuf(0))
'''                    lsExamName = Trim(gReadBuf(1))
'''                    lsSeqNo = Trim(gReadBuf(2))
'''
'''                    lsResRow = vasRes.DataRowCnt + 1
'''                    If vasRes.MaxRows < lsResRow Then
'''                        vasRes.MaxRows = lsResRow
'''                    End If
'''
'''                    '�Ҽ��� ó��, ��� ���� ó��
'''                    lsEquipRes = strResult
'''                    strResult = SetResult(strResult, strIntBase)
'''                    lsResult_Buff = strResult
'''
'''                    '-- Work List
'''                    SetText vasID, "Result", gRow, colState                 '11 �������
'''
'''                    '-- �������� seq
'''                    For intCol = colState + 1 To vasID.MaxCols
'''                        If lsExamCode = gArrEquip(intCol - colState, 3) Then
'''                            SetText vasID, strResult, gRow, intCol
'''                            SetText vasRes, gArrEquip(intCol - colState, 7), lsResRow, colSUBCODE               'subcode
'''                            Exit For
'''                        End If
'''                    Next
'''
'''
'''                    '-- ��� List
'''                    SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      '����ڵ�
'''                    SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '�˻��ڵ�
'''                    SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '�˻��
'''                    SetText vasRes, lsEquipRes, lsResRow, colMachResult     '�����
'''                    SetText vasRes, strResult, lsResRow, colRESULT          '���
'''                    SetText vasRes, lsSeqNo, lsResRow, colSeq               '����
'''                    SetText vasRes, strComm, lsResRow, 7                    'Flag
'''                    '-- ���� ����
'''                    SetLocalDB gRow, lsResRow, "1", lsEquipRes
'''
'''                    lsResult_Buff = ""
'''
'''                    strState = "R"
'''
'''                '-- ���� ���� ���
'''                Else
'''
'''                          SQL = "Select examcode, examname, seqno "
'''                    SQL = SQL & "  From EQPMASTER"
'''                    SQL = SQL & " Where equipno = '" & gEquip & "' "
'''                    SQL = SQL & "   and equipcode = '" & strIntBase & "' "
'''                    Res = GetDBSelectColumn(gLocal, SQL)
'''
'''                    If Res > 0 Then
'''                        lsExamCode = Trim(gReadBuf(0))
'''                        lsExamName = Trim(gReadBuf(1))
'''                        lsSeqNo = Trim(gReadBuf(2))
'''
'''                        lsResRow = vasRes.DataRowCnt + 1
'''                        If vasRes.MaxRows < lsResRow Then
'''                            vasRes.MaxRows = lsResRow
'''                        End If
'''
'''                        '�Ҽ��� ó��, ��� ���� ó��
'''                        lsEquipRes = strResult
'''                        strResult = SetResult(strResult, strIntBase)
'''                        lsResult_Buff = strResult
'''
'''                        '-- Work List
'''                        SetText vasID, "Result", gRow, colState                 '�������
'''
'''                        '-- �������� seq
'''                        For intCol = colState + 1 To vasID.MaxCols
'''                            If lsExamCode = gArrEquip(intCol - colState, 3) Then
'''                                SetText vasID, strResult, gRow, intCol
'''                                SetText vasRes, gArrEquip(intCol - colState, 7), lsResRow, colSUBCODE               'subcode
'''                                Exit For
'''                            End If
'''                        Next
'''
'''                        '-- ��� List
'''                        SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      '����ڵ�
'''                        SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '�˻��ڵ�
'''                        SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '�˻��
'''                        SetText vasRes, lsEquipRes, lsResRow, colMachResult     '�����
'''                        SetText vasRes, strResult, lsResRow, colRESULT          '���
'''                        SetText vasRes, lsSeqNo, lsResRow, colSeq               '����
'''                        SetText vasRes, strComm, lsResRow, colFLAG              'Flag
'''                        '-- ���� ����
'''                        SetLocalDB gRow, lsResRow, "1", lsEquipRes
'''
'''                        lsResult_Buff = ""
'''                        strState = "R"
'''                    End If
'''                End If
'''            End If
'''            vasRes.RowHeight(-1) = 14
'''
'''            '## DB�� �������
'''            If MnTransAuto.Checked = True And strState = "R" Then
'''                Res = SaveTransDataW(gRow)
'''
'''                If Res = -1 Then
'''                    '-- ���� ����
'''                    SetForeColor vasID, gRow, gRow, 1, colState, 255, 0, 0
'''                    SetText vasID, "Failed", gRow, colState
'''                Else
'''                    '-- ���� ����
'''                    SetBackColor vasID, gRow, gRow, 1, colState, 202, 255, 112
'''                    SetText vasID, "Trans", gRow, colState
'''                    SetText vasID, "0", gRow, colCheckBox
'''
'''                          SQL = "Update PATRESULT Set " & vbCrLf
'''                    SQL = SQL & " sendflag = '2' " & vbCrLf
'''                    SQL = SQL & " Where equipno = '" & gEquip & "' " & vbCrLf
'''                    SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(vasID, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
'''                    SQL = SQL & "   And barcode = '" & Trim(GetText(vasID, gRow, colBARCODE)) & "' " & vbCrLf
'''                    SQL = SQL & "   And saveseq = " & Trim(GetText(vasID, gRow, colSAVESEQ)) & vbCrLf
'''
'''                    Res = SendQuery(gLocal, SQL)
'''                    If Res = -1 Then
'''                        SaveQuery SQL
'''                        Exit Sub
'''                    End If
'''                End If
'''                strState = ""
'''            End If
'''
'''        End If
'''    'Next
'''
'''End Sub


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
' asRow2 = ��� List
Function SetLocalDB(ByVal asRow1 As Long, ByVal asRow2 As Long, asSend As String, Optional asEquipResult As String = "")
    Dim sCnt As String
    Dim sExamDate As String
    Dim strSaveSeq As String

    sExamDate = Format(dtpToday, "yyyymmddhhmmss")

    If Trim(GetText(vasID, asRow1, colSAVESEQ)) = "" Then
        Exit Function
    End If
    
    SQL = ""
    SQL = "DELETE FROM PATRESULT " & vbCrLf & _
          " WHERE EXAMDATE = '" & Mid(sExamDate, 1, 8) & "' " & vbCrLf & _
          "   AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
          "   AND BARCODE = '" & Trim(GetText(vasID, asRow1, colBARCODE)) & "' " & vbCrLf & _
          "   AND EQUIPCODE = '" & Trim(GetText(vasRes, asRow2, colEQUIPCODE)) & "'" & vbCrLf & _
          "   AND EXAMCODE = '" & Trim(GetText(vasRes, asRow2, colEXAMCODE)) & "'"
          
    Res = SendQuery(gLocal, SQL)
    
    If Res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
    SQL = ""
    SQL = SQL & "INSERT INTO PATRESULT (" & vbCrLf
    SQL = SQL & "SAVESEQ"                           '�������(��¥��)
    SQL = SQL & ", EXAMDATE"                        '�˻�����"
    SQL = SQL & ", HOSPDATE"                        '������������"
    SQL = SQL & ", EQUIPNO"                         '����ڵ�"
    SQL = SQL & ", BARCODE" & vbCrLf
    SQL = SQL & ", EQUIPCODE"                       '�˻�ä��"
    SQL = SQL & ", EXAMCODE"                        '�����˻��ڵ�"
    SQL = SQL & ", EXAMSUBCODE"                     '�����˻��ڵ�(SUB)"
    SQL = SQL & ", EXAMNAME"
    SQL = SQL & ", SEQNO" & vbCrLf                  '�˻��Ϸù�ȣ"
    SQL = SQL & ", SAMPLETYPE"                      '��ü����"
    SQL = SQL & ", INOUT"                           '��ü�ڵ�
    SQL = SQL & ", DISKNO"
    SQL = SQL & ", POSNO"
    SQL = SQL & ", EQUIPRESULT"                     '�����"
    SQL = SQL & ", RESULT" & vbCrLf                 '�Ҽ���������"
    SQL = SQL & ", REFFLAG"
    SQL = SQL & ", REFVALUE"
    SQL = SQL & ", CHARTNO"
    SQL = SQL & ", PID"                             '���Ϲ�ȣ(������ȣ)"
    SQL = SQL & ", PNAME" & vbCrLf
    SQL = SQL & ", PSEX"
    SQL = SQL & ", PAGE"
    SQL = SQL & ", PJUMIN"
    SQL = SQL & ", PANICVALUE"
    SQL = SQL & ", DELTAVALUE" & vbCrLf
    SQL = SQL & ", SENDFLAG"                        '���۱���(0:������,1:����)"
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
    SQL = SQL & "','" & Trim(GetText(vasRes, asRow2, colSUBCODE))
    SQL = SQL & "','" & Trim(GetText(vasRes, asRow2, colEXAMNAME))
    SQL = SQL & "','" & Trim(GetText(vasRes, asRow2, colSeq))
    SQL = SQL & "','"
    SQL = SQL & "','" & Trim(GetText(vasID, asRow1, colINOUT))
    SQL = SQL & "','" & Trim(GetText(vasID, asRow1, colDISKNO))
    SQL = SQL & "','" & Trim(GetText(vasID, asRow1, colPOSNO))
    SQL = SQL & "','" & Trim(GetText(vasRes, asRow2, colMachResult))
    SQL = SQL & "','" & Trim(GetText(vasRes, asRow2, colRESULT))
    SQL = SQL & "','" & Trim(GetText(vasRes, asRow2, colFLAG))
    SQL = SQL & "',''"
    SQL = SQL & ",'" & Trim(GetText(vasID, asRow1, colCHARTNO))
    SQL = SQL & "','" & Trim(GetText(vasID, asRow1, colPID))
    SQL = SQL & "','" & Trim(GetText(vasID, asRow1, colPNAME))
    SQL = SQL & "','" & Trim(GetText(vasID, asRow1, colPSEX))
    SQL = SQL & "','" & Trim(GetText(vasID, asRow1, colPAGE))
    SQL = SQL & "','" & Trim(GetText(vasID, asRow1, colPOSNO))
    SQL = SQL & "',''"
    SQL = SQL & ",''"
    SQL = SQL & ",'1'"
    SQL = SQL & ",''"
'    SQL = SQL & ",'" & gIFUser
'    SQL = SQL & ",'" & mResult.OperatorID
    SQL = SQL & ",'" & gUserID
    
    SQL = SQL & "','')"
    
    Res = SendQuery(gLocal, SQL)
    
    If Res = -1 Then
        SaveQuery SQL
        Exit Function
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
    sMsg = "�˻��ڸ� �Է����ּ���."
    lblUser.Caption = InputBox(sMsg, "�˻��� �Է�")

End Sub

Private Sub txtBarNum_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Not IsNumeric(txtBarNum) Then
            StatusBar1.Panels(3).Text = "���ڵ��ȣ�� ���ڸ� �Է��� �����մϴ�."
            txtBarNum = ""
            Exit Sub
        End If
        
        If Len(txtBarNum) <> 12 Then
            StatusBar1.Panels(3).Text = "���ڵ� �ڸ����� Ȯ���ϼ���"
            txtBarNum = ""
            Exit Sub
        End If
        
        If Trim(txtBarNum) <> "" Then
            Call GetWorkList(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"), Trim(txtBarNum))
        End If
        vasID.RowHeight(-1) = 12
        txtBarNum.Text = ""
    End If
    
End Sub



Private Sub vasID_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    Dim i As Integer
    
    If BlockRow <= 0 Then
        Exit Sub
    End If
    
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
    Dim RS          As ADODB.Recordset
    
    If Row = 0 Then
        With vasID
            .Col = 1: .Col2 = .MaxCols
            .Row = 2: .Row2 = .DataRowCnt
            .SortBy = 0
            .SortKey(1) = Col       '����Ű ����ȣ

            .SortKeyOrder(1) = SortKeyOrderAscending
    
            .Action = ActionSort
        End With
        Exit Sub
    End If
    
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
    
'    If lblSaveSeq.Caption = "" Then
'        Exit Sub
'    End If
    
    'Local���� �ҷ�����
    ClearSpread vasRes
    
    '����ڵ�, �˻��ڵ�, �˻��, ���, ����
          SQL = "SELECT EQUIPCODE, EXAMCODE, EXAMNAME, EQUIPRESULT, RESULT, SEQNO, REFFLAG, EXAMSUBCODE " & vbCrLf
    SQL = SQL & "  FROM PATRESULT " & vbCrLf
    SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "'" & vbCrLf
    'If lblSaveSeq.Caption <> "" Then
    '    SQL = SQL & "   AND SAVESEQ = " & lblSaveSeq.Caption & vbCrLf
    'End If
    SQL = SQL & "   AND BARCODE = '" & lsID & "' " & vbCrLf
    'SQL = SQL & "   AND EXAMDATE = '" & Mid(Trim(GetText(vasID, Row, colOrdDate)), 1, 8) & "' " & vbCrLf
    SQL = SQL & " GROUP BY SEQNO, EQUIPCODE, EXAMCODE, EXAMNAME, EQUIPRESULT, RESULT, SEQNO, REFFLAG, EXAMSUBCODE "
    SQL = SQL & " ORDER BY SEQNO * 10"
    
    Set RS = cn.Execute(SQL, , 1)

    If Not RS.EOF = True And Not RS.BOF = True Then
        vasRes.MaxRows = 0
        Do Until RS.EOF
            With vasRes
                .MaxRows = .MaxRows + 1
                SetText vasRes, "0", .MaxRows, colCHECKBOX
                SetText vasRes, Trim(RS.Fields("EQUIPCODE")) & "", .MaxRows, colEQUIPCODE
                SetText vasRes, Trim(RS.Fields("EXAMCODE")) & "", .MaxRows, colEXAMCODE
                SetText vasRes, Trim(RS.Fields("EXAMNAME")) & "", .MaxRows, colEXAMNAME
                SetText vasRes, Trim(RS.Fields("EQUIPRESULT")) & "", .MaxRows, colMachResult
                SetText vasRes, Trim(RS.Fields("RESULT")) & "", .MaxRows, colRESULT
                SetText vasRes, Trim(RS.Fields("SEQNO")) & "", .MaxRows, colSeq
                SetText vasRes, Trim(RS.Fields("REFFLAG")) & "", .MaxRows, colFLAG
                SetText vasRes, Trim(RS.Fields("EXAMSUBCODE")) & "", .MaxRows, colSUBCODE
                
                If Trim(RS.Fields("REFFLAG")) = "H" Then
                    .Row = .MaxRows
                    .Col = colRESULT
                    .ForeColor = vbRed
                ElseIf Trim(RS.Fields("REFFLAG")) = "L" Then
                    .Row = .MaxRows
                    .Col = colRESULT
                    .ForeColor = vbBlue
                End If
           
            End With
            RS.MoveNext
        Loop
    End If
    vasRes.RowHeight(-1) = 12
    
End Sub

Private Sub vasID_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim iRow    As Long
    Dim iCol    As Long
    Dim lsID    As String
    Dim lsTime  As String
    Dim lsPid   As String
    Dim lsSeq   As String
    Dim i       As Integer
    Dim strResult As String
    Dim blnModify As Boolean
    
    blnModify = False
    
    iRow = vasID.ActiveRow
    iCol = vasID.ActiveCol

    If KeyCode = vbKeyDelete Then
        If iRow < 1 Or iRow > vasID.DataRowCnt Then
            Exit Sub
        End If
        If iCol > colState Then
            Exit Sub
        End If
        lsID = Trim(GetText(vasID, iRow, colBARCODE))
        lsPid = Trim(GetText(vasID, iRow, colPID))
        lsSeq = Trim(GetText(vasID, iRow, colSAVESEQ))

'        If lsID = "" Or lsPid = "" Or lsSeq = "" Then
'            Exit Sub
'        End If
        If lsID = "" Then
            Exit Sub
        End If

        If MsgBox(lsID & " �� ����� �����Ͻðڽ��ϱ�?", vbInformation + vbYesNo, "�˸�") = vbNo Then
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
        
        vasID.MaxRows = vasID.MaxRows - 1
        
        blnModify = True

    ElseIf KeyCode = vbKeyReturn Then
        If iCol = colBARCODE Then
            'Exit Sub
            
            '-- �ٲ� ���ڵ�� ȯ������ �ҷ�����
            Call GetSampleInfoW_NCC(iRow)
            
            lsID = Trim(GetText(vasID, iRow, colBARCODE))
            
            
            '-- ���ڵ� ��ȣ�� ������ Ʋ���ٸ� ������Ʈ
            'If lsID <> lblChangeBar.Caption Then
            'If lsID <> lblBarcode(0).Caption Then
                      SQL = "UPDATE PATRESULT SET"
                SQL = SQL & " HOSPDATE = '" & Format(Mid(Trim(GetText(vasID, iRow, colHOSPDATE)), 1, 10), "########") & "' " & vbCrLf
                SQL = SQL & ",BARCODE = '" & lsID & "' " & vbCrLf
                SQL = SQL & ",CHARTNO = '" & Trim(GetText(vasID, iRow, colCHARTNO)) & "' " & vbCrLf
                SQL = SQL & ",PID = '" & Trim(GetText(vasID, iRow, colPID)) & "' " & vbCrLf
                SQL = SQL & ",PNAME = '" & Trim(GetText(vasID, iRow, colPNAME)) & "' " & vbCrLf
                SQL = SQL & ",PSEX = '" & Trim(GetText(vasID, iRow, colPSEX)) & "' " & vbCrLf
                SQL = SQL & ",PAGE = '" & Trim(GetText(vasID, iRow, colPAGE)) & "' " & vbCrLf
                SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf
                SQL = SQL & "   AND SAVESEQ = " & Trim(GetText(vasID, iRow, colSAVESEQ)) & vbCrLf
                SQL = SQL & "   AND BARCODE = '" & lblBarcode(0).Caption & "' "

                'SetRawData "[SQL]" & SQL
                Res = SendQuery(gLocal, SQL)
                
                If Res = -1 Then
                    SaveQuery SQL
                    Exit Sub
                End If

                blnModify = True
        ElseIf iCol = colCHARTNO Then
            'Exit Sub
            
            If Trim(GetText(vasID, iRow, colBARCODE)) <> "" Then
                Exit Sub
            End If
            
            '-- �ٲ� ���ڵ�� ȯ������ �ҷ�����
            Call GetSampleInfoW_NCC(iRow)
            
            lsID = Trim(GetText(vasID, iRow, colBARCODE))
            
            
            '-- ���ڵ� ��ȣ�� ������ Ʋ���ٸ� ������Ʈ
            'If lsID <> lblChangeBar.Caption Then
            'If lsID <> lblBarcode(0).Caption Then
                      SQL = "UPDATE PATRESULT SET"
                SQL = SQL & " HOSPDATE = '" & Format(Mid(Trim(GetText(vasID, iRow, colHOSPDATE)), 1, 10), "########") & "' " & vbCrLf
                SQL = SQL & ",BARCODE = '" & lsID & "' " & vbCrLf
                SQL = SQL & ",CHARTNO = '" & Trim(GetText(vasID, iRow, colCHARTNO)) & "' " & vbCrLf
                SQL = SQL & ",PID = '" & Trim(GetText(vasID, iRow, colPID)) & "' " & vbCrLf
                SQL = SQL & ",PNAME = '" & Trim(GetText(vasID, iRow, colPNAME)) & "' " & vbCrLf
                SQL = SQL & ",PSEX = '" & Trim(GetText(vasID, iRow, colPSEX)) & "' " & vbCrLf
                SQL = SQL & ",PAGE = '" & Trim(GetText(vasID, iRow, colPAGE)) & "' " & vbCrLf
                SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf
                SQL = SQL & "   AND SAVESEQ = " & Trim(GetText(vasID, iRow, colSAVESEQ)) & vbCrLf
                SQL = SQL & "   AND BARCODE = '" & lblBarcode(0).Caption & "' "

                'SetRawData "[SQL]" & SQL
                Res = SendQuery(gLocal, SQL)
                
                If Res = -1 Then
                    SaveQuery SQL
                    Exit Sub
                End If

                blnModify = True
                
            'End If
        Else
            Exit Sub
            vasID.Row = iRow
            vasID.Col = colState
            If Trim(vasID.Text) = "" Then
                Exit Sub
            End If

            '-- ����� �������� ����� ������Ʈ�� Delete >> Insert ������ �Ѵ�.
            '-- Delete
                  SQL = "DELETE FROM PATRESULT "
            SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf
            SQL = SQL & "   AND SAVESEQ = " & Trim(GetText(vasID, iRow, colSAVESEQ)) & vbCrLf
            SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Trim(GetText(vasID, iRow, colEXAMDATE)) & "' " & vbCrLf
            SQL = SQL & "   AND BARCODE = '" & Trim(GetText(vasID, iRow, colBARCODE)) & "' "

            Res = SendQuery(gLocal, SQL)
                
            If Res = -1 Then
                SaveQuery SQL
                Exit Sub
            End If

            '-- Insert
            For i = colState + 1 To vasID.MaxCols
                vasID.Row = iRow
                vasID.Col = i
                If Trim(vasID.Text) <> "" Then
                    '-- ��� �Ҽ��� ����
                    strResult = SetResult(Trim(GetText(vasID, iRow, i)), gArrEquip(i - colState, 2))
                    '-- H/L �϶� ��ǥ��
                    If gsFlag = "L" Then
                        vasID.Row = iRow
                        vasID.Col = i
                        vasID.ForeColor = vbBlue
                    ElseIf gsFlag = "H" Then
                        vasID.Row = iRow
                        vasID.Col = i
                        vasID.ForeColor = vbRed
                    End If
                    vasID.Text = strResult

                    SQL = ""
                    SQL = SQL & "INSERT INTO PATRESULT (" & vbCrLf
                    SQL = SQL & "SAVESEQ, EXAMDATE, HOSPDATE, EQUIPNO, BARCODE" & vbCrLf
                    SQL = SQL & ", EQUIPCODE, EXAMCODE, EXAMSUBCODE, EXAMNAME, SEQNO" & vbCrLf
                    SQL = SQL & ", SAMPLETYPE, DISKNO, POSNO, EQUIPRESULT, RESULT" & vbCrLf
                    SQL = SQL & ", REFFLAG, REFVALUE, CHARTNO, PID, PNAME" & vbCrLf
                    SQL = SQL & ", PSEX, PAGE, PJUMIN, PANICVALUE, DELTAVALUE" & vbCrLf
                    SQL = SQL & ", SENDFLAG, SENDDATE, EXAMUID, HOSPITAL)" & vbCrLf
                    SQL = SQL & " VALUES (" & vbCrLf
                    SQL = SQL & Trim(GetText(vasID, iRow, colSAVESEQ))
                    SQL = SQL & ",'" & Trim(GetText(vasID, iRow, colEXAMDATE))
                    SQL = SQL & "','" & Trim(GetText(vasID, iRow, colHOSPDATE))
                    SQL = SQL & "','" & gEquip
                    SQL = SQL & "','" & Trim(GetText(vasID, iRow, colBARCODE))
                    'equipcode , examcode, examname, resprec, seqno
                    SQL = SQL & "','" & gArrEquip(i - colState, 2) 'Trim(GetText(vasRes, asRow2, colEQUIPCODE))
                    SQL = SQL & "','" & gArrEquip(i - colState, 3) 'Trim(GetText(vasRes, asRow2, colEXAMCODE))
                    SQL = SQL & "','"                              'Trim(GetText(vasRes, asRow2, colSubCode))
                    SQL = SQL & "','" & gArrEquip(i - colState, 4) 'Trim(GetText(vasRes, asRow2, colEXAMNAME))
                    SQL = SQL & "','" & gArrEquip(i - colState, 6) 'Trim(GetText(vasRes, asRow2, colSeq))
                    SQL = SQL & "',''"
                    SQL = SQL & ",'" & Trim(GetText(vasID, iRow, colDISKNO))
                    SQL = SQL & "','" & Trim(GetText(vasID, iRow, colPOSNO))
                    SQL = SQL & "','" & Trim(GetText(vasID, iRow, i)) 'Trim(GetText(vasRes, asRow2, colMachResult))
                    SQL = SQL & "','" & strResult 'Trim(GetText(vasID, iRow, i)) 'Trim(GetText(vasRes, asRow2, colRESULT))
                    SQL = SQL & "','" & gsFlag & "'"
                    SQL = SQL & ",''"
                    SQL = SQL & ",'" & Trim(GetText(vasID, iRow, colCHARTNO))
                    SQL = SQL & "','" & Trim(GetText(vasID, iRow, colPID))
                    SQL = SQL & "','" & Trim(GetText(vasID, iRow, colPNAME))
                    SQL = SQL & "','" & Trim(GetText(vasID, iRow, colPSEX))
                    SQL = SQL & "','" & Trim(GetText(vasID, iRow, colPAGE))
                    SQL = SQL & "',''"
                    SQL = SQL & ",''"
                    SQL = SQL & ",''"
                    SQL = SQL & ",'3'"
                    SQL = SQL & ",''"
                    SQL = SQL & ",'" & gIFUser
                    SQL = SQL & "','')"

                    Res = SendQuery(gLocal, SQL)
                    SetText vasID, "����", iRow, colState

                End If
            Next
            blnModify = True
        End If
        'SetText vasID, "����", iRow, colState

    End If
    
'    If blnModify = True Then
'        Call cmdRsltSearch_Click
'    End If
    
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
    Dim strResult   As String
    
    With vasRes
        If KeyAscii = 13 And .ActiveCol = colRESULT And lblBarcode(0).Caption <> "" Then
            '-- ��� �Ҽ��� ����
            strResult = SetResult(Trim(GetText(vasRes, .ActiveRow, colRESULT)), Trim(GetText(vasRes, .ActiveRow, colEQUIPCODE)))
            .Col = colRESULT
            .Text = strResult
            '-- H/L �϶� ��ǥ��
            If gsFlag = "L" Then
                vasRes.Row = .ActiveRow
                vasRes.Col = colRESULT
                vasRes.ForeColor = vbBlue
            ElseIf gsFlag = "H" Then
                vasRes.Row = .ActiveRow
                vasRes.Col = colRESULT
                vasRes.ForeColor = vbRed
            End If
            
            SetText vasRes, gsFlag, .ActiveRow, colFLAG
            
            SQL = ""
            SQL = SQL & "UPDATE PATRESULT " & vbCrLf
            SQL = SQL & "   SET RESULT  ='" & strResult & "', " & vbCrLf
            SQL = SQL & "       REFFLAG    = '" & gsFlag & "' " & vbCrLf
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

