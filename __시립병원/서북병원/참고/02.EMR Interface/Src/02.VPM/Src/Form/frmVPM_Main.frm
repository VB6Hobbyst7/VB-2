VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C9460280-3EED-11D0-A647-00A0C91EF7B9}#1.0#0"; "ImageViewer2.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmVPM_Main 
   Caption         =   "Hi Interface EMR(VPM)"
   ClientHeight    =   10110
   ClientLeft      =   6060
   ClientTop       =   1650
   ClientWidth     =   15030
   BeginProperty Font 
      Name            =   "����ü"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVPM_Main.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10110
   ScaleWidth      =   15030
   Begin VB.TextBox txtSerialData 
      Height          =   4095
      Left            =   5040
      MultiLine       =   -1  'True
      ScrollBars      =   2  '����
      TabIndex        =   52
      Top             =   5580
      Width           =   9915
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   3840
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   5
      DTREnable       =   -1  'True
      InputLen        =   1
      RThreshold      =   1
   End
   Begin FPSpread.vaSpread sprEQ_INFO 
      Height          =   1335
      Left            =   1740
      TabIndex        =   47
      Top             =   12300
      Visible         =   0   'False
      Width           =   2415
      _Version        =   393216
      _ExtentX        =   4260
      _ExtentY        =   2355
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
      SpreadDesigner  =   "frmVPM_Main.frx":9F8A
   End
   Begin MSComctlLib.ProgressBar prgPatient 
      Height          =   75
      Left            =   5040
      TabIndex        =   46
      Top             =   1080
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   132
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "����(&S)"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   13980
      Picture         =   "frmVPM_Main.frx":A1AB
      Style           =   1  '�׷���
      TabIndex        =   15
      Top             =   60
      Width           =   975
   End
   Begin VB.Frame fraLimageList 
      Caption         =   "[Image List]"
      Height          =   3975
      Left            =   60
      TabIndex        =   40
      Top             =   5700
      Width           =   4875
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '���
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   1
         Left            =   960
         ScaleHeight     =   285
         ScaleWidth      =   3825
         TabIndex        =   48
         Top             =   240
         Width           =   3855
         Begin VB.OptionButton optImage���ۿ��� 
            Caption         =   "���ۿϷ�"
            Height          =   300
            Index           =   1
            Left            =   1440
            TabIndex        =   6
            Top             =   0
            Width           =   1095
         End
         Begin VB.OptionButton optImage���ۿ��� 
            Caption         =   "������"
            Height          =   300
            Index           =   0
            Left            =   60
            TabIndex        =   5
            ToolTipText     =   "Double Click �ϸ� Image ���Ϸ� ��ȯ�˴ϴ�."
            Top             =   0
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.PictureBox picImageList������ 
         Appearance      =   0  '���
         ForeColor       =   &H80000008&
         Height          =   3255
         Left            =   60
         ScaleHeight     =   3225
         ScaleWidth      =   4725
         TabIndex        =   50
         Top             =   660
         Width           =   4755
         Begin VB.CommandButton cmdViewLocal 
            Caption         =   "�ڷ�ҷ�����"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   60
            TabIndex        =   8
            Top             =   2880
            Width           =   1695
         End
         Begin VB.CommandButton cmdDeleteLocal 
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3720
            TabIndex        =   10
            Top             =   2880
            Width           =   975
         End
         Begin VB.CommandButton cmd�������������� 
            Caption         =   "��������������"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1800
            TabIndex        =   9
            Top             =   2880
            Width           =   1875
         End
         Begin FPSpread.vaSpread spr������ 
            Height          =   2775
            Left            =   60
            TabIndex        =   7
            Top             =   60
            Width           =   4635
            _Version        =   393216
            _ExtentX        =   8176
            _ExtentY        =   4895
            _StockProps     =   64
            ColsFrozen      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����ü"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   3
            MaxRows         =   20
            OperationMode   =   3
            SpreadDesigner  =   "frmVPM_Main.frx":C7E5
         End
      End
      Begin VB.PictureBox picImageList���ۿϷ� 
         Appearance      =   0  '���
         ForeColor       =   &H80000008&
         Height          =   3255
         Left            =   60
         ScaleHeight     =   3225
         ScaleWidth      =   4725
         TabIndex        =   51
         Top             =   660
         Width           =   4755
         Begin VB.CommandButton cmdViewFTP 
            Caption         =   "�ڷ�ҷ�����"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   60
            TabIndex        =   12
            Top             =   2880
            Width           =   1695
         End
         Begin VB.CommandButton cmdDeleteFTP 
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3720
            TabIndex        =   14
            Top             =   2880
            Width           =   975
         End
         Begin VB.CommandButton cmd���ۿϷ��������� 
            Caption         =   "���ۿϷ���������"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1800
            TabIndex        =   13
            Top             =   2880
            Width           =   1875
         End
         Begin FPSpread.vaSpread spr���ۿϷ� 
            Height          =   2775
            Left            =   60
            TabIndex        =   11
            Top             =   60
            Width           =   4635
            _Version        =   393216
            _ExtentX        =   8176
            _ExtentY        =   4895
            _StockProps     =   64
            ColsFrozen      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����ü"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   3
            MaxRows         =   20
            OperationMode   =   3
            SpreadDesigner  =   "frmVPM_Main.frx":CC9D
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "���ۿ���"
         Height          =   180
         Index           =   10
         Left            =   120
         TabIndex        =   49
         Top             =   300
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "[�����ڳ���]"
      Height          =   4995
      Left            =   60
      TabIndex        =   36
      Top             =   600
      Width           =   4875
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '���
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   960
         ScaleHeight     =   285
         ScaleWidth      =   3825
         TabIndex        =   39
         Top             =   240
         Width           =   3855
         Begin VB.OptionButton opt���ۿ��� 
            Caption         =   "������"
            Height          =   300
            Index           =   0
            Left            =   60
            TabIndex        =   0
            ToolTipText     =   "Double Click �ϸ� Image ���Ϸ� ��ȯ�˴ϴ�."
            Top             =   0
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton opt���ۿ��� 
            Caption         =   "���ۿϷ�"
            Height          =   300
            Index           =   1
            Left            =   1440
            TabIndex        =   1
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.CommandButton cmdView 
         Caption         =   "��ȸ(&V)"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2340
         Style           =   1  '�׷���
         TabIndex        =   3
         Top             =   600
         Width           =   2475
      End
      Begin MSComCtl2.DTPicker dtp�������� 
         Height          =   315
         Left            =   960
         TabIndex        =   2
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   21430273
         CurrentDate     =   40449
      End
      Begin FPSpread.vaSpread spr�����ڳ��� 
         Height          =   3975
         Left            =   60
         TabIndex        =   4
         Top             =   960
         Width           =   4755
         _Version        =   393216
         _ExtentX        =   8387
         _ExtentY        =   7011
         _StockProps     =   64
         BackColorStyle  =   1
         ColsFrozen      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   12
         MaxRows         =   20
         OperationMode   =   3
         SelectBlockOptions=   0
         SpreadDesigner  =   "frmVPM_Main.frx":D171
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "���ۿ���"
         Height          =   180
         Index           =   14
         Left            =   120
         TabIndex        =   38
         Top             =   300
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "��������"
         Height          =   180
         Index           =   13
         Left            =   120
         TabIndex        =   37
         Top             =   660
         Width           =   720
      End
   End
   Begin MSComctlLib.StatusBar staCondition 
      Align           =   2  '�Ʒ� ����
      Height          =   375
      Left            =   0
      TabIndex        =   17
      Top             =   9735
      Width           =   15030
      _ExtentX        =   26511
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   5398
            MinWidth        =   3528
            Text            =   "Copyright �� 2010 Medimate Corp."
            TextSave        =   "Copyright �� 2010 Medimate Corp."
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12462
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "DB"
            TextSave        =   "DB"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "FTP"
            TextSave        =   "FTP"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "COM"
            TextSave        =   "COM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "PRN"
            TextSave        =   "PRN"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1940
            MinWidth        =   1940
            TextSave        =   "2011-12-13"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "���� 4:55"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picImage 
      Appearance      =   0  '���
      ForeColor       =   &H80000008&
      Height          =   8475
      Left            =   5040
      ScaleHeight     =   8445
      ScaleWidth      =   9885
      TabIndex        =   45
      Top             =   1200
      Width           =   9915
      Begin SCRIBBLELib.ImageViewer imvResult 
         Height          =   7335
         Left            =   60
         TabIndex        =   16
         Top             =   1080
         Width           =   9795
         _Version        =   65536
         _ExtentX        =   17277
         _ExtentY        =   12938
         _StockProps     =   0
         LicenseKey      =   "12595"
      End
      Begin VB.PictureBox picControl 
         Height          =   555
         Left            =   60
         ScaleHeight     =   495
         ScaleWidth      =   9735
         TabIndex        =   65
         Top             =   480
         Width           =   9795
         Begin VB.CommandButton cmdCenter 
            Caption         =   "Center Image"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   8700
            TabIndex        =   74
            Top             =   0
            Width           =   1035
         End
         Begin VB.CommandButton cmdZmHeight 
            Caption         =   "Zoom to Height"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   7620
            TabIndex        =   73
            Top             =   0
            Width           =   1035
         End
         Begin VB.CommandButton cmdZmWidth 
            Caption         =   "Zoom to Width"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   6540
            TabIndex        =   72
            Top             =   0
            Width           =   1035
         End
         Begin VB.CommandButton cmdFit 
            Caption         =   "Fit to Window"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   5460
            TabIndex        =   71
            Top             =   0
            Width           =   1035
         End
         Begin VB.CommandButton cmdRotate 
            Caption         =   "Rotate"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   4380
            TabIndex        =   70
            Top             =   0
            Width           =   1035
         End
         Begin VB.CommandButton cmd100 
            Caption         =   "Original"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3300
            TabIndex        =   69
            Top             =   0
            Width           =   1035
         End
         Begin VB.ComboBox cboZoomValue 
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   12
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1920
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   68
            Top             =   60
            Width           =   1035
         End
         Begin VB.CommandButton cmdzoomin 
            Caption         =   "��"
            Height          =   495
            Left            =   0
            TabIndex        =   67
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdzoomout 
            BackColor       =   &H00FFFFFF&
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   480
            TabIndex        =   66
            Top             =   0
            Width           =   480
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '����
            Caption         =   "����"
            Height          =   180
            Index           =   12
            Left            =   1440
            TabIndex        =   75
            Top             =   180
            Width           =   360
         End
      End
      Begin VB.PictureBox picJPG 
         Appearance      =   0  '���
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   60
         ScaleHeight     =   405
         ScaleWidth      =   9765
         TabIndex        =   63
         Top             =   60
         Width           =   9795
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "JPG File"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   12
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   180
            TabIndex        =   64
            Top             =   60
            Width           =   1080
         End
      End
      Begin VB.PictureBox picTifPdf 
         Height          =   435
         Left            =   60
         ScaleHeight     =   375
         ScaleWidth      =   9735
         TabIndex        =   53
         Top             =   60
         Width           =   9795
         Begin VB.CommandButton cmdMultiLast 
            Caption         =   "->|"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1440
            TabIndex        =   59
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdMultiNext 
            Caption         =   "->"
            Height          =   375
            Left            =   960
            TabIndex        =   58
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdMultiPrev 
            Caption         =   "<-"
            Height          =   375
            Left            =   480
            TabIndex        =   57
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdMultiFirst 
            Caption         =   "|<-"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   0
            TabIndex        =   56
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdMultiJump 
            Caption         =   "Go"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3420
            TabIndex        =   55
            Top             =   0
            Width           =   495
         End
         Begin VB.TextBox txtMultiPno 
            Alignment       =   1  '������ ����
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   2940
            TabIndex        =   54
            Text            =   "1"
            Top             =   0
            Width           =   375
         End
         Begin VB.Label Label3 
            Caption         =   "Total Page"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4230
            TabIndex        =   62
            Top             =   60
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Page No."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2040
            TabIndex        =   61
            Top             =   60
            Width           =   855
         End
         Begin VB.Label lblMultiCnt 
            Alignment       =   1  '������ ����
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   5400
            TabIndex        =   60
            Top             =   60
            Width           =   615
         End
      End
   End
   Begin VB.Label lbló����� 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  '����
      Caption         =   "�ܷ�"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   10380
      TabIndex        =   77
      Top             =   840
      Width           =   390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "ó�����"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   180
      Index           =   15
      Left            =   9420
      TabIndex        =   76
      Top             =   840
      Width           =   780
   End
   Begin VB.Label lbló���ڵ� 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  '����
      Caption         =   "1234567890"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   5940
      TabIndex        =   44
      Top             =   600
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "ó���"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   180
      Index           =   6
      Left            =   7380
      TabIndex        =   43
      Top             =   600
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "ó��SEQ"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   180
      Index           =   5
      Left            =   11820
      TabIndex        =   42
      Top             =   600
      Width           =   705
   End
   Begin VB.Label lbló��SEQ 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  '����
      Caption         =   "1234567890"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   12720
      TabIndex        =   41
      Top             =   600
      Width           =   1050
   End
   Begin VB.Label lbl���� 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "�˻�����"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   180
      TabIndex        =   35
      Top             =   60
      Width           =   1800
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "���Ϲ�ȣ"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   180
      Index           =   0
      Left            =   5100
      TabIndex        =   34
      Top             =   360
      Width           =   780
   End
   Begin VB.Label lbl���Ϲ�ȣ 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  '����
      Caption         =   "1234567890"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   5940
      TabIndex        =   33
      Top             =   360
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "�����ڸ�"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   180
      Index           =   1
      Left            =   7380
      TabIndex        =   32
      Top             =   360
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "����/����"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   180
      Index           =   2
      Left            =   9420
      TabIndex        =   31
      Top             =   360
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "�����"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   180
      Index           =   3
      Left            =   5100
      TabIndex        =   30
      Top             =   840
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "ȯ�ڱ���"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   180
      Index           =   4
      Left            =   7380
      TabIndex        =   29
      Top             =   840
      Width           =   780
   End
   Begin VB.Label lbl�����ڸ� 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  '����
      Caption         =   "ȫ�浿"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   8340
      TabIndex        =   28
      Top             =   360
      Width           =   585
   End
   Begin VB.Label lbl���ɼ��� 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  '����
      Caption         =   "40/��"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   10380
      TabIndex        =   27
      Top             =   360
      Width           =   510
   End
   Begin VB.Label lbl����� 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  '����
      Caption         =   "IM"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   5940
      TabIndex        =   26
      Top             =   840
      Width           =   210
   End
   Begin VB.Label lbl�Կܱ��� 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  '����
      Caption         =   "�ܷ�"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   8340
      TabIndex        =   25
      Top             =   840
      Width           =   390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "Patient/Image Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   9
      Left            =   5160
      TabIndex        =   24
      Top             =   60
      Width           =   2640
   End
   Begin VB.Label lbl������� 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  '����
      Caption         =   "1234567890"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   12720
      TabIndex        =   23
      Top             =   840
      Width           =   1050
   End
   Begin VB.Label lbló������ 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  '����
      Caption         =   "1234567890"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   12720
      TabIndex        =   22
      Top             =   360
      Width           =   1050
   End
   Begin VB.Label lbló��� 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  '����
      Caption         =   "1234567890"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   8340
      TabIndex        =   21
      Top             =   600
      Width           =   3390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "�������"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   180
      Index           =   11
      Left            =   11820
      TabIndex        =   20
      Top             =   840
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "ó���ڵ�"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   180
      Index           =   8
      Left            =   5100
      TabIndex        =   19
      Top             =   600
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "ó������"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   180
      Index           =   7
      Left            =   11820
      TabIndex        =   18
      Top             =   360
      Width           =   780
   End
   Begin VB.Shape shpPatientInfo 
      BackColor       =   &H00808000&
      BackStyle       =   1  '�������� ����
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFC0C0&
      FillStyle       =   5  '���� �밢��
      Height          =   255
      Left            =   5040
      Shape           =   4  '�ձ� �簢��
      Top             =   60
      Width           =   8835
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808000&
      BackStyle       =   1  '�������� ����
      BorderColor     =   &H00000000&
      FillColor       =   &H0000C000&
      FillStyle       =   5  '���� �밢��
      Height          =   495
      Index           =   3
      Left            =   60
      Shape           =   4  '�ձ� �簢��
      Top             =   60
      Width           =   4875
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File    "
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuSetting 
      Caption         =   "����    "
      Begin VB.Menu mnuSettingSub 
         Caption         =   "DataBase Info Setting"
         Index           =   0
      End
      Begin VB.Menu mnuSettingSub 
         Caption         =   "Target Equipment Setting"
         Index           =   1
      End
      Begin VB.Menu mnuSettingSub 
         Caption         =   "Equipment Config"
         Index           =   2
      End
   End
   Begin VB.Menu mnuInfo 
      Caption         =   "����"
   End
End
Attribute VB_Name = "frmVPM_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lngMeHeight     As Long '/Me.Height�� �ʱⰪ
Dim lngMeWidth      As Long '/Me.Width�� �ʱⰪ

Private Type ConWhere   ' ����� ���� ������ ����ϴ�.
   Nm       As String
   Left     As Long
   Top      As Long
   Width    As Long
   Height   As Long
End Type
Dim CW()    As ConWhere

Private iPicCnt     As Integer
Private MMFTP       As New cls����_FTP
Private MMSFTP      As New cls����_SFTP

Private mintCurFrame As Integer ' ���� �������� ���Դϴ�.

Public Function FUNC_MM_CANCEL() As Boolean
    lbl���� = ""
    prgPatient.Max = 100
    prgPatient.Value = 100
    
    '/�����ڳ���
    opt���ۿ���(0).Value = True
    opt���ۿ���(0).ForeColor = RGB(0, 0, 255)
    opt���ۿ���(0).FontBold = True
    
    dtp��������.Value = Format(Now, "YYYY-MM-DD")
    
    optImage���ۿ���(0).ForeColor = RGB(0, 0, 255)
    optImage���ۿ���(0).FontBold = True
    
    Call FUNC_MM_KEY_CLEAR("1") '/�����ڳ��� Spread Clear
    Call FUNC_MM_KEY_CLEAR("2") '/Patient Information Clear
    Call FUNC_MM_KEY_CLEAR("3") '/Image Clear
    Call FUNC_MM_KEY_CLEAR("4") '/������ Spread Clear
    Call FUNC_MM_KEY_CLEAR("5") '/���ۿϷ� Spread Clear
    
    picImageList������.Visible = True
    picImageList���ۿϷ�.Visible = False

    mnuSetting.Visible = False '/�����޴� �Ⱥ��̱�

    txtSerialData.Visible = False
    
    picTifPdf.Visible = False
    picJPG.Visible = True
End Function

Public Function FUNC_MM_DELETE(ArgSection As String) As Boolean
    Dim strIMGFILEPATH  As String
    
    FUNC_MM_DELETE = False
    
On Error GoTo ERR_RTN
    
    Select Case ArgSection
        Case "1": GoSub DELETE1_RTN '/������ ����
        Case "2": GoSub DELETE2_RTN '/���ۿϷ� ����
    End Select
    
    FUNC_MM_DELETE = True

Exit Function

'/----------------------------------------------------------------------------------------------------/

DELETE1_RTN:
    imvResult.Filename = ""
    Call SUB_CHK_PDF_TIF("")

    For intX = 1 To spr������.MaxRows
        If GET_CELL(spr������, 1, intX) = "1" Then
            Kill gtypEQ_INFO.EQIMGFILEPATH & "\" & GET_CELL(spr������, 2, intX)
        End If
    Next intX
Return

'/----------------------------------------------------------------------------------------------------/

DELETE2_RTN:
    '/----------------------------------------------------------------------------------------------------/
    '/Step4.    FTP ����
    '/----------------------------------------------------------------------------------------------------/
    If OpenDB(gstrREG_DB_CONSTR) = False Then End
    
    gstrQuy = "SELECT * "
    gstrQuy = gstrQuy & vbCrLf & "  FROM MM_EMR_RES "
    gstrQuy = gstrQuy & vbCrLf & " WHERE PATNO     = '" & lbl���Ϲ�ȣ & "' "
    gstrQuy = gstrQuy & vbCrLf & "   AND ORDDATE   = '" & Replace(lbló������, "-", "") & "' "
    gstrQuy = gstrQuy & vbCrLf & "   AND ORDSEQ    =  " & Val(lbló��SEQ) & " "
    gstrQuy = gstrQuy & vbCrLf & "   AND EQUIPCODE = '" & gtypEQ_INFO.EQUIPCODE & "' "
    gstrQuy = gstrQuy & vbCrLf & "   AND EQUIPSEQ  =  " & gtypEQ_INFO.EQUIPSEQ & " "
    If ReadSQL(gstrQuy, ADR) = False Then Call CloseDB: End
    
    If Not ADR Is Nothing Then
        strIMGFILEPATH = Trim(ADR!IMGFILEPATH & "")                              '/FTP���
        
        ADR.Close: Set ADR = Nothing
    End If

    Call CloseDB
    
    Dim success As Long
    success = sftp.IsConnected
    If (success <> 1) Then
        If MMSFTP.OpenConnection(gstrFTP_RH, gstrFTP_RP, gstrFTP_UN, gstrFTP_PW) = False Then
            MsgBox "Image File Server�� ������ �� �����ϴ�." & vbCrLf & "����ǿ� ���ǹٶ��ϴ�.", vbCritical, "���ۺҰ�"
            Exit Function
        End If
    End If
    
    If MMSFTP.SetFTPDirectory(strIMGFILEPATH) = True Then
        For intX = 1 To spr���ۿϷ�.MaxRows
            If GET_CELL(spr���ۿϷ�, 1, intX) = "1" Then
'''                If MMFTP.DeleteFTPFile(GET_CELL(spr���ۿϷ�, 2, intX)) = True Then
'''                    Kill gtypEQ_INFO.FTPIMGFILEPATH & "\" & GET_CELL(spr���ۿϷ�, 2, intX)
'''                Else
'''                    Exit Function
'''                End If
            
                '/���Ϻ��� ����View OCX(��ȭ�� �����δ��� ��������)���Ͽ����� JPG ���ϸ� �����´�. BAK������ DownLoad ���� �����Ƿ� Rename�ص� �������.
                If MMSFTP.RenameFTPFile(strIMGFILEPATH & GET_CELL(spr���ۿϷ�, 2, intX), strIMGFILEPATH & Left(GET_CELL(spr���ۿϷ�, 2, intX), InStr(GET_CELL(spr���ۿϷ�, 2, intX), ".")) & "bak") = True Then
                    Kill gtypEQ_INFO.FTPIMGFILEPATH & "\" & GET_CELL(spr���ۿϷ�, 2, intX)
                Else
                    Exit Function
                End If
            End If
        Next intX
        
        Call MMSFTP.FtpScanDirectory(strIMGFILEPATH)
    
        If UBound(FtpScanFileName_IMG) = 0 Then '/�������� FTP�ڷᰡ ������...
            If OpenDB(gstrREG_DB_CONSTR) = False Then End
            
            ADC.BeginTrans
            
            gstrQuy = "DELETE FROM MM_EMR_RES "
            gstrQuy = gstrQuy & vbCrLf & " WHERE PATNO     = '" & lbl���Ϲ�ȣ & "' "
            gstrQuy = gstrQuy & vbCrLf & "   AND ORDDATE   = '" & Replace(lbló������, "-", "") & "' "
            gstrQuy = gstrQuy & vbCrLf & "   AND ORDSEQ    =  " & Val(lbló��SEQ) & " "
            gstrQuy = gstrQuy & vbCrLf & "   AND EQUIPCODE = '" & gtypEQ_INFO.EQUIPCODE & "' "
            gstrQuy = gstrQuy & vbCrLf & "   AND EQUIPSEQ  =  " & gtypEQ_INFO.EQUIPSEQ & " "
            If RunSQL(gstrQuy) = False Then ADC.RollbackTrans: Call CloseDB: End
            
            '/��ó��:   TMPRSCINFN
            gstrQuy = "UPDATE TMPRSCINFN SET "
            gstrQuy = gstrQuy & vbCrLf & "      PRSC_STAT_CD    = '440', "                                              '/440.����
            gstrQuy = gstrQuy & vbCrLf & "      UPDTR_ID        = '" & gtypUSER.USERID & "', "                          '/���������� ID
            gstrQuy = gstrQuy & vbCrLf & "      UPDT_DT         = SYSDATE "                                             '/������������
            gstrQuy = gstrQuy & vbCrLf & "WHERE PRSC_DATE       = '" & Format(CDate(lbló������), "YYYYMMDD") & "' "    '/ó������
            gstrQuy = gstrQuy & vbCrLf & "  AND PRSC_NO         =  " & lbló��SEQ & " "                                 '/ó���ȣ
            gstrQuy = gstrQuy & vbCrLf & "  AND PID             = '" & lbl���Ϲ�ȣ & "' "                               '/���Ϲ�ȣ
            gstrQuy = gstrQuy & vbCrLf & "  AND PRSC_CD         = '" & lbló���ڵ� & "' "                               '/ó���ڵ�
            gstrQuy = gstrQuy & vbCrLf & "  AND PRSC_VALD_YN    = 'Y' "                                                 '/��ó�� ����ִ� ó��
            gstrQuy = gstrQuy & vbCrLf & "  AND PRSC_HSTR_CD    = 'O' "                                                 '/ó��History ��ȣ
            If RunSQL(gstrQuy) = False Then ADC.RollbackTrans: Call CloseDB: End

            '/�ǽ�ó��: TMPRSCEXCN
            gstrQuy = "UPDATE TMPRSCEXCN SET "
            gstrQuy = gstrQuy & vbCrLf & "      CNDT_PRSC_STAT_CD   = '440', "                                              '/440.����
            gstrQuy = gstrQuy & vbCrLf & "      UPDTR_ID            = '" & gtypUSER.USERID & "', "                          '/���������� ID
            gstrQuy = gstrQuy & vbCrLf & "      UPDT_DT             = SYSDATE "                                             '/������������
            gstrQuy = gstrQuy & vbCrLf & "WHERE PRSC_DATE           = '" & Format(CDate(lbló������), "YYYYMMDD") & "' "    '/ó������
            gstrQuy = gstrQuy & vbCrLf & "  AND PRSC_NO             =  " & lbló��SEQ & " "                                 '/ó���ȣ
            gstrQuy = gstrQuy & vbCrLf & "  AND PID                 = '" & lbl���Ϲ�ȣ & "' "                               '/���Ϲ�ȣ
            gstrQuy = gstrQuy & vbCrLf & "  AND MEFE_CD             = '" & lbló���ڵ� & "' "                               '/ó���ڵ�
            gstrQuy = gstrQuy & vbCrLf & "  AND CNDT_PRSC_VALD_YN   = 'Y' "                                                 '/�ǽ�ó�� ����ִ� ó��
            If RunSQL(gstrQuy) = False Then ADC.RollbackTrans: Call CloseDB: End

            ADC.CommitTrans
            
            Call CloseDB
        End If
        
        Call MMSFTP.FtpScanDirectory(strIMGFILEPATH)
        If UBound(FtpScanFileName) = 0 Then '/BackUp �� FTP�ڷᵵ ������...
            Call MMSFTP.RemoveFTPDirectory(strIMGFILEPATH)
        End If
    End If
    
    Call MMSFTP.CloseConnection
Return

'/----------------------------------------------------------------------------------------------------/

ERR_RTN:
    MsgBox "���� ����!!!", vbCritical, "Ȯ��"
End Function

Public Sub FUNC_MM_INITIAL()
    '/Resize�� ���� �ʱ� Size Setting----------------------------------------------------------------------------------------------------/
    For intX = 0 To Me.Count - 1
        Select Case True
            Case TypeOf Me.Controls(intX) Is Menu
            Case TypeOf Me.Controls(intX) Is Line
            Case TypeOf Me.Controls(intX) Is MSComm
            Case TypeOf Me.Controls(intX) Is CommonDialog
            Case Else
                ReDim Preserve CW(intX)
                
                CW(intX).Nm = Me.Controls(intX).Name
                CW(intX).Left = Me.Controls(intX).Left
                CW(intX).Top = Me.Controls(intX).Top
                CW(intX).Width = Me.Controls(intX).Width
                CW(intX).Height = Me.Controls(intX).Height
        End Select
    Next intX
    
    '/Form Size Setting
    lngMeHeight = 10890
    lngMeWidth = 15150
    
    Me.Height = lngMeHeight
    Me.Width = lngMeWidth
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Show
    '/Resize�� ���� �ʱ� Size Setting----------------------------------------------------------------------------------------------------/
    
    '/�ʱ� �ڷ� Setting----------------------------------------------------------------------------------------------------/
    GoSub ADD_ITEM
    '/�ʱ� �ڷ� Setting----------------------------------------------------------------------------------------------------/
    
    '/���� ��Ʈ�� �ʱ�ȭ----------------------------------------------------------------------------------------------------/
    Call FUNC_MM_CANCEL
    '/���� ��Ʈ�� �ʱ�ȭ----------------------------------------------------------------------------------------------------/
    
    lbl���� = gtypEQ_INFO.EQUIPNM
    
    '/�۾� ���� Check----------------------------------------------------------------------------------------------------/
    staCondition.Panels.Item(3).Enabled = False
    staCondition.Panels.Item(4).Enabled = False
    staCondition.Panels.Item(5).Enabled = False
    staCondition.Panels.Item(6).Enabled = False

    staCondition.Panels.Item(5).Visible = False
    staCondition.Panels.Item(6).Visible = False

    If gstrSTAUS_DB = "Y" Then
        staCondition.Panels.Item(3).Enabled = True
    End If
    If gstrSTAUS_FTP = "Y" Then
        staCondition.Panels.Item(4).Enabled = True
    End If
    
    If gtypEQ_INFO.SERIALYN = "Y" Then
        staCondition.Panels.Item(5).Visible = True
        
        On Error GoTo RTN_ERR_PORT
        
        MSComm1.CommPort = gtypEQ_INFO.SERIALPORT
        MSComm1.RTSEnable = gtypEQ_INFO.SERIALRTS
        MSComm1.DTREnable = gtypEQ_INFO.SERIALDTR
        MSComm1.Settings = gtypEQ_INFO.SERIALBAUD & "," & gtypEQ_INFO.SERIALPARITY & "," & gtypEQ_INFO.SERIALDATABIT & "," & gtypEQ_INFO.SERIALSTOPBIT
        
        If MSComm1.PortOpen = False Then
            MSComm1.PortOpen = True
        
            staCondition.Panels.Item(5).Enabled = True
        End If
    End If

    If gtypEQ_INFO.ZIPYN = "Y" Then
        staCondition.Panels.Item(6).Visible = True
        
        If Trim(gtypEQ_INFO.ZIPNM) <> "" Then
            Dim X As Printer
            
            strTemp = ""
            For Each X In Printers
                If Trim(gtypEQ_INFO.ZIPNM) = X.DeviceName Then
                    staCondition.Panels.Item(6).Enabled = True
                    
                    strTemp = "Y"
                    Exit For
                End If
            Next X
            
            If strTemp = "" Then
                MsgBox "�ش� ���� ���������͸� ����ؾ��ϴ� ����Դϴ�." & vbCrLf & _
                       "���������� ������ ���ų� �ٸ��Ƿ� ���������͸� (��)�����Ͻʽÿ�!", vbInformation, "�˸�"
                frm����_Set_Equip_Config.Show vbModal
            End If
        Else
            MsgBox "�ش� ���� ���������͸� ����ؾ��ϴ� ����Դϴ�." & vbCrLf & _
                   "���������� ������ ���ų� �ٸ��Ƿ� ���������͸� (��)�����Ͻʽÿ�!", vbInformation, "�˸�"
            frm����_Set_Equip_Config.Show vbModal
        End If
    End If
    '/�۾� ���� Check----------------------------------------------------------------------------------------------------/
Exit Sub

'/----------------------------------------------------------------------------------------------------/

ADD_ITEM:
    Me.Caption = Me.Caption & Space(10) & "(�����: " & gtypUSER.USERNM & " )"
    
    '/Zoom ����
    cboZoomValue.Clear
    cboZoomValue.AddItem ""
    cboZoomValue.AddItem "25%"
    cboZoomValue.AddItem "33%"
    cboZoomValue.AddItem "50%"
    cboZoomValue.AddItem "75%"
    cboZoomValue.AddItem "100%"
    cboZoomValue.AddItem "150%"
    cboZoomValue.AddItem "200%"
    
    '''Call SUB_GET_REG_CLIENT_INFO '/���� ��� Image ���� ���� �� �������� �����ҷ��� ����(����� �ǹ̾���)
Return

'/----------------------------------------------------------------------------------------------------/

RTN_ERR_PORT:
    If Err = 8002 Then      'Port
        staCondition.Panels.Item(5).Enabled = False
        
        MsgBox "��� ��Ʈ�� Ȯ���ϼ���!", vbInformation, "�˸�"
        frm����_Set_Equip_Config.Show vbModal
    Else
        Resume Next
    End If
End Sub

Public Sub FUNC_MM_KEY_CLEAR(ArgSection As String)
    Select Case ArgSection
        Case "1" '/�����ڳ��� Spread Clear
            If spr�����ڳ���.MaxRows > 0 Then spr�����ڳ���.MaxRows = 0
            
        Case "2" '/Patient Information Clear
            lbl���Ϲ�ȣ = ""
            lbl�����ڸ� = ""
            lbl���ɼ��� = ""
            lbló���ڵ� = ""
            lbló��� = ""
            lbl����� = ""
            lbl�Կܱ��� = ""
            lbló������ = ""
            lbló��SEQ = ""
            lbl������� = ""
            lbló����� = ""
            lbló�����.ForeColor = RGB(0, 0, 0)
            
        Case "3": '/Image Clear
            imvResult.Filename = ""
            Call SUB_CHK_PDF_TIF("")
            
        Case "4": '/������ Spread Clear
            If spr������.MaxRows > 0 Then spr������.MaxRows = 0
    
        Case "5": '/���ۿϷ� Spread Clear
            If spr���ۿϷ�.MaxRows > 0 Then spr���ۿϷ�.MaxRows = 0
    End Select
End Sub

Public Function FUNC_MM_PRINT() As Boolean
'''    Dim strFont1  As String
'''    Dim strFont2  As String
'''    Dim strHead1  As String
'''
'''    If sprVIEW.MaxRows = 0 Then MsgBox "����� �ڷᰡ �����ϴ�.", vbInformation, "Ȯ��": Exit Function
'''
'''    If MsgBox("����ϰڽ��ϱ�?", vbQuestion + vbOKCancel, "��¿���") = vbCancel Then Exit Function
'''
'''    strFont1 = "/fn""����ü""/fz""15""/fb1/fi0/fu1/fk0/fs1"
'''    strFont2 = "/fn""����ü""/fz""10""/fb0/fi0/fu0/fk0/fs2"
'''
'''    strHead1 = "/f1/c" & "�ŷ�ó �ڵ�" & "/n/n/n"
'''
'''    With sprVIEW
'''        .PrintAbortMsg = "�ŷ�ó �ڵ� ��� ��..."
'''        .PrintHeader = strFont1 + strHead1 + strFont2
'''        .PrintFooter = "/c" & "PAGE : " & "/P"
'''        .PrintBorder = True
'''        .PrintGrid = True
'''        .PrintColHeaders = True
'''        .PrintRowHeaders = True
'''        .PrintColor = False
'''        .PrintMarginTop = 500
'''        .PrintMarginBottom = 500
'''        .PrintMarginLeft = 500
'''        .PrintMarginRight = 0
'''        .PrintType = PrintTypeAll
'''        .PrintShadows = False
'''        .PrintUseDataMax = False
'''        .Action = ActionSmartPrint
'''    End With
End Function

Public Function FUNC_MM_SAVE(argintRow�����ڳ��� As Integer) As Boolean
    Dim rtn
    Dim strFileName     As String
    Dim nImageSeq       As Integer
    Dim strTemp1
    Dim strIMGFILEPATH  As String
    
    Dim fstró������    As String
    Dim fstr���Ϲ�ȣ    As String
    Dim fstró��SEQ     As String
    Dim fstr�����      As String
    Dim fstró���ڵ�    As String
    Dim fstr�Կܱ���    As String
    
    fstró������ = GET_CELL(spr�����ڳ���, 7, argintRow�����ڳ���)
    fstr���Ϲ�ȣ = GET_CELL(spr�����ڳ���, 2, argintRow�����ڳ���)
    fstró��SEQ = GET_CELL(spr�����ڳ���, 8, argintRow�����ڳ���)
    fstr����� = GET_CELL(spr�����ڳ���, 4, argintRow�����ڳ���)
    fstró���ڵ� = GET_CELL(spr�����ڳ���, 11, argintRow�����ڳ���)
    fstr�Կܱ��� = GET_CELL(spr�����ڳ���, 3, argintRow�����ڳ���)
    
    FUNC_MM_SAVE = False
    
On Error GoTo RTN_ERR

    '/----------------------------------------------------------------------------------------------------/
    '/Step4.    FTP ����
    '/----------------------------------------------------------------------------------------------------/
    Dim success As Long
    success = sftp.IsConnected
    If (success <> 1) Then
        If MMSFTP.OpenConnection(gstrFTP_RH, gstrFTP_RP, gstrFTP_UN, gstrFTP_PW) = False Then
            MsgBox "Image File Server�� ������ �� �����ϴ�." & vbCrLf & "����ǿ� ���ǹٶ��ϴ�.", vbCritical, "���ۺҰ�"
            Exit Function
        End If
    End If
   '/----------------------------------------------------------------------------------------------------/
    '/Step5.    FTP ������ �ش� ���� ����/ ������ SKIP
    '/----------------------------------------------------------------------------------------------------/
    '/�⺻ Image ���� �̵�
    strIMGFILEPATH = ""
    
    If MMSFTP.SetFTPDirectory("upload/lis") = False Then
        'MsgBox "FTP ������ EMR_Image ������ �����ϴ�.", vbInformation, "Ȯ��"
        'Exit Sub
        Call MMSFTP.CreateFTPDirectory("upload/lis")
    End If
    strIMGFILEPATH = "upload/lis/"
    
    '/����ڵ� ���� �̵� �� ����
    If MMSFTP.SetFTPDirectory(strIMGFILEPATH & gtypEQ_INFO.EQUIPCODE & gtypEQ_INFO.EQUIPSEQ) = False Then
        If MMSFTP.CreateFTPDirectory(strIMGFILEPATH & gtypEQ_INFO.EQUIPCODE & gtypEQ_INFO.EQUIPSEQ) = False Then
            MsgBox "FTP ���� ����(����ڵ�&���SEQ) �� ������ �߻��Ͽ����ϴ�.", vbCritical, "���ۺҰ�"
            Exit Function
        End If
        Call MMSFTP.SetFTPDirectory(strIMGFILEPATH & gtypEQ_INFO.EQUIPCODE & gtypEQ_INFO.EQUIPSEQ)
    End If
    strIMGFILEPATH = strIMGFILEPATH & gtypEQ_INFO.EQUIPCODE & gtypEQ_INFO.EQUIPSEQ & "/"

    '/ó��⵵ ���� �̵� �� ����
    If MMSFTP.SetFTPDirectory(strIMGFILEPATH & Left(Replace(fstró������, "-", ""), 4)) = False Then
        If MMSFTP.CreateFTPDirectory(strIMGFILEPATH & Left(Replace(fstró������, "-", ""), 4)) = False Then
            MsgBox "FTP ���� ����(ó��⵵) �� ������ �߻��Ͽ����ϴ�.", vbCritical, "���ۺҰ�"
            Exit Function
        End If
        Call MMSFTP.SetFTPDirectory(strIMGFILEPATH & Left(Replace(fstró������, "-", ""), 4))
    End If
    strIMGFILEPATH = strIMGFILEPATH & Left(Replace(fstró������, "-", ""), 4) & "/"
    
    '/ó����� ���� �̵� �� ����
    If MMSFTP.SetFTPDirectory(strIMGFILEPATH & Mid(Replace(fstró������, "-", ""), 5)) = False Then
        If MMSFTP.CreateFTPDirectory(strIMGFILEPATH & Mid(Replace(fstró������, "-", ""), 5)) = False Then
            MsgBox "FTP ���� ����(ó�����) �� ������ �߻��Ͽ����ϴ�.", vbCritical, "���ۺҰ�"
            Exit Function
        End If
        Call MMSFTP.SetFTPDirectory(strIMGFILEPATH & Mid(Replace(fstró������, "-", ""), 5))
    End If
    strIMGFILEPATH = strIMGFILEPATH & Mid(Replace(fstró������, "-", ""), 5) & "/"

    '/���Ϲ�ȣ ���� �̵� �� ����
    If MMSFTP.SetFTPDirectory(strIMGFILEPATH & fstr���Ϲ�ȣ) = False Then
        If MMSFTP.CreateFTPDirectory(strIMGFILEPATH & fstr���Ϲ�ȣ) = False Then
            MsgBox "FTP ���� ����(���Ϲ�ȣ) �� ������ �߻��Ͽ����ϴ�.", vbCritical, "���ۺҰ�"
            Exit Function
        End If
        Call MMSFTP.SetFTPDirectory(strIMGFILEPATH & fstr���Ϲ�ȣ)
    End If
    strIMGFILEPATH = strIMGFILEPATH & fstr���Ϲ�ȣ & "/"

    '/ó��SEQ(���� ������ȣ) ���� �̵� �� ����
    If MMSFTP.SetFTPDirectory(strIMGFILEPATH & fstró��SEQ) = False Then
        If MMSFTP.CreateFTPDirectory(strIMGFILEPATH & fstró��SEQ) = False Then
            MsgBox "FTP ���� ����(ó��SEQ) �� ������ �߻��Ͽ����ϴ�.", vbCritical, "���ۺҰ�"
            Exit Function
        End If
        Call MMSFTP.SetFTPDirectory(strIMGFILEPATH & fstró��SEQ)
    End If
    strIMGFILEPATH = strIMGFILEPATH & fstró��SEQ & "/"
    
    '/----------------------------------------------------------------------------------------------------/
    '/Step6.    ������ Ȯ���� ������ ���ϸ� ����
    '/----------------------------------------------------------------------------------------------------/
    strFileName = Replace(fstró������, "-", "") & "@" & fstró��SEQ & "@" & fstr���Ϲ�ȣ & "@"

    '/----------------------------------------------------------------------------------------------------/
    '/Step7.    FTP ������ �ش� ������ �ڷᰡ ������ �ִ� Ȯ���� �� ã��/�ڷᰡ ������ 0
    '/----------------------------------------------------------------------------------------------------/
    Call MMSFTP.FtpScanDirectory(strIMGFILEPATH, strFileName & "*.*")
    
    nImageSeq = 0
    If UBound(FtpScanFileName) > 0 Then
        For intX = 1 To UBound(FtpScanFileName)
            strTemp1 = Split(FtpScanFileName(intX), "@")
            If UBound(strTemp1) = 3 Then
                If Val(Left(strTemp1(3), InStr(strTemp1(3), ".") - 1)) > nImageSeq Then
                    nImageSeq = Val(Left(strTemp1(3), InStr(strTemp1(3), ".") - 1))
                End If
            End If
        Next intX
    End If
    
    '/----------------------------------------------------------------------------------------------------/
    '/Step8.    ���õ� ������ �ڷ� Rename �ϸ鼭 ����
    '/----------------------------------------------------------------------------------------------------/
    For intX = 1 To spr������.MaxRows
        If GET_CELL(spr������, 1, intX) = "1" Then
            nImageSeq = nImageSeq + 1
            rtn = MMSFTP.FTPUploadFile(gtypEQ_INFO.EQIMGFILEPATH & "\" & GET_CELL(spr������, 2, intX), strIMGFILEPATH & strFileName & Format(nImageSeq, "000") & Mid(GET_CELL(spr������, 2, intX), InStr(GET_CELL(spr������, 2, intX), ".")))
        End If
    Next intX
    
    '/----------------------------------------------------------------------------------------------------/
    '/Step9.    FTP ����
    '/----------------------------------------------------------------------------------------------------/
    Call MMSFTP.CloseConnection
    
    '/----------------------------------------------------------------------------------------------------/
    '/Step10.   ��� ���� ���� �� HIS�� �˻�Ϸ� Falg Update ����.
    '/----------------------------------------------------------------------------------------------------/
    If OpenDB(gstrREG_DB_CONSTR) = False Then End
    
    ADC.BeginTrans
    
    '/HIS ������� Update
    '''If FUNC_HIS_RST_UPDATE = False Then ADC.RollbackTrans: Call CloseDB: End
    '/��ó��:   TMPRSCINFN
    gstrQuy = "UPDATE TMPRSCINFN SET "
    gstrQuy = gstrQuy & vbCrLf & "      PRSC_STAT_CD    = '560', "                      '/560.�ӽð�� �̻�
    gstrQuy = gstrQuy & vbCrLf & "      UPDTR_ID        = '" & gtypUSER.USERID & "', "  '/���������� ID
    gstrQuy = gstrQuy & vbCrLf & "      UPDT_DT         = SYSTIMESTAMP "                '/������������
    gstrQuy = gstrQuy & vbCrLf & "WHERE PRSC_DATE       = '" & Format(CDate(fstró������), "YYYYMMDD") & "' "  '/ó������
    gstrQuy = gstrQuy & vbCrLf & "  AND PRSC_NO         =  " & Val(fstró��SEQ) & " "                                     '/ó���ȣ
    gstrQuy = gstrQuy & vbCrLf & "  AND PID             = '" & fstr���Ϲ�ȣ & "' "                                   '/���Ϲ�ȣ
    gstrQuy = gstrQuy & vbCrLf & "  AND PRSC_CD         = '" & fstró���ڵ� & "' "                                   '/ó���ڵ�
    gstrQuy = gstrQuy & vbCrLf & "  AND PRSC_VALD_YN    = 'Y' "                                                     '/��ó�� ����ִ� ó��
    gstrQuy = gstrQuy & vbCrLf & "  AND PRSC_HSTR_CD    = 'O' "                                                     '/ó��History ��ȣ
    If RunSQL(gstrQuy) = False Then ADC.RollbackTrans: Call CloseDB: End

    '/�ǽ�ó��: TMPRSCEXCN
    gstrQuy = "UPDATE TMPRSCEXCN SET "
    gstrQuy = gstrQuy & vbCrLf & "      CNDT_PRSC_STAT_CD   = '560', "                      '/560.�ӽð�� �̻�
    gstrQuy = gstrQuy & vbCrLf & "      UPDTR_ID            = '" & gtypUSER.USERID & "', "  '/���������� ID
    gstrQuy = gstrQuy & vbCrLf & "      UPDT_DT             = SYSTIMESTAMP "                '/������������
    gstrQuy = gstrQuy & vbCrLf & "WHERE PRSC_DATE           = '" & Format(CDate(fstró������), "YYYYMMDD") & "' "  '/ó������
    gstrQuy = gstrQuy & vbCrLf & "  AND PRSC_NO             =  " & Val(fstró��SEQ) & " "                                     '/ó���ȣ
    gstrQuy = gstrQuy & vbCrLf & "  AND PID                 = '" & fstr���Ϲ�ȣ & "' "                                   '/���Ϲ�ȣ
    gstrQuy = gstrQuy & vbCrLf & "  AND MEFE_CD             = '" & fstró���ڵ� & "' "                                   '/ó���ڵ�
    gstrQuy = gstrQuy & vbCrLf & "  AND CNDT_PRSC_VALD_YN   = 'Y' "                                                     '/�ǽ�ó�� ����ִ� ó��
    If RunSQL(gstrQuy) = False Then ADC.RollbackTrans: Call CloseDB: End
    

    '/----------------------------------------------------------------------------------------------------/
    '/Step11.   ��� ���� ���� �� MM_EMR_RES(Image ��� ����)�� Insert
    '/----------------------------------------------------------------------------------------------------/
    gstrQuy = "SELECT * "
    gstrQuy = gstrQuy & vbCrLf & "  FROM MM_EMR_RES "
    gstrQuy = gstrQuy & vbCrLf & " WHERE PATNO     = '" & fstr���Ϲ�ȣ & "' "
    gstrQuy = gstrQuy & vbCrLf & "   AND ORDDATE   = '" & Replace(fstró������, "-", "") & "' "
    gstrQuy = gstrQuy & vbCrLf & "   AND ORDSEQ    =  " & Val(fstró��SEQ) & " "
    gstrQuy = gstrQuy & vbCrLf & "   AND EQUIPCODE = '" & gtypEQ_INFO.EQUIPCODE & "' "
    gstrQuy = gstrQuy & vbCrLf & "   AND EQUIPSEQ  =  " & gtypEQ_INFO.EQUIPSEQ & " "
    If ReadSQL(gstrQuy, ADR) = False Then Call CloseDB: End
                
    If Not ADR Is Nothing Then
        ADR.Close: Set ADR = Nothing
        
        '/Server DB�� ����� �Է��� �Ǿ� ������ �˻����ڸ� Update ��.
        gstrQuy = "UPDATE MM_EMR_RES SET "
        gstrQuy = gstrQuy & vbCrLf & "       EXAMDATE  = TO_CHAR(TRUNC(SYSDATE),'YYYYMMDD') "
        gstrQuy = gstrQuy & vbCrLf & " WHERE PATNO     = '" & fstr���Ϲ�ȣ & "' "
        gstrQuy = gstrQuy & vbCrLf & "   AND ORDDATE   = '" & Replace(fstró������, "-", "") & "' "
        gstrQuy = gstrQuy & vbCrLf & "   AND ORDSEQ    =  " & Val(fstró��SEQ) & " "
        gstrQuy = gstrQuy & vbCrLf & "   AND EQUIPCODE = '" & gtypEQ_INFO.EQUIPCODE & "' "
        gstrQuy = gstrQuy & vbCrLf & "   AND EQUIPSEQ  =  " & gtypEQ_INFO.EQUIPSEQ & " "
        If RunSQL(gstrQuy) = False Then ADC.RollbackTrans: Call CloseDB: End
    Else
        '/����ڵ庰 ó���ڵ� ��������
        gstrQuy = "INSERT INTO MM_EMR_RES "
        gstrQuy = gstrQuy & vbCrLf & " (PATNO,      ORDDATE,    ORDSEQ,     EXAMDATE,       DEPTCODE, "
        gstrQuy = gstrQuy & vbCrLf & "  PARTCODE,   EQUIPCODE,  EXAMCODE,   WORDNO,         ROOMNO, "
        gstrQuy = gstrQuy & vbCrLf & "  IOFLAG,     EXECID,     DRID,       IMGFILENAME,    IMGFILEPATH, "
        gstrQuy = gstrQuy & vbCrLf & "  RECEDATE,   RECESEQ,    EQUIPSEQ) "
        gstrQuy = gstrQuy & vbCrLf & " VALUES "
        gstrQuy = gstrQuy & vbCrLf & " ('" & fstr���Ϲ�ȣ & "', "                    '/PATNO(���Ϲ�ȣ)
        gstrQuy = gstrQuy & vbCrLf & "  '" & Replace(fstró������, "-", "") & "', "  '/ORDDATE(ó������)
        gstrQuy = gstrQuy & vbCrLf & "   " & Val(fstró��SEQ) & ", "                 '/ORDSEQ(ó��SEQ(�ǰ������� ��� ������ȣ))
        gstrQuy = gstrQuy & vbCrLf & "  TO_CHAR(TRUNC(SYSDATE),'YYYYMMDD'), "       '/EXAMDATE(����Է�����)
        gstrQuy = gstrQuy & vbCrLf & "  '" & fstr����� & "', "                      '/DEPTCODE(������ڵ�)
        gstrQuy = gstrQuy & vbCrLf & "  '', "                                       '/PARTCODE(������ڵ�)
        gstrQuy = gstrQuy & vbCrLf & "  '" & gtypEQ_INFO.EQUIPCODE & "', "          '/EQUIPCODE(����ڵ�)
        gstrQuy = gstrQuy & vbCrLf & "  '" & fstró���ڵ� & "', "                    '/EXAMCODE(�˻��ڵ�)
        gstrQuy = gstrQuy & vbCrLf & "  '', "                                       '/WORDNO(�����ڵ�)
        gstrQuy = gstrQuy & vbCrLf & "  '', "                                       '/ROOMNO(�����ڵ�)
        gstrQuy = gstrQuy & vbCrLf & "  '" & fstr�Կܱ��� & "', "                    '/IOFLAG(�Կ�/�ܷ�/���� ����)
        gstrQuy = gstrQuy & vbCrLf & "  '" & gtypUSER.USERID & "', "                '/EXECID(������ȣ)
        gstrQuy = gstrQuy & vbCrLf & "  '', "                                       '/DRID(ó���ǹ�ȣ)
        gstrQuy = gstrQuy & vbCrLf & "  '" & strFileName & "', "                    '/IMGFILENAME(����̹������ϸ�)
        gstrQuy = gstrQuy & vbCrLf & "  '" & strIMGFILEPATH & "', "                 '/IMGFILEPATH(����̹������ϰ��)
        gstrQuy = gstrQuy & vbCrLf & "  '', "                                       '/RECEDATE(��������)
        gstrQuy = gstrQuy & vbCrLf & "  '', "                                       '/RECESEQ(����SEQ)
        gstrQuy = gstrQuy & vbCrLf & "  '" & gtypEQ_INFO.EQUIPSEQ & "') "           '/EQUIPSEQ(���SEQ)
        If RunSQL(gstrQuy) = False Then ADC.RollbackTrans: Call CloseDB: End
    End If
    
    ADC.CommitTrans
    
    Call CloseDB

    FUNC_MM_SAVE = True

Exit Function

'/----------------------------------------------------------------------------------------------------/

RTN_ERR:
    MsgBox "���� �� ������ �߻��Ͽ����ϴ�." & vbCrLf & _
           "����� Ȥ�� ���޾�ü�� �����ֽñ� �ٶ��ϴ�.", vbCritical, "���ۿ���"

End Function

Public Function FUNC_MM_VIEW(ArgSection As Integer) As Boolean
    Dim stró���ڵ�     As String
    
    FUNC_MM_VIEW = False
    
    Select Case ArgSection
        Case 1: GoSub VIEW1_RTN '/�����ڳ���(������)
        Case 2: GoSub VIEW2_RTN '/�����ڳ���(����)
        Case 3: GoSub VIEW3_RTN '/������ �ڷ�ҷ�����
        Case 4: GoSub VIEW4_RTN '/���ۿϷ� �ڷ�ҷ�����
    End Select
    
    FUNC_MM_VIEW = True
Exit Function

'/----------------------------------------------------------------------------------------------------/

VIEW1_RTN: '/�����ڳ���(������)
    If OpenDB(gstrREG_DB_CONSTR) = True Then
        '/����ڵ庰 ó���ڵ� ��������
        gstrQuy = "SELECT ORDCD "
        gstrQuy = gstrQuy & vbCrLf & "  FROM MM_EMR_EQORD "
        gstrQuy = gstrQuy & vbCrLf & " WHERE EQUIPCODE = '" & gtypEQ_INFO.EQUIPCODE & "' "
        If ReadSQL(gstrQuy, ADR) = False Then Call CloseDB: End
        
        If Not ADR Is Nothing Then
            Do Until ADR.EOF
                stró���ڵ� = stró���ڵ� & ",'" & Trim(ADR!ORDCD & "") & "'"
                    
                ADR.MoveNext
            Loop
            ADR.Close: Set ADR = Nothing
                        
            stró���ڵ� = Mid(stró���ڵ�, 2)
        End If
        
        If FUNC_HIS_ORDER1_VIEW(stró���ڵ�) = True Then '/������ ������ ó����ȸQuery
            If ReadSQL(gstrQuy, ADR) = True Then
                If Not ADR Is Nothing Then
                    Do Until ADR.EOF
                        With spr�����ڳ���
                            .MaxRows = .MaxRows + 1: .Row = .MaxRows
                            
                            .Col = 2:  .Text = Trim(ADR!CHRTNO & "")                            '/���Ϲ�ȣ
                            .Col = 3:  .Text = Trim(ADR!IO_SECTION & "")                        '/ȯ�ڱ���
                            .Col = 4:  .Text = Trim(ADR!DETPCD & "")                            '/�����
                            .Col = 5:  .Text = Trim(ADR!PATNM & "")                             '/�����ڸ�
                            .Col = 6:  .Text = Trim(ADR!SEX & "") & "/" & Trim(ADR!AGE & "")    '/Seq/Age
                            .Col = 7:  .Text = Format(Trim(ADR!ORDDATE & ""), "@@@@-@@-@@")     '/ó������
                            .Col = 8:  .Text = Trim(ADR!ORDSEQ & "")                            '/ó��SEQ
                            .Col = 9:  .Text = ""                                               '/�������(�����ڷḸ)
                            .Col = 10: .Text = Trim(ADR!ORDNM & "")                             '/ó���
                            .Col = 11: .Text = Trim(ADR!ORDCD & "")                             '/ó���ڵ�
                            .Col = 12: .Text = Trim(ADR!CNDT_PRSC_STAT_CD & "")                 '/�ǽ�ó�� ó���������Flag
                        End With
                        
                        ADR.MoveNext
                    Loop
                    ADR.Close: Set ADR = Nothing
                    
                End If
            End If
        End If
        
        Call CloseDB
    End If
Return

'/----------------------------------------------------------------------------------------------------/

VIEW2_RTN: '/�����ڳ���(����)
    If OpenDB(gstrREG_DB_CONSTR) = True Then
        '/����ڵ庰 ó���ڵ� ��������
        gstrQuy = "SELECT ORDCD "
        gstrQuy = gstrQuy & vbCrLf & "  FROM MM_EMR_EQORD "
        gstrQuy = gstrQuy & vbCrLf & " WHERE EQUIPCODE = '" & gtypEQ_INFO.EQUIPCODE & "' "
        If ReadSQL(gstrQuy, ADR) = False Then Call CloseDB: End
        
        If Not ADR Is Nothing Then
            Do Until ADR.EOF
                stró���ڵ� = stró���ڵ� & ",'" & Trim(ADR!ORDCD & "") & "'"
                    
                ADR.MoveNext
            Loop
            ADR.Close: Set ADR = Nothing
                        
            stró���ڵ� = Mid(stró���ڵ�, 2)
        End If
        
        If FUNC_HIS_ORDER2_VIEW(stró���ڵ�) = True Then '/������ ������ ó����ȸQuery
            If ReadSQL(gstrQuy, ADR) = True Then
                If Not ADR Is Nothing Then
                    Do Until ADR.EOF
                        With spr�����ڳ���
                            .MaxRows = .MaxRows + 1: .Row = .MaxRows
                            
                            .Col = 2:  .Text = Trim(ADR!CHRTNO & "")                            '/���Ϲ�ȣ
                            .Col = 3:  .Text = Trim(ADR!IO_SECTION & "")                        '/ȯ�ڱ���
                            .Col = 4:  .Text = Trim(ADR!DETPCD & "")                            '/�����
                            .Col = 5:  .Text = Trim(ADR!PATNM & "")                             '/�����ڸ�
                            .Col = 6:  .Text = Trim(ADR!SEX & "") & "/" & Trim(ADR!AGE & "")    '/Seq/Age
                            .Col = 7:  .Text = Format(Trim(ADR!ORDDATE & ""), "@@@@-@@-@@")     '/ó������
                            .Col = 8:  .Text = Trim(ADR!ORDSEQ & "")                            '/ó��SEQ
                            .Col = 9:  .Text = Format(Trim(ADR!EXAMDATE & ""), "@@@@-@@-@@")    '/�������(�����ڷḸ)
                            .Col = 10: .Text = Trim(ADR!ORDNM & "")                             '/ó���
                            .Col = 11: .Text = Trim(ADR!ORDCD & "")                             '/ó���ڵ�
                            .Col = 12: .Text = Trim(ADR!CNDT_PRSC_STAT_CD & "")                 '/�ǽ�ó�� ó���������Flag
                        End With
                    
                        ADR.MoveNext
                    Loop
                    ADR.Close: Set ADR = Nothing
                End If
            End If
        End If
        
        Call CloseDB
    End If
Return

'/----------------------------------------------------------------------------------------------------/

VIEW3_RTN:
    Dim MyPath, MyName
    Dim sFileCnt As Integer
    
    If spr������.MaxRows > 0 Then spr������.MaxRows = 0
    
    MyPath = gtypEQ_INFO.EQIMGFILEPATH & "\"   ' ��θ� �����մϴ�.
    MyName = Dir(MyPath, vbDirectory)   ' ù��° �׸��� �˻��մϴ�.
    
'''    '/TIF or TIFF ���� ������ JPG ��ȯ �� ����
'''    Do While MyName <> ""   ' ����(loop)�� �����մϴ�.
'''        ' ���� ���丮�� �����ϴ� ���丮�� �����մϴ�.
'''        If MyName <> "." And MyName <> ".." Then
'''            If InStr(UCase(MyName), ".TIF") > 0 Or InStr(UCase(MyName), ".TIFF") > 0 Then
'''
'''                staCondition.Panels.Item(2).Text = "TIF or TIFF ������ JPG �� ��ȯ ���Դϴ�..."
'''
'''                sFileCnt = FUNC_TifToJpg(gtypEQ_INFO.EQIMGFILEPATH, CStr(MyName))
'''
'''                If sFileCnt = 0 Then
'''                    MsgBox "TIF Ȥ�� TIFF ������ JPG�� ��ȯ �߿� ������ �߻��Ͽ����ϴ�." & vbCrLf & _
'''                           "����� Ȥ�� ���޾�ü�� �����ֽñ� �ٶ��ϴ�.", vbCritical, "�ڷ�ҷ����� ����"
'''                End If
'''            End If
'''        End If
'''        MyName = Dir   ' ���� �׸��� �о���Դϴ�.
'''    Loop
'''    staCondition.Panels.Item(2).Text = ""
    
    MyName = Dir(MyPath, vbDirectory)   ' ù��° �׸��� �˻��մϴ�.
    Do While MyName <> ""   ' ����(loop)�� �����մϴ�.
       ' ���� ���丮�� �����ϴ� ���丮�� �����մϴ�.
        If MyName <> "." And MyName <> ".." Then
            If InStr(UCase(MyName), ".JPG") > 0 Or InStr(UCase(MyName), ".JPEG") > 0 Or _
               InStr(UCase(MyName), ".TIF") > 0 Or InStr(UCase(MyName), ".TIFF") > 0 Or _
               InStr(UCase(MyName), ".PDF") > 0 Then
                spr������.MaxRows = spr������.MaxRows + 1
                
                Call SET_CELL(spr������, 2, spr������.MaxRows, MyName) '/ȭ�ϸ�(�̸����氡��)
                Call SET_CELL(spr������, 3, spr������.MaxRows, FileDateTime(MyPath & MyName)) '/
            End If
        End If
        MyName = Dir   ' ���� �׸��� �о���Դϴ�.
    Loop
Return

'/----------------------------------------------------------------------------------------------------/

VIEW4_RTN:
    '/Step1.���ۿϷ� �ڷ� ���� ���� �� ���� Clear
    '/Step2.���ۿϷ� DB �ڷ� ��ȸ
    '/Step3.FTP �����ڷ� ��������
    Dim strIMGFILEPATH  As String
    
    If OpenDB(gstrREG_DB_CONSTR) = True Then
        '/----------------------------------------------------------------------------------------------------/
        '/Step1.���ۿϷ� �ڷ� ���� ���� �� ���� Clear
        '/----------------------------------------------------------------------------------------------------/
        If Dir(gtypEQ_INFO.FTPIMGFILEPATH, vbDirectory) = "" Then MkDir gtypEQ_INFO.FTPIMGFILEPATH
    
        If Len(Dir(gtypEQ_INFO.FTPIMGFILEPATH & "\*.*")) > 0 Then
            Kill gtypEQ_INFO.FTPIMGFILEPATH & "\*.*"
        End If
        
        '/----------------------------------------------------------------------------------------------------/
        '/Step2.���ۿϷ� DB �ڷ� ��ȸ
        '/----------------------------------------------------------------------------------------------------/
        gstrQuy = "SELECT * "
        gstrQuy = gstrQuy & vbCrLf & "  FROM MM_EMR_RES "
        gstrQuy = gstrQuy & vbCrLf & " WHERE PATNO     = '" & lbl���Ϲ�ȣ & "' "
        gstrQuy = gstrQuy & vbCrLf & "   AND ORDDATE   = '" & Replace(lbló������, "-", "") & "' "
        gstrQuy = gstrQuy & vbCrLf & "   AND ORDSEQ    =  " & Val(lbló��SEQ) & " "
        gstrQuy = gstrQuy & vbCrLf & "   AND EQUIPCODE = '" & gtypEQ_INFO.EQUIPCODE & "' "
        gstrQuy = gstrQuy & vbCrLf & "   AND EQUIPSEQ  =  " & gtypEQ_INFO.EQUIPSEQ & " "
        If ReadSQL(gstrQuy, ADR) = False Then Call CloseDB: End
        
        If Not ADR Is Nothing Then
            strIMGFILEPATH = Trim(ADR!IMGFILEPATH & "")                              '/FTP���
            
            ADR.Close: Set ADR = Nothing
        End If
            
        Call CloseDB
        
        '/----------------------------------------------------------------------------------------------------/
        '/Step3.FTP �����ڷ� ��������
        '/----------------------------------------------------------------------------------------------------/
        '/FTP ���� �õ�
        Dim success As Long
        success = sftp.IsConnected
        If (success <> 1) Then
            If MMSFTP.OpenConnection(gstrFTP_RH, gstrFTP_RP, gstrFTP_UN, gstrFTP_PW) = False Then
                MsgBox "Image File Server�� ������ �� �����ϴ�." & vbCrLf & "����ǿ� ���ǹٶ��ϴ�.", vbCritical, "FTP ���� ����"
                Exit Function
            End If
        End If
        
        '/���� ���� ����
        If MMSFTP.SetFTPDirectory(strIMGFILEPATH) = False Then
            '''MsgBox "�ش� �ڷᰡ �����ϴ�.", vbInformation, "Ȯ��"
            Exit Function
        End If
        
        '/����SEQ ���� ���� ã��
        Call MMSFTP.FtpScanDirectory(strIMGFILEPATH)
        
        With spr���ۿϷ�
            If UBound(FtpScanFileName_IMG) > 0 Then
                For intX = 1 To UBound(FtpScanFileName_IMG)
                    .MaxRows = .MaxRows + 1: .Row = .MaxRows
                    
                    .Col = 2: .Text = FtpScanFileName_IMG(intX)
                    '''.Col = 3: .Text = FtpScanFileDate(intX) '/��¥ ����� �ʿ信 ���� ��Ÿ����(������ڰ� �ִ� ����� ���� ������ �ʿ����)
                    
                    If .MaxTextRowHeight(.Row) > 13.3 Then .RowHeight(.Row) = .MaxTextRowHeight(.Row)
                    
                    If MMSFTP.FTPDownloadFile(gtypEQ_INFO.FTPIMGFILEPATH & "\" & GET_CELL(spr���ۿϷ�, 2, intX), strIMGFILEPATH & GET_CELL(spr���ۿϷ�, 2, intX)) = True Then FUNC_MM_VIEW = True
                    
                    '/Call sftp.DownloadFileByName(strIMGFILEPATH & GET_CELL(spr���ۿϷ�, 2, intX), gtypEQ_INFO.FTPIMGFILEPATH & "\" & GET_CELL(spr���ۿϷ�, 2, intX))
                    
                    
                    FUNC_MM_VIEW = True
                Next intX
            End If
        End With
        
        '/FTP ���� ����
        Call MMSFTP.CloseConnection
    End If
Return
End Function

Public Function FUNC_TifToJpg(ArgFilePath As String, ArgFileName As String) As Integer
'''    'tif ������ jpg �Ŀ÷� ��ȯ�ϴ� �Լ�
'''    Dim strFileExt        As String
'''    Dim strFilePath       As String
'''    Dim strFileName       As String
'''    Dim strFileNewName    As String
'''    Dim xPoint          As Integer
'''    Dim iImgCnt         As Integer
'''    Dim jCreateImgCnt   As Integer
'''
'''    FUNC_TifToJpg = 0
'''
'''On Error GoTo ERR_RTN
'''
'''    strFilePath = ArgFilePath '���ϰ��
'''    strFileName = ArgFileName '���ϸ�
'''    xPoint = InStr(1, strFileName, ".")
'''    strFileExt = Trim(Mid(strFileName, xPoint + 1)) 'Ȯ���ڸ�
'''    strFileNewName = Trim(Mid(strFileName, 1, xPoint - 1)) 'Ȯ���� ������ ���ϸ�
'''
'''    imvTif.LoadMultiPage strFilePath & "\" & strFileName, 0 '���Ϸε�
'''    iImgCnt = imvTif.GetTotalPage '�ε�� Tiff ������ Page ��
'''
'''    If LCase(strFileExt) = "tiff" Or LCase(strFileExt) = "tif" Then
'''        For jCreateImgCnt = 1 To iImgCnt '������ ����ŭ jpg ���� ����
'''            imvTif.ExportTIF strFilePath & "\" & strFileName, strFilePath & "\" & strFileNewName & "-" & Format(jCreateImgCnt, "00"), "JPG", jCreateImgCnt, 1
'''        Next jCreateImgCnt
'''        imvTif.Filename = ""
'''        '' jpg���� ������ tiff ���� ����
'''        Kill strFilePath & "\" & strFileName
'''    End If
'''    '������ jpg ���� ���� ��ȯ
'''    FUNC_TifToJpg = jCreateImgCnt
'''
'''ERR_RTN:
'''
End Function

Private Sub cboZoomValue_Click()
    Select Case cboZoomValue
        Case "25%": imvResult.View = 1
        Case "33%": imvResult.View = 2
        Case "50%": imvResult.View = 3
        Case "75%": imvResult.View = 4
        Case "100%": imvResult.View = 5
        Case "150%": imvResult.View = 6
        Case "200%": imvResult.View = 7
        Case Else: imvResult.View = 9
    End Select
    
    '/���Ϻ��� �������� �̹��� OCX���� ZOOM ����� ����.
'''    imvResult.Zoom Val(Replace(cboZoomValue, "%", "")), Val(Replace(cboZoomValue, "%", ""))
End Sub

Private Sub cboZoomValue_KeyDown(KeyCode As Integer, Shift As Integer)
    '/���Ϻ��� �������� �̹��� OCX���� ZOOM ����� ����.
'''    If KeyCode = vbKeyReturn Then
'''        strTemp = Replace(cboZoomValue, "%", "")
'''
'''        If IsNumeric(strTemp) = False Then
'''            MsgBox "1���� 200������ ���ڸ� (��)�Է��Ͻʽÿ�!", vbCritical, "���������Ұ�"
'''            cboZoomValue.SetFocus
'''            Exit Sub
'''        End If
'''
'''        If Not (Val(strTemp) >= 1 And Val(strTemp) <= 200) Then
'''            MsgBox "1���� 200������ ���ڸ� (��)�Է��Ͻʽÿ�!", vbCritical, "���������Ұ�"
'''            cboZoomValue.SetFocus
'''            Exit Sub
'''        End If
'''
'''        imvResult.Zoom Val(strTemp), Val(strTemp)
'''    End If
End Sub

Private Sub cmd��������������_Click()
    Dim Message, Title, Default, MyValue
    Dim OldName, NewName
    
    Message = "������ ������������θ� �Է��Ͻʽÿ�"   ' ������Ʈ ����.
    Title = "������ ���� ����"   ' ���� ����.
    ' �޽��� ȭ�� ǥ��, ����, �⺻��.
    MyValue = InputBox(Message, Title, gtypEQ_INFO.EQIMGFILEPATH)
    
    If Trim(MyValue) <> "" Then
        '/�������� ���� Ȯ��
        If Dir(MyValue, vbDirectory) = "" Then '/���������� ������...
            MsgBox "����� ������ ������ �������� �ʽ��ϴ�." & vbCrLf & vbCrLf & _
                   "������ �Ϸ��� [��������������]�� (��)�����Ͻʽÿ�.", vbCritical, "�������"
            Exit Sub
        Else
            '/2.��������
            gtypEQ_INFO.EQIMGFILEPATH = MyValue
        End If
        
        '/3.DB ����
        If OpenDB(gstrREG_DB_CONSTR) = False Then End
        
        gstrQuy = "SELECT * "
        gstrQuy = gstrQuy & vbCrLf & "  FROM MM_EMR_CONF "
        gstrQuy = gstrQuy & vbCrLf & " WHERE EQUIPCODE = '" & gtypEQ_INFO.EQUIPCODE & "' "
        gstrQuy = gstrQuy & vbCrLf & "   AND EQUIPSEQ  =  " & gtypEQ_INFO.EQUIPSEQ & " "
        If ReadSQL(gstrQuy, ADR) = False Then Call CloseDB: End
        
        If Not ADR Is Nothing Then
            ADR.Close: Set ADR = Nothing
            
            ADC.BeginTrans
    
            gstrQuy = "UPDATE MM_EMR_CONF SET "
            gstrQuy = gstrQuy & vbCrLf & "       EQIMGFILEPATH  = '" & gtypEQ_INFO.EQIMGFILEPATH & "' " '/Local �̹��� ���
            gstrQuy = gstrQuy & vbCrLf & " WHERE EQUIPCODE      = '" & gtypEQ_INFO.EQUIPCODE & "' "
            gstrQuy = gstrQuy & vbCrLf & "   AND EQUIPSEQ       =  " & gtypEQ_INFO.EQUIPSEQ & " "
            If RunSQL(gstrQuy) = False Then ADC.RollbackTrans: Call CloseDB: End
            
            ADC.CommitTrans
        End If
        
        Call CloseDB
    End If
End Sub

Private Sub cmd���ۿϷ���������_Click()
    Dim Message, Title, Default, MyValue
    Dim OldName, NewName
    
    Message = "������ ���ۿϷ�������θ� �Է��Ͻʽÿ�"   ' ������Ʈ ����.
    Title = "���ۿϷ� ���� ����"   ' ���� ����.
    ' �޽��� ȭ�� ǥ��, ����, �⺻��.
    MyValue = InputBox(Message, Title, gtypEQ_INFO.FTPIMGFILEPATH)
    
    If Trim(MyValue) <> "" Then
        '/�������� ���� Ȯ��
        If Dir(MyValue, vbDirectory) = "" Then '/���������� ������...
            MsgBox "����� ���ۿϷ� ������ �������� �ʽ��ϴ�." & vbCrLf & vbCrLf & _
                   "������ �Ϸ��� [���ۿϷ���������]�� (��)�����Ͻʽÿ�.", vbCritical, "�������"
            Exit Sub
        Else
            '/2.��������
            gtypEQ_INFO.FTPIMGFILEPATH = MyValue
        End If
    
        '/3.DB ����
        If OpenDB(gstrREG_DB_CONSTR) = False Then End
        
        gstrQuy = "SELECT * "
        gstrQuy = gstrQuy & vbCrLf & "  FROM MM_EMR_CONF "
        gstrQuy = gstrQuy & vbCrLf & " WHERE EQUIPCODE = '" & gtypEQ_INFO.EQUIPCODE & "' "
        gstrQuy = gstrQuy & vbCrLf & "   AND EQUIPSEQ  =  " & gtypEQ_INFO.EQUIPSEQ & " "
        If ReadSQL(gstrQuy, ADR) = False Then Call CloseDB: End
        
        If Not ADR Is Nothing Then
            ADR.Close: Set ADR = Nothing
            
            ADC.BeginTrans
    
            gstrQuy = "UPDATE MM_EMR_CONF SET "
            gstrQuy = gstrQuy & vbCrLf & "       FTPIMGFILEPATH = '" & gtypEQ_INFO.FTPIMGFILEPATH & "' " '/Local �̹��� ���
            gstrQuy = gstrQuy & vbCrLf & " WHERE EQUIPCODE      = '" & gtypEQ_INFO.EQUIPCODE & "' "
            gstrQuy = gstrQuy & vbCrLf & "   AND EQUIPSEQ       =  " & gtypEQ_INFO.EQUIPSEQ & " "
            If RunSQL(gstrQuy) = False Then ADC.RollbackTrans: Call CloseDB: End
            
            ADC.CommitTrans
        End If
        
        Call CloseDB
    End If
End Sub

Private Sub cmdDeleteFTP_Click()
    If Trim(gtypEQ_INFO.FTPIMGFILEPATH) = "" Then MsgBox "���ۿϷ� ���� ���� ���� ������ �����ϴ�.", vbCritical, "��ȸ�Ұ�": Exit Sub
    
    strTemp = "N"
    For intX = 1 To spr���ۿϷ�.MaxRows
        If GET_CELL(spr���ۿϷ�, 1, intX) = "1" Then strTemp = "Y": Exit For
    Next intX
    If strTemp <> "Y" Then MsgBox "������ ������ �����Ͻʽÿ�!", vbCritical, "�����Ұ�": Exit Sub
    
    If GET_CELL(spr�����ڳ���, 12, spr�����ڳ���.ActiveRow) >= "610" Then
        MsgBox "�ش� ó���� �ǽ� �Ϸ�� ó���Դϴ�. ������ �� �����ϴ�.!" & vbCrLf & vbCrLf & _
               "�ǽ� ���� �� ���� �ٶ��ϴ�.", vbCritical, "�����Ұ�": Exit Sub
    End If
    
    If MsgBox("������ ������ ��� ó���� �Ϸ�� �����̸�, " & vbCrLf & _
              "���� �� ������ �Ұ����Ͽ��� �����Ͻñ� �ٶ��ϴ�." & vbCrLf & vbCrLf & _
              "����ؼ� ���� ó���� �����ϰڽ��ϱ�?������ ���ۿϷ� ������ �����ϰڽ��ϱ�?", vbQuestion + vbYesNo, "��������") = vbNo Then Exit Sub
    
    Call FUNC_MM_KEY_CLEAR("3") '/Image Clear
    If FUNC_MM_DELETE("2") = True Then Call cmdView_Click 'Call cmdViewFTP_Click
End Sub

Private Sub cmdDeleteLocal_Click()
    If Trim(gtypEQ_INFO.EQIMGFILEPATH) = "" Then MsgBox "������ ���� ���� ���� ������ �����ϴ�.", vbCritical, "�����Ұ�": Exit Sub
    
    strTemp = "N"
    For intX = 1 To spr������.MaxRows
        If GET_CELL(spr������, 1, intX) = "1" Then strTemp = "Y": Exit For
    Next intX
    If strTemp <> "Y" Then MsgBox "������ ������ �����Ͻʽÿ�!", vbCritical, "�����Ұ�": Exit Sub
    
    If MsgBox("������ ������ ������ �����ϰڽ��ϱ�?", vbQuestion + vbOKCancel, "��������") = vbCancel Then Exit Sub
    
    Call FUNC_MM_KEY_CLEAR("3") '/Image Clear
    If FUNC_MM_DELETE("1") = True Then Call cmdViewLocal_Click
End Sub

Private Sub cmdMultiFirst_Click()
    txtMultiPno = 1
    Call cmdMultiJump_Click
End Sub

Private Sub cmdMultiJump_Click()
    imvResult.LoadMultiPage imvResult.Filename, Val(txtMultiPno)
End Sub

Private Sub cmdMultiLast_Click()
    txtMultiPno = imvResult.GetTotalPage
    Call cmdMultiJump_Click
End Sub

Private Sub cmdMultiNext_Click()
    If txtMultiPno < imvResult.GetTotalPage Then
        txtMultiPno = txtMultiPno + 1
    Else
        txtMultiPno = imvResult.GetTotalPage
    End If
    Call cmdMultiJump_Click
End Sub

Private Sub cmdMultiPrev_Click()
    If txtMultiPno > 1 Then
        txtMultiPno = txtMultiPno - 1
    Else
        txtMultiPno = 1
    End If
    Call cmdMultiJump_Click
End Sub

Private Sub cmdRotate_Click()
    imvResult.Rotate90
    imvResult.SetFocus
End Sub

Private Sub cmdSave_Click()
    Dim intRow�����ڳ���    As Integer
    
    '/----------------------------------------------------------------------------------------------------/
    '/## [Image List�� ���ۿϷ� Tab�� ��� ���۹�ư�� Enabled = False �� �ȴ�.
    '/Step1.    �����ڳ����� �ϳ� �̻��� ȯ�ڰ� ���õǾ� �ִ��� Ȯ��
    '/Step2.    �����ڳ����� �ϳ� �̻��� ȯ�ڰ� ���õǾ� ������ �ش� ȯ���� ���Ϲ�ȣ�� �ٸ��� Ȯ��
    '/Step3.    [Image List]�� ������ Tab�� �ϳ� �̻��� �ڷᰡ ���õǾ����� Ȯ��
    '/Step4.    FTP ����
    '/Step5.    FTP ������ �ش� ���� ����/ ������ SKIP
    '/Step6.    ������ Ȯ���� ������ ���ϸ� ����
    '/Step7.    FTP ������ �ش� ������ �ڷᰡ ������ �ִ� Ȯ���� �� ã��/�ڷᰡ ������ 0
    '/Step8.    ���õ� ������ �ڷ� Rename �ϸ鼭 ����
    '/Step9.    FTP ����
    '/Step10.   ��� ���� ���� �� HIS�� �˻�Ϸ� Falg Update ����.
    '/Step11.   ��� ���� ���� �� MM_EMR_RES(Image ��� ����)�� Insert
    '/Step12.   ��� ���� ���� �� ���� �ڷḦ Local�������� ���� �� �����ڳ��� ��ȸ �� ������ �ڷ�ҷ����� ����
    '/----------------------------------------------------------------------------------------------------/
    
    '/----------------------------------------------------------------------------------------------------/
    '/Step1.    �����ڳ����� �ϳ� �̻��� ȯ�ڰ� ���õǾ� �ִ��� Ȯ��
    '/----------------------------------------------------------------------------------------------------/
    strTemp = "N"
    For intX = 1 To spr�����ڳ���.MaxRows
        If GET_CELL(spr�����ڳ���, 1, intX) = "1" Then strTemp = "Y": Exit For
    Next intX
    If strTemp <> "Y" Then MsgBox "������ �����ڸ� �����Ͻʽÿ�!", vbCritical, "���ۺҰ�": Exit Sub

    '/----------------------------------------------------------------------------------------------------/
    '/Step2.    �����ڳ����� �ϳ� �̻��� ȯ�ڰ� ���õǾ� ������ �ش� ȯ���� ���Ϲ�ȣ�� �ٸ��� Ȯ��
    '/----------------------------------------------------------------------------------------------------/
    strTemp = ""
    For intX = 1 To spr�����ڳ���.MaxRows
        If GET_CELL(spr�����ڳ���, 1, intX) = "1" Then
            If strTemp <> GET_CELL(spr�����ڳ���, 2, intX) Then
                If strTemp = "" Then
                    strTemp = GET_CELL(spr�����ڳ���, 2, intX)
                Else
                    MsgBox "2�� �̻� ������, ������ �������� ���Ϲ�ȣ�� ���� �ٸ��ϴ�.", vbCritical, "���ۺҰ�": Exit Sub
                End If
            End If
        End If
    Next intX

    '/----------------------------------------------------------------------------------------------------/
    '/Step3.    [Image List]�� ������ Tab�� �ϳ� �̻��� �ڷᰡ ���õǾ����� Ȯ��
    '/----------------------------------------------------------------------------------------------------/
    strTemp = "N"
    For intX = 1 To spr������.MaxRows
        If GET_CELL(spr������, 1, intX) = "1" Then strTemp = "Y": Exit For
    Next intX
    If strTemp <> "Y" Then MsgBox "������ Image�� �����Ͻʽÿ�!", vbCritical, "���ۺҰ�": Exit Sub

    On Error GoTo ERN_ERR
    
    Screen.MousePointer = 11
    
    For intRow�����ڳ��� = 1 To spr�����ڳ���.MaxRows
        If GET_CELL(spr�����ڳ���, 1, intRow�����ڳ���) = "1" Then
            If FUNC_MM_SAVE(intRow�����ڳ���) = False Then GoTo ERN_ERR
        End If
    Next intRow�����ڳ���
    
    '/----------------------------------------------------------------------------------------------------/
    '/Step12.   ��� ���� ���� �� ���� �ڷḦ Local�������� ���� �� �����ڳ��� ��ȸ �� ������ �ڷ�ҷ����� ����
    '/----------------------------------------------------------------------------------------------------/
    imvResult.Filename = ""
    Call SUB_CHK_PDF_TIF("")
    For intX = 1 To spr������.MaxRows
        If GET_CELL(spr������, 1, intX) = "1" Then
            Kill gtypEQ_INFO.EQIMGFILEPATH & "\" & GET_CELL(spr������, 2, intX)
        End If
    Next intX
    Call cmdView_Click
    Call cmdViewLocal_Click
    
    Screen.MousePointer = 0
    
    MsgBox "���۵Ǿ����ϴ�.", vbInformation, "Ȯ��"
    
Exit Sub

'/----------------------------------------------------------------------------------------------------/

ERN_ERR:
    Screen.MousePointer = 0
End Sub

Private Sub cmdView_Click()
    Call FUNC_MM_KEY_CLEAR("1") '/�����ڳ��� Spread Clear
    Call FUNC_MM_KEY_CLEAR("2") '/Patient Information Clear
    If optImage���ۿ���(1).Value = True Then
        Call FUNC_MM_KEY_CLEAR("5") '/���ۿϷ� Spread Clear
        Call FUNC_MM_KEY_CLEAR("3") '/Image Clear
    End If
    
    Select Case True
        Case opt���ۿ���(0).Value
            Call FUNC_MM_VIEW("1") '/�����ڳ���(������)
            If spr�����ڳ���.MaxRows > 0 Then Call spr�����ڳ���_LeaveCell(0, 0, 1, 1, False)
            
        Case opt���ۿ���(1).Value
            Call FUNC_MM_VIEW("2") '/�����ڳ���(����)
            If spr�����ڳ���.MaxRows > 0 Then Call spr�����ڳ���_LeaveCell(0, 0, 1, 1, False)
            
            If optImage���ۿ���(1).Value = True Then Call cmdViewFTP_Click
    End Select
End Sub

Private Sub cmdViewFTP_Click()
    Call FUNC_MM_KEY_CLEAR("5") '/���ۿϷ� Spread Clear
    Call FUNC_MM_KEY_CLEAR("3") '/Image Clear
    
    If Trim(gtypEQ_INFO.FTPIMGFILEPATH) = "" Then '/DB���� ������ FTP������������ΰ� ���� ��...
        If MsgBox("���ۿϷ� ������ �������� �ʽ��ϴ�." & vbCrLf & vbCrLf & _
                   "�⺻ ������ �����ϰڽ��ϱ�?", vbQuestion + vbOKCancel, "���ۿϷ���������") = vbCancel Then Exit Sub
        
        GoSub RTN_�⺻���������׼���
    
    Else '/DB���� ������ FTP������������ΰ� ���� ��...
        '/�������� ���� Ȯ��
        If Dir(gtypEQ_INFO.FTPIMGFILEPATH, vbDirectory) = "" Then '/���������� ������...
            MsgBox "������ ������ ���ۿϷ� ������ �������� �ʽ��ϴ�." & vbCrLf & vbCrLf & _
                   "������ �Ϸ��� [���ۿϷ���������]�� �����Ͻʽÿ�.", vbInformation, "Ȯ��"
            Exit Sub
        End If
    End If
    If spr�����ڳ���.MaxRows > 0 And Trim(lbl���Ϲ�ȣ) = "" Then MsgBox "��ȸ�� �����ڸ� �����Ͻʽÿ�!", vbCritical, "��ȸ�Ұ�": Exit Sub
    
    Call FUNC_MM_VIEW("4") '/���ۿϷ� �ڷ�ҷ�����
    If spr���ۿϷ�.MaxRows > 0 Then
        imvResult.Filename = ""
        imvResult.Filename = gtypEQ_INFO.FTPIMGFILEPATH & "\" & GET_CELL(spr���ۿϷ�, 2, spr���ۿϷ�.ActiveRow)
        Call SUB_CHK_PDF_TIF(imvResult.Filename)
        imvResult.View = 9
    End If
Exit Sub

'/----------------------------------------------------------------------------------------------------/

RTN_�⺻���������׼���:
    '/1.�⺻���� ����
    If SET_DEFAULT_FOLDER("FTP") = False Then
        MsgBox "�⺻������ �������� �ʾҽ��ϴ�." & vbCrLf & vbCrLf & _
               "����� Ȥ�� ���޾�ü�� �����Ͻñ� �ٶ��ϴ�.", vbCritical, "�⺻������������"
               
        Exit Sub
    End If
    
    '/2.��������
    gtypEQ_INFO.FTPIMGFILEPATH = App.Path & "\" & gtypEQ_INFO.EQUIPCODE & "\" & gtypEQ_INFO.EQUIPSEQ & "\" & "EQUIP"
Return
End Sub

Private Sub cmdViewLocal_Click()
    Call FUNC_MM_KEY_CLEAR("4") '/������ Spread Clear
    Call FUNC_MM_KEY_CLEAR("3") '/Image Clear
    
    If Trim(gtypEQ_INFO.EQIMGFILEPATH) = "" Then '/DB���� ������ ��������������ΰ� ���� ��...
        If MsgBox("������ ������ �������� �ʽ��ϴ�." & vbCrLf & vbCrLf & _
                   "�⺻ ������ �����ϰڽ��ϱ�?", vbQuestion + vbOKCancel, "��������������") = vbCancel Then Exit Sub
        
        GoSub RTN_�⺻���������׼���
        
    Else '/DB���� ������ ��������������ΰ� ���� ��...
        '/�������� ���� Ȯ��
        If Dir(gtypEQ_INFO.EQIMGFILEPATH, vbDirectory) = "" Then '/���������� ������...
            MsgBox "������ ������ ������ ������ �������� �ʽ��ϴ�." & vbCrLf & vbCrLf & _
                   "������ �Ϸ��� [��������������]�� �����Ͻʽÿ�.", vbInformation, "Ȯ��"
            Exit Sub
        End If
    End If
    
    Call FUNC_MM_VIEW("3") '/������ �ڷ�ҷ�����
    If spr������.MaxRows > 0 Then
        imvResult.Filename = ""
        imvResult.Filename = gtypEQ_INFO.EQIMGFILEPATH & "\" & GET_CELL(spr������, 2, spr������.ActiveRow)
        Call SUB_CHK_PDF_TIF(imvResult.Filename)
        imvResult.View = 9
    End If
Exit Sub

'/----------------------------------------------------------------------------------------------------/

RTN_�⺻���������׼���:
    '/1.�⺻���� ����
    If SET_DEFAULT_FOLDER("EQUIP") = False Then
        MsgBox "�⺻������ �������� �ʾҽ��ϴ�." & vbCrLf & vbCrLf & _
               "����� Ȥ�� ���޾�ü�� �����Ͻñ� �ٶ��ϴ�.", vbCritical, "�⺻������������"
               
        Exit Sub
    End If
    
    '/2.��������
    gtypEQ_INFO.EQIMGFILEPATH = App.Path & "\" & gtypEQ_INFO.EQUIPCODE & "\" & gtypEQ_INFO.EQUIPSEQ & "\" & "EQUIP"
Return
End Sub

Private Sub cmdzoomin_Click()
'''    Dim strEXAMCODE As String
'''    Dim strUpEXAMCODE As String
'''    Dim nInc        As Integer
'''
'''    If OpenDB(gstrREG_DB_CONSTR) = True Then
'''        ADC.BeginTrans
'''
'''        gstrQuy = "SELECT EXAMCODE, ROWID "
'''        gstrQuy = gstrQuy & vbCrLf & "  FROM EXAMMASTER_1 "
'''        gstrQuy = gstrQuy & vbCrLf & " ORDER BY EXAMCODE "
'''        If ReadSQL(gstrQuy, ADR) = False Then Call CloseDB: End
'''
'''        If Not ADR Is Nothing Then
'''            Do Until ADR.EOF
'''                If strEXAMCODE = Trim(ADR!EXAMCODE & "") Then
'''                    nInc = nInc + 1
'''
'''                    strUpEXAMCODE = Trim(ADR!EXAMCODE & "") & "_" & CStr(nInc)
'''                Else
'''                    nInc = 1
'''
'''                    strUpEXAMCODE = Trim(ADR!EXAMCODE & "") & "_" & CStr(nInc)
'''
'''                    strEXAMCODE = Trim(ADR!EXAMCODE & "")
'''                End If
'''
'''                gstrQuy = "UPDATE EXAMMASTER_1 SET "
'''                gstrQuy = gstrQuy & vbCrLf & "       EXAMCODE1 = '" & strUpEXAMCODE & "' "
'''                gstrQuy = gstrQuy & vbCrLf & " WHERE ROWID  = '" & Trim(ADR!ROWID & "") & "' "
'''                If RunSQL(gstrQuy) = False Then ADC.RollbackTrans: Call CloseDB: End
'''
'''                ADR.MoveNext
'''            Loop
'''            ADR.Close: Set ADR = Nothing
'''        End If
'''
'''        ADC.CommitTrans
'''
'''        Call CloseDB
'''    End If
    imvResult.ZoomIn
End Sub

Private Sub cmdzoomout_Click()
    imvResult.ZoomOut
End Sub

Private Sub cmdZmHeight_Click()
    cmdzoomin.Enabled = True
    cmdzoomout.Enabled = True
    cboZoomValue.Enabled = True
    cmdRotate.Enabled = True
    
    imvResult.View = 11
    imvResult.SetFocus
End Sub

Private Sub cmdFit_Click()
    cmdzoomin.Enabled = True
    cmdzoomout.Enabled = True
    cboZoomValue.Enabled = True
    cmdRotate.Enabled = True
    
    imvResult.View = 9
    imvResult.SetFocus
End Sub

Private Sub cmdZmWidth_Click()
    cmdzoomin.Enabled = True
    cmdzoomout.Enabled = True
    cboZoomValue.Enabled = True
    cmdRotate.Enabled = True
    
    imvResult.View = 10
    imvResult.SetFocus
End Sub

Private Sub cmdCenter_Click()
    cmdzoomin.Enabled = False
    cmdzoomout.Enabled = False
    cboZoomValue.Enabled = False
    cmdRotate.Enabled = False
    
    imvResult.View = 12
    imvResult.SetFocus
End Sub

Private Sub cmd100_Click()
    cmdzoomin.Enabled = True
    cmdzoomout.Enabled = True
    cboZoomValue.Enabled = True
    cmdRotate.Enabled = True
    
    imvResult.View = 5
    imvResult.SetFocus
End Sub

Private Sub dtp��������_Change()
    Call FUNC_MM_KEY_CLEAR("1") '/�����ڳ��� Spread Clear
    Call FUNC_MM_KEY_CLEAR("2") '/Patient Information Clear
    If optImage���ۿ���(1).Value = True Then
        Call FUNC_MM_KEY_CLEAR("5") '/���ۿϷ� Spread Clear
        Call FUNC_MM_KEY_CLEAR("3") '/Image Clear
    End If
End Sub

Private Sub dtp��������_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim ShiftDown, AltDown, CtrlDown, Txt
   
    ShiftDown = (Shift And vbShiftMask) > 0
    AltDown = (Shift And vbAltMask) > 0
    CtrlDown = (Shift And vbCtrlMask) > 0
   
    If KeyCode = vbKeyM Then   ' Ű�� ���� ���¸� ����մϴ�.
        If mnuSetting.Visible = True Then
            mnuSetting.Visible = False
        Else
            mnuSetting.Visible = True
        End If
    End If
End Sub

Private Sub Form_Load()
    Call FUNC_MM_INITIAL
    
    DoEvents
    DoEvents
    DoEvents
    
    Call cmdView_Click
    
    DoEvents
    DoEvents
    DoEvents
    
    Call cmdViewLocal_Click
End Sub

Private Sub Form_Resize()
On Error Resume Next
    '/object.Move Left, Top, Width, Height
    '/(((Me.Height - lngMeHeight) / 3) * 2) : ���̰� �þ�� ��ü 3��, �����λ� �ش� ��ü ���� �þ ��ü�� 2��
    For intX = 0 To UBound(CW)
        Select Case CW(intX).Nm
            Case fraLimageList.Name:        fraLimageList.Move CW(intX).Left, CW(intX).Top, CW(intX).Width, CW(intX).Height + (Me.Height - lngMeHeight)
            Case picImageList������.Name:   picImageList������.Move CW(intX).Left, CW(intX).Top, CW(intX).Width, CW(intX).Height + (Me.Height - lngMeHeight)
            Case picImageList���ۿϷ�.Name: picImageList���ۿϷ�.Move CW(intX).Left, CW(intX).Top, CW(intX).Width, CW(intX).Height + (Me.Height - lngMeHeight)
            Case spr������.Name:            spr������.Move CW(intX).Left, CW(intX).Top, CW(intX).Width, CW(intX).Height + (Me.Height - lngMeHeight)
            Case spr���ۿϷ�.Name:          spr���ۿϷ�.Move CW(intX).Left, CW(intX).Top, CW(intX).Width, CW(intX).Height + (Me.Height - lngMeHeight)
            Case cmdViewLocal.Name:         cmdViewLocal.Move CW(intX).Left, CW(intX).Top + (Me.Height - lngMeHeight), CW(intX).Width, CW(intX).Height
            Case cmd��������������.Name:    cmd��������������.Move CW(intX).Left, CW(intX).Top + (Me.Height - lngMeHeight), CW(intX).Width, CW(intX).Height
            Case cmdDeleteLocal.Name:       cmdDeleteLocal.Move CW(intX).Left, CW(intX).Top + (Me.Height - lngMeHeight), CW(intX).Width, CW(intX).Height
            Case cmdViewFTP.Name:           cmdViewFTP.Move CW(intX).Left, CW(intX).Top + (Me.Height - lngMeHeight), CW(intX).Width, CW(intX).Height
            Case cmd���ۿϷ���������.Name:  cmd���ۿϷ���������.Move CW(intX).Left, CW(intX).Top + (Me.Height - lngMeHeight), CW(intX).Width, CW(intX).Height
            Case cmdDeleteFTP.Name:         cmdDeleteFTP.Move CW(intX).Left, CW(intX).Top + (Me.Height - lngMeHeight), CW(intX).Width, CW(intX).Height
            Case cmdSave.Name:              cmdSave.Move CW(intX).Left + (Me.Width - lngMeWidth), CW(intX).Top, CW(intX).Width, CW(intX).Height
            Case shpPatientInfo.Name:       shpPatientInfo.Move CW(intX).Left, CW(intX).Top, CW(intX).Width + (Me.Width - lngMeWidth), CW(intX).Height
            Case prgPatient.Name:           prgPatient.Move CW(intX).Left, CW(intX).Top, CW(intX).Width + (Me.Width - lngMeWidth), CW(intX).Height
            Case picImage.Name:             picImage.Move CW(intX).Left, CW(intX).Top, CW(intX).Width + (Me.Width - lngMeWidth), CW(intX).Height + (Me.Height - lngMeHeight)
            Case imvResult.Name:            imvResult.Move CW(intX).Left, CW(intX).Top, CW(intX).Width + (Me.Width - lngMeWidth), CW(intX).Height + (Me.Height - lngMeHeight)
            Case picTifPdf.Name:            picTifPdf.Move CW(intX).Left, CW(intX).Top, CW(intX).Width + (Me.Width - lngMeWidth), CW(intX).Height
            Case picControl.Name:           picControl.Move CW(intX).Left, CW(intX).Top, CW(intX).Width + (Me.Width - lngMeWidth), CW(intX).Height
            Case picJPG.Name:               picJPG.Move CW(intX).Left, CW(intX).Top, CW(intX).Width + (Me.Width - lngMeWidth), CW(intX).Height
        End Select
    Next intX
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call MMSFTP.CloseConnection

    Set MMSFTP = Nothing
    Call CloseDB
    
    If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
    
    imvResult.Filename = ""
    Call SUB_CHK_PDF_TIF("")
'''    imvTif.Filename = ""

    DoEvents
    DoEvents
    DoEvents
    DoEvents
    DoEvents

    '/FTP������������� ������ ��� ���� ����
    If Len(Dir(gtypEQ_INFO.FTPIMGFILEPATH & "\*.*")) > 0 Then
        Kill gtypEQ_INFO.FTPIMGFILEPATH & "\*.*"
    End If

    Set frmVPM_Main = Nothing
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuInfo_Click()
    frm����_Info.Show vbModal
End Sub

Private Sub mnuSettingSub_Click(Index As Integer)
    Select Case Index
        Case 0: frm����_Set_DataBase.Show vbModal
        Case 1: frm����_Set_Equipment_List.Show vbModal
        Case 2: frm����_Set_Equip_Config.Show vbModal
    End Select
End Sub

Private Sub MSComm1_OnComm()
    '/EMR Interface ��� ��� �� ��񿡼� ������ ��ȣ�� SM�϶�
    ' ����� �� �ִ� ����(Spread ��)�� ����� ���������ͷ� ��� ��Ų��.
    Select Case gtypEQ_INFO.EQUIPCODE
        Case "00008" '/AL2000(�Ȱ����)
            strTemp = MSComm1.Input
            
            Select Case strTemp
                Case Chr(1) 'SOH
                    gstrMSCOMM_Buff = ""
                    gstrMSCOMM_Buff = strTemp
                
                Case Chr(4) 'EOT
                    gstrMSCOMM_Buff = gstrMSCOMM_Buff & strTemp
                    txtSerialData = gstrMSCOMM_Buff
                    Call frmVPM_SM00008.FUNC_MM_PRINT(gstrMSCOMM_Buff)
                    
                    gstrMSCOMM_Buff = ""
                
                Case Else
                    gstrMSCOMM_Buff = gstrMSCOMM_Buff & strTemp
            End Select
        
        Case "00014" '/�������(�Ȱ����)
            strTemp = MSComm1.Input
            
            Select Case strTemp
                Case vbCr
                
                Case vbLf
                    If Mid(gstrMSCOMM_Buff, 1, 5) = "LM2RK" Then
                        txtSerialData = gstrMSCOMM_Buff
                        Call frmVPM_SM00014.FUNC_MM_PRINT(gstrMSCOMM_Buff)
                    End If
                    gstrMSCOMM_Buff = ""
                    
                Case Else
                    gstrMSCOMM_Buff = gstrMSCOMM_Buff & strTemp
            End Select
        
        Case "00016" '/CT80(�Ȱ����)
            strTemp = MSComm1.Input
            
            Select Case strTemp
                Case Chr(1)
                    gstrMSCOMM_Buff = strTemp
                
                Case Chr(4)
                    gstrMSCOMM_Buff = gstrMSCOMM_Buff & strTemp
                    txtSerialData = gstrMSCOMM_Buff
                    Call frmVPM_SM00016.FUNC_MM_PRINT(gstrMSCOMM_Buff)
                    
                    gstrMSCOMM_Buff = ""
                
                Case Else
                    gstrMSCOMM_Buff = gstrMSCOMM_Buff & strTemp
            End Select
    
        Case "00025" '/KR7100(�Ȱ����)
            strTemp = MSComm1.Input
            
            Select Case strTemp
                Case Chr(1)
                    gstrMSCOMM_Buff = strTemp
                
                Case Chr(4)
                    gstrMSCOMM_Buff = gstrMSCOMM_Buff & strTemp
                    txtSerialData = gstrMSCOMM_Buff
                    Call frmVPM_SM00025.FUNC_MM_PRINT(gstrMSCOMM_Buff)
                    
                    gstrMSCOMM_Buff = ""
                
                Case Else
                    gstrMSCOMM_Buff = gstrMSCOMM_Buff & strTemp
            End Select
    End Select
End Sub

Private Sub opt���ۿ���_Click(Index As Integer)
    Select Case Index
        Case 0:
            opt���ۿ���(0).ForeColor = RGB(0, 0, 255)
            opt���ۿ���(0).FontBold = True
            opt���ۿ���(1).ForeColor = RGB(0, 0, 0)
            opt���ۿ���(1).FontBold = False
        Case 1:
            opt���ۿ���(0).ForeColor = RGB(0, 0, 0)
            opt���ۿ���(0).FontBold = False
            opt���ۿ���(1).ForeColor = RGB(0, 0, 255)
            opt���ۿ���(1).FontBold = True
    End Select
    
    Call FUNC_MM_KEY_CLEAR("1") '/�����ڳ��� Spread Clear
    Call FUNC_MM_KEY_CLEAR("2") '/Patient Information Clear
    If optImage���ۿ���(1).Value = True Then
        Call FUNC_MM_KEY_CLEAR("5") '/���ۿϷ� Spread Clear
        Call FUNC_MM_KEY_CLEAR("3") '/Image Clear
    End If
End Sub

Private Sub opt���ۿ���_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub optImage���ۿ���_Click(Index As Integer)
    Select Case Index
        Case 0: '/������ Tab
            optImage���ۿ���(0).ForeColor = RGB(0, 0, 255)
            optImage���ۿ���(0).FontBold = True
            optImage���ۿ���(1).ForeColor = RGB(0, 0, 0)
            optImage���ۿ���(1).FontBold = False
            
            picImageList������.Visible = True
            picImageList���ۿϷ�.Visible = False
            
            cmdSave.Enabled = True
            
            Call cmdViewLocal_Click
            
            Call FUNC_MM_KEY_CLEAR("3") '/Image Clear
            
            imvResult.Filename = ""
            imvResult.Filename = gtypEQ_INFO.EQIMGFILEPATH & "\" & GET_CELL(spr������, 2, spr������.ActiveRow)
            Call SUB_CHK_PDF_TIF(imvResult.Filename)
            imvResult.View = 9
            
        Case 1: '/���ۿϷ� Tab
            optImage���ۿ���(0).ForeColor = RGB(0, 0, 0)
            optImage���ۿ���(0).FontBold = False
            optImage���ۿ���(1).ForeColor = RGB(0, 0, 255)
            optImage���ۿ���(1).FontBold = True
            
            picImageList������.Visible = False
            picImageList���ۿϷ�.Visible = True
            
            cmdSave.Enabled = False
            
            Call cmdViewFTP_Click
            
            Call FUNC_MM_KEY_CLEAR("3") '/Image Clear
            
            imvResult.Filename = ""
            imvResult.Filename = gtypEQ_INFO.FTPIMGFILEPATH & "\" & GET_CELL(spr���ۿϷ�, 2, spr���ۿϷ�.ActiveRow)
            Call SUB_CHK_PDF_TIF(imvResult.Filename)
            imvResult.View = 9
    End Select

End Sub

Private Sub spr������_Click(ByVal Col As Long, ByVal Row As Long)
    With spr������
        If Row > 0 Then Exit Sub
            
        If Col = 1 Then
            If GET_CELL(spr������, 1, 1) = "0" Then
                strTemp = "1"
            Else
                strTemp = "0"
            End If
            
            For intX = 1 To spr������.MaxRows
                If strTemp = "0" Then
                    Call SET_CELL(spr������, 1, intX, "0")
                Else
                    Call SET_CELL(spr������, 1, intX, "1")
                End If
            Next intX
        Else
            If Col < 2 Then Exit Sub
            
            .Col = -1
            .Row = 1
            .Col2 = -1
            .Row2 = .MaxRows
            .BlockMode = True
            .SortBy = SortByRow
            
            .SortKey(1) = Col
            If Val(Mid(spr������.Tag, 2)) = Col Then
                If Left(spr������.Tag, 1) = "A" Then
                    .SortKeyOrder(1) = SortKeyOrderDescending
                    spr������.Tag = "D" & CStr(Col)
                Else
                    .SortKeyOrder(1) = SortKeyOrderAscending
                    spr������.Tag = "A" & CStr(Col)
                End If
            Else
                .SortKeyOrder(1) = SortKeyOrderAscending
                spr������.Tag = "A" & CStr(Col)
            End If
            
            .Action = ActionSort
            .BlockMode = False
        
            Call spr������_LeaveCell(0, 0, 1, spr������.ActiveRow, False)
        End If
    End With
End Sub

Private Sub spr������_DblClick(ByVal Col As Long, ByVal Row As Long)
    If Row < 1 Then Exit Sub
        
    If GET_CELL(spr������, 1, Row) = "1" Then
        Call SET_CELL(spr������, 1, Row, "0")
    Else
        Call SET_CELL(spr������, 1, Row, "1")
    End If
End Sub

Private Sub spr������_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If GET_CELL(spr������, 1, spr������.ActiveRow) = "1" Then
            Call SET_CELL(spr������, 1, spr������.ActiveRow, "0")
        Else
            Call SET_CELL(spr������, 1, spr������.ActiveRow, "1")
        End If
    End If
End Sub

Private Sub spr������_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    If NewRow < 1 Then Exit Sub
    If Row = NewRow Then Exit Sub

    Call FUNC_MM_KEY_CLEAR("3") '/Image Clear
    
    imvResult.Filename = ""
    imvResult.Filename = gtypEQ_INFO.EQIMGFILEPATH & "\" & GET_CELL(spr������, 2, NewRow)
    Call SUB_CHK_PDF_TIF(imvResult.Filename)
    imvResult.View = 9
End Sub

Private Sub spr������_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    Dim Message, Title, Default, MyValue
    Dim OldName, NewName
    Dim strExtension    As String
    
    If Row < 1 Then Exit Sub
    
    Message = "������ ���ϸ��� �Է��Ͻʽÿ�"   ' ������Ʈ ����.
    Title = "Image ���ϸ� ����"   ' ���� ����.
    ' �޽��� ȭ�� ǥ��, ����, �⺻��.
    MyValue = InputBox(Message, Title, Left(GET_CELL(spr������, 2, Row), InStr(GET_CELL(spr������, 2, Row), ".") - 1))
    
    strExtension = Mid(GET_CELL(spr������, 2, Row), InStr(GET_CELL(spr������, 2, Row), ".") + 1)
    
    If Trim(MyValue) <> "" Then
        OldName = gtypEQ_INFO.EQIMGFILEPATH & "\" & GET_CELL(spr������, 2, Row)
        NewName = gtypEQ_INFO.EQIMGFILEPATH & "\" & MyValue & "." & strExtension  ' ���� �̸��� �����մϴ�.
        
        If Row = spr������.ActiveRow Then
            imvResult.Filename = ""
            Call SUB_CHK_PDF_TIF("")
        End If
        
        Name OldName As NewName   ' ���� �̸��� �����մϴ�.

        If Row = spr������.ActiveRow Then
            imvResult.Filename = ""
            imvResult.Filename = gtypEQ_INFO.EQIMGFILEPATH & "\" & MyValue & "." & strExtension
            Call SUB_CHK_PDF_TIF(imvResult.Filename)
            imvResult.View = 9
        End If
        
        Call SET_CELL(spr������, 2, Row, MyValue & "." & strExtension)
    End If
End Sub

Private Sub spr�����ڳ���_Click(ByVal Col As Long, ByVal Row As Long)
    With spr�����ڳ���
        If Row > 0 Then Exit Sub
            
        If Col = 1 Then
            If GET_CELL(spr�����ڳ���, 1, 1) = "0" Then
                strTemp = "1"
            Else
                strTemp = "0"
            End If
            
            For intX = 1 To spr�����ڳ���.MaxRows
                If strTemp = "0" Then
                    Call SET_CELL(spr�����ڳ���, 1, intX, "0")
                Else
                    Call SET_CELL(spr�����ڳ���, 1, intX, "1")
                End If
            Next intX
        Else
            If Col < 2 Then Exit Sub
            
            .Col = -1
            .Row = 1
            .Col2 = -1
            .Row2 = .MaxRows
            .BlockMode = True
            .SortBy = SortByRow
            
            .SortKey(1) = Col
            If Val(Mid(spr�����ڳ���.Tag, 2)) = Col Then
                If Left(spr�����ڳ���.Tag, 1) = "A" Then
                    .SortKeyOrder(1) = SortKeyOrderDescending
                    spr�����ڳ���.Tag = "D" & CStr(Col)
                Else
                    .SortKeyOrder(1) = SortKeyOrderAscending
                    spr�����ڳ���.Tag = "A" & CStr(Col)
                End If
            Else
                .SortKeyOrder(1) = SortKeyOrderAscending
                spr�����ڳ���.Tag = "A" & CStr(Col)
            End If
            
            .Action = ActionSort
            .BlockMode = False
        
            Call spr�����ڳ���_LeaveCell(0, 0, 1, spr�����ڳ���.ActiveRow, False)
        End If
    End With
End Sub

Private Sub spr�����ڳ���_DblClick(ByVal Col As Long, ByVal Row As Long)
    If Row < 1 Then Exit Sub
        
    If GET_CELL(spr�����ڳ���, 1, Row) = "1" Then
        Call SET_CELL(spr�����ڳ���, 1, Row, "0")
    Else
        Call SET_CELL(spr�����ڳ���, 1, Row, "1")
    End If
End Sub

Private Sub spr�����ڳ���_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If GET_CELL(spr�����ڳ���, 1, spr�����ڳ���.ActiveRow) = "1" Then
            Call SET_CELL(spr�����ڳ���, 1, spr�����ڳ���.ActiveRow, "0")
        Else
            Call SET_CELL(spr�����ڳ���, 1, spr�����ڳ���.ActiveRow, "1")
        End If
    End If
End Sub

Private Sub spr�����ڳ���_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    If NewRow < 1 Then Exit Sub
    If Row = NewRow Then Exit Sub
    
    Call FUNC_MM_KEY_CLEAR("2") '/Patient Information Clear

    lbl���Ϲ�ȣ = GET_CELL(spr�����ڳ���, 2, NewRow)
    lbl�����ڸ� = GET_CELL(spr�����ڳ���, 5, NewRow)
    lbl���ɼ��� = GET_CELL(spr�����ڳ���, 6, NewRow)
    lbló��� = GET_CELL(spr�����ڳ���, 10, NewRow)
    lbló���ڵ� = GET_CELL(spr�����ڳ���, 11, NewRow)

    lbl����� = GET_CELL(spr�����ڳ���, 4, NewRow)
    lbl�Կܱ��� = GET_CELL(spr�����ڳ���, 3, NewRow)

    lbló������ = GET_CELL(spr�����ڳ���, 7, NewRow)
    lbló��SEQ = GET_CELL(spr�����ڳ���, 8, NewRow)
    lbl������� = GET_CELL(spr�����ڳ���, 9, NewRow)
    
    Select Case GET_CELL(spr�����ڳ���, 12, NewRow) '/�ǽ�ó�� ó���������Flag
        Case "440": lbló����� = "����": lbló�����.ForeColor = RGB(0, 0, 0)
        Case "560": lbló����� = "�ӽð��": lbló�����.ForeColor = RGB(0, 255, 0)
        Case "610": lbló����� = "�ǽÿϷ�": lbló�����.ForeColor = RGB(255, 0, 0)
        Case Else:  lbló����� = GET_CELL(spr�����ڳ���, 12, NewRow)
    End Select
    
    If optImage���ۿ���(1).Value = True Then Call cmdViewFTP_Click
End Sub

Private Sub spr���ۿϷ�_Click(ByVal Col As Long, ByVal Row As Long)
    With spr���ۿϷ�
        If Row > 0 Then Exit Sub
            
        If Col = 1 Then
            If GET_CELL(spr���ۿϷ�, 1, 1) = "0" Then
                strTemp = "1"
            Else
                strTemp = "0"
            End If
            
            For intX = 1 To spr���ۿϷ�.MaxRows
                If strTemp = "0" Then
                    Call SET_CELL(spr���ۿϷ�, 1, intX, "0")
                Else
                    Call SET_CELL(spr���ۿϷ�, 1, intX, "1")
                End If
            Next intX
        Else
            If Col < 2 Then Exit Sub
            
            .Col = -1
            .Row = 1
            .Col2 = -1
            .Row2 = .MaxRows
            .BlockMode = True
            .SortBy = SortByRow
            
            .SortKey(1) = Col
            If Val(Mid(spr���ۿϷ�.Tag, 2)) = Col Then
                If Left(spr���ۿϷ�.Tag, 1) = "A" Then
                    .SortKeyOrder(1) = SortKeyOrderDescending
                    spr���ۿϷ�.Tag = "D" & CStr(Col)
                Else
                    .SortKeyOrder(1) = SortKeyOrderAscending
                    spr���ۿϷ�.Tag = "A" & CStr(Col)
                End If
            Else
                .SortKeyOrder(1) = SortKeyOrderAscending
                spr���ۿϷ�.Tag = "A" & CStr(Col)
            End If
            
            .Action = ActionSort
            .BlockMode = False
        
            Call spr���ۿϷ�_LeaveCell(0, 0, 1, spr���ۿϷ�.ActiveRow, False)
        End If
    End With
End Sub

Private Sub spr���ۿϷ�_DblClick(ByVal Col As Long, ByVal Row As Long)
    If Row < 1 Then Exit Sub
        
    If GET_CELL(spr���ۿϷ�, 1, Row) = "1" Then
        Call SET_CELL(spr���ۿϷ�, 1, Row, "0")
    Else
        Call SET_CELL(spr���ۿϷ�, 1, Row, "1")
    End If
End Sub

Private Sub spr���ۿϷ�_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If GET_CELL(spr���ۿϷ�, 1, spr���ۿϷ�.ActiveRow) = "1" Then
            Call SET_CELL(spr���ۿϷ�, 1, spr���ۿϷ�.ActiveRow, "0")
        Else
            Call SET_CELL(spr���ۿϷ�, 1, spr���ۿϷ�.ActiveRow, "1")
        End If
    End If
End Sub

Private Sub spr���ۿϷ�_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    If NewRow < 1 Then Exit Sub
    If Row = NewRow Then Exit Sub

    Call FUNC_MM_KEY_CLEAR("3") '/Image Clear
    
    imvResult.Filename = ""
    imvResult.Filename = gtypEQ_INFO.FTPIMGFILEPATH & "\" & GET_CELL(spr���ۿϷ�, 2, NewRow)
    Call SUB_CHK_PDF_TIF(imvResult.Filename)
    imvResult.View = 9
End Sub

Public Sub SUB_GET_REG_CLIENT_INFO()
    Dim strEQCD             As String
    Dim strEQNM             As String
    Dim strEQSEQ            As String
    Dim strEQPOS            As String
    Dim strEQTYPE           As String
    Dim strRECEIVETYPE      As String
    Dim strEQUIPPORT        As String
    Dim strORDYN            As String
    Dim strQUERYTYPE        As String
    Dim strZIPYN            As String
    Dim strSERIALYN         As String
    Dim strSERIALPORT       As String
    Dim strSERIALBAUD       As String
    Dim strSERIALDATABIT    As String
    Dim strSERIALSTARTBIT   As String
    Dim strSERIALSTOPBIT    As String
    Dim strSERIALPARITY     As String
    Dim strSERIALRTS        As String
    Dim strSERIALDTR        As String
    Dim strEQIMGFILEPATH    As String
    Dim strFTPIMGFILEPATH   As String
    
    Dim strEQCD_Array
    Dim strEQNM_Array
    Dim strEQSEQ_Array
    Dim strEQPOS_Array
    Dim strEQTYPE_Array
    Dim strRECEIVETYPE_Array
    Dim strEQUIPPORT_Array
    Dim strORDYN_Array
    Dim strQUERYTYPE_Array
    Dim strZIPYN_Array
    Dim strSERIALYN_Array
    Dim strSERIALPORT_Array
    Dim strSERIALBAUD_Array
    Dim strSERIALDATABIT_Array
    Dim strSERIALSTARTBIT_Array
    Dim strSERIALSTOPBIT_Array
    Dim strSERIALPARITY_Array
    Dim strSERIALRTS_Array
    Dim strSERIALDTR_Array
    Dim strEQIMGFILEPATH_Array
    Dim strFTPIMGFILEPATH_Array
    
    '/��� �Ƿ���� ����(��������) ��������
    strEQCD = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_EQCD)
    strEQNM = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_EQNM)
    strEQSEQ = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_EQSEQ)
    strEQPOS = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_EQPOS)
    strEQTYPE = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_EQTYPE)
    strRECEIVETYPE = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_RECEIVETYPE)
    strEQUIPPORT = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_EQUIPPORT)
    strORDYN = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_ORDYN)
    strQUERYTYPE = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_QUERYTYPE)
    strZIPYN = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_ZIPYN)
    strSERIALYN = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_SERIALYN)
    strSERIALPORT = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_SERIALPORT)
    strSERIALBAUD = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_SERIALBAUD)
    strSERIALDATABIT = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_SERIALDATABIT)
    strSERIALSTARTBIT = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_SERIALSTARTBIT)
    strSERIALSTOPBIT = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_SERIALSTOPBIT)
    strSERIALPARITY = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_SERIALPARITY)
    strSERIALRTS = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_SERIALRTS)
    strSERIALDTR = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_SERIALDTR)
    strEQIMGFILEPATH = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_EQIMGFILEPATH)
    strFTPIMGFILEPATH = GetSetting(REG_MAKER & "\" & REG_PRODUCT, REG_CLIENT_INFO, REG_CLIENT_FTPIMGFILEPATH)
    
    strEQCD_Array = Split(strEQCD, ",")
    strEQNM_Array = Split(strEQNM, ",")
    strEQSEQ_Array = Split(strEQSEQ, ",")
    strEQPOS_Array = Split(strEQPOS, ",")
    strEQTYPE_Array = Split(strEQTYPE, ",")
    strRECEIVETYPE_Array = Split(strRECEIVETYPE, ",")
    strEQUIPPORT_Array = Split(strEQUIPPORT, ",")
    strORDYN_Array = Split(strORDYN, ",")
    strQUERYTYPE_Array = Split(strQUERYTYPE, ",")
    strZIPYN_Array = Split(strZIPYN, ",")
    strSERIALYN_Array = Split(strSERIALYN, ",")
    strSERIALPORT_Array = Split(strSERIALPORT, ",")
    strSERIALBAUD_Array = Split(strSERIALBAUD, ",")
    strSERIALDATABIT_Array = Split(strSERIALDATABIT, ",")
    strSERIALSTARTBIT_Array = Split(strSERIALSTARTBIT, ",")
    strSERIALSTOPBIT_Array = Split(strSERIALSTOPBIT, ",")
    strSERIALPARITY_Array = Split(strSERIALPARITY, ",")
    strSERIALRTS_Array = Split(strSERIALRTS, ",")
    strSERIALDTR_Array = Split(strSERIALDTR, ",")
    strEQIMGFILEPATH_Array = Split(strEQIMGFILEPATH, ",")
    strFTPIMGFILEPATH_Array = Split(strFTPIMGFILEPATH, ",")
    
    On Error Resume Next
    
    With sprEQ_INFO
        If .MaxRows > 0 Then .MaxRows = 0
        
        For intX = 0 To UBound(strEQCD_Array)
            .MaxRows = .MaxRows + 1: .Row = .MaxRows
        
            .Col = 1:   .Text = strEQCD_Array(intX)
            .Col = 2:   .Text = strEQNM_Array(intX)
            .Col = 3:   .Text = strEQSEQ_Array(intX)
            .Col = 4:   .Text = strEQPOS_Array(intX)
            .Col = 5:   .Text = strEQTYPE_Array(intX)
            .Col = 6:   .Text = strRECEIVETYPE_Array(intX)
            .Col = 7:   .Text = strEQUIPPORT_Array(intX)
            .Col = 8:   .Text = strORDYN_Array(intX)
            .Col = 9:   .Text = strQUERYTYPE_Array(intX)
            .Col = 10:  .Text = strZIPYN_Array(intX)
            .Col = 11:  .Text = strSERIALYN_Array(intX)
            .Col = 12:  .Text = strSERIALPORT_Array(intX)
            .Col = 13:  .Text = strSERIALBAUD_Array(intX)
            .Col = 14:  .Text = strSERIALDATABIT_Array(intX)
            .Col = 15:  .Text = strSERIALSTARTBIT_Array(intX)
            .Col = 16:  .Text = strSERIALSTOPBIT_Array(intX)
            .Col = 17:  .Text = strSERIALPARITY_Array(intX)
            .Col = 18:  .Text = strSERIALRTS_Array(intX)
            .Col = 19:  .Text = strSERIALDTR_Array(intX)
            .Col = 20:  .Text = strEQIMGFILEPATH_Array(intX)
            .Col = 21:  .Text = strFTPIMGFILEPATH_Array(intX)
        Next intX
    End With
    
    On Error GoTo 0
End Sub

Private Sub staCondition_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel = "COM" Then
        If txtSerialData.Visible = False Then
            txtSerialData.Visible = True
        Else
            txtSerialData.Visible = False
        End If
    End If
End Sub

Public Sub SUB_CHK_PDF_TIF(ArgPathFileName As String)
    If InStr(UCase(ArgPathFileName), ".TIF") > 0 Or InStr(UCase(ArgPathFileName), ".TIFF") > 0 Or InStr(UCase(ArgPathFileName), ".PDF") > 0 Then
        imvResult.LoadMultiPage ArgPathFileName, 1
        
        txtMultiPno = "1"
        lblMultiCnt = str(imvResult.GetTotalPage)
        
        picTifPdf.Visible = True
        picJPG.Visible = False
    Else
        lblMultiCnt = ""
        
        picTifPdf.Visible = False
        picJPG.Visible = True
    End If
End Sub

Public Function FUNC_HIS_RST_UPDATE() As Boolean
    
    FUNC_HIS_RST_UPDATE = False
    
    Select Case gstrHOS_CUSCD
        Case 1 '/1.��õ�Ƿ��
            Select Case gtypEQ_INFO.QUERYTYPE
                Case "1" '/3����
                    gstrQuy = "UPDATE SY_MEODPRSC SET "
                    gstrQuy = gstrQuy & vbCrLf & "       CDIS_YN            = 'Y' "
                    gstrQuy = gstrQuy & vbCrLf & " WHERE PID                = '" & lbl���Ϲ�ȣ & "' "
                    gstrQuy = gstrQuy & vbCrLf & "   AND PRSC_DATE          = TO_DATE('" & dtp��������.Value & "','YYYY-MM-DD') "
                    gstrQuy = gstrQuy & vbCrLf & "   AND PRSC_SQNO          = '" & lbló��SEQ & "' "
                    If RunSQL(gstrQuy) = False Then Exit Function
        
                Case "3" '/���հ���
                    '/���հ����� ó������� �ϷῩ�θ� SETTING �� ���� ����.
        
                Case Else
                    gstrQuy = "UPDATE SY_MEODPRSC SET "
                    gstrQuy = gstrQuy & vbCrLf & "       CDIS_YN            = 'Y', "
                    gstrQuy = gstrQuy & vbCrLf & "       PRSC_PRGR_STAT_CD  ='C' "
                    gstrQuy = gstrQuy & vbCrLf & " WHERE PID                = '" & lbl���Ϲ�ȣ & "' "
                    gstrQuy = gstrQuy & vbCrLf & "   AND PRSC_DATE          = TO_DATE('" & dtp��������.Value & "','YYYY-MM-DD') "
                    gstrQuy = gstrQuy & vbCrLf & "   AND PRSC_SQNO          = '" & lbló��SEQ & "' "
                    If RunSQL(gstrQuy) = False Then Exit Function
            End Select
            
        Case 2 '/2.����ø����Ϻ���
            '/��ó��:   TMPRSCINFN
            gstrQuy = "UPDATE TMPRSCINFN SET "
            gstrQuy = gstrQuy & vbCrLf & "      PRSC_STAT_CD    = '560', "                      '/560.�ӽð�� �̻�
            gstrQuy = gstrQuy & vbCrLf & "      UPDTR_ID        = '" & gtypUSER.USERID & "', "  '/���������� ID
            gstrQuy = gstrQuy & vbCrLf & "      UPDT_DT         = SYSTIMESTAMP "                '/������������
            gstrQuy = gstrQuy & vbCrLf & "WHERE PRSC_DATE       = '" & Format(CDate(dtp��������.Value), "YYYYMMDD") & "' "  '/ó������
            gstrQuy = gstrQuy & vbCrLf & "  AND PRSC_NO         =  " & lbló��SEQ & " "                                     '/ó���ȣ
            gstrQuy = gstrQuy & vbCrLf & "  AND PID             = '" & lbl���Ϲ�ȣ & "' "                                   '/���Ϲ�ȣ
            gstrQuy = gstrQuy & vbCrLf & "  AND PRSC_CD         = '" & lbló���ڵ� & "' "                                   '/ó���ڵ�
            gstrQuy = gstrQuy & vbCrLf & "  AND PRSC_VALD_YN    = 'Y' "                                                     '/��ó�� ����ִ� ó��
            gstrQuy = gstrQuy & vbCrLf & "  AND PRSC_HSTR_CD    = 'O' "                                                     '/ó��History ��ȣ
            If RunSQL(gstrQuy) = False Then Exit Function

            '/�ǽ�ó��: TMPRSCEXCN
            gstrQuy = "UPDATE TMPRSCEXCN SET "
            gstrQuy = gstrQuy & vbCrLf & "      CNDT_PRSC_STAT_CD   = '560', "                      '/560.�ӽð�� �̻�
            gstrQuy = gstrQuy & vbCrLf & "      UPDTR_ID            = '" & gtypUSER.USERID & "', "  '/���������� ID
            gstrQuy = gstrQuy & vbCrLf & "      UPDT_DT             = SYSTIMESTAMP "                '/������������
            gstrQuy = gstrQuy & vbCrLf & "WHERE PRSC_DATE           = '" & Format(CDate(dtp��������.Value), "YYYYMMDD") & "' "  '/ó������
            gstrQuy = gstrQuy & vbCrLf & "  AND PRSC_NO             =  " & lbló��SEQ & " "                                     '/ó���ȣ
            gstrQuy = gstrQuy & vbCrLf & "  AND PID                 = '" & lbl���Ϲ�ȣ & "' "                                   '/���Ϲ�ȣ
            gstrQuy = gstrQuy & vbCrLf & "  AND MEFE_CD             = '" & lbló���ڵ� & "' "                                   '/ó���ڵ�
            gstrQuy = gstrQuy & vbCrLf & "  AND CNDT_PRSC_VALD_YN   = 'Y' "                                                     '/�ǽ�ó�� ����ִ� ó��
            If RunSQL(gstrQuy) = False Then Exit Function
  
        Case Else
            MsgBox "������� ó���� ���� HIS ������ �����ϴ�!", vbCritical, "���"
    End Select
    
    FUNC_HIS_RST_UPDATE = True
End Function

Public Function FUNC_HIS_ORDER1_VIEW(argOrderCode As String) As Boolean '/������ ó����ȸ(����ø����Ϻ���)
    FUNC_HIS_ORDER1_VIEW = False
    
    gstrQuy = " SELECT "
    gstrQuy = gstrQuy & vbCrLf & "       A.PID AS CHRTNO, "                                                            '/���Ϲ�ȣ
    gstrQuy = gstrQuy & vbCrLf & "       DECODE(A.PRSC_OCRR_DVCD, 'I', '�Կ�', 'O', '�ܷ�', '��Ÿ') AS IO_SECTION, "   '/�ܷ�/�Կ�����(I:�Կ�, O:�ܷ�)
    gstrQuy = gstrQuy & vbCrLf & "       C.DEPT_ENGL_ABNM AS DETPCD, "                                                 '/����� ��Ī(����)
    gstrQuy = gstrQuy & vbCrLf & "       B.PT_NM AS PATNM, "                                                           '/�����ڸ�
    gstrQuy = gstrQuy & vbCrLf & "       B.SEX_CD AS SEX, "                                                            '/����
    gstrQuy = gstrQuy & vbCrLf & "       B.RESD_NO_1 AS JUMIN1, "                                                      '/�ֹι�ȣ1
    gstrQuy = gstrQuy & vbCrLf & "       B.RESD_NO_2 AS JUMIN2, "                                                      '/�ֹι�ȣ2
    gstrQuy = gstrQuy & vbCrLf & "       fn_PaGetAge(B.RESD_NO_1, B.RESD_NO_2, B.DOBR, A.PRSC_DATE) AS AGE, "          '/HIS ���̰�� �Լ�
    gstrQuy = gstrQuy & vbCrLf & "       A.PRSC_DATE AS ORDDATE, "                                                     '/ó������
    gstrQuy = gstrQuy & vbCrLf & "       A.PRSC_NO AS ORDSEQ, "                                                        '/ó���ȣ
    gstrQuy = gstrQuy & vbCrLf & "       A.PRSC_CD AS ORDCD, "                                                         '/ó���ڵ�
    gstrQuy = gstrQuy & vbCrLf & "       A.PRSC_NM AS ORDNM, "                                                         '/ó���
    gstrQuy = gstrQuy & vbCrLf & "       A.DLVR_MATR, "                                                                '/���޻���
    gstrQuy = gstrQuy & vbCrLf & "       A.SUPT_DEPT_DLVR_MATR, "                                                      '/�����μ� ���޻���
    gstrQuy = gstrQuy & vbCrLf & "       A.CNDT_PRSC_STAT_CD "                                                         '/�ǽ�ó�� ó���������Flag
    gstrQuy = gstrQuy & vbCrLf & "  FROM VPRSCINFN A, TPAPTMASTN B, TZDEPTMSTN C "                                     '/VPRSCINFN(ó����ȸ VIEW), TPAPTMASTN(ȯ�ڸ�����), TZDEPTMSTN(�μ�������)
    gstrQuy = gstrQuy & vbCrLf & " WHERE A.PID                 = B.PID "
    gstrQuy = gstrQuy & vbCrLf & "   AND A.MDCR_DPMT_CD        = C.DEPT_CD "
'''    gstrQuy = gstrQuy & vbCrLf & "   AND A.PRSC_DATE           = '" & Format(CDate(dtp��������.Value), "YYYYMMDD") & "' "            '/ó������
    gstrQuy = gstrQuy & vbCrLf & "   AND A.CNDT_DATE           = '" & Format(CDate(dtp��������.Value), "YYYYMMDD") & "' "            '/�˻�� ��������
    gstrQuy = gstrQuy & vbCrLf & "   AND A.PRSC_VALD_YN        = 'Y' "                                                 '/��ó�� ����ִ� ó��
    gstrQuy = gstrQuy & vbCrLf & "   AND A.CNDT_PRSC_VALD_YN   = 'Y' "                                                 '/�ǽ�ó�� ����ִ� ó��
    gstrQuy = gstrQuy & vbCrLf & "   AND A.PRSC_HSTR_CD        = 'O' "                                                 '/ó��History ��ȣ
    gstrQuy = gstrQuy & vbCrLf & "   AND A.CNDT_PRSC_STAT_CD   = '440' "                                               '/�ǽ�ó�� ó���������Flag(440.����)
    gstrQuy = gstrQuy & vbCrLf & "   AND A.CNDT_DATE          <> '00000000' "                                          '/������������(������:00000000, ����:��ȿ����)
    gstrQuy = gstrQuy & vbCrLf & "   AND A.PRSC_CD            IN (" & argOrderCode & ") "                              '/ó���ڵ�
    
    FUNC_HIS_ORDER1_VIEW = True
End Function

Public Function FUNC_HIS_ORDER2_VIEW(argOrderCode As String) As Boolean '/������ ó����ȸ(����ø����Ϻ���)
    FUNC_HIS_ORDER2_VIEW = False
    
    gstrQuy = " SELECT "
    gstrQuy = gstrQuy & vbCrLf & "       A.PID AS CHRTNO, "                                                             '/���Ϲ�ȣ
    gstrQuy = gstrQuy & vbCrLf & "       DECODE(A.PRSC_OCRR_DVCD, 'I', '�Կ�', 'O', '�ܷ�', '��Ÿ') AS IO_SECTION, "    '/�ܷ�/�Կ�����(I:�Կ�, O:�ܷ�)
    gstrQuy = gstrQuy & vbCrLf & "       C.DEPT_ENGL_ABNM AS DETPCD, "                                                  '/����� ��Ī(����)
    gstrQuy = gstrQuy & vbCrLf & "       B.PT_NM AS PATNM, "                                                            '/�����ڸ�
    gstrQuy = gstrQuy & vbCrLf & "       B.SEX_CD AS SEX, "                                                             '/����
    gstrQuy = gstrQuy & vbCrLf & "       B.RESD_NO_1 AS JUMIN1, "                                                       '/�ֹι�ȣ1
    gstrQuy = gstrQuy & vbCrLf & "       B.RESD_NO_2 AS JUMIN2, "                                                       '/�ֹι�ȣ2
    gstrQuy = gstrQuy & vbCrLf & "       fn_PaGetAge(B.RESD_NO_1, B.RESD_NO_2, B.DOBR, A.PRSC_DATE) AS AGE, "           '/HIS ���̰�� �Լ�
    gstrQuy = gstrQuy & vbCrLf & "       A.PRSC_DATE AS ORDDATE, "                                                      '/ó������
    gstrQuy = gstrQuy & vbCrLf & "       A.PRSC_NO AS ORDSEQ, "                                                         '/ó���ȣ
    gstrQuy = gstrQuy & vbCrLf & "       A.PRSC_CD AS ORDCD, "                                                          '/ó���ڵ�
    gstrQuy = gstrQuy & vbCrLf & "       A.PRSC_NM AS ORDNM, "                                                          '/ó���
    gstrQuy = gstrQuy & vbCrLf & "       A.DLVR_MATR, "                                                                 '/���޻���
    gstrQuy = gstrQuy & vbCrLf & "       A.SUPT_DEPT_DLVR_MATR, "                                                       '/�����μ� ���޻���
    gstrQuy = gstrQuy & vbCrLf & "       D.EXAMDATE, "                                                                  '/�����������
    gstrQuy = gstrQuy & vbCrLf & "       A.CNDT_PRSC_STAT_CD "                                                          '/�ǽ�ó�� ó���������Flag
    gstrQuy = gstrQuy & vbCrLf & "  FROM VPRSCINFN A, TPAPTMASTN B, TZDEPTMSTN C, MM_EMR_RES D "                        '/VPRSCINFN(ó����ȸ VIEW), TPAPTMASTN(ȯ�ڸ�����), TZDEPTMSTN(�μ�������)
    gstrQuy = gstrQuy & vbCrLf & " WHERE A.PID                 = B.PID "
    gstrQuy = gstrQuy & vbCrLf & "   AND A.MDCR_DPMT_CD        = C.DEPT_CD "
    gstrQuy = gstrQuy & vbCrLf & "   AND A.PRSC_DATE           = D.ORDDATE "
    gstrQuy = gstrQuy & vbCrLf & "   AND A.PRSC_NO             = D.ORDSEQ "
'''    gstrQuy = gstrQuy & vbCrLf & "   AND A.PRSC_DATE           = '" & Format(CDate(dtp��������.Value), "YYYYMMDD") & "' "            '/ó������
    gstrQuy = gstrQuy & vbCrLf & "   AND A.CNDT_DATE           = '" & Format(CDate(dtp��������.Value), "YYYYMMDD") & "' "            '/�˻�� ��������
    gstrQuy = gstrQuy & vbCrLf & "   AND A.PRSC_VALD_YN        = 'Y' "                                                 '/��ó�� ����ִ� ó��
    gstrQuy = gstrQuy & vbCrLf & "   AND A.CNDT_PRSC_VALD_YN   = 'Y' "                                                 '/�ǽ�ó�� ����ִ� ó��
    gstrQuy = gstrQuy & vbCrLf & "   AND A.PRSC_HSTR_CD        = 'O' "                                                 '/ó��History ��ȣ
    gstrQuy = gstrQuy & vbCrLf & "   AND A.CNDT_PRSC_STAT_CD  >  '440' "                                               '/�ǽ�ó�� ó���������Flag(560.�ӽð�� �̻�)
    gstrQuy = gstrQuy & vbCrLf & "   AND A.CNDT_DATE          <> '00000000' "                                          '/������������(������:00000000, ����:��ȿ����)
    gstrQuy = gstrQuy & vbCrLf & "   AND A.PRSC_CD            IN (" & argOrderCode & ") "                              '/ó���ڵ�
    
    FUNC_HIS_ORDER2_VIEW = True
End Function

