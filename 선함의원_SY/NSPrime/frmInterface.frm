VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInterface 
   Caption         =   "NSPrime"
   ClientHeight    =   10095
   ClientLeft      =   405
   ClientTop       =   900
   ClientWidth     =   12930
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
   ScaleHeight     =   15315
   ScaleWidth      =   28560
   StartUpPosition =   1  '������ ���
   WindowState     =   2  '�ִ�ȭ
   Begin VB.Frame FrmHideControl 
      Caption         =   "HideControl"
      Height          =   9945
      Left            =   20730
      TabIndex        =   1
      Top             =   2340
      Visible         =   0   'False
      Width           =   10935
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
         Left            =   6180
         TabIndex        =   69
         Top             =   4320
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.Frame Frame6 
         Height          =   585
         Left            =   5460
         TabIndex        =   61
         Top             =   1170
         Visible         =   0   'False
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
            TabIndex        =   63
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
            TabIndex        =   62
            Top             =   180
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label Label3 
            BackColor       =   &H80000008&
            ForeColor       =   &H8000000E&
            Height          =   315
            Left            =   180
            TabIndex        =   68
            Top             =   720
            Width           =   1155
         End
         Begin VB.Label lblPname 
            Caption         =   "1234567890ab"
            Height          =   225
            Index           =   0
            Left            =   5130
            TabIndex        =   67
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
            TabIndex        =   66
            Top             =   240
            Width           =   945
         End
         Begin VB.Label lblBarcode 
            Caption         =   "12345"
            Height          =   165
            Index           =   0
            Left            =   1905
            TabIndex        =   65
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
            TabIndex        =   64
            Top             =   240
            Width           =   1380
         End
      End
      Begin VB.CommandButton cmdOrder 
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   6210
         TabIndex        =   60
         Top             =   2220
         Width           =   1395
      End
      Begin VB.CommandButton cmdResult 
         Appearance      =   0  '���
         Caption         =   "����ޱ�"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   7680
         TabIndex        =   59
         Top             =   2220
         Width           =   1395
      End
      Begin VB.CommandButton cmdExcelExport 
         Caption         =   "Excel"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   10590
         TabIndex        =   58
         Top             =   2220
         Width           =   1395
      End
      Begin VB.CommandButton cmdPatDelete 
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9270
         TabIndex        =   53
         Top             =   510
         Width           =   1425
      End
      Begin VB.TextBox txtStartNum 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   5220
         TabIndex        =   52
         Top             =   570
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox txtStopNum 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   5940
         TabIndex        =   51
         Top             =   570
         Visible         =   0   'False
         Width           =   555
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
         Left            =   6600
         List            =   "frmInterface.frx":1502
         TabIndex        =   50
         Top             =   570
         Visible         =   0   'False
         Width           =   975
      End
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
         TabIndex        =   38
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
         TabIndex        =   37
         Top             =   5160
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.Frame Frame2 
         Caption         =   "Error Log"
         Height          =   945
         Left            =   180
         TabIndex        =   34
         Top             =   8190
         Width           =   4530
         Begin VB.TextBox txtErrLog 
            Appearance      =   0  '���
            Height          =   615
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  '����
            TabIndex        =   35
            Top             =   240
            Width           =   4275
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Print"
         Height          =   2415
         Left            =   180
         TabIndex        =   31
         Top             =   5670
         Width           =   3045
         Begin FPSpread.vaSpread vasPrint 
            Height          =   1035
            Left            =   120
            TabIndex        =   32
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
            SpreadDesigner  =   "frmInterface.frx":151A
         End
         Begin FPSpread.vaSpread vasPrintBuf 
            Height          =   975
            Left            =   120
            TabIndex        =   33
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
            SpreadDesigner  =   "frmInterface.frx":2FA1
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
         Height          =   915
         Left            =   120
         TabIndex        =   15
         Top             =   2250
         Width           =   1695
         _Version        =   393216
         _ExtentX        =   2990
         _ExtentY        =   1614
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
         SpreadDesigner  =   "frmInterface.frx":31C7
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
         SpreadDesigner  =   "frmInterface.frx":33ED
      End
      Begin VB.PictureBox picLogin 
         Appearance      =   0  '���
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1530
         Picture         =   "frmInterface.frx":3613
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
            Left            =   1230
            Top             =   210
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
                  Picture         =   "frmInterface.frx":3B9D
                  Key             =   "RUN"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":4137
                  Key             =   "NOT"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":46D1
                  Key             =   "STOP"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":4C6B
                  Key             =   "LST"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":54FD
                  Key             =   "ITM"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":5657
                  Key             =   "ERR"
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":57B1
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
         SpreadDesigner  =   "frmInterface.frx":590B
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
         SpreadDesigner  =   "frmInterface.frx":5B31
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
         SpreadDesigner  =   "frmInterface.frx":5D57
      End
      Begin VB.Label Label2 
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
         Height          =   195
         Left            =   5790
         TabIndex        =   54
         Top             =   660
         Width           =   165
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
         TabIndex        =   41
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
         TabIndex        =   40
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
         TabIndex        =   39
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
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   120
      TabIndex        =   44
      Top             =   570
      Width           =   12705
      Begin VB.CommandButton cmdSearch 
         Caption         =   "ȯ����ȸ"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   6450
         TabIndex        =   70
         Top             =   180
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.CommandButton cmdRsltSearch 
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
         Height          =   435
         Left            =   4740
         TabIndex        =   57
         Top             =   180
         Width           =   1395
      End
      Begin VB.CommandButton cmdIFClear 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   11040
         TabIndex        =   56
         Top             =   180
         Width           =   1395
      End
      Begin VB.CommandButton cmdIFTrans 
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   9600
         TabIndex        =   55
         Top             =   180
         Width           =   1395
      End
      Begin MSComCtl2.DTPicker dtpStopDt 
         Height          =   345
         Left            =   3060
         TabIndex        =   45
         Top             =   210
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   131727361
         CurrentDate     =   40248
      End
      Begin MSComCtl2.DTPicker dtpStartDt 
         Height          =   345
         Left            =   1260
         TabIndex        =   46
         Top             =   210
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   131727361
         CurrentDate     =   40248
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
         Left            =   210
         TabIndex        =   48
         Top             =   270
         Width           =   915
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
         Left            =   2880
         TabIndex        =   47
         Top             =   300
         Width           =   105
      End
   End
   Begin VB.CommandButton Command16 
      Caption         =   "�����׽�Ʈ"
      Height          =   435
      Left            =   20940
      TabIndex        =   43
      Top             =   1470
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtTest 
      Height          =   1785
      Left            =   20970
      MultiLine       =   -1  'True
      TabIndex        =   42
      Top             =   10080
      Visible         =   0   'False
      Width           =   3225
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  '�� ����
      Height          =   810
      Left            =   0
      ScaleHeight     =   750
      ScaleWidth      =   28500
      TabIndex        =   30
      Top             =   525
      Width           =   28560
   End
   Begin VB.Frame Frame1 
      Height          =   8205
      Left            =   30
      TabIndex        =   26
      Top             =   1380
      Width           =   20835
      Begin VB.CommandButton cmdSL 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   210
         Width           =   495
      End
      Begin VB.CheckBox chkWAll 
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   690
         TabIndex        =   27
         Top             =   270
         Width           =   225
      End
      Begin FPSpread.vaSpread vasID 
         Height          =   7845
         Left            =   90
         TabIndex        =   29
         Top             =   180
         Width           =   11055
         _Version        =   393216
         _ExtentX        =   19500
         _ExtentY        =   13838
         _StockProps     =   64
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
         SpreadDesigner  =   "frmInterface.frx":5F7D
      End
      Begin FPSpread.vaSpread vasRes 
         Height          =   7845
         Left            =   11220
         TabIndex        =   28
         Top             =   180
         Width           =   6495
         _Version        =   393216
         _ExtentX        =   11456
         _ExtentY        =   13838
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
         SpreadDesigner  =   "frmInterface.frx":6B5E
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  '�� ����
      BackColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   28500
      TabIndex        =   18
      Top             =   0
      Width           =   28560
      Begin MSComCtl2.DTPicker dtpToday 
         Height          =   315
         Left            =   1110
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
         Format          =   131727360
         CurrentDate     =   40457
      End
      Begin VB.Label Label1 
         Appearance      =   0  '���
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '����
         Caption         =   "NSPrime"
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
         Left            =   3840
         TabIndex        =   36
         Top             =   90
         Width           =   1125
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
         Left            =   180
         TabIndex        =   25
         Top             =   150
         Width           =   780
      End
      Begin VB.Label Label1 
         Appearance      =   0  '���
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '����
         Caption         =   "NSPrime"
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
         Left            =   3870
         TabIndex        =   22
         Top             =   120
         Width           =   1125
      End
      Begin VB.Image imgPort 
         Height          =   240
         Left            =   10110
         Picture         =   "frmInterface.frx":71F6
         Top             =   150
         Width           =   240
      End
      Begin VB.Image imgSend 
         Height          =   240
         Left            =   11265
         Picture         =   "frmInterface.frx":7780
         Top             =   150
         Width           =   240
      End
      Begin VB.Image imgReceive 
         Height          =   240
         Left            =   12420
         Picture         =   "frmInterface.frx":7D0A
         Top             =   150
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "��Ʈ"
         Height          =   195
         Index           =   0
         Left            =   9600
         TabIndex        =   21
         Top             =   180
         Width           =   420
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "�۽�"
         Height          =   195
         Left            =   10785
         TabIndex        =   20
         Top             =   180
         Width           =   420
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "����"
         Height          =   195
         Left            =   11910
         TabIndex        =   19
         Top             =   180
         Width           =   420
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '�Ʒ� ����
      Height          =   405
      Left            =   0
      TabIndex        =   49
      Top             =   14910
      Width           =   28560
      _ExtentX        =   50377
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
            Object.Width           =   12347
            MinWidth        =   12347
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   2646
            MinWidth        =   2646
            TextSave        =   "2017-06-21"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2646
            MinWidth        =   2646
            TextSave        =   "���� 2:58"
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
         Visible         =   0   'False
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


Dim strRecvData()   As String
Dim intPhase        As Integer
Dim strState        As String
Dim intBufCnt       As Integer
Dim blnIsETB        As Boolean
Dim intSndPhase     As Integer
Dim intFrameNo      As Integer
'===============================

Dim ReceiveData As String


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
        MsgBox "������ �ڷᰡ �����ϴ�.", vbCritical + vbOKOnly, Me.Caption
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
    StatusBar1.Panels(3).Text = ""
    
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
    Dim intRow      As Integer
    Dim STM         As ADODB.Stream
    Dim blnSendXml  As Boolean
    Dim strHeader   As String
    Dim strBody     As String
    Dim strFileNm   As String
    Dim strAssayNm  As String
    
    blnSendXml = False
    
    strHeader = "<?xml version=""1.0"" encoding=""utf-8""?>" & vbCrLf
    strHeader = strHeader & "<GROUP NAME=""WORKLIST"" TYPE=""HOST"" VERSION=""00.00.00.00"">" & vbCrLf
    strBody = ""
    
    With vasID
        For intRow = 1 To .MaxRows
            .Row = intRow
            .Col = colCheckBox
            If .Value = 1 Then
                If Trim(GetText(vasID, intRow, colBARCODE)) <> "" Then
                    blnSendXml = True
                    '-- ������ XML
'''                    strBody = strBody & vbTab & "<GROUP TYPE=""Sample"">" & vbCrLf
'''                    strBody = strBody & vbTab & vbTab & "<GROUP TYPE=""Patient"" ID=""" & Trim(GetText(vasID, intRow, colBARCODE)) & """>" & vbCrLf
'''                    strBody = strBody & vbTab & vbTab & vbTab & "<PARAM TYPE=""String[32]"" ID=""Name"" FILTER=""-rrrr"">" & Trim(GetText(vasID, intRow, colPNAME)) & "</PARAM>" & vbCrLf
'''                    strBody = strBody & vbTab & vbTab & "</GROUP>" & vbCrLf
'''                    strBody = strBody & vbTab & vbTab & "<GROUP TYPE=""Assays"" ID="""">" & vbCrLf
'''                    Select Case Trim(GetText(vasID, intRow, colINOUT))
'''                        Case "INHALANT":    strAssayNm = gAssayNM.INHALANT
'''                        Case "FOOD":        strAssayNm = gAssayNM.FOOD
'''                        Case "ATOPY":       strAssayNm = gAssayNM.ATOPY
'''                    End Select
'''                    strBody = strBody & vbTab & vbTab & vbTab & "<GROUP TYPE=""Assay"" ID=""" & strAssayNm & """ STATE=""Host"" />" & vbCrLf
'''                    strBody = strBody & vbTab & vbTab & "</GROUP>" & vbCrLf
'''                    strBody = strBody & vbTab & "</GROUP>" & vbCrLf
                    
'<PARAM TYPE="String[32]" ID="Group" FILTER="-wwwr">����</PARAM>
'<PARAM TYPE="String[32]" ID="Name" FILTER="-ooor">��</PARAM>
'<PARAM TYPE="String[32]" ID="Surname" FILTER="-ooor">��</PARAM>
'<PARAM TYPE="Date" ID="BirthDate" FILTER="-ooor">2016-03-08 ���� 10:57:52</PARAM>
'<PARAM TYPE="SexType" ID="Sex" FILTER="-wwwr">M</PARAM>
'<PARAM TYPE="String[50]" ID="Note" FILTER="-wwwr">123456</PARAM>
'<PARAM TYPE="String[25]" ID="Code" FILTER="-ooor">111</PARAM>
                    
                    
                    strBody = strBody & vbTab & "<GROUP TYPE=""Sample"">" & vbCrLf
                    strBody = strBody & vbTab & vbTab & "<GROUP TYPE=""Patient"" ID=""" & Trim(GetText(vasID, intRow, colBARCODE)) & """>" & vbCrLf
                    '-- Group
                    strBody = strBody & vbTab & vbTab & vbTab & "<PARAM TYPE=""String[32]"" ID=""Group"" FILTER=""-wwwr"">" & Trim(GetText(vasID, intRow, colINOUT)) & "</PARAM>" & vbCrLf
                    '-- Name
                    strBody = strBody & vbTab & vbTab & vbTab & "<PARAM TYPE=""String[32]"" ID=""Name"" FILTER=""-ooor"">" & Trim(GetText(vasID, intRow, colPNAME)) & "</PARAM>" & vbCrLf
                    '-- Surname
                    strBody = strBody & vbTab & vbTab & vbTab & "<PARAM TYPE=""String[32]"" ID=""Surname"" FILTER=""-ooor"">" & Trim(GetText(vasID, intRow, colPNAME)) & "</PARAM>" & vbCrLf
                    '-- BirthDate
                    'strBody = strBody & vbTab & vbTab & vbTab & "<PARAM TYPE=""Date"" ID=""BirthDate"" FILTER=""-ooor"">" & Format(Trim(GetText(vasID, intRow, colPOSNO)), "yyyy-mm-dd") & "</PARAM>" & vbCrLf
                    strBody = strBody & vbTab & vbTab & vbTab & "<PARAM TYPE=""Date"" ID=""BirthDate"" FILTER=""-ooor"">" & Format(Now, "yyyy-mm-dd") & "</PARAM>" & vbCrLf
                    '-- Sex
                    strBody = strBody & vbTab & vbTab & vbTab & "<PARAM TYPE=""SexType"" ID=""Sex"" FILTER=""-wwwr"">" & IIf(Trim(GetText(vasID, intRow, colPSEX)) = "M", "M", "F") & "</PARAM>" & vbCrLf
                    '-- Note
                    'strBody = strBody & vbTab & vbTab & vbTab & "<PARAM TYPE=""String[50]"" ID=""Note"" FILTER=""-wwwr"">" & Trim(GetText(vasID, intRow, colDISKNO)) & "</PARAM>" & vbCrLf
                    strBody = strBody & vbTab & vbTab & vbTab & "<PARAM TYPE=""String[50]"" ID=""Note"" FILTER=""-wwwr"">" & "Note Test" & "</PARAM>" & vbCrLf
                    '-- Code (��񿡼� Patient ID)
                    strBody = strBody & vbTab & vbTab & vbTab & "<PARAM TYPE=""String[50]"" ID=""Note"" FILTER=""-wwwr"">" & Trim(GetText(vasID, intRow, colPID)) & "</PARAM>" & vbCrLf
                    strBody = strBody & vbTab & vbTab & vbTab & "<PARAM TYPE=""String[25]"" ID=""Code"" FILTER=""-ooor"">" & "Code Test" & "</PARAM>" & vbCrLf
                    
                    strBody = strBody & vbTab & vbTab & "</GROUP>" & vbCrLf
                    strBody = strBody & vbTab & vbTab & "<GROUP TYPE=""Assays"" ID="""">" & vbCrLf
                    Select Case Trim(GetText(vasID, intRow, colINOUT))
                        Case "INHALANT":    strAssayNm = gAssayNM.INHALANT
                        Case "FOOD":        strAssayNm = gAssayNM.FOOD
                        Case "ATOPY":       strAssayNm = gAssayNM.ATOPY
                    End Select
                    strBody = strBody & vbTab & vbTab & vbTab & "<GROUP TYPE=""Assay"" ID=""" & strAssayNm & """ STATE=""Host"" />" & vbCrLf
                    strBody = strBody & vbTab & vbTab & "</GROUP>" & vbCrLf
                    strBody = strBody & vbTab & "</GROUP>" & vbCrLf
                    
                    .Row = intRow
                    .Col = colCheckBox
                    .Value = "0"
                    
                    Call SetText(vasID, "Order", intRow, colState)
                End If
            End If
        Next
    End With
    
    If blnSendXml = True Then
    
        '## ������ ������ ������ ����
        strFileNm = gAssayNM.OrderPath & "\HostIn.xml"
    
        If Dir$(strFileNm, vbNormal) <> "" Then
            Kill strFileNm
        End If
        
         '## ���Ͽ���
        Set STM = New ADODB.Stream
        
        STM.Open
        STM.Type = adTypeText
        STM.Charset = "utf-8"
        STM.WriteText strHeader & strBody & "</GROUP>" & vbCrLf
                    
        STM.SaveToFile strFileNm, adSaveCreateNotExist
        STM.Close
        Set STM = Nothing
        
    End If
    
End Sub


'''Private Sub MakeXML(ByVal vasIDRow As Integer)
''''��������
'''    'Dim vasIDRow As Integer
'''    Dim vasResRow As Integer
'''    Dim iRow As Integer
'''    Dim liRet As Integer
'''    Dim FindFile As String
'''
''''    If MsgBox(" " & vbCrLf & "���������� �Ͻðڽ��ϱ�?" & vbCrLf & " ", vbInformation + vbOKCancel, "�˸�:��������") = vbCancel Then
''''        Exit Sub
''''    End If
'''
'''    With frmInterface
'''        For vasResRow = 1 To .vasTemp.DataRowCnt
'''            .vasID.Row = vasIDRow
'''            .vasID.Col = 1
'''            If .vasID.Value = 1 Then
'''                .vasTemp.Row = vasResRow
'''                liRet = -1
'''                If Trim(GetText(.vasTemp, vasResRow, 3)) <> "" Then
'''                    liRet = Make_XML(vasResRow)
'''                End If
'''
'''                If liRet = 1 Then
'''                    SetBackColor .vasID, vasIDRow, vasIDRow, colCheckBox, 12, 202, 255, 112
'''                    'SetText vasList, "���ۿϷ�", vasIDRow, colState
'''
'''                    FindFile = Dir("C:\UBCare\SINAI\IF\ExamIF_In.xml")
'''                    If FindFile <> "" Then
'''                        Kill "C:\UBCare\SINAI\IF\ExamIF_In.xml"     '���ۿϷᰡ ������ ���������
'''                    End If
'''
'''                          SQL = " Update pat_res Set "
'''                    SQL = SQL & " TransYN = '2', "
'''                    SQL = SQL & " TransDt = '" & Format(Now, "yyyymmdd") & "' "
'''                    .vasID.Row = vasIDRow: .vasID.Col = 4
'''                    SQL = SQL & " Where ChartNo  = '" & Trim(.vasID.Text) & "' "
'''                    .vasID.Row = vasIDRow: .vasID.Col = 12
'''                    SQL = SQL & "   and ExamID   = '" & Trim(.vasID.Text) & "' "
'''                    .vasID.Row = vasIDRow: .vasID.Col = 10
'''                    SQL = SQL & "   and CommDate = '" & Trim(.vasID.Text) & "'"
'''                    Res = SendQuery(gLocal, SQL)
'''
'''                Else
'''                    SetBackColor .vasID, vasIDRow, vasIDRow, colCheckBox, 12, 255, 0, 0
'''                    'SetText vasID, "����", vasIDRow, colState
'''                End If
'''                '.vasID.Col = 1
'''                '.vasID.Value = "0"
'''            Else
'''
'''            End If
'''        Next
'''    End With
'''
'''    If XmlTxtHead = "" Then
'''        '<?xml version="1.0" encoding="utf-8"?>
'''        '<GROUP NAME="WORKLIST" TYPE="HOST" VERSION="00.00.00.00">
'''
'''        XmlTxtHead = "<?xml version=""1.0"" encoding=""euc-kr""?>" & vbCrLf & _
'''                     "<?xml-stylesheet type=""text/xsl"" href=C:\UBCare\SINAI\IF\Form\ExamIF_Form_05.xsl""?>" & vbCrLf & "<UBCare�˻�����>"
'''    End If
'''
'''    If XmlTxtTail = "" Then
'''        XmlTxtTail = "</UBCare�˻�����>"
'''    End If
'''
''''    XMLAllTxt = XmlTxtHead & XMLAllTxt & XmlTxtTail
'''    SaveXMLFile XMLAllTxt
'''
'''End Sub


'Public Sub SaveXMLFile(argSQL As String, Optional argFlag As Integer = 0)
''argSQL�� ������ ���Ϸ� ����
'    Dim FilNum, FilNum1
'    Dim FindFile As String
'    Dim TxtString1 As String
'    Dim AllString1 As String
'    Dim i As Long
'
'    FindFile = Dir("C:\UBCare\SINAI\IF\ExamIF_Out.xml")
'
'
'    If FindFile <> "" Then
''        Kill "C:\UBCare\SINAI\IF\ExamIF_Out.xml"
'        FilNum1 = FreeFile
'        Open "C:\UBCare\SINAI\IF\ExamIF_out.xml" For Input As FilNum1
'
'        Do While Not EOF(FilNum1)
'            Input #FilNum1, TxtString1
'            Line Input #FilNum1, TxtString1
'            AllString1 = AllString1 & TxtString1
'        Loop
'
'        Close #FilNum1
'        i = InStr(1, AllString1, "</UBCare�˻�����>")
'        XmlBody = Mid(AllString1, 1, i - 1)
'        argSQL = XmlBody & argSQL & XmlTxtTail
'        Kill "C:\UBCare\SINAI\IF\ExamIF_Out.xml"
'    Else
'        argSQL = XmlTxtHead & argSQL & XmlTxtTail
'    End If
'
''    XMLAllTxt = XmlTxtHead & XMLAllTxt & XmlTxtTail
'
'    FilNum = FreeFile
'
'
'    If argFlag = 0 Then
'        Open "C:\UBCare\SINAI\IF\ExamIF_Out.xml" For Output As FilNum
'    Else
'        Open "C:\UBCare\SINAI\IF\ExamIF_Out.xml" For Append As FilNum
'    End If
'    Print #FilNum, argSQL
'    Close FilNum
'    argSQL = ""
'End Sub

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


Private Sub cmdResult_Click()
    Dim varXML      As Variant
    Dim strXmlName  As String
    Dim i As Integer
    
    StatusBar1.Panels(3).Text = ""

    ' gAssayNM.OrderPath
    CommonDialog1.Filter = "XML(*.xml)|*.xml"
    CommonDialog1.Action = 1
    
    If CommonDialog1.FileTitle = "" Then
        Exit Sub
    End If
    
    strXmlName = Trim(CommonDialog1.Filename)
    
    Call f_subSet_XMLWorkList(strXmlName)
    
'    Call EditRcvDataAPEX_Front
    
    Call EditRcvDataAPEX
    
End Sub


Private Function f_subSet_XMLWorkList(ByVal strXML As String) As Variant
    Dim strPath   As String
    Dim strBuffer As String
    Dim i         As Long
    Dim lngBufLen As Long
    Dim BufChar   As String
    Dim strTmp As String
    Dim intIDX As Integer
    Dim varTmp  As Variant
    Dim j       As Integer
    Dim k As Integer
    
    Dim blnAppend1  As Boolean
    Dim blnAppend2  As Boolean
    Dim varTmp1    As Variant
    Dim strTest   As String
    Dim strResult As String
    
On Error GoTo ErrorTrap
    
    Screen.MousePointer = 11
    
    j = 0
    blnAppend1 = False
    blnAppend2 = False
    
    '-- �������ϸ�� ��θ� �����Ѵ�.
    strPath = strXML

    
    '1���ξ� �������� MSDN����
    Dim TextLine
    Open strPath For Input As #1 ' ������ ���ϴ�.
    
    Do While Not EOF(1) ' ������ ���� ���� ������ �ݺ��մϴ�.
        Line Input #1, TextLine ' ������ ������ ���� �о���Դϴ�.
        strBuffer = strBuffer & TextLine
    Loop
    
    Close #1 ' ������ �ݽ��ϴ�
 
    intIDX = 0
    lngBufLen = Len(strBuffer)
        

    
    'strBuffer = Replace(strBuffer, Chr(9), "")
    varTmp = Split(strBuffer, "</GROUP>")

    Erase strRecvData

'    blnSameRecord = True
    
    For i = 0 To UBound(varTmp)
        'Debug.Print varTmp(i)
        If InStr(varTmp(i), """Patient""") > 0 Then 'blnAppend1 = False And
            strTmp = Mid(varTmp(i), InStr(varTmp(i), """Patient""") + 14)
            ReDim Preserve strRecvData(j)
            strRecvData(j) = strRecvData(j) & "P|" & mGetP(strTmp, 1, """")
            j = j + 1
            blnAppend1 = True
        End If
        
        If InStr(varTmp(i), """Assay""") > 0 Then 'blnAppend1 = True And
            strTmp = Mid(varTmp(i), InStr(varTmp(i), """Assay""") + 12)
            ReDim Preserve strRecvData(j)
            
            If gAssayNM.INHALANT = mGetP(strTmp, 1, """") Then
                strRecvData(j) = strRecvData(j) & "O|INHALANT"
            ElseIf gAssayNM.FOOD = mGetP(strTmp, 1, """") Then
                strRecvData(j) = strRecvData(j) & "O|FOOD"
            ElseIf gAssayNM.ATOPY = mGetP(strTmp, 1, """") Then
                strRecvData(j) = strRecvData(j) & "O|ATOPY"
            Else
                f_subSet_XMLWorkList = ""
                Exit Function
            End If
            
            j = j + 1
            blnAppend2 = True
            
            'varTmp1 = Split(varTmp(i), "</PARAM>")
        End If
        
        
        If InStr(varTmp(i), """Blot""") > 0 Then 'blnAppend1 = True And blnAppend2 = True And
            varTmp1 = Split(varTmp(i), "</PARAM>")
            strTest = ""
            strResult = ""
            For k = 0 To UBound(varTmp1)
                If InStr(varTmp1(k), """Code""") > 0 Then
                    strTest = Mid(varTmp1(k), InStr(varTmp1(k), """Code""") + 7)
                End If
                
                If strTest <> "" And InStr(varTmp1(k), """QntResult""") > 0 Then
                    strResult = Mid(varTmp1(k), InStr(varTmp1(k), """QntResult""") + 12)
                    
                    strResult = Replace(strResult, "&lt;", "<")
                    strResult = Replace(strResult, "&gt;", ">")
                    ReDim Preserve strRecvData(j)
                    strRecvData(j) = strRecvData(j) & "R|" & strTest & "^" & strResult
                    j = j + 1
                    strTest = ""
                    strResult = ""
                End If
            Next
        End If
    Next
    
    Screen.MousePointer = 0

    Exit Function
        
ErrorTrap:
    
'    blnSameRecord = False
    Screen.MousePointer = 0
    
    
End Function


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
    StatusBar1.Panels(3).Text = ""
          
          SQL = " SELECT '', SAVESEQ, MID(EXAMDATE,1,8) AS EXAMDATE, HOSPDATE AS ��������, BARCODE AS ���ڵ��ȣ, CHARTNO AS ��Ʈ��ȣ, PID AS ������ȣ, PNAME AS �̸�,PSEX AS ����, PAGE AS ����, DISKNO, POSNO, EXAMCODE, RESULT, REFFLAG, SENDFLAG " & vbCrLf
    SQL = SQL & "   FROM PATRESULT " & vbCrLf
    SQL = SQL & "  WHERE MID(EXAMDATE,1,8) Between '" & Format(dtpStartDt, "YYYYMMDD") & "' AND '" & Format(dtpStopDt, "YYYYMMDD") & "'" & vbCrLf
    SQL = SQL & "    AND EQUIPNO = '" & gEquip & "' " & vbCrLf
    SQL = SQL & " ORDER BY EXAMDATE,SAVESEQ,HOSPDATE,BARCODE "
    
    Set RS = cn.Execute(SQL, , 1)

    If Not RS.EOF = True And Not RS.BOF = True Then
        Do Until RS.EOF
            With vasID
                For i = 1 To .DataRowCnt
                    strDate = GetText(vasID, i, colHOSPDATE)
                    strChart = GetText(vasID, i, colBARCODE)
                    strSaveSeq = GetText(vasID, i, colSAVESEQ)
                    
                    'If Trim(RS("��������")) = strDate And Trim(RS("SAVESEQ")) = strSaveSeq And Trim(RS("���ڵ��ȣ")) = strChart Then
                    If Trim(RS("EXAMDATE")) = GetText(vasID, i, colEXAMDATE) And Trim(RS("SAVESEQ")) = strSaveSeq And Trim(RS("���ڵ��ȣ")) = strChart Then
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

                    SetText vasID, "0", .MaxRows, colCheckBox
                    SetText vasID, Trim(RS.Fields("SAVESEQ")) & "", .MaxRows, colSAVESEQ
                    SetText vasID, Trim(RS.Fields("EXAMDATE")) & "", .MaxRows, colEXAMDATE
                    SetText vasID, Trim(RS.Fields("��������")) & "", .MaxRows, colHOSPDATE
                    SetText vasID, Trim(RS.Fields("���ڵ��ȣ")) & "", .MaxRows, colBARCODE
                    SetText vasID, Trim(RS.Fields("��Ʈ��ȣ")) & "", .MaxRows, colCHARTNO
                    SetText vasID, Trim(RS.Fields("������ȣ")) & "", .MaxRows, colPID
                    SetText vasID, Trim(RS.Fields("�̸�")) & "", .MaxRows, colPNAME
                    SetText vasID, Trim(RS.Fields("����")) & "", .MaxRows, colPSEX
                    SetText vasID, Trim(RS.Fields("����")) & "", .MaxRows, colPAGE
                    SetText vasID, Trim(RS.Fields("POSNO")) & "", .MaxRows, colDISKNO
                    SetText vasID, Trim(RS.Fields("DISKNO")) & "", .MaxRows, colINOUT
                    'SetText vasID, Trim(RS.Fields("POSNO")) & "", .MaxRows, colPOSNO
                    
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
    Else
        StatusBar1.Panels(3).Text = "��ȸ ����ڰ� �����ϴ�."
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
                    SetText vasID, "1", .MaxRows, colCheckBox
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
                    SetText vasID, "1", .MaxRows, colCheckBox
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
                    SetText vasID, "1", .MaxRows, colCheckBox
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
                    SetText vasID, "1", .MaxRows, colCheckBox
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


Private Sub GetWorkList_NTL(ByVal pFrDt As String, ByVal pToDt As String, Optional pBarNo As String)
    Dim RS          As ADODB.Recordset
    Dim i           As Integer
    Dim iCnt        As Long
    Dim intRow      As Long
    Dim intCol      As Integer
    Dim strDate     As String
    Dim strChart    As String
    Dim blnSame     As Boolean
    Dim sqlRet      As Integer
    
    If pBarNo = "" Then
        vasID.MaxRows = 0
        intRow = 0
    End If
    
    blnSame = False
    vasID.ReDraw = False
    StatusBar1.Panels(3).Text = ""
    
    '   @StatustIndex         tinyint,          '+#13#10+    // �ʼ��Է�: 0-��ü, 1-�̿Ϸ�, 2-�Ϸ�
    '   @WorkListCode         varchar(50),      '+#13#10+    // �ʼ��Է�: ��ũ����Ʈ�ڵ�
    '   @BeginDate         smalldatetime,       '+#13#10+    // �ʼ��Է�: ��ȸ��-����
    '   @EndDate         smalldatetime,         '+#13#10+    // �ʼ��Է�: ��ȸ��-��
    '   @BeginNo         int,                   '+#13#10+    // �����Է�: ������ȣ ���� (�⺻�� : 0)
    '   @EndNo         int,                     '+#13#10+    // �����Է�: ������ȣ ���� (�⺻�� : 0)
    '   @TestCodes         varchar(200)         '+#13#10+    // �����Է�: �˻��ڵ�

    '-- Record Count ������
    cn_Ser.CursorLocation = adUseClient
    Set RS = cn_Ser.Execute("Exec interface_GetPatientResultList02 '" & cboChk.ListIndex & "','" & gWKCD & "','" & Format(dtpStartDt.Value, "yyyy-mm-dd") & "','" & Format(dtpStopDt.Value, "yyyy-mm-dd") & "'," & Val(txtStartNum.Text) & "," & Val(txtStopNum.Text) & ",''", sqlRet)
    
'    strRegDate = "2017-05-18"
'
'    lngRegNo = 1
'
'    Set RS = cn_Ser.Execute("Exec Interface_GetPatientResult02 '" & gWKCD & "','" & strRegDate & "','" & lngRegNo & "'")
    
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
                    strChart = GetText(vasID, i, colPID)
                    If Trim(RS("LabRegDate")) = strDate And Format(Trim(RS.Fields("LabRegDate")), "yymmdd") & PedLeftStr(Trim(RS.Fields("LabRegNo")), 5, "0") = strChart Then
                        blnSame = True
                    End If
                Next
                'If Trim(RS.Fields("PatientChartNo")) = "8608" Then Stop
                '    Debug.Print Trim(RS.Fields("PatientName")) & "" & "-" & Trim(RS.Fields("OrderCode")) & ""
                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    SetText vasID, "1", .MaxRows, colCheckBox
                    SetText vasID, Trim(RS.Fields("LabRegDate")) & "", .MaxRows, colHOSPDATE
                    SetText vasID, Format(Trim(RS.Fields("LabRegDate")), "yymmdd") & PedLeftStr(Trim(RS.Fields("LabRegNo")), 5, "0"), .MaxRows, colBARCODE
                    SetText vasID, Trim(RS.Fields("PatientChartNo")) & "", .MaxRows, colCHARTNO
                    SetText vasID, Trim(RS.Fields("LabRegNo")) & "", .MaxRows, colPID
                    SetText vasID, Trim(RS.Fields("PatientName")) & "", .MaxRows, colPNAME
                    SetText vasID, Trim(RS.Fields("CompanyCode")) & "", .MaxRows, colDISKNO
                    SetText vasID, Trim(RS.Fields("PatientBirthDay")) & "", .MaxRows, colPOSNO    '-- �������
                    SetText vasID, Trim(RS.Fields("PatientSex")) & "", .MaxRows, colPSEX
                    SetText vasID, Trim(RS.Fields("PatientAge")) & "", .MaxRows, colPAGE
                    Select Case Trim(RS.Fields("OrderCode")) & ""
                        Case "63100":   SetText vasID, "INHALANT", .MaxRows, colINOUT
                        Case "63200":   SetText vasID, "FOOD", .MaxRows, colINOUT
                        Case "63300":   SetText vasID, "ATOPY", .MaxRows, colINOUT
                        
                        Case "4044"     'Food 2016.07.15
                                        '-- �������Ϸ� ������ �� ��� ����
                                        'SetBackColor vasID, .MaxRows, .MaxRows, 1, colState, 202, 255, 112
                                        'SetText vasID, "Profile����", .MaxRows, colINOUT
                                        SetText vasID, "FOOD", .MaxRows, colINOUT
                        Case Else:
                                        SetBackColor vasID, .MaxRows, .MaxRows, 1, colState, 202, 255, 112
                                        SetText vasID, "ó�����", .MaxRows, colINOUT
                                                                            
                    End Select
                End If
                'Debug.Print Trim(RS.Fields("OrderCode")) & ""
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
                
                SetText vasID, "1", intRow, colCheckBox
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
                    
                    SetText vasID, "1", intRow, colCheckBox
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
        Case "NTL":         Call GetWorkList_NTL(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))
    End Select
    
    vasID.RowHeight(-1) = 12
    vasRes.MaxRows = 0
    
End Sub


Private Sub cmdSL_Click()
    If cmdSL.Caption = "��" Then
        cmdSL.Caption = "��"
        vasID.Width = 20655 '19605
    Else
        cmdSL.Caption = "��"
        vasID.Width = 11985
    End If

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

Private Sub lblclear_Click()
    lblChangePID.Caption = ""
    lblChangeBar.Caption = ""
    lblBarcode(0).Caption = ""
    lblPname(0).Caption = ""
    lblSaveSeq.Caption = ""
    lblExamDate.Caption = ""
End Sub

Private Sub Command16_Click()
    
    strBuffer = ":N1    80 81                 00620141422      15 1   7.0  2   4.1  3   0.5  4   4.5  5    34  6    20  7   417  8   239  9    97 14    85 15    14 16   0.7 18    93 19      T54     1 "
    
    strBuffer = txtTest.Text
    
    Call comEqp_OnComm
        

End Sub

Private Sub Form_Load()
    Dim sDate As String
    Dim i As Integer
    
On Error GoTo Rst

    If App.PrevInstance Then
        End
    End If
    
    Me.Left = 0
    Me.Top = 0
    
    'Me.Height = 11520
    Me.Width = 16545
    
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
        'fraWork.Visible = False
    
        chkMode.Caption = "Barcode"
        MnModeBarcode.Checked = True
        MnModeWorkList.Checked = False
        chkBar.Value = 1
    Else
        'fraBar.Visible = False
'        fraWork.Visible = True
    
        chkMode.Caption = "WorkList"
        MnModeBarcode.Checked = False
        MnModeWorkList.Checked = True
        chkBar.Value = 0
    End If
    
'    If gScreen = "����" Then
'        cmdSL.Caption = "��"
'        vasID.Width = 14595
'    Else
'        cmdSL.Caption = "��"
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
        frmInterface.StatusBar1.Panels(2).Text = "COM" & comEqp.CommPort & " ��Ʈ�� ���� �Ǿ����ϴ�"
        imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
        imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
        imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
    Else
        frmInterface.StatusBar1.Panels(2).Text = "�����Ʈ�� ���� ���� �ʾҽ��ϴ�"
        imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
        imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
        imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
    End If

    If Not Connect_Local Then
        MsgBox "������� �ʾҽ��ϴ�."
        cn_Local_Flag = False
        Exit Sub
    Else
        cn_Local_Flag = True
    End If
    
'    -- osw �߰�
    For i = 1 To 1
        If Not Connect_PRServer Then
            MsgBox "������� �ʾҽ��ϴ�."
            cn_Server_Flag = False
            Exit Sub
        Else
            cn_Server_Flag = True
        End If
    Next
    
    '-- osw �߰�
'    For i = 1 To 1
'        If Not Connect_DRServer Then
'            MsgBox "������� �ʾҽ��ϴ�."
'            cn_Server_Flag = False
'            Exit Sub
'        Else
            cn_Server_Flag = True
'        End If
'    Next
    
    GetExamCode
    
    SetExamCode
    
    dtpToday = Date
    dtpStartDt = Date
    dtpStopDt = Date
    
    sDate = Format(DateAdd("y", CDate(dtpToday.Value), -999), "yyyymmdd")
    
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
    
   ' Call cmdSL_Click
    
    '-- test
'    vasID.MaxRows = 10
    
Exit Sub
    
Rst:
    If Err.Number = "8002" Then
        If (MsgBox("��Ʈ ��ȣ�� �߸��Ǿ����ϴ�." & vbNewLine & vbNewLine & "   ��� �����Ͻðڽ��ϱ�?", vbYesNo + vbCritical, Me.Caption)) = vbYes Then
            Resume Next
        Else
            End
        End If
    Else
        If (MsgBox(Err.Description & vbNewLine & vbNewLine & "   ��� �����Ͻðڽ��ϱ�?", vbYesNo + vbCritical, Me.Caption)) = vbYes Then
            Resume Next
        Else
            End
        End If
    End If
        
        
End Sub

Private Sub SetExamCode()
    Dim i As Integer
    
    
    With vasID
        .MaxCols = colState + UBound(gArrEquip)
        
        For i = 0 To UBound(gArrEquip) - 1
            .Col = colState + (i + 1)
            .Row = -1
            '.CellType = CellTypeEdit
            '.TypeEditCharSet = TypeEditCharSetAlphanumeric
            .TypeEditCharCase = TypeEditCharCaseSetUpper
            .TypeHAlign = TypeHAlignCenter
            .TypeVAlign = TypeVAlignCenter
            'Call SetText(vasID, gArrEquip(i + 1, 2), 0, colState + (i + 1))
            Call SetText(vasID, gArrEquip(i + 1, 4), 0, colState + (i + 1))
            .ColWidth(colState + (i + 1)) = 20
        Next
    End With
    
End Sub


Function GetExamCode() As Integer
    Dim i, j As Long
    
    ClearSpread vasTemp
    GetExamCode = -1
    gAllExam = ""
    SQL = "Select equipcode, examcode, examname, resprec, seqno, '', examtype " & vbCrLf & _
          "  From EQPMASTER " & vbCrLf & _
          " Where equipno = '" & gEquip & "' " & vbCrLf & _
          " Order by seqno * 10 "
    Res = GetDBSelectVas(gLocal, SQL, vasCode)
    If Res > 0 Then
        ReDim gArrEquip(1 To vasCode.DataRowCnt, 1 To 9)
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
        For j = 1 To 7
            'Debug.Print Trim(GetText(vasCode, i, j))
            gArrEquip(i, j + 1) = Trim(GetText(vasCode, i, j))
        Next j
    Next i
    
    GetExamCode = 1
End Function

Private Sub Form_Unload(Cancel As Integer)
    If comEqp.PortOpen = True Then
        comEqp.PortOpen = False
    End If

'    Call dce_close_env      ' Server�� ������ ���� ��
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
'    MsgBox "�����Ʈ ��� ����", vbOKOnly + vbInformation, Me.Caption
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

'    If txtTest.Text <> "" And strBuffer = txtTest.Text Then
'        Buffer = strBuffer
'        GoTo Rst
'    End If
    
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
'Rst:
            SetRawData "[Rx]" & Buffer
            StatusBar1.Panels(3).Text = Buffer
            
            lngBufLen = Len(Buffer)
            
            'Debug.Print Buffer

            For i = 1 To lngBufLen
                BufChar = Mid$(Buffer, i, 1)
                Select Case BufChar
                    Case ENQ
                        ReceiveData = ""
                        comEqp.Output = ACK
                        SetRawData "[Tx]" & ACK
                    Case STX
                        intBufCnt = 1
                        ReceiveData = ""
                    Case ETX
                        comEqp.Output = ACK
                        SetRawData "[Tx]" & ACK
                        ReceiveData = ReceiveData & BufChar
                        Call EditRcvDataASTM
                    Case EOT
                        ReceiveData = ""
                    Case Else
                        ReceiveData = ReceiveData & BufChar
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
    
    '-- ���������� ǥ��
    Call SetText(vasID, "1", intRow, colCheckBox)
    Call SetText(vasID, pBarNo, intRow, colBARCODE)
    Call SetText(vasID, mResult.RsltDate, intRow, colEXAMDATE)
    Call SetText(vasID, mResult.RsltSeq, intRow, colSAVESEQ)
'    Call SetText(vasID, mResult.MnmNm, intRow, colINOUT)
    
    Call vasActiveCell(vasID, intRow, colBARCODE)
    
    '-- ����������� �����
    Call ClearSpread(vasRes)
    
    '-- �˻��� ���� �������̺��� ������ ǥ��(for ��ũ����Ʈ)  '6,7,8,9
    Call GetSampleInfoW_NTL(intRow)
    
    '-- ���ڵ��ȣ�� �����ϴ� �˻��ڵ� ��������(�μ� : ����ڵ�,íƮ��ȣ,������,������ȣ,������ȣ)
    'gOrderExam = GetOrderExamCode(gEquip, pBarNo)
    
    '-- ���� Row
    gRow = intRow
    
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
    Dim lsResult_Buff As String
    Dim lsExamDate As String
    Dim lsEquipRes As String
    Dim lsResRow    As String
    Dim ii As Integer
    Dim strTmp      As String
    Dim intIDX      As Integer
    Dim varRcvBuf   As Variant
    Dim intRow      As Integer
    Dim i As Integer
    Dim intCol As Integer
    Dim varHoriba As Variant
    Dim Pos As Integer
    Dim strSeqNo As String
    Dim varORQN As Variant
    Dim strHoleNo    As String
    Dim strPatID     As String
    Dim strExamSubCd As String
    
    SetRawData "[Rcv]" & ReceiveData
    
    strTemp1 = ReceiveData 'strBuffer
    
    If Mid$(strTemp1, 2, 1) <> "U" Then '"U" = ��ü "C" = ��Ʈ��
        Exit Sub
    End If
    
    strIntBase = "1"
    
    strRcvBuf = strTemp1
    
    strSeqNo = Trim(Mid(strRcvBuf, 3, 4))
    strRackNo = Trim(Mid(strRcvBuf, 8, 2))
    strTubePos = Trim(Mid(strRcvBuf, 10, 2))
    strBarNo = Trim(Mid(strRcvBuf, 13, 12))

'    strBarNo = "170518000001"
    
    strIntResult = Trim$(Mid$(strRcvBuf, 32, 6))   '����
    strResult = Trim$(Mid$(strRcvBuf, 43, 1))   '����
        
    If strResult = "-" Then
        If strIntResult = "0" Then
            strIntResult = "1"
        End If
        strResult = "Negative : " & strIntResult
    ElseIf strResult = "+" Then
        strResult = "Positive : " & strIntResult
    End If
    
    With mResult
        .BarNo = strBarNo
        .RackNo = strRackNo
        .TubePos = strTubePos
        .Seq = strSeq
        .RsltDate = Format(Now, "yyyymmddhhmmss")
        .RsltSeq = getMaxTestNum(Format(dtpToday, "yyyymmdd"))
    End With
    
    Call SetPatInfo(strBarNo)
    
    If strResult <> "" Then
        SQL = ""
        SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
        SQL = SQL & "  FROM EQPMASTER"
        SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
        SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
        SQL = SQL & "   AND EXAMCODE in (" & gOrderExam & ") "
        
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
            
            '-- �������� seq
            For intCol = colState + 1 To vasID.MaxCols
                If lsExamCode = gArrEquip(intCol - colState, 3) Then
                    SetText vasID, strResult, gRow, intCol
                    
                    strExamSubCd = gArrEquip(intCol - colState, 9)
                    Exit For
                End If
            Next

            '-- ��� List
            SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      '����ڵ�
            SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '�˻��ڵ�
            SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '�˻��
            SetText vasRes, lsEquipRes, lsResRow, colMachResult     '�����
            SetText vasRes, strResult, lsResRow, colRESULT          '���
            SetText vasRes, lsSeqNo, lsResRow, colSeq               '����
            SetText vasRes, strComm, lsResRow, 7                    'Flag
            
            
            SetText vasRes, strExamSubCd, lsResRow, colSUBCODE                    'colSUBCODE
            
            '-- ���� ����
            SetLocalDB gRow, lsResRow, "1", lsEquipRes
                        
            lsResult_Buff = ""
            
            strState = "R"
            
        '-- ���� ���� ���
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
                
                '�Ҽ��� ó��, ��� ���� ó��
                lsEquipRes = strResult
                strResult = SetResult(strResult, strIntBase)
                lsResult_Buff = strResult
                
                '-- Work List
                SetText vasID, "Result", gRow, colState                 '�������
                
                '-- �������� seq
                For intCol = colState + 1 To vasID.MaxCols
                    If lsExamCode = gArrEquip(intCol - colState, 3) Then
                        SetText vasID, strResult, gRow, intCol
                        Exit For
                    End If
                Next

                '-- ��� List
                SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      '����ڵ�
                SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '�˻��ڵ�
                SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '�˻��
                SetText vasRes, lsEquipRes, lsResRow, colMachResult     '�����
                SetText vasRes, strResult, lsResRow, colRESULT          '���
                SetText vasRes, lsSeqNo, lsResRow, colSeq               '����
                SetText vasRes, strComm, lsResRow, colFLAG              'Flag
                '-- ���� ����
                SetLocalDB gRow, lsResRow, "1", lsEquipRes
                
                lsResult_Buff = ""
                strState = "R"
            End If
        End If
    End If
            
    vasRes.RowHeight(-1) = 14
            
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
            SetText vasID, "0", gRow, colCheckBox
            
                  SQL = "Update PATRESULT Set " & vbCrLf
            SQL = SQL & " sendflag = '2' " & vbCrLf
            SQL = SQL & " Where equipno = '" & gEquip & "' " & vbCrLf
            SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(vasID, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
            SQL = SQL & "   And barcode = '" & Trim(GetText(vasID, gRow, colBARCODE)) & "' " & vbCrLf
            SQL = SQL & "   And saveseq = " & Trim(GetText(vasID, gRow, colSAVESEQ)) & vbCrLf
            
            Res = SendQuery(gLocal, SQL)
            If Res = -1 Then
                SaveQuery SQL
                Exit Sub
            End If
        End If
        strState = ""
    End If

End Sub

'-----------------------------------------------------------------------------'
'   ��� : ���κ� ������ ������ ����
'-----------------------------------------------------------------------------'
Private Sub EditRcvDataAPEX()
    Dim strRcvBuf    As String   '������ Data
    Dim strType      As String   '������ Record Type
    Dim strBarNo     As String   '������ ���ڵ��ȣ
    Dim strSeq       As String   '������ Sequence
    Dim strRackNo    As String   '������ Rack Or Disk No
    Dim strTubePos   As String   '������ Tube Position
    Dim strIntBase   As String   '������ ������ �˻��
    Dim strResult    As String   '������ ���(����)
    
    Dim strFIntBase   As String   '������ ������ �˻��
    Dim strFResult    As String   '������ ���(����)
    
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
    Dim lsResult_Buff As String
    Dim lsExamDate As String
    Dim lsEquipRes As String
    Dim lsResRow    As String
    Dim ii As Integer
    Dim strTmp      As String
    Dim intIDX      As Integer
    Dim varRcvBuf   As Variant
    Dim intRow      As Integer
    Dim i As Integer
    Dim intCol As Integer
    Dim varHoriba As Variant
    Dim strGubun  As String
    
    For intCnt = 0 To UBound(strRecvData)
        strRcvBuf = strRecvData(intCnt)
        
        strType = Mid$(strRcvBuf, 1, 1)
        If IsNumeric(strType) Then
            strType = Mid$(strRcvBuf, 2, 1)
        End If
        
        Select Case strType
            Case "H"    '## Header
            Case "P"    '## Order
                strBarNo = Trim(mGetP(strRcvBuf, 2, "|"))
                'strRackNo = mGetP(strTemp1, 1, "^")
                'strTubePos = mGetP(strTemp1, 2, "^")
                
                'If strBarNo = "" Then Exit Sub
                
            
            Case "O"    '## Order
                strGubun = Trim(mGetP(strRcvBuf, 2, "|"))
                                
                With mResult
                    .BarNo = strBarNo
                    '.RackNo = strRackNo
                    '.TubePos = strTubePos
                    .RsltDate = Format(Now, "yyyymmddhhmmss")
                    .RsltSeq = getMaxTestNum(Format(dtpToday, "yyyymmdd"))
                    .MnmNm = strGubun
                End With
                                
                Call SetPatInfo(strBarNo)
                
                If gRow <= 0 Then
                    'Exit Sub
                End If
                
                strState = "O"
                
                '-- ������ ���ȭ�� �ʱ�ȭ
                vasRes.MaxRows = 0

                
            Case "R"    '## Result
                strFIntBase = ""
                strFResult = ""
                '## ������ �˻��, ���, Abnormal Flag
                strIntBase = Trim(mGetP(mGetP(strRcvBuf, 2, "|"), 1, "^"))
                strResult = Trim(mGetP(mGetP(strRcvBuf, 2, "|"), 2, "^"))
                'Debug.Print strIntBase
                strFIntBase = strIntBase
                
                'strResult = Replace(strResult, "<", "")
                If InStr(strResult, "<0.15") > 0 Then
                    strResult = "0.00"
                End If
                'strResult = Replace(strResult, "<", "")
                strResult = Replace(strResult, ">", "")
                strFResult = strResult
                
                'code fx1(Fish mix) -> f313(��ġ)-1.00, f254(���ڹ�)-0.99, f413(����)-0.98
                'If strBarNo <> "" And mResult.MnmNm = "FOOD" And strIntBase = "fx1" And strResult <> "" and  IsNumeric(strResult)  Then
                If strBarNo <> "" And mResult.MnmNm = "FOOD" And strIntBase = "fx1" And strResult <> "" Then
                    For i = 1 To 3
                        If IsNumeric(strResult) Then
                            If i = 1 Then
                                strIntBase = "f313"
                                strResult = strResult
                            ElseIf i = 2 Then
                                strIntBase = "f254"
                                strResult = strResult * 0.99
                            ElseIf i = 3 Then
                                strIntBase = "f413"
                                strResult = strResult * 0.98
                            End If
                        Else
                            If i = 1 Then
                                strIntBase = "f313"
                            ElseIf i = 2 Then
                                strIntBase = "f254"
                            ElseIf i = 3 Then
                                strIntBase = "f413"
                            End If
                        End If
                        
                        
                        If strResult <> "" And Len(strIntBase) <= 6 Then
                            SQL = ""
                            SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
                            SQL = SQL & "  FROM EQPMASTER"
                            SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
                            SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
                            SQL = SQL & "   AND GUBUN = '" & strGubun & "'"
                            SQL = SQL & "   AND EXAMCODE in (" & gOrderExam & ") "
                            
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
                                
                                '-- �������� seq
                                For intCol = colState + 1 To vasID.MaxCols
                                    If lsExamCode = gArrEquip(intCol - colState, 3) And strGubun = gArrEquip(intCol - colState, 7) Then
                                        SetText vasID, strResult, gRow, intCol
                                        'SetText vasRes, gArrEquip(intCol - colState, 7), lsResRow, colSUBCODE               'subcode
                                        SetText vasRes, gArrEquip(intCol - colState, 9), lsResRow, colSUBCODE               'subcode
                                        Exit For
                                    End If
                                Next
                                
                                '-- ��� List
                                SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      '����ڵ�
                                SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '�˻��ڵ�
                                SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '�˻��
                                SetText vasRes, lsEquipRes, lsResRow, colMachResult     '�����
                                SetText vasRes, strResult, lsResRow, colRESULT          '���
                                SetText vasRes, lsSeqNo, lsResRow, colSeq               '����
                                SetText vasRes, strComm, lsResRow, 7                    'Flag
                                '-- ���� ����
                                SetLocalDB gRow, lsResRow, "1", lsEquipRes
                                            
                                lsResult_Buff = ""
                                
                                strState = "R"
                                
                            '-- ���� ���� ���
                            Else
                            
                                      SQL = "Select examcode, examname, seqno "
                                SQL = SQL & "  From EQPMASTER"
                                SQL = SQL & " Where equipno = '" & gEquip & "' "
                                SQL = SQL & "   and equipcode = '" & strIntBase & "' "
                                SQL = SQL & "   AND GUBUN = '" & strGubun & "'"
                                Res = GetDBSelectColumn(gLocal, SQL)
                                                        
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
                                    SetText vasID, "Result", gRow, colState                 '�������
                                    
                                    '-- �������� seq
                                    For intCol = colState + 1 To vasID.MaxCols
                                        If lsExamCode = gArrEquip(intCol - colState, 3) And strGubun = gArrEquip(intCol - colState, 7) Then
                                            SetText vasID, strResult, gRow, intCol
                                            'SetText vasRes, gArrEquip(intCol - colState, 7), lsResRow, colSUBCODE               'subcode
                                            SetText vasRes, gArrEquip(intCol - colState, 9), lsResRow, colSUBCODE               'subcode
                                            Exit For
                                        End If
                                    Next
                                    
                                    '-- ��� List
                                    SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      '����ڵ�
                                    SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '�˻��ڵ�
                                    SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '�˻��
                                    SetText vasRes, lsEquipRes, lsResRow, colMachResult     '�����
                                    SetText vasRes, strResult, lsResRow, colRESULT          '���
                                    SetText vasRes, lsSeqNo, lsResRow, colSeq               '����
                                    SetText vasRes, strComm, lsResRow, colFLAG              'Flag
                                    '-- ���� ����
                                    SetLocalDB gRow, lsResRow, "1", lsEquipRes
                                    
                                    lsResult_Buff = ""
                                    strState = "R"
                                End If
                            End If
                        End If
                    Next
                    strIntBase = ""
                    strResult = ""
                    strIntBase = strFIntBase
                    strResult = strFResult
                End If
                
                'code fx4(Nut mix) -> f20(�Ƹ��)-1.00, f253(��)-0.99, w204(�عٶ�⾾)-0.98
                'If strBarNo <> "" And mResult.MnmNm = "FOOD" And strIntBase = "fx4" And strResult <> "" And IsNumeric(strResult) Then
                If strBarNo <> "" And mResult.MnmNm = "FOOD" And strIntBase = "fx4" And strResult <> "" Then
                    For i = 1 To 3
                        If IsNumeric(strResult) Then
                            If i = 1 Then
                                strIntBase = "f20"
                                strResult = strResult
                            ElseIf i = 2 Then
                                strIntBase = "f253"
                                strResult = strResult * 0.99
                            ElseIf i = 3 Then
                                strIntBase = "w204"
                                strResult = strResult * 0.98
                            End If
                        Else
                            If i = 1 Then
                                strIntBase = "f20"
                            ElseIf i = 2 Then
                                strIntBase = "f253"
                            ElseIf i = 3 Then
                                strIntBase = "w204"
                            End If
                        End If
                        If strResult <> "" And Len(strIntBase) <= 6 Then
                            SQL = ""
                            SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
                            SQL = SQL & "  FROM EQPMASTER"
                            SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
                            SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
                            SQL = SQL & "   AND GUBUN = '" & strGubun & "'"
                            SQL = SQL & "   AND EXAMCODE in (" & gOrderExam & ") "
                            
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
                                
                                '-- �������� seq
                                For intCol = colState + 1 To vasID.MaxCols
                                    If lsExamCode = gArrEquip(intCol - colState, 3) And strGubun = gArrEquip(intCol - colState, 7) Then
                                        SetText vasID, strResult, gRow, intCol
                                        'SetText vasRes, gArrEquip(intCol - colState, 7), lsResRow, colSUBCODE               'subcode
                                        SetText vasRes, gArrEquip(intCol - colState, 9), lsResRow, colSUBCODE               'subcode
                                        Exit For
                                    End If
                                Next
                                
                                '-- ��� List
                                SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      '����ڵ�
                                SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '�˻��ڵ�
                                SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '�˻��
                                SetText vasRes, lsEquipRes, lsResRow, colMachResult     '�����
                                SetText vasRes, strResult, lsResRow, colRESULT          '���
                                SetText vasRes, lsSeqNo, lsResRow, colSeq               '����
                                SetText vasRes, strComm, lsResRow, 7                    'Flag
                                '-- ���� ����
                                SetLocalDB gRow, lsResRow, "1", lsEquipRes
                                            
                                lsResult_Buff = ""
                                
                                strState = "R"
                                
                            '-- ���� ���� ���
                            Else
                            
                                      SQL = "Select examcode, examname, seqno "
                                SQL = SQL & "  From EQPMASTER"
                                SQL = SQL & " Where equipno = '" & gEquip & "' "
                                SQL = SQL & "   and equipcode = '" & strIntBase & "' "
                                SQL = SQL & "   AND GUBUN = '" & strGubun & "'"
                                Res = GetDBSelectColumn(gLocal, SQL)
                                                        
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
                                    SetText vasID, "Result", gRow, colState                 '�������
                                    
                                    '-- �������� seq
                                    For intCol = colState + 1 To vasID.MaxCols
                                        If lsExamCode = gArrEquip(intCol - colState, 3) And strGubun = gArrEquip(intCol - colState, 7) Then
                                            SetText vasID, strResult, gRow, intCol
                                            'SetText vasRes, gArrEquip(intCol - colState, 7), lsResRow, colSUBCODE               'subcode
                                            SetText vasRes, gArrEquip(intCol - colState, 9), lsResRow, colSUBCODE               'subcode
                                            Exit For
                                        End If
                                    Next
                                    
                                    '-- ��� List
                                    SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      '����ڵ�
                                    SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '�˻��ڵ�
                                    SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '�˻��
                                    SetText vasRes, lsEquipRes, lsResRow, colMachResult     '�����
                                    SetText vasRes, strResult, lsResRow, colRESULT          '���
                                    SetText vasRes, lsSeqNo, lsResRow, colSeq               '����
                                    SetText vasRes, strComm, lsResRow, colFLAG              'Flag
                                    '-- ���� ����
                                    SetLocalDB gRow, lsResRow, "1", lsEquipRes
                                    
                                    lsResult_Buff = ""
                                    strState = "R"
                                End If
                            End If
                        End If
                    Next
                    strIntBase = ""
                    strResult = ""
                    strIntBase = strFIntBase
                    strResult = strFResult
                End If
                
                'code w71/e73(����/��) -> w71(����)-1.00, e73(��)-0.99
                'If strBarNo <> "" And mResult.MnmNm = "INHALANT" And strIntBase = "w71/e73" And strResult <> "" And IsNumeric(strResult) Then
                If strBarNo <> "" And mResult.MnmNm = "INHALANT" And strIntBase = "e71/e73" And strResult <> "" Then
                    For i = 1 To 2
                        If IsNumeric(strResult) Then
                            If i = 1 Then
                                strIntBase = "e71"
                                strResult = strResult
                            ElseIf i = 2 Then
                                strIntBase = "e73"
                                strResult = strResult * 0.99
                            End If
                        Else
                            If i = 1 Then
                                strIntBase = "e71"
                            ElseIf i = 2 Then
                                strIntBase = "e73"
                            End If
                        End If
                        
                        If strResult <> "" And Len(strIntBase) <= 6 Then
                            SQL = ""
                            SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
                            SQL = SQL & "  FROM EQPMASTER"
                            SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
                            SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
                            SQL = SQL & "   AND GUBUN = '" & strGubun & "'"
                            SQL = SQL & "   AND EXAMCODE in (" & gOrderExam & ") "
                            
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
                                
                                '-- �������� seq
                                For intCol = colState + 1 To vasID.MaxCols
                                    If lsExamCode = gArrEquip(intCol - colState, 3) And strGubun = gArrEquip(intCol - colState, 7) Then
                                        SetText vasID, strResult, gRow, intCol
                                        'SetText vasRes, gArrEquip(intCol - colState, 7), lsResRow, colSUBCODE               'subcode
                                        SetText vasRes, gArrEquip(intCol - colState, 9), lsResRow, colSUBCODE               'subcode
                                        Exit For
                                    End If
                                Next
                                
                                '-- ��� List
                                SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      '����ڵ�
                                SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '�˻��ڵ�
                                SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '�˻��
                                SetText vasRes, lsEquipRes, lsResRow, colMachResult     '�����
                                SetText vasRes, strResult, lsResRow, colRESULT          '���
                                SetText vasRes, lsSeqNo, lsResRow, colSeq               '����
                                SetText vasRes, strComm, lsResRow, 7                    'Flag
                                '-- ���� ����
                                SetLocalDB gRow, lsResRow, "1", lsEquipRes
                                            
                                lsResult_Buff = ""
                                
                                strState = "R"
                                
                            '-- ���� ���� ���
                            Else
                            
                                      SQL = "Select examcode, examname, seqno "
                                SQL = SQL & "  From EQPMASTER"
                                SQL = SQL & " Where equipno = '" & gEquip & "' "
                                SQL = SQL & "   and equipcode = '" & strIntBase & "' "
                                SQL = SQL & "   AND GUBUN = '" & strGubun & "'"
                                Res = GetDBSelectColumn(gLocal, SQL)
                                                        
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
                                    SetText vasID, "Result", gRow, colState                 '�������
                                    
                                    '-- �������� seq
                                    For intCol = colState + 1 To vasID.MaxCols
                                        If lsExamCode = gArrEquip(intCol - colState, 3) And strGubun = gArrEquip(intCol - colState, 7) Then
                                            SetText vasID, strResult, gRow, intCol
                                            'SetText vasRes, gArrEquip(intCol - colState, 7), lsResRow, colSUBCODE               'subcode
                                            SetText vasRes, gArrEquip(intCol - colState, 9), lsResRow, colSUBCODE               'subcode
                                            Exit For
                                        End If
                                    Next
                                    
                                    '-- ��� List
                                    SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      '����ڵ�
                                    SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '�˻��ڵ�
                                    SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '�˻��
                                    SetText vasRes, lsEquipRes, lsResRow, colMachResult     '�����
                                    SetText vasRes, strResult, lsResRow, colRESULT          '���
                                    SetText vasRes, lsSeqNo, lsResRow, colSeq               '����
                                    SetText vasRes, strComm, lsResRow, colFLAG              'Flag
                                    '-- ���� ����
                                    SetLocalDB gRow, lsResRow, "1", lsEquipRes
                                    
                                    lsResult_Buff = ""
                                    strState = "R"
                                End If
                            End If
                        End If
                    Next
                    strIntBase = ""
                    strResult = ""
                    strIntBase = strFIntBase
                    strResult = strFResult
                End If
                
                'code gx1(grass mix) -> g1(���Ǯ)-1.00, g3(������)-1.01, g7(����)-0.99, g9(�ܰ��̻�)-0.98
                'If strBarNo = "" And mResult.MnmNm = "INHALANT" And strIntBase = "gx1" And strResult <> "" And IsNumeric(strResult) Then
                If strBarNo <> "" And mResult.MnmNm = "INHALANT" And strIntBase = "gx1" And strResult <> "" Then
                    For i = 1 To 4
                        If IsNumeric(strResult) Then
                            If i = 1 Then
                                strIntBase = "g1"
                                strResult = strResult
                            ElseIf i = 2 Then
                                strIntBase = "g3"
                                strResult = strResult * 1.01
                            ElseIf i = 3 Then
                                strIntBase = "g7"
                                strResult = strResult * 0.99
                            ElseIf i = 4 Then
                                strIntBase = "g9"
                                strResult = strResult * 0.98
                            End If
                        Else
                            If i = 1 Then
                                strIntBase = "g1"
                            ElseIf i = 2 Then
                                strIntBase = "g3"
                            ElseIf i = 3 Then
                                strIntBase = "g7"
                            ElseIf i = 4 Then
                                strIntBase = "g9"
                            End If
                        End If
                        
                        If strResult <> "" And Len(strIntBase) <= 6 Then
                            SQL = ""
                            SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
                            SQL = SQL & "  FROM EQPMASTER"
                            SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
                            SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
                            SQL = SQL & "   AND GUBUN = '" & strGubun & "'"
                            SQL = SQL & "   AND EXAMCODE in (" & gOrderExam & ") "
                            
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
                                
                                '-- �������� seq
                                For intCol = colState + 1 To vasID.MaxCols
                                    If lsExamCode = gArrEquip(intCol - colState, 3) And strGubun = gArrEquip(intCol - colState, 7) Then
                                        SetText vasID, strResult, gRow, intCol
                                        'SetText vasRes, gArrEquip(intCol - colState, 7), lsResRow, colSUBCODE               'subcode
                                        SetText vasRes, gArrEquip(intCol - colState, 9), lsResRow, colSUBCODE               'subcode
                                        Exit For
                                    End If
                                Next
                                
                                '-- ��� List
                                SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      '����ڵ�
                                SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '�˻��ڵ�
                                SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '�˻��
                                SetText vasRes, lsEquipRes, lsResRow, colMachResult     '�����
                                SetText vasRes, strResult, lsResRow, colRESULT          '���
                                SetText vasRes, lsSeqNo, lsResRow, colSeq               '����
                                SetText vasRes, strComm, lsResRow, 7                    'Flag
                                '-- ���� ����
                                SetLocalDB gRow, lsResRow, "1", lsEquipRes
                                            
                                lsResult_Buff = ""
                                
                                strState = "R"
                                
                            '-- ���� ���� ���
                            Else
                            
                                      SQL = "Select examcode, examname, seqno "
                                SQL = SQL & "  From EQPMASTER"
                                SQL = SQL & " Where equipno = '" & gEquip & "' "
                                SQL = SQL & "   and equipcode = '" & strIntBase & "' "
                                SQL = SQL & "   AND GUBUN = '" & strGubun & "'"
                                Res = GetDBSelectColumn(gLocal, SQL)
                                                        
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
                                    SetText vasID, "Result", gRow, colState                 '�������
                                    
                                    '-- �������� seq
                                    For intCol = colState + 1 To vasID.MaxCols
                                        If lsExamCode = gArrEquip(intCol - colState, 3) And strGubun = gArrEquip(intCol - colState, 7) Then
                                            SetText vasID, strResult, gRow, intCol
                                            'SetText vasRes, gArrEquip(intCol - colState, 7), lsResRow, colSUBCODE               'subcode
                                            SetText vasRes, gArrEquip(intCol - colState, 9), lsResRow, colSUBCODE               'subcode
                                            Exit For
                                        End If
                                    Next
                                    
                                    '-- ��� List
                                    SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      '����ڵ�
                                    SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '�˻��ڵ�
                                    SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '�˻��
                                    SetText vasRes, lsEquipRes, lsResRow, colMachResult     '�����
                                    SetText vasRes, strResult, lsResRow, colRESULT          '���
                                    SetText vasRes, lsSeqNo, lsResRow, colSeq               '����
                                    SetText vasRes, strComm, lsResRow, colFLAG              'Flag
                                    '-- ���� ����
                                    SetLocalDB gRow, lsResRow, "1", lsEquipRes
                                    
                                    lsResult_Buff = ""
                                    strState = "R"
                                End If
                            End If
                        End If
                    Next
                    strIntBase = ""
                    strResult = ""
                    strIntBase = strFIntBase
                    strResult = strFResult
                End If
                                
                'If strResult <> "" And Len(strIntBase) <= 6 Then
                If strResult <> "" And Len(strIntBase) > 0 Then
                    SQL = ""
                    SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
                    SQL = SQL & "  FROM EQPMASTER"
                    SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
                    SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
                    SQL = SQL & "   AND GUBUN = '" & strGubun & "'"
                    SQL = SQL & "   AND EXAMCODE in (" & gOrderExam & ") "
                    
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
                        
                        '-- �������� seq
                        For intCol = colState + 1 To vasID.MaxCols
                            If lsExamCode = gArrEquip(intCol - colState, 3) And strGubun = gArrEquip(intCol - colState, 7) Then
                                SetText vasID, strResult, gRow, intCol
                                'SetText vasRes, gArrEquip(intCol - colState, 7), lsResRow, colSUBCODE               'subcode
                                SetText vasRes, gArrEquip(intCol - colState, 9), lsResRow, colSUBCODE               'subcode
                                Exit For
                            End If
                        Next
                        
                        '-- ��� List
                        SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      '����ڵ�
                        SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '�˻��ڵ�
                        SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '�˻��
                        SetText vasRes, lsEquipRes, lsResRow, colMachResult     '�����
                        SetText vasRes, strResult, lsResRow, colRESULT          '���
                        SetText vasRes, lsSeqNo, lsResRow, colSeq               '����
                        SetText vasRes, strComm, lsResRow, 7                    'Flag
                        '-- ���� ����
                        SetLocalDB gRow, lsResRow, "1", lsEquipRes
                                    
                        lsResult_Buff = ""
                        
                        strState = "R"
                        
                    '-- ���� ���� ���
                    Else
                    
                              SQL = "Select examcode, examname, seqno "
                        SQL = SQL & "  From EQPMASTER"
                        SQL = SQL & " Where equipno = '" & gEquip & "' "
                        SQL = SQL & "   and equipcode = '" & strIntBase & "' "
                        SQL = SQL & "   AND GUBUN = '" & strGubun & "'"
                        Res = GetDBSelectColumn(gLocal, SQL)
                                                
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
                            SetText vasID, "Result", gRow, colState                 '�������
                            
                            '-- �������� seq
                            For intCol = colState + 1 To vasID.MaxCols
                                If lsExamCode = gArrEquip(intCol - colState, 3) And strGubun = gArrEquip(intCol - colState, 7) Then
                                    SetText vasID, strResult, gRow, intCol
                                    'SetText vasRes, gArrEquip(intCol - colState, 7), lsResRow, colSUBCODE               'subcode
                                    'SetText vasRes, gArrEquip(intCol - colState, 8), lsResRow, colSUBCODE               'subcode
                                    SetText vasRes, gArrEquip(intCol - colState, 9), lsResRow, colSUBCODE               'subcode
                                    Exit For
                                End If
                            Next
                            
                            '-- ��� List
                            SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      '����ڵ�
                            SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '�˻��ڵ�
                            SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '�˻��
                            SetText vasRes, lsEquipRes, lsResRow, colMachResult     '�����
                            SetText vasRes, strResult, lsResRow, colRESULT          '���
                            SetText vasRes, lsSeqNo, lsResRow, colSeq               '����
                            SetText vasRes, strComm, lsResRow, colFLAG              'Flag
                            '-- ���� ����
                            SetLocalDB gRow, lsResRow, "1", lsEquipRes
                            
                            lsResult_Buff = ""
                            strState = "R"
                        End If
                    End If
                End If
                vasRes.RowHeight(-1) = 10
                
            Case "L"    '## Terminator
                
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
                        SetText vasID, "0", gRow, colCheckBox
                        
                              SQL = "Update PATRESULT Set " & vbCrLf
                        SQL = SQL & " sendflag = '2' " & vbCrLf
                        SQL = SQL & " Where equipno = '" & gEquip & "' " & vbCrLf
                        SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(vasID, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                        SQL = SQL & "   And barcode = '" & Trim(GetText(vasID, gRow, colBARCODE)) & "' " & vbCrLf
                        SQL = SQL & "   And saveseq = " & Trim(GetText(vasID, gRow, colSAVESEQ)) & vbCrLf
                        
                        Res = SendQuery(gLocal, SQL)
                        If Res = -1 Then
                            SaveQuery SQL
                            Exit Sub
                        End If
                    End If
                    strState = ""
                End If
        
        End Select
    Next

End Sub

'-----------------------------------------------------------------------------'
'   ��� : ���κ� ������ ������ ����
'-----------------------------------------------------------------------------'
Private Sub EditRcvDataAPEX_Front()
    Dim intCnt       As Integer
    Dim strRcvBuf    As String   '������ Data
    Dim strType      As String   '������ Record Type
    Dim strBarNo     As String   '������ ���ڵ��ȣ
    Dim strGubun     As String
    Dim strIntBase   As String   '������ ������ �˻��
    Dim strResult    As String   '������ ���(����)
    
    For intCnt = 0 To UBound(strRecvData)
        strRcvBuf = strRecvData(intCnt)
        
        strType = Mid$(strRcvBuf, 1, 1)
        If IsNumeric(strType) Then
            strType = Mid$(strRcvBuf, 2, 1)
        End If
        
        Select Case strType
            Case "H"    '## Header
            Case "P"    '## Order
                strBarNo = Trim(mGetP(strRcvBuf, 2, "|"))
                
            Case "O"    '## Order
                strGubun = Trim(mGetP(strRcvBuf, 2, "|"))
                                
            Case "R"    '## Result
                '## ������ �˻��, ���, Abnormal Flag
                strIntBase = Trim(mGetP(mGetP(strRcvBuf, 2, "|"), 1, "^"))
                strResult = Trim(mGetP(mGetP(strRcvBuf, 2, "|"), 2, "^"))
                
                
                If strResult <> "" And Len(strIntBase) <= 6 Then
                    
                End If
        
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
    
          SQL = "SELECT resprec, reflow, refhigh " & vbCr
    SQL = SQL & "  FROM EQPMASTER " & vbCr
    SQL = SQL & " WHERE equipcode = '" & sEquipCode & "' " & vbCr
    SQL = SQL & "   AND EQUIPNO = '" & gEquip & "' " & vbCr
    
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
'Function SetLocalDB(ByVal asRow1 As Long, ByVal asRow2 As Long, asSend As String, Optional asEquipResult As String = "")
'    Dim sCnt As String
'    Dim sExamDate As String
'    Dim strSaveSeq As String
'    Dim RS          As ADODB.Recordset
'    Dim blnUpdate   As Boolean
'    Dim intUpCnt    As Integer
'    Dim strUpData   As String
'    Dim strGubuns   As String
'    Dim varGubuns   As Variant
'    Dim intCnt      As Integer
'    Dim intCnt1     As Integer
'    Dim strChannel  As String
'    Dim strGubun    As String
'
'
'    blnUpdate = False
'    sExamDate = Format(dtpToday, "yyyymmddhhmmss")
'    'sExamDate = Trim(GetText(vasID, asRow1, colOrdDate))
'    strChannel = Trim(GetText(vasRes, asRow2, colEQUIPCODE))
'    strGubun = Trim(GetText(vasID, asRow1, colINOUT))
'
'    SQL = ""
'    SQL = "DELETE FROM PATRESULT " & vbCrLf & _
'          " WHERE EXAMDATE = '" & Mid(sExamDate, 1, 8) & "' " & vbCrLf & _
'          "   AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
'          "   AND SAVESEQ = " & Trim(GetText(vasID, asRow1, colSAVESEQ)) & vbCrLf & _
'          "   AND BARCODE = '" & Trim(GetText(vasID, asRow1, colBARCODE)) & "' " & vbCrLf & _
'          "   AND EQUIPCODE = '" & Trim(GetText(vasRes, asRow2, colEQUIPCODE)) & "'" & vbCrLf & _
'          "   AND EXAMCODE = '" & Trim(GetText(vasRes, asRow2, colEXAMCODE)) & "'" & vbCrLf
'    SQL = SQL & "   AND DISKNO = '" & Trim(GetText(vasID, asRow1, colINOUT)) & "'"
'    Res = SendQuery(gLocal, SQL)
'
'    If Res = -1 Then
'        SaveQuery SQL
'        Exit Function
'    End If
'
'    'strSaveSeq = getMaxTestNum(Mid(sExamDate, 1, 8))
'
'    '-- �����ڵ��� ���
'    For intCnt = 1 To UBound(gArrEquip)
'        If strChannel = gArrEquip(intCnt, 2) And strGubun = gArrEquip(intCnt, 7) And gArrEquip(intCnt, 8) = "����" Then
'            '-- ����
'            '-- ���� ���ڵ忡 (������ Ʋ���� ��) ���� ä���� ����� �ִ��� Ȯ���Ѵ�.
'            '-- Select
'                  SQL = " SELECT RESULT, DISKNO " & vbCrLf
'            SQL = SQL & "   FROM PATRESULT " & vbCrLf
'            SQL = SQL & "  WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf
'            SQL = SQL & "    AND DISKNO =  '" & strGubun & "' "
'            SQL = SQL & "    AND EXAMDATE = '" & sExamDate & "'" & vbCrLf
'            SQL = SQL & "    AND BARCODE = '" & Trim(GetText(vasID, asRow1, colBARCODE)) & "'" & vbCrLf
'            SQL = SQL & "    AND EQUIPCODE = '" & Trim(GetText(vasRes, asRow2, colEQUIPCODE)) & "'"
'            SQL = SQL & "    AND SENDFLAG <> '2'"
'            Set RS = cn.Execute(SQL, , 1)
'            strUpData = ""
'            strGubuns = ""
'            intUpCnt = 0
'
'            If Not RS.EOF = True And Not RS.BOF = True Then
'                Do Until RS.EOF
'                    If IsNumeric(RS.Fields("RESULT")) Then
'                        blnUpdate = True
'                        intUpCnt = intUpCnt + 1
'                        strUpData = Val(strUpData) + Val(RS.Fields("RESULT").Value)
'                        strGubuns = strGubuns & RS.Fields("DISKNO").Value & "|"
'
'                        strUpData = SetResult(strUpData, strChannel)
'                    Else
'                        Exit For
'                    End If
'                    RS.MoveNext
'                Loop
'            Else
'                Exit For
'            End If
'
'            If IsNumeric(Trim(GetText(vasRes, asRow2, colRESULT))) Then
'                strUpData = Val(strUpData) + Val(Trim(GetText(vasRes, asRow2, colRESULT)))
'                strUpData = strUpData / (intUpCnt + 1)
'
'                strUpData = SetResult(strUpData, strChannel)
'            Else
'                Exit For
'            End If
'            '-- ���� ���� ����� ������ �װ���� ���ݰ���� ���ؼ� ������ ���������ŭ �Ͽ� ���� �����Ѵ�.
'            '-- Update
'            If blnUpdate = True Then
'                varGubuns = Split(strGubuns, "|")
'                For intCnt1 = 0 To UBound(varGubuns) - 1
'                    '�Ҽ��� ó��, ��� ���� ó��
'                    'strUpData = SetResult(strUpData, strChannel)
'
'                          SQL = "UPDATE  PATRESULT Set "
'                    SQL = SQL & " RESULT = '" & strUpData & "'" & vbCrLf
'                    SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf
'                    SQL = SQL & "  AND EXAMDATE = '" & sExamDate & "'" & vbCrLf
'                    SQL = SQL & "  AND BARCODE = '" & Trim(GetText(vasID, asRow1, colBARCODE)) & "'" & vbCrLf
'                    SQL = SQL & "  AND EQUIPCODE = '" & Trim(GetText(vasRes, asRow2, colEQUIPCODE)) & "'" & vbCrLf
'                    SQL = SQL & "  AND DISKNO = '" & varGubuns(intCnt1) & "'"
'                    Res = SendQuery(gLocal, SQL)
'
'                    If Res = -1 Then
'                        SaveQuery SQL
'                        Exit Function
'                    End If
'                Next
'            End If
'            '-- Insert ��� ����
'            Exit For
'        End If
'    Next
'
'
'    SQL = ""
'    SQL = SQL & "INSERT INTO PATRESULT (" & vbCrLf
'    SQL = SQL & "SAVESEQ"                           '�������(��¥��)
'    SQL = SQL & ", EXAMDATE"                        '�˻�����"
'    SQL = SQL & ", HOSPDATE"                        '������������"
'    SQL = SQL & ", EQUIPNO"                         '����ڵ�"
'    SQL = SQL & ", BARCODE" & vbCrLf
'    SQL = SQL & ", EQUIPCODE"                       '�˻�ä��"
'    SQL = SQL & ", EXAMCODE"                        '�����˻��ڵ�"
'    SQL = SQL & ", EXAMSUBCODE"                     '�����˻��ڵ�(SUB)"
'    SQL = SQL & ", EXAMNAME"
'    SQL = SQL & ", SEQNO" & vbCrLf                  '�˻��Ϸù�ȣ"
'    SQL = SQL & ", SAMPLETYPE"                      '��ü����"
'    SQL = SQL & ", DISKNO"
'    SQL = SQL & ", POSNO"
'    SQL = SQL & ", EQUIPRESULT"                     '�����"
'    SQL = SQL & ", RESULT" & vbCrLf                 '�Ҽ���������"
'    SQL = SQL & ", REFFLAG"
'    SQL = SQL & ", REFVALUE"
'    SQL = SQL & ", CHARTNO"
'    SQL = SQL & ", PID"                             '���Ϲ�ȣ(������ȣ)"
'    SQL = SQL & ", PNAME" & vbCrLf
'    SQL = SQL & ", PSEX"
'    SQL = SQL & ", PAGE"
'    SQL = SQL & ", PJUMIN"
'    SQL = SQL & ", PANICVALUE"
'    SQL = SQL & ", DELTAVALUE" & vbCrLf
'    SQL = SQL & ", SENDFLAG"                        '���۱���(0:������,1:����)"
'    SQL = SQL & ", SENDDATE"
'    SQL = SQL & ", EXAMUID"
'    SQL = SQL & ", HOSPITAL)" & vbCrLf
'    SQL = SQL & " VALUES (" & vbCrLf
''    SQL = SQL & strSaveSeq
'    SQL = SQL & Trim(GetText(vasID, asRow1, colSAVESEQ))
'    SQL = SQL & ",'" & sExamDate
'    SQL = SQL & "','" & Trim(GetText(vasID, asRow1, colHOSPDATE))
'    SQL = SQL & "','" & gEquip
'    SQL = SQL & "','" & Trim(GetText(vasID, asRow1, colBARCODE))
'    SQL = SQL & "','" & Trim(GetText(vasRes, asRow2, colEQUIPCODE))
'    SQL = SQL & "','" & Trim(GetText(vasRes, asRow2, colEXAMCODE))
'    SQL = SQL & "','" & Trim(GetText(vasRes, asRow2, colSUBCODE))
'    SQL = SQL & "','" & Trim(GetText(vasRes, asRow2, colEXAMNAME))
'    SQL = SQL & "','" & Trim(GetText(vasRes, asRow2, colSeq))
'    SQL = SQL & "',''"
'    'SQL = SQL & ",'" & Trim(GetText(vasID, asRow1, colDISKNO))
'    SQL = SQL & ",'" & Trim(GetText(vasID, asRow1, colINOUT))
'    SQL = SQL & "','" & Trim(GetText(vasID, asRow1, colDISKNO))
'    SQL = SQL & "','" & Trim(GetText(vasRes, asRow2, colMachResult))
'
'    If strUpData <> "" Then
'        '�Ҽ��� ó��, ��� ���� ó��
'        'strUpData = SetResult(strUpData, strChannel)
'
'        SQL = SQL & "','" & strUpData
'    Else
'        SQL = SQL & "','" & Trim(GetText(vasRes, asRow2, colRESULT))
'    End If
'    SQL = SQL & "','" & Trim(GetText(vasRes, asRow2, colFLAG))
'    SQL = SQL & "',''"
'    SQL = SQL & ",'" & Trim(GetText(vasID, asRow1, colCHARTNO))
'    SQL = SQL & "','" & Trim(GetText(vasID, asRow1, colPID))
'    SQL = SQL & "','" & Trim(GetText(vasID, asRow1, colPNAME))
'    SQL = SQL & "','" & Trim(GetText(vasID, asRow1, colPSEX))
'    SQL = SQL & "','" & Trim(GetText(vasID, asRow1, colPAGE))
'    SQL = SQL & "',''"
'    SQL = SQL & ",''"
'    SQL = SQL & ",''"
'    SQL = SQL & ",'1'"
'    SQL = SQL & ",''"
'    SQL = SQL & ",'" & gIFUser
'    SQL = SQL & "','')"
'
'    Res = SendQuery(gLocal, SQL)
'
'    If Res = -1 Then
'        SaveQuery SQL
'        Exit Function
'    End If
'
'End Function

Function SetLocalDB(ByVal asRow1 As Long, ByVal asRow2 As Long, asSend As String, Optional asEquipResult As String = "")
    Dim sCnt As String
    Dim sExamDate As String
    Dim strSaveSeq As String
    
    sExamDate = Format(dtpToday, "yyyymmddhhmmss")
'    sExamDate = Trim(GetText(vasID, asRow1, colEXAMDATE))
    If Trim(GetText(vasID, asRow1, colSAVESEQ)) = "" Then
        Exit Function
    End If
    
    SQL = ""
    SQL = "DELETE FROM PATRESULT " & vbCrLf & _
          " WHERE EXAMDATE = '" & Mid(sExamDate, 1, 8) & "' " & vbCrLf & _
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
    SQL = SQL & ", EXAMNAME"                        '�˻��
    SQL = SQL & ", SEQNO" & vbCrLf                  '�˻��Ϸù�ȣ"
    SQL = SQL & ", SAMPLETYPE"                      '��ü��"
    SQL = SQL & ", INOUT"                           '��/��
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
    SQL = SQL & "','" & Trim(GetText(vasRes, asRow2, colEXAMNAME))      '�˻��
    SQL = SQL & "','" & Trim(GetText(vasRes, asRow2, colSeq))
    SQL = SQL & "',''"        '��ü��
    SQL = SQL & ",'" & Trim(GetText(vasID, asRow1, colINOUT))
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
    SQL = SQL & "',''"
    SQL = SQL & ",''"
    SQL = SQL & ",''"
    SQL = SQL & ",''"
    SQL = SQL & ",'1'"
    SQL = SQL & ",''"
    SQL = SQL & ",'" & gIFUser
    SQL = SQL & "','')"
    'MsgBox "2"
    Res = SendQuery(gLocal, SQL)
'    Call SetSQLData("��������", SQL)
    If Res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
End Function


'-- ���� �˻��� ��¥�� Max + 1 ��ȣ�� �����´�
Private Function getMaxTestNum(ByVal strDate As String) As Long

    getMaxTestNum = 1
    
    '-- ���������Ʈ
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

    txtStartNum.Text = "0"
    txtStopNum.Text = "0"

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
    
    If Row < 1 Or Row > vasID.DataRowCnt Then
        Exit Sub
    End If
    
    'Local���� �ҷ�����
    ClearSpread vasRes
    
    
'    lblDate.Caption = Trim(GetText(vasID, Row, colHOSPDATE))
    lsID = Trim(GetText(vasID, Row, colBARCODE))
    lblChangeBar.Caption = lsID
    lblBarcode(0).Caption = lsID
    lblPname(0).Caption = Trim(GetText(vasID, Row, colPNAME))
    lblSaveSeq.Caption = Trim(GetText(vasID, Row, colSAVESEQ))
    lblExamDate.Caption = Trim(GetText(vasID, Row, colEXAMDATE))
    
    If lblSaveSeq.Caption = "" Then
        Exit Sub
    End If
    
    
    '����ڵ�, �˻��ڵ�, �˻��, ���, ����
          SQL = "SELECT EQUIPCODE, EXAMCODE, EXAMNAME, EQUIPRESULT, RESULT, SEQNO, REFFLAG, EXAMSUBCODE " & vbCrLf
    SQL = SQL & "  FROM PATRESULT " & vbCrLf
    SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "'" & vbCrLf
    SQL = SQL & "   AND SAVESEQ = " & lblSaveSeq.Caption & vbCrLf
    SQL = SQL & "   AND BARCODE = '" & lsID & "' " & vbCrLf
    'SQL = SQL & "   AND EXAMDATE = '" & Mid(Trim(GetText(vasID, Row, colOrdDate)), 1, 8) & "' " & vbCrLf
    SQL = SQL & "   AND DISKNO = '" & Trim(GetText(vasID, Row, colINOUT)) & "' " & vbCrLf
    SQL = SQL & " GROUP BY SEQNO, EQUIPCODE, EXAMCODE, EXAMNAME, EQUIPRESULT, RESULT, SEQNO, REFFLAG, EXAMSUBCODE "
    SQL = SQL & " ORDER BY SEQNO * 10"
    
    Set RS = cn.Execute(SQL, , 1)

    If Not RS.EOF = True And Not RS.BOF = True Then
        vasRes.MaxRows = 0
        Do Until RS.EOF
            With vasRes
                .MaxRows = .MaxRows + 1
                SetText vasRes, "0", .MaxRows, colCheckBox
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
    vasRes.RowHeight(-1) = 10
    
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
        If lsSeq = "" Then
            Exit Sub
        End If

        If MsgBox(lsSeq & " �� ����� �����Ͻðڽ��ϱ�?", vbInformation + vbYesNo, "�˸�") = vbNo Then
            Exit Sub
        End If

              SQL = "DELETE FROM PATRESULT " & vbCrLf
        SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf
        SQL = SQL & "   AND BARCODE = '" & lsID & "' " & vbCrLf
        SQL = SQL & "   AND PID = '" & lsPid & "' " & vbCrLf
        SQL = SQL & "   AND SAVESEQ = " & lsSeq & vbCrLf
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Trim(GetText(vasID, iRow, colEXAMDATE)) & "' "
        'SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Format(dtpToday.Value, "yyyymmdd") & "' "
        Res = SendQuery(gLocal, SQL)

        If Res = -1 Then
            SaveQuery SQL
            Exit Sub
        End If

        DeleteRow vasID, iRow, iRow
        vasRes.MaxRows = 0
        blnModify = True

    ElseIf KeyCode = vbKeyReturn Then
        If iCol = colBARCODE Then
            '-- �ٲ� ���ڵ�� ȯ������ �ҷ�����
            Call GetSampleInfoW_NTL(iRow)
            
            lsID = Trim(GetText(vasID, iRow, colBARCODE))
            
            
            '-- ���ڵ� ��ȣ�� ������ Ʋ���ٸ� ������Ʈ
            'If lsID <> lblChangeBar.Caption Then
            If lsID <> lblBarcode(0).Caption Then
                      SQL = "UPDATE PATRESULT SET"
                SQL = SQL & " HOSPDATE = '" & Format(Mid(Trim(GetText(vasID, iRow, colHOSPDATE)), 1, 10), "yyyymmdd") & "' " & vbCrLf
                SQL = SQL & ",BARCODE = '" & lsID & "' " & vbCrLf
                SQL = SQL & ",CHARTNO = '" & Trim(GetText(vasID, iRow, colCHARTNO)) & "' " & vbCrLf
                SQL = SQL & ",PID = '" & Trim(GetText(vasID, iRow, colPID)) & "' " & vbCrLf
                SQL = SQL & ",PNAME = '" & Trim(GetText(vasID, iRow, colPNAME)) & "' " & vbCrLf
                SQL = SQL & ",PSEX = '" & Trim(GetText(vasID, iRow, colPSEX)) & "' " & vbCrLf
                SQL = SQL & ",PAGE = '" & Trim(GetText(vasID, iRow, colPAGE)) & "' " & vbCrLf
                SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf
                SQL = SQL & "   AND SAVESEQ = " & Trim(GetText(vasID, iRow, colSAVESEQ)) & vbCrLf
                'SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Trim(GetText(vasID, iRow, colEXAMDATE)) & "' " & vbCrLf
                SQL = SQL & "   AND BARCODE = '" & lblBarcode(0).Caption & "' "

                'SetRawData "[SQL]" & SQL
                Res = SendQuery(gLocal, SQL)
                
                If Res = -1 Then
                    SaveQuery SQL
                    Exit Sub
                End If

                blnModify = True

            End If
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

