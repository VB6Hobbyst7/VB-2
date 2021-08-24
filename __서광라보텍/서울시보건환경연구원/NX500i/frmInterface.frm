VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInterface 
   Caption         =   "LAB Interface"
   ClientHeight    =   10110
   ClientLeft      =   345
   ClientTop       =   840
   ClientWidth     =   15675
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
   Picture         =   "frmInterface.frx":1272
   ScaleHeight     =   10110
   ScaleWidth      =   15675
   StartUpPosition =   1  '������ ���
   WindowState     =   2  '�ִ�ȭ
   Begin VB.Frame FrmHideControl 
      Caption         =   "HideControl"
      Height          =   9945
      Left            =   16200
      TabIndex        =   23
      Top             =   1890
      Visible         =   0   'False
      Width           =   6975
      Begin FPSpread.vaSpread vasExcel 
         Height          =   1005
         Left            =   4470
         TabIndex        =   77
         Top             =   8580
         Visible         =   0   'False
         Width           =   2535
         _Version        =   393216
         _ExtentX        =   4471
         _ExtentY        =   1773
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
         SpreadDesigner  =   "frmInterface.frx":14F5
      End
      Begin VB.Frame Frame6 
         Height          =   525
         Left            =   0
         TabIndex        =   69
         Top             =   0
         Visible         =   0   'False
         Width           =   5685
         Begin VB.Label Label8 
            Caption         =   "��ü��ȣ :"
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
            Left            =   120
            TabIndex        =   74
            Top             =   210
            Width           =   1200
         End
         Begin VB.Label lblBarcode 
            Caption         =   "12345"
            Height          =   165
            Index           =   0
            Left            =   1395
            TabIndex        =   73
            Top             =   210
            Width           =   1485
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
            Left            =   3000
            TabIndex        =   72
            Top             =   210
            Width           =   945
         End
         Begin VB.Label lblPname 
            Caption         =   "1234567890ab"
            Height          =   225
            Index           =   0
            Left            =   4050
            TabIndex        =   71
            Top             =   210
            Width           =   1305
         End
         Begin VB.Label Label3 
            BackColor       =   &H80000008&
            ForeColor       =   &H8000000E&
            Height          =   315
            Left            =   180
            TabIndex        =   70
            Top             =   720
            Width           =   1155
         End
      End
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
         Left            =   4020
         TabIndex        =   67
         Top             =   7710
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
         Left            =   3540
         TabIndex        =   66
         Top             =   7740
         Visible         =   0   'False
         Width           =   285
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
         Left            =   6150
         TabIndex        =   65
         Top             =   5580
         Width           =   735
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
         Left            =   4860
         TabIndex        =   64
         Top             =   5550
         Width           =   1185
      End
      Begin VB.ComboBox cboTest 
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
         Height          =   300
         ItemData        =   "frmInterface.frx":171B
         Left            =   4980
         List            =   "frmInterface.frx":171D
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   63
         Top             =   4200
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.TextBox txtStartNum 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   3450
         TabIndex        =   62
         Text            =   "000001"
         Top             =   4230
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.TextBox txtStopNum 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   4260
         TabIndex        =   61
         Text            =   "009999"
         Top             =   4230
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.CheckBox chkQC 
         Alignment       =   1  '������ ����
         Caption         =   "QC�� ��ȸ"
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
         Left            =   5100
         TabIndex        =   60
         Top             =   4980
         Visible         =   0   'False
         Width           =   1245
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
         Left            =   6210
         TabIndex        =   59
         Text            =   "A"
         Top             =   3510
         Visible         =   0   'False
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
         Left            =   5610
         TabIndex        =   58
         Text            =   "1"
         Top             =   3540
         Visible         =   0   'False
         Width           =   525
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
         ItemData        =   "frmInterface.frx":171F
         Left            =   5430
         List            =   "frmInterface.frx":172C
         TabIndex        =   57
         Top             =   1530
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Frame FrmUseControl 
         Caption         =   "UseControl"
         Height          =   870
         Left            =   1860
         TabIndex        =   46
         Top             =   2310
         Width           =   2835
         Begin VB.Timer tmrReceive 
            Enabled         =   0   'False
            Interval        =   100
            Left            =   1740
            Top             =   300
         End
         Begin VB.Timer tmrSend 
            Enabled         =   0   'False
            Interval        =   100
            Left            =   2220
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
                  Picture         =   "frmInterface.frx":1742
                  Key             =   "RUN"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":1CDC
                  Key             =   "NOT"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":2276
                  Key             =   "STOP"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":2810
                  Key             =   "LST"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":30A2
                  Key             =   "ITM"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":31FC
                  Key             =   "ERR"
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":3356
                  Key             =   "NOF"
               EndProperty
            EndProperty
         End
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
         TabIndex        =   45
         Top             =   3180
         Value           =   1  'Ȯ��
         Width           =   1065
      End
      Begin VB.TextBox txtErr 
         ForeColor       =   &H000000FF&
         Height          =   585
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  '�����
         TabIndex        =   44
         Top             =   3840
         Width           =   1635
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
         TabIndex        =   43
         Top             =   3705
         Width           =   645
      End
      Begin VB.TextBox txtTemp 
         Height          =   435
         Left            =   2730
         TabIndex        =   42
         Top             =   3690
         Width           =   645
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
         TabIndex        =   41
         Top             =   3240
         Width           =   1665
      End
      Begin VB.CommandButton lblclear 
         Caption         =   "lblclear"
         Height          =   495
         Left            =   180
         TabIndex        =   40
         Top             =   4560
         Width           =   1215
      End
      Begin VB.PictureBox picLogin 
         Appearance      =   0  '���
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1530
         Picture         =   "frmInterface.frx":34B0
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   39
         Top             =   4710
         Width           =   285
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
         TabIndex        =   36
         Top             =   3210
         Value           =   1  'Ȯ��
         Width           =   1065
      End
      Begin VB.Frame Frame4 
         Caption         =   "Print"
         Height          =   2415
         Left            =   180
         TabIndex        =   33
         Top             =   5670
         Width           =   3045
         Begin FPSpread.vaSpread vasPrint 
            Height          =   1035
            Left            =   120
            TabIndex        =   34
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
            SpreadDesigner  =   "frmInterface.frx":3A3A
         End
         Begin FPSpread.vaSpread vasPrintBuf 
            Height          =   975
            Left            =   120
            TabIndex        =   35
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
            SpreadDesigner  =   "frmInterface.frx":5493
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Error Log"
         Height          =   945
         Left            =   180
         TabIndex        =   31
         Top             =   8190
         Width           =   4530
         Begin VB.TextBox txtErrLog 
            Appearance      =   0  '���
            Height          =   615
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  '����
            TabIndex        =   32
            Top             =   240
            Width           =   4275
         End
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
         TabIndex        =   30
         Top             =   5160
         Value           =   -1  'True
         Width           =   735
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
         TabIndex        =   29
         Top             =   5160
         Width           =   735
      End
      Begin VB.TextBox txtTest 
         Height          =   1485
         Left            =   3450
         MultiLine       =   -1  'True
         TabIndex        =   28
         Top             =   6150
         Visible         =   0   'False
         Width           =   2145
      End
      Begin VB.CommandButton Command16 
         Caption         =   "�����׽�Ʈ"
         Height          =   435
         Left            =   3510
         TabIndex        =   27
         Top             =   5670
         Width           =   1215
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
         Left            =   5400
         TabIndex        =   26
         Top             =   330
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.CheckBox chkGum 
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
         Left            =   5460
         TabIndex        =   25
         Top             =   2160
         Visible         =   0   'False
         Width           =   855
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
         Left            =   5430
         TabIndex        =   24
         Top             =   2820
         Visible         =   0   'False
         Width           =   1035
      End
      Begin FPSpread.vaSpread vasCode 
         Height          =   945
         Left            =   120
         TabIndex        =   37
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
         SpreadDesigner  =   "frmInterface.frx":56B9
      End
      Begin FPSpread.vaSpread vasTemp1 
         Height          =   945
         Left            =   1860
         TabIndex        =   38
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
         SpreadDesigner  =   "frmInterface.frx":58DF
      End
      Begin FPSpread.vaSpread vasList 
         Height          =   975
         Left            =   120
         TabIndex        =   47
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
         SpreadDesigner  =   "frmInterface.frx":5B05
      End
      Begin FPSpread.vaSpread vasResTemp 
         Height          =   1035
         Left            =   1860
         TabIndex        =   48
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
         SpreadDesigner  =   "frmInterface.frx":5D2B
      End
      Begin FPSpread.vaSpread vasTemp 
         Height          =   975
         Left            =   120
         TabIndex        =   49
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
         SpreadDesigner  =   "frmInterface.frx":5F51
      End
      Begin VB.Label lblChangePID 
         BackColor       =   &H000000FF&
         Height          =   405
         Left            =   3390
         TabIndex        =   55
         Top             =   4650
         Width           =   435
      End
      Begin VB.Label lblChangeBar 
         BackColor       =   &H000000FF&
         Height          =   405
         Left            =   2880
         TabIndex        =   54
         Top             =   4650
         Width           =   465
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
         TabIndex        =   53
         Top             =   4680
         Width           =   825
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
         TabIndex        =   52
         Top             =   5250
         Width           =   1005
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
         TabIndex        =   51
         Top             =   5250
         Width           =   615
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
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  '�� ����
      BackColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   15615
      TabIndex        =   15
      Top             =   0
      Width           =   15675
      Begin MSComCtl2.DTPicker dtpToday 
         Height          =   315
         Left            =   1050
         TabIndex        =   16
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
         Format          =   126091264
         CurrentDate     =   40457
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "����"
         Height          =   195
         Left            =   14670
         TabIndex        =   22
         Top             =   120
         Width           =   420
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "�۽�"
         Height          =   195
         Left            =   13545
         TabIndex        =   21
         Top             =   120
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "��Ʈ"
         Height          =   195
         Index           =   0
         Left            =   12360
         TabIndex        =   20
         Top             =   120
         Width           =   420
      End
      Begin VB.Image imgReceive 
         Height          =   240
         Left            =   15180
         Picture         =   "frmInterface.frx":6177
         Top             =   90
         Width           =   240
      End
      Begin VB.Image imgSend 
         Height          =   240
         Left            =   14025
         Picture         =   "frmInterface.frx":6701
         Top             =   90
         Width           =   240
      End
      Begin VB.Image imgPort 
         Height          =   240
         Left            =   12870
         Picture         =   "frmInterface.frx":6C8B
         Top             =   90
         Width           =   240
      End
      Begin VB.Label lblMachNm 
         Appearance      =   0  '���
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '����
         Caption         =   "HemaVet950"
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
         Index           =   1
         Left            =   3810
         TabIndex        =   19
         Top             =   120
         Width           =   1680
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
         TabIndex        =   18
         Top             =   150
         Width           =   780
      End
      Begin VB.Label lblMachNm 
         Appearance      =   0  '���
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '����
         Caption         =   "HemaVet950"
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
         Index           =   0
         Left            =   3780
         TabIndex        =   17
         Top             =   90
         Width           =   1680
      End
   End
   Begin VB.Frame Frame1 
      Height          =   8235
      Left            =   90
      TabIndex        =   10
      Top             =   1290
      Width           =   15495
      Begin VB.CheckBox chkWAll 
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   690
         TabIndex        =   12
         Top             =   270
         Width           =   225
      End
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
         TabIndex        =   11
         Top             =   210
         Width           =   495
      End
      Begin FPSpread.vaSpread vasID 
         Height          =   7995
         Left            =   90
         TabIndex        =   13
         Top             =   180
         Width           =   8475
         _Version        =   393216
         _ExtentX        =   14949
         _ExtentY        =   14102
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
         SpreadDesigner  =   "frmInterface.frx":7215
      End
      Begin FPSpread.vaSpread vasRes 
         Height          =   8010
         Left            =   8670
         TabIndex        =   14
         Top             =   180
         Width           =   6645
         _Version        =   393216
         _ExtentX        =   11721
         _ExtentY        =   14129
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
         SpreadDesigner  =   "frmInterface.frx":7E83
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  '�� ����
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   15615
      TabIndex        =   0
      Top             =   525
      Width           =   15675
      Begin VB.Frame Frame3 
         Height          =   735
         Left            =   90
         TabIndex        =   1
         Top             =   -60
         Width           =   15495
         Begin VB.OptionButton optPrint 
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
            Height          =   180
            Index           =   1
            Left            =   6810
            TabIndex        =   76
            Top             =   420
            Width           =   705
         End
         Begin VB.OptionButton optPrint 
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
            Height          =   225
            Index           =   0
            Left            =   6810
            TabIndex        =   75
            Top             =   180
            Value           =   -1  'True
            Width           =   705
         End
         Begin VB.CommandButton cmdWorkPrint 
            Caption         =   "������"
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
            Left            =   7590
            TabIndex        =   68
            Top             =   150
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
            Left            =   4290
            TabIndex        =   5
            Top             =   150
            Width           =   1155
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
            Left            =   14220
            TabIndex        =   4
            Top             =   150
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
            Left            =   13020
            TabIndex        =   3
            Top             =   150
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.CommandButton cmdExcelExport 
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
            Height          =   495
            Left            =   5520
            TabIndex        =   2
            Top             =   150
            Width           =   1155
         End
         Begin MSComCtl2.DTPicker dtpStopDt 
            Height          =   345
            Left            =   2790
            TabIndex        =   6
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
            Format          =   126091265
            CurrentDate     =   40248
         End
         Begin MSComCtl2.DTPicker dtpStartDt 
            Height          =   345
            Left            =   1170
            TabIndex        =   7
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
            Format          =   126091265
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
            Left            =   180
            TabIndex        =   9
            Top             =   300
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
            Left            =   2610
            TabIndex        =   8
            Top             =   330
            Width           =   105
         End
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '�Ʒ� ����
      Height          =   405
      Left            =   0
      TabIndex        =   56
      Top             =   9705
      Width           =   15675
      _ExtentX        =   27649
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
            TextSave        =   "2017-08-04"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2646
            MinWidth        =   2646
            TextSave        =   "���� 2:47"
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
      Visible         =   0   'False
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

Dim blnLDLCal As Boolean

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
'
'    Dim iRow As Integer
'    Dim j As Integer
'    Dim k As Integer
'
'
'    Dim sCurDate As String
'    Dim sSerDate As String
'    Dim sHead As String
'    Dim sFoot As String
'    Dim sFileName As String
'
'    Dim sA1c As String
'    Dim sIFCC As String
'    Dim seAg As String
'    Dim blnWrite As Variant
'
'    ClearSpread vasPrint
'
'    blnWrite = False
'    vasExcel.MaxRows = vasID.MaxRows
'    'vasExcel.MaxCols = vasID.MaxCols
'    vasExcel.MaxCols = 6
'
'    For iRow = 1 To vasID.DataRowCnt
'        vasID.Row = iRow
'        vasID.Col = 1
'
'        If vasID.Value = 1 Then
'            If blnWrite = False Then
'                For j = 1 To vasID.MaxCols
'                    If j = 3 Or j = 5 Or j = 17 Or j = 18 Or j = 19 Or j = 20 Then
'                        If j = 3 Then k = 1
'                        If j = 5 Then k = 2
'                        If j = 17 Then k = 3
'                        If j = 18 Then k = 4
'                        If j = 19 Then k = 5
'                        If j = 20 Then k = 6
'
'                        'SetText vasExcel, Trim(GetText(vasID, 0, j)), 0, j
'                        SetText vasExcel, Trim(GetText(vasID, 0, j)), 0, k
'                    End If
'                Next
'            End If
'
'            For j = 1 To vasID.MaxCols
'                If j = 3 Or j = 5 Or j = 17 Or j = 18 Or j = 19 Or j = 20 Then
'                    If j = 3 Then k = 1
'                    If j = 5 Then k = 2
'                    If j = 17 Then k = 3
'                    If j = 18 Then k = 4
'                    If j = 19 Then k = 5
'                    If j = 20 Then k = 6
'
'                    'SetText vasExcel, Trim(GetText(vasID, iRow, j)), iRow, j
'                    SetText vasExcel, Trim(GetText(vasID, iRow, j)), iRow, k
'                End If
'            Next
'        End If
'    Next iRow
'
'    If vasExcel.DataRowCnt < 1 Then
'        MsgBox "������ �ڷᰡ �����ϴ�.", , "�� ��"
'        Exit Sub
'    Else
'        CommonDialog1.Filter = "Excel Files (*.xls)|*.xls|All Files (*.*)|*.*"
'        CommonDialog1.ShowSave
'        sFileName = CommonDialog1.Filename
'        SaveExcel sFileName, vasExcel
'        MsgBox "���� ����Ϸ�", vbOKOnly + vbInformation, Me.Caption
'    End If
    

    Dim iRow As Integer
    Dim j, k As Integer
    
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
                    If j = 3 Or j = 4 Or j = 5 Or j = 6 Or j >= 17 Then
                        If j = 3 Then k = 1
                        If j = 4 Then k = 2
                        If j = 5 Then k = 3
                        If j = 6 Then k = 4
                        
                        If j >= 17 Then
                            k = j - 12
                        End If
                        'If j = 18 Then k = 4
                        'If j = 19 Then k = 5
                        'If j = 20 Then k = 6

                        SetText vasPrint, Trim(GetText(vasID, 0, j)), 0, k
                    End If
                Next
            End If

            For j = 1 To vasID.MaxCols
                If j = 3 Or j = 4 Or j = 5 Or j = 6 Or j >= 17 Then
                    If j = 3 Then k = 1
                    If j = 4 Then k = 2
                    If j = 5 Then k = 3
                    If j = 6 Then k = 4
                    If j >= 17 Then
                        k = j - 12
                    End If
                        
                    SetText vasPrint, Trim(GetText(vasID, iRow, j)), iRow, k
                End If
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
    
    txtRack.Text = "1"
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



Private Sub cmdResult_Click()
    Dim sSeq As String
    Dim sBarcode As String
    Dim strEqpResult As String
    Dim strLisResult As String
    Dim strIntBase As String
    Dim lsExamCode As String
    Dim lsExamName As String
    Dim lsSeqNo As String
    Dim lsResRow    As String
    Dim lsEquipRes As String
    Dim lsResult_Buff As String
    
    Dim lRow As Integer
    Dim lRow1 As Integer
    Dim intRow As Integer
    Dim sWellOld As String
    Dim sWell As String
    Dim sExamCode As String
    Dim sExamName As String
    Dim sEquipCode As String
    Dim sItemCode As String
    Dim strAge As String
    Dim strSex As String
    Dim strPtno As String
    Dim strPtname As String
    Dim varTmp As Variant
    Dim intTstCnt As Integer
    Dim intCol   As Integer
    
    Screen.MousePointer = 11
    
    vasExcel.MaxRows = 0
    
    Call Excel_Open

    intTstCnt = 0
    
    With vasExcel
        For intRow = 2 To .DataRowCnt
            
            .GetText 6, intRow, varTmp: sSeq = varTmp
            .GetText 2, intRow, varTmp: sWell = varTmp
            If sSeq <> "" Then
                With mResult
                    .BarNo = sSeq
                    .RsltDate = Format(Now, "yyyymmddhhmmss")
                    .RsltSeq = getMaxTestNum(Format(dtpToday, "yyyymmdd"))
                    .RackNo = Val(Mid(sWell, 2))
                    .TubePos = Mid(sWell, 1, 1)
                End With
                
                .GetText 3, intRow, varTmp: strIntBase = varTmp
                '.GetText 10, intRow, varTmp: strEqpResult = varTmp
                
                If strIntBase = "FAM" Then
                    Call SetPatInfo(sSeq)
                    
                    vasID.GetText colBARCODE, gRow, varTmp: sBarcode = varTmp
                    SetText vasID, "Result", gRow, colState
                    
                    '-- ä��
                    .GetText 3, intRow, varTmp: strIntBase = varTmp
    
                    '-- ���
                    .GetText 10, intRow, varTmp: strEqpResult = varTmp
                    
                    If Val(strEqpResult) = 0 Then
                        strLisResult = "Not-Detected"
                    Else
                       ' strLisResult = CSng(strEqpResult)
                        strLisResult = Convert2EXP(strEqpResult, "")
                    End If
                    
                    
                    If strLisResult <> "" Then
                              SQL = "Select examcode, examname, seqno "
                        SQL = SQL & "  From EQPMASTER"
                        SQL = SQL & " Where equipno = '" & gEquip & "' "
                        SQL = SQL & "   AND EXAMNAME = '" & Trim(mGetP(cboTest.Text, 1, "|")) & "'"
                        SQL = SQL & "   and equipcode = '" & strIntBase & "' "
                        SQL = SQL & "   and examcode in (" & gOrderExam & ") "      '"'36721','36722','36723','36724'"
                        
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
                            'lsEquipRes = strLisResult
                            'strLisResult = SetResult(strLisResult, strIntBase)
                            'lsResult_Buff = strLisResult
                        
                            For intCol = colState + 1 To vasID.MaxCols
                                If lsExamCode = gArrEquip(intCol - colState, 3) Then
                                    SetText vasID, strLisResult, gRow, intCol
                                    Exit For
                                End If
                            Next
                        
                            SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      '����ڵ�
                            SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '�˻��ڵ�
                            SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '�˻��
                            SetText vasRes, lsEquipRes, lsResRow, colMachResult     '�����
                            SetText vasRes, strLisResult, lsResRow, colRESULT          '���
                            SetText vasRes, lsSeqNo, lsResRow, colSeq               '����
'                            SetText vasRes, strComm, lsResRow, 7                    'Flag
                            '-- ���� ����
                            SetLocalDB gRow, lsResRow, "1", lsEquipRes
                            
                            lsResult_Buff = ""
    
                        Else
                            '-- ���� ���� ���
                                  SQL = "Select examcode, examname, seqno "
                            SQL = SQL & "  From EQPMASTER"
                            SQL = SQL & " Where equipno = '" & gEquip & "' "
                            SQL = SQL & "   AND EXAMNAME = '" & Trim(mGetP(cboTest.Text, 1, "|")) & "'"
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
                                'lsEquipRes = strLisResult
                                'strLisResult = SetResult(strLisResult, strIntBase)
                                'lsResult_Buff = strLisResult
                                
                                For intCol = colState + 1 To vasID.MaxCols
                                    If lsExamCode = gArrEquip(intCol - colState, 3) Then
                                        SetText vasID, strLisResult, gRow, intCol
                                        Exit For
                                    End If
                                Next
'
                                SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      '����ڵ�
                                SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '�˻��ڵ�
                                SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '�˻��
                                SetText vasRes, lsEquipRes, lsResRow, colMachResult     '�����
                                SetText vasRes, strLisResult, lsResRow, colRESULT          '���
                                SetText vasRes, lsSeqNo, lsResRow, colSeq               '����
'                                SetText vasRes, strComm, lsResRow, 7                    'Flag
                                '-- ���� ����
                                SetLocalDB gRow, lsResRow, "1", lsEquipRes
                                
                                
                                lsResult_Buff = ""
                                strState = ""
                            End If
                        End If
                    End If
    
                    strState = "R"
                End If
            End If
        Next
    End With

    Screen.MousePointer = 0

End Sub


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
          
          SQL = " SELECT '', SAVESEQ, MID(EXAMDATE,1,8) AS EXAMDATE, HOSPDATE AS ��������, BARCODE AS ���ڵ��ȣ, CHARTNO AS ��Ʈ��ȣ, PID AS ������ȣ, PNAME AS �̸�,PSEX AS ����, PAGE AS ����, DISKNO, POSNO, EXAMCODE, RESULT, REFFLAG, SENDFLAG,INOUT " & vbCrLf
    SQL = SQL & "   FROM PATRESULT " & vbCrLf
    SQL = SQL & "  WHERE MID(EXAMDATE,1,8) Between '" & Format(dtpStartDt, "YYYYMMDD") & "' AND '" & Format(dtpStopDt, "YYYYMMDD") & "'" & vbCrLf
    SQL = SQL & "    AND EQUIPNO = '" & gEquip & "' " & vbCrLf
    If chkQC.Value = "1" Then
        SQL = SQL & "     AND MID(BARCODE,1,2) = 'QC' "
        SQL = SQL & "     AND EQUIPCODE IN ('pH','pCO2','pO2','cCa2+')"
    End If
    SQL = SQL & " ORDER BY EXAMDATE,SAVESEQ,HOSPDATE,BARCODE "
    
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
        Case "JBUNIV":      Call GetWorkList_JBUNIV(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))
        Case "MSINFOTEC":   Call GetWorkList_MSINFOTEC(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))
        Case "JWINFO":      Call GetWorkList_JWINFO(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))
        Case "AMIS":        Call GetWorkList_AMIS(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))
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
    SQL = SQL & "              AND f72m.testcd  = '" & mGetP(cboTest.Text, 2, "|") & "'   " & vbCr
    SQL = SQL & "              AND r010m.testcd = f72m.testcd            " & vbCr
    SQL = SQL & " WHERE j011m.colldt BETWEEN '" & pFrDt & "000000" & "' AND '" & pToDt & "235959" & "'  " & vbCr
    SQL = SQL & "   and r010m.wkno between '" & txtStartNum.Text & "' AND '" & txtStopNum.Text & "' " & vbCr
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
    SQL = SQL & "                AND f72m.testcd  = '" & mGetP(cboTest.Text, 2, "|") & "'  " & vbCr
    SQL = SQL & "               AND r010m.testcd = f72m.testcd       " & vbCr
    SQL = SQL & "  WHERE j011m.colldt BETWEEN '" & pFrDt & "000000" & "' AND '" & pToDt & "235959" & "'  " & vbCr
    SQL = SQL & "   and r010m.wkno BETWEEN '" & txtStartNum.Text & "' AND '" & txtStopNum.Text & "' " & vbCr
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
                    SetText vasID, "1", .MaxRows, colCheckBox
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



Private Sub GetWorkList_MSINFOTEC(ByVal pFrDt As String, ByVal pToDt As String, Optional pBarNo As String)
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
    
                '-- ó������,ó���Ϸù�ȣ,ȯ�ڸ�,��ü��ȣ,�Կܱ���,�Ϸù�ȣ,����,����,������ȣ,ó���ڵ�
    SQL = ""
    SQL = SQL & "Select DISTINCT a.ORDT as ��������,'0',b.PANM as �̸�,a.SPNO as ���ڵ��ȣ,a.PAID as íƮ��ȣ, a.OIFL,'0',b.SEXS as ����,b.AGES as ����,a.NWNO as ������ȣ,a.ORCD as ITEM " & vbCr
    SQL = SQL & "  From LRESULT a, APATINF b" & vbCr
    SQL = SQL & " Where a.ORDT between  '" & pFrDt & "' and '" & pToDt & "'" & vbCr
    SQL = SQL & "   And a.PAID = b.PAID " & vbCr
    SQL = SQL & "   And a.ORCD in (" & gAllExam & ")" & vbCr
    SQL = SQL & "   And a.OKFL <> 'Y' "   '-- ���Ȯ������

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
'                    For intCol = colState + 1 To vasID.MaxCols
'                        If Trim(RS.Fields("ITEM")) = gArrEquip(intCol - colState, 3) Then
'                            vasID.Row = .MaxRows
'                            vasID.Col = intCol
'                            vasID.BackColor = vbYellow
'                            Exit For
'                        End If
'                    Next
                Next
                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    SetText vasID, "1", .MaxRows, colCheckBox
                    SetText vasID, Trim(RS.Fields("��������")) & "", .MaxRows, colHOSPDATE
                    SetText vasID, Trim(RS.Fields("���ڵ��ȣ")) & "", .MaxRows, colBARCODE
                    SetText vasID, Trim(RS.Fields("íƮ��ȣ")) & "", .MaxRows, colCHARTNO
                    SetText vasID, Trim(RS.Fields("������ȣ")) & "", .MaxRows, colPID
                    SetText vasID, Trim(RS.Fields("�̸�")) & "", .MaxRows, colPNAME
                    SetText vasID, Trim(RS.Fields("����")) & "", .MaxRows, colPSEX
                    SetText vasID, Trim(RS.Fields("����")) & "", .MaxRows, colPAGE
                    'SetText vasID, Trim(RS.Fields("SPCCD")) & "", .MaxRows, colINOUT
                    SetText vasID, txtRack.Text, .MaxRows, colDISKNO
                    
                    '.MaxRows = .MaxRows + 1
                    
                    SetText vasID, txtRack.Text, .MaxRows, colDISKNO
                    SetText vasID, txtPos.Text, .MaxRows, colPOSNO
                    
                    txtRack.Text = txtRack.Text + 1
                    
                    If txtRack.Text = "31" Then
                        txtRack.Text = "1"
                    End If
                    
                    
'                    For intCol = colState + 1 To vasID.MaxCols
'                        If Trim(RS.Fields("ITEM")) = gArrEquip(intCol - colState, 3) Then
'                            vasID.Row = .MaxRows
'                            vasID.Col = intCol
'                            vasID.BackColor = vbYellow
'                            Exit For
'                        End If
'                    Next
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


Private Sub GetWorkList_JWINFO(ByVal pFrDt As String, ByVal pToDt As String, Optional pBarNo As String)
    Dim RS          As ADODB.Recordset
    Dim i           As Integer
    Dim iCnt        As Long
    Dim intRow      As Long
    Dim intCol      As Integer
    Dim strDate     As String
    Dim strChart    As String
    Dim blnSame     As Boolean
    
'    If pBarNo = "" Then
'        vasID.MaxRows = 0
'        intRow = 0
'    End If
    
    blnSame = False
    vasID.ReDraw = False
    

'          SQL = "SELECT DISTINCT RECEIPTDATE as ��������, SPECIMENNUM as ���ڵ��ȣ, RECEIPTNO as íƮ��ȣ, IPDOPD, PTNO as ������ȣ, SNAME as �̸�, LABCODE as ITEM,ORDERCODE,RECEIPTTIME"
'    SQL = SQL & vbCrLf & "  FROM SLA_LabMaster "
'    SQL = SQL & vbCrLf & " WHERE RECEIPTDATE between '" & Format(pFrDt, "####-##-##") & "' and '" & Format(pToDt, "####-##-##") & "'"
'    SQL = SQL & vbCrLf & "   AND LABCODE IN (" & gAllExam & ") "
'    SQL = SQL & vbCrLf & "   AND JSTATUS < '3'" & vbLf
'    '�Կ�
'    If cboChk.ListIndex = 1 Then
'        SQL = SQL & "    AND IpdOpd = 1 "
'    '�ܷ�
'    ElseIf cboChk.ListIndex = 2 Then
'        SQL = SQL & "    AND IpdOpd = 0 "
'    End If
'
'    SQL = SQL & "  ORDER BY RECEIPTDATE,RECEIPTTIME "

          SQL = "SELECT DISTINCT a.RECEIPTDATE as ��������, a.SPECIMENNUM as ���ڵ��ȣ, a.RECEIPTNO as íƮ��ȣ, a.IPDOPD, a.PTNO as ������ȣ, a.SNAME as �̸�, a.LABCODE as ITEM,a.ORDERCODE,a.RECEIPTTIME"
    SQL = SQL & vbCrLf & "  FROM SLA_LabMaster a, SLA_LabResult b"
    SQL = SQL & vbCrLf & " WHERE a.RECEIPTDATE between '" & Format(pFrDt, "####-##-##") & "' and '" & Format(pToDt, "####-##-##") & "'"
    SQL = SQL & vbCrLf & "   AND a.LABCODE IN (" & gAllExam & ") "
    SQL = SQL & vbCrLf & "   AND a.JSTATUS < '2'" & vbLf
    SQL = SQL & vbCrLf & "   AND a.ReceiptNo = b.ReceiptNo"
    SQL = SQL & vbCrLf & "   AND a.OrderCode = b.OrderCode"
    
'    '�Կ�
'    If cboChk.ListIndex = 1 Then
'        SQL = SQL & "    AND a.IpdOpd = 1 "
'    '�ܷ�
'    ElseIf cboChk.ListIndex = 2 Then
'        SQL = SQL & "    AND a.IpdOpd = 0 "
'    End If
    SQL = SQL & "  AND (b.Result = '' OR b.Result IS NULL)"

    SQL = SQL & "  ORDER BY a.RECEIPTDATE, a.RECEIPTTIME "

    Call SetSQLData("��ũ��ȸ", SQL)

    '-- Record Count ������
    cn_Ser.CursorLocation = adUseClient
    Set RS = cn_Ser.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then
'        frmProgress.Show
'        frmProgress.ZOrder 0
'        frmProgress.Xprog.Min = 1
'        frmProgress.Xprog.Max = RS.RecordCount + 1
        
        Do Until RS.EOF
            iCnt = iCnt + 1
            With vasID
                .ReDraw = False
                For i = 1 To .DataRowCnt
                    strDate = GetText(vasID, i, colHOSPDATE)
                    strChart = GetText(vasID, i, colCHARTNO)
                    If Trim(RS("��������")) = strDate And Trim(RS("íƮ��ȣ")) = strChart Then
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
                    SetText vasID, "0", .MaxRows, colCheckBox
                    SetText vasID, Trim(RS.Fields("��������")) & "", .MaxRows, colHOSPDATE
                    If Trim(RS.Fields("���ڵ��ȣ")) & "" = "0" Then
                        SetText vasID, Trim(RS.Fields("íƮ��ȣ")) & "", .MaxRows, colBARCODE
                    Else
                        SetText vasID, Trim(RS.Fields("���ڵ��ȣ")) & "", .MaxRows, colBARCODE
                    End If
                    SetText vasID, Trim(RS.Fields("íƮ��ȣ")) & "", .MaxRows, colCHARTNO
                    SetText vasID, Trim(RS.Fields("������ȣ")) & "", .MaxRows, colPID
                    SetText vasID, Trim(RS.Fields("�̸�")) & "", .MaxRows, colPNAME
                    SetText vasID, IIf(Trim(RS.Fields("IPDOPD")) = 1, "�Կ�", "�ܷ�"), .MaxRows, colINOUT
                    SetText vasID, Trim(RS.Fields("ORDERCODE")) & "", .MaxRows, colPSEX
                    
                  
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
        'chkWAll.Value = "1"
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




Private Sub GetWorkList_AMIS(ByVal pFrDt As String, ByVal pToDt As String, Optional pBarNo As String)
    Dim RS          As ADODB.Recordset
    Dim i           As Integer
    Dim iCnt        As Long
    Dim intRow      As Long
    Dim intCol      As Integer
    Dim strDate     As String
    Dim strChart    As String
    Dim blnSame     As Boolean
    
    blnSame = False
    vasID.ReDraw = False
    
'          SQL = "SELECT DISTINCT a.RECEIPTDATE as ��������, a.SPECIMENNUM as ���ڵ��ȣ, a.RECEIPTNO as íƮ��ȣ, a.IPDOPD, a.PTNO as ������ȣ, a.SNAME as �̸�, a.LABCODE as ITEM,a.ORDERCODE,a.RECEIPTTIME"
'    SQL = SQL & vbCrLf & "  FROM SLA_LabMaster a, SLA_LabResult b"
'    SQL = SQL & vbCrLf & " WHERE a.RECEIPTDATE between '" & Format(pFrDt, "####-##-##") & "' and '" & Format(pToDt, "####-##-##") & "'"
'    SQL = SQL & vbCrLf & "   AND a.LABCODE IN (" & gAllExam & ") "
'    SQL = SQL & vbCrLf & "   AND a.JSTATUS < '2'" & vbLf
'    SQL = SQL & vbCrLf & "   AND a.ReceiptNo = b.ReceiptNo"
'    SQL = SQL & vbCrLf & "   AND a.OrderCode = b.OrderCode"
'    SQL = SQL & "  AND (b.Result = '' OR b.Result IS NULL)"
'    SQL = SQL & "  ORDER BY a.RECEIPTDATE, a.RECEIPTTIME "


    SQL = ""
    SQL = SQL & "SELECT P.PATID as íƮ��ȣ, P.PATNAME as �̸�, P.SEX, O.ACPTDATE as ��������, O.ACPTSEQ, O.RSVACPTSTATE, O.RESULTSTATE, O.DEPTCODE, O.ORDERDATE, O.SLIPNO, O.IOFLAG, O.ORDERCODE, O.ORDERNAME, R.SPCMNO as ���ڵ��ȣ, R.RESULTFLAG, R.RESULTNO, R.RESULTITEMCODE as ITEM " & vbCr
    SQL = SQL & "  FROM registinfos O, resultofnum R, PATMST P " & vbCr
    SQL = SQL & " WHERE O.acptdate = R.acptdate " & vbCr
    SQL = SQL & "   AND O.acptdate between '" & pFrDt & "' and '" & pToDt & "'"
    SQL = SQL & "   AND R.resultitemcode in (" & gAllExam & ")"
    SQL = SQL & "   AND O.patid = R.patid " & vbCr
    SQL = SQL & "   AND O.acptseq = R.acptseq " & vbCr
    SQL = SQL & "   AND O.patid = P.patid " & vbCr
    SQL = SQL & "   AND O.CLAS = 4 " & vbCr '�ӻ󺴸�
    SQL = SQL & "   AND O.ORDERCODE = 'LC8006' "
    If chkSaveAll.Value = "0" Then
        SQL = SQL & "   AND R.RESULTFLAG = 0 " & vbCr
    End If
    SQL = SQL & "  ORDER BY R.SPCMNO"
    
'    ���ڵ�� ��ȸ
'    SQL = ""
'    SQL = SQL & "SELECT P.PATID, P.PATNAME, P.SEX, O.ACPTDATE, O.ACPTSEQ, O.RSVACPTSTATE, O.RESULTSTATE, O.DEPTCODE, O.ORDERDATE, O.SLIPNO, O.IOFLAG, O.ORDERCODE, O.ORDERNAME, R.SPCMNO, R.RESULTFLAG, R.RESULTNO, R.RESULTITEMCODE " & vbCr
'    SQL = SQL & "  FROM registinfos O, resultofnum R, PATMST P " & vbCr
'    SQL = SQL & " WHERE o.acptdate = r.acptdate " & vbCr
'    SQL = SQL & "   AND o.patid = r.patid " & vbCr
'    SQL = SQL & "   AND o.acptseq = r.acptseq " & vbCr
'    SQL = SQL & "   AND O.patid = p.patid " & vbCr
'    SQL = SQL & "   AND O.CLAS = 4 " & vbCr '�ӻ󺴸�
'    SQL = SQL & "   AND R.SPCMNO = '''+TMaster.FBarCode+''' ';"


    Call SetSQLData("��ũ��ȸ", SQL)

    '-- Record Count ������
    cn_Ser.CursorLocation = adUseClient
    Set RS = cn_Ser.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then
        Do Until RS.EOF
            iCnt = iCnt + 1
            With vasID
                .ReDraw = False
                For i = 1 To .DataRowCnt
                    strDate = GetText(vasID, i, colHOSPDATE)
                    strChart = GetText(vasID, i, colCHARTNO)
                    If Trim(RS("��������")) = strDate And Trim(RS("íƮ��ȣ")) = strChart Then
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
                    If Trim(RS.Fields("���ڵ��ȣ")) & "" = "0" Then
                        SetText vasID, Trim(RS.Fields("íƮ��ȣ")) & "", .MaxRows, colBARCODE
                    Else
                        SetText vasID, Trim(RS.Fields("���ڵ��ȣ")) & "", .MaxRows, colBARCODE
                    End If
                    SetText vasID, Trim(RS.Fields("íƮ��ȣ")) & "", .MaxRows, colCHARTNO
                    'SetText vasID, Trim(RS.Fields("������ȣ")) & "", .MaxRows, colPID
                    SetText vasID, Trim(RS.Fields("�̸�")) & "", .MaxRows, colPNAME
                    'SetText vasID, IIf(Trim(RS.Fields("IPDOPD")) = 1, "�Կ�", "�ܷ�"), .MaxRows, colINOUT
                    'SetText vasID, Trim(RS.Fields("ORDERCODE")) & "", .MaxRows, colPSEX
                    
                  
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
            'frmProgress.Xprog.Value = iCnt
            DoEvents
                        
            RS.MoveNext
        Loop
        'chkWAll.Value = "1"
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
        'vasID.Width = 15225
        vasID.Width = Frame1.Width - 200
    Else
        cmdSL.Caption = "��"
        'vasID.Width = 8475
        vasID.Width = Me.Width - Frame6.Width - 710
    End If

    Call Form_Resize
    
End Sub

Private Sub cmdWorkPrint_Click()
'Dim iRow As Integer
'Dim j As Integer
'
'Dim sCurDate As String
'Dim sSerDate As String
'Dim sHead As String
'Dim sFoot As String
'
'    ClearSpread vasPrint
'
'    j = 1
'
'    vasPrint.RowHeight(-1) = 25.9
'
'    For iRow = 1 To vasID.DataRowCnt
'        vasID.Row = iRow
'        vasID.Col = colCheckBox
'
'        If vasID.Value = 1 Then
'            SetText vasPrint, Trim(GetText(vasID, iRow, colBARCODE)), j, 1     '��ü��ȣ
'            SetText vasPrint, Trim(GetText(vasID, iRow, colCHARTNO)), j, 2     'ȯ�ڹ�ȣ
'            SetText vasPrint, Trim(GetText(vasID, iRow, colPNAME)), j, 3     'ȯ���̸�
'
'            SetText vasPrint, Trim(GetText(vasID, iRow, colPSEX)), j, 4     '����
'            SetText vasPrint, Trim(GetText(vasID, iRow, colPAGE)), j, 5     '����
'            SetText vasPrint, Trim(GetText(vasID, iRow, colHOSPDATE)), j, 7     'ó������
'            SetText vasPrint, Trim(GetText(vasID, iRow, colHOSPDATE)), j, 8     'ó������
'
'            j = j + 1
'        End If
'    Next iRow
'
'    If vasPrint.DataRowCnt < 1 Then
'        MsgBox "����� �ڷᰡ �����ϴ�.", , "�� ��"
'        Exit Sub
'    End If
'
'    sCurDate = GetDateFull
'
'    sSerDate = Trim(dtpStartDt.Value) & " - " & Trim(dtpStopDt.Value)
'
'    vasPrint.PrintOrientation = 1   ' SS_PRINTORIENT_PORTRAIT
'    vasPrint.PrintAbortMsg = "�μ��� �Դϴ� ..."
'    vasPrint.PrintJobName = "WorkList ���"
'
'
'    sHead = "/fn""�ü�ü"" /fz""12"" /fb1 /fi0 /fu0 " & "/c" & "�� WorkList ��" & "/n/n " & _
'            "/fn""����ü"" /fz""10"" /fb0 /fi0 /fu0 " & "/c" & "ó������ : " & dtpStartDt & " ~ " & dtpStopDt
'
'    sFoot = "/fn""����ü"" /fz""10"" /fb1 /fi0 /fu0 " & "/l" & sCurDate & "/fn""�ü�ü"" /fz""11"" /fb1 /fi0 /fu0 /r" & " �˻��"
'
'    vasPrint.PrintHeader = sHead
'    vasPrint.PrintFooter = sFoot
'
'    vasPrint.PrintMarginTop = 680
'    vasPrint.PrintMarginBottom = 680
''���� SS�� ���Ī���� �����
''    vaslist.PrintMarginLeft = 720
'    vasPrint.PrintMarginLeft = 0
'    vasPrint.PrintMarginRight = 0
'
'    vasPrint.PrintColor = True
'    vasPrint.PrintGrid = True
'
''Set printing range
'    vasPrint.PrintType = 0  'SS_PRINT_ALL(default)
'
'    vasPrint.PrintShadows = True
'
'    vasPrint.Action = 13 'SS_ACTION_PRINT


    Dim iRow As Integer
    Dim i, j, k As Integer
    
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
                    If j = 3 Or j = 4 Or j = 5 Or j = 6 Or j >= 17 Then
                        If j = 3 Then k = 1
                        If j = 4 Then k = 2
                        If j = 5 Then k = 3
                        If j = 6 Then k = 4
                        
                        If j >= 17 Then
                            k = j - 12
                        End If
                        'If j = 18 Then k = 4
                        'If j = 19 Then k = 5
                        'If j = 20 Then k = 6

                        SetText vasPrint, Trim(GetText(vasID, 0, j)), 0, k
                    End If
                Next
            End If

            For j = 1 To vasID.MaxCols
                If j = 3 Or j = 4 Or j = 5 Or j = 6 Or j >= 17 Then
                    If j = 3 Then k = 1
                    If j = 4 Then k = 2
                    If j = 5 Then k = 3
                    If j = 6 Then k = 4
                    If j >= 17 Then
                        k = j - 12
                    End If
                        
                    SetText vasPrint, Trim(GetText(vasID, iRow, j)), iRow, k
                End If
            Next
        End If
    Next iRow
    
    
    
    If vasPrint.DataRowCnt < 1 Then
        MsgBox "����� �ڷᰡ �����ϴ�.", , "�� ��"
        Exit Sub
    Else
        If optPrint(0).Value = True Then
            For i = 6 To vasPrint.MaxCols
                vasPrint.ColWidth(i) = 5
            Next
            
            vasPrint.PrintOrientation = PrintOrientationLandscape '�������
            vasPrint.Action = 13
        Else
            vasPrint.PrintOrientation = PrintOrientationPortrait '�������
            vasPrint.Action = 13
        End If
        MsgBox "��� ��¿Ϸ�", vbOKOnly + vbInformation, Me.Caption
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
        vasRes.Height = vasID.Height '550
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
    
    
    imgPort.Picture = imlStatus.ListImages("NOT").ExtractIcon
    imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
    imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
    
    cmdIFClear_Click
    lblclear_Click
    
    GetSetup
    
    lblMachNm(0).Caption = gEquip
    lblMachNm(1).Caption = gEquip
    
    
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
'        fraWork.Visible = False
    
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
    
    'Call cmdSL_Click
    
'    StatusBar1.Panels(2).Text = Winsock1.LocalIP
    '-- test
'    vasID.MaxRows = 10


'    Me.Height = 10890
'    Me.Width = 16000
    
Exit Sub
Rst:
    If Err = 8002 Then      'Port
        MsgBox "��� ��Ʈ�� Ȯ���ϼ���!", vbExclamation, "�˸�"
        
        frmConfig.Show 1  '/��ż���
'        Me.Height = 10890
'        Me.Width = 16000
        
        Call GetExamCode
    Else
        Resume Next
    End If
    
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
            .ColWidth(colState + (i + 1)) = 6
            
            cboTest.AddItem gArrEquip(i + 1, 4) & Space(20) & "|" & gArrEquip(i + 1, 3)
        Next
        
        cboTest.ListIndex = 0
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
        Case -1  '## EOT
            strState = ""
            comEqp.Output = EOT
            SetRawData "[Tx]" & EOT
            intFrameNo = 1
            
            Exit Sub
        Case 0  '## Header
            '## Header
            strOutput = "H|\^&|||" & vbCr
            
            '## Patient
            strOutput = strOutput & "P|1||" & Format$(mOrder.BarNo, String$(12, "@")) & vbCr

            '## Order
            If mOrder.NoOrder = False Then
                '## ���������� �ִ°��: ��ü����, ���������� �����Ѱ��� ����!
                strOutput = strOutput & "O|1|" & Format$(mOrder.BarNo, String$(12, "@")) & "||" & mOrder.Order & "|||||||||||S||||||||||X" & vbCr
            Else
                '## ���������� ���°��: �˻��׸� ������ ������ ����!
                strOutput = strOutput & "O|1|" & Format$(mOrder.BarNo, String$(12, "@")) & "|||||||||||||S||||||||||X" & vbCr
            End If
            
            '## Termianator
            strOutput = strOutput & "L|1|N" & vbCr
            strOutput = intFrameNo & strOutput
            
            intFrameNo = intFrameNo + 1
            
        Case 1      '## ���� ������ ���� ���ڿ��� �ִ� ���
            strOutput = intFrameNo & mOrder.Order


    End Select
    
    If Len(strOutput) >= 230 Then
        mOrder.Order = Mid$(strOutput, 231)
        strOutput = Mid$(strOutput, 1, 230) & ETB
        intSndPhase = 1
    Else
        strOutput = strOutput & ETX
        intSndPhase = -1
    End If
    
    
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

            SetRawData "[Rx]" & Buffer
            StatusBar1.Panels(3).Text = Buffer
            
            lngBufLen = Len(Buffer)
            
            
            For i = 1 To lngBufLen
                BufChar = Mid$(Buffer, i, 1)
                Select Case BufChar
                    Case STX
                        strBuffer = ""
                    Case ETX
                        '-- ��񿡼� �Ѿ�� �ð��� �쿬�� 11:59:59�ʳ� ���Ͽ� ����� �ð��� ���
                        '-- ��� ����� �������� ������ �� �����Ƿ� ��¥�� �ǽð� ������Ʈ �Ѵ�.
                        strDate = DateCompare(Format(CDate(dtpToday.Value), "yyyymmdd"))
                        dtpToday.Value = Format(strDate, "####-##-##")
                        
                        DoEvents
                        
                        If gMach = "FDC7000" Then
                            Call EditRcvData_FDC7000
                            strBuffer = ""
                        ElseIf gMach = "FDC7000i" Then
                            Call EditRcvData_FDC7000i
                            strBuffer = ""
                        ElseIf gMach = "HemaVet950" Then
                            Call EditRcvData_HemaVet950
                            strBuffer = ""
                        ElseIf gMach = "AU10V" Then
                            Call EditRcvData_AU10V
                            strBuffer = ""
                        ElseIf gMach = "NX500" Then
                            Call EditRcvData_NX500
                            strBuffer = ""
                        ElseIf gMach = "NX500i" Then
                            Call EditRcvData_NX500
                            strBuffer = ""
                        Else
                            Call EditRcvDataASTM
                            strBuffer = ""
                        End If
                    Case Else
                        strBuffer = strBuffer & BufChar
                        
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
    'Call SetText(vasID, mOrder.RackNo, intRow, colDISKNO)       '-- Rack
    'Call SetText(vasID, mOrder.TubePos, intRow, colPOSNO)       '-- Pos
    
    '-- ȯ������ ǥ��
    Call vasActiveCell(vasID, intRow, colBARCODE)
    
    '-- ����������� �����
    Call ClearSpread(vasRes)
    
    '-- �˻��� ���� ��������
    Call GetSampleInfoW_AMIS(intRow)
    
    '-- ���ڵ��ȣ�� �ش��ϴ� �˻��ڵ� ��������
    'gOrderExam = GetOrderExamCode(gEquip, pBarNo)

    '-- �������̺��� �˻��׸� �ش��ϴ� �˻�ä�� ã�ƿ��� (intRow = ���� �˻��ߴ� ���ڵ尡 �ٽ� �ö�� ��� ��ġ�� ��ã�´�.)
    strItems = GetEquipExamCode_GEMINI(gEquip, pBarNo, intRow)

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
Function GetEquipExamCode_GEMINI(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim i As Integer
    Dim strExamCode As String
    Dim sBarcode     As String
    Dim strCBC As String
    Dim strDiff As String
    
    GetEquipExamCode_GEMINI = ""
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If
    
    sBarcode = Trim(GetText(frmInterface.vasID, intRow, colBARCODE))   '2 ���� ���ڵ� ��ȣ
    'SetRawData "[sBarcode]" & sBarcode
    
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

    
    For i = 0 To UBound(gReadBuf)
        If gReadBuf(i) <> "" Then
            strExamCode = strExamCode & "^^^" & Trim(gReadBuf(i)) & "^\"
        End If
    Next
    
    If strExamCode <> "" Then
        strExamCode = Mid(strExamCode, 1, Len(strExamCode) - 1)
    End If
    
    GetEquipExamCode_GEMINI = strExamCode
    
End Function

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
    
'    For i = 1 To vasID.DataRowCnt
'        If Trim(GetText(vasID, i, colBARCODE)) = pBarNo Then
'            intRow = i
'            Exit For
'        End If
'    Next i

    If intRow < 0 Then
        intRow = vasID.DataRowCnt + 1
        If vasID.MaxRows < intRow Then
            vasID.MaxRows = intRow
        End If
    End If
    
    
    '-- ���������� ǥ��
    Call SetText(vasID, "1", intRow, colCheckBox)
    If pBarNo = "" Then
        Call SetText(vasID, mResult.PatNo, intRow, colBARCODE)
        Call SetText(vasID, mResult.PatNo, intRow, colCHARTNO)
    Else
        Call SetText(vasID, mResult.BarNo, intRow, colBARCODE)
        Call SetText(vasID, mResult.PatNo, intRow, colCHARTNO)
    End If
'    Call SetText(vasID, mResult.RackNo, intRow, colDISKNO)
    Call SetText(vasID, mResult.SpcmNo, intRow, colHOSPDATE)
    Call SetText(vasID, mResult.RsltDate, intRow, colEXAMDATE)
    Call SetText(vasID, mResult.RsltSeq, intRow, colSAVESEQ)
    Call SetText(vasID, mResult.SpcPos, intRow, colCHARTNO)
    
    
    
    
    Call vasActiveCell(vasID, intRow, colBARCODE)
    
    '-- ����������� �����
    Call ClearSpread(vasRes)
    
    '-- �˻��� ���� �������̺��� ������ ǥ��(for ��ũ����Ʈ)  '6,7,8,9
'    Call GetSampleInfoW_AMIS(intRow)
    
    '-- ���� Row
    gRow = intRow
    
End Sub

'-----------------------------------------------------------------------------'
'   ��� : ���κ� ������ ������ ����
'-----------------------------------------------------------------------------'
Private Sub EditRcvData_NX500()
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
    Dim blnSame As Boolean
    
    Dim strSndBuffer As String
    Dim strSaveSeq As String
    Dim strSaveDate As String
    
    Dim strTC As String
    Dim strTG As String
    Dim strHDL As String
    Dim strTestDt As String
    
    strRcvBuf = strBuffer
                
    strType = Mid$(strRcvBuf, 1, 1)
    If IsNumeric(strType) Then
        strType = Mid$(strRcvBuf, 2, 1)
    End If
    '12345678901234567890123456789012345678901234567890
    'NORMAL 2017-05-1523:596                         01GGT-P  =1        U/l   1  @         
    
    '-- Type1 �϶� ���(������û)
    If UCase(strType) = "W" Then
        strBarNo = Trim(mGetP(strRcvBuf, 2, ","))
        
    ElseIf UCase(strType) = "N" Then
        strBarNo = ""
        strTestDt = Trim(Mid(strRcvBuf, 8, 10))
        strBarNo = Trim(Mid(strRcvBuf, 30, 16))
        strSeq = Trim(Mid(strRcvBuf, 23, 1))
    
        If strBarNo = "" Then
            strBarNo = strTestDt
            'Exit Sub
        End If
        
        With mResult
            .BarNo = strBarNo
            .SpcPos = strSeq
            .RsltDate = Format(Now, "yyyymmddhhmmss")
            .RsltSeq = getMaxTestNum(Format(dtpToday, "yyyymmdd"))
        End With
    

        
        Call SetPatInfo(strBarNo)

        If gRow < 0 Then
            Exit Sub
        End If
                    
        For i = 51 To Len(strRcvBuf) Step 36
            '## ������ �˻��, ���
            strIntBase = Trim(Mid(strRcvBuf, i, 7))
            strResult = Trim(Mid(strRcvBuf, i + 8, 8))
        
            strResult = Replace(strResult, "=", "")
            strResult = Replace(strResult, "  ", " ")
            strResult = Replace(strResult, "  ", " ")
            strResult = Replace(strResult, "  ", " ")
        
            If strIntBase = "TCHO-P" Then
                strTC = strResult
            End If
            
            If strIntBase = "TG-P" Then
                strTG = strResult
            End If
            
            If strIntBase = "HDLC-P" Then
                strHDL = strResult
            End If
        
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
                    
                    '-- vasID �� ǥ��
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
                    SetText vasRes, strComm, lsResRow, colFLAG                    'Flag
                    '-- ���� ����
                    SetLocalDB gRow, lsResRow, "1", lsEquipRes
                                
                    lsResult_Buff = ""
                    
                    If strState <> "R" Then
                        strState = ""
                    End If
                    
                '-- ���� ���� ���
                Else
                    SQL = ""
                    SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
                    SQL = SQL & "  FROM EQPMASTER"
                    SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
                    SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
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
                        
                        '-- vasID �� ǥ��
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
        Next
    
'''        If strTC <> "" And strTG <> "" And strHDL <> "" And IsNumeric(strTC) And IsNumeric(strTG) And IsNumeric(strHDL) Then
'''            blnLDLCal = False
'''            strIntBase = "LDL-P"
'''            strResult = strTC - ((strTG / 5) + strHDL)
'''            If strResult < 0 Then
'''                strResult = "0"
'''            End If
'''            'MsgBox strResult
'''            SQL = ""
'''            SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
'''            SQL = SQL & "  FROM EQPMASTER"
'''            SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
'''            SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
'''            SQL = SQL & "   AND EXAMCODE in (" & gOrderExam & ") "
'''            'SQL = SQL & "   AND EXAMCODE in ('C3791','C3792','C3793') "
'''            Res = GetDBSelectColumn(gLocal, SQL)
'''
'''            '-- ���� ���� ���
'''            If Res > 0 Then
'''                lsExamCode = Trim(gReadBuf(0))
'''                lsExamName = Trim(gReadBuf(1))
'''                lsSeqNo = Trim(gReadBuf(2))
'''
'''                lsResRow = vasRes.DataRowCnt + 1
'''                If vasRes.MaxRows < lsResRow Then
'''                    vasRes.MaxRows = lsResRow
'''                End If
'''
'''                '�Ҽ��� ó��, ��� ���� ó��
'''                lsEquipRes = strResult
'''                'strResult = SetResult(strResult, strIntBase)
'''                lsResult_Buff = strResult
'''
'''                '-- Work List
'''                SetText vasID, "Result", gRow, colState                 '11 �������
'''
'''
'''                '-- ��� List
'''                SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      '����ڵ�
'''                SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '�˻��ڵ�
'''                SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '�˻��
'''                SetText vasRes, lsEquipRes, lsResRow, colMachResult     '�����
'''                SetText vasRes, strResult, lsResRow, colRESULT          '���
'''                SetText vasRes, lsSeqNo, lsResRow, colSeq               '����
'''                SetText vasRes, strComm, lsResRow, 7                    'Flag
'''                '-- ���� ����
'''                SetLocalDB gRow, lsResRow, "1", lsEquipRes
'''
'''                lsResult_Buff = ""
'''
'''                strState = "R"
'''
'''            '-- ���� ���� ���
'''            Else
'''
'''                      SQL = "Select examcode, examname, seqno "
'''                SQL = SQL & "  From EQPMASTER"
'''                SQL = SQL & " Where equipno = '" & gEquip & "' "
'''                SQL = SQL & "   and equipcode = '" & strIntBase & "' "
'''                Res = GetDBSelectColumn(gLocal, SQL)
'''
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
'''                    'strResult = SetResult(strResult, strIntBase)
'''                    lsResult_Buff = strResult
'''
'''                    '-- Work List
'''                    SetText vasID, "Result", gRow, colState                 '�������
'''
'''                    '-- ��� List
'''                    SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      '����ڵ�
'''                    SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '�˻��ڵ�
'''                    SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '�˻��
'''                    SetText vasRes, lsEquipRes, lsResRow, colMachResult     '�����
'''                    SetText vasRes, strResult, lsResRow, colRESULT          '���
'''                    SetText vasRes, lsSeqNo, lsResRow, colSeq               '����
'''                    SetText vasRes, strComm, lsResRow, colFLAG              'Flag
'''                    '-- ���� ����
'''                    SetLocalDB gRow, lsResRow, "1", lsEquipRes
'''
'''                    lsResult_Buff = ""
'''                    strState = "R"
'''                End If
'''            End If
'''        End If
    
    
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
    End If
    
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
    Dim varBuffer   As Variant
    Dim strTestDt   As String
    Dim strTestTm   As String
    Dim strTestNo   As String
    
    varRcvBuf = Split(strBuffer, vbCrLf)
    
    For i = 0 To UBound(varRcvBuf)
        strRcvBuf = varRcvBuf(i)
        
        If i = 0 Then
            strBarNo = Trim$(mGetP(strRcvBuf, 5, ","))   '-- Patient id
            strSeqNo = Trim$(mGetP(strRcvBuf, 7, ","))  '-- Test No
            strTestDt = Trim$(mGetP(strRcvBuf, 8, ","))
            strTestDt = Format(strTestDt, "yyyy-mm-dd")
            
            '-- ������ ���ȭ�� �ʱ�ȭ
            vasRes.MaxRows = 0
            
            If strBarNo <> "" Then
                With mResult
                    .BarNo = strBarNo
                    .PatNo = strSeqNo
                    .RsltDate = Format(Now, "yyyymmddhhmmss")
                    .RsltSeq = getMaxTestNum(Format(dtpToday, "yyyymmdd"))
                    .SpcmNo = strTestDt '�������
                End With
                        
                Call SetPatInfo(strBarNo)
                
                If gRow <= 0 Then
                    Exit Sub
                End If
            End If
        Else
            strIntBase = Trim$(mGetP(varRcvBuf(i), 1, ","))
            strResult = Trim$(mGetP(varRcvBuf(i), 2, ","))
            
            If strIntBase <> "" And strResult <> "" Then
                SQL = ""
                SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
                SQL = SQL & "  FROM EQPMASTER"
                SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
                SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
                
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
                    
                    '-- vasID �� ǥ��
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
                    SetText vasRes, strComm, lsResRow, 7                    'Flag
                    '-- ���� ����
                    SetLocalDB gRow, lsResRow, "1", lsEquipRes
                                
                    lsResult_Buff = ""
                    
                    If strState <> "R" Then
                        strState = ""
                    End If
                End If
            End If
        End If
        
        SetText vasID, "Result", gRow, colState
        vasRes.RowHeight(-1) = 14
    Next

End Sub

'-----------------------------------------------------------------------------'
'   ��� : ���κ� ������ ������ ����
'-----------------------------------------------------------------------------'
Private Sub EditRcvData_AU10V()
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
    Dim varBuffer   As Variant
    Dim strTestDt   As String
    Dim strTestTm   As String
    Dim strTestNo   As String
    
    varRcvBuf = Split(strBuffer, vbCrLf)
    
    'For i = 0 To UBound(varRcvBuf)
    strRcvBuf = strBuffer
    strType = Mid(strRcvBuf, 1, 1)
    
    Select Case strType
        Case "N"    '## Normal result
    
            strSeqNo = Trim$(Mid(strRcvBuf, 23, 13))
            strBarNo = Trim$(Mid(strRcvBuf, 36, 13))
            strTestDt = Trim$(Mid(strRcvBuf, 8, 10)) & " " & Trim$(Mid(strRcvBuf, 18, 5))
            strTestDt = Format(strTestDt, "yyyy-mm-dd")
            If strBarNo = "" Then
                strBarNo = strSeqNo
            End If
            'strDevice = Trim$(Mid(strRcvBuf, 49, 2))
            
            '-- ������ ���ȭ�� �ʱ�ȭ
            vasRes.MaxRows = 0
            
            If strBarNo <> "" Then
                With mResult
                    .BarNo = strBarNo
                    .PatNo = strSeqNo
                    .RsltDate = Format(Now, "yyyymmddhhmmss")
                    .RsltSeq = getMaxTestNum(Format(dtpToday, "yyyymmdd"))
                    .SpcmNo = strTestDt '�������
                End With
                        
                Call SetPatInfo(strBarNo)
                
                If gRow <= 0 Then
                    Exit Sub
                End If
            End If

            strIntBase = Trim(Mid(strRcvBuf, 51, 7))
            strResult = Trim(Mid(strRcvBuf, 58, 10))
            strResult = Replace(strResult, "=", "")
            strResult = Replace(strResult, "  ", " ")
            strResult = Replace(strResult, "  ", " ")
            strResult = Replace(strResult, "  ", " ")
            strResult = strResult & " " & Trim(Mid(strRcvBuf, 68, 5))
            
            If strIntBase <> "" And strResult <> "" Then
                SQL = ""
                SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
                SQL = SQL & "  FROM EQPMASTER"
                SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
                SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
                
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
                    
                    '-- vasID �� ǥ��
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
                    SetText vasRes, strComm, lsResRow, 7                    'Flag
                    '-- ���� ����
                    SetLocalDB gRow, lsResRow, "1", lsEquipRes
                                
                    lsResult_Buff = ""
                    
                    If strState <> "R" Then
                        strState = ""
                    End If
                End If
            End If
        
            SetText vasID, "Result", gRow, colState
            vasRes.RowHeight(-1) = 14
        
        Case Else
        
    End Select

End Sub

'-----------------------------------------------------------------------------'
'   ��� : ���κ� ������ ������ ����
'-----------------------------------------------------------------------------'
Private Sub EditRcvData_HemaVet950()
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
    Dim varBuffer   As Variant
    Dim strTestDt   As String
    Dim strTestTm   As String
    Dim strTestNo   As String
    
    varRcvBuf = Split(strBuffer, vbCrLf)
    
    For i = 0 To UBound(varRcvBuf)
        strRcvBuf = varRcvBuf(i)
        
        If i = 0 Then
            strBarNo = Trim$(mGetP(strRcvBuf, 5, ","))   '-- Patient id
            strSeqNo = Trim$(mGetP(strRcvBuf, 7, ","))  '-- Test No
            strTestDt = Trim$(mGetP(strRcvBuf, 8, ","))
            strTestDt = Format(strTestDt, "yyyy-mm-dd")
            
            '-- ������ ���ȭ�� �ʱ�ȭ
            vasRes.MaxRows = 0
            
            If strBarNo <> "" Then
                With mResult
                    .BarNo = strBarNo
                    .PatNo = strSeqNo
                    .RsltDate = Format(Now, "yyyymmddhhmmss")
                    .RsltSeq = getMaxTestNum(Format(dtpToday, "yyyymmdd"))
                    .SpcmNo = strTestDt '�������
                End With
                        
                Call SetPatInfo(strBarNo)
                
                If gRow <= 0 Then
                    Exit Sub
                End If
            End If
        Else
            strIntBase = Trim$(mGetP(varRcvBuf(i), 1, ","))
            strResult = Trim$(mGetP(varRcvBuf(i), 2, ","))
            
            If strIntBase <> "" And strResult <> "" Then
                SQL = ""
                SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
                SQL = SQL & "  FROM EQPMASTER"
                SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
                SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
                
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
                    
                    '-- vasID �� ǥ��
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
                    SetText vasRes, strComm, lsResRow, 7                    'Flag
                    '-- ���� ����
                    SetLocalDB gRow, lsResRow, "1", lsEquipRes
                                
                    lsResult_Buff = ""
                    
                    If strState <> "R" Then
                        strState = ""
                    End If
                End If
            End If
        End If
        
        SetText vasID, "Result", gRow, colState
        vasRes.RowHeight(-1) = 14
    Next

End Sub


'-----------------------------------------------------------------------------'
'   ��� : ���κ� ������ ������ ����
'-----------------------------------------------------------------------------'
Private Sub EditRcvData_FDC7000()
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
    Dim varBuffer   As Variant
    Dim strTestDt   As String
    Dim strTestTm   As String
    
    Dim strTC As String
    Dim strTG As String
    Dim strHDL As String
    
    strRcvBuf = strBuffer
    
    strType = mGetP(strRcvBuf, 1, ",")
            
    Select Case strType
        Case "R"
            '-- ������ ���ȭ�� �ʱ�ȭ
            vasRes.MaxRows = 0
            
            strTestDt = Trim(mGetP(strRcvBuf, 3, ","))
            strTestTm = Trim(mGetP(strRcvBuf, 4, ","))
            strSeqNo = Trim(mGetP(strRcvBuf, 5, ","))
            strBarNo = Trim(mGetP(strRcvBuf, 6, ","))
            
            
            If strBarNo <> "" Then
                With mResult
                    .BarNo = strBarNo
                    .PatNo = strSeqNo
                    .RsltDate = Format(Now, "yyyymmddhhmmss")
                    .RsltSeq = getMaxTestNum(Format(dtpToday, "yyyymmdd"))
                    .SpcmNo = strTestDt '�������
                End With
                        
                Call SetPatInfo(strBarNo)
                
                If gRow <= 0 Then
                    Exit Sub
                End If
            End If
            
            For i = 13 To Len(strRcvBuf) Step 7
                strIntBase = Trim(mGetP(strRcvBuf, i, ","))
                strIntBase = mGetP(strIntBase, 1, "-")
                '-- ���� , ���� ����
                'strResult = Trim(mGetP(strRcvBuf, i + 1, ",")) & Trim(mGetP(strRcvBuf, i + 2, ","))
                strResult = Trim(mGetP(strRcvBuf, i + 1, ",")) & Mid(Trim(mGetP(strRcvBuf, i + 2, ",")), 1, 8)
                
                strResult = Replace(strResult, "=", "")
                strResult = Replace(strResult, "  ", " ")
                strResult = Replace(strResult, "  ", " ")
                strResult = Replace(strResult, "  ", " ")
            
                strResult = Trim(strResult)
                
                If strIntBase = "TCHO" Then
                    strTC = strResult
                    MsgBox "strTC:" & strTC
                End If
                
                If strIntBase = "TG" Then
                    strTG = strResult
                    MsgBox "strTG:" & strTG
                End If
                
                If strIntBase = "HDLC" Then
                    strHDL = strResult
                    MsgBox "strHDL:" & strHDL
                End If
                
                
                If strIntBase <> "" And strResult <> "" Then
                    SQL = ""
                    SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
                    SQL = SQL & "  FROM EQPMASTER"
                    SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
                    SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
                    
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
                        
                        '-- vasID �� ǥ��
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
                        SetText vasRes, strComm, lsResRow, 7                    'Flag
                        '-- ���� ����
                        SetLocalDB gRow, lsResRow, "1", lsEquipRes
                                    
                        lsResult_Buff = ""
                        
                        If strState <> "R" Then
                            strState = ""
                        End If
                    End If
                End If
            Next
            
            'LDL ���� ����
            If strTC <> "" And strTG <> "" And strHDL <> "" And IsNumeric(strTC) And IsNumeric(strTG) And IsNumeric(strHDL) Then
                MsgBox "1"
                blnLDLCal = False
                strIntBase = "LDL-CAL"
                strResult = strTC - ((strTG / 5) + strHDL)
                If strResult < 0 Then
                    strResult = "0"
                End If
                
                If strIntBase <> "" And strResult <> "" Then
                    SQL = ""
                    SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
                    SQL = SQL & "  FROM EQPMASTER"
                    SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
                    SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
                    
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
                        
                        '-- vasID �� ǥ��
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
                        SetText vasRes, strComm, lsResRow, 7                    'Flag
                        '-- ���� ����
                        SetLocalDB gRow, lsResRow, "1", lsEquipRes
                                    
                        lsResult_Buff = ""
                        
                        If strState <> "R" Then
                            strState = ""
                        End If
                    End If
                End If
                
            End If
        
            
            SetText vasID, "Result", gRow, colState
            vasRes.RowHeight(-1) = 14
    End Select

End Sub

'-----------------------------------------------------------------------------'
'   ��� : ���κ� ������ ������ ����
'-----------------------------------------------------------------------------'
Private Sub EditRcvData_FDC7000i()
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
    Dim varBuffer   As Variant
    Dim strTestDt   As String
    Dim strTestTm   As String
    
    Dim strTC As String
    Dim strTG As String
    Dim strHDL As String
    
    strRcvBuf = strBuffer
    
    strType = Mid(strRcvBuf, 1, 1)
            
    Select Case strType
        Case "N"
            '-- ������ ���ȭ�� �ʱ�ȭ
            vasRes.MaxRows = 0
            
            'NORMAL 2002-03-0521:299            1512150228   03GGT-P  =67       U/l   01           CPK-P  =121      U/l   01           CRE-P  =1.0      mg/dl 01           
            'NORMAL 2002-03-0521:3110           1512150228   03BUN-P  =10.0     mg/dl 01           TBIL-P =0.5      mg/dl 01           LDH-P  =193      U/l   01           ALB-P  =5.0      g/dl  01           HDLC-P =42       mg/dl 01           TP-P   =7.6      g/dl  01           GLU-P  =76       mg/dl 01           GOT-P  =35       U/l   01           TG-P   =410      mg/dl 01H          GPT-P  =48       U/l   01H          TCHO-P =236      mg/dl 01H          ALP-P  =250      U/l   01           

            
            strTestDt = Trim(Mid(strRcvBuf, 8, 10))
            strTestTm = Trim(Mid(strRcvBuf, 18, 5))
            strSeqNo = Trim(Mid(strRcvBuf, 23, 5))
            strBarNo = Trim(Mid(strRcvBuf, 30, 16))
            
            If strBarNo = "" Then
                strBarNo = strSeqNo
            End If
            
            If strBarNo <> "" Then
                With mResult
                    .BarNo = strBarNo
                    .PatNo = strSeqNo
                    .RsltDate = Format(Now, "yyyymmddhhmmss")
                    .RsltSeq = getMaxTestNum(Format(dtpToday, "yyyymmdd"))
                    .SpcmNo = strTestDt '�������
                End With
                        
                Call SetPatInfo(strBarNo)
                
                If gRow <= 0 Then
                    Exit Sub
                End If
            End If
            

            
            For i = 51 To Len(strRcvBuf) Step 36
                strIntBase = Trim(Mid(strRcvBuf, i, 7))
                strIntBase = mGetP(strIntBase, 1, "-")
                strResult = Trim(Mid(strRcvBuf, i + 8, 8))
            
                If strIntBase = "TCHO" Then
                    strTC = strResult
                End If
                
                If strIntBase = "TG" Then
                    strTG = strResult
                End If
                
                If strIntBase = "HDLC" Then
                    strHDL = strResult
                End If
            
            
                If strIntBase <> "" And strResult <> "" Then
                    SQL = ""
                    SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
                    SQL = SQL & "  FROM EQPMASTER"
                    SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
                    SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
                    
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
                        
                        '-- vasID �� ǥ��
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
                        SetText vasRes, strComm, lsResRow, 7                    'Flag
                        '-- ���� ����
                        SetLocalDB gRow, lsResRow, "1", lsEquipRes
                                    
                        lsResult_Buff = ""
                        
                        If strState <> "R" Then
                            strState = ""
                        End If
                    End If
                End If
            Next
            
            'LDL ���� ����
            If strTC <> "" And strTG <> "" And strHDL <> "" And IsNumeric(strTC) And IsNumeric(strTG) And IsNumeric(strHDL) Then
                blnLDLCal = False
                strIntBase = "LDL-CAL"
                strResult = strTC - ((strTG / 5) + strHDL)
                If strResult < 0 Then
                    strResult = "0"
                End If
                
                If strIntBase <> "" And strResult <> "" Then
                    SQL = ""
                    SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
                    SQL = SQL & "  FROM EQPMASTER"
                    SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
                    SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
                    
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
                        
                        '-- vasID �� ǥ��
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
                        SetText vasRes, strComm, lsResRow, 7                    'Flag
                        '-- ���� ����
                        SetLocalDB gRow, lsResRow, "1", lsEquipRes
                                    
                        lsResult_Buff = ""
                        
                        If strState <> "R" Then
                            strState = ""
                        End If
                    End If
                End If
                
            End If
            SetText vasID, "Result", gRow, colState
            vasRes.RowHeight(-1) = 14
    End Select

End Sub

'-----------------------------------------------------------------------------'
'   ��� : ���κ� ������ ������ ����
'-----------------------------------------------------------------------------'
Private Sub EditRcvDataASTMi()
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
    Dim varBuffer   As Variant
    Dim strTestDt   As String
    Dim strTestTm   As String
    
    strRcvBuf = strBuffer
    
    strType = Mid(strRcvBuf, 1, 1)
            
    Select Case strType
        Case "N"
            '-- ������ ���ȭ�� �ʱ�ȭ
            vasRes.MaxRows = 0
            
            'NORMAL 2002-03-0521:299            1512150228   03GGT-P  =67       U/l   01           CPK-P  =121      U/l   01           CRE-P  =1.0      mg/dl 01           
            'NORMAL 2002-03-0521:3110           1512150228   03BUN-P  =10.0     mg/dl 01           TBIL-P =0.5      mg/dl 01           LDH-P  =193      U/l   01           ALB-P  =5.0      g/dl  01           HDLC-P =42       mg/dl 01           TP-P   =7.6      g/dl  01           GLU-P  =76       mg/dl 01           GOT-P  =35       U/l   01           TG-P   =410      mg/dl 01H          GPT-P  =48       U/l   01H          TCHO-P =236      mg/dl 01H          ALP-P  =250      U/l   01           

            
            strTestDt = Trim(Mid(strRcvBuf, 8, 10))
            strTestTm = Trim(Mid(strRcvBuf, 18, 5))
            strSeqNo = Trim(Mid(strRcvBuf, 23, 5))
            strBarNo = Trim(Mid(strRcvBuf, 30, 16))
            
            
            If strBarNo <> "" Then
                With mResult
                    .BarNo = strBarNo
                    .PatNo = strSeqNo
                    .RsltDate = Format(Now, "yyyymmddhhmmss")
                    .RsltSeq = getMaxTestNum(Format(dtpToday, "yyyymmdd"))
                    .SpcmNo = strTestDt '�������
                End With
                        
                Call SetPatInfo(strBarNo)
                
                If gRow <= 0 Then
                    Exit Sub
                End If
            End If
            
            For i = 51 To Len(strRcvBuf) Step 36
                strIntBase = Trim(Mid(strRcvBuf, i, 7))
                strIntBase = mGetP(strIntBase, 1, "-")
                strResult = Trim(Mid(strRcvBuf, i + 8, 8))
            
                If strIntBase <> "" And strResult <> "" Then
                    SQL = ""
                    SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
                    SQL = SQL & "  FROM EQPMASTER"
                    SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
                    SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
                    
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
                        
                        '-- vasID �� ǥ��
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
                        SetText vasRes, strComm, lsResRow, 7                    'Flag
                        '-- ���� ����
                        SetLocalDB gRow, lsResRow, "1", lsEquipRes
                                    
                        lsResult_Buff = ""
                        
                        If strState <> "R" Then
                            strState = ""
                        End If
                    End If
                End If
            Next
            
            SetText vasID, "Result", gRow, colState
            vasRes.RowHeight(-1) = 14
    End Select

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
' asRow2 = ��� List
Function SetLocalDB(ByVal asRow1 As Long, ByVal asRow2 As Long, asSend As String, Optional asEquipResult As String = "")
    Dim sCnt As String
    Dim sExamDate As String
    Dim strSaveSeq As String
    
    sExamDate = Format(dtpToday, "yyyymmddhhmmss")
    'sExamDate = Trim(GetText(vasID, asRow1, colOrdDate))
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
   ' SetRawData "[SQL]" & SQL
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
    SQL = SQL & "','" & Trim(GetText(vasID, asRow1, colPSEX))   '-- ORDERCODE ����
    SQL = SQL & "','" & Trim(GetText(vasID, asRow1, colPAGE))
    SQL = SQL & "',''"
    SQL = SQL & ",''"
    SQL = SQL & ",''"
    SQL = SQL & ",'1'"
    SQL = SQL & ",''"
    SQL = SQL & ",'" & gIFUser
    SQL = SQL & "','')"
    
'    SetRawData "[SQL]" & SQL
    Res = SendQuery(gLocal, SQL)
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


Private Sub txtRack_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    
    If KeyAscii = 13 Then
        With vasID
            For i = .ActiveRow To .MaxRows
                .Row = i
                .Col = colDISKNO
                .Text = txtRack.Text
                txtRack.Text = txtRack.Text + 1
                If txtRack.Text = "31" Then
                    txtRack.Text = "1"
                End If
            Next
        End With
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
    
    If lblSaveSeq.Caption = "" Then
        Exit Sub
    End If
    
    'Local���� �ҷ�����
    ClearSpread vasRes
    
    '����ڵ�, �˻��ڵ�, �˻��, ���, ����
          SQL = "SELECT EQUIPCODE, EXAMCODE, EXAMNAME, EQUIPRESULT, RESULT, SEQNO, REFFLAG, EXAMSUBCODE " & vbCrLf
    SQL = SQL & "  FROM PATRESULT " & vbCrLf
    SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "'" & vbCrLf
    SQL = SQL & "   AND SAVESEQ = " & lblSaveSeq.Caption & vbCrLf
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
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Format(dtpToday.Value, "yyyymmdd") & "' "
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
            'Exit Sub
            
            '-- �ٲ� ���ڵ�� ȯ������ �ҷ�����
            Call GetSampleInfoW_AMIS(iRow)
            
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
                SQL = SQL & ",INOUT = '" & Trim(GetText(vasID, iRow, colINOUT)) & "' " & vbCrLf
                SQL = SQL & ",PSEX = '" & Trim(GetText(vasID, iRow, colPSEX)) & "' " & vbCrLf
                SQL = SQL & ",PAGE = '" & Trim(GetText(vasID, iRow, colPAGE)) & "' " & vbCrLf
                SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf
                SQL = SQL & "   AND SAVESEQ = " & Trim(GetText(vasID, iRow, colSAVESEQ)) & vbCrLf
                'SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Trim(GetText(vasID, iRow, colEXAMDATE)) & "' " & vbCrLf
                SQL = SQL & "   AND BARCODE = '" & lblBarcode(0).Caption & "' "

                SetRawData "[SQL]" & SQL
                Res = SendQuery(gLocal, SQL)
                
                If Res = -1 Then
                    SaveQuery SQL
                    Exit Sub
                End If

                blnModify = True

            End If
        ElseIf iCol = colDISKNO Then
            
        
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



