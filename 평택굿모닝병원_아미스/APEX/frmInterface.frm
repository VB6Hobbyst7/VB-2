VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInterface 
   BorderStyle     =   1  '���� ����
   Caption         =   "APEX"
   ClientHeight    =   13530
   ClientLeft      =   330
   ClientTop       =   825
   ClientWidth     =   18960
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
   MaxButton       =   0   'False
   Picture         =   "frmInterface.frx":1272
   ScaleHeight     =   13530
   ScaleWidth      =   18960
   StartUpPosition =   1  '������ ���
   Begin VB.Frame FrmHideControl 
      Caption         =   "HideControl"
      Height          =   9945
      Left            =   21960
      TabIndex        =   2
      Top             =   2520
      Visible         =   0   'False
      Width           =   8295
      Begin VB.TextBox txtBarNum 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
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
         Height          =   360
         Left            =   5130
         TabIndex        =   71
         Top             =   1020
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.Frame Frame6 
         Height          =   585
         Left            =   5160
         TabIndex        =   69
         Top             =   4860
         Visible         =   0   'False
         Width           =   5985
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
         Left            =   5250
         TabIndex        =   64
         Top             =   4290
         Visible         =   0   'False
         Width           =   285
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
         Height          =   495
         Left            =   5880
         TabIndex        =   63
         Top             =   3870
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.TextBox txtTest 
         Height          =   1785
         Left            =   4770
         MultiLine       =   -1  'True
         TabIndex        =   62
         Top             =   1560
         Visible         =   0   'False
         Width           =   3225
      End
      Begin VB.CommandButton Command16 
         Caption         =   "�����׽�Ʈ"
         Height          =   435
         Left            =   6000
         TabIndex        =   61
         Top             =   450
         Visible         =   0   'False
         Width           =   1215
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
         Left            =   3960
         TabIndex        =   59
         Top             =   6720
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
         Left            =   4680
         TabIndex        =   58
         Top             =   6720
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
         Left            =   5340
         List            =   "frmInterface.frx":1502
         TabIndex        =   57
         Top             =   6720
         Visible         =   0   'False
         Width           =   975
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
         Left            =   6450
         TabIndex        =   56
         Top             =   6720
         Visible         =   0   'False
         Width           =   1425
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
         TabIndex        =   39
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
         TabIndex        =   38
         Top             =   5160
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.Frame Frame2 
         Caption         =   "Error Log"
         Height          =   945
         Left            =   180
         TabIndex        =   35
         Top             =   8190
         Width           =   4530
         Begin VB.TextBox txtErrLog 
            Appearance      =   0  '���
            Height          =   615
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  '����
            TabIndex        =   36
            Top             =   240
            Width           =   4275
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Print"
         Height          =   2415
         Left            =   180
         TabIndex        =   32
         Top             =   5670
         Width           =   3045
         Begin FPSpread.vaSpread vasPrint 
            Height          =   1035
            Left            =   120
            TabIndex        =   33
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
            TabIndex        =   34
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
         TabIndex        =   24
         Top             =   3210
         Value           =   1  'Ȯ��
         Width           =   1065
      End
      Begin FPSpread.vaSpread vasCode 
         Height          =   915
         Left            =   120
         TabIndex        =   16
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
         TabIndex        =   3
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
         TabIndex        =   17
         Top             =   4710
         Width           =   285
      End
      Begin VB.CommandButton lblclear 
         Caption         =   "lblclear"
         Height          =   495
         Left            =   180
         TabIndex        =   15
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
         TabIndex        =   9
         Top             =   3240
         Width           =   1665
      End
      Begin VB.TextBox txtTemp 
         Height          =   435
         Left            =   2730
         TabIndex        =   8
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
         TabIndex        =   7
         Top             =   3705
         Width           =   645
      End
      Begin VB.TextBox txtErr 
         ForeColor       =   &H000000FF&
         Height          =   585
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  '�����
         TabIndex        =   6
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
         TabIndex        =   5
         Top             =   3180
         Value           =   1  'Ȯ��
         Width           =   1065
      End
      Begin VB.Frame FrmUseControl 
         Caption         =   "UseControl"
         Height          =   870
         Left            =   1860
         TabIndex        =   4
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
         TabIndex        =   10
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
         TabIndex        =   11
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
         TabIndex        =   12
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
         Left            =   4500
         TabIndex        =   60
         Top             =   6810
         Visible         =   0   'False
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
         TabIndex        =   42
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
         TabIndex        =   41
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
         TabIndex        =   40
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
         TabIndex        =   18
         Top             =   4680
         Width           =   825
      End
      Begin VB.Label lblChangeBar 
         BackColor       =   &H000000FF&
         Height          =   405
         Left            =   2880
         TabIndex        =   14
         Top             =   4650
         Width           =   465
      End
      Begin VB.Label lblChangePID 
         BackColor       =   &H000000FF&
         Height          =   405
         Left            =   3390
         TabIndex        =   13
         Top             =   4650
         Width           =   435
      End
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   120
      TabIndex        =   43
      Top             =   570
      Width           =   18765
      Begin VB.TextBox txtBarcode 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
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
         Height          =   360
         Left            =   5460
         TabIndex        =   0
         Top             =   240
         Width           =   1875
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
         Height          =   495
         Left            =   8670
         TabIndex        =   55
         Top             =   150
         Width           =   1155
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
         Height          =   495
         Left            =   9870
         TabIndex        =   54
         Top             =   150
         Width           =   1155
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
         Height          =   495
         Left            =   14280
         TabIndex        =   53
         Top             =   150
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
         Height          =   495
         Left            =   15720
         TabIndex        =   52
         Top             =   150
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
         Height          =   495
         Left            =   17190
         TabIndex        =   51
         Top             =   150
         Width           =   1395
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "��ȸ"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   4440
         TabIndex        =   45
         Top             =   180
         Width           =   945
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
         Left            =   7440
         TabIndex        =   44
         Top             =   150
         Width           =   1185
      End
      Begin MSComCtl2.DTPicker dtpStopDt 
         Height          =   345
         Left            =   2820
         TabIndex        =   46
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   123076609
         CurrentDate     =   40248
      End
      Begin MSComCtl2.DTPicker dtpStartDt 
         Height          =   345
         Left            =   1110
         TabIndex        =   47
         Top             =   240
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   123076609
         CurrentDate     =   40248
      End
      Begin VB.Label lblBarcode 
         Caption         =   "12345"
         Height          =   165
         Index           =   0
         Left            =   12390
         TabIndex        =   68
         Top             =   180
         Width           =   2235
      End
      Begin VB.Label lblPname 
         Caption         =   "1234567890ab"
         Height          =   225
         Index           =   0
         Left            =   12390
         TabIndex        =   67
         Top             =   450
         Width           =   2055
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
         Left            =   11340
         TabIndex        =   66
         Top             =   450
         Width           =   945
      End
      Begin VB.Label Label8 
         Caption         =   "���ڵ� :"
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
         Left            =   11340
         TabIndex        =   65
         Top             =   180
         Width           =   930
      End
      Begin VB.Label Label20 
         Caption         =   "��ȸ����"
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
         Left            =   180
         TabIndex        =   49
         Top             =   330
         Width           =   795
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
         TabIndex        =   48
         Top             =   330
         Width           =   105
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  '�� ����
      Height          =   810
      Left            =   0
      ScaleHeight     =   750
      ScaleWidth      =   18900
      TabIndex        =   31
      Top             =   525
      Width           =   18960
   End
   Begin VB.Frame Frame1 
      Height          =   11625
      Left            =   30
      TabIndex        =   27
      Top             =   1380
      Width           =   18825
      Begin FPSpread.vaSpread vasRes 
         Height          =   9795
         Left            =   10050
         TabIndex        =   29
         Top             =   180
         Width           =   8625
         _Version        =   393216
         _ExtentX        =   15214
         _ExtentY        =   17277
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
         MaxCols         =   9
         MaxRows         =   10
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "frmInterface.frx":A383
      End
      Begin VB.TextBox txtCmnt 
         Appearance      =   0  '���
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1530
         Left            =   10650
         MultiLine       =   -1  'True
         TabIndex        =   74
         Top             =   10020
         Width           =   8025
      End
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
         TabIndex        =   50
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
         Height          =   11355
         Left            =   90
         TabIndex        =   30
         Top             =   180
         Width           =   9885
         _Version        =   393216
         _ExtentX        =   17436
         _ExtentY        =   20029
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
         SpreadDesigner  =   "frmInterface.frx":AA50
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   24
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   10110
         TabIndex        =   73
         Top             =   10830
         Width           =   480
      End
      Begin VB.Label lblCmnt 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   24
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   10110
         TabIndex        =   72
         Top             =   10230
         Width           =   480
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  '�� ����
      BackColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   18900
      TabIndex        =   19
      Top             =   0
      Width           =   18960
      Begin MSComCtl2.DTPicker dtpToday 
         Height          =   315
         Left            =   1110
         TabIndex        =   25
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
         Format          =   123076608
         CurrentDate     =   40457
      End
      Begin VB.Label Label1 
         Appearance      =   0  '���
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '����
         Caption         =   "APEX"
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
         TabIndex        =   37
         Top             =   90
         Width           =   735
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
         TabIndex        =   26
         Top             =   150
         Width           =   780
      End
      Begin VB.Label Label1 
         Appearance      =   0  '���
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '����
         Caption         =   "APEX"
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
         TabIndex        =   23
         Top             =   120
         Width           =   735
      End
      Begin VB.Image imgPort 
         Height          =   240
         Left            =   13560
         Picture         =   "frmInterface.frx":B638
         Top             =   90
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgSend 
         Height          =   240
         Left            =   14715
         Picture         =   "frmInterface.frx":BBC2
         Top             =   90
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgReceive 
         Height          =   240
         Left            =   15870
         Picture         =   "frmInterface.frx":C14C
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
         Top             =   120
         Visible         =   0   'False
         Width           =   420
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '�Ʒ� ����
      Height          =   405
      Left            =   0
      TabIndex        =   1
      Top             =   13125
      Width           =   18960
      _ExtentX        =   33443
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
            Object.Width           =   21696
            MinWidth        =   21696
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   2646
            MinWidth        =   2646
            TextSave        =   "2017-08-01"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2646
            MinWidth        =   2646
            TextSave        =   "���� 9:00"
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
         Visible         =   0   'False
      End
      Begin VB.Menu MnExamConfig 
         Caption         =   "�ڵ弳��"
      End
      Begin VB.Menu MnCmmtConfig 
         Caption         =   "�Ұ߼���"
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
    txtCmnt = ""
    
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
                    
                    strBody = strBody & vbTab & "<GROUP TYPE=""Sample"">" & vbCrLf
                    strBody = strBody & vbTab & vbTab & "<GROUP TYPE=""Patient"" ID=""" & Trim(GetText(vasID, intRow, colBARCODE)) & """>" & vbCrLf
                    '-- Group
                    strBody = strBody & vbTab & vbTab & vbTab & "<PARAM TYPE=""String[32]"" ID=""Group"" FILTER=""-wwwr"">" & Trim(GetText(vasID, intRow, colINOUT)) & "</PARAM>" & vbCrLf
                    '-- Name
                    strBody = strBody & vbTab & vbTab & vbTab & "<PARAM TYPE=""String[32]"" ID=""Name"" FILTER=""-ooor"">" & Trim(GetText(vasID, intRow, colPNAME)) & "</PARAM>" & vbCrLf
                    '-- Surname
                    strBody = strBody & vbTab & vbTab & vbTab & "<PARAM TYPE=""String[32]"" ID=""Surname"" FILTER=""-ooor"">" & Trim(GetText(vasID, intRow, colPNAME)) & "</PARAM>" & vbCrLf
                    '-- BirthDate
                    strBody = strBody & vbTab & vbTab & vbTab & "<PARAM TYPE=""Date"" ID=""BirthDate"" FILTER=""-ooor"">" & Format(Now, "yyyy-mm-dd") & "</PARAM>" & vbCrLf
                    '-- Sex
                    strBody = strBody & vbTab & vbTab & vbTab & "<PARAM TYPE=""SexType"" ID=""Sex"" FILTER=""-wwwr"">" & IIf(Trim(GetText(vasID, intRow, colPSEX)) = "M", "M", "F") & "</PARAM>" & vbCrLf
                    '-- Note
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
    
    If gAssayNM.CMTVIEW = "1" Then
        Call SetComment
    End If
    
End Sub

Private Function XMLFileOpen(strPath) As String
    Dim myXml As New MSXML2.DOMDocument60
    Dim node1 As IXMLDOMNode
    Dim node2 As IXMLDOMNode
    Dim strMSG As String
    Dim objElem As IXMLDOMNodeList
    Dim i       As Integer
    Dim j       As Integer
    Dim k       As Integer
    Dim l       As Integer
    Dim m       As Integer
    Dim N       As Integer
    Dim blnEdit As Boolean
    
    
   ' On Error Resume Next
    
    N = 0
    myXml.async = False
    If myXml.Load(strPath) = True Then
        Set objElem = myXml.selectNodes("GROUP//GROUP")

        For i = 0 To objElem.Length - 1
            For j = 0 To objElem.Item(i).childNodes.Length - 1
                For k = 0 To objElem.Item(i).childNodes(j).childNodes.Length - 1
                    If objElem.Item(i).childNodes(j).Attributes(k).nodeName = "TYPE" Then
                        If objElem.Item(i).childNodes(j).Attributes(k).nodeValue = "Patient" Then
                            blnEdit = True
                        Else
                            blnEdit = False
                        End If
                    End If
                    
                    If blnEdit = True And objElem.Item(i).childNodes(j).Attributes(k).nodeName = "ID" Then
                        ReDim Preserve strRecvData(N)
                        strRecvData(N) = strRecvData(N) & "P|" & objElem.Item(i).childNodes(j).Attributes(k).nodeValue
                        N = N + 1
                        Exit For
                        'Set node1 = myXml.selectSingleNode("GROUP//GROUP//GROUP//GROUP//GROUP//PARAM")
                        'Debug.Print node1.XML
                    End If
                    'Call EditRcvDataAPEX
                    
                Next
                
            Next
        Next
        Set myXml = Nothing
    Else
      MsgBox "�б⿡��", vbCritical
    End If
    
    For i = 0 To UBound(strRecvData)
        Debug.Print strRecvData(i)
    Next
    
End Function


Private Function f_subSet_XMLWorkList(ByVal strXML As String) As Variant
    Dim strPath   As String
    Dim strBuffer As String
    Dim i         As Long
    Dim lngBufLen As Long
    Dim BufChar   As String
    Dim strTmp As String
    Dim intIdx As Integer
    Dim varTmp  As Variant
    Dim j       As Integer
    Dim k As Integer
    
    Dim blnAppend1  As Boolean
    Dim blnAppend2  As Boolean
    Dim varTmp1    As Variant
    Dim strTest   As String
    Dim strResult As String
    Dim strClass  As String
    Dim strIntensity    As String
    
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
 
    intIdx = 0
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
            
            If i > 30 Then
                strRecvData(j) = strRecvData(j) & "L|1|N"
                j = j + 1
                ReDim Preserve strRecvData(j)
            End If
            
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
        
        'Debug.Print varTmp(i)
        '<PARAM TYPE="Long" ID="Intensity">0
        If InStr(varTmp(i), """Blot""") > 0 Then 'blnAppend1 = True And blnAppend2 = True And
            strIntensity = Mid(varTmp(i), InStr(varTmp(i), """Intensity""") + 12)
            strIntensity = Mid(strIntensity, 1, InStr(strIntensity, "<") - 1)
            varTmp1 = Split(varTmp(i), "</PARAM>")
            strTest = ""
            strResult = ""
            For k = 0 To UBound(varTmp1)
                If InStr(varTmp1(k), """Code""") > 0 Then
                    strTest = Mid(varTmp1(k), InStr(varTmp1(k), """Code""") + 7)
                End If
                
                If strTest <> "" And InStr(varTmp1(k), """QntResult""") > 0 Then
                    strResult = Mid(varTmp1(k), InStr(varTmp1(k), """QntResult""") + 12)
                    strClass = Mid(varTmp1(k + 1), InStr(varTmp1(k + 1), """Result""") + 9)
                    
                    strResult = Replace(strResult, "&lt;", "<")
                    strResult = Replace(strResult, "&gt;", ">")
                    
                    
                    
                    If strIntensity = "-1" Then
                        'Stop
                        strResult = "-"
                    End If
                    
                    ReDim Preserve strRecvData(j)
                    strRecvData(j) = strRecvData(j) & "R|" & strTest & "^" & strResult & "^" & strClass
                    j = j + 1
                    strTest = ""
                    strResult = ""
                End If
            Next
        End If
    Next
    
    If UBound(varTmp) > 0 Then
        ReDim Preserve strRecvData(j)
        strRecvData(j) = strRecvData(j) & "L|1|N"
    End If
    
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
    txtCmnt = ""

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


Private Sub SetComment()
    Dim strDate As String
    Dim strBarNo As String
    
    Dim strSaveSeq As String
    Dim RS          As ADODB.Recordset
    Dim i As Integer

    Dim strClass  As String
    Dim strClass0 As String '-- Total IgE
    Dim strClass1 As String
    Dim strClass2 As String
    Dim strClass3 As String
    Dim strClass4 As String
    Dim strClass5 As String
    Dim strClass6 As String
    Dim intClass  As Integer
    Dim blnIgE    As Boolean
    
    Dim strResult As String
    Dim strIntBase As String
    Dim strGubun   As String
        
    Dim lsExamCode As String
    Dim lsExamName As String
    Dim lsSeqNo    As String
    
    For i = 1 To vasID.DataRowCnt
        strDate = GetText(vasID, i, colHOSPDATE)
        strBarNo = GetText(vasID, i, colBARCODE)
        strSaveSeq = GetText(vasID, i, colSAVESEQ)
        strGubun = GetText(vasID, i, colINOUT)
        txtCmnt = ""
        strClass0 = ""
        strClass1 = ""
        strClass2 = ""
        strClass3 = ""
        strClass4 = ""
        strClass5 = ""
        strClass6 = ""
        intClass = 0
        blnIgE = False
        
              SQL = " SELECT DISTINCT SEQNO, EQUIPCODE,EXAMCODE, RESULT, REFVALUE " & vbCrLf
        SQL = SQL & "   FROM PATRESULT " & vbCrLf
        SQL = SQL & "  WHERE BARCODE = '" & strBarNo & "'" & vbCrLf
        SQL = SQL & "    AND EQUIPNO = '" & gEquip & "' " & vbCrLf
        SQL = SQL & "    AND SAVESEQ = " & strSaveSeq
        SQL = SQL & " ORDER BY SEQNO"
        
        Set RS = cn.Execute(SQL, , 1)
    
        If Not RS.EOF = True And Not RS.BOF = True Then
            Do Until RS.EOF
                If Trim(RS("EQUIPCODE")) = "tIgE" Then
                    If InStr(Trim(RS("RESULT")), "��") > 0 Then
                        strResult = "�� 100"
                        'strClass0 = "Total IgE�� �����Դϴ�."
                    ElseIf InStr(Trim(RS("RESULT")), ">") > 0 Then
                        strResult = "> 100"
                        strClass0 = "Total IgE�� �����Ͽ����ϴ�."
                    End If
                    
'                    If Trim(RS("RESULT")) > 100 Then
'                        strResult = "> 100"
'                        strClass0 = "Total IgE�� �����Ͽ����ϴ�."
'                    Else
'                        strResult = "�� 100"
'                        strClass0 = "Total IgE�� �����Դϴ�."
'                    End If
                End If
            
                strClass = Trim(RS("REFVALUE"))
                Select Case strClass
                    Case "1": strClass1 = strClass1 & Trim(RS("SEQNO")) & ",": intClass = intClass + 1
                    Case "2": strClass2 = strClass2 & Trim(RS("SEQNO")) & ",": intClass = intClass + 1
                    Case "3": strClass3 = strClass3 & Trim(RS("SEQNO")) & ",": intClass = intClass + 1
                    Case "4": strClass4 = strClass4 & Trim(RS("SEQNO")) & ",": intClass = intClass + 1
                    Case "5": strClass5 = strClass5 & Trim(RS("SEQNO")) & ",": intClass = intClass + 1
                    Case "6": strClass6 = strClass6 & Trim(RS("SEQNO")) & ",": intClass = intClass + 1
                End Select
                            
    
                RS.MoveNext
            Loop
        End If
        
        strResult = ""
        strIntBase = "cmnt"
        If strClass1 <> "" Then
            strResult = strResult & Mid(strClass1, 1, Len(strClass1) - 1) & "���� Low " '& vbNewLine
            'intClass = intClass + 1
        End If
        If strClass2 <> "" Then
            strResult = strResult & Mid(strClass2, 1, Len(strClass2) - 1) & "���� Increased " '& vbNewLine
            'intClass = intClass + 1
        End If
        If strClass3 <> "" Then
            strResult = strResult & Mid(strClass3, 1, Len(strClass3) - 1) & "���� Significantly Increased " '& vbNewLine
            'intClass = intClass + 1
        End If
        If strClass4 <> "" Then
            strResult = strResult & Mid(strClass4, 1, Len(strClass4) - 1) & "���� High " '& vbNewLine
            'intClass = intClass + 1
        End If
        If strClass5 <> "" Then
            strResult = strResult & Mid(strClass5, 1, Len(strClass5) - 1) & "���� Very High " '& vbNewLine
            'intClass = intClass + 1
        End If
        If strClass6 <> "" Then
            strResult = strResult & Mid(strClass6, 1, Len(strClass6) - 1) & "���� Extremely High " '& vbNewLine
            'intClass = intClass + 1
        End If
        
        If strResult = "" Then
            If strClass0 <> "" Then '����
                strResult = strClass0 & vbNewLine & "Allergen�� ������ ��Ÿ���� �ʾҽ��ϴ�." & vbNewLine
            Else
                strResult = strClass0 & vbNewLine
            
            End If
        Else
            strResult = strClass0 & vbNewLine & "Allergen�� " & strResult & " ������ ��Ÿ�½��ϴ�." & vbNewLine
        End If
                        
'''        'If intClass >= 2 Then
'''            strResult = strResult & vbNewLine & "�������� �˷������� �缺������ ��Ÿ�½��ϴ�."
'''            strResult = strResult & vbNewLine & "�̴� �˷����� cross reaction�� ���� �������� �ǴܵǹǷ�"
'''            strResult = strResult & vbNewLine & "�ֵ� �ӻ�Ұ��� �����Ͻð� Skin test�� ���մϴ�."
'''        'Else
'''        '    strResult = strResult & vbNewLine & "�˷������� �缺������ ��Ÿ�½��ϴ�."
'''        '    strResult = strResult & vbNewLine & "�̴� �˷����� cross reaction�� ���� �������� �ǴܵǹǷ�"
'''        '    strResult = strResult & vbNewLine & "�ֵ� �ӻ�Ұ��� �����Ͻð� Skin test�� ���մϴ�."
'''        'End If
'''
'''        If intClass = 0 Then
'''            strResult = strResult & CMNT.N
'''        ElseIf intClass = 1 Then
'''            strResult = strResult & CMNT.P1
'''        ElseIf intClass > 1 Then
'''            strResult = strResult & CMNT.P2
'''        End If
            
        '-- IgE ����
        If blnIgE = True Then
            strResult = strResult & CMNT.N & vbNewLine & vbNewLine
        End If
        
        '-- ��������( => 10)
        If intClass >= 10 Then
            strResult = strResult & CMNT.P1
        End If
        
        
        '## �Ұֱ߳� ################################################################
        If strResult <> "" And Len(strIntBase) > 0 Then
            SQL = ""
            SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO " & vbCr
            SQL = SQL & "  FROM EQPMASTER" & vbCr
            SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' " & vbCr
            SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' " & vbCr
            SQL = SQL & "   AND GUBUN = '" & strGubun & "'" & vbCr
            SQL = SQL & "   AND EXAMCODE in (" & gOrderExam & ") "
            
            Res = GetDBSelectColumn(gLocal, SQL)
            
            '-- ���� ���� ���
            If Res > 0 Then
                lsExamCode = Trim(gReadBuf(0))
                lsExamName = Trim(gReadBuf(1))
                lsSeqNo = Trim(gReadBuf(2))
                
                '-- ��� List
                SetText vasRes, strIntBase, vasRes.MaxRows, colEQUIPCODE        '����ڵ�
                SetText vasRes, lsExamCode, vasRes.MaxRows, colEXAMCODE       '�˻��ڵ�
                SetText vasRes, lsExamName, vasRes.MaxRows, colEXAMNAME       '�˻��
                SetText vasRes, strResult, vasRes.MaxRows, colRESULT          '���
                SetText vasRes, strClass, vasRes.MaxRows, colCLASS            'CLASS
                SetText vasRes, lsSeqNo, vasRes.MaxRows, colSEQ               '����
                
                '-- ���� Update
                'SetLocalDB gRow, vasRes.MaxRows, "1", strResult
                      
                      SQL = "UPDATE  PATRESULT Set "
                SQL = SQL & "  RESULT = '" & strResult & "'" & vbCrLf
                SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf
                SQL = SQL & "   AND BARCODE = '" & strBarNo & "'" & vbCrLf
                SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "'" & vbCrLf
                SQL = SQL & "   AND DISKNO = '" & strGubun & "'" & vbCr
                Res = SendQuery(gLocal, SQL)
                
                If Res = -1 Then
                    SaveQuery SQL
                End If
                
            '-- ���� ���� ���
            Else
            
                      SQL = "Select examcode, examname, seqno " & vbCr
                SQL = SQL & "  From EQPMASTER" & vbCr
                SQL = SQL & " Where equipno = '" & gEquip & "' " & vbCr
                SQL = SQL & "   and equipcode = '" & strIntBase & "' " & vbCr
                SQL = SQL & "   AND GUBUN = '" & strGubun & "'"
                Res = GetDBSelectColumn(gLocal, SQL)
                                        
                If Res > 0 Then
                    lsExamCode = Trim(gReadBuf(0))
                    lsExamName = Trim(gReadBuf(1))
                    lsSeqNo = Trim(gReadBuf(2))
                    
                    '-- ��� List
                    SetText vasRes, strIntBase, vasRes.MaxRows, colEQUIPCODE       '����ڵ�
                    SetText vasRes, lsExamCode, vasRes.MaxRows, colEXAMCODE       '�˻��ڵ�
                    SetText vasRes, lsExamName, vasRes.MaxRows, colEXAMNAME       '�˻��
                    SetText vasRes, strResult, vasRes.MaxRows, colRESULT          '���
                    SetText vasRes, strClass, vasRes.MaxRows, colCLASS            'CLASS
                    SetText vasRes, lsSeqNo, vasRes.MaxRows, colSEQ               '����
                    '-- ���� ����
                    'SetLocalDB gRow, vasRes.MaxRows, "1", strResult
                    
                          SQL = "UPDATE  PATRESULT Set "
                    SQL = SQL & "  RESULT = '" & strResult & "'" & vbCrLf
                    SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf
                    SQL = SQL & "   AND BARCODE = '" & strBarNo & "'" & vbCrLf
                    SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "'" & vbCrLf
                    SQL = SQL & "   AND DISKNO = '" & strGubun & "'" & vbCr
                    Res = SendQuery(gLocal, SQL)
                    
                    If Res = -1 Then
                        SaveQuery SQL
                    End If
                    
                End If
            End If
        End If
        
        '## �Ұֱ߳� ################################################################
            
        RS.Close
        
        vasID.RowHeight(-1) = 12
    Next
    
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


Private Sub GetWorkList_PHILL(ByVal pFrDt As String, ByVal pToDt As String, Optional pBarNo As String)
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

    SQL = ""
    SQL = SQL & "SELECT DISTINCT P.request_date AS ��������, P.exam_no AS ������ȣ, P.company_code AS �Ƿ�ó, P.chart_no AS ��Ʈ��ȣ, p.personal_id, p.person_name AS �̸�, " & vbCr
    SQL = SQL & "       P.worker_code, P.patient_kind, P.person_sex AS ����, P.person_age AS ����, R.pro_code AS ó���ڵ� " & vbCr
    SQL = SQL & "  FROM trust P, trures R " & vbCr
    SQL = SQL & " WHERE P.request_date = '" & pFrDt & "'" & vbCr
    SQL = SQL & "   AND R.pro_code IN ('" & gAssayNM.INHALANT_CD & "','" & gAssayNM.FOOD_CD & "','" & gAssayNM.ATOPY_CD & "') " & vbCr
    SQL = SQL & "   AND R.exam_code <> 'X999' " & vbCr
    SQL = SQL & "   AND P.request_date = R.request_date " & vbCr
    SQL = SQL & "   AND P.exam_no = R.exam_no " & vbCr
    SQL = SQL & " ORDER BY P.request_date, P.exam_no "

    Call SetSQLData("��ũ��ȸ", SQL)

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
                    strChart = GetText(vasID, i, colPID)
                    If Trim(RS("��������")) = strDate And Mid(Trim(RS.Fields("��������")), 3, 6) & PedLeftStr(Trim(RS.Fields("������ȣ")), 5, "0") = strChart Then
                        blnSame = True
                    End If
                Next
                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    SetText vasID, "1", .MaxRows, colCheckBox
                    SetText vasID, Trim(RS.Fields("��������")) & "", .MaxRows, colHOSPDATE
                    
'                    MsgBox PedLeftStr(Trim(RS.Fields("������ȣ")), 5, "0")
'                    MsgBox Mid(Trim(RS.Fields("��������")), 3, 6)
                    
                    'SetText vasID, "20" & Mid(Trim(RS.Fields("��������")), 3, 6) & PedLeftStr(Trim(RS.Fields("������ȣ")), 5, "0"), .MaxRows, colBARCODE
                    SetText vasID, Trim(RS.Fields("��������")) & PedLeftStr(Trim(RS.Fields("������ȣ")), 5, "0"), .MaxRows, colBARCODE
                    SetText vasID, Trim(RS.Fields("��Ʈ��ȣ")) & "", .MaxRows, colCHARTNO
                    SetText vasID, Trim(RS.Fields("������ȣ")) & "", .MaxRows, colPID
                    SetText vasID, Trim(RS.Fields("�̸�")) & "", .MaxRows, colPNAME
                    SetText vasID, Trim(RS.Fields("�Ƿ�ó")) & "", .MaxRows, colDISKNO
                    'SetText vasID, Trim(RS.Fields("PatientBirthDay")) & "", .MaxRows, colPOSNO    '-- �������
                    SetText vasID, Trim(RS.Fields("����")) & "", .MaxRows, colPSEX
                    SetText vasID, Trim(RS.Fields("����")) & "", .MaxRows, colPAGE
                    
                    Select Case Trim(RS.Fields("ó���ڵ�")) & ""        'ó���ڵ� ??
                        Case gAssayNM.INHALANT_CD: SetText vasID, "INHALANT", .MaxRows, colINOUT
                        Case gAssayNM.FOOD_CD:     SetText vasID, "FOOD", .MaxRows, colINOUT
                        Case gAssayNM.ATOPY_CD:    SetText vasID, "ATOPY", .MaxRows, colINOUT
                        Case Else:
                                                   SetBackColor vasID, .MaxRows, .MaxRows, 1, colState, 202, 255, 112
                                                   SetText vasID, "ó�����", .MaxRows, colINOUT
                                                                            
                    End Select
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
    
    SQL = ""
    SQL = SQL & "SELECT P.PATID AS ��Ʈ��ȣ, P.PATNAME AS �̸�, P.SEX AS ����, O.ACPTDATE AS �������� " & vbCr
    SQL = SQL & ", O.ACPTSEQ, O.RSVACPTSTATE, O.RESULTSTATE, O.DEPTCODE, O.ORDERDATE, O.SLIPNO AS ������ȣ, O.IOFLAG, O.ORDERCODE, O.ORDERNAME " & vbCr
'    SQL = SQL & ", R.SPCMNO as ���ڵ��ȣ, R.RESULTFLAG, R.RESULTNO, R.RESULTITEMCODE as ITEM " & vbCr
    SQL = SQL & ", R.SPCMNO as ���ڵ��ȣ, R.RESULTFLAG, R.RESULTNO, R.RESULTITEMCODE as ITEM " & vbCr
    SQL = SQL & "  FROM registinfos O, resultofnum R, PATMST P " & vbCr
    SQL = SQL & " WHERE O.acptdate = R.acptdate " & vbCr
    SQL = SQL & "   AND O.acptdate between '" & pFrDt & "' and '" & pToDt & "'" & vbCr
    SQL = SQL & "   AND O.patid = R.patid " & vbCr
    SQL = SQL & "   AND O.acptseq = R.acptseq " & vbCr
    SQL = SQL & "   AND O.patid = P.patid " & vbCr
    SQL = SQL & "   AND O.CLAS = 4 " & vbCr '�ӻ󺴸�
    SQL = SQL & "   AND O.ORDERCODE IN ('" & gAssayNM.INHALANT_CD & "','" & gAssayNM.FOOD_CD & "') " & vbCr
    If chkSaveAll.Value = "0" Then
        SQL = SQL & "   AND R.RESULTFLAG = 0 " & vbCr
    End If
    'SQL = SQL & "   AND R.resultitemcode in (" & gAllExam & ")" & vbCr
    'SQL = SQL & "   AND R.resultitemcode IN ('" & gAssayNM.INHALANT_CD & "','" & gAssayNM.FOOD_CD & "') " & vbCr
    SQL = SQL & "  ORDER BY R.SPCMNO"
    
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
                    If Trim(RS("��������")) = strDate And Trim(RS("��Ʈ��ȣ")) = strChart Then
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
                    'If Trim(RS.Fields("���ڵ��ȣ")) & "" = "0" Then
                    '    SetText vasID, Trim(RS.Fields("��Ʈ��ȣ")) & "", .MaxRows, colBARCODE
                    'Else
                        SetText vasID, Trim(RS.Fields("���ڵ��ȣ")) & "", .MaxRows, colBARCODE
                    'End If
                    SetText vasID, Trim(RS.Fields("��Ʈ��ȣ")) & "", .MaxRows, colCHARTNO
                    SetText vasID, Trim(RS.Fields("�̸�")) & "", .MaxRows, colPNAME
                    SetText vasID, Trim(RS.Fields("������ȣ")) & "", .MaxRows, colPID
                    SetText vasID, Trim(RS.Fields("����")) & "", .MaxRows, colPSEX
                    
                    'Select Case Trim(RS.Fields("ITEM")) & ""        'ó���ڵ� ??
                    Select Case Trim(RS.Fields("ORDERCODE")) & ""        'ó���ڵ� ??
                        Case gAssayNM.INHALANT_CD: SetText vasID, "INHALANT", .MaxRows, colINOUT
                        Case gAssayNM.FOOD_CD:     SetText vasID, "FOOD", .MaxRows, colINOUT
                        Case gAssayNM.ATOPY_CD:    SetText vasID, "ATOPY", .MaxRows, colINOUT
                        Case Else:
                                                   SetBackColor vasID, .MaxRows, .MaxRows, 1, colState, 202, 255, 112
                                                   SetText vasID, "ó�����", .MaxRows, colINOUT
                                                                            
                    End Select
                    
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
            DoEvents
                        
            RS.MoveNext
        Loop
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
                
    txtCmnt = ""
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
        Case "PHILL":       Call GetWorkList_PHILL(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))
        Case "AMIS":        Call GetWorkList_AMIS(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))
    End Select
    
    
   ' Call GetWorkList_NTL(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))
    
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

'Private Sub Form_Resize()
'    On Error Resume Next
'
'    If frmInterface.ScaleHeight = 0 Then Exit Sub
'
'
'    If cmdSL.Caption = "��" Then
'        Frame1.Height = frmInterface.ScaleHeight - (Picture2.Top) - 1200
'        vasID.Height = Frame1.Height - 300
'
'        Frame1.Width = frmInterface.ScaleWidth - 200
'        vasID.Width = frmInterface.ScaleWidth - 7300
'
'
'        'Frame6.Left = vasID.Width + 300
'        vasRes.Height = vasID.Height - 550
'        vasRes.Left = vasID.Width + 300
'
'    Else
'        Frame1.Height = frmInterface.ScaleHeight - (Picture2.Top) - 1200
'        vasID.Height = Frame1.Height - 300
'
'        Frame1.Width = frmInterface.ScaleWidth - 200
'        vasID.Width = frmInterface.ScaleWidth - 300
'
'        'Frame6.Left = frmInterface.ScaleWidth - vasID.Width
'        'vasRes.Height = vasID.Height - 550
'        'vasRes.Left = Frame6.Left
'
'    End If
'
'    Picture2.Width = Frame1.Width
'
'    StatusBar1.Panels(3).Width = Frame1.Width - 8500
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
'            cn_Server_Flag = True
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
    
    '-- �Ұ� ��������
          SQL = " Select TITLE, CONTENT " & vbCr
    SQL = SQL & "  FROM CONFIG  " & vbCr
    SQL = SQL & " Where CATEGORY = 'COMMENT'"
          
    Res = GetDBSelectVas(gLocal, SQL, vasTemp)
        
    If Res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If
            
    vasTemp.MaxRows = vasTemp.DataRowCnt + 1

    '-- ������ ����� �����ϱ�
    For i = 1 To vasTemp.DataRowCnt
        Select Case Trim(GetText(vasTemp, i, 1))
            Case "N": CMNT.N = Trim(GetText(vasTemp, i, 2))
            Case "P1": CMNT.P1 = Trim(GetText(vasTemp, i, 2))
            Case "P2": CMNT.P2 = Trim(GetText(vasTemp, i, 2))
            Case "P3": CMNT.P3 = Trim(GetText(vasTemp, i, 2))
        End Select
    Next
    
    
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
    
    If gAssayNM.CMTVIEW = "1" Then
        vasRes.Height = 9795
        lblCmnt.Visible = True
        txtCmnt.Visible = True
    Else
        vasRes.Height = 11385
        lblCmnt.Visible = False
        txtCmnt.Visible = False
    End If
    
   ' Call cmdSL_Click
    
    '-- test
'    vasID.MaxRows = 10
    
End Sub

Private Sub SetExamCode()
    Dim i As Integer
    
    
    With vasID
        .MaxCols = colState + UBound(gArrEquip)
        
        For i = 0 To UBound(gArrEquip) - 1
            .Col = colState + (i + 1)
            .Row = -1
            .CellType = CellTypeEdit
            .TypeEditCharSet = TypeEditCharSetAlphanumeric
            .TypeEditCharCase = TypeEditCharCaseSetUpper
            .TypeHAlign = TypeHAlignCenter
            .TypeVAlign = TypeVAlignCenter
            'Call SetText(vasID, gArrEquip(i + 1, 2), 0, colState + (i + 1))
            Call SetText(vasID, gArrEquip(i + 1, 4), 0, colState + (i + 1))
            .ColWidth(colState + (i + 1)) = 6
        Next
    End With
    
End Sub


Function GetExamCode() As Integer
    Dim i, j As Long
    
    ClearSpread vasTemp
    GetExamCode = -1
    gAllExam = ""
    SQL = "Select equipcode, examcode, examname, resprec, seqno, gubun, examtype " & vbCrLf & _
          "  From EQPMASTER " & vbCrLf & _
          " Where equipno = '" & gEquip & "' " & vbCrLf & _
          " Order by gubun, seqno * 10 "
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

Private Sub MnCmmtConfig_Click()
    frmRemark.Show
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
    MsgBox "�����Ʈ ��� ����", vbOKOnly + vbInformation, Me.Caption
    'frmConfig.Show
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
                                '-- ��񿡼� �Ѿ�� �ð��� �쿬�� 11:59:59�ʳ� ���Ͽ� ����� �ð��� ���
                                '-- ��� ����� �������� ������ �� �����Ƿ� ��¥�� �ǽð� ������Ʈ �Ѵ�.
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
                                '-- ��񿡼� �Ѿ�� �ð��� �쿬�� 11:59:59�ʳ� ���Ͽ� ����� �ð��� ���
                                '-- ��� ����� �������� ������ �� �����Ƿ� ��¥�� �ǽð� ������Ʈ �Ѵ�.
                                strDate = DateCompare(Format(CDate(dtpToday.Value), "yyyymmdd"))
                                dtpToday.Value = Format(strDate, "####-##-##")

                                DoEvents
                                
                                Call EditRcvDataASTM
                                
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
        If Trim(GetText(vasID, i, colBARCODE)) = pBarNo And Trim(GetText(vasID, i, colINOUT)) = mResult.MnmNm And Trim(GetText(vasID, i, colSAVESEQ)) = mResult.RsltSeq Then
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
    Call SetText(vasID, mResult.MnmNm, intRow, colINOUT)
    
    Call vasActiveCell(vasID, intRow, colBARCODE)
    
    '-- ����������� �����
    Call ClearSpread(vasRes)
    
    '-- �˻��� ���� �������̺��� ������ ǥ��(for ��ũ����Ʈ)
    Call GetSampleInfoW_AMIS(intRow)
        
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
    Dim intIdx      As Integer
    Dim varRcvBuf   As Variant
    Dim intRow      As Integer
    Dim i As Integer
    Dim intCol As Integer
    Dim varHoriba As Variant
        
    For intCnt = 1 To UBound(strRecvData)
        strRcvBuf = strRecvData(intCnt)
        
        strType = Mid$(strRcvBuf, 1, 1)
        If IsNumeric(strType) Then
            strType = Mid$(strRcvBuf, 2, 1)
        End If
        
        Select Case strType
            Case "H"    '## Header
            Case "P"    '## Patient
            Case "O"    '## Order
                strBarNo = Trim(mGetP(mGetP(strRcvBuf, 3, "|"), 1, "^"))
                'strRackNo = mGetP(strTemp1, 1, "^")
                'strTubePos = mGetP(strTemp1, 2, "^")
                
                If strBarNo = "" Then Exit Sub
                
                With mResult
                    .BarNo = strBarNo
                    '.RackNo = strRackNo
                    '.TubePos = strTubePos
                    .RsltDate = Format(Now, "yyyymmddhhmmss")
                    .RsltSeq = getMaxTestNum(Format(dtpToday, "yyyymmdd"))
                End With
                                
                Call SetPatInfo(strBarNo)
                
                If gRow <= 0 Then
                    Exit Sub
                End If
                
                strState = "O"
                
                '-- ������ ���ȭ�� �ʱ�ȭ
                vasRes.MaxRows = 0
                
            Case "R"    '## Result
                '## ������ �˻��, ���, Abnormal Flag
                strIntBase = Trim(mGetP(mGetP(strRcvBuf, 3, "|"), 4, "^"))
                strResult = Trim(mGetP(strRcvBuf, 4, "|"))
'                If InStr(strTemp2, "^") > 0 Then
'                    '## ������� ����
'                    strResult = mGetP(strTemp2, 2, "^")
'                Else
'                    '## ������� ����
'                    strResult = strTemp2
'                End If
                
                If strResult <> "" And Len(strIntBase) <= 6 Then
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
                                SetText vasRes, gArrEquip(intCol - colState, 7), lsResRow, colSUBCODE               'subcode
                                Exit For
                            End If
                        Next
                        

                        '-- ��� List
                        SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      '����ڵ�
                        SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '�˻��ڵ�
                        SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '�˻��
                        SetText vasRes, lsEquipRes, lsResRow, colMachResult     '�����
                        SetText vasRes, strResult, lsResRow, colRESULT          '���
                        SetText vasRes, lsSeqNo, lsResRow, colSEQ               '����
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
                                    SetText vasRes, gArrEquip(intCol - colState, 7), lsResRow, colSUBCODE               'subcode
                                    Exit For
                                End If
                            Next
                            
                            '-- ��� List
                            SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      '����ڵ�
                            SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '�˻��ڵ�
                            SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '�˻��
                            SetText vasRes, lsEquipRes, lsResRow, colMachResult     '�����
                            SetText vasRes, strResult, lsResRow, colRESULT          '���
                            SetText vasRes, lsSeqNo, lsResRow, colSEQ               '����
                            SetText vasRes, strComm, lsResRow, colFLAG              'Flag
                            '-- ���� ����
                            SetLocalDB gRow, lsResRow, "1", lsEquipRes
                            
                            lsResult_Buff = ""
                            strState = "R"
                        End If
                    End If
                End If
                vasRes.RowHeight(-1) = 14
                
            Case "C"    '## Comment
                '## Abnormal ����϶� Comment ����
'                If strFlag <> "N" Then
'                    strTemp1 = mGetP(strRcvBuf, 4, "|")
'                    strComm = "[Flag]: " & mGetP(strTemp1, 1, "^") & ", " & mGetP(strTemp1, 2, "^")
'                End If
                
            Case "L"    '## Terminator
                '-- HCT%
                strIntBase = "HCT%"
                strResult = "%"
                
                If strResult <> "" And Len(strIntBase) <= 6 Then
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
                        
                        '-- �������� seq
                        For intCol = colState + 1 To vasID.MaxCols
                            If lsExamCode = gArrEquip(intCol - colState, 3) Then
                                SetText vasID, strResult, gRow, intCol
                                SetText vasRes, gArrEquip(intCol - colState, 7), lsResRow, colSUBCODE               'subcode
                                Exit For
                            End If
                        Next
                        

                        '-- ��� List
                        SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      '����ڵ�
                        SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '�˻��ڵ�
                        SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '�˻��
                        SetText vasRes, lsEquipRes, lsResRow, colMachResult     '�����
                        SetText vasRes, strResult, lsResRow, colRESULT          '���
                        SetText vasRes, lsSeqNo, lsResRow, colSEQ               '����
                        SetText vasRes, strComm, lsResRow, 7                    'Flag
                        '-- ���� ����
                        SetLocalDB gRow, lsResRow, "1", lsEquipRes
                                    
                        lsResult_Buff = ""
                        
                        strState = "R"
                    End If
                End If
                
                '-- LUC
                strIntBase = "LUC"
                strResult = "0.0"
                
                If strResult <> "" And Len(strIntBase) <= 6 Then
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
                        
                        '-- �������� seq
                        For intCol = colState + 1 To vasID.MaxCols
                            If lsExamCode = gArrEquip(intCol - colState, 3) Then
                                SetText vasID, strResult, gRow, intCol
                                SetText vasRes, gArrEquip(intCol - colState, 7), lsResRow, colSUBCODE               'subcode
                                Exit For
                            End If
                        Next
                        

                        '-- ��� List
                        SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      '����ڵ�
                        SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '�˻��ڵ�
                        SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '�˻��
                        SetText vasRes, lsEquipRes, lsResRow, colMachResult     '�����
                        SetText vasRes, strResult, lsResRow, colRESULT          '���
                        SetText vasRes, lsSeqNo, lsResRow, colSEQ               '����
                        SetText vasRes, strComm, lsResRow, 7                    'Flag
                        '-- ���� ����
                        SetLocalDB gRow, lsResRow, "1", lsEquipRes
                                    
                        lsResult_Buff = ""
                        
                        strState = "R"
                    End If
                End If
                
                '-- Diff
                strIntBase = "Diff"
                strResult = "100"
                
                If strResult <> "" And Len(strIntBase) <= 6 Then
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
                        
                        '-- �������� seq
                        For intCol = colState + 1 To vasID.MaxCols
                            If lsExamCode = gArrEquip(intCol - colState, 3) Then
                                SetText vasID, strResult, gRow, intCol
                                SetText vasRes, gArrEquip(intCol - colState, 7), lsResRow, colSUBCODE               'subcode
                                Exit For
                            End If
                        Next
                        

                        '-- ��� List
                        SetText vasRes, strIntBase, lsResRow, colEQUIPCODE      '����ڵ�
                        SetText vasRes, lsExamCode, lsResRow, colEXAMCODE       '�˻��ڵ�
                        SetText vasRes, lsExamName, lsResRow, colEXAMNAME       '�˻��
                        SetText vasRes, lsEquipRes, lsResRow, colMachResult     '�����
                        SetText vasRes, strResult, lsResRow, colRESULT          '���
                        SetText vasRes, lsSeqNo, lsResRow, colSEQ               '����
                        SetText vasRes, strComm, lsResRow, 7                    'Flag
                        '-- ���� ����
                        SetLocalDB gRow, lsResRow, "1", lsEquipRes
                                    
                        lsResult_Buff = ""
                        
                        strState = "R"
                    End If
                End If
                
                
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
Private Sub EditRcvDataAPEX()
    Dim strRcvBuf    As String   '������ Data
    Dim strType      As String   '������ Record Type
    Dim strBarNo     As String   '������ ���ڵ��ȣ
    Dim strSeq       As String   '������ Sequence
    Dim strRackNo    As String   '������ Rack Or Disk No
    Dim strTubePos   As String   '������ Tube Position
    Dim strIntBase   As String   '������ ������ �˻��
    Dim strResult    As String   '������ ���(����)
    Dim strClass     As String
    
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
    Dim intIdx      As Integer
    Dim varRcvBuf   As Variant
    Dim intRow      As Integer
    Dim i As Integer
    Dim intCol As Integer
    Dim varHoriba As Variant
    Dim strGubun  As String
    
    Dim strClass0 As String '-- Total IgE
    Dim strClass1 As String
    Dim strClass2 As String
    Dim strClass3 As String
    Dim strClass4 As String
    Dim strClass5 As String
    Dim strClass6 As String
    Dim intClass  As Integer
    Dim blnIgE    As Boolean
    
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
                
                With mResult
                    .BarNo = strBarNo
                    .RsltDate = Format(Now, "yyyymmddhhmmss")
                    .RsltSeq = getMaxTestNum(Format(dtpToday, "yyyymmdd"))
                    .MnmNm = strGubun
                End With
                
                strClass0 = ""
                strClass1 = ""
                strClass2 = ""
                strClass3 = ""
                strClass4 = ""
                strClass5 = ""
                strClass6 = ""
                intClass = 0
                blnIgE = False
                
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
                strFResult = strResult
                strClass = Trim(mGetP(mGetP(strRcvBuf, 2, "|"), 3, "^"))
                strFIntBase = strIntBase
                
                If strIntBase <> "PC" Then  'S2,S3,tIgE
                    If strClass = "-" Then
                        strResult = "-"
                        strFResult = "-"
                    End If
                End If
                
                '## �������� ����ֱ� ################################################################
                If strIntBase = "tIgE" Then
                    strFIntBase = ""
                    strFResult = ""
                    '## ������ �˻��, ���, Abnormal Flag
                    strIntBase = strGubun
                    strResult = "��������"
                    
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
                            SetText vasRes, strClass, lsResRow, colCLASS            'CLASS
                            SetText vasRes, lsSeqNo, lsResRow, colSEQ               '����
                            '-- ���� ����
                            SetLocalDB gRow, lsResRow, "1", lsEquipRes, strClass
                                        
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
                                SetText vasRes, strClass, lsResRow, colCLASS            'CLASS
                                SetText vasRes, lsSeqNo, lsResRow, colSEQ               '����
                                '-- ���� ����
                                SetLocalDB gRow, lsResRow, "1", lsEquipRes, strClass
                                
                                lsResult_Buff = ""
                                strState = "R"
                            End If
                        End If
                    End If
                    
                    strIntBase = Trim(mGetP(mGetP(strRcvBuf, 2, "|"), 1, "^"))
                    strResult = Trim(mGetP(mGetP(strRcvBuf, 2, "|"), 2, "^"))
                
                End If
                '## �������� ����ֱ� ################################################################
                
                If strIntBase = "tIgE" Then
                    If IsNumeric(strResult) Then
                        If strResult > 100 Then
                            strResult = "> 100"
                            strClass0 = "Total IgE�� �����Ͽ����ϴ�."
                            blnIgE = True
                        Else
                            strResult = "�� 100"
    '                        strClass0 = "Total IgE�� �����Դϴ�."
                            blnIgE = False
                        End If
                    Else
                        If InStr(strResult, "2000") > 0 Then   '>2000
                            strResult = "> 100"
                            strClass0 = "Total IgE�� �����Ͽ����ϴ�."
                            blnIgE = True
                        ElseIf InStr(strResult, "<0.15") > 0 Then '<0.15
                            strResult = "�� 100"
                            blnIgE = False
                        End If
                    End If
                Else
'''                    If InStr(strResult, "<0.15") > 0 Then
'''                        'strResult = "0.00"
'''                        strResult = "0.15"
'''                    End If
'''                    strResult = Replace(strResult, ">", "")
                End If
                
                
                strFResult = strResult
                
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
                        
                        Select Case strClass
                            Case "1": strClass1 = strClass1 & lsSeqNo & ",": intClass = intClass + 1
                            Case "2": strClass2 = strClass2 & lsSeqNo & ",": intClass = intClass + 1
                            Case "3": strClass3 = strClass3 & lsSeqNo & ",": intClass = intClass + 1
                            Case "4": strClass4 = strClass4 & lsSeqNo & ",": intClass = intClass + 1
                            Case "5": strClass5 = strClass5 & lsSeqNo & ",": intClass = intClass + 1
                            Case "6": strClass6 = strClass6 & lsSeqNo & ",": intClass = intClass + 1
                        End Select
                        
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
                        SetText vasRes, strClass, lsResRow, colCLASS            'CLASS
                        SetText vasRes, lsSeqNo, lsResRow, colSEQ               '����
                        '-- ���� ����
                        'SetLocalDB gRow, lsResRow, "1", lsEquipRes, strClass
                        SetLocalDB gRow, lsResRow, "1", strFResult, strClass
                                    
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
                            
                            Select Case strClass
                                Case "1": strClass1 = strClass1 & lsSeqNo & ","
                                Case "2": strClass2 = strClass2 & lsSeqNo & ","
                                Case "3": strClass3 = strClass3 & lsSeqNo & ","
                                Case "4": strClass4 = strClass4 & lsSeqNo & ","
                                Case "5": strClass5 = strClass5 & lsSeqNo & ","
                                Case "6": strClass6 = strClass6 & lsSeqNo & ","
                            End Select
                            
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
                            SetText vasRes, strClass, lsResRow, colCLASS            'CLASS
                            SetText vasRes, lsSeqNo, lsResRow, colSEQ               '����
                            '-- ���� ����
                            SetLocalDB gRow, lsResRow, "1", lsEquipRes, strClass
                            
                            lsResult_Buff = ""
                            strState = "R"
                        End If
                    End If
                End If
                
                vasRes.RowHeight(-1) = 10
                
            Case "L"    '## Terminator
                strResult = ""
                strIntBase = "cmnt"
                If strClass1 <> "" Then
                    strResult = strResult & Mid(strClass1, 1, Len(strClass1) - 1) & "���� Low" & vbNewLine
                '    intClass = intClass + 1
                End If
                If strClass2 <> "" Then
                    strResult = strResult & Mid(strClass2, 1, Len(strClass2) - 1) & "���� Increased" & vbNewLine
                    'intClass = intClass + 1
                End If
                If strClass3 <> "" Then
                    strResult = strResult & Mid(strClass3, 1, Len(strClass3) - 1) & "���� Significantly Increased" & vbNewLine
                    'intClass = intClass + 1
                End If
                If strClass4 <> "" Then
                    strResult = strResult & Mid(strClass4, 1, Len(strClass4) - 1) & "���� High" & vbNewLine
                    'intClass = intClass + 1
                End If
                If strClass5 <> "" Then
                    strResult = strResult & Mid(strClass5, 1, Len(strClass5) - 1) & "���� Very High" & vbNewLine
                    'intClass = intClass + 1
                End If
                If strClass6 <> "" Then
                    strResult = strResult & Mid(strClass6, 1, Len(strClass6) - 1) & "���� Extremely High" & vbNewLine
                    'intClass = intClass + 1
                End If
                
                If strResult = "" Then
                    strResult = strClass0 & vbNewLine
                Else
                    strResult = strClass0 & vbNewLine & "Allergen�� " & strResult & " ������ ��Ÿ�½��ϴ�." & vbNewLine
                End If
                             
                '-- IgE ����
                If blnIgE = True Then
                    strResult = strResult & CMNT.N & vbNewLine
                End If
                
                '-- ��������( => 10)
                If intClass >= 10 Then
                    strResult = strResult & CMNT.P1
                End If

'''                '-- IgE ����
'''                If blnIgE = False Then
'''                    '-- Class 2 �̻� ����
'''                    If intClass >= 1 Then
'''                        strResult = strResult & CMNT.P1
'''                    '-- Class 2 �̻� ����
'''                    Else
'''                        strResult = strResult & CMNT.N
'''                    End If
'''                '-- IgE ����
'''                Else
'''                    '-- Class 2 �̻� ����
'''                    If intClass >= 1 Then
'''                        strResult = strResult & CMNT.P3
'''                    '-- Class 2 �̻� ����
'''                    Else
'''                        strResult = strResult & CMNT.P2
'''                    End If
'''                End If
                
'''                If intClass >= 1 Then
''''                    strResult = strResult & vbNewLine & "�������� �˷������� �缺������ ��Ÿ�½��ϴ�."
''''                    strResult = strResult & vbNewLine & "�̴� �˷����� cross reaction�� ���� �������� �ǴܵǹǷ�"
''''                    strResult = strResult & vbNewLine & "�ֵ� �ӻ�Ұ��� �����Ͻð� Skin test�� ���մϴ�."
'''
'''                    strResult = strResult & CMNT.P1
'''                Else
''''                    strResult = strResult & vbNewLine & "�˷������� �缺������ ��Ÿ�½��ϴ�."
''''                    strResult = strResult & vbNewLine & "�̴� �˷����� cross reaction�� ���� �������� �ǴܵǹǷ�"
''''                    strResult = strResult & vbNewLine & "�ֵ� �ӻ�Ұ��� �����Ͻð� Skin test�� ���մϴ�."
'''
'''                    strResult = strResult & CMNT.N
'''                End If
                
                '## �Ұֱ߳� ################################################################
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
'                        strResult = SetResult(strResult, strIntBase)
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
'                        SetText vasRes, lsEquipRes, lsResRow, colMachResult     '�����
                        SetText vasRes, strResult, lsResRow, colRESULT          '���
                        SetText vasRes, strClass, lsResRow, colCLASS            'CLASS
                        SetText vasRes, lsSeqNo, lsResRow, colSEQ               '����
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
'                            strResult = SetResult(strResult, strIntBase)
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
'                            SetText vasRes, lsEquipRes, lsResRow, colMachResult     '�����
                            SetText vasRes, strResult, lsResRow, colRESULT          '���
                            SetText vasRes, strClass, lsResRow, colCLASS            'CLASS
                            SetText vasRes, lsSeqNo, lsResRow, colSEQ               '����
                            '-- ���� ����
                            SetLocalDB gRow, lsResRow, "1", lsEquipRes
                            
                            lsResult_Buff = ""
                            strState = "R"
                        End If
                    End If
                End If
                    
                '## �Ұֱ߳� ################################################################
                
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
Rst:
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
Function SetLocalDB(ByVal asRow1 As Long, ByVal asRow2 As Long, asSend As String, Optional asEquipResult As String = "", Optional asEquipClass As String = "")
    Dim sCnt As String
    Dim sExamDate As String
    Dim strSaveSeq As String
    Dim RS          As ADODB.Recordset
    Dim blnUpdate   As Boolean
    Dim intUpCnt    As Integer
    Dim strUpData   As String
    Dim strGubuns   As String
    Dim varGubuns   As Variant
    Dim intCnt      As Integer
    Dim intCnt1     As Integer
    Dim strChannel  As String
    Dim strGubun    As String
    
    Dim strOrgVal   As String
    Dim strOldVal   As String
    Dim strNewVal   As String
    Dim strClass    As String
    Dim strFlag     As String
    
    blnUpdate = False
    'sExamDate = Format(dtpToday, "yyyymmddhhmmss")
    sExamDate = Mid(Trim(GetText(vasID, asRow1, colEXAMDATE)), 1, 8)
    
    strChannel = Trim(GetText(vasRes, asRow2, colEQUIPCODE))
    strGubun = Trim(GetText(vasID, asRow1, colINOUT))
    
    SQL = ""
'          " WHERE EXAMDATE = '" & Mid(sExamDate, 1, 8) & "' " & vbCrLf & _

    SQL = "DELETE FROM PATRESULT " & vbCrLf & _
          " WHERE EXAMDATE = '" & sExamDate & "'" & vbCrLf & _
          "   AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
          "   AND SAVESEQ = " & Trim(GetText(vasID, asRow1, colSAVESEQ)) & vbCrLf & _
          "   AND BARCODE = '" & Trim(GetText(vasID, asRow1, colBARCODE)) & "' " & vbCrLf & _
          "   AND EQUIPCODE = '" & Trim(GetText(vasRes, asRow2, colEQUIPCODE)) & "'" & vbCrLf & _
          "   AND EXAMCODE = '" & Trim(GetText(vasRes, asRow2, colEXAMCODE)) & "'" & vbCrLf
    SQL = SQL & "   AND DISKNO = '" & Trim(GetText(vasID, asRow1, colINOUT)) & "'"
    Res = SendQuery(gLocal, SQL)
    
    If Res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
    'strSaveSeq = getMaxTestNum(Mid(sExamDate, 1, 8))
    
    '-- �����ڵ��� ���
    For intCnt = 1 To UBound(gArrEquip)
        If strChannel = gArrEquip(intCnt, 2) And strGubun = gArrEquip(intCnt, 7) And gArrEquip(intCnt, 8) = "����" Then
            '-- ����
            '-- ���� ���ڵ忡 (������ Ʋ���� ��) ���� ä���� ����� �ִ��� Ȯ���Ѵ�.
            '-- Select
                  SQL = " SELECT RESULT, DISKNO, REFVALUE, REFFLAG " & vbCrLf
            SQL = SQL & "   FROM PATRESULT " & vbCrLf
            SQL = SQL & "  WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf
            'SQL = SQL & "    AND DISKNO =  '" & strGubun & "' "
            SQL = SQL & "    AND EXAMDATE = '" & sExamDate & "'" & vbCrLf
            SQL = SQL & "    AND BARCODE = '" & Trim(GetText(vasID, asRow1, colBARCODE)) & "'" & vbCrLf
            SQL = SQL & "    AND EQUIPCODE = '" & Trim(GetText(vasRes, asRow2, colEQUIPCODE)) & "'" & vbCrLf
            SQL = SQL & "    AND SENDFLAG <> '2'"
            Set RS = cn.Execute(SQL, , 1)
            strUpData = ""
            strGubuns = ""
            intUpCnt = 0
            
            If Not RS.EOF = True And Not RS.BOF = True Then
                Do Until RS.EOF
                
'''                    If strChannel = "cmnt" Then
'''                        '-- �Ұ� ������Ʈ
'''                        strOldVal = RS.Fields("RESULT")
'''                        If strOldVal <> asEquipResult Then
'''                            blnUpdate = True
'''                            strUpData = asEquipResult
'''                            strGubuns = RS.Fields("DISKNO")
'''                        Else
'''                            blnUpdate = False
'''                            asEquipResult = RS.Fields("RESULT")
'''                        End If
'''                    Else
                        '-- ��� ������Ʈ
                        strOrgVal = RS.Fields("RESULT")
                        strOldVal = RS.Fields("RESULT")
                        strOldVal = Replace(strOldVal, "<", "")
                        strOldVal = Replace(strOldVal, ">", "")
                        strOldVal = Replace(strOldVal, "��", "")
                        strOldVal = Trim(strOldVal)
                        
                        If Val(strOldVal) < Val(asEquipResult) Then
                            blnUpdate = True
                            strGubuns = RS.Fields("DISKNO")
                            
                            '-- 2017.02.22
                            'strUpData = strOrgVal 'asEquipResult
                            strUpData = asEquipResult
                            
                            strClass = asEquipClass
                            strFlag = RS.Fields("REFFLAG")
                        Else
                            If Val(strOldVal) > Val(asEquipResult) Then
                                blnUpdate = True
                                strGubuns = RS.Fields("DISKNO")
                                strUpData = strOrgVal 'strOldVal
                                strClass = RS.Fields("REFVALUE")
                                strFlag = RS.Fields("REFFLAG")
                            Else
                                blnUpdate = False
                                asEquipResult = RS.Fields("RESULT")
                                asEquipClass = RS.Fields("REFVALUE")
                            End If
                        End If



'''                    End If
                    RS.MoveNext
                Loop
            Else
                Exit For
            End If
            
            '-- ���� ���� �׸� ���� ����� ������ ���� ����� �����Ѵ�.
            '-- Update
            If blnUpdate = True Then
                      SQL = "UPDATE  PATRESULT Set "
                SQL = SQL & "  RESULT = '" & strUpData & "'" & vbCrLf
                SQL = SQL & " ,REFVALUE = '" & strClass & "'" & vbCrLf
                SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf
                SQL = SQL & "   AND EXAMDATE = '" & sExamDate & "'" & vbCrLf
                SQL = SQL & "   AND BARCODE = '" & Trim(GetText(vasID, asRow1, colBARCODE)) & "'" & vbCrLf
                SQL = SQL & "   AND EQUIPCODE = '" & Trim(GetText(vasRes, asRow2, colEQUIPCODE)) & "'" & vbCrLf
                'SQL = SQL & "   AND DISKNO = '" & strGubuns & "'"
                Res = SendQuery(gLocal, SQL)
                
                If Res = -1 Then
                    SaveQuery SQL
                    Exit Function
                End If
            
                '-- �Ұ� ����
'                      SQL = "UPDATE  PATRESULT Set "
'                SQL = SQL & " RESULT = '" & strUpData & "'" & vbCrLf
'                SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf
'                SQL = SQL & "   AND EXAMDATE = '" & sExamDate & "'" & vbCrLf
'                SQL = SQL & "   AND BARCODE = '" & Trim(GetText(vasID, asRow1, colBARCODE)) & "'" & vbCrLf
'                SQL = SQL & "   AND EQUIPCODE = '" & Trim(GetText(vasRes, asRow2, colEQUIPCODE)) & "'" & vbCrLf
'                SQL = SQL & "   AND DISKNO = '" & strGubuns & "'"
'                Res = SendQuery(gLocal, SQL)
'
'                If Res = -1 Then
'                    SaveQuery SQL
'                    Exit Function
'                End If
            End If
            '-- Insert ��� ����
            Exit For
        End If
    Next
    
    
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
    SQL = SQL & "','" & Trim(GetText(vasRes, asRow2, colSEQ))
    SQL = SQL & "',''"
    'SQL = SQL & ",'" & Trim(GetText(vasID, asRow1, colDISKNO))
    SQL = SQL & ",'" & Trim(GetText(vasID, asRow1, colINOUT))
    SQL = SQL & "','" & Trim(GetText(vasID, asRow1, colDISKNO))
    SQL = SQL & "','" & Trim(GetText(vasRes, asRow2, colMachResult))
    
    If strUpData <> "" Then
        SQL = SQL & "','" & strUpData
        SQL = SQL & "','" & strFlag
        SQL = SQL & "','" & strClass
    Else
        SQL = SQL & "','" & Trim(GetText(vasRes, asRow2, colRESULT))
        SQL = SQL & "','" & Trim(GetText(vasRes, asRow2, colFLAG))
        SQL = SQL & "','" & asEquipClass
    End If
    
'    SQL = SQL & "','" & Trim(GetText(vasRes, asRow2, colFLAG))
    
    'SQL = SQL & "','" & Trim(GetText(vasRes, asRow2, colCLASS))
    
    SQL = SQL & "','" & Trim(GetText(vasID, asRow1, colCHARTNO))
    SQL = SQL & "','" & Trim(GetText(vasID, asRow1, colPID))
    SQL = SQL & "','" & Trim(GetText(vasID, asRow1, colPNAME))
    SQL = SQL & "','" & Trim(GetText(vasID, asRow1, colPSEX))
    SQL = SQL & "','" & Trim(GetText(vasID, asRow1, colPAGE))
    SQL = SQL & "',''"
    SQL = SQL & ",''"
    SQL = SQL & ",''"
    SQL = SQL & ",'1'"
    SQL = SQL & ",''"
    SQL = SQL & ",'" & gIFUser
    SQL = SQL & "','')"
    
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

    txtStartNum.Text = "0"
    txtStopNum.Text = "0"

End Sub



Private Sub picLogin_Click()

    Dim sMsg As String
    sMsg = "�˻��ڸ� �Է����ּ���."
    lblUser.Caption = InputBox(sMsg, "�˻��� �Է�")

End Sub

Private Sub txtBarcode_KeyPress(KeyAscii As Integer)
    Dim intRow As Integer
    
    If KeyAscii = 13 And Len(txtBarcode) > 0 Then
        
        vasID.MaxRows = vasID.MaxRows + 1
        intRow = vasID.MaxRows
        
        Call SetText(vasID, txtBarcode, intRow, colBARCODE)
        
        Call GetSampleInfoW_AMIS(intRow)
        
        SelectFocus txtBarcode
    End If

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
    txtCmnt = ""
    
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
          SQL = "SELECT EQUIPCODE, EXAMCODE, EXAMNAME, EQUIPRESULT, RESULT, SEQNO, REFFLAG, EXAMSUBCODE, REFVALUE " & vbCrLf
    SQL = SQL & "  FROM PATRESULT " & vbCrLf
    SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "'" & vbCrLf
    SQL = SQL & "   AND SAVESEQ = " & lblSaveSeq.Caption & vbCrLf
    SQL = SQL & "   AND BARCODE = '" & lsID & "' " & vbCrLf
    'SQL = SQL & "   AND EXAMDATE = '" & Mid(Trim(GetText(vasID, Row, colOrdDate)), 1, 8) & "' " & vbCrLf
    SQL = SQL & "   AND DISKNO = '" & Trim(GetText(vasID, Row, colINOUT)) & "' " & vbCrLf
    SQL = SQL & " GROUP BY SEQNO, EQUIPCODE, EXAMCODE, EXAMNAME, EQUIPRESULT, RESULT, SEQNO, REFFLAG, EXAMSUBCODE, REFVALUE "
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
                SetText vasRes, Trim(RS.Fields("REFVALUE")) & "", .MaxRows, colCLASS
                SetText vasRes, Trim(RS.Fields("SEQNO")) & "", .MaxRows, colSEQ
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
            If Trim(RS.Fields("EQUIPCODE")) & "" = "cmnt" Then
                txtCmnt.Text = Trim(RS.Fields("RESULT")) & ""
            End If
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
            Call GetSampleInfoW_AMIS(iRow)
            
            lsID = Trim(GetText(vasID, iRow, colBARCODE))
            
            
            '-- ���ڵ� ��ȣ�� ������ Ʋ���ٸ� ������Ʈ
            'If lsID <> lblChangeBar.Caption Then
            If lsID <> lblBarcode(0).Caption Then
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

