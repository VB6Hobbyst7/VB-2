VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmProgCfg 
   Caption         =   "�������̽� ���α׷� ȯ�漳��"
   ClientHeight    =   8595
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11880
   ClipControls    =   0   'False
   Icon            =   "PROGCFG.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'ȭ�� ���
   Begin Threed.SSFrame SSFrame3 
      Height          =   7275
      Left            =   9600
      TabIndex        =   16
      Top             =   0
      Width           =   2235
      _Version        =   65536
      _ExtentX        =   3942
      _ExtentY        =   12832
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin FPSpread.vaSpread spdPosSet 
         Height          =   5385
         Left            =   330
         TabIndex        =   24
         Top             =   1740
         Width           =   1650
         _Version        =   196608
         _ExtentX        =   2910
         _ExtentY        =   9499
         _StockProps     =   64
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
         MaxCols         =   2
         MaxRows         =   20
         NoBeep          =   -1  'True
         ScrollBars      =   2
         SpreadDesigner  =   "PROGCFG.frx":08CA
         UserResize      =   0
         VisibleCols     =   500
         VisibleRows     =   500
      End
      Begin VB.TextBox txtMaxRack 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   8  '����
         Left            =   1500
         MaxLength       =   2
         TabIndex        =   23
         Text            =   "20"
         Top             =   1380
         Width           =   345
      End
      Begin VB.TextBox txtPosDig 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   8  '����
         Left            =   1500
         MaxLength       =   1
         TabIndex        =   22
         Text            =   "2"
         Top             =   1050
         Width           =   225
      End
      Begin VB.TextBox txtRackDig 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   8  '����
         Left            =   1500
         MaxLength       =   1
         TabIndex        =   21
         Text            =   "3"
         Top             =   720
         Width           =   225
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   285
         Left            =   390
         TabIndex        =   18
         Top             =   720
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "Rack�ڸ���"
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   315
         Left            =   270
         TabIndex        =   17
         Top             =   330
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "Rack Setting"
         ForeColor       =   12648447
         BackColor       =   16448
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   285
         Left            =   390
         TabIndex        =   19
         Top             =   1050
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "Pos�ڸ���"
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   285
         Left            =   390
         TabIndex        =   20
         Top             =   1380
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "MaxRack"
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   1275
      Left            =   30
      TabIndex        =   2
      Top             =   0
      Width           =   9555
      _Version        =   65536
      _ExtentX        =   16854
      _ExtentY        =   2249
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.Frame Frame5 
         Caption         =   "Parity"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Left            =   8190
         TabIndex        =   12
         Top             =   150
         Width           =   1215
         Begin VB.OptionButton OptParity 
            Caption         =   "None"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   15
            Top             =   270
            Width           =   1000
         End
         Begin VB.OptionButton OptParity 
            Caption         =   "Odd"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   150
            TabIndex        =   14
            Top             =   750
            Width           =   1000
         End
         Begin VB.OptionButton OptParity 
            Caption         =   "Even"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   150
            TabIndex        =   13
            Top             =   510
            Width           =   1000
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Baud Rate"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Left            =   5520
         TabIndex        =   10
         Top             =   150
         Width           =   1965
         Begin VB.ComboBox cboBaudRate 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "PROGCFG.frx":0BA6
            Left            =   90
            List            =   "PROGCFG.frx":0BC8
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   11
            Top             =   390
            Width           =   1785
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Stop Bit"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Left            =   3870
         TabIndex        =   7
         Top             =   150
         Width           =   1005
         Begin VB.OptionButton OptStop 
            Caption         =   "2bit"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   150
            TabIndex        =   9
            Top             =   600
            Width           =   795
         End
         Begin VB.OptionButton OptStop 
            Caption         =   "1bit"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   8
            Top             =   330
            Width           =   765
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Data Bit"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Left            =   2220
         TabIndex        =   4
         Top             =   150
         Width           =   1035
         Begin VB.OptionButton OptData 
            Caption         =   "8bit"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   150
            TabIndex        =   6
            Top             =   600
            Width           =   705
         End
         Begin VB.OptionButton OptData 
            Caption         =   "7bit"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   5
            Top             =   330
            Width           =   705
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "��Ʈ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Left            =   510
         TabIndex        =   3
         Top             =   150
         Width           =   1155
         Begin VB.ComboBox cboPort 
            Height          =   276
            ItemData        =   "PROGCFG.frx":0C0B
            Left            =   72
            List            =   "PROGCFG.frx":0C4B
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   117
            Top             =   432
            Width           =   1020
         End
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   5010
         Picture         =   "PROGCFG.frx":0CD2
         Top             =   150
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   60
         Picture         =   "PROGCFG.frx":1114
         Top             =   180
         Width           =   480
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   3360
         Picture         =   "PROGCFG.frx":1556
         Top             =   150
         Width           =   480
      End
      Begin VB.Image Image5 
         Height          =   480
         Left            =   1740
         Picture         =   "PROGCFG.frx":1998
         Top             =   150
         Width           =   480
      End
      Begin VB.Image Image6 
         Height          =   480
         Left            =   7620
         Picture         =   "PROGCFG.frx":1DDA
         Top             =   150
         Width           =   480
      End
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   990
      Left            =   10710
      TabIndex        =   1
      Top             =   7290
      Width           =   1035
      _Version        =   65536
      _ExtentX        =   1826
      _ExtentY        =   1746
      _StockProps     =   78
      Caption         =   "�� ��"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      RoundedCorners  =   0   'False
      Picture         =   "PROGCFG.frx":221C
   End
   Begin Threed.SSCommand cmdSave 
      Height          =   990
      Left            =   9690
      TabIndex        =   0
      Top             =   7290
      Width           =   1035
      _Version        =   65536
      _ExtentX        =   1826
      _ExtentY        =   1746
      _StockProps     =   78
      Caption         =   "�� ��"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      RoundedCorners  =   0   'False
      Picture         =   "PROGCFG.frx":307A
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   3315
      Left            =   30
      TabIndex        =   25
      Top             =   1200
      Width           =   9555
      _Version        =   65536
      _ExtentX        =   16854
      _ExtentY        =   5847
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtOCOM 
         Height          =   300
         Left            =   300
         TabIndex        =   109
         Top             =   1350
         Width           =   1485
      End
      Begin Threed.SSFrame SSFrame5 
         Height          =   3105
         Left            =   2130
         TabIndex        =   35
         Top             =   120
         Width           =   7245
         _Version        =   65536
         _ExtentX        =   12779
         _ExtentY        =   5477
         _StockProps     =   14
         Caption         =   "�ʵ� ����"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.ComboBox cboOFsize 
            Height          =   300
            Index           =   8
            Left            =   5880
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   65
            Top             =   2430
            Width           =   615
         End
         Begin VB.ComboBox cboOFsize 
            Height          =   300
            Index           =   7
            Left            =   5880
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   64
            Top             =   2160
            Width           =   615
         End
         Begin VB.ComboBox cboOFsize 
            Height          =   300
            Index           =   6
            Left            =   5880
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   63
            Top             =   1890
            Width           =   615
         End
         Begin VB.ComboBox cboOFsize 
            Height          =   300
            Index           =   5
            Left            =   5880
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   62
            Top             =   1620
            Width           =   615
         End
         Begin VB.ComboBox cboOFsize 
            Height          =   300
            Index           =   4
            Left            =   5880
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   61
            Top             =   1350
            Width           =   615
         End
         Begin VB.ComboBox cboOFsize 
            Height          =   300
            Index           =   3
            Left            =   5880
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   60
            Top             =   1080
            Width           =   615
         End
         Begin VB.ComboBox cboOFsize 
            Height          =   300
            Index           =   9
            Left            =   5880
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   59
            Top             =   2700
            Width           =   615
         End
         Begin VB.ComboBox cboOFsize 
            Height          =   300
            Index           =   2
            Left            =   5880
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   58
            Top             =   810
            Width           =   615
         End
         Begin VB.ComboBox cboOFsize 
            Height          =   300
            Index           =   1
            Left            =   5880
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   57
            Top             =   540
            Width           =   615
         End
         Begin VB.ComboBox cboOFsize 
            Height          =   300
            Index           =   0
            Left            =   5880
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   56
            Top             =   270
            Width           =   615
         End
         Begin VB.ComboBox cboOFord 
            Height          =   300
            Index           =   8
            Left            =   4260
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   55
            Top             =   2430
            Width           =   615
         End
         Begin VB.ComboBox cboOFord 
            Height          =   300
            Index           =   7
            Left            =   4260
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   54
            Top             =   2160
            Width           =   615
         End
         Begin VB.ComboBox cboOFord 
            Height          =   300
            Index           =   6
            Left            =   4260
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   53
            Top             =   1890
            Width           =   615
         End
         Begin VB.ComboBox cboOFord 
            Height          =   300
            Index           =   5
            Left            =   4260
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   52
            Top             =   1620
            Width           =   615
         End
         Begin VB.ComboBox cboOFord 
            Height          =   300
            Index           =   4
            Left            =   4260
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   51
            Top             =   1350
            Width           =   615
         End
         Begin VB.ComboBox cboOFord 
            Height          =   300
            Index           =   3
            Left            =   4260
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   50
            Top             =   1080
            Width           =   615
         End
         Begin VB.ComboBox cboOFord 
            Height          =   300
            Index           =   9
            Left            =   4260
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   49
            Top             =   2700
            Width           =   615
         End
         Begin VB.ComboBox cboOFord 
            Height          =   300
            Index           =   2
            Left            =   4260
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   48
            Top             =   810
            Width           =   615
         End
         Begin VB.ComboBox cboOFord 
            Height          =   300
            Index           =   1
            Left            =   4260
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   47
            Top             =   540
            Width           =   615
         End
         Begin VB.ComboBox cboOFord 
            Height          =   300
            Index           =   0
            Left            =   4260
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   46
            Top             =   270
            Width           =   615
         End
         Begin VB.CheckBox chkOFUse 
            Alignment       =   1  '������ ����
            Caption         =   "��Ÿ"
            Height          =   285
            Index           =   8
            Left            =   270
            TabIndex        =   45
            Top             =   2430
            Width           =   1215
         End
         Begin VB.CheckBox chkOFUse 
            Alignment       =   1  '������ ����
            Caption         =   "��ü"
            Height          =   285
            Index           =   7
            Left            =   270
            TabIndex        =   44
            Top             =   2160
            Width           =   1215
         End
         Begin VB.CheckBox chkOFUse 
            Alignment       =   1  '������ ����
            Caption         =   "�ܷ�/�Կ�"
            Height          =   285
            Index           =   6
            Left            =   270
            TabIndex        =   43
            Top             =   1890
            Width           =   1215
         End
         Begin VB.CheckBox chkOFUse 
            Alignment       =   1  '������ ����
            Caption         =   "����"
            Height          =   285
            Index           =   5
            Left            =   270
            TabIndex        =   42
            Top             =   1620
            Width           =   1215
         End
         Begin VB.CheckBox chkOFUse 
            Alignment       =   1  '������ ����
            Caption         =   "�̸�"
            Height          =   285
            Index           =   4
            Left            =   270
            TabIndex        =   41
            Top             =   1350
            Width           =   1215
         End
         Begin VB.CheckBox chkOFUse 
            Alignment       =   1  '������ ����
            Caption         =   "��Ϲ�ȣ"
            Height          =   285
            Index           =   3
            Left            =   270
            TabIndex        =   40
            Top             =   1080
            Width           =   1215
         End
         Begin VB.CheckBox chkOFUse 
            Alignment       =   1  '������ ����
            Caption         =   "�˻��ڵ�"
            Height          =   285
            Index           =   9
            Left            =   270
            TabIndex        =   39
            Top             =   2700
            Width           =   1215
         End
         Begin VB.CheckBox chkOFUse 
            Alignment       =   1  '������ ����
            Caption         =   "���ڵ��ȣ"
            Height          =   285
            Index           =   2
            Left            =   270
            TabIndex        =   38
            Top             =   810
            Width           =   1215
         End
         Begin VB.CheckBox chkOFUse 
            Alignment       =   1  '������ ����
            Caption         =   "�Ϲ�/�Ű�"
            Height          =   285
            Index           =   1
            Left            =   270
            TabIndex        =   37
            Top             =   540
            Width           =   1215
         End
         Begin VB.CheckBox chkOFUse 
            Alignment       =   1  '������ ����
            Caption         =   "��������"
            Height          =   285
            Index           =   0
            Left            =   270
            TabIndex        =   36
            Top             =   270
            Width           =   1215
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  '����
            Caption         =   "�ʵ� ũ��"
            Height          =   180
            Left            =   5070
            TabIndex        =   68
            Top             =   330
            Width           =   780
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  '����
            Caption         =   "�ʵ� ����"
            Height          =   180
            Left            =   3450
            TabIndex        =   67
            Top             =   330
            Width           =   780
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '����
            Caption         =   "����� ���ϴ� �ʵ����        üũ�ϰ�                         üũ�� �ʵ��� �ʵ� ������ �ʵ� ũ�⸦ ����"
            Height          =   1455
            Left            =   1740
            TabIndex        =   66
            Top             =   690
            Width           =   1155
         End
      End
      Begin VB.TextBox txtOStoragePath 
         Height          =   300
         Left            =   300
         TabIndex        =   34
         Top             =   2880
         Width           =   1485
      End
      Begin VB.ComboBox cboOStorage 
         Height          =   300
         ItemData        =   "PROGCFG.frx":411C
         Left            =   300
         List            =   "PROGCFG.frx":4129
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   32
         Top             =   2010
         Width           =   1515
      End
      Begin VB.CheckBox chkOUse 
         Alignment       =   1  '������ ����
         Caption         =   "Order ���"
         Height          =   210
         Left            =   270
         TabIndex        =   29
         Top             =   630
         Width           =   1185
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   315
         Left            =   120
         TabIndex        =   27
         Top             =   150
         Width           =   1725
         _Version        =   65536
         _ExtentX        =   3043
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "   Order Setting"
         ForeColor       =   12640511
         BackColor       =   4210816
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         Alignment       =   1
      End
      Begin VB.Label Label9 
         BackStyle       =   0  '����
         Caption         =   "Order ���� �������� ������Ʈ"
         Height          =   375
         Left            =   300
         TabIndex        =   108
         Top             =   960
         Width           =   1665
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '����
         Caption         =   "Order ���� ����� ���"
         Height          =   360
         Left            =   300
         TabIndex        =   33
         Top             =   2490
         Width           =   1485
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "Order ���� �����"
         Height          =   180
         Left            =   300
         TabIndex        =   31
         Top             =   1800
         Width           =   1485
      End
   End
   Begin Threed.SSFrame SSFrame4 
      Height          =   3885
      Left            =   30
      TabIndex        =   26
      Top             =   4440
      Width           =   9555
      _Version        =   65536
      _ExtentX        =   16854
      _ExtentY        =   6853
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtRCOM 
         Height          =   300
         Left            =   300
         TabIndex        =   115
         Top             =   1320
         Width           =   1485
      End
      Begin VB.ComboBox cboRStorage 
         Height          =   300
         ItemData        =   "PROGCFG.frx":413F
         Left            =   300
         List            =   "PROGCFG.frx":414F
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   111
         Top             =   1980
         Width           =   1515
      End
      Begin VB.TextBox txtRStoragePath 
         Height          =   300
         Left            =   300
         TabIndex        =   110
         Top             =   2850
         Width           =   1485
      End
      Begin VB.CheckBox chkRUse 
         Alignment       =   1  '������ ����
         Caption         =   "Result ���"
         Height          =   210
         Left            =   270
         TabIndex        =   30
         Top             =   600
         Width           =   1245
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   315
         Left            =   120
         TabIndex        =   28
         Top             =   150
         Width           =   1725
         _Version        =   65536
         _ExtentX        =   3043
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "   Result Setting"
         ForeColor       =   12648384
         BackColor       =   16384
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         Alignment       =   1
      End
      Begin Threed.SSFrame SSFrame6 
         Height          =   3675
         Left            =   2130
         TabIndex        =   69
         Top             =   120
         Width           =   7245
         _Version        =   65536
         _ExtentX        =   12779
         _ExtentY        =   6482
         _StockProps     =   14
         Caption         =   "�ʵ� ����"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.ComboBox cboRFsize 
            Height          =   300
            Index           =   11
            Left            =   5880
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   107
            Top             =   3240
            Width           =   615
         End
         Begin VB.ComboBox cboRFord 
            Height          =   300
            Index           =   11
            Left            =   4260
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   106
            Top             =   3240
            Width           =   615
         End
         Begin VB.CheckBox chkRFUse 
            Alignment       =   1  '������ ����
            Caption         =   "��� 2"
            Height          =   285
            Index           =   11
            Left            =   270
            TabIndex        =   105
            Top             =   3240
            Width           =   1215
         End
         Begin VB.CheckBox chkRFUse 
            Alignment       =   1  '������ ����
            Caption         =   "��������"
            Height          =   285
            Index           =   0
            Left            =   270
            TabIndex        =   102
            Top             =   270
            Width           =   1215
         End
         Begin VB.CheckBox chkRFUse 
            Alignment       =   1  '������ ����
            Caption         =   "�Ϲ�/�Ű�"
            Height          =   285
            Index           =   1
            Left            =   270
            TabIndex        =   101
            Top             =   540
            Width           =   1215
         End
         Begin VB.CheckBox chkRFUse 
            Alignment       =   1  '������ ����
            Caption         =   "���ڵ��ȣ"
            Height          =   285
            Index           =   2
            Left            =   270
            TabIndex        =   100
            Top             =   810
            Width           =   1215
         End
         Begin VB.CheckBox chkRFUse 
            Alignment       =   1  '������ ����
            Caption         =   "�˻��ڵ�"
            Height          =   285
            Index           =   9
            Left            =   270
            TabIndex        =   99
            Top             =   2700
            Width           =   1215
         End
         Begin VB.CheckBox chkRFUse 
            Alignment       =   1  '������ ����
            Caption         =   "��Ϲ�ȣ"
            Height          =   285
            Index           =   3
            Left            =   270
            TabIndex        =   98
            Top             =   1080
            Width           =   1215
         End
         Begin VB.CheckBox chkRFUse 
            Alignment       =   1  '������ ����
            Caption         =   "�̸�"
            Height          =   285
            Index           =   4
            Left            =   270
            TabIndex        =   97
            Top             =   1350
            Width           =   1215
         End
         Begin VB.CheckBox chkRFUse 
            Alignment       =   1  '������ ����
            Caption         =   "����"
            Height          =   285
            Index           =   5
            Left            =   270
            TabIndex        =   96
            Top             =   1620
            Width           =   1215
         End
         Begin VB.CheckBox chkRFUse 
            Alignment       =   1  '������ ����
            Caption         =   "�ܷ�/�Կ�"
            Height          =   285
            Index           =   6
            Left            =   270
            TabIndex        =   95
            Top             =   1890
            Width           =   1215
         End
         Begin VB.CheckBox chkRFUse 
            Alignment       =   1  '������ ����
            Caption         =   "��ü"
            Height          =   285
            Index           =   7
            Left            =   270
            TabIndex        =   94
            Top             =   2160
            Width           =   1215
         End
         Begin VB.CheckBox chkRFUse 
            Alignment       =   1  '������ ����
            Caption         =   "��Ÿ"
            Height          =   285
            Index           =   8
            Left            =   270
            TabIndex        =   93
            Top             =   2430
            Width           =   1215
         End
         Begin VB.CheckBox chkRFUse 
            Alignment       =   1  '������ ����
            Caption         =   "��� 1"
            Height          =   285
            Index           =   10
            Left            =   270
            TabIndex        =   92
            Top             =   2970
            Width           =   1215
         End
         Begin VB.ComboBox cboRFord 
            Height          =   300
            Index           =   0
            Left            =   4260
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   91
            Top             =   270
            Width           =   615
         End
         Begin VB.ComboBox cboRFord 
            Height          =   300
            Index           =   1
            Left            =   4260
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   90
            Top             =   540
            Width           =   615
         End
         Begin VB.ComboBox cboRFord 
            Height          =   300
            Index           =   2
            Left            =   4260
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   89
            Top             =   810
            Width           =   615
         End
         Begin VB.ComboBox cboRFord 
            Height          =   300
            Index           =   9
            Left            =   4260
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   88
            Top             =   2700
            Width           =   615
         End
         Begin VB.ComboBox cboRFord 
            Height          =   300
            Index           =   3
            Left            =   4260
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   87
            Top             =   1080
            Width           =   615
         End
         Begin VB.ComboBox cboRFord 
            Height          =   300
            Index           =   4
            Left            =   4260
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   86
            Top             =   1350
            Width           =   615
         End
         Begin VB.ComboBox cboRFord 
            Height          =   300
            Index           =   5
            Left            =   4260
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   85
            Top             =   1620
            Width           =   615
         End
         Begin VB.ComboBox cboRFord 
            Height          =   300
            Index           =   6
            Left            =   4260
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   84
            Top             =   1890
            Width           =   615
         End
         Begin VB.ComboBox cboRFord 
            Height          =   300
            Index           =   7
            Left            =   4260
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   83
            Top             =   2160
            Width           =   615
         End
         Begin VB.ComboBox cboRFord 
            Height          =   300
            Index           =   8
            Left            =   4260
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   82
            Top             =   2430
            Width           =   615
         End
         Begin VB.ComboBox cboRFsize 
            Height          =   300
            Index           =   0
            Left            =   5880
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   81
            Top             =   270
            Width           =   615
         End
         Begin VB.ComboBox cboRFsize 
            Height          =   300
            Index           =   1
            Left            =   5880
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   80
            Top             =   540
            Width           =   615
         End
         Begin VB.ComboBox cboRFsize 
            Height          =   300
            Index           =   2
            Left            =   5880
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   79
            Top             =   810
            Width           =   615
         End
         Begin VB.ComboBox cboRFsize 
            Height          =   300
            Index           =   9
            Left            =   5880
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   78
            Top             =   2700
            Width           =   615
         End
         Begin VB.ComboBox cboRFsize 
            Height          =   300
            Index           =   3
            Left            =   5880
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   77
            Top             =   1080
            Width           =   615
         End
         Begin VB.ComboBox cboRFsize 
            Height          =   300
            Index           =   4
            Left            =   5880
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   76
            Top             =   1350
            Width           =   615
         End
         Begin VB.ComboBox cboRFsize 
            Height          =   300
            Index           =   5
            Left            =   5880
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   75
            Top             =   1620
            Width           =   615
         End
         Begin VB.ComboBox cboRFsize 
            Height          =   300
            Index           =   6
            Left            =   5880
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   74
            Top             =   1890
            Width           =   615
         End
         Begin VB.ComboBox cboRFsize 
            Height          =   300
            Index           =   7
            Left            =   5880
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   73
            Top             =   2160
            Width           =   615
         End
         Begin VB.ComboBox cboRFsize 
            Height          =   300
            Index           =   8
            Left            =   5880
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   72
            Top             =   2430
            Width           =   615
         End
         Begin VB.ComboBox cboRFord 
            Height          =   300
            Index           =   10
            Left            =   4260
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   71
            Top             =   2970
            Width           =   615
         End
         Begin VB.ComboBox cboRFsize 
            Height          =   300
            Index           =   10
            Left            =   5880
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   70
            Top             =   2970
            Width           =   615
         End
         Begin VB.Label Label8 
            BackStyle       =   0  '����
            Caption         =   "����� ���ϴ� �ʵ����        üũ�ϰ�                         üũ�� �ʵ��� �ʵ� ������ �ʵ� ũ�⸦ ����"
            Height          =   1455
            Left            =   1710
            TabIndex        =   116
            Top             =   690
            Width           =   1155
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  '����
            Caption         =   "�ʵ� ����"
            Height          =   180
            Left            =   3450
            TabIndex        =   104
            Top             =   330
            Width           =   780
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  '����
            Caption         =   "�ʵ� ũ��"
            Height          =   180
            Left            =   5070
            TabIndex        =   103
            Top             =   330
            Width           =   780
         End
      End
      Begin VB.Label Label12 
         BackStyle       =   0  '����
         Caption         =   "Result ����ϴ�    ������Ʈ"
         Height          =   375
         Left            =   300
         TabIndex        =   114
         Top             =   930
         Width           =   1665
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "Result ��� �����"
         Height          =   180
         Left            =   300
         TabIndex        =   113
         Top             =   1770
         Width           =   1545
      End
      Begin VB.Label Label10 
         BackStyle       =   0  '����
         Caption         =   "Result ��� ����� ���"
         Height          =   360
         Left            =   300
         TabIndex        =   112
         Top             =   2460
         Width           =   1545
      End
   End
End
Attribute VB_Name = "frmProgCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sPPosDig$
Dim iActiveKey%
Dim iInitRowCnt%

Private Sub DisplayInit()
    Dim i%
    Dim j%
        
    For i = 0 To MAXORDERFIELD - 1
        cboOFord(i).AddItem ""
        cboOFsize(i).AddItem ""
        
        For j = 1 To MAXORDERFIELD
            cboOFord(i).AddItem CStr(j)
        Next
        
        For j = 1 To 20
            cboOFsize(i).AddItem CStr(j)
        Next
        
        '��Ÿ ���� ����� ���� ũ��
        If i = OTHERFIELD - 1 Then
            For j = 21 To 100
                cboOFsize(i).AddItem CStr(j)
            Next
        End If
    Next
    
    For i = 0 To MAXRESULTFIELD - 1
        cboRFord(i).AddItem ""
        cboRFsize(i).AddItem ""
    
        For j = 1 To MAXRESULTFIELD
            cboRFord(i).AddItem CStr(j)
        Next
        
        For j = 1 To 20
            cboRFsize(i).AddItem CStr(j)
        Next
        
        '��Ÿ ���� ����� ���� ũ��
        If i = OTHERFIELD - 1 Then
            For j = 21 To 100
                cboRFsize(i).AddItem CStr(j)
            Next
        End If
    Next
    
    'spdPosSet �ʱ�ȭ
    With spdPosSet
        .BlockMode = True
        .Col = -1
        .Col2 = -1
        .Row = -1
        .Row2 = -1
        .BackColorStyle = BackColorStyleUnderGrid
        .BackColor = RGB(255, 255, 255)
        .EditModePermanent = True
        .EditModeReplace = True
        .NoBeep = True
        .BlockMode = False
                
        .BlockMode = True
        .Col = 1
        .Col2 = 1
        .Row = -1
        .Row2 = -1
        .Lock = True
        .BlockMode = False
    End With
End Sub

Private Sub CommCfg()
    Dim sBuf$
    Dim bRetVal As Boolean
    
'Port Setting
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "ComPort")
        
    If sBuf = "" Then
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                "Software\Ack_if\Interface Config\" & gsMachineCd, "ComPort", "1")
    
        If bRetVal = True Then
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
        End If
        
        cboPort.ListIndex = Val(sBuf) - 1
    Else
        cboPort.ListIndex = Val(sBuf) - 1
    End If

'BaudRate Setting
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "BaudRate")
        
    If sBuf = "" Then
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                "Software\Ack_if\Interface Config\" & gsMachineCd, "BaudRate", "9600")
    
        If bRetVal = True Then
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
        End If

        cboBaudRate.Text = "9600"
    Else
        cboBaudRate.Text = sBuf
    End If
    
'DataBit Setting
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "DataBit")
        
    If sBuf = "" Then
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                "Software\Ack_if\Interface Config\" & gsMachineCd, "DataBit", "8")
    
        If bRetVal = True Then
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
        End If

        OptData(1).Value = True
    Else
        OptData(CInt(sBuf) - 7).Value = True
    End If

'StopBit Setting
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "StopBit")
        
    If sBuf = "" Then
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                "Software\Ack_if\Interface Config\" & gsMachineCd, "StopBit", "1")
    
        If bRetVal = True Then
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
        End If

        OptStop(0).Value = True
    Else
        OptStop(CInt(sBuf) - 1).Value = True
    End If

'Parity Setting
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "Parity")
        
    If sBuf = "" Then
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                "Software\Ack_if\Interface Config\" & gsMachineCd, "Parity", "N")
    
        If bRetVal = True Then
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
        End If

        OptParity(0).Value = True
    Else
        If sBuf = "N" Then
            OptParity(0).Value = True
        ElseIf sBuf = "O" Then
            OptParity(1).Value = True
        ElseIf sBuf = "E" Then
            OptParity(2).Value = True
        End If
    End If

End Sub

Private Sub RackCfg()
    Dim sBuf$
    Dim bRetVal As Boolean
    Dim sPosSet$
    Dim i%
    
'RackDigit Setting
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "RackDig")
        
    If sBuf = "" Then
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                "Software\Ack_if\Interface Config\" & gsMachineCd, "RackDig", "3")
    
        If bRetVal = True Then
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
        End If

        txtRackDig = "3"
    Else
        txtRackDig = sBuf
    End If
    
'PosDigit Setting
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "PosDig")
        
    If sBuf = "" Then
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                "Software\Ack_if\Interface Config\" & gsMachineCd, "PosDig", "2")
    
        If bRetVal = True Then
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
        End If

        txtPosDig = "2"
    Else
        txtPosDig = sBuf
    End If
    
'MaxRack Setting
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "MaxRack")
        
    If sBuf = "" Then
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                "Software\Ack_if\Interface Config\" & gsMachineCd, "MaxRack", "20")
    
        If bRetVal = True Then
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
        End If
        
        txtMaxRack = 20
        spdPosSet.MaxRows = 20
        iInitRowCnt = 20
    Else
        txtMaxRack = CInt(sBuf)
        spdPosSet.MaxRows = CInt(sBuf)
        iInitRowCnt = CInt(sBuf)
    End If
    
'PosSetting
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "PosSetting")
        
    If sBuf = "" Then
        For i = 1 To spdPosSet.MaxRows
            sPosSet = sPosSet & "10" & "|"
        Next
        
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                "Software\Ack_if\Interface Config\" & gsMachineCd, "PosSetting", sPosSet)
    
        If bRetVal = True Then
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
        End If

        For i = 1 To spdPosSet.MaxRows
            Call spdPosSet.SetText(1, i, Format$(CStr(i), "000") & "")
            Call spdPosSet.SetText(2, i, "10")
        Next
    Else
        sPosSet = ""
        sPosSet = sBuf
        
        For i = 1 To spdPosSet.MaxRows
            Call spdPosSet.SetText(1, i, Format$(CStr(i), RackFormat(txtRackDig)) & "")
            Call spdPosSet.SetText(2, i, GetByOne(sPosSet, sPosSet) & "")
        Next
    End If

End Sub

Private Sub OrderResultCfg()
    On Error GoTo ErrHandler
    
    Dim sBuf$
    Dim i%
    
'Order.Use
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "Order.Use")
        
    chkOUse = sBuf
    
'Order.Component
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "Order.Component")
        
    txtOCOM = sBuf
    
'Order.Storage.Type
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "Order.Storage.Type")
        
    For i = 0 To 2
        If cboOStorage.List(i) = sBuf Then
            cboOStorage.ListIndex = i
            Exit For
        End If
    Next
    
'Order.Storage.Path
    If cboOStorage = "" Then
        txtOStoragePath = ""
    ElseIf cboOStorage = "File" Then
        sBuf = GetKeyValue(HKEY_CURRENT_USER, _
            "Software\Ack_if\Interface Config\" & gsMachineCd, "Order.FILE.Path")
            
        txtOStoragePath = sBuf
    ElseIf cboOStorage = "Database" Then
        sBuf = GetKeyValue(HKEY_CURRENT_USER, _
            "Software\Ack_if\Interface Config\" & gsMachineCd, "Order.DB.DSN")
            
        txtOStoragePath = sBuf
    Else
    End If
    
'Result.Use
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "Result.Use")
        
    chkRUse = sBuf
    
'Result.Component
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "Result.Component")
        
    txtRCOM = sBuf
    
'Result.Storage.Type
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "Result.Storage.Type")
        
    For i = 0 To 3
        If cboRStorage.List(i) = sBuf Then
            cboRStorage.ListIndex = i
            Exit For
        End If
    Next
    
'Result.Storage.Path
    If cboRStorage = "" Then
        txtRStoragePath = ""
    ElseIf cboRStorage = "File" Then
        sBuf = GetKeyValue(HKEY_CURRENT_USER, _
            "Software\Ack_if\Interface Config\" & gsMachineCd, "Result.FILE.Path")
            
        txtRStoragePath = sBuf
    ElseIf cboRStorage = "Database" Then
        sBuf = GetKeyValue(HKEY_CURRENT_USER, _
            "Software\Ack_if\Interface Config\" & gsMachineCd, "Result.DB.DSN")
            
        txtRStoragePath = sBuf
    Else
        txtRStoragePath = ""
    End If

'Order.Field.Use
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "Order.Field.Use")
    
    For i = 0 To MAXORDERFIELD - 1
        chkOFUse(i) = CStr(Val(GetByOne(sBuf, sBuf)))
    Next

'Order.Field.FName
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "Order.Field.FName")
    
    For i = 0 To MAXORDERFIELD - 2
        chkOFUse(i).Caption = CStr(GetByOne(sBuf, sBuf))
    Next
    
'Order.Field.FOrder
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "Order.Field.FOrder")
    
    For i = 0 To MAXORDERFIELD - 1
        cboOFord(i).ListIndex = Val(GetByOne(sBuf, sBuf))
    Next

'Order.Field.Size
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "Order.Field.Size")
    
    For i = 0 To MAXORDERFIELD - 1
        cboOFsize(i).ListIndex = Val(GetByOne(sBuf, sBuf))
    Next

'Result.Field.Use
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "Result.Field.Use")
    
    For i = 0 To MAXRESULTFIELD - 1
        chkRFUse(i) = CStr(Val(GetByOne(sBuf, sBuf)))
    Next

'Result.Field.FName
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "Result.Field.FName")
    
    For i = 0 To MAXRESULTFIELD - 4
        chkRFUse(i).Caption = CStr(GetByOne(sBuf, sBuf))
    Next

'Result.Field.FOrder
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "Result.Field.FOrder")
    
    For i = 0 To MAXRESULTFIELD - 1
        cboRFord(i).ListIndex = Val(GetByOne(sBuf, sBuf))
    Next

'Result.Field.Size
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "Result.Field.Size")
    
    For i = 0 To MAXRESULTFIELD - 1
        cboRFsize(i).ListIndex = Val(GetByOne(sBuf, sBuf))
    Next
    
    Exit Sub
    
ErrHandler:
    
End Sub

Private Sub SaveCommCfg()
    Dim bRetVal As Boolean
    Dim i%

'ComPort
    bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\Ack_if\Interface Config\" & gsMachineCd, "ComPort", CStr(cboPort.ListIndex + 1))
    
    If bRetVal = True Then
    Else
        MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
    End If
    
'BaudRate
    bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                "Software\Ack_if\Interface Config\" & gsMachineCd, "BaudRate", cboBaudRate.Text)
    
    If bRetVal = True Then
    Else
        MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
    End If
    
'DataBit
    If OptData(0).Value = True Then
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                "Software\Ack_if\Interface Config\" & gsMachineCd, "DataBit", "7")
    ElseIf OptData(1).Value = True Then
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                "Software\Ack_if\Interface Config\" & gsMachineCd, "DataBit", "8")
    End If
    
    If bRetVal = True Then
    Else
        MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
    End If

'StopBit
    If OptStop(0).Value = True Then
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                "Software\Ack_if\Interface Config\" & gsMachineCd, "StopBit", "1")
    ElseIf OptStop(1).Value = True Then
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                "Software\Ack_if\Interface Config\" & gsMachineCd, "StopBit", "2")
    End If
    
    If bRetVal = True Then
    Else
        MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
    End If
    
'Parity
    If OptParity(0).Value = True Then
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                "Software\Ack_if\Interface Config\" & gsMachineCd, "Parity", "N")
    ElseIf OptParity(1).Value = True Then
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                "Software\Ack_if\Interface Config\" & gsMachineCd, "Parity", "O")
    ElseIf OptParity(2).Value = True Then
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                "Software\Ack_if\Interface Config\" & gsMachineCd, "Parity", "E")
    End If
    
    If bRetVal = True Then
    Else
        MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
    End If
End Sub

Private Sub SaveRackCfg()
    Dim bRetVal As Boolean
    Dim vPosNo
    Dim sPosSet$
    Dim i%
    
'RackDig
    bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                "Software\Ack_if\Interface Config\" & gsMachineCd, "RackDig", txtRackDig)
    
    If bRetVal = True Then
    Else
        MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
    End If
    
'PosDig
    bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                "Software\Ack_if\Interface Config\" & gsMachineCd, "PosDig", txtPosDig)
    
    If bRetVal = True Then
    Else
        MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
    End If
    
'MaxRack
    bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                "Software\Ack_if\Interface Config\" & gsMachineCd, "MaxRack", txtMaxRack)
    
    If bRetVal = True Then
    Else
        MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
    End If
    
'PosSetting
    sPosSet = ""
    
    For i = 1 To spdPosSet.MaxRows
        Call spdPosSet.GetText(2, i, vPosNo)
        
        sPosSet = sPosSet & CStr(vPosNo) & "|"
    Next
    
    bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                "Software\Ack_if\Interface Config\" & gsMachineCd, "PosSetting", sPosSet)
    
    If bRetVal = True Then
    Else
        MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
    End If
End Sub

Private Sub SaveOrderResultCfg()
    Dim bRetVal As Boolean
    Dim vPosNo
    Dim sTOFUse$, sTOFName$, sTOFDOrder$, sTOFOrder$, sTOFSize$, sTRFUse$, sTRFName$, sTRFDOrder$, sTRFOrder$, sTRFSize$
    Dim i%
        
    sTOFUse = ""
    
    For i = 0 To MAXORDERFIELD - 1
    'Order.Field.Use
        sTOFUse = sTOFUse & chkOFUse(i) & Chr(124)
        
        If chkOFUse(i) = "1" Then
        'Order.Field.FOrder
            sTOFOrder = sTOFOrder & cboOFord(i) & Chr(124)
        'Order.Field.Size
            sTOFSize = sTOFSize & cboOFsize(i) & Chr(124)
        Else
            sTOFOrder = sTOFOrder & Chr(124)
            sTOFSize = sTOFSize & Chr(124)
        End If
    Next
    
    sTOFName = ""
    
    For i = 0 To MAXORDERFIELD - 2
    'Order.Field.FName
        sTOFName = sTOFName & chkOFUse(i).Caption & Chr(124)
    Next
    
    For i = 0 To MAXRESULTFIELD - 1
    'Result.Field.Use
        sTRFUse = sTRFUse & chkRFUse(i) & Chr(124)
        
        If chkRFUse(i) = "1" Then
        'Result.Field.FOrder
            sTRFOrder = sTRFOrder & cboRFord(i) & Chr(124)
        'Result.Field.Size
            sTRFSize = sTRFSize & cboRFsize(i) & Chr(124)
        Else
            sTRFOrder = sTRFOrder & Chr(124)
            sTRFSize = sTRFSize & Chr(124)
        End If
    Next
    
    sTRFName = ""
    For i = 0 To MAXRESULTFIELD - 4
    'Result.Field.FName
        sTRFName = sTRFName & chkRFUse(i).Caption & Chr(124)
    Next
    
    '�˻��ڵ��� SEQ�� ������Ʈ���� �Է�
    If chkOFUse(MAXORDERFIELD - 1) = "1" Then
        'Order.TestCd.Seq
        '�˻��ڵ�� ȭ�鿡 ���� �������� �׷��� ��
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\Ack_if\Interface Config\" & gsMachineCd, "Order.TestCd.Seq", MAXORDERFIELD)
        
        If bRetVal = True Then
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
        End If
    End If
    
    If chkRFUse(MAXRESULTFIELD - 1) = "1" Then
        'Result.TestCd.Seq
        '�˻��ڵ�� ȭ�鿡 Order�� ���� ���� ��ġ�� �׷��� ��
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\Ack_if\Interface Config\" & gsMachineCd, "Result.TestCd.Seq", MAXORDERFIELD)
        
        If bRetVal = True Then
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
        End If
    End If
    
'Order.Use
    bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                "Software\Ack_if\Interface Config\" & gsMachineCd, "Order.Use", chkOUse)
    
    If bRetVal = True Then
    Else
        MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
    End If
    
'Order ��� YES
    If chkOUse = "1" Then
    'Order.Component
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\Ack_if\Interface Config\" & gsMachineCd, "Order.Component", txtOCOM)
        
        If bRetVal = True Then
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
        End If
        
    'Order.Storage.Type
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\Ack_if\Interface Config\" & gsMachineCd, "Order.Storage.Type", cboOStorage)
        
        If bRetVal = True Then
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
        End If
        
    'Order.Storage.Path
        If cboOStorage = "" Then
        ElseIf cboOStorage = "File" Then
            bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\Ack_if\Interface Config\" & gsMachineCd, "Order.FILE.Path", txtOStoragePath)
        
        ElseIf cboOStorage = "Database" Then
            bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\Ack_if\Interface Config\" & gsMachineCd, "Order.DB.DSN", txtOStoragePath)
        
        Else
            bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\Ack_if\Interface Config\" & gsMachineCd, "Order.Storage.Type", cboOStorage)
    
        End If
        
        If bRetVal = True Then
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
        End If
                
    'Order.Field.Use
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\Ack_if\Interface Config\" & gsMachineCd, "Order.Field.Use", sTOFUse)
        
        If bRetVal = True Then
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
        End If
    
    'Order.Field.FName
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\Ack_if\Interface Config\" & gsMachineCd, "Order.Field.FName", sTOFName)
        
        If bRetVal = True Then
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
        End If
        
    'Order.Field.FOrder
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\Ack_if\Interface Config\" & gsMachineCd, "Order.Field.FOrder", sTOFOrder)
        
        If bRetVal = True Then
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
        End If
    
    'Order.Field.Size
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\Ack_if\Interface Config\" & gsMachineCd, "Order.Field.Size", sTOFSize)
        
        If bRetVal = True Then
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
        End If
                
'Order ��� NO
    Else
            'Order.Component
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\Ack_if\Interface Config\" & gsMachineCd, "Order.Component", "")
        
        If bRetVal = True Then
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
        End If
        
    'Order.Storage.Type
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\Ack_if\Interface Config\" & gsMachineCd, "Order.Storage.Type", "")
        
        If bRetVal = True Then
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
        End If
        
    'Order.Storage.Path
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                "Software\Ack_if\Interface Config\" & gsMachineCd, "Order.Storage.Path", "")
    
        If bRetVal = True Then
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
        End If
        
    'Order.Field.Use
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\Ack_if\Interface Config\" & gsMachineCd, "Order.Field.Use", "")
        
        If bRetVal = True Then
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
        End If
            
    'Order.Field.FOrder
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\Ack_if\Interface Config\" & gsMachineCd, "Order.Field.FOrder", "")
        
        If bRetVal = True Then
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
        End If
    
    'Order.Field.Size
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\Ack_if\Interface Config\" & gsMachineCd, "Order.Field.Size", "")
        
        If bRetVal = True Then
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
        End If
        
    End If
    
'Result.Use
    bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                "Software\Ack_if\Interface Config\" & gsMachineCd, "Result.Use", chkRUse)
    
    If bRetVal = True Then
    Else
        MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
    End If
    
'Result ��� YES
    If chkRUse = "1" Then
    'Result.Component
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\Ack_if\Interface Config\" & gsMachineCd, "Result.Component", txtRCOM)
        
        If bRetVal = True Then
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
        End If
        
    'Result.Storage.Type
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\Ack_if\Interface Config\" & gsMachineCd, "Result.Storage.Type", cboRStorage)
        
        If bRetVal = True Then
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
        End If
        
    'Result.Storage.Path
        If cboRStorage = "" Then
        ElseIf cboRStorage = "File" Then
            bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\Ack_if\Interface Config\" & gsMachineCd, "Result.FILE.Path", txtRStoragePath)
        
        ElseIf cboRStorage = "Database" Then
            bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\Ack_if\Interface Config\" & gsMachineCd, "Result.DB.DSN", txtRStoragePath)
        
        Else
            bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\Ack_if\Interface Config\" & gsMachineCd, "Result.Storage.Type", cboRStorage)
    
        End If
        
        If bRetVal = True Then
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
        End If
        
    'Result.Field.Use
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\Ack_if\Interface Config\" & gsMachineCd, "Result.Field.Use", sTRFUse)
        
        If bRetVal = True Then
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
        End If
    
    'Result.Field.FName
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\Ack_if\Interface Config\" & gsMachineCd, "Result.Field.FName", sTRFName)
        
        If bRetVal = True Then
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
        End If
        
    'Result.Field.FOrder
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\Ack_if\Interface Config\" & gsMachineCd, "Result.Field.FOrder", sTRFOrder)
        
        If bRetVal = True Then
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
        End If
    
    'Result.Field.Size
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\Ack_if\Interface Config\" & gsMachineCd, "Result.Field.Size", sTRFSize)
        
        If bRetVal = True Then
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
        End If
        
'Result ��� NO
    Else
    'Result.Component
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\Ack_if\Interface Config\" & gsMachineCd, "Result.Component", "")
        
        If bRetVal = True Then
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
        End If
        
    'Result.Storage.Type
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\Ack_if\Interface Config\" & gsMachineCd, "Result.Storage.Type", "")
        
        If bRetVal = True Then
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
        End If
        
    'Result.Storage.Path
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                "Software\Ack_if\Interface Config\" & gsMachineCd, "Result.Storage.Path", txtRStoragePath)
    
        If bRetVal = True Then
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
        End If
        
    'Result.Field.Use
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\Ack_if\Interface Config\" & gsMachineCd, "Result.Field.Use", sTRFUse)
        
        If bRetVal = True Then
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
        End If
            
    'Result.Field.FOrder
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\Ack_if\Interface Config\" & gsMachineCd, "Result.Field.FOrder", sTRFOrder)
        
        If bRetVal = True Then
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
        End If
    
    'Result.Field.Size
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\Ack_if\Interface Config\" & gsMachineCd, "Result.Field.Size", sTRFSize)
        
        If bRetVal = True Then
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
        End If
        
    End If
    
End Sub

Private Sub chkOFUse_Click(Index As Integer)
    Dim i%
    
    If iActiveKey = 1 Then
        If chkOFUse(Index) = "0" Then
            cboOFord(Index).Enabled = False
            cboOFsize(Index).Enabled = False
        ElseIf chkOFUse(Index) = "1" Then
            cboOFord(Index).Enabled = True
            cboOFsize(Index).Enabled = True
        End If
    End If
End Sub

Private Sub chkOFUse_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim sDefault$
    
    If KeyCode = vbKeyF2 Then
        Select Case Index
            Case 0
                sDefault = "��������"
            Case 1
                sDefault = "��������"
            Case 2
                sDefault = "��ü��ȣ"
            Case 3
                sDefault = "��Ϲ�ȣ"
            Case 4
                sDefault = "�̸�"
            Case 5
                sDefault = "����"
            Case 6
                sDefault = "����"
            Case 7
                sDefault = "���"
            Case 8
                sDefault = "��Ÿ"
        End Select
        
        chkOFUse(Index).Caption = InputBox("FIELD" & CStr(Index) & " HEADER ?", "Change Header", sDefault)
    End If
End Sub

Private Sub chkOUse_Click()
    Dim i%
    
    If iActiveKey = 1 Then
        If chkOUse = "0" Then
            txtOCOM.Enabled = False
            cboOStorage.Enabled = False
            txtOStoragePath.Enabled = False
            
            For i = 0 To MAXORDERFIELD - 1
                chkOFUse(i).Enabled = False
                cboOFord(i).Enabled = False
                cboOFsize(i).Enabled = False
            Next
        ElseIf chkOUse = "1" Then
            txtOCOM.Enabled = True
            cboOStorage.Enabled = True
            txtOStoragePath.Enabled = True
            
            For i = 0 To MAXORDERFIELD - 1
                chkOFUse(i).Enabled = True
                cboOFord(i).Enabled = True
                cboOFsize(i).Enabled = True
            Next
        End If
    End If
End Sub

Private Sub chkRFUse_Click(Index As Integer)
    Dim i%
    
    If iActiveKey = 1 Then
        If chkRFUse(Index) = "0" Then
            cboRFord(Index).Enabled = False
            cboRFsize(Index).Enabled = False
        ElseIf chkRFUse(Index) = "1" Then
            cboRFord(Index).Enabled = True
            cboRFsize(Index).Enabled = True
        End If
    End If
End Sub

Private Sub chkRFUse_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim sDefault$
    
    If KeyCode = vbKeyF2 Then
        Select Case Index
            Case 0
                sDefault = "��������"
            Case 1
                sDefault = "��������"
            Case 2
                sDefault = "��ü��ȣ"
            Case 3
                sDefault = "��Ϲ�ȣ"
            Case 4
                sDefault = "�̸�"
            Case 5
                sDefault = "����"
            Case 6
                sDefault = "����"
            Case 7
                sDefault = "���"
            Case 8
                sDefault = "��Ÿ"
        End Select
        
        chkRFUse(Index).Caption = InputBox("FIELD" & CStr(Index) & " HEADER ?", "Change Header", sDefault)
    End If
End Sub

Private Sub chkRUse_Click()
    Dim i%
    
    If iActiveKey = 1 Then
        If chkRUse = "0" Then
            txtRCOM.Enabled = False
            cboRStorage.Enabled = False
            txtRStoragePath.Enabled = False
            
            For i = 0 To MAXRESULTFIELD - 1
                chkRFUse(i).Enabled = False
                cboRFord(i).Enabled = False
                cboRFsize(i).Enabled = False
            Next
        ElseIf chkRUse = "1" Then
            txtRCOM.Enabled = True
            cboRStorage.Enabled = True
            txtRStoragePath.Enabled = True
            
            For i = 0 To MAXRESULTFIELD - 1
                chkRFUse(i).Enabled = True
                cboRFord(i).Enabled = True
                cboRFsize(i).Enabled = True
            Next
        End If
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Call SaveCommCfg
    Call SaveRackCfg
    Call SaveOrderResultCfg
    
    MsgBox "�����۾��� ���������� ����Ǿ����ϴ�!!"
End Sub

Private Sub Form_Activate()
    iActiveKey = 1
End Sub

Private Sub Form_DblClick()
    frmIntMcd.Left = 9000
    frmIntMcd.Top = 5500
    Load frmIntMcd
    frmIntMcd.Show vbModal, frmProgCfg
End Sub

Private Sub Form_Load()
    iActiveKey = 0
    
    Call DisplayInit
    Call CommCfg
    Call RackCfg
    Call OrderResultCfg

    iActiveKey = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    RegEditCurFrmTitle "ProgCfg", ""
    ViewMsg ""
End Sub

Private Sub txtMaxRack_Change()
    Dim i%
    
    If iActiveKey = 1 Then
        If Len(txtMaxRack) = txtMaxRack.MaxLength Then
            spdPosSet.MaxRows = CInt(txtMaxRack)
            
            If spdPosSet.MaxRows > iInitRowCnt Then
                For i = iInitRowCnt + 1 To spdPosSet.MaxRows
                    Call spdPosSet.SetText(1, i, Format$(CStr(i), RackFormat(txtRackDig)) & "")
                Next
            End If
            
            spdPosSet.SetFocus
            If spdPosSet.MaxRows > iInitRowCnt Then
                spdPosSet.Col = 2
                spdPosSet.Row = iInitRowCnt + 1
            Else
                spdPosSet.Col = 2
                spdPosSet.Row = 1
            End If
            spdPosSet.Action = SS_ACTION_ACTIVE_CELL
        End If
    End If
End Sub

Private Sub txtMaxRack_Click()
    Call Txt_Highlight(txtMaxRack)
    iInitRowCnt = Val(txtMaxRack)
End Sub

Private Sub txtMaxRack_GotFocus()
    Call Txt_Highlight(txtMaxRack)
    iInitRowCnt = Val(txtMaxRack)
End Sub

Private Sub txtPosDig_Change()
    Dim i%
    Dim vTmp
    
    If iActiveKey = 1 Then
        If Len(txtPosDig) = txtPosDig.MaxLength Then
            With spdPosSet
                For i = 1 To .MaxRows
                    Call .GetText(2, i, vTmp)
                    
                    If Len(vTmp) > CInt(txtPosDig) Then
                        MsgBox "PosDigit�� Pos������ �۽��ϴ�. " & vbCrLf & _
                         "Pos������ �����Ͻʽÿ�!!"
                         txtPosDig = sPPosDig
                         Exit For
                    End If
                Next
            End With
        End If
        
        txtMaxRack.SetFocus
    End If
End Sub

Private Sub txtPosDig_Click()
    Call Txt_Highlight(txtPosDig)
    sPPosDig = txtPosDig
End Sub

Private Sub txtPosDig_GotFocus()
    Call Txt_Highlight(txtPosDig)
    sPPosDig = txtPosDig
End Sub

Private Sub txtRackDig_Change()
    Dim i%
    
    If iActiveKey = 1 Then
        If Len(txtRackDig) = txtRackDig.MaxLength Then
        
            With spdPosSet
                For i = 1 To .MaxRows
                    Call .SetText(1, i, Format$(CStr(i), RackFormat(txtRackDig)) & "")
                Next
            End With
        End If
        
        txtPosDig.SetFocus
    End If
End Sub

Private Sub txtRackDig_Click()
    Call Txt_Highlight(txtRackDig)
End Sub

Private Sub txtRackDig_GotFocus()
    Call Txt_Highlight(txtRackDig)
End Sub
