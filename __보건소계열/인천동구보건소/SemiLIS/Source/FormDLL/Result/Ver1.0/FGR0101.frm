VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FGR0101 
   BorderStyle     =   0  '����
   Caption         =   "Form2"
   ClientHeight    =   7845
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11925
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   7845
   ScaleWidth      =   11925
   ShowInTaskbar   =   0   'False
   Begin Threed.SSFrame SSFrame1 
      Height          =   1695
      Left            =   3600
      TabIndex        =   20
      Top             =   30
      Width           =   8295
      _Version        =   65536
      _ExtentX        =   14631
      _ExtentY        =   2990
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
      Begin VB.TextBox txt�Ұ� 
         Height          =   855
         Left            =   5160
         MultiLine       =   -1  'True
         ScrollBars      =   2  '����
         TabIndex        =   51
         Top             =   750
         Width           =   3075
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   285
         Left            =   60
         TabIndex        =   24
         Top             =   150
         Width           =   3195
         _Version        =   65536
         _ExtentX        =   5636
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "   Patient Info....."
         ForeColor       =   12648447
         BackColor       =   32896
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         BorderWidth     =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         Alignment       =   2
      End
      Begin Threed.SSPanel SSPanel15 
         Height          =   285
         Left            =   150
         TabIndex        =   25
         Top             =   450
         Width           =   945
         _Version        =   65536
         _ExtentX        =   1667
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "�۾���ȣ"
         ForeColor       =   0
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
         BorderWidth     =   1
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   285
         Left            =   150
         TabIndex        =   29
         Top             =   750
         Width           =   945
         _Version        =   65536
         _ExtentX        =   1667
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "��Ϲ�ȣ"
         ForeColor       =   0
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
         BorderWidth     =   1
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   285
         Left            =   150
         TabIndex        =   30
         Top             =   1050
         Width           =   945
         _Version        =   65536
         _ExtentX        =   1667
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "��  ��"
         ForeColor       =   0
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
         BorderWidth     =   1
      End
      Begin Threed.SSPanel SSPanel13 
         Height          =   285
         Left            =   150
         TabIndex        =   31
         Top             =   1350
         Width           =   945
         _Version        =   65536
         _ExtentX        =   1667
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "��  ��"
         ForeColor       =   0
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
         BorderWidth     =   1
         FloodColor      =   12582912
      End
      Begin Threed.SSPanel SSPanel14 
         Height          =   285
         Left            =   1830
         TabIndex        =   32
         Top             =   1350
         Width           =   945
         _Version        =   65536
         _ExtentX        =   1667
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "��  ��"
         ForeColor       =   0
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
         BorderWidth     =   1
         FloodColor      =   12582912
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   285
         Left            =   3300
         TabIndex        =   39
         Top             =   150
         Width           =   945
         _Version        =   65536
         _ExtentX        =   1667
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "�����"
         ForeColor       =   0
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
         BorderWidth     =   1
         FloodColor      =   12582912
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   285
         Left            =   3300
         TabIndex        =   41
         Top             =   450
         Width           =   945
         _Version        =   65536
         _ExtentX        =   1667
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "����(����)"
         ForeColor       =   0
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
         BorderWidth     =   1
         FloodColor      =   12582912
      End
      Begin Threed.SSPanel SSPanel7 
         Height          =   285
         Left            =   3300
         TabIndex        =   43
         Top             =   750
         Width           =   945
         _Version        =   65536
         _ExtentX        =   1667
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "��������"
         ForeColor       =   0
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
         BorderWidth     =   1
         FloodColor      =   12582912
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   285
         Left            =   3300
         TabIndex        =   45
         Top             =   1050
         Width           =   945
         _Version        =   65536
         _ExtentX        =   1667
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "���ޱ���"
         ForeColor       =   0
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
         BorderWidth     =   1
         FloodColor      =   12582912
      End
      Begin Threed.SSPanel SSPanel12 
         Height          =   285
         Left            =   3300
         TabIndex        =   47
         Top             =   1350
         Width           =   945
         _Version        =   65536
         _ExtentX        =   1667
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "Ư������"
         ForeColor       =   0
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
         BorderWidth     =   1
         FloodColor      =   12582912
      End
      Begin Threed.SSPanel SSPanel16 
         Height          =   285
         Left            =   6000
         TabIndex        =   49
         Top             =   150
         Width           =   945
         _Version        =   65536
         _ExtentX        =   1667
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "�ǻ��"
         ForeColor       =   0
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
         BorderWidth     =   1
         FloodColor      =   12582912
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   285
         Left            =   6000
         TabIndex        =   52
         Top             =   450
         Width           =   2235
         _Version        =   65536
         _ExtentX        =   3942
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "�����ȣ(�ӻ�Ұ�)"
         ForeColor       =   0
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
         BorderWidth     =   1
         FloodColor      =   12582912
      End
      Begin VB.Label lbl�ǻ� 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '���� ����
         Caption         =   "ȭŸ"
         Height          =   270
         Left            =   6960
         TabIndex        =   50
         Top             =   150
         Width           =   1245
      End
      Begin VB.Label lblƯ������ 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '���� ����
         Caption         =   "Y"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4260
         TabIndex        =   48
         Top             =   1350
         Width           =   255
      End
      Begin VB.Label lbl���ޱ��� 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '���� ����
         Caption         =   "Y"
         Height          =   270
         Left            =   4260
         TabIndex        =   46
         Top             =   1050
         Width           =   255
      End
      Begin VB.Label lbl�������� 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '���� ����
         Caption         =   "�ܷ�"
         Height          =   270
         Left            =   4260
         TabIndex        =   44
         Top             =   750
         Width           =   825
      End
      Begin VB.Label lbl���� 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '���� ����
         Caption         =   "��1����"
         Height          =   270
         Left            =   4260
         TabIndex        =   42
         Top             =   450
         Width           =   1665
      End
      Begin VB.Label lbl����� 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '���� ����
         Caption         =   "�񴢱��"
         Height          =   270
         Left            =   4260
         TabIndex        =   40
         Top             =   150
         Width           =   1665
      End
      Begin VB.Label lbl�̸� 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '���� ����
         Caption         =   "���ÿ��о��ֱ�"
         Height          =   270
         Left            =   1110
         TabIndex        =   37
         Top             =   1050
         Width           =   1575
      End
      Begin VB.Label lbl��Ϲ�ȣ 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '���� ����
         Caption         =   "720121-1840518"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1110
         TabIndex        =   36
         Top             =   750
         Width           =   1575
      End
      Begin VB.Label lbl���� 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '���� ����
         Caption         =   "120"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1110
         TabIndex        =   34
         Top             =   1350
         Width           =   675
      End
      Begin VB.Label lbl���� 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '���� ����
         Caption         =   "��"
         Height          =   270
         Left            =   2790
         TabIndex        =   33
         Top             =   1350
         Width           =   315
      End
      Begin VB.Label lbl�۾����� 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '���� ����
         Caption         =   "19980105"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1110
         TabIndex        =   28
         Top             =   450
         Width           =   945
      End
      Begin VB.Label lblSlip 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '���� ����
         Caption         =   "H01"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2070
         TabIndex        =   27
         Top             =   450
         Width           =   405
      End
      Begin VB.Label lbl�۾���ȣ 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '���� ����
         Caption         =   "00001"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2490
         TabIndex        =   26
         Top             =   450
         Width           =   615
      End
   End
   Begin Threed.SSFrame SSFrame6 
      Height          =   7785
      Left            =   30
      TabIndex        =   7
      Top             =   30
      Width           =   3555
      _Version        =   65536
      _ExtentX        =   6271
      _ExtentY        =   13732
      _StockProps     =   14
      ForeColor       =   4194304
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin FPSpread.vaSpread spdList 
         Height          =   4335
         Left            =   60
         OleObjectBlob   =   "FGR0101.frx":0000
         TabIndex        =   8
         Top             =   3120
         Width           =   3435
      End
      Begin Threed.SSFrame SSFrame7 
         Height          =   1515
         Left            =   60
         TabIndex        =   14
         Top             =   570
         Width           =   3435
         _Version        =   65536
         _ExtentX        =   6059
         _ExtentY        =   2672
         _StockProps     =   14
         Caption         =   "����Է� ��ȸ Option"
         ForeColor       =   4194304
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.OptionButton optGbn 
            Appearance      =   0  '���
            BackColor       =   &H00C0FFFF&
            Caption         =   "�ش� �۾���ȣ"
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   0
            Left            =   630
            TabIndex        =   1
            Top             =   300
            Width           =   2055
         End
         Begin VB.OptionButton optGbn 
            Appearance      =   0  '���
            BackColor       =   &H00C0FFC0&
            Caption         =   "�ش��� ��Ϲ�ȣ"
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   1
            Left            =   630
            TabIndex        =   2
            Top             =   690
            Width           =   2055
         End
         Begin VB.OptionButton optGbn 
            Appearance      =   0  '���
            BackColor       =   &H00C0C0FF&
            Caption         =   "��� ��/�̿� List"
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   2
            Left            =   630
            TabIndex        =   3
            Top             =   1080
            Width           =   2055
         End
      End
      Begin MSComCtl2.DTPicker dtpLabDate 
         Height          =   315
         Left            =   1170
         TabIndex        =   0
         Top             =   180
         Width           =   1485
         _ExtentX        =   2619
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
         CustomFormat    =   "yyy-MM-dd"
         Format          =   24641539
         CurrentDate     =   36165
      End
      Begin Threed.SSPanel pnlLabDate 
         Height          =   375
         Left            =   90
         TabIndex        =   35
         Top             =   150
         Width           =   1065
         _Version        =   65536
         _ExtentX        =   1879
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "��������"
         ForeColor       =   8454143
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
      End
      Begin Threed.SSPanel pnl�۾���ȣ 
         Height          =   825
         Left            =   60
         TabIndex        =   9
         Top             =   2190
         Width           =   3435
         _Version        =   65536
         _ExtentX        =   6059
         _ExtentY        =   1455
         _StockProps     =   15
         ForeColor       =   4194304
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
         Begin VB.TextBox txt�۾����� 
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
            Left            =   990
            MaxLength       =   8
            TabIndex        =   4
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txtSlip 
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
            IMEMode         =   8  '����
            Left            =   1950
            MaxLength       =   3
            TabIndex        =   5
            Top             =   240
            Width           =   465
         End
         Begin VB.TextBox txt�۾���ȣ 
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
            Left            =   2700
            MaxLength       =   5
            TabIndex        =   6
            Top             =   240
            Width           =   645
         End
         Begin Threed.SSPanel SSPanel25 
            Height          =   315
            Left            =   90
            TabIndex        =   10
            Top             =   240
            Width           =   885
            _Version        =   65536
            _ExtentX        =   1561
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "�۾���ȣ"
            ForeColor       =   0
            BackColor       =   12648447
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   1
         End
         Begin Threed.SSCommand cmdButtonSlip 
            Height          =   330
            Left            =   2430
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   240
            Width           =   270
            _Version        =   65536
            _ExtentX        =   476
            _ExtentY        =   582
            _StockProps     =   78
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            RoundedCorners  =   0   'False
            Picture         =   "FGR0101.frx":06CD
         End
      End
      Begin Threed.SSPanel pnl��Ϲ�ȣ 
         Height          =   825
         Left            =   60
         TabIndex        =   11
         Top             =   2190
         Width           =   3435
         _Version        =   65536
         _ExtentX        =   6059
         _ExtentY        =   1455
         _StockProps     =   15
         ForeColor       =   4194304
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
         Begin VB.TextBox txt��Ϲ�ȣ 
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
            Left            =   1290
            MaxLength       =   15
            TabIndex        =   13
            Top             =   240
            Width           =   1755
         End
         Begin Threed.SSPanel SSPanel23 
            Height          =   315
            Left            =   210
            TabIndex        =   12
            Top             =   240
            Width           =   1065
            _Version        =   65536
            _ExtentX        =   1879
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "��Ϲ�ȣ"
            ForeColor       =   0
            BackColor       =   12648384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   1
         End
      End
      Begin Threed.SSPanel pnl�� 
         Height          =   825
         Left            =   60
         TabIndex        =   15
         Top             =   2190
         Width           =   3435
         _Version        =   65536
         _ExtentX        =   6059
         _ExtentY        =   1455
         _StockProps     =   15
         ForeColor       =   4194304
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
         Begin VB.OptionButton opt�� 
            Caption         =   "��"
            Height          =   525
            Index           =   0
            Left            =   1920
            TabIndex        =   17
            Top             =   150
            Width           =   525
         End
         Begin Threed.SSPanel SSPanel26 
            Height          =   315
            Left            =   180
            TabIndex        =   18
            Top             =   270
            Width           =   1665
            _Version        =   65536
            _ExtentX        =   2937
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "��� ��/�̿� List"
            ForeColor       =   0
            BackColor       =   12632319
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   1
         End
         Begin VB.OptionButton opt�� 
            Caption         =   "�̿�"
            Height          =   525
            Index           =   1
            Left            =   2550
            TabIndex        =   16
            Top             =   150
            Width           =   675
         End
      End
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   4935
      Left            =   3600
      TabIndex        =   23
      Top             =   1650
      Width           =   8295
      _Version        =   65536
      _ExtentX        =   14631
      _ExtentY        =   8705
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
      Begin FPSpread.vaSpread spdRst 
         Height          =   3795
         Left            =   150
         OleObjectBlob   =   "FGR0101.frx":07EF
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   765
         Width           =   6465
      End
      Begin Threed.SSFrame SSFrame3 
         Height          =   1455
         Left            =   6660
         TabIndex        =   54
         Top             =   120
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   2566
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
         Begin Threed.SSCommand cmdReg 
            Height          =   615
            Left            =   90
            TabIndex        =   59
            Top             =   150
            Width           =   1395
            _Version        =   65536
            _ExtentX        =   2461
            _ExtentY        =   1085
            _StockProps     =   78
            Caption         =   "����        F8  "
            ForeColor       =   16576
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   3
            Picture         =   "FGR0101.frx":2932
         End
         Begin Threed.SSCommand cmdExit 
            Height          =   615
            Left            =   90
            TabIndex        =   60
            Top             =   750
            Width           =   1395
            _Version        =   65536
            _ExtentX        =   2461
            _ExtentY        =   1085
            _StockProps     =   78
            Caption         =   "����      ESC "
            ForeColor       =   12582912
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   3
            Picture         =   "FGR0101.frx":3784
         End
      End
      Begin Threed.SSFrame SSFrame5 
         Height          =   1485
         Left            =   6660
         TabIndex        =   70
         Top             =   1470
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   2619
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
         Begin Threed.SSCommand cmdReRun 
            Height          =   615
            Left            =   90
            TabIndex        =   71
            Top             =   180
            Width           =   1395
            _Version        =   65536
            _ExtentX        =   2461
            _ExtentY        =   1085
            _StockProps     =   78
            Caption         =   "��           �� "
            ForeColor       =   128
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   3
            Picture         =   "FGR0101.frx":3BD6
         End
         Begin Threed.SSCommand cmdSlipPrint 
            Height          =   615
            Left            =   90
            TabIndex        =   72
            Top             =   780
            Width           =   1395
            _Version        =   65536
            _ExtentX        =   2461
            _ExtentY        =   1085
            _StockProps     =   78
            Caption         =   "  ���       ���  "
            ForeColor       =   8388608
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   3
            Picture         =   "FGR0101.frx":3EF0
         End
      End
      Begin Threed.SSFrame fraOthers 
         Height          =   2025
         Left            =   6660
         TabIndex        =   66
         Top             =   2850
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   3572
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
         Begin Threed.SSCommand cmdSecond 
            Height          =   615
            Left            =   90
            TabIndex        =   67
            Top             =   150
            Width           =   1395
            _Version        =   65536
            _ExtentX        =   2461
            _ExtentY        =   1085
            _StockProps     =   78
            Caption         =   "2���˻��׸�"
            BevelWidth      =   3
         End
         Begin Threed.SSCommand cmdMDC 
            Height          =   615
            Left            =   90
            TabIndex        =   68
            Top             =   750
            Width           =   1395
            _Version        =   65536
            _ExtentX        =   2461
            _ExtentY        =   1085
            _StockProps     =   78
            Caption         =   "M.D.C"
            BevelWidth      =   3
         End
         Begin Threed.SSCommand cmdMorph 
            Height          =   615
            Left            =   90
            TabIndex        =   69
            Top             =   1350
            Width           =   1395
            _Version        =   65536
            _ExtentX        =   2461
            _ExtentY        =   1085
            _StockProps     =   78
            Caption         =   "Morphology"
            BevelWidth      =   3
         End
      End
      Begin Threed.SSPanel SSPanel10 
         Height          =   285
         Left            =   60
         TabIndex        =   53
         Top             =   150
         Width           =   3195
         _Version        =   65536
         _ExtentX        =   5636
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "   Result View and Edit....."
         ForeColor       =   65535
         BackColor       =   128
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         BorderWidth     =   0
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         FloodColor      =   0
         Alignment       =   2
      End
      Begin Threed.SSPanel SSPanel11 
         Height          =   285
         Left            =   3300
         TabIndex        =   55
         Top             =   150
         Width           =   1245
         _Version        =   65536
         _ExtentX        =   2196
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "����Ϸ� �ð�"
         ForeColor       =   12640511
         BackColor       =   16576
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         Alignment       =   8
      End
      Begin Threed.SSPanel SSPanel17 
         Height          =   285
         Left            =   150
         TabIndex        =   57
         Top             =   450
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "�� ü ��"
         ForeColor       =   12640511
         BackColor       =   16576
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         Alignment       =   8
      End
      Begin Threed.SSPanel SSPanel18 
         Height          =   285
         Left            =   150
         TabIndex        =   61
         Top             =   4590
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "�� �� ġ"
         ForeColor       =   12640511
         BackColor       =   16576
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         Alignment       =   8
      End
      Begin Threed.SSPanel SSPanel19 
         Height          =   285
         Left            =   3870
         TabIndex        =   64
         Top             =   4590
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "��  ��"
         ForeColor       =   12640511
         BackColor       =   16576
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         Alignment       =   8
      End
      Begin VB.Label lbl���� 
         Alignment       =   2  '��� ����
         BackColor       =   &H00EAEAFF&
         BorderStyle     =   1  '���� ����
         Caption         =   "12.5 - 84.7"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4920
         TabIndex        =   65
         Top             =   4590
         Width           =   1695
      End
      Begin VB.Label lbl����ġ 
         Alignment       =   2  '��� ����
         BackColor       =   &H00EAEAFF&
         BorderStyle     =   1  '���� ����
         Caption         =   "12.5 - 84.7"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1200
         TabIndex        =   62
         Top             =   4590
         Width           =   2625
      End
      Begin VB.Label lbl��ü�� 
         BackColor       =   &H00EAEAFF&
         BorderStyle     =   1  '���� ����
         Caption         =   "EDTA WHOLE BLOOD EDTA WHOLE BLOOD"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1200
         TabIndex        =   58
         Top             =   450
         Width           =   5415
      End
      Begin VB.Label lbl�ð� 
         BackColor       =   &H00EAEAFF&
         BorderStyle     =   1  '���� ����
         Caption         =   "1999-01-05 18:25:34"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4560
         TabIndex        =   56
         Top             =   150
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1350
      Left            =   3600
      TabIndex        =   21
      Top             =   6480
      Width           =   8295
      Begin FPSpread.vaSpread spdCmt 
         Height          =   825
         Left            =   150
         OleObjectBlob   =   "FGR0101.frx":4342
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   450
         Width           =   8085
      End
      Begin Threed.SSCommand cmdAddCmt 
         Height          =   285
         Left            =   3240
         TabIndex        =   73
         Top             =   150
         Width           =   1605
         _Version        =   65536
         _ExtentX        =   2831
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "Add Comment"
         ForeColor       =   32768
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
      Begin Threed.SSPanel SSPanel5 
         Height          =   285
         Left            =   60
         TabIndex        =   38
         Top             =   150
         Width           =   3135
         _Version        =   65536
         _ExtentX        =   5530
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "   Comment....."
         ForeColor       =   12648384
         BackColor       =   32768
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         BorderWidth     =   0
         BevelInner      =   1
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
         Alignment       =   2
      End
   End
End
Attribute VB_Name = "FGR0101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type �����
    ������ڵ� As String
    ������� As String
End Type

Private Type ��ü��
    ��ü�ڵ� As String
    ��ü��� As String
    ��ü���� As String
End Type

Private Type �Ұ���
    �Ұ��ڵ� As String
    �Ұ߳��� As String
End Type

Private Type FLAGDETAIL
    STATUSCD As String
    �̻���� As String
    ������� As String
    ���� As String
    ����� As String
End Type

Private Type FLAGMAIN
    �����ڵ� As String
    FLAG���� As Integer
    FLAG(1 To 10) As FLAGDETAIL
End Type

Private Type PATFLAG
    �˻��ڵ� As String
    �����ڵ� As String
    STATUSCD As String
    ���� As String
    ����� As String
End Type

Dim ����() As �����
Dim ��ü() As ��ü��
Dim �Ұ�() As �Ұ���
Dim FLAG() As FLAGDETAIL
Dim FLAGCHK() As FLAGMAIN
Dim PFLAG() As PATFLAG

Dim ����Cnt%
Dim ��üCnt%
Dim �Ұ�Cnt%
Dim FLAGCnt%
Dim iHlpClick%
Dim iDtpChange%
Dim sPrevSlipCd$
Dim iColorCnt%
Dim iCurRow%
Dim iReRunYN%

Private Sub ChkFromSearchFrm()
    Dim sBuf$

'<----------------- �޴� �� �����ڷ��� OTHER�� ��� ���θ� Registry�� ���� �о� �Ǵ� ----------->
    
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
                    "Software\SemiLIS\Program Config\Menu.Setting\Others.Visible", "Check YN")
    
    
    If sBuf = "" Then     '���� ������Ʈ��Ű�� ���� �� Default �� ���
'        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
'                      "Software\SemiLIS\Program Config\Menu.Setting\Others.Visible", "Check YN", "N")
'
'        If bRetVal = True Then
'            fraOthers.Visible = False
'        Else
'            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
'        End If
    End If
'<---------------------------------------------------------------------------------------->
    
End Sub

Private Sub Com_E_CdHlp(ByVal iRow As Integer)
    Dim CCom As DCB0101
    Dim j%, i%
    Dim sTot01$, sTot02$, sTot03$, sTot04$, sTot05$
    Dim sTmp1$
    Dim vPartGbn
    
    Set CCom = New DCB0101
    
    Call spdList.GetText(2, iCurRow, vPartGbn)
    
    CCom.Get_COMCD Left$(CStr(vPartGbn), 1), "E"
        
    j = CCom.CurItemCnt
    
    iHlpClick = 1
    
    Erase gCodeHlpTable '�迭 �ʱ�ȭ
    
    ReDim gCodeHlpTable(j) As CodeTBL
    
    With CCom
        sTot01 = .TotField01    'CommentCd
        sTot02 = .TotField02    'CommentNm
    End With
    
    Set CCom = Nothing

    For i = 1 To j
        gCodeHlpTable(i).sSeq = Format$(i, "00000")
        gCodeHlpTable(i).sCode = Left$(CStr(vPartGbn), 1) & GetByOne(sTot01, sTot01)
        gCodeHlpTable(i).sCodeNm = GetByOne(sTot02, sTot02)
    Next

    giCodeHlpCnt = j
    giCodeHlpMode = 2
    giCallSpdRow = iRow
    
    Set gCallObject = FGR0101.spdRst
    
    FSR0201.Left = 7800
    FSR0201.Top = 2000
    
    Load FSR0201
    FSR0201.Show vbModal
    
    giCallSpdRow = 0
    
    FGR0101.spdRst.SetFocus
    FGR0101.spdRst.Row = iRow
    FGR0101.spdRst.Col = 2
    FGR0101.spdRst.Action = SS_ACTION_ACTIVE_CELL

End Sub

Private Sub Com_S_CdHlp(ByVal iRow As Integer)
    Dim CCom As DCB0101
    Dim j%, i%
    Dim sTot01$, sTot02$, sTot03$, sTot04$, sTot05$
    Dim sTmp1$
    Dim vPartGbn
    
    Set CCom = New DCB0101
    
    Call spdList.GetText(2, iCurRow, vPartGbn)
    
    CCom.Get_COMCD Left$(CStr(vPartGbn), 1), "S"
        
    j = CCom.CurItemCnt
    
    iHlpClick = 1
    
    Erase gCodeHlpTable '�迭 �ʱ�ȭ
    
    ReDim gCodeHlpTable(j) As CodeTBL
    
    With CCom
        sTot01 = .TotField01    'CommentCd
        sTot02 = .TotField02    'CommentNm
    End With
    
    Set CCom = Nothing

    For i = 1 To j
        gCodeHlpTable(i).sSeq = Format$(i, "00000")
        gCodeHlpTable(i).sCode = Left$(CStr(vPartGbn), 1) & GetByOne(sTot01, sTot01)
        gCodeHlpTable(i).sCodeNm = GetByOne(sTot02, sTot02)
    Next

    giCodeHlpCnt = j
    giCodeHlpMode = 3
    giCallSpdRow = iRow
    
    Set gCallObject = FGR0101.spdCmt
    
    FSR0201.Left = 5000
    FSR0201.Top = 2000
    
    Load FSR0201
    FSR0201.Show vbModal
    
    giCallSpdRow = 0
    
    FGR0101.spdCmt.SetFocus
    FGR0101.spdCmt.Row = iRow
    FGR0101.spdCmt.Col = 3
    FGR0101.spdCmt.Action = SS_ACTION_ACTIVE_CELL

End Sub

Private Sub Get�����()
    Dim C����� As DCR0101
    Dim i%, j%
    Dim sField01$, sField02$
    
    Set C����� = New DCR0101
        
    C�����.Get_�����
    
    j = C�����.CurItemCnt
    
    If j = 0 Then
        ViewMsg ("����� �����ڷᰡ �����ϴ�.")
        Set C����� = Nothing
        Exit Sub
    End If
        
    sField01 = C�����.Tot�����
    sField02 = C�����.Tot�������

    ����Cnt = j
    
    ReDim ����(1 To j) As �����
    
    For i = 1 To j
        ����(i).������ڵ� = GetByOne(sField01, sField01)
        ����(i).������� = GetByOne(sField02, sField02)
    Next
    
    Set C����� = Nothing
    
End Sub

Private Sub GetPatient()
    Dim CPat As DCR0101
    Dim i%, j%
    Dim sField01$, sField02$, sField03$, sField04$, sField05$
    Dim sField06$, sField07$, sField08$, sField09$, sField10$
    Dim sField11$, sField12$, sField13$, sField14$, sField15$
    Dim sField16$, sField17$, sField19$
    Dim tmpCd  As String
    
    Set CPat = New DCR0101
        
    If optGbn(0).Value = True Then
        CPat.Get_Pat 1, IIf(pnlLabDate = "��������", 1, 2), Trim(Format(dtpLabDate.Value, "YYYYMMDD")), Trim(txt�۾�����), Trim(Right(txtSlip, 2)), Trim(txt�۾���ȣ), Left$(txtSlip, 1)
    ElseIf optGbn(1).Value = True Then
        CPat.Get_Pat 2, IIf(pnlLabDate = "��������", 1, 2), Trim(Format(dtpLabDate.Value, "YYYYMMDD")), Trim(txt��Ϲ�ȣ), "", "", fCurUserPartCd
    ElseIf optGbn(2).Value = True Then
        If opt��(0).Value = True Then
            CPat.Get_Pat 3, IIf(pnlLabDate = "��������", 1, 2), Trim(Format(dtpLabDate.Value, "YYYYMMDD")), "", "", "", fCurUserPartCd
        Else
            CPat.Get_Pat 4, IIf(pnlLabDate = "��������", 1, 2), Trim(Format(dtpLabDate.Value, "YYYYMMDD")), "", "", "", fCurUserPartCd
        End If
    End If
    
    j = CPat.CurItemCnt
    
    If j = 0 Then
        ViewMsg "�ش系���� �����ϴ�..."
        Set CPat = Nothing
        Exit Sub
    End If
        
    sField01 = CPat.TotLABDATE
    sField02 = CPat.TotPARTGBN
    sField03 = CPat.TotLABSEQ
    sField04 = CPat.Tot��Ϲ�ȣ
    sField05 = CPat.Tot�̸�
    sField06 = CPat.Tot����
    sField07 = CPat.Tot����
    sField08 = CPat.Tot�����
    sField09 = CPat.Tot����
    sField10 = CPat.Tot��������
    sField11 = CPat.Tot����
    sField12 = CPat.TotƯ��
    sField13 = CPat.Tot�ǻ�
    sField14 = CPat.Tot�Ұ�
    sField15 = CPat.Tot��ü�ڵ�
    sField16 = CPat.Tot�����
    sField17 = CPat.Tot����ð�
    sField19 = CPat.Tot����Ϸ�
    
    Set CPat = Nothing
    
    'spdList�� ȯ�ڸ� �Ѹ���
    For i = 1 To j
        spdList.MaxRows = spdList.MaxRows + 1
        Call spdList.SetText(1, i, GetByOne(sField01, sField01))
        tmpCd = GetByOne(sField02, sField02)
        If tmpCd = "01" Then
            Call spdList.SetText(2, i, "C01")
        ElseIf tmpCd = "02" Then
            Call spdList.SetText(2, i, "H02")
        Else
            Call spdList.SetText(2, i, "U04")
        End If
        'Call spdList.SetText(2, i, Left$(txtSlip, 1) & GetByOne(sField02, sField02))
        Call spdList.SetText(3, i, GetByOne(sField03, sField03))
        Call spdList.SetText(4, i, GetByOne(sField04, sField04))
        Call spdList.SetText(5, i, GetByOne(sField05, sField05))
        Call spdList.SetText(6, i, GetByOne(sField06, sField06))
        Call spdList.SetText(7, i, GetByOne(sField07, sField07))
        Call spdList.SetText(8, i, GetByOne(sField08, sField08))
        Call spdList.SetText(9, i, GetByOne(sField09, sField09))
        Call spdList.SetText(10, i, GetByOne(sField10, sField10))
        Call spdList.SetText(11, i, GetByOne(sField11, sField11))
        Call spdList.SetText(12, i, GetByOne(sField12, sField12))
        Call spdList.SetText(13, i, GetByOne(sField13, sField13))
        Call spdList.SetText(14, i, GetByOne(sField14, sField14))
        Call spdList.SetText(15, i, GetByOne(sField15, sField15))
        Call spdList.SetText(16, i, GetByOne(sField16, sField16))
        Call spdList.SetText(17, i, GetByOne(sField17, sField17))
        Call spdList.SetText(19, i, GetByOne(sField19, sField19))
    Next
    
End Sub

Private Sub Get��ü��()
    Dim C��ü�� As DCR0101
    Dim i%, j%
    Dim sField01$, sField02$, sField03$
    
    Set C��ü�� = New DCR0101
        
    C��ü��.Get_��ü
    
    j = C��ü��.CurItemCnt
    
    If j = 0 Then
        ViewMsg ("��ü�ڵ� �����ڷᰡ �����ϴ�.")
        Set C��ü�� = Nothing
        Exit Sub
    End If
        
    sField01 = C��ü��.Tot��ü�ڵ�
    sField02 = C��ü��.Tot��ü���
    sField03 = C��ü��.Tot��ü����
    
    ��üCnt = j
        
    ReDim ��ü(1 To j) As ��ü��
    
    For i = 1 To j
        ��ü(i).��ü�ڵ� = GetByOne(sField01, sField01)
        ��ü(i).��ü��� = GetByOne(sField02, sField02)
        ��ü(i).��ü���� = GetByOne(sField03, sField03)
    Next

    Set C��ü�� = Nothing
    
End Sub

Private Sub ClearData(spdChk As Integer)
 
    If spdChk = 0 Then spdList.MaxRows = 0
    
    spdRst.MaxRows = 0
    spdCmt.MaxRows = 0
    
    lbl�۾����� = ""
    lblSlip = ""
    lbl�۾���ȣ = ""
    lbl��Ϲ�ȣ = ""
    lbl�̸� = ""
    lbl���� = ""
    lbl���� = ""
    lbl����� = ""
    lbl���� = ""
    lbl�������� = ""
    lbl���ޱ��� = ""
    lblƯ������ = ""
    lbl�ǻ� = ""
    txt�Ұ� = ""
    lbl�ð� = ""
    lbl��ü�� = ""
    lbl����ġ = ""
    lbl���� = ""
    
End Sub
Private Sub GetResult()
    On Error GoTo ErrHandler
    
    Dim C��� As DCR0101
    Dim i%, j%, j1%, j2%, j3%, SpdRow As Long
    Dim sField01$, sField02$, sField03$, sField04$, sField05$
    Dim sField06$, sField07$, sField08$, sField09$, sField10$
    Dim sField11$, sField12$, sField13$, sField14$, sField15$
    Dim sField16$, sField17$, sField18$, sField19$, sField20$
    Dim sField21$, sField22$, sField23$, sField24$, sField25$
    Dim sField26$, sField27$, sField28$, sField29$, sField30$
    Dim sField31$, sField32$
    Dim spdChk1, spdChk2, spdChk3, spdChk4, spdChk5, spdChk6, vSlipCd
    Dim sColor$
    
    Set C��� = New DCR0101
    
    Call spdList.GetText(1, iCurRow, spdChk1)
    Call spdList.GetText(2, iCurRow, spdChk2)
    Call spdList.GetText(3, iCurRow, spdChk3)
    Call spdList.GetText(15, iCurRow, spdChk4)
    
    If lbl���� = "��" Then
        spdChk5 = "2"
    Else
        spdChk5 = "1"
    End If
    
    C���.Get_��� spdChk1, Right(spdChk2, 2), spdChk3, spdChk4, spdChk5, Left(spdChk2, 1)
    
    j1 = C���.CurItemCnt
    j2 = C���.CurItemCnt2
    j3 = C���.CurItemCnt3
    
    If j1 = 0 Then
        ViewMsg ("����� �������� �ʽ��ϴ�.")
        Set C��� = Nothing
        Exit Sub
    End If
    
    sField01 = C���.Tot�˻��ڵ�
    sField02 = C���.Tot�˻���
    sField03 = C���.Tot�˻���
    sField04 = C���.Tot����ġ
    sField05 = C���.TotPANIC
    sField06 = C���.TotDELTA
    sField07 = C���.Tot�Ұ��ڵ�
    sField08 = C���.Tot�Ұ߳���
    sField09 = C���.Tot�����ڵ�
    sField10 = C���.TotREFGBN
    sField11 = C���.TotPANICGBN
    sField12 = C���.TotDELTAGBN
    sField13 = C���.TotDELTAVAL
    sField14 = C���.TotPANJUNGGBN
    sField15 = C���.TotPANICLOW
    sField16 = C���.TotPANICHIGH
    sField17 = C���.TotREFLOW
    sField18 = C���.TotREFHIGH
    sField19 = C���.TotUPPERLIMIT
    sField20 = C���.TotLOWERLIMIT
    sField21 = C���.TotGRAYUPPER
    sField22 = C���.TotGRAYLOWER
    sField23 = C���.TotOTHERFLAG
    sField24 = C���.TotREFLETTER
    sField25 = C���.TotUNIT
    sField26 = C���.Tot�˻��ڵ�F
    sField27 = C���.Tot�����ڵ�F
    sField28 = C���.TotSTATUSCD
    sField29 = C���.Tot����
    sField30 = C���.Tot�����
    sField31 = C���.TotFLAGYN
    sField32 = C���.Tot�������
    
    Set C��� = Nothing
    
    SpdRow = 0
    
    For i = 1 To j1
        SpdRow = SpdRow + 1
        spdRst.MaxRows = SpdRow
        spdChk6 = GetByOne(sField03, sField03)
        
        Call spdRst.SetText(1, SpdRow, GetByOne(sField02, sField02))
        Call spdRst.SetText(2, SpdRow, spdChk6)
        Call spdRst.SetText(4, SpdRow, GetByOne(sField04, sField04))
        Call spdRst.SetText(5, SpdRow, GetByOne(sField05, sField05))
        Call spdRst.SetText(6, SpdRow, GetByOne(sField06, sField06))
        Call spdRst.SetText(7, SpdRow, GetByOne(sField01, sField01))
        Call spdRst.SetText(8, SpdRow, GetByOne(sField09, sField09))
        Call spdRst.SetText(9, SpdRow, GetByOne(sField10, sField10))
        Call spdRst.SetText(10, SpdRow, GetByOne(sField11, sField11))
        Call spdRst.SetText(11, SpdRow, GetByOne(sField12, sField12))
        Call spdRst.SetText(12, SpdRow, GetByOne(sField13, sField13))
        Call spdRst.SetText(13, SpdRow, GetByOne(sField14, sField14))
        Call spdRst.SetText(14, SpdRow, GetByOne(sField15, sField15))
        Call spdRst.SetText(15, SpdRow, GetByOne(sField16, sField16))
        Call spdRst.SetText(16, SpdRow, GetByOne(sField17, sField17))
        Call spdRst.SetText(17, SpdRow, GetByOne(sField18, sField18))
        Call spdRst.SetText(18, SpdRow, GetByOne(sField19, sField19))
        Call spdRst.SetText(19, SpdRow, GetByOne(sField20, sField20))
        Call spdRst.SetText(20, SpdRow, GetByOne(sField21, sField21))
        Call spdRst.SetText(21, SpdRow, GetByOne(sField22, sField22))
        Call spdRst.SetText(22, SpdRow, GetByOne(sField23, sField23))
        Call spdRst.SetText(23, SpdRow, GetByOne(sField24, sField24))
        Call spdRst.SetText(24, SpdRow, GetByOne(sField25, sField25))
        Call spdRst.SetText(28, SpdRow, GetByOne(sField31, sField31))
        Call spdRst.SetText(29, SpdRow, spdChk6)
        Call spdRst.SetText(30, SpdRow, GetByOne(sField32, sField32))
        Call spdRst.SetText(31, SpdRow, spdChk6)
        
        'SUB�˻��׸��̸� ����� ó��
        Call spdRst.GetText(8, SpdRow, spdChk1)
        
        If Left(spdChk1, 2) = "NN" Then
            Call SpdForeBack(spdRst, 1, 1, i, i, RGB(0, 0, 0), �����)
        ElseIf IsNumeric(Left(spdChk1, 2)) = True Then
            If Left(spdChk1, 2) = "00" Then
                iColorCnt = iColorCnt + 1
                
                If (iColorCnt Mod 2) = 1 Then
                    sColor = ���ϴ�
                Else
                    sColor = ���ʷ�
                End If
                
                With spdRst
                    Call .SetText(2, i, "SUB �˻��׸�")
                    
                    .BlockMode = True
                    .Col = 1
                    .Col2 = .MaxCols
                    .Row = i
                    .Row2 = i
                    .Lock = True
                    .BlockMode = False
                End With
            End If
                
            Call SpdForeBack(spdRst, 1, 1, i, i, RGB(0, 0, 0), sColor)
        Else
            sColor = RGB(255, 255, 255)
        End If
        
        '���� ����
        Call spdRst.GetText(13, SpdRow, spdChk1)
        '����ġ ����
        Call spdRst.GetText(9, SpdRow, spdChk2)
        
        'REFGBN = 0 (����) , = 1 (����), = 2 (�� ~ ��), = 3 ( < ��), 4 = ( > ��)
        If spdChk2 <> "0" Then
            Call spdRst.GetText(20, SpdRow, spdChk3)
            Call spdRst.GetText(21, SpdRow, spdChk4)
            
            If spdChk2 = "2" Then
                If spdChk3 <> "" Then
                    Call spdRst.GetText(17, SpdRow, spdChk5)
                    spdChk5 = Trim(Str(Val(spdChk5) + Val(spdChk3)))
                    Call spdRst.SetText(25, SpdRow, spdChk5)
                End If
                If spdChk4 <> "" Then
                    Call spdRst.GetText(16, SpdRow, spdChk6)
                    spdChk6 = Trim(Str(Val(spdChk6) - Val(spdChk4)))
                    Call spdRst.SetText(26, SpdRow, spdChk6)
                End If
            ElseIf spdChk2 = "3" Then
                If spdChk3 <> "" Then
                    Call spdRst.GetText(18, SpdRow, spdChk5)
                    spdChk5 = Trim(Str(Val(spdChk5) + Val(spdChk3)))
                    Call spdRst.SetText(25, SpdRow, spdChk5)
                End If
            ElseIf spdChk2 = "4" Then
                If spdChk4 <> "" Then
                    Call spdRst.GetText(19, SpdRow, spdChk6)
                    spdChk6 = Trim(Str(Val(spdChk6) - Val(spdChk4)))
                    Call spdRst.SetText(26, SpdRow, spdChk6)
                End If
            End If
        End If
    Next
    
    'FLAG �Ѹ���
    ReDim PFLAG(j2) As PATFLAG
    
    For i = 1 To j2
        PFLAG(i).�˻��ڵ� = GetByOne(sField26, sField26)
        PFLAG(i).�����ڵ� = GetByOne(sField27, sField27)
        PFLAG(i).STATUSCD = GetByOne(sField28, sField28)
        PFLAG(i).���� = GetByOne(sField29, sField29)
        PFLAG(i).����� = GetByOne(sField30, sField30)
    Next
    
    For i = 1 To j1
        Call spdRst.GetText(7, i, spdChk2)
        Call spdRst.GetText(8, i, spdChk3)
        For j = 1 To j2
            If Trim(spdChk2) & Trim(spdChk3) = PFLAG(j).�˻��ڵ� & PFLAG(j).�����ڵ� Then
                Call spdRst.GetText(27, i, spdChk4)
                If spdChk4 = "" Then
                    Call spdRst.SetText(27, i, PFLAG(j).STATUSCD)
                    Call spdRst.SetText(32, i, PFLAG(j).STATUSCD)
                    Call SpdForeBack(spdRst, 27, 27, i, i, PFLAG(j).�����, PFLAG(j).����)
                Else
                    Call spdRst.SetText(27, i, spdChk4 & " " & PFLAG(j).STATUSCD)
                    Call spdRst.SetText(32, i, spdChk4 & " " & PFLAG(j).STATUSCD)
                    Call SpdForeBack(spdRst, 27, 27, i, i, RGB(255, 255, 255), RGB(0, 0, 0))
                End If
            End If
        Next
    Next
    
    '�Ұ߻Ѹ���
    SpdRow = 0
    For i = 1 To j3
        SpdRow = SpdRow + 1
        spdCmt.MaxRows = SpdRow
        
        Call spdCmt.SetText(1, SpdRow, GetByOne(sField07, sField07))
        Call spdCmt.GetText(1, SpdRow, spdChk1)
        
        Call spdList.GetText(2, iCurRow, vSlipCd)
        
        If spdChk1 = "" Then
            Call spdCmt.SetText(3, SpdRow, GetByOne(sField08, sField08))
        Else
            For j = 1 To �Ұ�Cnt
                If �Ұ�(j).�Ұ��ڵ� = Left$(CStr(vSlipCd), 1) & spdChk1 Then
                    '�������� �Ѿ������ �׳� GetByOne
                    Call GetByOne(sField08, sField08)
                    Call spdCmt.SetText(3, SpdRow, �Ұ�(j).�Ұ߳���)
                    Exit For
                End If
            Next
        End If
        Call spdCmt.GetText(3, SpdRow, spdChk2)
        'Call spdCmt.SetText(5, SpdRow, spdChk1 & spdChk2)
        Call spdList.SetText(18, iCurRow, CStr(j3))
    Next
    
    If spdCmt.MaxRows > 0 Then
        spdCmt.BlockMode = True
        spdCmt.Col = 1
        spdCmt.Col2 = 3
        spdCmt.Row = 1
        spdCmt.Row2 = spdCmt.MaxRows
        spdCmt.Lock = True
        spdCmt.BlockMode = False
    End If
    
    If spdRst.MaxRows > 0 Then
        spdRst.Col = 2
        spdRst.Row = 1
        spdRst.Action = SS_ACTION_ACTIVE_CELL
        spdRst.SetFocus
    End If
    
    Exit Sub
    
ErrHandler:
End Sub

Private Sub SpdBack()
    
    With spdList
        .Row = 1
        .Row2 = .MaxRows
        .Col = 1
        .Col2 = .MaxCols
        .BlockMode = True
        .BackColor = &HDFFFDF
        .BlockMode = False
    End With
    
End Sub

Private Sub Get�Ұ�()
    Dim C�Ұ� As DCR0101
    Dim i%, j%
    Dim sField01$, sField02$, sField03$
    
    Set C�Ұ� = New DCR0101
        
    C�Ұ�.Get_�Ұ�
    
    j = C�Ұ�.CurItemCnt
    
    If j = 0 Then
        Set C�Ұ� = Nothing
        Exit Sub
    End If
        
    sField01 = C�Ұ�.TotPARTCD
    sField02 = C�Ұ�.Tot�Ұ��ڵ�
    sField03 = C�Ұ�.Tot�Ұ߳���
    
    �Ұ�Cnt = j
        
    ReDim �Ұ�(1 To j) As �Ұ���
    
    For i = 1 To j
        �Ұ�(i).�Ұ��ڵ� = GetByOne(sField01, sField01) & GetByOne(sField02, sField02)
        �Ұ�(i).�Ұ߳��� = GetByOne(sField03, sField03)
    Next

    Set C�Ұ� = Nothing
    
End Sub

Private Sub GetFLAG()
    Dim CFLAG As DCR0101
    Dim i%, j%, k%, l%
    Dim sField01$, sField02$, sField03$, sField04$, sField05$
    Dim sField06$, sField07$, sField08$, sField09$, sField10$
    Dim FLAGKEY$
    
    Set CFLAG = New DCR0101
        
    CFLAG.Get_FLAG
    
    j = CFLAG.CurItemCnt
    
    If j = 0 Then
        Set CFLAG = Nothing
        Exit Sub
    End If
        
    sField01 = CFLAG.TotSTATUSCD
    sField02 = CFLAG.TotPARTCD
    sField03 = CFLAG.TotPARTGBN
    sField04 = CFLAG.Tot��ü�ڵ�
    sField05 = CFLAG.Tot�˻��ڵ�
    sField06 = CFLAG.Tot�����ڵ�
    sField07 = CFLAG.Tot�̻����
    sField08 = CFLAG.Tot�������
    sField09 = CFLAG.Tot����
    sField10 = CFLAG.Tot�����
    
    FLAGCnt = j
        
    ReDim FLAG(1 To j) As FLAGDETAIL
    ReDim FLAGCHK(1 To j) As FLAGMAIN
    
    k = 0
    FLAGKEY = ""
    
    For i = 1 To j
        FLAGCHK(i).�����ڵ� = GetByOne(sField02, sField02) & _
                              GetByOne(sField03, sField03) & _
                              GetByOne(sField04, sField04) & _
                              GetByOne(sField05, sField05) & _
                              GetByOne(sField06, sField06)
        If FLAGKEY <> FLAGCHK(i).�����ڵ� Then
            If k <> 0 Then FLAGCHK(k).FLAG���� = l
            k = k + 1
            FLAGKEY = FLAGCHK(i).�����ڵ�
            FLAGCHK(k).�����ڵ� = FLAGCHK(i).�����ڵ�
            l = 0
        End If
        l = l + 1
        FLAGCHK(k).FLAG(l).STATUSCD = GetByOne(sField01, sField01)
        FLAGCHK(k).FLAG(l).�̻���� = GetByOne(sField07, sField07)
        FLAGCHK(k).FLAG(l).������� = GetByOne(sField08, sField08)
        FLAGCHK(k).FLAG(l).���� = GetByOne(sField09, sField09)
        FLAGCHK(k).FLAG(l).����� = GetByOne(sField10, sField10)
    Next
    
    FLAGCHK(k).FLAG���� = l
    FLAGCnt = k
    
    Set CFLAG = Nothing
    
End Sub

Private Sub RstJudge(ByVal CurRst As String)
    Dim i%, j%, iFlagOK%
    Dim vRefGbn, vPanjungGbn, vRefL, vRefH, vRefLetter, vOtherFlag
    Dim vPanicYN, vPanicL, vPanicH
    Dim vDeltaYN, vPreRstVal, vDeltaVal
    Dim vFlagYN, vSlipCd, vSpecimenCd, vTestItemSeq, vSubmCd
    Dim vRefFlag, vPanicFlag, vDeltaFlag, vFlag

    With spdRst
        
        'Reference chk
        vRefFlag = ""
        If CurRst <> "" Then
            Call .GetText(9, .ActiveRow, vRefGbn)
            Call .GetText(13, .ActiveRow, vPanjungGbn)
            If vRefGbn <> "0" And vPanjungGbn <> "0" Then
                If vRefGbn = "1" Then '����
                    Call .GetText(22, .ActiveRow, vOtherFlag)
                    Call .GetText(23, .ActiveRow, vRefLetter)
                    If CurRst <> vRefLetter Then
                        vRefFlag = vOtherFlag
                    End If
                ElseIf vRefGbn = "2" Then '����(LOW,HIGH)
                    Call .GetText(26, .ActiveRow, vRefL)
                    Call .GetText(25, .ActiveRow, vRefH)
                    
                    If vRefL <> "" And (Val(CurRst) < Val(vRefL)) Then
                        If vPanjungGbn = "1" Then
                            vRefFlag = "L"
                        ElseIf vPanjungGbn = "2" Then
                            vRefFlag = "Pos"
                        ElseIf vPanjungGbn = "3" Then
                            Call .GetText(22, .ActiveRow, vOtherFlag)
                            vRefFlag = vOtherFlag
                        End If
                    End If
                    If vRefH <> "" And (Val(CurRst) > Val(vRefH)) Then
                        If vPanjungGbn = "1" Then
                            vRefFlag = "H"
                        ElseIf vPanjungGbn = "2" Then
                            vRefFlag = "Pos"
                        ElseIf vPanjungGbn = "3" Then
                            Call .GetText(22, .ActiveRow, vOtherFlag)
                            vRefFlag = vOtherFlag
                        End If
                    End If
                ElseIf vRefGbn = "3" Then '����(����)
                    Call .GetText(25, .ActiveRow, vRefH)
                    If vRefH <> "" And (Val(CurRst) > Val(vRefH) Or Val(CurRst) = Val(vRefH)) Then
                        If vPanjungGbn = "1" Then
                            vRefFlag = "H"
                        ElseIf vPanjungGbn = "2" Then
                            vRefFlag = "Pos"
                        ElseIf vPanjungGbn = "3" Then
                            Call .GetText(22, .ActiveRow, vOtherFlag)
                            vRefFlag = vOtherFlag
                        End If
                    End If
                ElseIf vRefGbn = "4" Then '����(����)
                    Call .GetText(26, .ActiveRow, vRefL)
                    If vRefL <> "" And (Val(CurRst) < Val(vRefL) Or Val(CurRst) = Val(vRefL)) Then
                        If vPanjungGbn = "1" Then
                            vRefFlag = "L"
                        ElseIf vPanjungGbn = "2" Then
                            vRefFlag = "Pos"
                        ElseIf vPanjungGbn = "3" Then
                            Call .GetText(22, .ActiveRow, vOtherFlag)
                            vRefFlag = vOtherFlag
                        End If
                    End If
                End If
                
                If vRefGbn = "2" Or vRefGbn = "3" Or vRefGbn = "4" Then
                    If IsNumeric(CurRst) = False Then
                        If vPanjungGbn = "1" Then
                            vRefFlag = "H"
                        ElseIf vPanjungGbn = "2" Then
                            vRefFlag = "Pos"
                        ElseIf vPanjungGbn = "3" Then
                            vRefFlag = vOtherFlag
                        End If
                    End If
                End If
            End If
        End If
        'ȭ�鿡 �ݿ�
        Call .SetText(4, .ActiveRow, vRefFlag)
        
        'PANIC check
        vPanicFlag = ""
        If CurRst <> "" Then
            Call .GetText(10, .ActiveRow, vPanicYN)
            If vPanicYN = "1" Then
                Call .GetText(14, .ActiveRow, vPanicL)
                Call .GetText(15, .ActiveRow, vPanicH)
                If vPanicL <> "" And (Val(CurRst) < Val(vPanicL)) Then
                    vPanicFlag = "P"
                End If
                If vPanicH <> "" And (Val(CurRst) > Val(vPanicH)) Then
                    vPanicFlag = "P"
                End If
            End If
        End If
        Call .SetText(5, .ActiveRow, vPanicFlag)
        
        'DELTA chk
        vDeltaFlag = ""
        If CurRst <> "" Then
            Call .GetText(11, .ActiveRow, vDeltaYN)
            If vDeltaYN <> "0" Then
                Call .GetText(30, .ActiveRow, vPreRstVal)
                If vPreRstVal <> "" Then
                    Call .GetText(12, .ActiveRow, vDeltaVal)
                    If vDeltaYN = "1" Then
                        If vDeltaVal <> "" And (Abs(Val(CurRst) - Val(vPreRstVal)) > Val(vDeltaVal) Or Abs(Val(CurRst) - Val(vPreRstVal)) = Val(vDeltaVal)) Then
                            vDeltaFlag = "D"
                        End If
                    ElseIf vDeltaYN = "2" And CurRst <> "0" Then
                        If vDeltaVal <> "" And (((Abs(Val(CurRst) - Val(vPreRstVal)) / Val(CurRst)) * 100#) > Val(vDeltaVal) Or ((Abs(Val(CurRst) - Val(vPreRstVal)) / Val(CurRst)) * 100#) = Val(vDeltaVal)) Then
                            vDeltaFlag = "D"
                        End If
                    End If
                End If
            End If
        End If
        Call .SetText(6, .ActiveRow, vDeltaFlag)
        
        'FLAG chk
        iFlagOK = 0
        Call .SetText(27, .ActiveRow, "")
        If CurRst <> "" Then
            Call spdList.GetText(2, iCurRow, vSlipCd)
            Call spdList.GetText(15, iCurRow, vSpecimenCd)
            Call .GetText(7, .ActiveRow, vTestItemSeq)
            Call .GetText(8, .ActiveRow, vSubmCd)
            For i = 1 To FLAGCnt
                If FLAGCHK(i).�����ڵ� = vSlipCd & vSpecimenCd & vTestItemSeq & vSubmCd Then
                    For j = 1 To FLAGCHK(i).FLAG����
                        Call .GetText(27, .ActiveRow, vFlag)
                        If FLAGCHK(i).FLAG(j).������� = "�̻�" Then
                            If Val(CurRst) > Val(FLAGCHK(i).FLAG(j).�̻����) Or Val(CurRst) = Val(FLAGCHK(i).FLAG(j).�̻����) Then
                                iFlagOK = 1
                                If vFlag = "" Then
                                    Call .SetText(27, .ActiveRow, FLAGCHK(i).FLAG(j).STATUSCD)
                                    Call SpdForeBack(spdRst, 27, 27, .ActiveRow, .ActiveRow, FLAGCHK(i).FLAG(j).�����, FLAGCHK(i).FLAG(j).����)
                                Else
                                    'Flag��  �� �̻��� ���
                                    Call .SetText(27, .ActiveRow, vFlag & " " & FLAGCHK(i).FLAG(j).STATUSCD)
                                    Call SpdForeBack(spdRst, 27, 27, .ActiveRow, .ActiveRow, RGB(255, 255, 255), RGB(0, 0, 0))
                                End If
                            End If
                        Else
                            If Val(CurRst) < Val(FLAGCHK(i).FLAG(j).�̻����) Or Val(CurRst) = Val(FLAGCHK(i).FLAG(j).�̻����) Then
                                iFlagOK = 1
                                If vFlag = "" Then
                                    Call .SetText(27, .ActiveRow, FLAGCHK(i).FLAG(j).STATUSCD)
                                    Call SpdForeBack(spdRst, 27, 27, .ActiveRow, .ActiveRow, FLAGCHK(i).FLAG(j).�����, FLAGCHK(i).FLAG(j).����)
                                Else
                                    'Flag��  �� �̻��� ���
                                    Call .SetText(27, .ActiveRow, vFlag & " " & FLAGCHK(i).FLAG(j).STATUSCD)
                                    Call SpdForeBack(spdRst, 27, 27, .ActiveRow, .ActiveRow, RGB(255, 255, 255), RGB(0, 0, 0))
                                End If
                            End If
                        End If
                    Next
                End If
            Next
        End If
        
        If iFlagOK = 0 Then Call SpdForeBack(spdRst, 27, 27, .ActiveRow, .ActiveRow, RGB(0, 0, 0), RGB(255, 255, 255))
        
        '����� ����� �����ϱ�
        Call .SetText(29, .ActiveRow, CurRst)
    End With
End Sub


Private Sub SpdInit()
    'SpreadBackColor Option
    iSpdBackColorOption = 2
    
    With spdList
        .BlockMode = True
        .Col = -1
        .Col2 = -1
        .Row = -1
        .Row2 = -1
        .BackColorStyle = BackColorStyleUnderGrid
        .BackColor = SpdBackcolor(iSpdBackColorOption)   'GBR
        .EditModePermanent = True
        .Protect = True
        .NoBeep = True
        .BlockMode = False
        
        .BlockMode = True
        .Col = 1
        .Col2 = 3
        .Row = -1
        .Row2 = -1
        .Lock = True
        .BlockMode = False
        
        .MaxRows = 0
    End With
    
    With spdRst
        .BlockMode = True
        .Col = -1
        .Col2 = -1
        .Row = -1
        .Row2 = -1
        .BackColorStyle = BackColorStyleUnderGrid
        '.BackColor = SpdBackcolor(iSpdBackColorOption)   'GBR
        .EditModePermanent = True
        .Protect = True
        .NoBeep = True
        .BlockMode = False
                
        .MaxRows = 0
    End With
End Sub

Private Sub RefShow(ByVal lnRow As Long)
    Dim spdChk1, spdChk2, spdChk3, spdChk4, spdChk5, spdChk6, spdChk7
    
    With spdRst
        Call .GetText(9, lnRow, spdChk2)
        lbl����ġ = "": lbl���� = ""
        If spdChk2 <> "0" Then
            If spdChk2 = "1" Then
                Call .GetText(23, lnRow, spdChk3)
                lbl����ġ = spdChk3
            ElseIf spdChk2 = "2" Then
                Call .GetText(16, lnRow, spdChk3)
                Call .GetText(17, lnRow, spdChk4)
                Call .GetText(24, lnRow, spdChk5)
                Call .GetText(20, lnRow, spdChk6)
                Call .GetText(21, lnRow, spdChk7)
                lbl����ġ = spdChk3
                If spdChk7 <> "0" Then lbl����ġ = lbl����ġ & "(-" & spdChk7 & ") "
                lbl����ġ = lbl����ġ & " ~ " & spdChk4
                If spdChk6 <> "0" Then lbl����ġ = lbl����ġ & "(+" & spdChk6 & ") "
                lbl���� = spdChk5
            ElseIf spdChk2 = "3" Then
                Call .GetText(18, lnRow, spdChk3)
                Call .GetText(24, lnRow, spdChk4)
                Call .GetText(20, lnRow, spdChk5)
                lbl����ġ = spdChk3
                If spdChk5 <> "0" Then lbl����ġ = lbl����ġ & "(+" & spdChk5 & ") "
                lbl����ġ = "<  " & lbl����ġ
                lbl���� = spdChk4
            ElseIf spdChk2 = "4" Then
                Call .GetText(19, lnRow, spdChk3)
                Call .GetText(24, lnRow, spdChk4)
                Call .GetText(21, lnRow, spdChk5)
                If spdChk5 <> "0" Then lbl����ġ = lbl����ġ & "(-" & spdChk5 & ") "
                lbl����ġ = ">  " & lbl����ġ
                lbl���� = spdChk4
            End If
        End If
    End With
End Sub

Private Sub spdBlock(ByVal lnCol1 As Long, ByVal lnCol2 As Long, ByVal lnRow1 As Long, ByVal lnRow2 As Long, ByVal iMode As Integer, ByVal iAct As Integer)
    With spdRst
        .Col = lnCol1: .Col2 = lnCol2
        .Row = lnRow1: .Row2 = lnRow2
        
        .BlockMode = True
        If iMode = 0 Then
            If iAct = 0 Then
                '.FontBold = False
            Else
                '.FontBold = True
            End If
        ElseIf iMode = 1 Then
            If iAct = 0 Then
                .Lock = False
            Else
                .Lock = True
            End If
        End If
        .BlockMode = False
        
        If iMode = 1 And iAct = 1 Then
            .BlockMode = True
            .Col = 1
            .Col2 = 1
            .Row = lnRow1
            .Row2 = lnRow2
            .ShadowText = True
            .BlockMode = False
        End If
    End With
End Sub

Private Sub cmdAddCmt_Click()
    Dim vTmp
    
    Call spdCmt.GetText(2, spdCmt.MaxRows, vTmp)
    
    If vTmp = "" And IsNull(vTmp) = True Then
        ViewMsg "���� ���� COMMENT ������ �ۼ����� �ʾҽ��ϴ�..."
    Else
        spdCmt.MaxRows = spdCmt.MaxRows + 1
        Call SpdForeBack(spdCmt, -1, -1, spdCmt.MaxRows, spdCmt.MaxRows, RGB(0, 0, 0), ���ʷ�)
        
        If spdCmt.MaxRows > 3 Then
            spdCmt.TopRow = spdCmt.MaxRows - 2
        End If
    End If
End Sub

Private Sub cmdButtonSlip_Click()
    Dim i%
    Dim j%
    Dim CPart As DCB0101
    Dim sTot01$
    Dim sTot02$
    Dim sTot03$
    
    Set CPart = New DCB0101
    
    CPart.Get_PART
    
    j = CPart.CurItemCnt
    
    iHlpClick = 1
    
    Erase gCodeHlpTable '�迭 �ʱ�ȭ
    
    ReDim gCodeHlpTable(j) As CodeTBL
    
    With CPart
        sTot01 = .TotField01
        sTot02 = .TotField02
        sTot03 = .TotField03
    End With
    
    Set CPart = Nothing
    
    For i = 1 To j
        gCodeHlpTable(i).sSeq = Format$(i, "00000")
        gCodeHlpTable(i).sCode = GetByOne(sTot01, sTot01) & GetByOne(sTot02, sTot02)
        gCodeHlpTable(i).sCodeNm = GetByOne(sTot03, sTot03)
    Next
    
    giCodeHlpCnt = j
    
    hWndCd = txtSlip.hwnd
    
    FSR0101.Left = 2700
    FSR0101.Top = 1400
    
    Load FSR0101
    FSR0101.Show vbModal
    
    iHlpClick = 0
    txt�۾���ȣ.SetFocus
End Sub

Private Sub cmdReg_Click()
    Dim i%, iRstCnt%, iFlagCnt%, iAbnormalCnt%, iCmtCnt%
    Dim C���� As DCR0101
    Dim vLabdate, vPartGbn, vLabSeq, vCmt, vRst, vLastRst
    Dim vCmtCd, vCmtCnt, vTotalCmt, vSpc, vFlag, vLastFlag
    Dim vExamCd, vSubCd, vRstData, vRefMark, vPanicMark, vDeltaMark, vFlagYN
    Dim sLabInfo$, sRstDate$, sRstTime$, sJubsu$, sFlag$
    Dim sCmt$, sCmtFlag$, sRst$, sSpc$, sAllData$, sAbnormal$
    Dim RtnCd$
    Dim vRstVal, vRstYN
    
' ��������
    Call spdList.GetText(1, iCurRow, vLabdate)
    Call spdList.GetText(2, iCurRow, vPartGbn)
    Call spdList.GetText(3, iCurRow, vLabSeq)
    sLabInfo = vLabdate & vPartGbn & vLabSeq
    
' �������(X_RESULT) & �̻���(X_ABNORMAL) & ����Ϸ��Ͻ�(X_JUBSE) & �Ұ�(COMMENT) & FLAG
    sJubsu = ""
    sFlag = ""
    sRst = ""
    sAbnormal = ""
    sAllData = "Y"
    iRstCnt = 0
    iFlagCnt = 0
    iAbnormalCnt = 0
    iCmtCnt = 0
    
    '�����ȣ(�ӻ�Ұ�)����Ȯ��
    Call spdList.GetText(14, iCurRow, vCmt)
    If vCmt <> Trim(txt�Ұ�) Then
        If Trim(txt�Ұ�) = "" Then
            sJubsu = Chr(13) & "|"
        Else
            sJubsu = Trim(txt�Ұ�) & "|"
        End If
    Else
        sJubsu = "|"
    End If
    
    '����Ϸ����� �ƴ��� üũ
    Call spdList.GetText(19, iCurRow, vRstYN)
    
    '�������Ȯ��
    For i = 1 To spdRst.MaxRows
        Call spdRst.GetText(8, i, vSubCd)

        If Left(Trim(vSubCd), 2) <> "00" Then
            '���
            Call spdRst.GetText(2, i, vRst)
            Call spdRst.GetText(31, i, vLastRst)
            
            If (vRstYN = "0" Or vRstYN = "") Then
                If Trim(vRst) = "" Then sAllData = "N"
            End If
            
            If vRst <> vLastRst Then
                iRstCnt = iRstCnt + 1
                
                If Trim(vRst) = "" Then sAllData = "N"
                
                Call spdRst.GetText(7, i, vExamCd)
                Call spdRst.GetText(4, i, vRefMark)
                Call spdRst.GetText(5, i, vPanicMark)
                Call spdRst.GetText(6, i, vDeltaMark)
                
                sRst = sRst & vExamCd & "|" & vSubCd & "|" & vRst & "|" & _
                       vRefMark & "|" & vPanicMark & "|" & vDeltaMark & "|"
                
                sAbnormal = sAbnormal & vExamCd & "|" & vSubCd & "|" & vRst & "|" & _
                   vRefMark & "|" & vPanicMark & "|" & vDeltaMark & "|"
                
                iAbnormalCnt = iAbnormalCnt + 1
                
                'FLAG üũ�� ����
                Call spdRst.GetText(27, i, vFlag)
                Call spdRst.GetText(32, i, vLastFlag)
                
                sFlag = sFlag & vExamCd & "|" & vSubCd & "|" & vFlag & " |"
                
                iFlagCnt = iFlagCnt + 1
            End If
        Else
        End If
    Next
    
    If iRstCnt <> 0 Then
        sRst = Str(iRstCnt) & "|" & sRst
        
        '����� �� ���� ó��(����Ϸ��Ͻ�)
        If (vRstYN = "0" Or vRstYN = "") And sAllData = "Y" Then
            sJubsu = sJubsu & "Y|"
        ElseIf vRstYN = "1" And sAllData = "N" Then
            sJubsu = sJubsu & "N|"
        Else
            sJubsu = sJubsu & "|"
        End If
        
        Call spdList.GetText(15, iCurRow, vSpc)
        sSpc = vSpc
        sRstDate = Format(Now, "YYYYMMDD")
        sRstTime = Format(Now, "HHMMSS")
    Else
        sJubsu = sJubsu & "|"
    End If

'ABNORMAL
    If iAbnormalCnt <> 0 Then
        sAbnormal = CStr(iAbnormalCnt) & "|" & sAbnormal
    Else
    End If
    
' ȯ�� Comment(COMMENT)
    sCmt = ""
    'Comment����Ȯ��
    sCmtFlag = "N"
    
    For i = 1 To spdCmt.MaxRows
        Call spdCmt.GetText(1, i, vCmtCd)
        Call spdCmt.GetText(3, i, vCmt)
        'Call spdCmt.GetText(5, i, vTotalCmt)
        spdCmt.Row = i
        
        If spdCmt.BackColor = ���ʷ� Then
            If vCmtCd = "" Then
                sCmt = sCmt & "" & "|" & vCmt & "|"
            Else
                sCmt = sCmt & vCmtCd & "|" & "" & "|"
            End If
            
            iCmtCnt = iCmtCnt + 1
            sCmtFlag = "Y"
        End If
    Next
    
    Call spdList.GetText(18, iCurRow, vCmtCnt)
    
    If spdCmt.MaxRows <> Val(vCmtCnt) Then
        sCmtFlag = "Y"
        If spdCmt.MaxRows = 0 Then sCmt = "NULL"
    End If
    
    'RERUN�� ��� X_JUBSU�� RERUNYN�� 1�� UPDATE ����
    If iReRunYN = 1 Then
        sJubsu = sJubsu & "R|"
    Else
        sJubsu = sJubsu & "|"
    End If
    
    '����� ������ ������ ����������
    If Len(sJubsu) = 3 And iRstCnt = 0 And sCmtFlag = "N" Then
        ViewMsg "������ ������ �����ϴ�..."
        Exit Sub
    End If
    
    'Transaction ó��-----------------
    Set C���� = New DCR0101
    
    RtnCd = C����.Tran_Result(sLabInfo, sJubsu, sCmt, iCmtCnt, sRst, sRstDate, sRstTime, sSpc, sFlag, iFlagCnt, sAbnormal, CStr(iReRunYN))
    
    Set C���� = Nothing
    '---------------------------------

    '�������ο����� ��� �ݿ�
    If RtnCd = "SUCCESS" Then
        ViewMsg "���������� ����Ǿ����ϴ�..."
        
        '�����ȣ ������ ȭ�鿡 �ݿ�
        Call spdList.SetText(14, iCurRow, Trim$(txt�Ұ�) & "")
        
        '����Ϸ� ������ ȭ�鿡 �ݿ�
        If (vRstYN = "0" Or vRstYN = "") And sAllData = "Y" Then
            Call spdList.SetText(16, iCurRow, sRstDate)
            Call spdList.SetText(17, iCurRow, sRstTime)
            Call spdList.SetText(19, iCurRow, "1")
            lbl�ð� = Left(sRstDate, 4) & "-" & Mid(sRstDate, 5, 2) & "-" & Right(sRstDate, 2) & " " & _
                      Left(sRstTime, 2) & ":" & Mid(sRstTime, 3, 2) & ":" & Right(sRstTime, 2)
        ElseIf vRstYN = "1" And sAllData = "N" Then
            Call spdList.SetText(16, iCurRow, "")
            Call spdList.SetText(17, iCurRow, "")
            Call spdList.SetText(19, iCurRow, "0")
            
            lbl�ð� = ""
        End If
        
        For i = 1 To spdRst.MaxRows
            Call spdRst.GetText(2, i, vRstVal)
            Call spdRst.SetText(29, i, vRstVal)
            Call spdRst.SetText(31, i, vRstVal)
        Next
        
        If iReRunYN = 1 Then
            'COMMENT�� �����Ǿ����Ƿ� ������ ȭ�鿡 �ݿ�
            Call spdList.SetText(18, iCurRow, "0")
        Else
            For i = 1 To spdCmt.MaxRows
                Call SpdForeBack(spdCmt, -1, -1, i, i, RGB(0, 0, 0), RGB(255, 255, 255))
            Next
                        
            spdCmt.BlockMode = True
            spdCmt.Col = 1
            spdCmt.Col2 = 3
            spdCmt.Row = 1
            spdCmt.Row2 = spdCmt.MaxRows
            spdCmt.Lock = True
            spdCmt.BlockMode = False
            
            Call spdList.SetText(18, iCurRow, CStr(spdCmt.MaxRows) & "")
        End If
    Else
        ViewMsg "���忡 �����Ͽ����ϴ�..."
    End If
End Sub

Private Sub cmdReRun_Click()
    Dim i%
    
    If MsgBox("��� ó���� �Ͻðڽ��ϱ�?" & vbCrLf & _
                "��� ó���� �ϸ� �ش� �۾���ȣ�� ����� ��� �����˴ϴ�.", vbYesNo, _
                "��� ó�� Ȯ��") = vbYes Then
        
        iReRunYN = 1
        
        For i = 1 To spdRst.MaxRows
            'RESULT CLEAR
            Call spdRst.SetText(2, i, "")
            'REF CHECK CLEAR
            Call spdRst.SetText(4, i, "")
            'PANIC CHECK CLEAR
            Call spdRst.SetText(5, i, "")
            'DELTA CHECK CLEAR
            Call spdRst.SetText(6, i, "")
            
            'FLAG CLEAR
            Call spdRst.SetText(27, i, "")
            Call SpdForeBack(spdRst, 27, 27, i, i, RGB(0, 0, 0), RGB(255, 255, 255))
        Next
        
        spdCmt.MaxRows = 0
        
        lbl����ġ = ""
        lbl���� = ""
        
        Call cmdReg_Click
        
        iReRunYN = 0
    End If
End Sub

Private Sub cmdSlipPrint_Click()
    Dim CPrint As DCO0101
    Dim iRetVal As Integer
    
    Set CPrint = New DCO0101
    
    If lbl�۾����� = "" Then
        ViewMsg "����� ����� ����� �����ϴ�..."
    Else
        iRetVal = CPrint.Get_Print_Info("00", lblSlip)
        iRetVal = CPrint.Print_LabResult(lbl�۾�����, lblSlip, lbl�۾���ȣ)
        
        If iRetVal = 1 Then
            ViewMsg lbl�۾����� & "-" & lblSlip & "-" & lbl�۾���ȣ & "�� ����� ����Ͽ����ϴ�..."
        End If
    End If
    
    Set CPrint = Nothing
End Sub

Private Sub dtpLabDate_Change()
    iDtpChange = 1
    txt�۾����� = Trim$(Format$(dtpLabDate.Value, "YYYYMMDD"))
End Sub

Private Sub dtpLabDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If optGbn(0).Value = True Then
            optGbn(0).SetFocus
        ElseIf optGbn(1).Value = True Then
            optGbn(1).SetFocus
        ElseIf optGbn(2).Value = True Then
            optGbn(2).SetFocus
        End If
    End If
End Sub

Private Sub dtpLabDate_Validate(Cancel As Boolean)
    txt�۾����� = Trim(Format(dtpLabDate.Value, "YYYYMMDD"))
End Sub

Private Sub Form_Activate()
    If iHlpClick = 0 Then
        optGbn(0).Value = True
    ElseIf iHlpClick = 1 Then
        iHlpClick = 0
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1
            Call cmdButtonSlip_Click
        Case vbKeyF3
            If optGbn(0).Value = True Then
                txt�۾���ȣ.SetFocus
            ElseIf optGbn(1).Value = True Then
                txt��Ϲ�ȣ.SetFocus
            ElseIf optGbn(2).Value = True Then
                If opt��(0).Value = True Then
                    opt��(0).SetFocus
                Else
                    opt��(1).SetFocus
                End If
            End If
        Case vbKeyF8
            Call spdRst_KeyDown(13, 0)
            Call cmdReg_Click
        Case vbKeyEscape
            Call cmdExit_Click
    End Select
End Sub

Private Sub Form_Load()
    Dim sUseYN$
    Dim bRetVal As Boolean
    
'''    Me.Left = 0
'''    Me.Top = 0
'''    Me.Width = 11920
'''    Me.Height = 7950
    
    iHlpClick = 0
    iDtpChange = 0
    iColorCnt = 0
    iReRunYN = 0
    
    Me.KeyPreview = True
    
    dtpLabDate.Value = Format(Now, "yyyy-mm-dd")
    
    Get�����
    Get��ü��
    Get�Ұ�
    GetFLAG
    
'<----------------- �޴� �� �����ڷ��� OTHER�� ��� ���θ� Registry�� ���� �о� �Ǵ� ----------->
    
    sUseYN = GetKeyValue(HKEY_CURRENT_USER, _
                    "Software\SemiLIS\Program Config\Menu.Setting\Others.Visible", "Check YN")
    
    If sUseYN = "Y" Then
        fraOthers.Visible = True
    ElseIf sUseYN = "N" Then
        fraOthers.Visible = False
    ElseIf sUseYN = "" Then     '���� ������Ʈ��Ű�� ���� �� Default �� ���
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                      "Software\SemiLIS\Program Config\Menu.Setting\Others.Visible", "Check YN", "N")

        If bRetVal = True Then
            fraOthers.Visible = False
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
        End If
    End If
'<---------------------------------------------------------------------------------------->

    Call SpdInit
    
    Call ChkFromSearchFrm
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call InitRegCurFrmTitle
    ViewMsg ""
    
End Sub

Private Sub opt��_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    opt��(Index).Value = True
    ClearData 0
    GetPatient
    
End Sub

Private Sub optGbn_Click(Index As Integer)

    pnl�۾���ȣ.Visible = False
    pnl��Ϲ�ȣ.Visible = False
    pnl��.Visible = False
    ClearData 0
    
    Select Case Index
        Case 0
            pnl�۾���ȣ.Visible = True
            txt�۾����� = Trim(Format(dtpLabDate.Value, "YYYYMMDD"))
            txtSlip = fCurUserSlipCd
            txt�۾���ȣ = ""
            txt�۾���ȣ.SetFocus
         Case 1
            pnl��Ϲ�ȣ.Visible = True
            txt��Ϲ�ȣ = ""
            txt��Ϲ�ȣ.SetFocus
        Case 2
            pnl��.Visible = True
            opt��(0).Value = True
            opt��(0).SetFocus
    End Select
     
End Sub

Private Sub optGbn_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        
        If optGbn(0).Value = True Then
            txt�۾���ȣ.SetFocus
        ElseIf optGbn(1).Value = True Then
            txt��Ϲ�ȣ.SetFocus
        ElseIf optGbn(2).Value = True Then
            opt��(0).Value = True
            opt��(0).SetFocus
        End If
    End If
End Sub

Private Sub pnlLabDate_Click()

    If pnlLabDate.Caption = "��������" Then
        pnlLabDate.Caption = "�˻�Ϸ���"
    ElseIf pnlLabDate.Caption = "�˻�Ϸ���" Then
        pnlLabDate.Caption = "��������"
    End If
    
End Sub

Private Sub cmdExit_Click()

    Unload Me
    
End Sub

Private Sub spdCmt_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    Dim vTmp, vCmtCd, vCmt
    Dim sCmt$, sLabInfo
    Dim CCmt As DCR0101
    
    vTmp = ""
    vCmtCd = ""
    vCmt = ""
    
    With spdCmt
        If Col = 2 Then
            Call Com_S_CdHlp(Row)
        ElseIf Col = 4 Then
            .Row = Row
            
            Call .GetText(3, Row, vTmp)
            
            If (vTmp = "" Or IsNull(vTmp) = True) And .BackColor = ���ʷ� Then
                .Action = SS_ACTION_DELETE_ROW
                .MaxRows = .MaxRows - 1
            Else
                If MsgBox("COMMENT ������ �ֽ��ϴ�. ������ �����Ͻðڽ��ϱ�?", vbYesNo) = vbYes Then
                    Set CCmt = New DCR0101
                    
                    Call .GetText(1, Row, vCmtCd)
                    Call .GetText(3, Row, vCmt)
                    
                    If vCmtCd = "" Then
                        sCmt = "" & "|" & CStr(vCmt) & "|"
                    Else
                        sCmt = CStr(vCmtCd) & "|" & "" & "|"
                    End If
                    
                    sLabInfo = lbl�۾����� & lblSlip & lbl�۾���ȣ
                    
                    If CCmt.Del_CmtByOne(sLabInfo, sCmt, 1) = "SUCCESS" Then
                        .Action = SS_ACTION_DELETE_ROW
                        .MaxRows = .MaxRows - 1
                        Call spdList.SetText(18, iCurRow, CStr(spdCmt.MaxRows) & "")
                    End If
                                        
                    Set CCmt = Nothing
                End If
            End If
        End If
    End With
End Sub

Private Sub spdList_Click(ByVal Col As Long, ByVal Row As Long)
    
    Dim vCol01, vCol02, vCol03, vCol04, vCol05
    Dim vCol06, vCol07, vCol08, vCol09, vCol10
    Dim vCol11, vCol12, vCol13, vCol14, vCol15
    Dim vCol16, vCol17
    Dim i%
    Dim vChk As Variant
    
    If Row = 0 Then Exit Sub
    
    Call spdList.GetText(1, Row, vChk)
    If Trim(vChk) = "" Then Exit Sub
    
    ClearData 1
    
    Call spdReverse(spdList, -1, -1, Row, Row, RGB(255, 230, 230), iSpdBackColorOption)
    
    iCurRow = Row
    
    Call spdList.GetText(1, Row, vCol01)
    Call spdList.GetText(2, Row, vCol02)
    Call spdList.GetText(3, Row, vCol03)
    Call spdList.GetText(4, Row, vCol04)
    Call spdList.GetText(5, Row, vCol05)
    Call spdList.GetText(6, Row, vCol06)
    Call spdList.GetText(7, Row, vCol07)
    Call spdList.GetText(8, Row, vCol08)
    Call spdList.GetText(9, Row, vCol09)
    Call spdList.GetText(10, Row, vCol10)
    Call spdList.GetText(11, Row, vCol11)
    Call spdList.GetText(12, Row, vCol12)
    Call spdList.GetText(13, Row, vCol13)
    Call spdList.GetText(14, Row, vCol14)
    Call spdList.GetText(15, Row, vCol15)
    Call spdList.GetText(16, Row, vCol16)
    Call spdList.GetText(17, Row, vCol17)

    lbl�۾����� = vCol01
    lblSlip = vCol02
    lbl�۾���ȣ = vCol03
    lbl��Ϲ�ȣ = vCol04
    lbl�̸� = vCol05
    lbl���� = vCol06

    If vCol07 = "1" Then
        lbl���� = "��"
    ElseIf vCol07 = "2" Then
        lbl���� = "��"
    Else
        lbl���� = "?"
    End If

    For i = 1 To ����Cnt
        If ����(i).������ڵ� = UCase$(vCol08) Then
            lbl����� = ����(i).�������
            Exit For
        End If
    Next

    lbl���� = vCol09
    
    If vCol10 = "0" Then
        lbl�������� = "�ܷ�"
    ElseIf vCol10 = "1" Then
        lbl�������� = "�Կ�"
    ElseIf vCol10 = "2" Then
        lbl�������� = "��Ź"
    End If
    
    If vCol11 = "0" Then
        lbl���ޱ��� = "N"
    ElseIf vCol11 = "1" Then
        lbl���ޱ��� = "Y"
    Else    'Zerolength�� ���
        lbl���ޱ��� = "N"
    End If
    
    If vCol12 = "0" Then
        lblƯ������ = "N"
    ElseIf vCol12 = "1" Then
        lblƯ������ = "Y"
    Else
        lblƯ������ = "N"
    End If
    
    lbl�ǻ� = vCol13
    txt�Ұ� = vCol14
    
    For i = 1 To ��üCnt
        If ��ü(i).��ü�ڵ� = vCol15 Then
            lbl��ü�� = ��ü(i).��ü���� & "(" & ��ü(i).��ü��� & ")"
            Exit For
        End If
    Next
    
    If vCol16 <> "" Then
        lbl�ð� = Left(vCol16, 4) & "-" & Mid(vCol16, 5, 2) & "-" & Right(vCol16, 2) & " "
    End If
    If vCol17 <> "" Then
        lbl�ð� = lbl�ð� & Left(vCol17, 2) & ":" & Mid(vCol17, 3, 2) & ":" & Right(vCol17, 2)
    End If
    
    GetResult '�����������
    
End Sub

Private Sub spdRst_Change(ByVal Col As Long, ByVal Row As Long)
    Dim spdChk1, spdChk2
    
    '����� �����Ǿ����� Ȯ��
    Call spdRst.GetText(2, Row, spdChk1)
    Call spdRst.GetText(29, Row, spdChk2)
    
    If Trim(spdChk1) <> Trim(spdChk2) Then
        RstJudge (Trim$(CStr(spdChk1)))
    End If
End Sub

Private Sub spdRst_Click(ByVal Col As Long, ByVal Row As Long)
    If Row = 0 Then Exit Sub
    
    If Col = 3 Then
        Call Com_E_CdHlp(Row)
    Else
        RefShow (Row)
    End If
End Sub

Private Sub spdRst_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim spdChk1, spdChk2
    
    If KeyCode = 13 Then
        KeyCode = 0
        
        Call spdRst.GetText(2, spdRst.ActiveRow, spdChk1)
        Call spdRst.GetText(29, spdRst.ActiveRow, spdChk2)
        
        If Trim(spdChk1) <> Trim(spdChk2) Then
            RstJudge (Trim$(CStr(spdChk1)))
        End If
    End If
End Sub

Private Sub spdRst_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim spdChk1, spdChk2
    
    If Col = 2 And NewCol = 2 Then
        Call spdRst.GetText(2, Row, spdChk1)
        Call spdRst.GetText(29, Row, spdChk2)
        
        If Trim(spdChk1) <> Trim(spdChk2) Then
            RstJudge (Trim$(CStr(spdChk1)))
        End If
        
        RefShow NewRow
    End If
End Sub

Private Sub txt��Ϲ�ȣ_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Len(txt��Ϲ�ȣ) > 0 Then
            ClearData 0
            GetPatient
        End If
    Else
        If Len(txt��Ϲ�ȣ) >= fDigRegNo Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txt��Ϲ�ȣ_Click()
    Call Txt_Highlight(txt��Ϲ�ȣ)
End Sub

Private Sub txt��Ϲ�ȣ_GotFocus()
    Call Txt_Highlight(txt��Ϲ�ȣ)
End Sub

Private Sub txt��Ϲ�ȣ_KeyDown(KeyCode As Integer, Shift As Integer)
'''    If KeyCode = 13 Then
'''        KeyCode = 0
'''        If Len(txt��Ϲ�ȣ) > 0 Then
'''            ClearData 0
'''            GetPatient
'''        End If
'''    End If
End Sub

Private Sub txt��Ϲ�ȣ_Validate(Cancel As Boolean)
'    Dim i%
'
'    txt��Ϲ�ȣ = UCase(txt��Ϲ�ȣ)
'
'    For i = 1 To giPartCnt
'        If gPartTable(i).sPartInit = UCase(txt��Ϲ�ȣ) Then
'            'lblPart.Caption = gPartTable(i).sPartName
'            Exit For
'        End If
'    Next
End Sub


Private Sub txt�۾���ȣ_Change()
    On Error GoTo ErrHandler

    Dim i%

    If Len(txt�۾���ȣ) = txt�۾���ȣ.MaxLength Then
        ClearData 0
        GetPatient
'        spdList.Row = 1: spdList.Col = 1
'        spdList.Action = SS_ACTION_ACTIVE_CELL
        Call spdList_Click(1, 1)
    End If
ErrHandler:

End Sub

Private Sub txt�۾���ȣ_Click()
    Call Txt_Highlight(txt�۾���ȣ)
End Sub

Private Sub txt�۾���ȣ_GotFocus()
    Call Txt_Highlight(txt�۾���ȣ)
End Sub

Private Sub txt�۾���ȣ_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        
        If Len(txt�۾���ȣ) < txt�۾���ȣ.MaxLength Then
            txt�۾���ȣ = Format$(txt�۾���ȣ, "00000")
            Call Txt_Highlight(txt�۾���ȣ)
        End If
    End If
End Sub

Private Sub txt�۾���ȣ_LostFocus()
    If Len(txt�۾���ȣ) < txt�۾���ȣ.MaxLength Then
        txt�۾���ȣ = Format$(txt�۾���ȣ, "00000")
    End If
End Sub

Private Sub txt�۾�����_Change()
    If iDtpChange = 1 Then
    Else
        If Len(txt�۾�����) = txt�۾�����.MaxLength Then
            dtpLabDate.Year = Left(txt�۾�����, 4)
            dtpLabDate.Month = Mid(txt�۾�����, 5, 2)
            dtpLabDate.Day = Right(txt�۾�����, 2)
            txtSlip.SetFocus
        End If
    End If
    
    iDtpChange = 0
End Sub

Private Sub txt�۾�����_Click()
    Call Txt_Highlight(txt�۾�����)
End Sub

Private Sub txt�۾�����_GotFocus()
    Call Txt_Highlight(txt�۾�����)
End Sub

Private Sub txtSlip_Change()
    Dim CPart As DCB0101
    Dim i%
    
    If Len(txtSlip) = txtSlip.MaxLength Then
        If sPrevSlipCd = txtSlip Then
            If txt�۾���ȣ.Enabled = True Then
                txt�۾���ȣ.SetFocus
            End If
        Else
            Set CPart = New DCB0101
            
            CPart.Get_PART Left$(txtSlip, 1), Right$(txtSlip, 2)
            
            i = CPart.CurItemCnt
            
            If i = 0 Then
                MsgBox "�������� �ʴ� �����ڵ��Դϴ�!!"
                Call Txt_Highlight(txtSlip)
                Set CPart = Nothing
                Exit Sub
            ElseIf i = 1 Then
                If iHlpClick = 1 Then
                Else
                    txt�۾���ȣ.SetFocus
                End If
                Set CPart = Nothing
            ElseIf i > 1 Then
                MsgBox "�ڵ弳���� ������ �ֽ��ϴ�!!"
                Call Txt_Highlight(txtSlip)
                Set CPart = Nothing
                Exit Sub
            End If
        End If
    End If
End Sub

Private Sub txtSlip_Click()
    Call Txt_Highlight(txtSlip)
    sPrevSlipCd = txtSlip
End Sub

Private Sub txtSlip_GotFocus()
    Call Txt_Highlight(txtSlip)
    sPrevSlipCd = txtSlip
End Sub

Private Sub txtSlip_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        
        txt�۾���ȣ.SetFocus
    End If
End Sub
