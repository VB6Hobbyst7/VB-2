VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FGB0301 
   Caption         =   "�����ڷ� - TESTITEM"
   ClientHeight    =   7335
   ClientLeft      =   495
   ClientTop       =   780
   ClientWidth     =   11775
   Icon            =   "FGB0301.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   11775
   StartUpPosition =   2  'ȭ�� ���
   Begin FPSpread.vaSpread spdBaseCode 
      Height          =   2055
      Left            =   30
      OleObjectBlob   =   "FGB0301.frx":030A
      TabIndex        =   86
      Top             =   5280
      Width           =   11715
   End
   Begin VB.Frame Fra 
      Appearance      =   0  '���
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   5325
      Left            =   30
      TabIndex        =   45
      Top             =   -30
      Width           =   10350
      Begin Threed.SSCommand cmdSub 
         Height          =   345
         Left            =   5370
         TabIndex        =   99
         Top             =   1320
         Visible         =   0   'False
         Width           =   1155
         _Version        =   65536
         _ExtentX        =   2037
         _ExtentY        =   609
         _StockProps     =   78
         Caption         =   "SUB ITEM"
         ForeColor       =   192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   3
      End
      Begin VB.TextBox txtOtherFlag 
         Height          =   315
         Left            =   8805
         TabIndex        =   90
         Top             =   2430
         Width           =   1395
      End
      Begin Threed.SSPanel pnlNoOpt 
         Height          =   1785
         Left            =   150
         TabIndex        =   76
         Top             =   3480
         Width           =   4095
         _Version        =   65536
         _ExtentX        =   7223
         _ExtentY        =   3149
         _StockProps     =   15
         Caption         =   "SSPanel18"
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
         Begin Threed.SSFrame SSFrame1 
            Height          =   1515
            Left            =   150
            TabIndex        =   77
            Top             =   90
            Width           =   3765
            _Version        =   65536
            _ExtentX        =   6641
            _ExtentY        =   2672
            _StockProps     =   14
            Caption         =   "����ġ ���� Option"
            ForeColor       =   128
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
            Begin VB.OptionButton optNoOpt 
               Appearance      =   0  '���
               BackColor       =   &H00C0FFFF&
               Caption         =   "(UserX-) Low ~ High (+UserX)"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   285
               Index           =   1
               Left            =   210
               TabIndex        =   20
               ToolTipText     =   "��) WBC : 4.8 - 10.8 �� ���� �׸� ����"
               Top             =   330
               Width           =   3375
            End
            Begin VB.OptionButton optNoOpt 
               Appearance      =   0  '���
               BackColor       =   &H00C0C0FF&
               Caption         =   "  ( UserX -  ) ����  >"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   285
               Index           =   3
               Left            =   210
               TabIndex        =   22
               ToolTipText     =   "��) ����ġ > 0.002 �� ���� �׸��� ����"
               Top             =   1080
               Width           =   3375
            End
            Begin VB.OptionButton optNoOpt 
               Appearance      =   0  '���
               BackColor       =   &H00FFFFC0&
               Caption         =   " <  ���� ( + UserX )"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   285
               Index           =   2
               Left            =   210
               TabIndex        =   21
               ToolTipText     =   "��) ����ġ < 100 �� ���� �׸� ������"
               Top             =   690
               Width           =   3375
            End
         End
      End
      Begin VB.TextBox txtSlip 
         Alignment       =   1  '������ ����
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   8  '����
         Left            =   1710
         MaxLength       =   3
         TabIndex        =   0
         Text            =   "001"
         Top             =   240
         Width           =   540
      End
      Begin VB.TextBox txtRefLetter 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   8  '����
         Left            =   1710
         MaxLength       =   15
         TabIndex        =   19
         Text            =   "ABCDEFGHIJKMLN"
         Top             =   3090
         Width           =   2520
      End
      Begin VB.TextBox txtSub 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   8  '����
         Left            =   7980
         MaxLength       =   2
         TabIndex        =   5
         Text            =   "01"
         Top             =   1330
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtDelta 
         Alignment       =   1  '������ ����
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   8  '����
         Left            =   9180
         MaxLength       =   15
         TabIndex        =   40
         Text            =   "100000"
         Top             =   3120
         Width           =   1020
      End
      Begin VB.TextBox txtPrintSpace 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   8  '����
         Left            =   7260
         MaxLength       =   2
         TabIndex        =   11
         Text            =   "Y"
         Top             =   2395
         Width           =   240
      End
      Begin VB.TextBox txtPrintOrd 
         Alignment       =   1  '������ ����
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   8  '����
         Left            =   4380
         MaxLength       =   25
         TabIndex        =   10
         Text            =   "1000"
         Top             =   2385
         Width           =   900
      End
      Begin VB.TextBox txtDisplayOrd 
         Alignment       =   1  '������ ����
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   8  '����
         Left            =   1710
         MaxLength       =   25
         TabIndex        =   9
         Text            =   "100"
         Top             =   2385
         Width           =   930
      End
      Begin VB.TextBox txtUnit 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   8  '����
         Left            =   6240
         MaxLength       =   15
         TabIndex        =   8
         Text            =   "ABCDEFGHIJKMLN"
         Top             =   2035
         Width           =   1920
      End
      Begin VB.TextBox txtPrintNm 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   8  '����
         Left            =   1710
         MaxLength       =   25
         TabIndex        =   7
         Text            =   "WBC MANUAL DIFFERENTIAL C"
         Top             =   2025
         Width           =   3570
      End
      Begin VB.TextBox txtTestNm 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   8  '����
         Left            =   1710
         MaxLength       =   35
         TabIndex        =   6
         Text            =   "WBC MANUAL DIFFERENTIAL COUNT AND M"
         Top             =   1665
         Width           =   4830
      End
      Begin VB.TextBox txtTestSeq 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   8  '����
         Left            =   1710
         MaxLength       =   3
         TabIndex        =   2
         Text            =   "001"
         Top             =   975
         Width           =   540
      End
      Begin VB.TextBox txtSpecimen 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   8  '����
         Left            =   1710
         MaxLength       =   3
         TabIndex        =   1
         Text            =   "001"
         Top             =   622
         Width           =   540
      End
      Begin Threed.SSPanel Panel3D3 
         Height          =   345
         Left            =   150
         TabIndex        =   46
         Top             =   240
         Width           =   1545
         _Version        =   65536
         _ExtentX        =   2725
         _ExtentY        =   609
         _StockProps     =   15
         Caption         =   "SLIP �ڵ�"
         ForeColor       =   8454143
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
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
         Left            =   2250
         TabIndex        =   47
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
         Picture         =   "FGB0301.frx":16E5
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   345
         Left            =   150
         TabIndex        =   49
         Top             =   600
         Width           =   1545
         _Version        =   65536
         _ExtentX        =   2725
         _ExtentY        =   609
         _StockProps     =   15
         Caption         =   "��ü �ڵ�"
         ForeColor       =   8454143
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
      End
      Begin Threed.SSCommand cmdButtonSpc 
         Height          =   330
         Left            =   2250
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   615
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
         Picture         =   "FGB0301.frx":1807
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   345
         Left            =   150
         TabIndex        =   52
         Top             =   960
         Width           =   1545
         _Version        =   65536
         _ExtentX        =   2725
         _ExtentY        =   609
         _StockProps     =   15
         Caption         =   "�˻��׸����"
         ForeColor       =   8454143
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   345
         Left            =   150
         TabIndex        =   53
         Top             =   1320
         Width           =   1545
         _Version        =   65536
         _ExtentX        =   2725
         _ExtentY        =   609
         _StockProps     =   15
         Caption         =   "�˻��׸񱸺�"
         ForeColor       =   8454143
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   330
         Left            =   1710
         TabIndex        =   54
         Top             =   1335
         Width           =   3570
         _Version        =   65536
         _ExtentX        =   6297
         _ExtentY        =   582
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Begin VB.OptionButton optTestGbn 
            Caption         =   "NORMAL"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   570
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   30
            Width           =   1245
         End
         Begin VB.OptionButton optTestGbn 
            Caption         =   "SUB"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   2220
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   30
            Visible         =   0   'False
            Width           =   885
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   345
         Left            =   150
         TabIndex        =   55
         Top             =   1680
         Width           =   1545
         _Version        =   65536
         _ExtentX        =   2725
         _ExtentY        =   609
         _StockProps     =   15
         Caption         =   "�˻��׸��"
         ForeColor       =   8454143
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   345
         Left            =   150
         TabIndex        =   56
         Top             =   2040
         Width           =   1545
         _Version        =   65536
         _ExtentX        =   2725
         _ExtentY        =   609
         _StockProps     =   15
         Caption         =   "��� �׸��"
         ForeColor       =   0
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
      End
      Begin Threed.SSPanel SSPanel7 
         Height          =   345
         Left            =   5400
         TabIndex        =   57
         Top             =   2040
         Width           =   825
         _Version        =   65536
         _ExtentX        =   1455
         _ExtentY        =   609
         _StockProps     =   15
         Caption         =   "����"
         ForeColor       =   0
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   345
         Left            =   150
         TabIndex        =   58
         Top             =   2400
         Width           =   1545
         _Version        =   65536
         _ExtentX        =   2725
         _ExtentY        =   609
         _StockProps     =   15
         Caption         =   "DisplayOrder"
         ForeColor       =   0
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   345
         Index           =   0
         Left            =   2760
         TabIndex        =   59
         Top             =   2390
         Width           =   1605
         _Version        =   65536
         _ExtentX        =   2831
         _ExtentY        =   609
         _StockProps     =   15
         Caption         =   "PrintOrder"
         ForeColor       =   0
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
      End
      Begin Threed.SSPanel SSPanel10 
         Height          =   345
         Left            =   5400
         TabIndex        =   60
         Top             =   2400
         Width           =   1845
         _Version        =   65536
         _ExtentX        =   3254
         _ExtentY        =   609
         _StockProps     =   15
         Caption         =   "��½� �ڸ�����"
         ForeColor       =   0
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
      End
      Begin Threed.SSPanel SSPanel11 
         Height          =   375
         Left            =   4320
         TabIndex        =   62
         Top             =   2745
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "��������"
         ForeColor       =   0
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.74
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
      End
      Begin Threed.SSPanel SSPanel12 
         Height          =   360
         Left            =   5430
         TabIndex        =   63
         Top             =   2760
         Width           =   4770
         _Version        =   65536
         _ExtentX        =   8414
         _ExtentY        =   635
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Begin VB.OptionButton optJudge 
            Caption         =   "Other Flag"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   3270
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   30
            Width           =   1275
         End
         Begin VB.OptionButton optJudge 
            Caption         =   "No"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   30
            Width           =   555
         End
         Begin VB.OptionButton optJudge 
            Caption         =   "Neg/Pos"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   2040
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   30
            Width           =   1155
         End
         Begin VB.OptionButton optJudge 
            Caption         =   "Low/High"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   780
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   30
            Width           =   1185
         End
      End
      Begin Threed.SSPanel SSPanel13 
         Height          =   345
         Left            =   150
         TabIndex        =   64
         Top             =   2760
         Width           =   1545
         _Version        =   65536
         _ExtentX        =   2725
         _ExtentY        =   609
         _StockProps     =   15
         Caption         =   "����ġ ����"
         ForeColor       =   0
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
      End
      Begin Threed.SSPanel SSPanel14 
         Height          =   330
         Left            =   1710
         TabIndex        =   65
         Top             =   2760
         Width           =   2525
         _Version        =   65536
         _ExtentX        =   4454
         _ExtentY        =   582
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Begin VB.OptionButton optRefGbn 
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   1710
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   30
            Width           =   765
         End
         Begin VB.OptionButton optRefGbn 
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   90
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   30
            Width           =   765
         End
         Begin VB.OptionButton optRefGbn 
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   900
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   30
            Width           =   765
         End
      End
      Begin Threed.SSPanel SSPanel20 
         Height          =   345
         Left            =   4320
         TabIndex        =   72
         Top             =   3135
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   609
         _StockProps     =   15
         Caption         =   "Delta ����"
         ForeColor       =   0
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
      End
      Begin Threed.SSPanel SSPanel21 
         Height          =   360
         Left            =   5430
         TabIndex        =   73
         Top             =   3120
         Width           =   2700
         _Version        =   65536
         _ExtentX        =   4762
         _ExtentY        =   635
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Begin VB.OptionButton optDelta 
            Caption         =   "�ۼ�Ʈ"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   1740
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   60
            Width           =   885
         End
         Begin VB.OptionButton optDelta 
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   60
            Width           =   705
         End
         Begin VB.OptionButton optDelta 
            Caption         =   "���밪"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   810
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   60
            Width           =   915
         End
      End
      Begin Threed.SSPanel SSPanel22 
         Height          =   345
         Left            =   8190
         TabIndex        =   85
         Top             =   3150
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   609
         _StockProps     =   15
         Caption         =   "Delta��"
         ForeColor       =   0
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
      End
      Begin Threed.SSPanel pnlSub 
         Height          =   345
         Left            =   6540
         TabIndex        =   87
         Top             =   1330
         Visible         =   0   'False
         Width           =   1410
         _Version        =   65536
         _ExtentX        =   2487
         _ExtentY        =   609
         _StockProps     =   15
         Caption         =   "SUB �˻��׸�"
         ForeColor       =   8454143
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   2
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel pnlRefLetter 
         Height          =   345
         Left            =   150
         TabIndex        =   88
         Top             =   3120
         Width           =   1545
         _Version        =   65536
         _ExtentX        =   2725
         _ExtentY        =   609
         _StockProps     =   15
         Caption         =   "����ġ ����"
         ForeColor       =   0
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
      End
      Begin VB.Frame fraRefNum 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   4320
         TabIndex        =   66
         Top             =   3510
         Width           =   5925
         Begin VB.TextBox txtUpperGrayF 
            Alignment       =   1  '������ ����
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3840
            MaxLength       =   9
            TabIndex        =   34
            Top             =   930
            Width           =   705
         End
         Begin VB.TextBox txtLowerGrayF 
            Alignment       =   1  '������ ����
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2430
            MaxLength       =   9
            TabIndex        =   32
            Top             =   930
            Width           =   675
         End
         Begin VB.TextBox txtUpperGrayM 
            Alignment       =   1  '������ ����
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1590
            MaxLength       =   9
            TabIndex        =   30
            Top             =   930
            Width           =   675
         End
         Begin VB.TextBox txtLowerGrayM 
            Alignment       =   1  '������ ����
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   150
            MaxLength       =   9
            TabIndex        =   28
            Top             =   930
            Width           =   705
         End
         Begin VB.TextBox txtPanicLo 
            Alignment       =   1  '������ ����
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4710
            TabIndex        =   35
            Top             =   570
            Width           =   1035
         End
         Begin VB.TextBox txtPanicHi 
            Alignment       =   1  '������ ����
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4710
            TabIndex        =   36
            Top             =   930
            Width           =   1035
         End
         Begin VB.TextBox txtRefLoM 
            Alignment       =   1  '������ ����
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   150
            MaxLength       =   20
            TabIndex        =   27
            Text            =   "123"
            Top             =   570
            Width           =   945
         End
         Begin VB.TextBox txtRefHiM 
            Alignment       =   1  '������ ����
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1320
            TabIndex        =   29
            Top             =   570
            Width           =   945
         End
         Begin VB.TextBox txtRefLoF 
            Alignment       =   1  '������ ����
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2430
            TabIndex        =   31
            Top             =   570
            Width           =   945
         End
         Begin VB.TextBox txtRefHiF 
            Alignment       =   1  '������ ����
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3600
            TabIndex        =   33
            Top             =   570
            Width           =   945
         End
         Begin Threed.SSPanel SSPanel15 
            Height          =   345
            Left            =   900
            TabIndex        =   67
            Top             =   930
            Width           =   645
            _Version        =   65536
            _ExtentX        =   1138
            _ExtentY        =   609
            _StockProps     =   15
            Caption         =   "��"
            ForeColor       =   65535
            BackColor       =   8388608
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.76
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSPanel SSPanel16 
            Height          =   345
            Left            =   3150
            TabIndex        =   68
            Top             =   930
            Width           =   645
            _Version        =   65536
            _ExtentX        =   1138
            _ExtentY        =   609
            _StockProps     =   15
            Caption         =   "��"
            ForeColor       =   65535
            BackColor       =   128
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.76
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSPanel SSPanel9 
            Height          =   255
            Index           =   1
            Left            =   4710
            TabIndex        =   74
            Top             =   300
            Width           =   1035
            _Version        =   65536
            _ExtentX        =   1826
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "Panic"
            ForeColor       =   8454143
            BackColor       =   8421376
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label14 
            BackStyle       =   0  '����
            Caption         =   "High Ref."
            Height          =   195
            Left            =   3750
            TabIndex        =   98
            Top             =   330
            Width           =   915
         End
         Begin VB.Label Label13 
            BackStyle       =   0  '����
            Caption         =   "Low Ref."
            Height          =   195
            Left            =   2430
            TabIndex        =   97
            Top             =   330
            Width           =   885
         End
         Begin VB.Label Label12 
            BackStyle       =   0  '����
            Caption         =   "+ UserX"
            Height          =   195
            Left            =   3810
            TabIndex        =   96
            Top             =   1320
            Width           =   705
         End
         Begin VB.Label Label11 
            BackStyle       =   0  '����
            Caption         =   "UserX - "
            Height          =   255
            Left            =   2430
            TabIndex        =   95
            Top             =   1320
            Width           =   825
         End
         Begin VB.Label Label9 
            BackStyle       =   0  '����
            Caption         =   "High Ref."
            Height          =   195
            Left            =   1470
            TabIndex        =   94
            Top             =   330
            Width           =   915
         End
         Begin VB.Label Label8 
            BackStyle       =   0  '����
            Caption         =   "+ UserX"
            Height          =   255
            Left            =   1560
            TabIndex        =   93
            Top             =   1320
            Width           =   795
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '����
            Caption         =   "Low Ref."
            Height          =   255
            Left            =   150
            TabIndex        =   92
            Top             =   330
            Width           =   1005
         End
         Begin VB.Label Label7 
            BackStyle       =   0  '����
            Caption         =   "UserX - "
            Height          =   255
            Left            =   150
            TabIndex        =   89
            Top             =   1320
            Width           =   825
         End
         Begin VB.Label Label5 
            Caption         =   "I"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   5190
            TabIndex        =   75
            Top             =   810
            Width           =   135
         End
         Begin VB.Label Label6 
            Caption         =   "~"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   1140
            TabIndex        =   70
            Top             =   660
            Width           =   135
         End
         Begin VB.Label Label4 
            Caption         =   "~"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   3420
            TabIndex        =   69
            Top             =   660
            Width           =   135
         End
      End
      Begin VB.Frame fraUpLow 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   4320
         TabIndex        =   78
         Top             =   3570
         Width           =   5925
         Begin VB.TextBox txtGrayF 
            Alignment       =   1  '������ ����
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4200
            MaxLength       =   20
            TabIndex        =   26
            Top             =   900
            Width           =   1185
         End
         Begin VB.TextBox txtGrayM 
            Alignment       =   1  '������ ����
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1590
            MaxLength       =   20
            TabIndex        =   24
            Top             =   900
            Width           =   1185
         End
         Begin VB.TextBox txtUpLowF 
            Alignment       =   1  '������ ����
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4200
            MaxLength       =   20
            TabIndex        =   25
            Top             =   450
            Width           =   1185
         End
         Begin VB.TextBox txtUpLowM 
            Alignment       =   1  '������ ����
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1590
            MaxLength       =   20
            TabIndex        =   23
            Top             =   450
            Width           =   1185
         End
         Begin Threed.SSPanel SSPanel17 
            Height          =   345
            Left            =   450
            TabIndex        =   79
            Top             =   450
            Width           =   645
            _Version        =   65536
            _ExtentX        =   1138
            _ExtentY        =   609
            _StockProps     =   15
            Caption         =   "��"
            ForeColor       =   65535
            BackColor       =   8388608
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   10.51
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSPanel SSPanel19 
            Height          =   345
            Left            =   3060
            TabIndex        =   80
            Top             =   450
            Width           =   645
            _Version        =   65536
            _ExtentX        =   1138
            _ExtentY        =   609
            _StockProps     =   15
            Caption         =   "��"
            ForeColor       =   65535
            BackColor       =   128
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   10.51
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label10 
            Caption         =   "UserX"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3480
            TabIndex        =   84
            Top             =   930
            Width           =   645
         End
         Begin VB.Label lblUpLow 
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   3750
            TabIndex        =   83
            Top             =   480
            Width           =   525
         End
         Begin VB.Label Label3 
            Caption         =   "UserX"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   900
            TabIndex        =   82
            Top             =   930
            Width           =   645
         End
         Begin VB.Label lblUpLow 
            Caption         =   "����"
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
            Index           =   0
            Left            =   1140
            TabIndex        =   81
            Top             =   480
            Width           =   525
         End
      End
      Begin VB.Label lblOtherFlag 
         Caption         =   "OtherFlag ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   8790
         TabIndex        =   91
         Top             =   2220
         Width           =   1365
      End
      Begin VB.Label lblTestNm 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '���� ����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2280
         TabIndex        =   71
         Top             =   975
         Width           =   5955
      End
      Begin VB.Label Label2 
         Caption         =   "(Y/N)"
         Height          =   225
         Left            =   7620
         TabIndex        =   61
         Top             =   2460
         Width           =   525
      End
      Begin VB.Label lblSpecimenNm 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '���� ����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2550
         TabIndex        =   51
         Top             =   615
         Width           =   5685
      End
      Begin VB.Label lblSlipNm 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '���� ����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2550
         TabIndex        =   48
         Top             =   240
         Width           =   4185
      End
   End
   Begin Threed.SSCommand cmdDelete 
      Height          =   1005
      Left            =   10530
      TabIndex        =   43
      Top             =   2070
      Width           =   1125
      _Version        =   65536
      _ExtentX        =   1976
      _ExtentY        =   1773
      _StockProps     =   78
      Caption         =   "���� F4"
      ForeColor       =   32768
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
      Picture         =   "FGB0301.frx":1929
   End
   Begin Threed.SSCommand cmdSearch 
      Height          =   1005
      Left            =   10530
      TabIndex        =   42
      Top             =   1050
      Width           =   1125
      _Version        =   65536
      _ExtentX        =   1976
      _ExtentY        =   1773
      _StockProps     =   78
      Caption         =   "��ȸ F3"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
      Picture         =   "FGB0301.frx":2203
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   1005
      Left            =   10530
      TabIndex        =   44
      Top             =   3090
      Width           =   1125
      _Version        =   65536
      _ExtentX        =   1984
      _ExtentY        =   1773
      _StockProps     =   78
      Caption         =   "���� ESC"
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
      Picture         =   "FGB0301.frx":2ADD
   End
   Begin Threed.SSCommand cmdReg 
      Height          =   945
      Left            =   10530
      TabIndex        =   41
      Top             =   90
      Width           =   1125
      _Version        =   65536
      _ExtentX        =   1984
      _ExtentY        =   1667
      _StockProps     =   78
      Caption         =   "��� F2"
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
      Picture         =   "FGB0301.frx":33B7
   End
End
Attribute VB_Name = "FGB0301"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TestItemTable() As TestItemTBL
Dim iValid%     'Validation Check�� ���
Dim ioptNoOptFocus%
Dim iDefaultView%
Dim iCurSelRow%
Dim sPrevSlip$
Dim sPrevSpc$
Dim iSpdClick%

Private Function fBlankToZero(ByVal sBuf As String) As String
    If sBuf = "" Then
        fBlankToZero = "0"
    Else
        fBlankToZero = sBuf
    End If
End Function

Private Sub BaseCodeInit()
    Dim CTestItem As DCB0101
    Dim i%
    Dim j%
    Dim s01$, s02$, s03$, s04$, s05$, s06$, s07$, s08$, s09$, s10$
    Dim s11$, s12$, s13$, s14$, s15$, s16$, s17$, s18$, s19$, s20$
    Dim s21$, s22$, s23$, s24$, s25$, s26$, s27$, s28$, s29$, s30$
    Dim s31$, s32$, s33$, s34$
    
    Dim vTmp, vTmp2
    Dim bMatch As Boolean
    
    If txtSlip = "" Or txtSpecimen = "" Then
        Exit Sub
    End If
    
    Set CTestItem = New DCB0101
    
    CTestItem.Get_TESTITEM 1, Left$(txtSlip, 1), Mid$(txtSlip, 2, 2), txtSpecimen
    
    i = CTestItem.CurItemCnt
    
    If i = 0 Then
        MsgBox "���� �����ڷῡ � �׸� ��ϵǾ� ���� �ʽ��ϴ�!!"
        Set CTestItem = Nothing
        Exit Sub
    End If
    
    s01 = CTestItem.TotField01: s02 = CTestItem.TotField02:  s03 = CTestItem.TotField03
    s04 = CTestItem.TotField04: s05 = CTestItem.TotField05:  s06 = CTestItem.TotField06
    s07 = CTestItem.TotField07: s08 = CTestItem.TotField08:  s09 = CTestItem.TotField09
    s10 = CTestItem.TotField10: s11 = CTestItem.TotField11:  s12 = CTestItem.TotField12
    s13 = CTestItem.TotField13: s14 = CTestItem.TotField14:  s15 = CTestItem.TotField15
    s16 = CTestItem.TotField16: s17 = CTestItem.TotField17:  s18 = CTestItem.TotField18
    s19 = CTestItem.TotField19: s20 = CTestItem.TotField20:  s21 = CTestItem.TotField21
    s22 = CTestItem.TotField22: s23 = CTestItem.TotField23:  s24 = CTestItem.TotField24
    s25 = CTestItem.TotField25: s26 = CTestItem.TotField26:  s27 = CTestItem.TotField27
    s28 = CTestItem.TotField28: s29 = CTestItem.TotField29:  s30 = CTestItem.TotField30
    s31 = CTestItem.TotField31: s32 = CTestItem.TotField32:   s33 = CTestItem.TotField33
    s34 = CTestItem.TotField34
    
    Erase TestItemTable
    
    ReDim TestItemTable(i) As TestItemTBL
    
    For j = 1 To i
    
        TestItemTable(j).s01 = GetByOne(s01, s01):        TestItemTable(j).s02 = GetByOne(s02, s02)
        TestItemTable(j).s03 = GetByOne(s03, s03):        TestItemTable(j).s04 = GetByOne(s04, s04)
        TestItemTable(j).s05 = GetByOne(s05, s05):        TestItemTable(j).s06 = GetByOne(s06, s06)
        TestItemTable(j).s07 = GetByOne(s07, s07):        TestItemTable(j).s08 = GetByOne(s08, s08)
        TestItemTable(j).s09 = GetByOne(s09, s09):        TestItemTable(j).s10 = GetByOne(s10, s10)
        TestItemTable(j).s11 = GetByOne(s11, s11):        TestItemTable(j).s12 = GetByOne(s12, s12)
        TestItemTable(j).s13 = GetByOne(s13, s13):        TestItemTable(j).s14 = GetByOne(s14, s14)
        TestItemTable(j).s15 = GetByOne(s15, s15):        TestItemTable(j).s16 = GetByOne(s16, s16)
        TestItemTable(j).s17 = GetByOne(s17, s17):        TestItemTable(j).s18 = GetByOne(s18, s18)
        TestItemTable(j).s19 = GetByOne(s19, s19):        TestItemTable(j).s20 = GetByOne(s20, s20)
        TestItemTable(j).s21 = GetByOne(s21, s21):        TestItemTable(j).s22 = GetByOne(s22, s22)
        TestItemTable(j).s23 = GetByOne(s23, s23):        TestItemTable(j).s24 = GetByOne(s24, s24)
        TestItemTable(j).s25 = GetByOne(s25, s25):        TestItemTable(j).s26 = GetByOne(s26, s26)
        TestItemTable(j).s27 = GetByOne(s27, s27):        TestItemTable(j).s28 = GetByOne(s28, s28)
        TestItemTable(j).s29 = GetByOne(s29, s29):        TestItemTable(j).s30 = GetByOne(s30, s30)
        TestItemTable(j).s31 = GetByOne(s31, s31):        TestItemTable(j).s32 = GetByOne(s32, s32)
        TestItemTable(j).s33 = GetByOne(s33, s33):        TestItemTable(j).s34 = GetByOne(s34, s34)
        
        bMatch = False
        
        With spdBaseCode
            
            If j = 1 Then
            Else
                Call .GetText(1, .MaxRows, vTmp)
                Call .GetText(2, .MaxRows, vTmp2)
                
                If IsNull(vTmp2) = True Then
                    vTmp2 = ""
                End If
                
                If vTmp = TestItemTable(j).s01 & TestItemTable(j).s02 & TestItemTable(j).s03 & TestItemTable(j).s04 Then
                    If vTmp2 = "N" Or vTmp2 = "C" Then
                        bMatch = True
                    ElseIf Right$(vTmp2, 2) <> Left$(TestItemTable(j).s05, 2) Or Right$(vTmp2, 2) <> Right$(TestItemTable(j).s05, 2) Then  'Sub Or MultiRef
                        bMatch = False
                    End If
                End If
            End If
            
            If bMatch = False Then
                .MaxRows = .MaxRows + 1
                Call .SetText(1, .MaxRows, TestItemTable(j).s01 & TestItemTable(j).s02 & TestItemTable(j).s03 & TestItemTable(j).s04 & "")
            
                If TestItemTable(j).s05 = "NNNN" Then   'Normal
                    Call .SetText(2, .MaxRows, "N")
                ElseIf IsNumeric(Left$(TestItemTable(j).s05, 2)) = True Then   'sub Yes
                    Call .SetText(2, .MaxRows, "S" & Left$(TestItemTable(j).s05, 2) & "")
                End If
                
                Call .SetText(3, .MaxRows, TestItemTable(j).s07 & "")      '�˻��
                
                If TestItemTable(j).s13 = "0" Then    '<-------------- ����ġ ����
                    '����ġ����üũ
                    Call .SetText(4, .MaxRows, "1")
                    
                ElseIf TestItemTable(j).s13 = "1" Then    '<----------- ����ġ ����
                    '����ġ ������ ���븸 ȭ�� ���
                    Call .SetText(5, .MaxRows, TestItemTable(j).s28 & "")
                    
                ElseIf TestItemTable(j).s13 = "2" Then    '<-------------------- ����ġ ���� - LowerGray/Low - High/UpperGray
                        
                    Call .SetText(6, .MaxRows, TestItemTable(j).s21 & "")      'LowM
                    Call .SetText(7, .MaxRows, TestItemTable(j).s22 & "")       'HighM
                    Call .SetText(8, .MaxRows, TestItemTable(j).s29 & "")      'LowF
                    Call .SetText(9, .MaxRows, TestItemTable(j).s30 & "")      'HighF
                                            
                    Call .SetText(12, .MaxRows, TestItemTable(j).s25 & "")      'UpperGrayM
                    Call .SetText(13, .MaxRows, TestItemTable(j).s26 & "")      'LowerGrayM
                    Call .SetText(16, .MaxRows, TestItemTable(j).s31 & "")      'UpperGrayF
                    Call .SetText(17, .MaxRows, TestItemTable(j).s32 & "")      'LowerGrayF
                
                    'Low - High ������ Panic ����
                    If TestItemTable(j).s14 = "1" Then    'Panic ����
                        Call .SetText(19, .MaxRows, TestItemTable(j).s19 & "")
                        Call .SetText(20, .MaxRows, TestItemTable(j).s20 & "")
                    End If
                    
                ElseIf TestItemTable(j).s13 = "3" Then  '<--------------------- ����ġ ���� - UpperLimit / UpperGrayZone
                
                    Call .SetText(10, .MaxRows, TestItemTable(j).s23 & "")     'UpperLimitM
                    Call .SetText(12, .MaxRows, TestItemTable(j).s25 & "")     'UpperGrayZoneM
                    Call .SetText(14, .MaxRows, TestItemTable(j).s31 & "")     'UpperLimitF
                    Call .SetText(16, .MaxRows, TestItemTable(j).s33 & "")     'UpperGrayZoneF
            
                ElseIf TestItemTable(j).s13 = "4" Then      '<-------------------- ����ġ ���� - LowerGrayZone / LowerLimit
                
                    Call .SetText(11, .MaxRows, TestItemTable(j).s24 & "")      'LowerLimitM
                    Call .SetText(13, .MaxRows, TestItemTable(j).s26 & "")      'LowerGrayZoneM
                    Call .SetText(15, .MaxRows, TestItemTable(j).s32 & "")      'LowerLimitF
                    Call .SetText(17, .MaxRows, TestItemTable(j).s34 & "")      'LowerGrayZoneF
                                    
                End If
                
                If TestItemTable(j).s27 <> "" Then
                    Call .SetText(18, .MaxRows, TestItemTable(j).s27 & "")
                End If
            End If
        End With
    Next
    
    Set CTestItem = Nothing
    
End Sub

Private Sub CompareSlip()
    Dim CPart As DCB0101
    Dim i%
    
    Set CPart = New DCB0101
        
    If txtSlip = "" Then
        Exit Sub
    End If
    
    CPart.Get_PART Left$(txtSlip, 1), Right$(txtSlip, 2)
    
    i = CPart.CurItemCnt
    
    If i = 0 Then
        lblSlipNm = ""
        txtSpecimen = ""
        lblSpecimenNm = ""
        txtTestSeq = ""
        'Call Txt_Highlight(txtSlip)
        Set CPart = Nothing
        Exit Sub
    ElseIf i = 1 Then
        lblSlipNm = GetByOne(CPart.TotField03, CPart.TotField03)
        'txtSpecimen.SetFocus
        Set CPart = Nothing
    ElseIf i > 1 Then
        MsgBox "�ڵ弳���� ������ �ֽ��ϴ�!!"
        'Call Txt_Highlight(txtSlip)
        Set CPart = Nothing
        Exit Sub
    End If
    
End Sub

Private Sub CompareSpecimen()
    Dim CSpecimen As DCB0101
    Dim i%
    
    Set CSpecimen = New DCB0101
    
    CSpecimen.Get_SPC txtSpecimen
    
    i = CSpecimen.CurItemCnt
    
    If i = 0 Then
        Call Txt_Highlight(txtSpecimen)
        lblSpecimenNm = ""
        txtTestSeq = ""
        txtTestNm = ""
        txtPrintNm = ""
        Set CSpecimen = Nothing
        Exit Sub
    ElseIf i = 1 Then
        lblSpecimenNm = GetByOne(CSpecimen.TotField02, CSpecimen.TotField02)
        Set CSpecimen = Nothing
    ElseIf i > 1 Then
        MsgBox "�ڵ弳���� ������ �ֽ��ϴ�!!"
        Set CSpecimen = Nothing
        Exit Sub
    End If
    
End Sub

Private Sub CompareTestNm(Optional ByVal iMode As Integer)
    Dim CTestItem As DCB0101
    Dim i%, j%
    Dim sSub$, sMulti$
    Dim s01$, s02$, s03$, s04$, s05$, s06$, s07$, s08$, s09$, s10$
    Dim s11$, s12$, s13$, s14$, s15$, s16$, s17$, s18$, s19$, s20$
    Dim s21$, s22$, s23$, s24$, s25$, s26$, s27$, s28$, s29$, s30$
    Dim s31$, s32$, s33$, s34$, s35$
    
    Set CTestItem = New DCB0101
        
    If iMode = 1 Then   'With SubCd
        CTestItem.Get_TESTITEM 4, Left$(txtSlip, 1), Right$(txtSlip, 2), txtSpecimen, txtTestSeq, txtSub
    ElseIf iMode = 2 Then  'Default Cd Normal�̳� Sub��
        CTestItem.Get_TESTITEM 3, Left$(txtSlip, 1), Right$(txtSlip, 2), txtSpecimen, txtTestSeq
    ElseIf iMode = 3 Then   'SUB �� ���� ������ �׸�
        CTestItem.Get_TESTITEM 5, Left$(txtSlip, 1), Right$(txtSlip, 2), txtSpecimen, txtTestSeq
    End If
    
    i = CTestItem.CurItemCnt
    
    If i = 0 Then
        If iMode = 1 Then   'With SubCd�� �����ϴ���
            Call DisplayInit(3)
            Set CTestItem = Nothing
            txtTestNm.SetFocus
            Exit Sub
        ElseIf iMode = 2 Then 'DefaultCd�� �����ϴ���
            Call DisplayInit(2)
            Set CTestItem = Nothing
            txtTestNm.SetFocus
            Exit Sub
        ElseIf iMode = 3 Then   'SUB�� ���� �������׸� �����ϴ³�
            txtSub = "00"
            Call DisplayInit(4)
        End If
    ElseIf i < 3 Then
                   
        s01 = CTestItem.TotField01: s02 = CTestItem.TotField02:  s03 = CTestItem.TotField03
        s04 = CTestItem.TotField04: s05 = CTestItem.TotField05:  s06 = CTestItem.TotField06
        s07 = CTestItem.TotField07: s08 = CTestItem.TotField08:  s09 = CTestItem.TotField09
        s10 = CTestItem.TotField10: s11 = CTestItem.TotField11:  s12 = CTestItem.TotField12
        s13 = CTestItem.TotField13: s14 = CTestItem.TotField14:  s15 = CTestItem.TotField15
        s16 = CTestItem.TotField16: s17 = CTestItem.TotField17:  s18 = CTestItem.TotField18
        s19 = CTestItem.TotField19: s20 = CTestItem.TotField20:  s21 = CTestItem.TotField21
        s22 = CTestItem.TotField22: s23 = CTestItem.TotField23:  s24 = CTestItem.TotField24
        s25 = CTestItem.TotField25: s26 = CTestItem.TotField26:  s27 = CTestItem.TotField27
        s28 = CTestItem.TotField28: s29 = CTestItem.TotField29:  s30 = CTestItem.TotField30
        s31 = CTestItem.TotField31: s32 = CTestItem.TotField32:  s33 = CTestItem.TotField33
        s34 = CTestItem.TotField34
        
        Set CTestItem = Nothing
        
        Erase TestItemTable
        
        ReDim TestItemTable(i) As TestItemTBL
        
        For j = 1 To i
        
            TestItemTable(j).s01 = GetByOne(s01, s01):        TestItemTable(j).s02 = GetByOne(s02, s02)
            TestItemTable(j).s03 = GetByOne(s03, s03):        TestItemTable(j).s04 = GetByOne(s04, s04)
            TestItemTable(j).s05 = GetByOne(s05, s05):        TestItemTable(j).s06 = GetByOne(s06, s06)
            TestItemTable(j).s07 = GetByOne(s07, s07):        TestItemTable(j).s08 = GetByOne(s08, s08)
            TestItemTable(j).s09 = GetByOne(s09, s09):        TestItemTable(j).s10 = GetByOne(s10, s10)
            TestItemTable(j).s11 = GetByOne(s11, s11):        TestItemTable(j).s12 = GetByOne(s12, s12)
            TestItemTable(j).s13 = GetByOne(s13, s13):        TestItemTable(j).s14 = GetByOne(s14, s14)
            TestItemTable(j).s15 = GetByOne(s15, s15):        TestItemTable(j).s16 = GetByOne(s16, s16)
            TestItemTable(j).s17 = GetByOne(s17, s17):        TestItemTable(j).s18 = GetByOne(s18, s18)
            TestItemTable(j).s19 = GetByOne(s19, s19):        TestItemTable(j).s20 = GetByOne(s20, s20)
            TestItemTable(j).s21 = GetByOne(s21, s21):        TestItemTable(j).s22 = GetByOne(s22, s22)
            TestItemTable(j).s23 = GetByOne(s23, s23):        TestItemTable(j).s24 = GetByOne(s24, s24)
            TestItemTable(j).s25 = GetByOne(s25, s25):        TestItemTable(j).s26 = GetByOne(s26, s26)
            TestItemTable(j).s27 = GetByOne(s27, s27):        TestItemTable(j).s28 = GetByOne(s28, s28)
            TestItemTable(j).s29 = GetByOne(s29, s29):        TestItemTable(j).s30 = GetByOne(s30, s30)
            TestItemTable(j).s31 = GetByOne(s31, s31):        TestItemTable(j).s32 = GetByOne(s32, s32)
            TestItemTable(j).s33 = GetByOne(s33, s33):        TestItemTable(j).s34 = GetByOne(s34, s34)
        Next
        
        Call DisplayInfoOfSpecificTestItem(i)
        
        If iMode = 1 Or iMode = 2 Then
        Else
            txtTestNm.SetFocus
        End If
    ElseIf i >= 3 Then
        MsgBox "�ڵ弳���� ������ �ֽ��ϴ�!!"
        Set CTestItem = Nothing
        Exit Sub
    End If
    
End Sub

Private Sub DefaultItemDisplay()
    txtSlip = fCurUserSlipCd

    lblSlipNm.Caption = fCurUserSlipNm
        
    txtSpecimen = fCurUserSpcCd
    
    lblSpecimenNm = fCurUserSpcNm
    
    If txtSlip = "" Then
        Exit Sub
    End If
    
    If txtSpecimen = "" Then
        txtSpecimen.TabIndex = 0
        Exit Sub
    End If
    
    iDefaultView = 2
    
    'txtTestSeq.TabIndex = 0
    
End Sub

Private Sub DisplayInit(ByVal iMode As Integer)
    If iMode = 1 Then
        txtSlip = ""
        txtSpecimen = ""
        txtTestSeq = ""
    End If
    
    If iMode = 1 Or iMode = 2 Or iMode = 3 Or iMode = 4 Then
        txtTestNm = ""
        lblTestNm.Caption = ""
        txtPrintNm = ""
    End If
    
    txtUnit = ""
    txtDisplayOrd = ""
    txtPrintOrd = ""
    txtPrintSpace = "N"
    txtDelta = ""
    txtRefLetter = ""
    txtOtherFlag = ""
    
    If iMode = 3 Or iMode = 4 Then
    Else
        optTestGbn(1).Enabled = True
        optTestGbn(2).Enabled = True
        
        optTestGbn(1).Value = True  'Sub or Multi�� ��� ������ ����� ���� ����� �� ����
        
        pnlSub.Visible = False
        txtSub.Visible = False
        txtSub = ""
        'pnlMulti.Visible = False
        'txtMulti.Visible = False
    End If
    
    optJudge(1).Value = True
    optRefGbn(1).Value = True
    optDelta(1).Value = True
    
    fraRefNum.Visible = False
    txtRefLoM = ""
    txtLowerGrayM = ""
    txtRefHiM = ""
    txtUpperGrayM = ""
    txtRefLoF = ""
    txtLowerGrayF = ""
    txtRefHiF = ""
    txtUpperGrayF = ""
        
    fraUpLow.Visible = False
    txtUpLowM = ""
    txtGrayM = ""
    txtUpLowF = ""
    txtGrayF = ""
    pnlNoOpt.Visible = False
    lblOtherFlag.Visible = False
    txtOtherFlag.Visible = False
    txtOtherFlag = ""
    
    If iMode = 1 Then
        'SpreadBackColor Option
        iSpdBackColorOption = 2
        
        With spdBaseCode
            .MaxRows = 0
            
            .BlockMode = True
            .Col = -1
            .Col2 = -1
            .Row = -1
            .Row2 = -1
            .BackColorStyle = BackColorStyleUnderGrid
            .BackColor = SpdBackcolor(iSpdBackColorOption)      'GBR
            .EditModePermanent = True
            .Protect = True
            .BlockMode = False
            
            .BlockMode = True
            .Col = -1
            .Col2 = -1
            .Row = -1
            .Row2 = -1
            .Lock = True
            .BlockMode = False
            
            .ColsFrozen = 3
            
        End With
    End If
End Sub

Private Sub DisplayInfoOfSpecificTestItem(ByVal i As Integer)
    Dim j%
    
    
'''    pnlSub.Visible = True
'''    txtSub.Visible = True
    
'''    optTestGbn(1).Enabled = True
'''    optTestGbn(2).Enabled = True
        
    If TestItemTable(1).s05 = "NNNN" Then   'Normal
        pnlSub.Visible = False
        txtSub.Visible = False
        txtSub = ""
        optTestGbn(1).Value = True
        optTestGbn(1).Enabled = True
        optTestGbn(2).Enabled = False
    End If
    
    txtTestSeq = TestItemTable(1).s04
    
    If IsNumeric(Left$(TestItemTable(1).s05, 2)) = True Then 'Sub
        pnlSub.Visible = True
        txtSub.Visible = True
        optTestGbn(2).Value = True
        optTestGbn(2).Enabled = True
        optTestGbn(1).Enabled = False
        txtSub = Left$(TestItemTable(1).s05, 2)
    End If
    
    txtSlip = TestItemTable(1).s01 & TestItemTable(1).s02
    txtSpecimen = TestItemTable(1).s03
    txtTestSeq = TestItemTable(1).s04
    
    
    txtTestNm = TestItemTable(1).s07        'TestName
    lblTestNm.Caption = TestItemTable(1).s07
    txtPrintNm = TestItemTable(1).s08       'PrintName
    txtUnit = TestItemTable(1).s09          'Unit
    txtDisplayOrd = TestItemTable(1).s10    'DisplayOrder
    txtPrintOrd = TestItemTable(1).s11
    
    If TestItemTable(1).s12 = "0" Then
        txtPrintSpace = "N"
    ElseIf TestItemTable(1).s12 = "1" Then
        txtPrintSpace = "Y"
    End If
    
    txtDelta = "": txtRefLetter = "": txtRefLoM = "": txtRefHiM = ""
    txtRefLoF = "": txtRefHiF = "": txtPanicLo = "": txtPanicHi = ""
    txtUpLowM = "": txtUpLowF = "": txtGrayM = "": txtGrayF = ""
    txtLowerGrayM = "": txtLowerGrayF = "": txtUpperGrayM = "": txtUpperGrayF = ""
    txtOtherFlag = ""
    
'<--------------- ����ġ ���� flag 0 -----------------------------
    If TestItemTable(1).s13 = "0" Then
        optRefGbn(1).Value = True
        optJudge(1).Value = True
        
        txtRefLetter.Locked = True
        txtDelta.Locked = True
        
        pnlNoOpt.Visible = False
        fraRefNum.Visible = False
        txtRefLoM = ""
        txtLowerGrayM = ""
        txtRefHiM = ""
        txtUpperGrayM = ""
        txtRefLoF = ""
        txtLowerGrayF = ""
        txtRefHiF = ""
        txtUpperGrayF = ""
        
        fraUpLow.Visible = False
        txtUpLowM = ""
        txtGrayM = ""
        txtUpLowF = ""
        txtGrayF = ""
        
        lblOtherFlag.Visible = False
        txtOtherFlag.Visible = False
        txtOtherFlag = ""
        
'<--------------- ����ġ ���� flag 1 -------------------------------
    ElseIf TestItemTable(1).s13 = "1" Then
        txtRefLetter.Locked = False
        optRefGbn(2).Value = True
        
        pnlNoOpt.Visible = False
        fraRefNum.Visible = False
        txtRefLoM = ""
        txtLowerGrayM = ""
        txtRefHiM = ""
        txtUpperGrayM = ""
        txtRefLoF = ""
        txtLowerGrayF = ""
        txtRefHiF = ""
        txtUpperGrayF = ""
        
        txtRefLoM = ""
        txtLowerGrayM = ""
        txtRefHiM = ""
        txtUpperGrayM = ""
        txtRefLoF = ""
        txtLowerGrayF = ""
        txtRefHiF = ""
        txtUpperGrayF = ""
        fraUpLow.Visible = False
        txtUpLowM = ""
        txtGrayM = ""
        txtUpLowF = ""
        txtGrayF = ""
    
        lblOtherFlag.Visible = False
        txtOtherFlag.Visible = False
        txtOtherFlag = ""
        
        '����ġ ������ ���� ��Ÿ��
        txtRefLetter = TestItemTable(1).s28
        txtDelta.Locked = True
        
        If TestItemTable(1).s18 = "0" Then  '��������
            optJudge(1).Value = True
            
        ElseIf TestItemTable(1).s18 = "1" Then '���� Low/High
        ElseIf TestItemTable(1).s18 = "2" Then  '���� Neg/Pos
            optJudge(3).Value = True
            
        ElseIf TestItemTable(1).s18 = "3" Then  '���� OtherFlag
            optJudge(4).Value = True
            lblOtherFlag.Visible = True
            txtOtherFlag.Visible = True
            txtOtherFlag = TestItemTable(1).s27
        End If
    
'<------------------ ����ġ ���� flag 2, 3, 4 ----------------------------
    Else
        txtRefLetter.Locked = True
        optRefGbn(3).Value = True
        
        pnlNoOpt.Visible = True
                
        '<-- Flag 2 ----- LowerGrayZone/Low - High/UpperGrayZone-------
        If TestItemTable(1).s13 = "2" Then
            optNoOpt(1).Value = True
            fraUpLow.Visible = False
            txtUpLowM = ""
            txtGrayM = ""
            txtUpLowF = ""
            txtGrayF = ""
            
            fraRefNum.Visible = True
            
            For j = 1 To i
                txtRefLoM = TestItemTable(j).s21        'Low
                txtRefHiM = TestItemTable(j).s22        'High
                txtLowerGrayM = TestItemTable(j).s26    'LowerGray
                txtUpperGrayM = TestItemTable(j).s25    'UpperGray
                
                txtRefLoF = TestItemTable(j).s29
                txtRefHiF = TestItemTable(j).s30
                txtLowerGrayF = TestItemTable(j).s34
                txtUpperGrayF = TestItemTable(j).s33
            Next
            
            If TestItemTable(1).s14 = "1" Then  'Panic  Low - High������
                txtPanicLo = TestItemTable(1).s19
                txtPanicHi = TestItemTable(1).s20
            End If
        ''<-- Flag 3 ----- UpperLimit/UpperGrayZone-------
        ElseIf TestItemTable(1).s13 = "3" Then
            optNoOpt(2).Value = True
            fraUpLow.Visible = True
            fraRefNum.Visible = False
            txtRefLoM = ""
            txtLowerGrayM = ""
            txtRefHiM = ""
            txtUpperGrayM = ""
            txtRefLoF = ""
            txtLowerGrayF = ""
            txtRefHiF = ""
            txtUpperGrayF = ""
        
            txtRefLoM = ""
            txtLowerGrayM = ""
            txtRefHiM = ""
            txtUpperGrayM = ""
            txtRefLoF = ""
            txtLowerGrayF = ""
            txtRefHiF = ""
            txtUpperGrayF = ""
            
            For j = 1 To i
                txtUpLowM = TestItemTable(j).s23    'UpperLimit
                txtGrayM = TestItemTable(j).s25     'UpperGrayZone
                txtUpLowF = TestItemTable(j).s31
                txtGrayF = TestItemTable(j).s33
            Next
        '<-- Flag 4 ----- LowerGrayZone/LowerLimit-------
        ElseIf TestItemTable(1).s13 = "4" Then
            optNoOpt(3).Value = True
            fraUpLow.Visible = True
            fraRefNum.Visible = False
            txtRefLoM = ""
            txtLowerGrayM = ""
            txtRefHiM = ""
            txtUpperGrayM = ""
            txtRefLoF = ""
            txtLowerGrayF = ""
            txtRefHiF = ""
            txtUpperGrayF = ""
            
            For j = 1 To i
                txtUpLowM = TestItemTable(j).s24    'LowerLimit
                txtGrayM = TestItemTable(j).s26     'LowerGrayZone
                txtUpLowF = TestItemTable(j).s32
                txtGrayF = TestItemTable(j).s34
            Next
        End If
                
        lblOtherFlag.Visible = False
        txtOtherFlag.Visible = False
        txtOtherFlag = ""
        
        txtDelta.Locked = True
        
        If TestItemTable(1).s15 = "0" Then      'Delta 0,1,2
            optDelta(1).Value = True
        ElseIf TestItemTable(1).s15 = "1" Then
            txtDelta.Locked = False
            txtDelta = TestItemTable(1).s16
            optDelta(2).Value = True
        ElseIf TestItemTable(1).s15 = "2" Then
            txtDelta.Locked = False
            txtDelta = TestItemTable(1).s16
            optDelta(3).Value = True
        End If
        
        If TestItemTable(1).s18 = "0" Then  '��������
            optJudge(1).Value = True
            
        ElseIf TestItemTable(1).s18 = "1" Then '���� Low/High
            optJudge(2).Value = True
            
        ElseIf TestItemTable(1).s18 = "2" Then  '���� Neg/Pos
            optJudge(3).Value = True
            
        ElseIf TestItemTable(1).s18 = "3" Then  '���� OtherFlag
            optJudge(4).Value = True
            lblOtherFlag.Visible = True
            txtOtherFlag.Visible = True
            txtOtherFlag = TestItemTable(1).s27
        End If
    End If
    
End Sub

Private Sub DisplaySearchResult()
    
End Sub
Private Sub InsertItem(ByVal SubMCd As String, _
                       ByVal RefGbn As String, ByVal PanicGbn As String, _
                       ByVal DeltaGbn As String, _
                       ByVal PanJungGbn As String)
    
    On Error GoTo ErrHandler
    
    Dim iLowHighYes%
    Dim i%
    Dim sTestGbn$
    Dim vTestCd, vTestGbn
    Dim iDisplayRow%
    Dim CTestItem As DCB0101
    
    Set CTestItem = New DCB0101
    
    With CTestItem
        .TotField01 = Left$(txtSlip, 1)
        .TotField02 = Right$(txtSlip, 2)
        .TotField03 = txtSpecimen
        .TotField04 = txtTestSeq
        .TotField05 = SubMCd
        .TotField06 = "0"       'FLAGYN --> DEFAULT '0'
        .TotField07 = txtTestNm
        .TotField08 = txtPrintNm
        .TotField09 = txtUnit
        .TotField10 = txtDisplayOrd
        .TotField11 = txtPrintOrd
        
        If txtPrintSpace = "Y" Then
            .TotField12 = "1"
        ElseIf txtPrintSpace = "N" Then
            .TotField12 = "0"
        End If
        
        .TotField13 = RefGbn
        .TotField14 = PanicGbn
        .TotField15 = DeltaGbn
        
        If optDelta(1).Value = False Then
            .TotField16 = txtDelta
        Else
            .TotField16 = ""
        End If
    
        .TotField17 = ""
        .TotField18 = PanJungGbn
        
        If optRefGbn(3).Value = True And optNoOpt(1).Value = True Then
            .TotField19 = txtPanicLo        '�дл�� (����ġ ����(Low - High))
            .TotField20 = txtPanicHi
        Else                                '�̿��� ��� ���
            .TotField19 = ""
            .TotField20 = ""
        End If
        
        If RefGbn = 2 Then              '����ġ ���� ( LOW - HIGH )
            If Len(txtRefLoM) = 0 Or Len(txtRefHiM) = 0 Or _
                Len(txtRefLoF) = 0 Or Len(txtRefHiF) = 0 Then
                    
                MsgBox "����ġ ������ ��, �� �� �� �Ϻθ� �Է����� �ʾҽ��ϴ�!!"
                Exit Sub
            End If
            
            .TotField21 = txtRefLoM
            .TotField22 = txtRefHiM
            .TotField23 = ""
            .TotField24 = ""
            .TotField25 = fBlankToZero(txtUpperGrayM)
            .TotField26 = fBlankToZero(txtLowerGrayM)
            '.TotField27
            '.TotField28
            .TotField29 = txtRefLoF
            .TotField30 = txtRefHiF
            .TotField31 = ""
            .TotField32 = ""
            .TotField33 = fBlankToZero(txtUpperGrayF)
            .TotField34 = fBlankToZero(txtLowerGrayF)
            
        ElseIf RefGbn = 3 Then          '����ġ ���� ( - ����)
            If Len(txtUpLowM) = 0 Or Len(txtGrayM) = 0 Then
                MsgBox "����ġ ������ ��, �� �� �� �Ϻθ� �Է����� �ʾҽ��ϴ�!!"
                Exit Sub
            End If
            
            .TotField21 = ""
            .TotField22 = ""
            .TotField23 = txtUpLowM
            .TotField24 = ""
            .TotField25 = fBlankToZero(txtGrayM)
            .TotField26 = ""
            '.TotField27
            '.TotField28
            .TotField29 = ""
            .TotField30 = ""
            .TotField31 = txtUpLowF
            .TotField32 = ""
            .TotField33 = fBlankToZero(txtGrayF)
            .TotField34 = ""
            
        ElseIf RefGbn = 4 Then          '����ġ ���� ( ���� - )
            .TotField21 = ""
            .TotField22 = ""
            .TotField23 = ""
            .TotField24 = txtUpLowM
            .TotField25 = ""
            .TotField26 = fBlankToZero(txtGrayM)
            '.TotField27
            '.TotField28
            .TotField29 = ""
            .TotField30 = ""
            .TotField31 = ""
            .TotField32 = txtUpLowF
            .TotField33 = ""
            .TotField34 = fBlankToZero(txtGrayF)
            
        Else                            '�̿��� ��� ���
            .TotField21 = ""
            .TotField22 = ""
            .TotField23 = ""
            .TotField24 = ""
            .TotField25 = ""
            .TotField26 = ""
            '.TotField27
            '.TotField28
            .TotField29 = ""
            .TotField30 = ""
            .TotField31 = ""
            .TotField32 = ""
            .TotField33 = ""
            .TotField34 = ""
            
        End If
    
        If optJudge(4).Value = True Then
            .TotField27 = txtOtherFlag
        Else
            .TotField27 = ""
        End If
    
        If optRefGbn(2).Value = True Then
            .TotField28 = txtRefLetter
        Else
            .TotField28 = ""
        End If
        
    End With
    
    CTestItem.Add_TESTITEM "1"
    
    If CTestItem.AdoErrNum = 0 Or CTestItem.AdoErrNum = 1 Then
        ViewMsg "���������� ��ϵǾ����ϴ�..."
    Else
        ViewMsg "����۾��� �����Ͽ����ϴ�. �����޼����� �����ϼ���. - ( " & CTestItem.AdoErrNum & " )"
    End If
    

'Spread�� �ٲﳻ�� ����
'ȭ�鿡 �ݿ�
    If CTestItem.AdoErrNum = 0 Or CTestItem.AdoErrNum = 1 Then
        With spdBaseCode
            iDisplayRow = 0
            
            For i = 1 To .MaxRows
                Call .GetText(1, i, vTestCd)
                Call .GetText(2, i, vTestGbn)
                
                If optTestGbn(1).Value = True Then
                    sTestGbn = "N"
                ElseIf optTestGbn(2).Value = True Then
                    sTestGbn = "S" & txtSub
                    txtSub.Enabled = True
                End If
                
                If vTestCd = txtSlip & txtSpecimen & txtTestSeq And vTestGbn = sTestGbn Then
                    iDisplayRow = i
                    Exit For
                End If
            Next
                    
            If iDisplayRow = 0 Then
                .MaxRows = .MaxRows + 1
                iDisplayRow = .MaxRows
            Else
            End If
                        
            Call .SetText(1, iDisplayRow, txtSlip & txtSpecimen & txtTestSeq & "")
            
            If optTestGbn(1).Value = True Then
                Call .SetText(2, iDisplayRow, "N")
                optTestGbn(2).Enabled = False
            ElseIf optTestGbn(2).Value = True Then
                Call .SetText(2, iDisplayRow, "S" & txtSub & "")
                optTestGbn(1).Enabled = False
            End If
            
            Call .SetText(3, iDisplayRow, txtTestNm & "")
            lblTestNm = txtTestNm
            
            If optRefGbn(1).Value = True Then
                Call .SetText(4, iDisplayRow, "1")
            Else
                Call .SetText(4, iDisplayRow, "0")
            End If
            
            If optRefGbn(2).Value = True Then
                Call .SetText(5, iDisplayRow, txtRefLetter & "")
            End If
            
            If RefGbn = "2" Then
                Call .SetText(6, iDisplayRow, txtRefLoM & "")
                Call .SetText(7, iDisplayRow, txtRefHiM & "")
                
                Call .SetText(12, iDisplayRow, fBlankToZero(txtUpperGrayM) & "")
                Call .SetText(13, iDisplayRow, fBlankToZero(txtLowerGrayM) & "")
            
                Call .SetText(8, iDisplayRow, txtRefLoF & "")
                Call .SetText(9, iDisplayRow, txtRefHiF & "")
                
                Call .SetText(16, iDisplayRow, fBlankToZero(txtUpperGrayF) & "")
                Call .SetText(17, iDisplayRow, fBlankToZero(txtLowerGrayF) & "")
            
            ElseIf RefGbn = "3" Then
                Call .SetText(10, iDisplayRow, txtUpLowM & "")
                Call .SetText(12, iDisplayRow, fBlankToZero(txtGrayM) & "")

                Call .SetText(14, iDisplayRow, txtUpLowF & "")
                Call .SetText(16, iDisplayRow, fBlankToZero(txtGrayF) & "")

            ElseIf RefGbn = "4" Then
                Call .SetText(11, iDisplayRow, txtUpLowM & "")
                Call .SetText(13, iDisplayRow, fBlankToZero(txtGrayM) & "")
            
                Call .SetText(15, iDisplayRow, txtUpLowF & "")
                Call .SetText(17, iDisplayRow, fBlankToZero(txtGrayF) & "")
                
            End If
            
            If PanJungGbn = "3" Then
                Call .SetText(18, iDisplayRow, txtOtherFlag & "")
            End If
            
            If PanicGbn = "1" Then
                Call .SetText(19, iDisplayRow, txtPanicLo & "")
                Call .SetText(20, iDisplayRow, txtPanicHi & "")
            End If
            
            If iDisplayRow > 5 Then
                .TopRow = iDisplayRow - 5 + 1
            End If
        End With
        
        txtTestSeq.SetFocus
    End If
    
    Set CTestItem = Nothing
    
    Exit Sub
    
ErrHandler:
    If CTestItem.AdoErrNum = 0 Or CTestItem.AdoErrNum = 1 Then
    Else
       MsgBox "ErrHandler : ����۾��� �����Ͽ����ϴ�. �����޼����� �����ϼ���"
    End If
    
    Set CTestItem = Nothing
    
End Sub

Private Sub ShortKeyOrTabOrderInit()
    Me.KeyPreview = True
    
    txtSlip.TabIndex = 0
    txtSpecimen.TabIndex = 1
    txtTestSeq.TabIndex = 2
    optTestGbn(1).TabIndex = 3
    txtSub.TabIndex = 4

    txtTestNm.TabIndex = 6
    txtPrintNm.TabIndex = 7
    txtUnit.TabIndex = 8
    txtDisplayOrd.TabIndex = 9
    txtPrintOrd.TabIndex = 10
    txtPrintSpace.TabIndex = 11
    optRefGbn(1).TabIndex = 12
    optJudge(1).TabIndex = 13
    optDelta(1).TabIndex = 14
    txtDelta.TabIndex = 15
    cmdReg.TabIndex = 16
    cmdSearch.TabIndex = 17
    cmdDelete.TabIndex = 18
    cmdExit.TabIndex = 19
    
End Sub

Private Sub ValidChk()

    If LenH(txtSlip) <> 3 Then
        MsgBox "SLIP CODE�� ��Ʈ 1�ڸ��� ��Ʈ ���� 2�ڸ��� �Ǿ� �־�� �մϴ�!!"
        txtSlip.SetFocus
        iValid = -1
        Exit Sub
    End If
    
    If LenH(txtSpecimen) <> 3 Then
        MsgBox "��ü�ڵ�� 3�ڸ��� �Ǿ� �־�� �մϴ�!!"
        txtSpecimen.SetFocus
        iValid = -1
        Exit Sub
    End If
    
    If LenH(txtTestSeq) <> 3 And IsNumeric(txtTestSeq) = False Then
        MsgBox "�˻��׸������ 3�ڸ��� ���ڷ� �Ǿ�� �մϴ�!!"
        txtTestSeq.SetFocus
        iValid = -1
        Exit Sub
    End If
    
    If LenH(txtSub) <> 2 And IsNumeric(txtSub) = False And txtSub.Visible = True Then
        MsgBox "SUB �ڵ�� 2�ڸ��� ���ڷ� �Ǿ�� �մϴ�!!"
        txtSub.SetFocus
        iValid = -1
        Exit Sub
    End If
    
    If LenH(txtTestNm) > 35 Then
        MsgBox "�˻���� �ѱ� 2 Byte, ���� 1 Byte�� ����� 35 Byte �����̾�� �մϴ�!!"
        txtTestNm.SetFocus
        iValid = -1
        Exit Sub
    End If
    
    If LenH(txtPrintNm) > 25 Then
        MsgBox "�˻���� �ѱ� 2 Byte, ���� 1 Byte�� ����� 25 Byte �����̾�� �մϴ�!!"
        txtPrintNm.SetFocus
        iValid = -1
        Exit Sub
    End If
    
    If LenH(txtUnit) > 15 Then
        MsgBox "������ �ѱ� 2 Byte, ���� 1 Byte�� ����� 15 Byte �����̾�� �մϴ�!!"
        txtUnit.SetFocus
        iValid = -1
        Exit Sub
    End If
    
    If IsNumeric(txtDisplayOrd) = False Then
        MsgBox "DiplayOrder�� ���ڷ� �Ǿ�� �մϴ�!!"
        txtDisplayOrd.SetFocus
        iValid = -1
        Exit Sub
    End If
    
    If IsNumeric(txtPrintOrd) = False Then
        MsgBox "PrintOrder�� ���ڷ� �Ǿ�� �մϴ�!!"
        txtPrintOrd.SetFocus
        iValid = -1
        Exit Sub
    End If
    
    If txtPrintSpace = "Y" Or txtPrintSpace = "N" Then
    Else
        MsgBox "��½� �ڸ������� Y �Ǵ� N ���� �Ǿ�� �մϴ�!!"
        txtPrintSpace.SetFocus
        iValid = -1
        Exit Sub
    End If
    
    If LenH(txtRefLetter) > 20 Then
        MsgBox "����ġ ���ڴ� �ѱ� 2 Byte, ���� 1 Byte�� ����� 20 Byte �����̾�� �մϴ�!!"
        txtRefLetter.SetFocus
        iValid = -1
        Exit Sub
    End If
    
    If LenH(txtOtherFlag) > 20 And txtOtherFlag.Visible = True Then
        MsgBox "Other Flag�� �ѱ� 2 Byte, ���� 1 Byte�� ����� 20 Byte �����̾�� �մϴ�!!"
        txtOtherFlag.SetFocus
        iValid = -1
        Exit Sub
    End If
    
    If LenH(txtDelta) > 20 Then
        MsgBox "Delta���� �ѱ� 2 Byte, ���� 1 Byte�� ����� 20 Byte �����̾�� �մϴ�!!"
        txtDelta.SetFocus
        iValid = -1
        Exit Sub
    End If
    
    If LenH(txtRefLoM) > 20 And IsNumeric(txtRefLoM) = False Then
        MsgBox "����ġ�� 20�� ������ ���ڷ� �Ǿ�� �մϴ�!!"
        txtRefLoM.SetFocus
        iValid = -1
        Exit Sub
    End If
    
    If LenH(txtRefLoF) > 20 And IsNumeric(txtRefLoF) = False Then
        MsgBox "����ġ�� 20�� ������ ���ڷ� �Ǿ�� �մϴ�!!"
        txtRefLoF.SetFocus
        iValid = -1
        Exit Sub
    End If
    
    If LenH(txtRefHiM) > 20 And IsNumeric(txtRefHiM) = False Then
        MsgBox "����ġ�� 20�� ������ ���ڷ� �Ǿ�� �մϴ�!!"
        txtRefHiM.SetFocus
        iValid = -1
        Exit Sub
    End If
    
    If LenH(txtRefHiF) > 20 And IsNumeric(txtRefHiF) = False Then
        MsgBox "����ġ�� 20�� ������ ���ڷ� �Ǿ�� �մϴ�!!"
        txtRefHiF.SetFocus
        iValid = -1
        Exit Sub
    End If
    
    If LenH(txtPanicLo) > 20 And IsNumeric(txtPanicLo) = False Then
        MsgBox "����ġ�� 20�� ������ ���ڷ� �Ǿ�� �մϴ�!!"
        txtPanicLo.SetFocus
        iValid = -1
        Exit Sub
    End If
    
    If LenH(txtPanicHi) > 20 And IsNumeric(txtPanicHi) = False Then
        MsgBox "����ġ�� 20�� ������ ���ڷ� �Ǿ�� �մϴ�!!"
        txtPanicHi.SetFocus
        iValid = -1
        Exit Sub
    End If
    
    If LenH(txtUpLowM) > 20 And IsNumeric(txtUpLowM) = False Then
        MsgBox "����ġ�� 20�� ������ ���ڷ� �Ǿ�� �մϴ�!!"
        txtUpLowM.SetFocus
        iValid = -1
        Exit Sub
    End If
    
    If LenH(txtUpLowF) > 20 And IsNumeric(txtUpLowF) = False Then
        MsgBox "����ġ�� 20�� ������ ���ڷ� �Ǿ�� �մϴ�!!"
        txtUpLowF.SetFocus
        iValid = -1
        Exit Sub
    End If
    
    If LenH(txtGrayM) > 20 And IsNumeric(txtGrayM) = False Then
        MsgBox "����ġ�� 20�� ������ ���ڷ� �Ǿ�� �մϴ�!!"
        txtGrayM.SetFocus
        iValid = -1
        Exit Sub
    End If
    
    If LenH(txtGrayF) > 20 And IsNumeric(txtGrayF) = False Then
        MsgBox "����ġ�� 20�� ������ ���ڷ� �Ǿ�� �մϴ�!!"
        txtGrayF.SetFocus
        iValid = -1
        Exit Sub
    End If
    
    If LenH(txtUpperGrayM) > 20 And IsNumeric(txtUpperGrayM) = False Then
        MsgBox "����ġ�� 20�� ������ ���ڷ� �Ǿ�� �մϴ�!!"
        txtUpperGrayM.SetFocus
        iValid = -1
        Exit Sub
    End If
    
    If LenH(txtLowerGrayM) > 20 And IsNumeric(txtLowerGrayM) = False Then
        MsgBox "����ġ�� 20�� ������ ���ڷ� �Ǿ�� �մϴ�!!"
        txtLowerGrayM.SetFocus
        iValid = -1
        Exit Sub
    End If
    
    If LenH(txtUpperGrayF) > 20 And IsNumeric(txtUpperGrayF) = False Then
        MsgBox "����ġ�� 20�� ������ ���ڷ� �Ǿ�� �մϴ�!!"
        txtUpperGrayF.SetFocus
        iValid = -1
        Exit Sub
    End If
    
    If LenH(txtLowerGrayF) > 20 And IsNumeric(txtLowerGrayF) = False Then
        MsgBox "����ġ�� 20�� ������ ���ڷ� �Ǿ�� �մϴ�!!"
        txtLowerGrayF.SetFocus
        iValid = -1
        Exit Sub
    End If
    
    If optDelta(1).Value = False Then
        If txtDelta = "" Then
            MsgBox "DELTA ���� ������ ��Ÿ ���밪 �Ǵ� �ۼ�Ʈ �� ���� DELTA ������ �Է��ؾ��մϴ�!!"
            iValid = -1
            Exit Sub
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
    
    FSB0101.Left = 2700
    FSB0101.Top = 1400
    
    Load FSB0101
    FSB0101.Show vbModal
End Sub

Private Sub cmdButtonSpc_Click()
    Dim i%
    Dim j%
    Dim CSpecimen As DCB0101
    Dim sTot01$
    Dim sTot02$
    
    Set CSpecimen = New DCB0101
    CSpecimen.Get_SPC
    j = CSpecimen.CurItemCnt
    
    Erase gCodeHlpTable '�迭 �ʱ�ȭ
    
    ReDim gCodeHlpTable(j) As CodeTBL
    
    With CSpecimen
        sTot01 = .TotField01
        sTot02 = .TotField02
    End With
    
    Set CSpecimen = Nothing
    
    For i = 1 To j
        gCodeHlpTable(i).sSeq = Format$(i, "00000")
        gCodeHlpTable(i).sCode = GetByOne(sTot01, sTot01)
        gCodeHlpTable(i).sCodeNm = GetByOne(sTot02, sTot02)
    Next
    
    giCodeHlpCnt = j
    
    hWndCd = txtSpecimen.hwnd
    
    FSB0101.Left = 2700
    FSB0101.Top = 1750
    
    Load FSB0101
    FSB0101.Show vbModal
End Sub

Private Sub cmdDelete_Click()
    On Err GoTo ErrHandler
    
    Dim vTestCd
    Dim vTestGbn
    Dim vTestNm
    Dim CTestItem As DCB0101
    Dim sPart$, sPartGbn$, sSpecimen$, sTestSeq$, sSubM$
    Dim iRetVal%
    
    If iCurSelRow = 0 Then
        MsgBox "������ ���ϴ� �׸��� �Ʒ��� ǥ���� Ŭ���� �� ������ �Ͻʽÿ�!!"
        Exit Sub
    End If
    
    Call spdBaseCode.GetText(1, iCurSelRow, vTestCd)
    Call spdBaseCode.GetText(2, iCurSelRow, vTestGbn)
    Call spdBaseCode.GetText(3, iCurSelRow, vTestNm)
    
    sPart = Left$(CStr(vTestCd), 1)
    sPartGbn = Mid$(CStr(vTestCd), 2, 2)
    sSpecimen = Mid$(CStr(vTestCd), 4, 3)
    sTestSeq = Mid$(CStr(vTestCd), 7, 3)
        
    If Left$(vTestGbn, 1) = "N" Then
        sSubM = "NNNN"
    ElseIf Left$(vTestGbn, 1) = "S" Then
        sSubM = Right$(vTestGbn, 2) & "NN"
    End If
    
    iRetVal = MsgBox("�˻��׸��ڵ尡 " & CStr(vTestCd) & vbCrLf & "�˻��׸񱸺��� " & CStr(vTestGbn) & " �� " & CStr(vTestNm) & " ��(��) �����Ͻðڽ��ϱ�?", vbOKCancel, "�˻��׸� ���� Ȯ��")
    
    If iRetVal = 1 Then
        Set CTestItem = New DCB0101
        
        CTestItem.Delete_TESTITEM sPart, sPartGbn, sSpecimen, sTestSeq, sSubM
        
        If CTestItem.AdoErrNum = 0 Then
            With spdBaseCode
                .Row = iCurSelRow
                .Action = SS_ACTION_DELETE_ROW
                .MaxRows = .MaxRows - 1
            End With
            
            ViewMsg "�����۾��� ���������� �̷�� �����ϴ�!!"
            
            txtTestSeq.SetFocus
        End If
        
        Set CTestItem = Nothing
            
        
    Else
    End If
    
    Exit Sub
    
ErrHandler:
    Set CTestItem = Nothing
    
    Select Case Err.Number
        Case 13
            MsgBox Err.Description, vbInformation, "Ȯ��"
        Case Else
            MsgBox Err.Description, vbCritical, "����"
    End Select
    
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdReg_Click()
    Dim sSubMCd$, sRefGbn$, sDeltaGbn$, sPanicGbn$, sPanJungGbn$
    Dim i%
    
    iValid = 0
    
    Call ValidChk
    
    If iValid = -1 Then     'Validation Error
        Exit Sub
    End If
    
    If optTestGbn(1).Value = True Then
        sSubMCd = "NNNN"
    ElseIf optTestGbn(2).Value = True Then
        sSubMCd = txtSub & "NN"
    End If
    
    '����ġ������ 0, 1, 2, 3, 4�� �ټ�����
    If optRefGbn(1).Value = True Then
        sRefGbn = "0"
    ElseIf optRefGbn(2).Value = True Then
        sRefGbn = "1"
    ElseIf optRefGbn(3).Value = True Then
        If optNoOpt(1).Value = True Then
            sRefGbn = "2"
        ElseIf optNoOpt(2).Value = True Then
            sRefGbn = "3"
        ElseIf optNoOpt(3).Value = True Then
            sRefGbn = "4"
        End If
    End If
    
    For i = 1 To 4
        If optJudge(i).Value = True Then
            sPanJungGbn = CStr(i - 1)
        End If
    Next
    
    For i = 1 To 3
        If optDelta(i).Value = True Then
            sDeltaGbn = CStr(i - 1)
        End If
    Next
    
    If txtPanicLo = "" And txtPanicHi = "" Then
        sPanicGbn = "0"
    Else
        sPanicGbn = "1"
    End If
    
    Call InsertItem(sSubMCd, sRefGbn, sPanicGbn, sDeltaGbn, sPanJungGbn)
    
End Sub

Private Sub cmdSearch_Click()
    Dim CTestItem As DCB0101
    Dim i%
    
    Set CTestItem = New DCB0101
        
    If txtSlip = "" Then
        'All TestItem Search
        
        CTestItem.Get_TESTITEM 6, "", "", ""
        
        i = CTestItem.CurItemCnt
        
        If i = 0 Then
            MsgBox "���� �����ڷῡ � �׸� ��ϵǾ� ���� �ʽ��ϴ�!!"
            Set CTestItem = Nothing
            Exit Sub
        End If
    Else
        If txtSpecimen = "" Then
            'Specific Slip TestItem Search
            CTestItem.Get_TESTITEM 7, Left$(txtSlip, 1), Right$(txtSlip, 2), ""
        
            i = CTestItem.CurItemCnt
            
            If i = 0 Then
                MsgBox "���� �����ڷῡ �ش� ������ �׸��� ��ϵǾ� ���� �ʽ��ϴ�!!"
                Set CTestItem = Nothing
                Exit Sub
            End If
        Else
            If txtTestSeq = "" Then
            'Specific Slip And Specimen Search
                CTestItem.Get_TESTITEM 1, Left$(txtSlip, 1), Right$(txtSlip, 2), txtSpecimen
        
                i = CTestItem.CurItemCnt
                
                If i = 0 Then
                    MsgBox "���� �����ڷῡ �ش� ����, �ش� ��ü�� �׸��� ��ϵǾ� ���� �ʽ��ϴ�!!"
                    Set CTestItem = Nothing
                    Exit Sub
                End If
            Else
                If optTestGbn(1).Value = True Then
                    'Normal
                    CTestItem.Get_TESTITEM 9, Left$(txtSlip, 1), Right$(txtSlip, 2), txtSpecimen, txtTestSeq
        
                    i = CTestItem.CurItemCnt
                    
                    If i = 0 Then
                        MsgBox "���� �����ڷῡ �ش� ����, �ش� ��ü, �ش� ������ Normal �׸��� ��ϵǾ� ���� �ʽ��ϴ�!!"
                        Call DisplayInit(2)
                        Set CTestItem = Nothing
                        Exit Sub
                    End If
                ElseIf optTestGbn(2).Value = True Then  'SuB
                    If txtSub = "" Then
                        CTestItem.Get_TESTITEM 10, Left$(txtSlip, 1), Right$(txtSlip, 2), txtSpecimen, txtTestSeq
        
                        i = CTestItem.CurItemCnt
                        
                        If i = 0 Then
                            MsgBox "���� �����ڷῡ �ش� ����, �ش� ��ü, �ش� ������ SUB �׸��� ��ϵǾ� ���� �ʽ��ϴ�!!"
                            Set CTestItem = Nothing
                            Exit Sub
                        End If
                    Else
                        CTestItem.Get_TESTITEM 4, Left$(txtSlip, 1), Right$(txtSlip, 2), txtSpecimen, txtTestSeq, txtSub
        
                        i = CTestItem.CurItemCnt
                        
                        If i = 0 Then
                            MsgBox "���� �����ڷῡ �ش� ����, �ش� ��ü, �ش� ������ SUB �׸��� ��ϵǾ� ���� �ʽ��ϴ�!!"
                            Set CTestItem = Nothing
                            Exit Sub
                        End If
                    End If
                End If
            
            End If
        End If
    End If
    
    'DisplaySearchResult
    Dim j%
    Dim s01$, s02$, s03$, s04$, s05$, s06$, s07$, s08$, s09$, s10$
    Dim s11$, s12$, s13$, s14$, s15$, s16$, s17$, s18$, s19$, s20$
    Dim s21$, s22$, s23$, s24$, s25$, s26$, s27$, s28$, s29$, s30$
    Dim s31$, s32$, s33$, s34$, s35$
    Dim vTmp, vTmp2
    
    s01 = CTestItem.TotField01: s02 = CTestItem.TotField02:  s03 = CTestItem.TotField03
    s04 = CTestItem.TotField04: s05 = CTestItem.TotField05:  s06 = CTestItem.TotField06
    s07 = CTestItem.TotField07: s08 = CTestItem.TotField08:  s09 = CTestItem.TotField09
    s10 = CTestItem.TotField10: s11 = CTestItem.TotField11:  s12 = CTestItem.TotField12
    s13 = CTestItem.TotField13: s14 = CTestItem.TotField14:  s15 = CTestItem.TotField15
    s16 = CTestItem.TotField16: s17 = CTestItem.TotField17:  s18 = CTestItem.TotField18
    s19 = CTestItem.TotField19: s20 = CTestItem.TotField20:  s21 = CTestItem.TotField21
    s22 = CTestItem.TotField22: s23 = CTestItem.TotField23:  s24 = CTestItem.TotField24
    s25 = CTestItem.TotField25: s26 = CTestItem.TotField26:  s27 = CTestItem.TotField27
    s28 = CTestItem.TotField28: s29 = CTestItem.TotField29:  s30 = CTestItem.TotField30
    s31 = CTestItem.TotField31: s32 = CTestItem.TotField32:  s33 = CTestItem.TotField33
    s34 = CTestItem.TotField34
    
    Erase TestItemTable
    
    ReDim TestItemTable(i) As TestItemTBL
    
    spdBaseCode.MaxRows = 0
    
    For j = 1 To i
    
        TestItemTable(j).s01 = GetByOne(s01, s01):        TestItemTable(j).s02 = GetByOne(s02, s02)
        TestItemTable(j).s03 = GetByOne(s03, s03):        TestItemTable(j).s04 = GetByOne(s04, s04)
        TestItemTable(j).s05 = GetByOne(s05, s05):        TestItemTable(j).s06 = GetByOne(s06, s06)
        TestItemTable(j).s07 = GetByOne(s07, s07):        TestItemTable(j).s08 = GetByOne(s08, s08)
        TestItemTable(j).s09 = GetByOne(s09, s09):        TestItemTable(j).s10 = GetByOne(s10, s10)
        TestItemTable(j).s11 = GetByOne(s11, s11):        TestItemTable(j).s12 = GetByOne(s12, s12)
        TestItemTable(j).s13 = GetByOne(s13, s13):        TestItemTable(j).s14 = GetByOne(s14, s14)
        TestItemTable(j).s15 = GetByOne(s15, s15):        TestItemTable(j).s16 = GetByOne(s16, s16)
        TestItemTable(j).s17 = GetByOne(s17, s17):        TestItemTable(j).s18 = GetByOne(s18, s18)
        TestItemTable(j).s19 = GetByOne(s19, s19):        TestItemTable(j).s20 = GetByOne(s20, s20)
        TestItemTable(j).s21 = GetByOne(s21, s21):        TestItemTable(j).s22 = GetByOne(s22, s22)
        TestItemTable(j).s23 = GetByOne(s23, s23):        TestItemTable(j).s24 = GetByOne(s24, s24)
        TestItemTable(j).s25 = GetByOne(s25, s25):        TestItemTable(j).s26 = GetByOne(s26, s26)
        TestItemTable(j).s27 = GetByOne(s27, s27):        TestItemTable(j).s28 = GetByOne(s28, s28)
        TestItemTable(j).s29 = GetByOne(s29, s29):        TestItemTable(j).s30 = GetByOne(s30, s30)
        TestItemTable(j).s31 = GetByOne(s31, s31):        TestItemTable(j).s32 = GetByOne(s32, s32)
        TestItemTable(j).s33 = GetByOne(s33, s33):        TestItemTable(j).s34 = GetByOne(s34, s34)
        
        With spdBaseCode
            
            .MaxRows = .MaxRows + 1
            
            Call .SetText(1, .MaxRows, TestItemTable(j).s01 & TestItemTable(j).s02 & TestItemTable(j).s03 & TestItemTable(j).s04 & "")
        
            If TestItemTable(j).s05 = "NNNN" Then   'Normal
                Call .SetText(2, .MaxRows, "N")
            ElseIf IsNumeric(Left$(TestItemTable(j).s05, 2)) = True Then   'sub Yes
                Call .SetText(2, .MaxRows, "S" & Left$(TestItemTable(j).s05, 2) & "")
            End If
            
            Call .SetText(3, .MaxRows, TestItemTable(j).s07 & "")      '�˻��
            
            If TestItemTable(j).s13 = "0" Then    '<-------------- ����ġ ����
                '����ġ����üũ
                Call .SetText(4, .MaxRows, "1")
                
            ElseIf TestItemTable(j).s13 = "1" Then    '<----------- ����ġ ����
                                    
                '����ġ ������ ���븸 ȭ�� ���
                Call .SetText(5, .MaxRows, TestItemTable(j).s28 & "")
                
            ElseIf TestItemTable(j).s13 = "2" Then    '<-------------------- ����ġ ���� - LowerGray/Low - High/UpperGray
                    
                Call .SetText(6, .MaxRows, TestItemTable(j).s21 & "")      'LowM
                Call .SetText(7, .MaxRows, TestItemTable(j).s22 & "")      'HighM
                Call .SetText(8, .MaxRows, TestItemTable(j).s29 & "")      'LowF
                Call .SetText(9, .MaxRows, TestItemTable(j).s30 & "")      'HighF
                                        
                Call .SetText(12, .MaxRows, TestItemTable(j).s25 & "")      'UpperGrayM
                Call .SetText(13, .MaxRows, TestItemTable(j).s26 & "")      'LowerGrayM
                Call .SetText(16, .MaxRows, TestItemTable(j).s33 & "")      'UpperGrayF
                Call .SetText(17, .MaxRows, TestItemTable(j).s34 & "")      'LowerGrayF
                
                'Low - High ������ Panic ����
                If TestItemTable(j).s14 = "1" Then    'Panic ����
                    Call .SetText(19, .MaxRows, TestItemTable(j).s19 & "")
                    Call .SetText(20, .MaxRows, TestItemTable(j).s20 & "")
                End If
                
            ElseIf TestItemTable(j).s13 = "3" Then  '<--------------------- ����ġ ���� - UpperLimit / UpperGrayZone
            
                Call .SetText(10, .MaxRows, TestItemTable(j).s23 & "")     'UpperLimitM
                Call .SetText(12, .MaxRows, TestItemTable(j).s25 & "")     'UpperGrayZoneM
                Call .SetText(14, .MaxRows, TestItemTable(j).s31 & "")     'UpperLimitF
                Call .SetText(16, .MaxRows, TestItemTable(j).s33 & "")     'UpperGrayZoneF
            
            ElseIf TestItemTable(j).s13 = "4" Then      '<-------------------- ����ġ ���� - LowerGrayZone / LowerLimit
                    
                Call .SetText(11, .MaxRows, TestItemTable(j).s24 & "")      'LowerLimitM
                Call .SetText(13, .MaxRows, TestItemTable(j).s26 & "")      'LowerGrayZoneM
                Call .SetText(15, .MaxRows, TestItemTable(j).s32 & "")      'LowerLimitF
                Call .SetText(17, .MaxRows, TestItemTable(j).s34 & "")      'LowerGrayZoneF
            
            End If
            
            If TestItemTable(j).s27 <> "" Then
                Call .SetText(18, .MaxRows, TestItemTable(j).s27 & "")
            End If
                
        End With
    Next
    '<-------------------
    
    Set CTestItem = Nothing
    
End Sub

Private Sub cmdSub_Click()
    If optTestGbn(2).Value = True Then
        If txtTestSeq = "" Then
            ViewMsg "�˻��׸� ������ �Է��� �� �˻��׸� ������ �����Ͻʽÿ�!!"
            txtTestSeq.SetFocus
            Exit Sub
        Else
            If pnlSub.Visible = False Then
                pnlSub.Visible = True
                txtSub.Visible = True
                
                Call CompareTestNm(3)
                
                If lblTestNm = "" Then
                    txtSub.Enabled = False
                Else
                    txtSub.Enabled = True
                    txtSub.SetFocus
                End If
            Else
                If lblTestNm = "" Then
                    txtSub.Enabled = True
                Else
                    txtSub.Enabled = True
                    txtSub.SetFocus
                End If
            
            End If
        End If
    Else
        MsgBox "�� �׸��� SUB Ÿ���� �׸����� �������� �ʾҽ��ϴ�!!"
    End If
End Sub

Private Sub Form_Activate()
    If txtSlip = "" Then
        txtSlip.SetFocus
    ElseIf txtSpecimen = "" Then
        txtSpecimen.SetFocus
    ElseIf txtTestSeq = "" Then
        txtTestSeq.SetFocus
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        'Case vbKeyF1:        Call cmdButtonPart_Click
        Case vbKeyF2:        Call cmdReg_Click
        Case vbKeyF3:        Call cmdSearch_Click
        Case vbKeyF4:        Call cmdDelete_Click
        Case vbKeyEscape:    Call cmdExit_Click
    End Select
End Sub

Private Sub Form_Load()
    
    If Me.StartUpPosition = 2 Then
    Else
        Me.Left = 0
        Me.Top = 0
        Me.Width = 11900
        Me.Height = 7900
    End If
    
    iSpdClick = 0
    iDefaultView = 0
    
    Call DisplayInit(1)
    Call ShortKeyOrTabOrderInit
    Call DefaultItemDisplay
    Call BaseCodeInit
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call InitRegCurFrmTitle
    ViewMsg ""
End Sub

Private Sub optDelta_Click(Index As Integer)
    If Index = 1 Then
        txtDelta.Locked = True
    Else
        txtDelta.Locked = False
        If optRefGbn(3).Value = False Then
            MsgBox "Delta �ɼ��� ����ġ ������ ������ ���� �����մϴ�!!"
            optDelta(1).Value = True
            optDelta(1).SetFocus
        End If
    End If
End Sub

Private Sub optDelta_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If optDelta(1).Value = True Then
        Else
            txtDelta.SetFocus
        End If
        
        KeyAscii = 0
    End If
End Sub

Private Sub optJudge_Click(Index As Integer)
    On Error GoTo ErrHandler
    
    lblOtherFlag.Visible = False
    txtOtherFlag.Visible = False
    txtOtherFlag = ""
    
    If Index = 1 Then
    ElseIf Index = 2 Then
        If optRefGbn(3).Value = False Then
            MsgBox "Low/High �ɼ��� ����ġ ������ ������ ���� �����մϴ�!!" & vbCrLf & _
                "���� ������ ���� �Ǵ� Other Flag�� �����Ͽ��� �մϴ�!!"
            optJudge(1).Value = True
            optJudge(1).SetFocus
        End If
    ElseIf Index = 3 Then
        If optRefGbn(3).Value = False Then
            MsgBox "Low/High �ɼ��� ����ġ ������ ������ ���� �����մϴ�!!" & vbCrLf & _
                "���� ������ ���� �Ǵ� Other Flag�� �����Ͽ��� �մϴ�!!"
            optJudge(1).Value = True
            optJudge(1).SetFocus
        End If
    ElseIf Index = 4 Then
        lblOtherFlag.Visible = True
        txtOtherFlag.Visible = True
    End If
    
ErrHandler:
    
End Sub

Private Sub optJudge_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If optRefGbn(1).Value Then
            optDelta(1).SetFocus
        ElseIf optRefGbn(2).Value = True Then
            txtRefLetter.SetFocus
        ElseIf optRefGbn(3).Value = True Then
            ioptNoOptFocus = 1
            optNoOpt(1).SetFocus
        End If
        
        KeyAscii = 0
    End If
End Sub

Private Sub optNoOpt_Click(Index As Integer)
    If Index = 1 Then
        fraUpLow.Visible = False
        txtUpLowM = ""
        txtGrayM = ""
        txtUpLowF = ""
        txtGrayF = ""
        
        fraRefNum.Visible = True
    ElseIf Index = 2 Then
        fraRefNum.Visible = False
        txtRefLoM = ""
        txtLowerGrayM = ""
        txtRefHiM = ""
        txtUpperGrayM = ""
        txtRefLoF = ""
        txtLowerGrayF = ""
        txtRefHiF = ""
        txtUpperGrayF = ""
        fraUpLow.Visible = True
        
        lblUpLow(0).Caption = "����"
        lblUpLow(1).Caption = "����"
        
        ioptNoOptFocus = 1
    ElseIf Index = 3 Then
        fraRefNum.Visible = False
        txtRefLoM = ""
        txtLowerGrayM = ""
        txtRefHiM = ""
        txtUpperGrayM = ""
        txtRefLoF = ""
        txtLowerGrayF = ""
        txtRefHiF = ""
        txtUpperGrayF = ""
        
        fraUpLow.Visible = True
        
        lblUpLow(0).Caption = "����"
        lblUpLow(1).Caption = "����"
        
        ioptNoOptFocus = 1
    End If
End Sub

Private Sub optNoOpt_GotFocus(Index As Integer)
    If Index = 1 And ioptNoOptFocus = 0 Then
        optRefGbn(3).SetFocus
    End If
End Sub

Private Sub optNoOpt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Index = 1 Then
            txtRefLoM.SetFocus
        Else
            txtUpLowM.SetFocus
        End If
        KeyAscii = 0
    End If
End Sub

Private Sub optRefGbn_Click(Index As Integer)
    If Index = 1 Then
        optJudge(1).Value = True
        optJudge(2).Enabled = False
        optJudge(3).Enabled = False
        optJudge(4).Enabled = False
        txtRefLetter.Locked = True
        pnlNoOpt.Visible = False
        fraUpLow.Visible = False
        txtUpLowM = ""
        txtGrayM = ""
        txtUpLowF = ""
        txtGrayF = ""
        
        fraRefNum.Visible = False
        txtRefLoM = ""
        txtLowerGrayM = ""
        txtRefHiM = ""
        txtUpperGrayM = ""
        txtRefLoF = ""
        txtLowerGrayF = ""
        txtRefHiF = ""
        txtUpperGrayF = ""
        
        lblOtherFlag.Visible = False
        txtOtherFlag.Visible = False
        txtOtherFlag = "'"
        
'''        If optDelta(1).Value = True Then
'''            optDelta(1).SetFocus
'''        ElseIf optDelta(2).Value = True Then
'''            optDelta(2).SetFocus
'''        ElseIf optDelta(3).Value = True Then
'''            optDelta(3).SetFocus
'''        End If
    ElseIf Index = 2 Then
        txtRefLetter.Locked = False
        pnlNoOpt.Visible = False
        fraRefNum.Visible = False
        txtRefLoM = ""
        txtLowerGrayM = ""
        txtRefHiM = ""
        txtUpperGrayM = ""
        txtRefLoF = ""
        txtLowerGrayF = ""
        txtRefHiF = ""
        txtUpperGrayF = ""
        
        fraUpLow.Visible = False
        txtUpLowM = ""
        txtGrayM = ""
        txtUpLowF = ""
        txtGrayF = ""
        
        optJudge(1).Enabled = True
        optJudge(2).Enabled = True
        optJudge(3).Enabled = True
        optJudge(4).Enabled = True
        
'''        If optJudge(1).Value = True Then
'''            optJudge(1).SetFocus
'''        ElseIf optJudge(2).Value = True Then
'''            optJudge(2).SetFocus
'''        ElseIf optJudge(3).Value = True Then
'''            optJudge(3).SetFocus
'''        ElseIf optJudge(4).Value = True Then
'''            optJudge(4).SetFocus
'''        Else
'''            optJudge(1).SetFocus
'''        End If
    ElseIf Index = 3 Then
        txtRefLetter.Locked = True
        pnlNoOpt.Visible = True
        fraRefNum.Visible = True
        
        optJudge(1).Enabled = True
        optJudge(2).Enabled = True
        optJudge(3).Enabled = True
        optJudge(4).Enabled = True
        
        ioptNoOptFocus = 0
        optNoOpt(1).SetFocus
    End If
End Sub

Private Sub optRefGbn_KeyPress(Index As Integer, KeyAscii As Integer)
      
    If KeyAscii = 13 Then
        If optJudge(1).Value = True Then
            optJudge(1).SetFocus
        ElseIf optJudge(2).Value = True Then
            optJudge(2).SetFocus
        ElseIf optJudge(3).Value = True Then
            optJudge(3).SetFocus
        ElseIf optJudge(4).Value = True Then
            optJudge(4).SetFocus
        End If
        KeyAscii = 0
    End If
End Sub

Private Sub optTestGbn_Click(Index As Integer)
    On Error GoTo ErrHandler
    
    If Index = 1 Then
        pnlSub.Visible = False
        txtSub.Visible = False
        txtSub = ""
        'pnlMulti.Visible = False
        'txtMulti.Visible = False
        'cmdCal.Visible = False
        
        If txtTestSeq = "" Then
        Else
            txtTestNm.SetFocus
        End If
    ElseIf Index = 2 Then
        'pnlSub.Visible = False
        'txtSub.Visible = False
        'pnlMulti.Visible = False
        'txtMulti.Visible = False
        'cmdCal.Visible = False
        
        Call cmdSub_Click
        
    ElseIf Index = 3 Then
        pnlSub.Visible = False
        txtSub.Visible = False
        txtSub = ""
        'pnlMulti.Visible = True
        'txtMulti.Visible = True
        'cmdCal.Visible = False
        
    ElseIf Index = 4 Then
        pnlSub.Visible = False
        txtSub.Visible = False
        txtSub = ""
        'pnlMulti.Visible = False
        'txtMulti.Visible = False
        'cmdCal.Visible = True
            
        'cmdCal.SetFocus
    End If
    
ErrHandler:
    
End Sub

Private Sub optTestGbn_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        If Index = 1 Then
            txtTestNm.SetFocus
        ElseIf Index = 2 Then
            txtSub.SetFocus
        End If
    End If
End Sub

Private Sub spdBaseCode_Click(ByVal Col As Long, ByVal Row As Long)
    Dim vField01
    Dim vField02
    Dim i%, j%
    Dim sField01$, sField02$, sField03$, sField04$, sField05$
    Dim CTestItem As DCB0101
    
    Dim s01$, s02$, s03$, s04$, s05$, s06$, s07$, s08$, s09$, s10$
    Dim s11$, s12$, s13$, s14$, s15$, s16$, s17$, s18$, s19$, s20$
    Dim s21$, s22$, s23$, s24$, s25$, s26$, s27$, s28$, s29$, s30$
    Dim s31$, s32$, s33$, s34$, s35$
    
    If Row = 0 Then
        Exit Sub
    End If
    
    iCurSelRow = CInt(Row)
    
    Call spdReverse(spdBaseCode, -1, -1, Row, Row, RGB(255, 230, 230), 2)
    
    Call spdBaseCode.GetText(1, Row, vField01)
    Call spdBaseCode.GetText(2, Row, vField02)
        
    sField01 = Left$(CStr(vField01), 1)
    sField02 = Mid$(CStr(vField01), 2, 2)
    sField03 = Mid$(CStr(vField01), 4, 3)
    sField04 = Mid$(CStr(vField01), 7, 3)
    
    If Left$(CStr(vField02), 1) = "N" Then
        sField05 = "NNNN"
    ElseIf Left$(CStr(vField02), 1) = "C" Then
        sField05 = "CCCC"
    ElseIf Left$(CStr(vField02), 1) = "S" Then
        sField05 = Right$(CStr(vField02), 2) & "NN"
    ElseIf Left$(CStr(vField02), 1) = "M" Then
        sField05 = "NN" & Right$(CStr(vField02), 2)
    End If

    Set CTestItem = New DCB0101
    
    CTestItem.Get_TESTITEM 2, sField01, sField02, sField03, sField04, sField05
    
    i = CTestItem.CurItemCnt
    
    If i = 0 Then
        MsgBox "���� �����ڷῡ � �׸� ��ϵǾ� ���� �ʽ��ϴ�!!"
        Set CTestItem = Nothing
        Exit Sub
    End If
    
    iSpdClick = 1
    
    s01 = CTestItem.TotField01: s02 = CTestItem.TotField02:  s03 = CTestItem.TotField03
    s04 = CTestItem.TotField04: s05 = CTestItem.TotField05:  s06 = CTestItem.TotField06
    s07 = CTestItem.TotField07: s08 = CTestItem.TotField08:  s09 = CTestItem.TotField09
    s10 = CTestItem.TotField10: s11 = CTestItem.TotField11:  s12 = CTestItem.TotField12
    s13 = CTestItem.TotField13: s14 = CTestItem.TotField14:  s15 = CTestItem.TotField15
    s16 = CTestItem.TotField16: s17 = CTestItem.TotField17:  s18 = CTestItem.TotField18
    s19 = CTestItem.TotField19: s20 = CTestItem.TotField20:  s21 = CTestItem.TotField21
    s22 = CTestItem.TotField22: s23 = CTestItem.TotField23:  s24 = CTestItem.TotField24
    s25 = CTestItem.TotField25: s26 = CTestItem.TotField26:  s27 = CTestItem.TotField27
    s28 = CTestItem.TotField28: s29 = CTestItem.TotField29:  s30 = CTestItem.TotField30
    s31 = CTestItem.TotField31: s32 = CTestItem.TotField32:  s33 = CTestItem.TotField33
    s34 = CTestItem.TotField34
    
    Set CTestItem = Nothing
    
    Erase TestItemTable
    
    ReDim TestItemTable(i) As TestItemTBL
    
    For j = 1 To i
    
        TestItemTable(j).s01 = GetByOne(s01, s01):        TestItemTable(j).s02 = GetByOne(s02, s02)
        TestItemTable(j).s03 = GetByOne(s03, s03):        TestItemTable(j).s04 = GetByOne(s04, s04)
        TestItemTable(j).s05 = GetByOne(s05, s05):        TestItemTable(j).s06 = GetByOne(s06, s06)
        TestItemTable(j).s07 = GetByOne(s07, s07):        TestItemTable(j).s08 = GetByOne(s08, s08)
        TestItemTable(j).s09 = GetByOne(s09, s09):        TestItemTable(j).s10 = GetByOne(s10, s10)
        TestItemTable(j).s11 = GetByOne(s11, s11):        TestItemTable(j).s12 = GetByOne(s12, s12)
        TestItemTable(j).s13 = GetByOne(s13, s13):        TestItemTable(j).s14 = GetByOne(s14, s14)
        TestItemTable(j).s15 = GetByOne(s15, s15):        TestItemTable(j).s16 = GetByOne(s16, s16)
        TestItemTable(j).s17 = GetByOne(s17, s17):        TestItemTable(j).s18 = GetByOne(s18, s18)
        TestItemTable(j).s19 = GetByOne(s19, s19):        TestItemTable(j).s20 = GetByOne(s20, s20)
        TestItemTable(j).s21 = GetByOne(s21, s21):        TestItemTable(j).s22 = GetByOne(s22, s22)
        TestItemTable(j).s23 = GetByOne(s23, s23):        TestItemTable(j).s24 = GetByOne(s24, s24)
        TestItemTable(j).s25 = GetByOne(s25, s25):        TestItemTable(j).s26 = GetByOne(s26, s26)
        TestItemTable(j).s27 = GetByOne(s27, s27):        TestItemTable(j).s28 = GetByOne(s28, s28)
        TestItemTable(j).s29 = GetByOne(s29, s29):        TestItemTable(j).s30 = GetByOne(s30, s30)
        TestItemTable(j).s31 = GetByOne(s31, s31):        TestItemTable(j).s32 = GetByOne(s32, s32)
        TestItemTable(j).s33 = GetByOne(s33, s33):        TestItemTable(j).s34 = GetByOne(s34, s34)
    Next
    
    Call DisplayInfoOfSpecificTestItem(i)
        
    iSpdClick = 0
    txtTestNm.SetFocus
End Sub

Private Sub txtDelta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtDisplayOrd_Click()
    Call Txt_Highlight(txtDisplayOrd)
End Sub

Private Sub txtDisplayOrd_GotFocus()
    Call Txt_Highlight(txtDisplayOrd)
    
    If txtDisplayOrd = "" Then
        Dim CTestItem As DCB0101
        
        Set CTestItem = New DCB0101
        
        CTestItem.Get_TESTITEMORD Left$(txtSlip, 1), Right$(txtSlip, 2), txtSpecimen
        
        If CTestItem.CurItemCnt = 0 Then
            txtDisplayOrd = "1"
        Else
            txtDisplayOrd = CTestItem.TotField01
        End If
        
        Set CTestItem = Nothing
    End If
End Sub

Private Sub txtDisplayOrd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtPrintOrd.SetFocus
        txtPrintOrd = txtDisplayOrd
        KeyAscii = 0
    End If
End Sub

Private Sub txtDisplayOrd_Validate(Cancel As Boolean)
    txtPrintOrd = txtDisplayOrd
End Sub

Private Sub txtGrayF_Click()
    Call Txt_Highlight(txtGrayF)
End Sub

Private Sub txtGrayF_GotFocus()
    Call Txt_Highlight(txtGrayF)
End Sub

Private Sub txtGrayF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If optDelta(1).Value = True Then
            optDelta(1).SetFocus
        ElseIf optDelta(2).Value = True Then
            optDelta(2).SetFocus
        ElseIf optDelta(3).Value = True Then
            optDelta(3).SetFocus
        End If
    End If
End Sub

Private Sub txtGrayM_Click()
    Call Txt_Highlight(txtGrayM)
End Sub

Private Sub txtGrayM_GotFocus()
    Call Txt_Highlight(txtGrayM)
End Sub

Private Sub txtGrayM_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtUpLowF.SetFocus
    End If
End Sub

Private Sub txtLowerGrayF_Click()
    Call Txt_Highlight(txtLowerGrayF)
End Sub

Private Sub txtLowerGrayF_GotFocus()
    Call Txt_Highlight(txtLowerGrayF)
End Sub

Private Sub txtLowerGrayF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtRefHiF.SetFocus
        KeyAscii = 0
    End If
End Sub

Private Sub txtLowerGrayM_Click()
    Call Txt_Highlight(txtLowerGrayM)
End Sub

Private Sub txtLowerGrayM_GotFocus()
    Call Txt_Highlight(txtLowerGrayM)
End Sub

Private Sub txtLowerGrayM_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtRefHiM.SetFocus
        KeyAscii = 0
    End If
End Sub

Private Sub txtPanicHi_Click()
    Call Txt_Highlight(txtPanicHi)
End Sub

Private Sub txtPanicHi_GotFocus()
    Call Txt_Highlight(txtPanicHi)
End Sub

Private Sub txtPanicHi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If optDelta(1).Value = True Then
            optDelta(1).SetFocus
        ElseIf optDelta(2).Value = True Then
            optDelta(2).SetFocus
        ElseIf optDelta(3).Value = True Then
            optDelta(3).SetFocus
        End If
    End If
End Sub

Private Sub txtPanicLo_Click()
    Call Txt_Highlight(txtPanicLo)
End Sub

Private Sub txtPanicLo_GotFocus()
    Call Txt_Highlight(txtPanicLo)
End Sub

Private Sub txtPanicLo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtPanicHi.SetFocus
        KeyAscii = 0
    End If
End Sub

Private Sub txtPrintNm_Click()
    Call Txt_Highlight(txtPrintNm)
End Sub

Private Sub txtPrintNm_GotFocus()
    Call Txt_Highlight(txtPrintNm)
End Sub

Private Sub txtPrintNm_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtUnit.SetFocus
        KeyAscii = 0
    End If
End Sub

Private Sub txtPrintOrd_Click()
    Call Txt_Highlight(txtPrintOrd)
End Sub

Private Sub txtPrintOrd_GotFocus()
    Call Txt_Highlight(txtPrintOrd)
End Sub

Private Sub txtPrintOrd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtPrintSpace.SetFocus
        KeyAscii = 0
    End If
End Sub

Private Sub txtPrintSpace_Change()
    On Error GoTo ErrHandler
    
    If Len(txtPrintSpace) = 1 Then
        If txtPrintSpace = "Y" Or txtPrintSpace = "N" Then
            If optRefGbn(1).Value = True Then
                optRefGbn(1).SetFocus
            ElseIf optRefGbn(2).Value = True Then
                optRefGbn(2).SetFocus
            ElseIf optRefGbn(3).Value = True Then
                optRefGbn(3).SetFocus
            Else
                optRefGbn(1).SetFocus
            End If
        Else
            MsgBox "��½� �ڸ������� Y (Yes) �� N (No) �θ� ������ �����մϴ�!!"
            txtPrintSpace.SetFocus
        End If
    End If
    
ErrHandler:
End Sub

Private Sub txtPrintSpace_Click()
    Call Txt_Highlight(txtPrintSpace)
End Sub

Private Sub txtPrintSpace_GotFocus()
    Call Txt_Highlight(txtPrintSpace)
End Sub

Private Sub txtPrintSpace_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txtPrintSpace = "Y" Or txtPrintSpace = "N" Then
            If optRefGbn(1).Value = True Then
                optRefGbn(1).SetFocus
            ElseIf optRefGbn(2).Value = True Then
                optRefGbn(2).SetFocus
            ElseIf optRefGbn(3).Value = True Then
                optRefGbn(3).SetFocus
            Else
                optRefGbn(1).SetFocus
            End If
        Else
            MsgBox "��½� �ڸ������� Y (Yes) �� N (No) �θ� ������ �����մϴ�!!"
            txtPrintSpace.SetFocus
        End If
    End If
End Sub

Private Sub txtRefHiF_Click()
    Call Txt_Highlight(txtRefHiF)
End Sub

Private Sub txtRefHiF_GotFocus()
    Call Txt_Highlight(txtRefHiF)
End Sub

Private Sub txtRefHiF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtUpperGrayF.SetFocus
        KeyAscii = 0
    End If
End Sub

Private Sub txtRefHiM_Click()
    Call Txt_Highlight(txtRefHiM)
End Sub

Private Sub txtRefHiM_GotFocus()
    Call Txt_Highlight(txtRefHiM)
End Sub

Private Sub txtRefHiM_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtUpperGrayM.SetFocus
        KeyAscii = 0
    End If
End Sub

Private Sub txtRefLetter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        optDelta(1).SetFocus
    End If
End Sub

Private Sub txtRefLoF_Click()
    Call Txt_Highlight(txtRefLoF)
End Sub

Private Sub txtRefLoF_GotFocus()
    Call Txt_Highlight(txtRefLoF)
End Sub

Private Sub txtRefLoF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtLowerGrayF.SetFocus
        KeyAscii = 0
    End If
End Sub

Private Sub txtRefLoM_Click()
    Call Txt_Highlight(txtRefLoM)
End Sub

Private Sub txtRefLoM_GotFocus()
    Call Txt_Highlight(txtRefLoM)
End Sub

Private Sub txtRefLoM_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtLowerGrayM.SetFocus
        KeyAscii = 0
    End If
End Sub

Private Sub txtSlip_Change()
    On Error GoTo ErrHandler
    'MsgBox iDefaultView
    
    If iDefaultView = 0 Then
        Exit Sub
    End If
    
    If Len(txtSlip) = txtSlip.MaxLength Then
        If sPrevSlip = txtSlip Then
            txtSpecimen.SetFocus
        Else
            'MsgBox ""
            Call CompareSlip
            txtSpecimen.SetFocus
            
            If iSpdClick = 1 Then
                '�������带 Ŭ����
            Else
                'Ű���带 ���� �����ڵ尡 �ٲ���� ���
                txtSpecimen = ""
            End If
        End If
    ElseIf Len(txtSlip) = 0 Then
        lblSlipNm = ""
        Call DisplayInit(1)
    End If
ErrHandler:
End Sub

Private Sub txtSlip_Click()
    Call Txt_Highlight(txtSlip)
    sPrevSlip = txtSlip
End Sub

Private Sub txtSlip_GotFocus()
    Call Txt_Highlight(txtSlip)
    sPrevSlip = txtSlip
End Sub

Private Sub txtSlip_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1:        Call cmdButtonSlip_Click
    End Select
End Sub

Private Sub txtSlip_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtSpecimen.SetFocus
    End If
End Sub

Private Sub txtSlip_Validate(Cancel As Boolean)
    If Len(txtSlip) = txtSlip.MaxLength Then
        txtSpecimen.SetFocus
        
        Call CompareSlip
    End If
End Sub

Private Sub txtSpecimen_Change()
    On Error GoTo ErrHandler
    
    'MsgBox iDefaultView
    
    If iDefaultView = 0 Then
        Exit Sub
    End If
    
    If Len(txtSpecimen) = txtSpecimen.MaxLength Then
        
        If sPrevSpc = txtSpecimen Then
        Else
            Call CompareSpecimen
            
            If iSpdClick = 1 Then
                '�������� Ŭ���ÿ��� txtTestSeq�� �ʱ�ȭ ���� �ʴ´�.
            Else
                'Ű����� ��ü�ڵ带 �Է��� ��쿡�� �ʱ�ȭ
                txtTestSeq = ""
            End If
        End If
        
        txtTestSeq.SetFocus
        
    ElseIf Len(txtSpecimen) = 0 Then
        txtSpecimen = ""
        sPrevSpc = ""
        lblSpecimenNm = ""
        txtTestSeq = ""
        Call DisplayInit(2)
    End If
    
ErrHandler:
End Sub

Private Sub txtSpecimen_Click()
    Call Txt_Highlight(txtSpecimen)
    sPrevSpc = txtSpecimen
End Sub

Private Sub txtSpecimen_GotFocus()
    Call Txt_Highlight(txtSpecimen)
    sPrevSpc = txtSpecimen
End Sub

Private Sub txtSpecimen_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1:        Call cmdButtonSpc_Click
    End Select
End Sub

Private Sub txtSpecimen_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtTestSeq.SetFocus
    End If
End Sub

Private Sub txtSpecimen_LostFocus()
    If Len(txtSpecimen) < txtSpecimen.MaxLength Then
        txtSpecimen = Format$(txtSpecimen, "000")
    End If
End Sub

Private Sub txtSpecimen_Validate(Cancel As Boolean)
    If Len(txtSpecimen) = txtSpecimen.MaxLength Then
        Call CompareSpecimen
        txtTestSeq.SetFocus
    End If
End Sub

Private Sub txtSub_Change()
    If Len(txtSub) = txtSub.MaxLength Then
        
        If iSpdClick = 1 Then
            'iSpdClick = 0
        Else
            Call CompareTestNm(1)
        End If
        
        txtTestNm.SetFocus
        
    End If
End Sub

Private Sub txtSub_Click()
    Call Txt_Highlight(txtSub)
End Sub

Private Sub txtSub_GotFocus()
    Call Txt_Highlight(txtSub)
End Sub

Private Sub txtSub_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtTestNm.SetFocus
    End If
End Sub

Private Sub txtSub_LostFocus()
    If Len(txtSub) < txtSub.MaxLength Then
        txtSub = Format$(txtSub, "00")
    End If
End Sub

Private Sub txtTestNm_Click()
    Call Txt_Highlight(txtTestNm)
End Sub

Private Sub txtTestNm_GotFocus()
    Call Txt_Highlight(txtTestNm)
End Sub

Private Sub txtTestNm_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtPrintNm = "" Then
            txtPrintNm = LeftH$(txtTestNm, 25)
        End If
        
        txtPrintNm.SetFocus
        KeyAscii = 0
    End If
End Sub

Private Sub txtTestNm_Validate(Cancel As Boolean)
    If txtPrintNm = "" Then
        txtPrintNm = LeftH$(txtTestNm, 25)
    End If
End Sub

Private Sub txtTestSeq_Change()
    If Len(txtTestSeq) = txtTestSeq.MaxLength Then
        If iSpdClick = 1 Then
            '�������� Ŭ����
        Else
            'Ű���� �Է��� ���� �˻��׸� ������ �ٲ���� ���
            Call CompareTestNm(2)
        End If
        
        txtTestNm.SetFocus
    ElseIf Len(txtTestSeq) = 0 Then
        Call DisplayInit(2)
        
    End If
End Sub

Private Sub txtTestSeq_Click()
    Call Txt_Highlight(txtTestSeq)
End Sub

Private Sub txtTestSeq_GotFocus()
    Call Txt_Highlight(txtTestSeq)
End Sub

Private Sub txtTestSeq_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If optTestGbn(1).Value = True Then
            optTestGbn(1).SetFocus
        ElseIf optTestGbn(2).Value = True Then
            optTestGbn(2).SetFocus
        End If
    End If
End Sub

Private Sub txtTestSeq_LostFocus()
    If Len(txtTestSeq) < txtTestSeq.MaxLength Then
        txtTestSeq = Format$(txtTestSeq, "000")
    End If
End Sub

Private Sub txtTestSeq_Validate(Cancel As Boolean)
    If Len(txtTestSeq) = txtTestSeq.MaxLength Then
        If txtSub = "" Then
            Call CompareTestNm
        End If
        txtTestNm.SetFocus
    End If
End Sub

Private Sub txtUnit_Click()
    Call Txt_Highlight(txtUnit)
End Sub

Private Sub txtUnit_GotFocus()
    Call Txt_Highlight(txtUnit)
End Sub

Private Sub txtUnit_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        
        If txtDisplayOrd = "" Then
            Dim CTestItem As DCB0101
            
            Set CTestItem = New DCB0101
            
            'Diplay Order�� ����, ��ü�� ������ ���� ���� ���� ����
            CTestItem.Get_TESTITEMORD Left$(txtSlip, 1), Right$(txtSlip, 2), txtSpecimen
            
            If CTestItem.CurItemCnt = 0 Then
                txtDisplayOrd = "1"
            Else
                txtDisplayOrd = CTestItem.TotField01
            End If
            
            Set CTestItem = Nothing
        End If
        
        txtDisplayOrd.SetFocus
        KeyAscii = 0
    End If
End Sub

Private Sub txtUpLowF_Click()
    Call Txt_Highlight(txtUpLowF)
End Sub

Private Sub txtUpLowF_GotFocus()
    Call Txt_Highlight(txtUpLowF)
End Sub

Private Sub txtUpLowF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtGrayF.SetFocus
    End If
End Sub

Private Sub txtUpLowM_Click()
    Call Txt_Highlight(txtUpLowM)
End Sub

Private Sub txtUpLowM_GotFocus()
    Call Txt_Highlight(txtUpLowM)
End Sub

Private Sub txtUpLowM_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtGrayM.SetFocus
    End If
End Sub

Private Sub txtUpperGrayF_Click()
    Call Txt_Highlight(txtUpperGrayF)
End Sub

Private Sub txtUpperGrayF_GotFocus()
    Call Txt_Highlight(txtUpperGrayF)
End Sub

Private Sub txtUpperGrayF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtPanicLo.SetFocus
        KeyAscii = 0
    End If
End Sub

Private Sub txtUpperGrayM_Click()
    Call Txt_Highlight(txtUpperGrayM)
End Sub

Private Sub txtUpperGrayM_GotFocus()
    Call Txt_Highlight(txtUpperGrayM)
End Sub

Private Sub txtUpperGrayM_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtRefLoF.SetFocus
        
        If txtRefLoF = "" And txtRefHiF = "" Then
            txtRefLoF = txtRefLoM
            txtRefHiF = txtRefHiM
            txtUpperGrayF = txtUpperGrayM
            txtLowerGrayF = txtLowerGrayM
        End If
        
        KeyAscii = 0
    End If
End Sub
