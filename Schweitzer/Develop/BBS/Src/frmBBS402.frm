VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBBS402 
   BackColor       =   &H00DBE6E6&
   Caption         =   "������ ����"
   ClientHeight    =   9015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14700
   Icon            =   "frmBBS402.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9015
   ScaleWidth      =   14700
   WindowState     =   2  '�ִ�ȭ
   Begin VB.CommandButton cmdCallAsk 
      BackColor       =   &H00F4F0F2&
      Caption         =   "�������(&N)"
      Height          =   510
      Left            =   2265
      Style           =   1  '�׷���
      TabIndex        =   14
      Tag             =   "15101"
      Top             =   7545
      Width           =   1320
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00F4F0F2&
      Caption         =   "�������"
      Height          =   510
      Left            =   6915
      Style           =   1  '�׷���
      TabIndex        =   12
      Tag             =   "15101"
      Top             =   7575
      Width           =   1320
   End
   Begin VB.CommandButton cmdNew 
      BackColor       =   &H00F4F0F2&
      Caption         =   "�ű� ����"
      Height          =   510
      Left            =   7335
      Style           =   1  '�׷���
      TabIndex        =   2
      Tag             =   "15101"
      Top             =   3135
      Width           =   1320
   End
   Begin MSComctlLib.TabStrip tabAccDt 
      Height          =   315
      Left            =   2265
      TabIndex        =   15
      Top             =   2025
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "2000-01-01"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "����(&X)"
      Height          =   510
      Left            =   10875
      Style           =   1  '�׷���
      TabIndex        =   13
      Tag             =   "128"
      Top             =   7575
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "ȭ������(&C)"
      Height          =   510
      Left            =   9555
      Style           =   1  '�׷���
      TabIndex        =   11
      Tag             =   "124"
      Top             =   7575
      Width           =   1320
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "����(&S)"
      Height          =   510
      Left            =   8235
      Style           =   1  '�׷���
      TabIndex        =   10
      Tag             =   "15101"
      Top             =   7575
      Width           =   1320
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   315
      Left            =   2280
      TabIndex        =   41
      Top             =   480
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   556
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
      Caption         =   "  �� �� �� ��"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   315
      Left            =   2280
      TabIndex        =   42
      Top             =   1695
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   556
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
      Caption         =   "  �� �� �� ��"
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   615
      Left            =   2280
      TabIndex        =   46
      Top             =   2265
      Width           =   9945
      Begin MedControls1.LisLabel lblStsNm 
         Height          =   315
         Left            =   1095
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   195
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   556
         ForeColor       =   -2147483634
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblStsCd 
         Height          =   315
         Left            =   2340
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   195
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         ForeColor       =   -2147483634
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblOkDiv1Nm 
         Height          =   315
         Left            =   3630
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   195
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   556
         ForeColor       =   -2147483634
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblOkDiv1Cd 
         Height          =   315
         Left            =   4575
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   195
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         ForeColor       =   -2147483634
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblOkDiv2Nm 
         Height          =   315
         Left            =   5895
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   195
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   556
         ForeColor       =   -2147483634
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblOkDiv2Cd 
         Height          =   315
         Left            =   6840
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   195
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         ForeColor       =   -2147483634
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblOkDiv3Nm 
         Height          =   315
         Left            =   8160
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   195
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   556
         ForeColor       =   -2147483634
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblOkDiv3Cd 
         Height          =   315
         Left            =   9120
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   195
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         ForeColor       =   -2147483634
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   315
         Index           =   6
         Left            =   90
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   195
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
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
         Caption         =   "�������"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   315
         Index           =   7
         Left            =   2640
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   195
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
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
         Caption         =   "�������"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   315
         Index           =   8
         Left            =   4890
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   195
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
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
         Caption         =   "�������"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   315
         Index           =   9
         Left            =   7155
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   195
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
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
         Caption         =   "�˻���"
         Appearance      =   0
      End
   End
   Begin VB.Frame fraDonorCd 
      BackColor       =   &H00DBE6E6&
      Height          =   1575
      Left            =   2280
      TabIndex        =   16
      Top             =   2790
      Width           =   6975
      Begin VB.OptionButton optDonorCd 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Phlebotomy"
         Height          =   435
         Index           =   4
         Left            =   5130
         Style           =   1  '�׷���
         TabIndex        =   56
         Top             =   900
         Width           =   1245
      End
      Begin VB.OptionButton optDonorCd 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Autologous"
         Height          =   435
         Index           =   2
         Left            =   2610
         Style           =   1  '�׷���
         TabIndex        =   20
         Top             =   900
         Width           =   1245
      End
      Begin VB.OptionButton optDonorCd 
         BackColor       =   &H00DBE6E6&
         Caption         =   "���� ����"
         Height          =   435
         Index           =   0
         Left            =   105
         Style           =   1  '�׷���
         TabIndex        =   19
         Top             =   900
         Width           =   1245
      End
      Begin VB.OptionButton optDonorCd 
         BackColor       =   &H00DBE6E6&
         Caption         =   "���� ����"
         Height          =   435
         Index           =   1
         Left            =   1365
         Style           =   1  '�׷���
         TabIndex        =   18
         Top             =   900
         Width           =   1245
      End
      Begin VB.OptionButton optDonorCd 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Pheresis"
         Height          =   435
         Index           =   3
         Left            =   3870
         Style           =   1  '�׷���
         TabIndex        =   17
         Top             =   900
         Width           =   1245
      End
      Begin MSComCtl2.DTPicker dtpAccDt 
         Height          =   330
         Left            =   1140
         TabIndex        =   1
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   60489731
         CurrentDate     =   36797
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   315
         Index           =   10
         Left            =   120
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   360
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
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
         Caption         =   "��������"
         Appearance      =   0
      End
   End
   Begin VB.Frame fraAcc 
      BackColor       =   &H00DBE6E6&
      Height          =   3540
      Left            =   9255
      TabIndex        =   25
      Top             =   2790
      Width           =   2955
      Begin VB.TextBox txtWeight 
         Alignment       =   1  '������ ����
         Appearance      =   0  '���
         Height          =   330
         Left            =   1140
         TabIndex        =   3
         Top             =   720
         Width           =   930
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   12
         Left            =   135
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   720
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   582
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
         Caption         =   "ü��"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   13
         Left            =   135
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   1080
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   582
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
         Caption         =   "�ƹ�"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   14
         Left            =   135
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   1455
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   582
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
         Caption         =   "ü��"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   15
         Left            =   135
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   1815
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   582
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
         Caption         =   "����"
         Appearance      =   0
      End
      Begin VB.TextBox txtHeight 
         Alignment       =   1  '������ ����
         Appearance      =   0  '���
         Height          =   330
         Left            =   1140
         TabIndex        =   6
         Top             =   1815
         Width           =   945
      End
      Begin VB.TextBox txtPulse 
         Alignment       =   1  '������ ����
         Appearance      =   0  '���
         Height          =   315
         Left            =   1140
         TabIndex        =   4
         Top             =   1080
         Width           =   930
      End
      Begin VB.TextBox txtBldPres1 
         Alignment       =   1  '������ ����
         Appearance      =   0  '���
         Height          =   315
         Left            =   1140
         TabIndex        =   7
         Top             =   2175
         Width           =   585
      End
      Begin VB.TextBox txtBldPres2 
         Alignment       =   1  '������ ����
         Appearance      =   0  '���
         Height          =   315
         Left            =   1860
         TabIndex        =   8
         Top             =   2175
         Width           =   555
      End
      Begin VB.TextBox txtBodyTemp 
         Alignment       =   1  '������ ����
         Appearance      =   0  '���
         Height          =   315
         Left            =   1140
         TabIndex        =   5
         Top             =   1455
         Width           =   930
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   16
         Left            =   135
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   2175
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   582
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
         Caption         =   "����"
         Appearance      =   0
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "Kg"
         Height          =   180
         Left            =   2130
         TabIndex        =   29
         Top             =   840
         Width           =   225
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "Cm"
         Height          =   180
         Left            =   2130
         TabIndex        =   28
         Top             =   1950
         Width           =   300
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "/Min"
         Height          =   180
         Left            =   2100
         TabIndex        =   27
         Top             =   1200
         Width           =   405
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "/"
         Height          =   180
         Left            =   1755
         TabIndex        =   26
         Top             =   2235
         Width           =   90
      End
   End
   Begin VB.Frame fraDonation 
      BackColor       =   &H00DBE6E6&
      Height          =   2055
      Left            =   2280
      TabIndex        =   21
      Top             =   4275
      Width           =   6975
      Begin VB.CommandButton cmdReserved 
         BackColor       =   &H00E0E0E0&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5835
         MousePointer    =   14  'ȭ��ǥ�� ����ǥ
         Style           =   1  '�׷���
         TabIndex        =   23
         Top             =   345
         Width           =   360
      End
      Begin VB.TextBox txtReservedID 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1125
         MaxLength       =   10
         TabIndex        =   22
         Top             =   360
         Width           =   1305
      End
      Begin MedControls1.LisLabel lblReservedNm 
         Height          =   315
         Left            =   2460
         TabIndex        =   24
         Top             =   360
         Width           =   3360
         _ExtentX        =   5927
         _ExtentY        =   556
         BackColor       =   13622494
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
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblTmpPtId 
         Height          =   315
         Left            =   120
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   360
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
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
         Caption         =   "����ȯ��"
         Appearance      =   0
      End
   End
   Begin VB.Frame fraResult 
      BackColor       =   &H00DBE6E6&
      Height          =   1215
      Left            =   2280
      TabIndex        =   30
      Top             =   6240
      Width           =   9930
      Begin VB.CheckBox chkHold 
         BackColor       =   &H00DBE6E6&
         Caption         =   "����"
         Height          =   255
         Left            =   540
         TabIndex        =   45
         Top             =   720
         Width           =   675
      End
      Begin VB.TextBox txtRmk 
         Height          =   825
         Left            =   2520
         MultiLine       =   -1  'True
         ScrollBars      =   2  '����
         TabIndex        =   9
         Top             =   240
         Width           =   6750
      End
      Begin VB.OptionButton optOk 
         BackColor       =   &H00DBF2FD&
         Caption         =   "������"
         Height          =   375
         Index           =   1
         Left            =   1380
         Style           =   1  '�׷���
         TabIndex        =   32
         Top             =   660
         Width           =   1095
      End
      Begin VB.OptionButton optOk 
         BackColor       =   &H00DBF2FD&
         Caption         =   "��   ��"
         Height          =   375
         Index           =   0
         Left            =   1380
         Style           =   1  '�׷���
         TabIndex        =   31
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "�� �������"
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
         Left            =   180
         TabIndex        =   43
         Top             =   300
         Width           =   990
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00DBE6E6&
      Height          =   975
      Left            =   2280
      TabIndex        =   33
      Top             =   720
      Width           =   9945
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   3
         Left            =   5655
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   165
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   582
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
         Caption         =   "��/����"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   4
         Left            =   5655
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   525
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   582
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
         Caption         =   "�� ������"
         Appearance      =   0
      End
      Begin VB.TextBox txtDonorNm 
         Appearance      =   0  '���
         Height          =   330
         Left            =   1050
         TabIndex        =   0
         Top             =   165
         Width           =   1515
      End
      Begin VB.CommandButton cmdNewReg 
         BackColor       =   &H00F4F0F2&
         Caption         =   "�űԵ��"
         Height          =   375
         Left            =   1050
         Style           =   1  '�׷���
         TabIndex        =   55
         TabStop         =   0   'False
         Tag             =   "15101"
         Top             =   510
         Width           =   1500
      End
      Begin MedControls1.LisLabel lblDOB 
         Height          =   330
         Left            =   4290
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   165
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   582
         BackColor       =   13622494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "2001-01-01"
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblSex 
         Height          =   330
         Left            =   6645
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   165
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   582
         BackColor       =   13622494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "M/100"
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblABO 
         Height          =   330
         Left            =   8955
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   165
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   582
         BackColor       =   13622494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblCnt 
         Height          =   330
         Left            =   4290
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   525
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   582
         BackColor       =   13622494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblTotVol 
         Height          =   330
         Left            =   6645
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   525
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   582
         BackColor       =   13622494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblDonorID 
         Height          =   315
         Left            =   1050
         TabIndex        =   40
         Top             =   540
         Visible         =   0   'False
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   556
         BackColor       =   13622494
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
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblSSN 
         Height          =   315
         Left            =   1815
         TabIndex        =   44
         Top             =   540
         Visible         =   0   'False
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   556
         BackColor       =   13622494
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
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   315
         Index           =   0
         Left            =   60
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   165
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
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
         Caption         =   "��   ��"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   1
         Left            =   3300
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   165
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   582
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
         Caption         =   "�������"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   2
         Left            =   3300
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   525
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   582
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
         Caption         =   "����Ƚ��"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   5
         Left            =   7965
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   165
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   582
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
         Caption         =   "������"
         Appearance      =   0
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "cc"
         Height          =   180
         Left            =   7605
         TabIndex        =   39
         Top             =   660
         Width           =   210
      End
   End
End
Attribute VB_Name = "frmBBS402"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'
'������ ����

Private objMySQL As clsBBSSQLStatement
Private WithEvents GetPtInfo As frmPtInfo
Attribute GetPtInfo.VB_VarHelpID = -1
Private AccDtform As Long
'2001-11-27�߰�
Private strSaveDonorId As String
Private strSaveDonorNm As String
Private blnClearFg As Boolean
Private blnDonorFind As Boolean

'2001-11-27 �߰�
Private Sub cmdCallAsk_Click()
    frmBBS403.Show
    frmBBS403.txtDonorNm.Text = strSaveDonorNm
    Call frmBBS403.CallDonorNmLostFocus
End Sub

Private Sub cmdCancel_Click()
    Dim DonorId As String
    Dim accdt As String
    Dim objSql As clsBBSSQLStatement
    
    If tabAccDt.SelectedItem.Index > 1 Then
        '���� �������ڰ� �ƴϴ�. ������� �� �� ����.
        MsgBox "������Ҹ� �� �� �����ϴ�.", vbCritical, Me.Caption
        Exit Sub
    Else
        '�������� ���¸� �ľ��Ѵ�.
        If lblStsCd.Caption > DonorStatus.stsAskSave Then
            MsgBox "������Ҹ� �� �� �����ϴ�.", vbCritical, Me.Caption
            Exit Sub
        End If
    End If
    
    DonorId = lblDonorId.Caption
    accdt = Format(tabAccDt.SelectedItem.Caption, PRESENTDATE_FORMAT)
    
    Set objSql = New clsBBSSQLStatement
'    objSql.setDbConn DBConn
    If objSql.SetDonorStatus(DonorId, accdt, DonorStatus.stsAccessSave) = True Then
        txtDonorNm = ""
        FormInitialize
    End If
    Set objSql = Nothing
End Sub

Private Sub Form_Activate()
    medMain.lblSubMenu.Caption = Me.Caption
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objMySQL = Nothing
    Set GetPtInfo = New frmPtInfo
End Sub

Private Sub GetPtInfo_Click(ByVal isSELECT As Boolean, ByVal ptInfo As S2BBS_Library.clsPtInformation)
    If isSELECT = False Then Exit Sub
    
    txtReservedID.Text = "": lblReservedNm.Caption = ""
    
    With ptInfo
            txtReservedID = .PtId
            lblReservedNm.Caption = .ptnm
    End With
End Sub

Private Sub cmdClear_Click()
    txtDonorNm = ""
    Call FormInitialize
End Sub

Private Sub cmdExit_Click()
    Set objMySQL = Nothing
    Unload Me
    Set frmBBS402 = Nothing
End Sub

Private Sub cmdNew_Click()
    If tabAccDt.Tabs.Count <> 0 Then
        If Format(tabAccDt.Tabs.Item(1).Caption, PRESENTDATE_FORMAT) = Format(GetSystemDate, PRESENTDATE_FORMAT) Then
            MsgBox "�ű������� �� �� �����ϴ�.", vbCritical, Me.Caption
            Exit Sub
        End If
    End If

    dtpAccDt.Enabled = True
    
    tabAccDt.Visible = True
    tabAccDt.Tabs.Add 1, , Format(dtpAccDt.value, "yyyy-MM-dd")
    
    
    optDonorCd(0).value = False
    optDonorCd(1).value = False
    optDonorCd(2).value = False
    optDonorCd(3).value = False
    optDonorCd(4).value = False
    
    txtReservedID.Text = ""
    txtReservedID.Enabled = False
    
    cmdReserved.Enabled = False
    lblReservedNm.Caption = ""
    
    
    lblTmpPtId.ToolTipText = ""
    
    Call tabAccDt_Click
    cmdNew.Enabled = False
'    lvwPtList.Enabled = False
    cmdSave.Enabled = True
End Sub

Private Sub cmdNewReg_Click()
   
    With frmBBS401
        .Show vbModal
    End With
End Sub

Private Sub FormInitialize()
    lblDonorId.Caption = ""
    lblDOB.Caption = ""
    lblSex.Caption = ""
    lblABO.Caption = ""
    lblCnt.Caption = ""
    lblTotVol.Caption = ""
    lblSSN.Caption = ""
    tabAccDt.Tabs.Clear
    tabAccDt.Visible = False
    dtpAccDt.value = GetSystemDate
    dtpAccDt.Enabled = False
    
    fraDonorCd.Enabled = False
    fraAcc.Enabled = False
    fraDonation.Enabled = False
    fraResult.Enabled = False
    
    Call FrameInitialize
    
    cmdNew.Enabled = False
    cmdSave.Enabled = False
    cmdCancel.Enabled = False
    blnClearFg = True
    
End Sub

Private Sub FrameInitialize()
    
    lblStsNm.Caption = ""
    lblStsCd.Caption = ""
    lblOkDiv1Nm.Caption = ""
    lblOkDiv1Cd.Caption = ""
    lblOkDiv2Nm.Caption = ""
    lblOkDiv2Cd.Caption = ""
    lblOkDiv3Nm.Caption = ""
    lblOkDiv3Cd.Caption = ""
    
    
    optDonorCd(0).value = False
    optDonorCd(1).value = False
    optDonorCd(2).value = False
    optDonorCd(3).value = False
    
    txtReservedID.Text = ""
    txtReservedID.Enabled = False
    cmdReserved.Enabled = False
    lblReservedNm.Caption = ""
    
    txtWeight.Text = ""
    txtHeight.Text = ""
    txtPulse.Text = ""
    txtBldPres1.Text = "": txtBldPres2.Text = ""
    txtBodyTemp.Text = ""
    
    lblTmpPtId.ToolTipText = ""
    
    chkHold.value = 0
    optOk(0).value = False
    optOk(1).value = False
    txtrmk = ""
End Sub

Private Sub cmdReserved_Click()
    
    
    
    Set GetPtInfo = New frmPtInfo
    GetPtInfo.Show 1



End Sub

Private Function GetDonorCd() As Long
    If optDonorCd(0).value = True Then
        GetDonorCd = 0
    ElseIf optDonorCd(1).value = True Then
        GetDonorCd = 1
    ElseIf optDonorCd(2).value = True Then
        GetDonorCd = 2
    ElseIf optDonorCd(3).value = True Then
        GetDonorCd = 3
    Else
        GetDonorCd = -1
    End If
End Function

Private Function Save_chk() As Boolean
    Dim lngDonorCd As Long
    
    
    If tabAccDt.Tabs.Count < 1 Then
        MsgBox "�ű�������ư�� �������������ϼ���", vbInformation, "�ű�����"
        Exit Function
    End If
    
    lngDonorCd = GetDonorCd
    If lngDonorCd < 0 Then
        MsgBox "���� ������ �����ϼ���.", vbInformation, "����Ȯ��"
        Exit Function
    End If
    
    'lblStsCd Status�� ���ؼ� ȭ���ɰ� ���� ���°� �ٸ� ��� ó�� �Ұ��� ó��
    If Val(lblStsCd.Caption) > 2 Then
        MsgBox "�̹� ������ �Ǿ� �ֽ��ϴ�.", vbExclamation
        Exit Function
    End If
    
    Save_chk = True
'    If Trim(txtWeight.Text) = "" Then
'        MsgBox "ü���� �Է��ϼ���.", vbInformation, "����Ȯ��"
'        txtWeight.SetFocus
'        Exit Function
'    End If
'
'    If Trim(txtHeight.Text) = "" Then
'        MsgBox "������ �Է��ϼ���", vbInformation, "����Ȯ��"
'        txtHeight.SetFocus
'        Exit Function
'    End If
'
'    If Trim(txtPulse.Text) = "" Then
'        MsgBox "�ƹ��� �Է��ϼ���", vbInformation, "����Ȯ��"
'        txtPulse.SetFocus
'        Exit Function
'    End If
'
'    If Trim(txtBldPres1.Text) = "" Then
'        MsgBox "������ �Է��ϼ���", vbInformation, "����Ȯ��"
'        txtBldPres1.SetFocus
'        Exit Function
'    End If
'
'    If Trim(txtBldPres2.Text) = "" Then
'        MsgBox "������ �Է��ϼ���", vbInformation, "����Ȯ��"
'        txtBldPres2.SetFocus
'        Exit Function
'    End If
'
'    If Trim(txtBodyTemp.Text) = "" Then
'        MsgBox "ü���� �Է��ϼ���", vbInformation, "����Ȯ��"
'        txtBodyTemp.SetFocus
'        Exit Function
'    End If
End Function

Private Function GetFormattedTmpPtID(ByVal tmpptid As String) As String
    Dim objcom003   As clsCom003
    Dim DrRS        As Recordset
    Dim fmt         As String
    Dim ii          As Integer
    
    Set objcom003 = New clsCom003
    Set DrRS = objcom003.OpenRecordSet(BC2_TMP_ID, "0")
    
    With DrRS
        If .EOF = True Then
'            'dbconn.DisplayErrors
        Else
            If .RecordCount > 0 Then
                fmt = ""
                For ii = 1 To .Fields("text1").value & ""
                    fmt = fmt & "0"
                Next ii
                tmpptid = .Fields("field3").value & "" & Format(tmpptid, fmt) & .Fields("field4").value & ""
            End If
        End If
    End With
    Set DrRS = Nothing
    Set objcom003 = Nothing
    
    GetFormattedTmpPtID = tmpptid

End Function

Private Function Save() As Boolean
    Dim RsPtid          As Recordset
    Dim arySQL()        As String
    Dim strAccDt        As String
    Dim strTmp          As String
    Dim okdiv           As String
    Dim rmk             As String
    Dim reserid         As String
    
    Dim accdt           As String
    Dim accseq          As String
    Dim WorkArea        As String
    Dim tmpid           As String
    Dim tmpptid         As String
    Dim strSEX          As String
    Dim GetAccdt        As String
    Dim strDonorCd      As String
    
    Dim IsHold          As Boolean
    Dim blnupchk        As Boolean
    
    
    Dim lngMin          As Long
    Dim lngsql          As Long
On Error GoTo Err_Trap

    If Not blnDonorFind Then
        If Not SaveDonorMst Then
            Save = False
            Exit Function
        End If
    End If
    
    Set objMySQL = New clsBBSSQLStatement
    
    '�ӽ�ȯ�ڹ�ȣ�� ���Ѵ�.
    '
    lngMin = GetTmpIDRange(False)
    
    If lngMin = 0 Then Save = False: GoTo Err_Trap

    Set RsPtid = New Recordset
    RsPtid.Open objMySQL.GetNoGiveInfo(BN_TMP_ID), DBConn
    
    If RsPtid.EOF Then
        '��ȣ�ο����� insert ����
        tmpptid = lngMin
        lngsql = lngsql + 1: ReDim Preserve arySQL(lngsql - 1)
        arySQL(lngsql - 1) = objMySQL.SetNoGiveInfo(False, BN_TMP_ID, tmpptid)
    Else
        '��ȣ�ο����� update ����
        tmpptid = GetNextPtID
        lngsql = lngsql + 1: ReDim Preserve arySQL(lngsql - 1)
        arySQL(lngsql - 1) = objMySQL.SetNoGiveInfo(True, BN_TMP_ID, tmpptid)
    End If
    Set RsPtid = Nothing
    
    tmpid = GetFormattedTmpPtID(tmpptid)
    
    strSEX = IIf(lblSex.Caption <> "", Mid(lblSex.Caption, 1, 1), "M")
    
    'ȯ�ڸ����� ����
    lngsql = lngsql + 1: ReDim Preserve arySQL(lngsql - 1)
    arySQL(lngsql - 1) = objMySQL.GetPtMasterInsertSQL(tmpid, txtDonorNm.Text, strSEX, Format(lblDOB.Caption, "yyyyMMdd"), lblSSN.Caption)
    
    reserid = txtReservedID.Text
    
    accdt = ""
    accseq = ""
    WorkArea = ""
    

    
    GetAccdt = Format(tabAccDt.SelectedItem.Caption, PRESENTDATE_FORMAT)
    strAccDt = GetAccdt
    
    
    blnupchk = UpdateChk(Trim(lblDonorId.Caption), strAccDt)
    strDonorCd = GetDonorCd
    If strDonorCd = "0" Then reserid = ""
    
    
    okdiv = IIf(optOk(0).value, "1", "0")
    rmk = Trim(txtrmk)
    
    
    IsHold = (chkHold.value = 1)
    
    If blnupchk = True Then
        '�������� UPDATE ���� ------------------------------------------
        strTmp = objMySQL.SetDonorAccHistory(blnupchk, Trim(lblDonorId.Caption), GetAccdt, strDonorCd, reserid, _
                                             tmpid, txtWeight, txtHeight, txtPulse, _
                                             txtBodyTemp, txtBldPres1, txtBldPres2, "", "", "", 0, _
                                             "", "", okdiv, rmk, _
                                             accdt, accseq, WorkArea, IsHold)
    Else
        '�������� INSERT ���� ------------------------------------------
        strTmp = objMySQL.SetDonorAccHistory(blnupchk, Trim(lblDonorId.Caption), GetAccdt, strDonorCd, reserid, _
                                             tmpid, txtWeight, txtHeight, txtPulse, _
                                             txtBodyTemp, txtBldPres1, txtBldPres2, "", "", "", 0, _
                                             "", "", okdiv, rmk, _
                                             accdt, accseq, WorkArea, IsHold)
    End If
    
    lngsql = lngsql + 1: ReDim Preserve arySQL(lngsql - 1)
    arySQL(lngsql - 1) = medGetP(strTmp, 1, ";")
    lngsql = lngsql + 1: ReDim Preserve arySQL(lngsql - 1)
    arySQL(lngsql - 1) = medGetP(strTmp, 2, ";")
    
    ReDim Preserve arySQL(lngsql): arySQL(lngsql) = ""
    
    If InsertData(arySQL) = False Then
        MsgBox "���������� ó������ �ʾҽ��ϴ�.", vbInformation, "����Ȯ��"
        Save = False
        GoTo Err_Trap
    Else
        Save = True
    End If
    
    Set objMySQL = Nothing
    Exit Function

Err_Trap:
    Save = False
    
End Function

Private Sub cmdSave_Click()
    '���� ���� üũ
    If Save_chk = False Then Exit Sub
    If Save = True Then
        Call FrameInitialize
    End If
End Sub

Private Function UpdateChk(ByVal DonorId As String, ByVal donoraccdt As String) As Boolean
    Dim Rs As New Recordset
    Dim objSql As New clsBBSSQLStatement
    
    Set Rs = objSql.GetDonorAccHistory(DonorId, donoraccdt)
    
    If Rs.EOF Then
        UpdateChk = False
    Else
        UpdateChk = True
    End If
    
    Set Rs = Nothing
    Set objSql = Nothing
End Function

Private Sub dtpAccDt_Change()
    tabAccDt.SelectedItem.Caption = Format(dtpAccDt.value, "yyyy-MM-dd")
End Sub

Private Sub Form_Load()
    txtDonorNm = ""
    Call FormInitialize
End Sub


Private Sub optDonorCd_Click(Index As Integer)
    Select Case Index
        Case 0:     '��������
            txtReservedID.Enabled = False
            cmdReserved.Enabled = False
        Case 1:     '��������
            txtReservedID.Enabled = True
            cmdReserved.Enabled = True
        Case 2:     'Autologous
            txtReservedID.Enabled = True
            cmdReserved.Enabled = True
        Case 3:     'Pheresis
            txtReservedID.Enabled = True
            cmdReserved.Enabled = True
        Case 4:     'Phlebotomy
            txtReservedID.Enabled = True
            cmdReserved.Enabled = True
        Case Else:  'Unknown
            txtReservedID.Enabled = False
            cmdReserved.Enabled = False
    End Select
End Sub
Private Function AccDtformat() As Long
    Dim objNum As New clsBBSNumbers
    
    With objNum
'        .setDbConn DBConn
        AccDtformat = Len(.Get_AccdtFormat)
        
'        .getCollectDt_cdval Format(GetSystemDate, PRESENTDATE_FORMAT)       '���� ������ ������ �����´�.(AccDt�� ������ ��¥)
'        If .field1 = "0" Then
'            AccDtformat = 8
'        ElseIf .field1 = "1" Then
'            AccDtformat = 6
'        ElseIf .field1 = "2" Then
'            AccDtformat = 4
'        End If
    End With
    Set objNum = Nothing
End Function
Private Sub tabAccDt_Click()
'
    Dim strAccDt As String
    Dim canEdit As Boolean
    
    strAccDt = Format(tabAccDt.SelectedItem.Caption, PRESENTDATE_FORMAT)
    
    Call FrameInitialize
    Call ShowDonorValue(Trim(lblDonorId.Caption), strAccDt)
    Call SetDonorStatus(Trim(lblDonorId.Caption), strAccDt)
    Call SetDonorResult(Trim(lblDonorId.Caption), strAccDt)
    
    canEdit = GetCanEdit
    fraDonorCd.Enabled = canEdit
    
    fraAcc.Enabled = canEdit
    fraDonation.Enabled = canEdit
    fraResult.Enabled = canEdit
    
  '  cmdSave.Enabled = canEdit
    
    cmdCancel.Enabled = Not canEdit
End Sub

Private Function GetCanEdit() As Boolean
    '������ ���������� �Ǵ��Ѵ�.
    If tabAccDt.SelectedItem.Index > 1 Then
        '���� �������ڰ� �ƴϴ�. ������ �� ����.
        GetCanEdit = False
    Else
        Select Case lblStsCd.Caption
            Case DonorStatus.stsAccessSave
                GetCanEdit = True
            Case DonorStatus.stsAccessVerify
                GetCanEdit = False
            Case DonorStatus.stsAskSave
                GetCanEdit = False
            Case DonorStatus.stsAskVerify
                GetCanEdit = False
            Case DonorStatus.stsDonation
                GetCanEdit = False
            Case DonorStatus.stsFinish
                GetCanEdit = False
            Case DonorStatus.stsPrint
                GetCanEdit = False
            Case Else
                If Val(lblStsCd.Caption) < DonorStatus.stsAccessSave Then
                    GetCanEdit = True
                Else
                    GetCanEdit = False
                End If
        End Select
    End If
End Function

Private Sub SetDonorResult(ByVal DonorId As String, ByVal accdt As String)
    Dim objDonor As clsBBSSQLStatement
    Dim DrRS As Recordset
    
    Set objDonor = New clsBBSSQLStatement
    Set DrRS = objDonor.GetDonorResult(DonorId, accdt)
    Set objDonor = Nothing
    
    If DrRS Is Nothing Then Exit Sub
    
    With DrRS
        If .RecordCount > 0 Then
            Select Case .Fields("okdiv1").value & ""
                Case 0:
                    optOk(1).value = True
                    optOk(0).value = False
                Case 1:
                    optOk(1).value = False
                    optOk(0).value = True
                Case Else:
                    optOk(1).value = False
                    optOk(0).value = False
            End Select
            txtrmk = .Fields("rmk1").value & ""
        End If
    End With
End Sub

Private Sub SetDonorStatus(ByVal DonorId As String, ByVal accdt As String)
    Dim objDonor As clsBBSSQLStatement
    Dim strStatus As String
    Dim IsPhere As Boolean
    
    
    
    
    Set objDonor = New clsBBSSQLStatement
    strStatus = objDonor.GetDonorStatus(DonorId, accdt, IsPhere)
    Set objDonor = Nothing
    
    lblStsNm.Caption = medGetP(strStatus, 1, vbTab)
    lblStsCd.Caption = medGetP(strStatus, 2, vbTab)
    lblOkDiv1Nm.Caption = medGetP(strStatus, 3, vbTab)
    lblOkDiv1Cd.Caption = medGetP(strStatus, 4, vbTab)
    lblOkDiv2Nm.Caption = medGetP(strStatus, 5, vbTab)
    lblOkDiv2Cd.Caption = medGetP(strStatus, 6, vbTab)
    lblOkDiv3Nm.Caption = medGetP(strStatus, 7, vbTab)
    lblOkDiv3Cd.Caption = medGetP(strStatus, 8, vbTab)
    
    If lblOkDiv1Nm.Caption = "������" Then
        lblOkDiv1Nm.ForeColor = vbRed
        lblOkDiv1Cd.ForeColor = vbRed
    Else
        lblOkDiv1Nm.ForeColor = vbBlack
        lblOkDiv1Cd.ForeColor = vbBlack
    End If
    
    If lblOkDiv2Nm.Caption = "������" Then
        lblOkDiv2Nm.ForeColor = vbRed
        lblOkDiv2Cd.ForeColor = vbRed
    Else
        lblOkDiv2Nm.ForeColor = vbBlack
        lblOkDiv2Cd.ForeColor = vbBlack
    End If
    
    If lblOkDiv3Nm.Caption = "������" Then
        lblOkDiv3Nm.ForeColor = vbRed
        lblOkDiv3Cd.ForeColor = vbRed
    Else
        lblOkDiv3Nm.ForeColor = vbBlack
        lblOkDiv3Cd.ForeColor = vbBlack
    End If
    
    
End Sub

Private Sub ShowDonorValue(ByVal DonorId As String, ByVal accdt As String)
    Dim objDonor As New clsBBSSQLStatement
    Dim Rs       As New Recordset
    
    
    With objDonor
'        .setDbConn DBConn
        Set Rs = .GetDonorAccHistory(DonorId, accdt)
    End With
    
    If Rs.EOF = False Then
        With Rs
            lblTmpPtId.ToolTipText = .Fields("tmpid").value & ""
            optDonorCd(Val(.Fields("donorcd").value & "")).value = True
            txtReservedID.Text = Trim(.Fields("reservedid").value & "")
            If Trim(.Fields("reservedid").value & "") = 0 Then txtReservedID.Text = ""
            
            Call txtReservedID_LostFocus
            
            txtWeight.Text = .Fields("weight").value & ""
            txtHeight.Text = .Fields("height").value & ""
            txtPulse.Text = .Fields("pulse").value & ""
            txtBldPres1.Text = .Fields("bldpres1").value & ""
            txtBldPres2.Text = .Fields("bldpres2").value & ""
            txtBodyTemp.Text = .Fields("bodytemp").value & ""
            
        End With
    Else
        lblTmpPtId.ToolTipText = ""
    End If
    
    Set Rs = Nothing
    Set objDonor = Nothing
End Sub
Private Sub txtAccNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txtBldPres1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txtBldPres2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txtBodyTemp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txtDonorNm_Change()
    If Not blnClearFg Then FormInitialize
End Sub

Private Sub txtDonorNm_GotFocus()
    txtDonorNm.tag = txtDonorNm
End Sub

Private Sub txtDonorNm_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call DonorFind
        blnClearFg = False
        txtDonorNm.tag = txtDonorNm
    End If
End Sub

Private Sub txtDonorNm_LostFocus()
    If txtDonorNm.tag <> txtDonorNm Then
        Call DonorFind
        blnClearFg = False
    End If
End Sub

Private Sub DonorFind()
    Dim objDonor As clsBBSBldDonationBusi
    
    blnClearFg = True
    If txtDonorNm = "" Then Call FrameInitialize: Exit Sub
    
    Set objDonor = New clsBBSBldDonationBusi
    With objDonor
        
        blnDonorFind = .DonorFind(txtDonorNm)
        If blnDonorFind = True Then
            Call FrameInitialize
            
            lblDonorId.Caption = .mDonorID
            txtDonorNm = .mDonorNm
            '2001-11-27�߰�
            strSaveDonorId = lblDonorId.Caption
            strSaveDonorNm = txtDonorNm.Text
            '
            lblDOB.Caption = .mDOB
            lblSex.Caption = .mSEX
            lblABO.Caption = .mABO
            lblCnt.Caption = .Mcnt
            lblTotVol.Caption = .mTotVol
            lblSSN.Caption = .mSSN
        
            Call ShowAccList
            cmdNew.Enabled = True
            optDonorCd(0).Enabled = True
            optDonorCd(1).Enabled = True
            optDonorCd(3).Enabled = True
        Else
            Call FrameInitialize
            cmdNew.Enabled = True
            optDonorCd(0).Enabled = False
            optDonorCd(1).Enabled = False
            optDonorCd(3).Enabled = False
        End If
    End With
    Set objDonor = Nothing
    blnClearFg = False
End Sub

Private Sub ShowAccList()
    Dim strAccDt    As String
    Dim Rs          As Recordset
    Dim objMySQL    As clsBBSSQLStatement
    '�����ڿ� ���ؼ� ������ ������ ���� ��쿡 ���� ������ �����ش�.

    Set objMySQL = New clsBBSSQLStatement

'    objMySQL.setDbConn DBConn
    Set Rs = objMySQL.GetDonorAccHistory(Trim(lblDonorId.Caption))
    
    If Rs.EOF Then
        tabAccDt.Tabs.Clear
        tabAccDt.Visible = False
    Else
        tabAccDt.Tabs.Clear
        tabAccDt.Visible = True
        
        Do Until Rs.EOF
            strAccDt = Format(Rs.Fields("donoraccdt").value & "", "####-##-##")
            tabAccDt.Tabs.Add , , strAccDt
            Rs.MoveNext
        Loop
        
        cmdSave.Enabled = True
        Call tabAccDt_Click
    End If

End Sub

Private Sub txtHeight_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txtPulse_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txtReservedID_Change()
    lblReservedNm.Caption = ""
End Sub

Private Sub txtReservedID_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(txtReservedID.Text) = "" Then Exit Sub
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtReservedID_LostFocus()
    Dim Rs As New Recordset
    Dim objPtInfo As clsPtInformation
    Dim Reserved As clsBBSSQLStatement
    
    If Trim(txtReservedID.Text) = "" Or Trim(txtReservedID.Text) = 0 Then Exit Sub
    
    Set objPtInfo = New clsPtInformation
    Set Reserved = New clsBBSSQLStatement
    
    
    Set objMySQL = New clsBBSSQLStatement
    
    Set Rs = New Recordset
    Rs.Open objPtInfo.GetPtInfo(Trim(txtReservedID.Text), True, GetSystemDate), DBConn
    
    If Rs.EOF Then
        MsgBox "��ϵ� ȯ�ڰ� �ƴϰų� �߸��� �����Դϴ�.", vbInformation, "����Ȯ��"
        With txtReservedID
            .Enabled = True
            .SetFocus
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
        
        Set Reserved = Nothing
        Set objPtInfo = Nothing
        Set objMySQL = Nothing
        Exit Sub
    End If
    
    txtReservedID.Text = Rs.Fields("ptid").value & ""
    lblReservedNm.Caption = Rs.Fields("ptnm").value & ""
    
    Set Reserved = Nothing
    Set objPtInfo = Nothing
    Set objMySQL = Nothing
End Sub

Private Function GetTmpIDRange(ByVal pMaxMin As Boolean) As Long
'�����ڵ� ������(COM003) ���� �˻��Ƿڿ� �ӽ�ȯ�� ID�� ������ üũ�Ѵ�.
'pMaxMin : True = �ִ밪, pMaxMin : False = �ּҰ�

    Dim RsRange As Recordset
    Dim objPtID As clsBBSSQLStatement
    
    Set objPtID = New clsBBSSQLStatement

    Set RsRange = New Recordset
    RsRange.Open objPtID.GetTestReq4IDRange, DBConn
    
    If RsRange.EOF Then
'        MsgBox "�˻��Ƿڿ� �ӽ�ȯ�� ID�� ������ �����ϼ���.", vbInformation, "����Ȯ��"
        
        GetTmpIDRange = 0
        Set RsRange = Nothing
        Set objPtID = Nothing
        Exit Function
    Else
        If pMaxMin Then
            GetTmpIDRange = Val(Trim(RsRange.Fields("field2").value & ""))
        Else
            GetTmpIDRange = Val(Trim(RsRange.Fields("field1").value & ""))
        End If
    End If
    
    Set RsRange = Nothing
    Set objPtID = Nothing
End Function

Private Function GetNextPtID() As Long
'��ȣ�ο� �������� ������ ����� ID�� ���´�.
    Dim RsNext As Recordset
    Dim objSSql As clsBBSSQLStatement
    
    Set objSSql = New clsBBSSQLStatement
    
    With objSSql
'        .setDbConn DBConn
        Set RsNext = New Recordset
        RsNext.Open .GetNoGiveMaxSeq(BN_TMP_ID), DBConn
    End With
    
    GetNextPtID = Val(Trim(RsNext.Fields("maxseq").value & "")) + 1
    Set objSSql = Nothing
    
End Function








Private Function GetPtID() As String
'PtID�� �����Ѵ�.
'    Dim RsPtid As New RECORDSET
'    Dim arysql(1) As String
'    Dim lngMin As Long
'    Dim objcom003 As clsCom003
'    Dim DrRS As RECORDSET
'    Dim fmt As String
'    Dim i As Long
'
'
'    lngMin = GetTmpIDRange(False)
'    If lngMin = 0 Then Exit Function
'
'    Set objMySQL = New clsBBSSQLStatement
'
'    With objMySQL
'        .setDbConn DbConn
'        Set RsPtid = OpenRecordSet(.GetNoGiveInfo(BN_TMP_ID))
'    End With
'
'    If RsPtid.EOF Then
'        'PtID�� ����
'        arysql(0) = objMySQL.SetNoGiveInfo(False, BN_TMP_ID, lngMin)
'        Call insertdata(arysql)
'        GetPtID = lngMin
'    Else
'        GetPtID = GetNextPtID
'    End If
'    '���⼭ ���������� �����ؾ� ���� ������?
'
'    'Ptid�� ������Ʈ
'    Call SetPtIDUpdate(GetPtID)
'
'    Set RsPtid = Nothing
'    Set objMySQL = Nothing
'
'    Dim objcom003 As clsCom003
'    Set objcom003 = New clsCom003
'    Set DrRS = objcom003.OpenRecordSet(BC2_TMP_ID, "0")
'
'    With DrRS
'        If .DBerror = True Then
'            'dbconn.DisplayErrors
'        Else
'            If .RecordCount > 0 Then
'                fmt = ""
'                For i = 1 To .Fields("text1")
'                    fmt = fmt & "0"
'                Next i
'                GetPtID = .Fields("field3") & Format(GetPtID, fmt) & .Fields("field4")
'            End If
'            .RsClose
'        End If
'    End With
'
'    Set DrRS = Nothing
'    Set objcom003 = Nothing
End Function
Private Function SetPtIDUpdate(ByVal pGetPtID As String) As Boolean
'    'PtID�� ������ �ٷ� ������Ʈ �Ѵ�.
'    'Dim objMySQL As New clsBBSSQLStatement
'    Set objMySQL = New clsBBSSQLStatement
'On Error GoTo SetPtIDUpdate
'    With DbConn
'        .BeginTrans
'        .Execute objMySQL.SetNoGiveInfo(True, BN_TMP_ID, pGetPtID)
'        .CommitTrans
'    End With
'
'    Set objMySQL = Nothing
'    SetPtIDUpdate = True
'    Exit Function
'
'SetPtIDUpdate:
'
'    With DbConn
'        .RollbackTrans
'        .DisplayErrors
'        SetPtIDUpdate = False
'    End With
End Function

Private Sub txtWeight_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
    
End Sub

Private Function SaveDonorMst() As Boolean
    Dim strSeq      As String
    Dim blnUpdateFg As Boolean
    Dim arySSN()    As String
    Dim strSSN      As String
    Dim aryZipCd()  As String
    Dim strZipCd    As String

    Dim SSQL        As String
    Dim objBg       As clsBeginTrans

    SaveDonorMst = True
    Set objBg = New clsBeginTrans

On Error GoTo SAVE_ERROR

    strSSN = ""
    strZipCd = ""


    '��� �ֱ�����. Seq �� ����
    blnUpdateFg = IIf(GetNoGiveInfo, True, False)

    strSeq = GetNoGiveSeq
    '�����ڸ���������.............
    SSQL = objBg.SetDonorMST(False, strSeq, Trim(txtDonorNm.Text), _
                             strSSN, "", "", _
                             strZipCd, "", "", "", _
                             "", "", "", 0, 0)
    DBConn.Execute SSQL
    '��ȣ�ο���������............
    SSQL = objBg.SetNoGiveInfo(blnUpdateFg, BN_DONOR_ID, Val(strSeq))
    DBConn.Execute SSQL

    Exit Function

SAVE_ERROR:
    SaveDonorMst = False
    Set objBg = Nothing

End Function

Private Function GetNoGiveInfo() As Boolean
'��ȣ�ο� ���� ������Ʈ üũ
    
    Dim Rs As New Recordset
    Dim objNoGive As clsBBSSQLStatement
    Dim arySQL(1) As String
    
    Set objNoGive = New clsBBSSQLStatement
    With objNoGive
'        .setDbConn DBConn
        Set Rs = New Recordset
        Rs.Open .GetNoGiveInfo(BN_DONOR_ID), DBConn
    End With
           
    If Rs.EOF Then
    '�ʵ尡 �������� �ʴ� ��� Insert ����
        arySQL(0) = objNoGive.SetNoGiveInfo(False, BN_DONOR_ID, 0)
        Call InsertData(arySQL, False)
    End If
    
    GetNoGiveInfo = True
    
    Set Rs = Nothing
    Set objNoGive = Nothing
End Function

Private Function GetNoGiveSeq() As String
'��ȣ�ο� �������� �ְ��� ���´�.

    Dim Rs As New Recordset
    Dim objMaxSeq As clsBBSSQLStatement
    
    Set objMaxSeq = New clsBBSSQLStatement
    With objMaxSeq
'        .setDbConn DBConn
        Rs.Open .GetNoGiveMaxSeq(BN_DONOR_ID), DBConn
    End With
    
    If Rs.EOF Then
        GetNoGiveSeq = 1
    Else
        GetNoGiveSeq = Rs.Fields("maxseq").value & "" + 1
    End If
    
    Set Rs = Nothing
    Set objMaxSeq = Nothing
End Function


