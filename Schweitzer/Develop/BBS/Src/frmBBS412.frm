VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBBS412 
   BackColor       =   &H00DBE6E6&
   Caption         =   "PHERESIS ���"
   ClientHeight    =   9015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14700
   Icon            =   "frmBBS412.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9015
   ScaleWidth      =   14700
   WindowState     =   2  '�ִ�ȭ
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "����(&X)"
      Height          =   510
      Left            =   10875
      Style           =   1  '�׷���
      TabIndex        =   7
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
      TabIndex        =   6
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
      TabIndex        =   5
      Tag             =   "15101"
      Top             =   7575
      Width           =   1320
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00F4F0F2&
      Caption         =   "������"
      Height          =   510
      Left            =   6915
      Style           =   1  '�׷���
      TabIndex        =   8
      Tag             =   "15101"
      Top             =   7575
      Visible         =   0   'False
      Width           =   1320
   End
   Begin MSComctlLib.TabStrip tabAccDt 
      Height          =   315
      Left            =   2280
      TabIndex        =   9
      Top             =   2025
      Width           =   9930
      _ExtentX        =   17515
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
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   315
      Left            =   2280
      TabIndex        =   10
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
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   315
      Left            =   2280
      TabIndex        =   11
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
   Begin VB.Frame Frame4 
      BackColor       =   &H00DBE6E6&
      Height          =   975
      Left            =   2280
      TabIndex        =   19
      Top             =   720
      Width           =   9930
      Begin VB.TextBox txtDonorNm 
         Appearance      =   0  '���
         Height          =   330
         Left            =   1035
         TabIndex        =   0
         Top             =   180
         Width           =   1515
      End
      Begin MedControls1.LisLabel lblDOB 
         Height          =   315
         Left            =   4260
         TabIndex        =   20
         Top             =   180
         Width           =   1260
         _ExtentX        =   2223
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
         Caption         =   "2001-01-01"
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblSex 
         Height          =   330
         Left            =   6615
         TabIndex        =   21
         Top             =   180
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
         Left            =   8925
         TabIndex        =   22
         Top             =   180
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
         Left            =   4260
         TabIndex        =   23
         Top             =   540
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
         Left            =   6615
         TabIndex        =   24
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
         Left            =   1035
         TabIndex        =   25
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
         TabIndex        =   26
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
         Left            =   45
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   180
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
         Left            =   3270
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   180
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
         Left            =   3270
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   540
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
         Index           =   3
         Left            =   5625
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   180
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
         Left            =   5625
         TabIndex        =   31
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
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   5
         Left            =   7935
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   180
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
         Left            =   7530
         TabIndex        =   33
         Top             =   690
         Width           =   210
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00DBE6E6&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   2280
      TabIndex        =   34
      Top             =   2280
      Width           =   9930
      Begin VB.ComboBox cboDonorCd 
         Appearance      =   0  '���
         Height          =   300
         ItemData        =   "frmBBS412.frx":076A
         Left            =   1035
         List            =   "frmBBS412.frx":077A
         Locked          =   -1  'True
         Style           =   1  '�ܼ� �޺�
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   225
         Width           =   2055
      End
      Begin VB.TextBox txtReservedID 
         Alignment       =   2  '��� ����
         BackColor       =   &H00CFDCDE&
         Height          =   330
         Left            =   4245
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   225
         Width           =   1125
      End
      Begin MedControls1.LisLabel lblReservedNm 
         Height          =   330
         Left            =   5370
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   225
         Width           =   2640
         _ExtentX        =   4657
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
         Caption         =   ""
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   315
         Index           =   10
         Left            =   45
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   225
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
      Begin MedControls1.LisLabel lblTmpPtId 
         Height          =   315
         Left            =   3255
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   225
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00DBE6E6&
      Height          =   615
      Left            =   2280
      TabIndex        =   40
      Top             =   2880
      Width           =   9930
      Begin MedControls1.LisLabel lblStsNm 
         Height          =   315
         Left            =   1050
         TabIndex        =   41
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
         Left            =   2295
         TabIndex        =   42
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
         Left            =   3585
         TabIndex        =   43
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
         Left            =   4530
         TabIndex        =   44
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
         Left            =   5850
         TabIndex        =   45
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
         Left            =   6795
         TabIndex        =   46
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
         Left            =   8115
         TabIndex        =   47
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
         Left            =   9075
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
      Begin MedControls1.LisLabel lbldt 
         Height          =   315
         Index           =   6
         Left            =   45
         TabIndex        =   49
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
         Left            =   2595
         TabIndex        =   50
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
         Left            =   4845
         TabIndex        =   51
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
         Left            =   7110
         TabIndex        =   52
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
   Begin VB.Frame fraDonation 
      BackColor       =   &H00DBE6E6&
      Height          =   4065
      Left            =   2280
      TabIndex        =   12
      Top             =   3420
      Width           =   9930
      Begin VB.ComboBox cboCompo 
         Height          =   300
         ItemData        =   "frmBBS412.frx":07A8
         Left            =   1395
         List            =   "frmBBS412.frx":07B5
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   3
         Top             =   1110
         Width           =   2175
      End
      Begin VB.TextBox txtRemark 
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2505
         Left            =   1395
         MultiLine       =   -1  'True
         ScrollBars      =   2  '����
         TabIndex        =   4
         Top             =   1455
         Width           =   7785
      End
      Begin VB.TextBox txtVolumn 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   7170
         MaxLength       =   10
         TabIndex        =   17
         Top             =   1140
         Width           =   825
      End
      Begin VB.CheckBox chkBar 
         BackColor       =   &H00DBE6E6&
         Caption         =   "���ڵ�ó��"
         Height          =   195
         Left            =   3675
         TabIndex        =   16
         Top             =   810
         Width           =   1215
      End
      Begin VB.TextBox txtBldNo 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1395
         MaxLength       =   12
         TabIndex        =   2
         Top             =   735
         Width           =   2130
      End
      Begin VB.OptionButton optVo 
         BackColor       =   &H00DBE6E6&
         Caption         =   "320cc"
         Height          =   270
         Index           =   0
         Left            =   4665
         TabIndex        =   15
         Top             =   1170
         Width           =   795
      End
      Begin VB.OptionButton optVo 
         BackColor       =   &H00DBE6E6&
         Caption         =   "400cc"
         Height          =   270
         Index           =   1
         Left            =   5535
         TabIndex        =   14
         Top             =   1155
         Value           =   -1  'True
         Width           =   795
      End
      Begin VB.OptionButton optVo 
         BackColor       =   &H00DBE6E6&
         Caption         =   "��Ÿ"
         Height          =   270
         Index           =   2
         Left            =   6420
         TabIndex        =   13
         Top             =   1155
         Width           =   675
      End
      Begin MSComCtl2.DTPicker dtpDonationDt 
         Height          =   315
         Left            =   1380
         TabIndex        =   1
         Top             =   375
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   60620803
         CurrentDate     =   36797
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   12
         Left            =   360
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   375
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
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   13
         Left            =   360
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   735
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
         Caption         =   "���׹�ȣ"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   14
         Left            =   360
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   1095
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
         Caption         =   "��������"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   15
         Left            =   360
         TabIndex        =   56
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
         Caption         =   "Remark"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   16
         Left            =   3585
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   1110
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
         Caption         =   "���׷�"
         Appearance      =   0
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "cc"
         Height          =   180
         Left            =   8070
         TabIndex        =   18
         Top             =   1275
         Width           =   210
      End
   End
End
Attribute VB_Name = "frmBBS412"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Enum TblColumn
    tcSEL = 1
    tcName
    tcCODE
    tcQTY
End Enum
Private objMySQL As New clsBBSSQLStatement
Private objMyOrder As New clsDonorBusiOrder
Private objMyCollection As New clsDonorBusiCollection

Private AccDtform As Long
'2001-11-27�߰�
Private strSaveDonorId As String
Private strSaveDonorNm As String

Private Sub cboCompo_Click()
    If cboDonorCd.ListIndex <> 3 Then Exit Sub
    Dim objSQL     As clsBBSSQLStatement
    Dim aryOrdCd() As String
    Dim today      As Date
    Dim Volumn     As String
    Dim CompoCd    As String
    Dim Cnt        As Long
    Dim i          As Long
    
    today = GetSystemDate
    
    Volumn = "0"
    
    Set objSQL = New clsBBSSQLStatement
'    objSql.setDbConn DBConn
'    CompoCd = medGetP(cboCompo.Text, 1, " ")
    Cnt = objSQL.GetOrdCd(Volumn, CompoCd, Format(today, PRESENTDATE_FORMAT), aryOrdCd)
    Set objSQL = Nothing
    
'    cboNewTest.Clear
'    If cnt > 0 Then
'        For i = 1 To cnt
'            cboNewTest.AddItem aryOrdCd(i - 1)
'        Next i
'        cboNewTest.ListIndex = 0
'    End If
End Sub

Private Sub chkBar_Click()
    txtBldNo.SetFocus
End Sub

Private Sub cmdCancel_Click()
'�������(602�� cancelfg="1",401 ���ڵ� ����,lab102(Dcfg='1')
    If cboDonorCd.ListIndex = 3 Then
        MsgBox "Pheresis ������ ��� �ϽǼ� �����ϴ�.", vbInformation + vbOKOnly, "�������"
        Exit Sub
    End If
    Dim objSQL      As New clsBBSSQLStatement
    Dim BldSrc      As String
    Dim BldYY       As String
    Dim BldNo       As String
    Dim CompoCd     As String
    Dim donorid     As String
    Dim donoraccdt  As String
    Dim tmpptid     As String
    
    donoraccdt = Format(tabAccDt.SelectedItem.Caption, PRESENTDATE_FORMAT)
    donorid = lblDonorID.Caption
'    CompoCd = medGetP(cboCompo.Text, 1, " ")
    BldSrc = medGetP(txtBldNo, 1, "-")
    BldYY = medGetP(txtBldNo, 2, "-")
    BldNo = medGetP(txtBldNo, 3, "-")
    tmpptid = lblTmpPtId.ToolTipText
    
'    objSql.setDbConn DBConn
    
    If objSQL.SetPheresisCancel(donorid, donoraccdt, tmpptid) Then
        MsgBox "��������� ��ҵǾ����ϴ�.", vbInformation + vbOKOnly, "�������"
        FrameInitialize
    End If
'    If objSql.SetBldCancel(donorid, donoraccdt, tmpptid, BldSrc, BldYY, BldNo, CompoCd) Then
'        MsgBox "��������� ��ҵǾ����ϴ�.", vbInformation + vbOKOnly, "�������"
'        FrameInitialize
'    End If
    
    Set objSQL = Nothing
End Sub

Private Sub cmdClear_Click()
    FrameInitialize
End Sub

Private Sub cmdExit_Click()
    Set objMySQL = Nothing
    Set objMyOrder = Nothing
    Set objMyCollection = Nothing
    Unload Me
    Set frmBBS412 = Nothing
End Sub


Private Sub cmdSave_Click()
    
    '�Է��� ���׹�ȣ�� �԰����� �����Ѵٸ�, ���� �ϸ� �ʵȴ�.
    If Bld_Check(txtBldNo) = False Then Exit Sub
    
    '�������� ���(�������� ����)2001/10/04
    
    If SetPheresisSave = True Then FrameInitialize
    
    
    
    'If SaveAll = True Then FrameInitialize
    
End Sub
Private Function Bld_Check(ByVal BldNum As String) As Boolean
    Dim objSQL As clsBBSSQLStatement
    Dim BldSrc As String
    Dim BldYY  As String
    Dim BldNo  As String
    Dim CompoCd As String
    
    If txtBldNo = "" Then
        MsgBox "���������� �Է��� �۾��� �����Ͻʽÿ�", vbInformation + vbOKOnly, "��������"
        Exit Function
    End If
    
    If cboCompo.ListIndex < 1 Then
        MsgBox "���������� �����Ͻ��� ��� �Ͻʽÿ�", vbInformation + vbOKOnly, "������������"
        Exit Function
    End If
    If txtVolumn.Text = "" Then
        MsgBox "Volumn�� �Է��Ͻʽÿ�.", vbCritical, Me.Caption
        Exit Function
    End If
    
    Set objSQL = New clsBBSSQLStatement
    
    If chkBar.value <> 1 Then
        BldSrc = medGetP(BldNum, 1, "-")
        BldYY = medGetP(BldNum, 2, "-")
        BldNo = medGetP(BldNum, 3, "-")
    Else
        BldSrc = Mid(BldNum, 1, 2)
        BldYY = Mid(BldNum, 3, 2)
        BldNo = Mid(BldNum, 5, 6)
    End If
    
    'CompoCd = ObjSQL.GetCompocdPheresis ' medGetP(lblCompoCd.Caption, 1, " ")
    
    CompoCd = medGetP(cboCompo.Text, 1, " ")
    
    
    If objSQL.GetBloodCheck(BldSrc, BldYY, BldNo, CompoCd) = True Then
        Bld_Check = True
    Else
        MsgBox "�̹� �԰�� ���׹�ȣ�Դϴ�. Ȯ���� ����ϼ���", vbInformation + vbOKOnly, "�������"
    End If
    Set objSQL = Nothing
End Function
Private Sub Form_Activate()
    medMain.lblSubMenu.Caption = Me.Caption
'    lblTestChk.Visible = False
End Sub

Private Sub Form_Load()

    dtpDonationDt.value = GetSystemDate
    AccDtform = AccDtformat
    
    
    '������������(���������ΰ��)
    '2001/09/28
    Call SetCboCompList
    
    Call SetMaterial
    Call FrameInitialize
    Call ClassInitialize

End Sub
Private Sub SetCboCompList()
    Dim objSQL As clsBBSADDSQL
    Dim Rs     As Recordset
    Dim i      As Integer
    
    Set objSQL = New clsBBSADDSQL
    Set Rs = objSQL.Get_PheresisCompoNm
    
    If Not Rs.EOF Then
        cboCompo.Clear
        cboCompo.AddItem "������������"
        
        For i = 1 To Rs.RecordCount
            Do Until Rs.EOF
                cboCompo.AddItem Rs.Fields("compocd").value & "" & " " & _
                                 Rs.Fields("abbrnm").value & "" & Space(20) & COL_DIV & _
                                 Rs.Fields("keepday").value & ""
                Rs.MoveNext
            Loop
        Next i
        cboCompo.ListIndex = 0
    End If
    
    Set objSQL = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set objMySQL = Nothing
    Set objMyOrder = Nothing
    Set objMyCollection = Nothing
End Sub

Private Function AccDtformat() As Long
    Dim objNum As New clsBBSNumbers
    
    AccDtformat = Len(objNum.Get_AccdtFormat)
    
    Set objNum = Nothing
End Function

Private Sub optVo_Click(Index As Integer)
    Select Case Index
        Case 0: txtVolumn = "320"
        Case 1: txtVolumn = "400"
        Case 2: txtVolumn = ""
    End Select
End Sub

Private Sub tabAccDt_Click()
    
    Dim donorid As String
    Dim canEdit As Boolean
    
    donorid = lblDonorID.Caption
    Call tabAccdtClickCode(donorid, Format(tabAccDt.SelectedItem.Caption, PRESENTDATE_FORMAT))
    Call SetDonorStatus(donorid, Format(tabAccDt.SelectedItem.Caption, PRESENTDATE_FORMAT))
    'Call SetDonorMaterial
    '---------------------------------------------
    '����/������ ������ ��� ��������� �����ϰ�
    '�Ұ����̸� Ǯ���ָ� �ȴ�(2001/08/17)
    '---------------------------------------------
'    canEdit = GetCanEdit
'    fraDonation.Enabled = canEdit
'    canEdit = GetCanEdit
'    fraDonation.Enabled = canEdit

End Sub

Private Function GetCanEdit() As Boolean
    '������ ���������� �Ǵ��Ѵ�.
    If tabAccDt.SelectedItem.Index > 1 Then
        '���� �������ڰ� �ƴϴ�. ������ �� ����.
        GetCanEdit = False
    Else
        Select Case lblStsCd.Caption
            Case DonorStatus.stsAccessSave
                GetCanEdit = False
            Case DonorStatus.stsAccessVerify
                GetCanEdit = False
            Case DonorStatus.stsAskSave
                GetCanEdit = False
            Case DonorStatus.stsAskVerify
                GetCanEdit = (lblOkDiv2Cd.Caption = "1")
            Case DonorStatus.stsDonation
                GetCanEdit = True
            Case DonorStatus.stsFinish
                GetCanEdit = False
            Case DonorStatus.stsPrint
                GetCanEdit = False
            Case Else
                GetCanEdit = False
        End Select
    End If
End Function

Private Sub SetDonorStatus(ByVal donorid As String, ByVal accdt As String)
    Dim objDonor As clsBBSSQLStatement
    Dim strStatus As String
    Dim IsPhere As Boolean
    
    
    Set objDonor = New clsBBSSQLStatement
    strStatus = objDonor.GetDonorStatus(donorid, accdt, IsPhere)
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

Private Sub tblMaterial_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
'    With tblMaterial
'        If Row < 1 Or Row > .MaxRows Then Exit Sub
'
'        Select Case Col
'            Case TblColumn.tcSEL:
'                .Row = Row
'                .Col = TblColumn.tcSEL
'
'                '������ 1��-------------------
'                If .value = 1 Then
'                    .Row = Row
'                    .Col = TblColumn.tcQTY: .value = 1
'                    .Row2 = Row
'                    .Col2 = TblColumn.tcQTY
'                    .BlockMode = True
'                    .Lock = False
'                    .BlockMode = False
'                Else
'                    .Row = Row
'                    .Col = TblColumn.tcQTY: .value = ""
'                    .Row2 = Row
'                    .Col2 = TblColumn.tcQTY
'                    .BlockMode = True
'                    .Lock = True
'                    .BlockMode = False
'                End If
'        End Select
'    End With
End Sub

Private Sub SetAccNo()
'    Dim objSql As clsBBSSQLStatement
'    Dim Rs     As RECORDSET
'    Dim accdt  As String
'    Dim accseq As String
'
'    If txtAccNo = "" Then
'        txtAccNo = ""
'        lblOrdNm.Caption = ""
'        Exit Sub
'    End If
'
'    accdt = medGetP(txtAccNo, 1, "-")
'    accseq = medGetP(txtAccNo, 2, "-")
'
'    Set objSql = New clsBBSSQLStatement
'
''    objSql.setDbConn DBConn
'    Set Rs = objSql.Get_PheresisInfo(accdt, accseq)
'    If Not Rs.EOF Then
'        If txtReservedID <> Rs.Fields("ptid") Then
'            MsgBox "����ȯ���� ������ȣ�� �ƴմϴ�.", vbInformation, "ó��������ȸ"
'        Else
'            lblNewTestDiv = Rs.Fields("newtestdiv")
'            lblOrdNm.Caption = Rs.Fields("testnm")
'            XM_Method txtAccNo
'        End If
'    Else
'        MsgBox "�ش�Ǵ� ������ �����ϴ�.", vbInformation + vbOKOnly, "ó��������ȸ"
'    End If
'    Set Rs = Nothing
'    Set objSql = Nothing
End Sub

Private Sub txtAccNo_GotFocus()
'    txtAccNo.tag = txtAccNo
End Sub

Private Sub txtAccNo_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'        Call SetAccNo
'        txtAccNo.tag = txtAccNo
'    End If
End Sub

Private Sub txtAccNo_LostFocus()
'    If txtAccNo.tag <> txtAccNo Then
'        Call SetAccNo
'    End If
End Sub

Private Sub txtAccNo_Change()
'    Dim lngLen As Long
'
'    With txtAccNo
'        lngLen = Len(Trim(.Text))
'        If lngLen = AccDtform Then
'                .Text = .Text & "-"
'                .SelStart = Len(.Text)
'        End If
'    End With
End Sub


Private Sub txtAccNo_KeyPress(KeyAscii As Integer)

'    If KeyAscii = vbKeyBack Then
'        With txtAccNo
'            If .Text = "" Then Exit Sub
'            If Mid(.Text, Len(.Text)) = "-" Then
'                .Text = Mid(.Text, 1, Len(.Text) - 2)
'                .SelStart = Len(.Text)
'                KeyAscii = 0
'            End If
'        End With
'    End If

End Sub
Private Sub txtDonorNm_GotFocus()
    txtDonorNm.tag = txtDonorNm
End Sub

Private Sub txtDonorNm_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call DonorFind
        txtDonorNm.tag = txtDonorNm
    End If
End Sub

Private Sub txtDonorNm_LostFocus()
    If txtDonorNm.tag <> txtDonorNm Then
        Call DonorFind
    End If
End Sub

Private Sub DonorFind()
    Dim objDonor As clsBBSBldDonationBusi
    
    If txtDonorNm = "" Then Call FrameInitialize: Exit Sub
    
    Set objDonor = New clsBBSBldDonationBusi
    With objDonor

        If .DonorFind(txtDonorNm) = True Then
            Call FrameInitialize
            
            lblDonorID.Caption = .mDonorID
            txtDonorNm = .mDonorNm
            '2001-11-27 �߰�
            strSaveDonorId = lblDonorID.Caption
            strSaveDonorNm = txtDonorNm.Text
            '
            lblDOB.Caption = .mDOB
            lblSex.Caption = .mSEX
            lblABO.Caption = .mABO
            lblCnt.Caption = .Mcnt
            lblTotVol.Caption = .mTotVol
        
            Call ShowAccList
'            cmdNew.Enabled = True
        End If
    End With
    Set objDonor = Nothing
End Sub

Private Sub ShowAccList()
    Dim strAccDt As String
    Dim Rs As Recordset
    Dim objMySQL As clsBBSSQLStatement
    '�����ڿ� ���ؼ� ������ ������ ���� ��쿡 ���� ������ �����ش�.

    Set objMySQL = New clsBBSSQLStatement

'    objMySQL.setDbConn DBConn
    'Set Rs = objMySQL.GetDonorAccHistory(Trim(lblDonorID.Caption))
    
    
    '���������ڸ� ��ȸ�ؼ� �����ֱ� ���� �߰�(2001/10/04,��굿������ ����)
    
    Set Rs = objMySQL.GetDonorAccdtHistoryDivPheresis(Trim(lblDonorID.Caption), , True)
    
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

Private Sub FrameInitialize()
    dtpDonationDt = GetSystemDate
    tabAccDt.Tabs.Clear
    tabAccDt.Visible = False
    
    cboDonorCd.ListIndex = -1
    txtReservedID.Text = ""
    lblReservedNm.Caption = ""
    
    lblStsNm.Caption = ""
    lblStsCd.Caption = ""
    lblOkDiv1Nm.Caption = ""
    lblOkDiv1Cd.Caption = ""
    lblOkDiv2Nm.Caption = ""
    lblOkDiv2Cd.Caption = ""
    lblOkDiv3Nm.Caption = ""
    lblOkDiv3Cd.Caption = ""

'    fraTest.Visible = False
    
    lblTmpPtId.ToolTipText = ""
'    txtAccNo = ""
    txtDonorNm = ""
    lblDonorID.Caption = ""
    lblSex.Caption = ""
    lblABO.Caption = ""
    lblCnt.Caption = ""
    lblTotVol.Caption = ""
    lblDOB.Caption = ""
    cmdSave.Enabled = False
    cmdCancel.Enabled = False
    txtVolumn.Enabled = False
    optVo(0).value = True
    txtVolumn.Text = "320"
    cboCompo.ListIndex = 0
    
    Clear
End Sub

Private Sub Clear()
    Dim r As Long
    
'    cboCompo.ListIndex = -1
    txtVolumn = ""
    txtBldNo = ""
    
'    With tblMaterial
'        For r = 1 To .MaxRows
'            .Row = r
'            .Col = TblColumn.tcSEL:  .value = 0
'            .Col = TblColumn.tcQTY:  .value = ""
'        Next r
'    End With
End Sub

Private Sub SetCboCompo(ByVal TF As Boolean)
'    'tf:t�̸�pheresis ��������
'    Dim objCompo  As New clsBBSSQLStatement
'    Dim Rs        As New RECORDSET
'    Dim i         As Integer
'
''    objCompo.setDbConn DBConn
'
'    Set Rs = objCompo.Compolist(TF)
'
'    If Not Rs.EOF Then
'        cboCompo.Clear
'        For i = 1 To Rs.RecordCount
'            Do Until Rs.EOF
'                cboCompo.AddItem Rs.Fields("compocd") & " " & Rs.Fields("abbrnm") & Space(20) & COL_DIV & Rs.Fields("keepday")
'                Rs.MoveNext
'            Loop
'        Next i
'    End If
'
'    Set Rs = Nothing
'    Set objCompo = Nothing
End Sub

Private Sub SetMaterial()
'    Dim objcom003 As clsCom003
'    Dim DrRS As RECORDSET
'    Dim i As Long
'
'    tblMaterial.MaxRows = 0
'
'    Set objcom003 = New clsCom003
'
'    Set DrRS = objcom003.OpenRecordSet(BC2_MATERIAL)
'    If DrRS Is Nothing Then Exit Sub
'
'    With tblMaterial
'        .MaxRows = DrRS.RecordCount
'        For i = 1 To DrRS.RecordCount
'            .Row = i
'            .Col = TblColumn.tcSEL:  .value = 0
'            .Col = TblColumn.tcCODE: .value = DrRS.Fields("cdval1")
'            .Col = TblColumn.TcName: .value = DrRS.Fields("field1")
'
'            DrRS.MoveNext
'        Next i
'    End With
'
'    Set DrRS = Nothing
'    Set objcom003 = Nothing
End Sub

Private Sub SetDonorMaterial()
    Dim donorid As String
    Dim donoraccdt As String
    Dim objDonorMaterial As clsDonorMaterial
    Dim DrRS As Recordset
    Dim i As Long
    Dim r As Long
    
    Dim RsTestReq As Recordset
    Clear
   
    donorid = lblDonorID.Caption
    donoraccdt = Format(tabAccDt.SelectedItem.Caption, PRESENTDATE_FORMAT)
    
    
    
    Set objMySQL = New clsBBSSQLStatement
    With objMySQL
'        .setDbConn DBConn
        Set RsTestReq = .GetDonorAccHistory(donorid, donoraccdt)
    End With
    
    If Not RsTestReq.EOF Then
        '�ӽ�ȯ��id
        lblTmpPtId.ToolTipText = RsTestReq.Fields("tmpid")
        
        '��������
        Select Case RsTestReq.Fields("donorcd")
            Case "0": cboDonorCd.ListIndex = 0
            Case "1": cboDonorCd.ListIndex = 1
            Case "2": cboDonorCd.ListIndex = 2
            Case "3": cboDonorCd.ListIndex = 3
        End Select
        txtReservedID = RsTestReq.Fields("reservedid")
    End If
    Set RsTestReq = Nothing
    Set objMySQL = Nothing
    
    
    '--------------------------------------------------
    '�� �����Ͽ� ��ϵ� ���������� ��᳻���� �ҷ��´�.
    '--------------------------------------------------
    
    
    Set objDonorMaterial = New clsDonorMaterial
    '------------------------------------------------------------------------
    '�� �� �� ��
    '------------------------------------------------------------------------
    Set DrRS = objDonorMaterial.GetDonorDonation(donorid, donoraccdt)
    If DrRS Is Nothing Then Exit Sub
    With DrRS
        If .RecordCount > 0 Then
'            cboCompo.ListIndex = medComboFind(cboCompo, .Fields("compocd") & " " & .Fields("componm"))
            txtVolumn = .Fields("volumn") & ""
            If Trim(.Fields("donationdt")) <> "" Then
                dtpDonationDt = Format(.Fields("donationdt"), "####-##-##")
            End If
            If Trim(.Fields("bldsrc")) <> "" Then
                txtBldNo = .Fields("bldsrc") & "-" & .Fields("bldyy") & "-" & .Fields("bldno")
            End If
        End If
    End With
    Set DrRS = Nothing
    
    '------------------------------------------------------------------------
    '�� �� �� ��
    '------------------------------------------------------------------------
'    Set DrRS = objDonorMaterial.GetDonorMaterial(Donorid, DonorAccdt)
'    If DrRS Is Nothing Then Exit Sub
'
'    With tblMaterial
'        For i = 1 To DrRS.RecordCount
'            For r = 1 To .MaxRows
'                .Row = r
'                .Col = TblColumn.tcCODE
'                If Trim(.value) = Trim(DrRS.Fields("ordcd")) Then
'                    .Col = TblColumn.tcSEL: .value = 1
'                    .Col = TblColumn.tcQTY: .value = DrRS.Fields("qty")
'
'                    Exit For
'                End If
'            Next r
'            DrRS.MoveNext
'        Next i
'    End With
'
'    Set DrRS = Nothing
    Set objDonorMaterial = Nothing
End Sub
Private Sub Used_Material(ByVal donorid, ByVal donoraccdt As String)
'   '------------------------------------------------------------------------
'   ' �� �� �� ��
'   ' ------------------------------------------------------------------------
'   Dim DrRS             As New RECORDSET
'   Dim objDonorMaterial As New clsDonorMaterial
'   Dim i                As Integer
'   Dim r                As Integer
'
'    Set DrRS = objDonorMaterial.GetDonorMaterial(donorid, donoraccdt)
'    If DrRS Is Nothing Then Exit Sub
'
'    With tblMaterial
'        For i = 1 To DrRS.RecordCount
'            For r = 1 To .MaxRows
'                .Row = r
'                .Col = TblColumn.tcCODE
'                If Trim(.value) = Trim(DrRS.Fields("ordcd")) Then
'                    .Col = TblColumn.tcSEL: .value = 1
'                    .Col = TblColumn.tcQTY: .value = DrRS.Fields("qty")
'                End If
'            Next r
'            DrRS.MoveNext
'        Next i
'    End With
'
'    Set DrRS = Nothing
'    Set objDonorMaterial = Nothing
End Sub




Private Function Save() As Boolean
    
    Dim objSQL         As clsBBSSQLStatement
    Dim Rs             As Recordset
    Dim BldSrc         As String
    Dim BldYY          As String
    Dim BldNo          As String
    Dim CompoCd        As String
    Dim Volumn         As String
    Dim ABO            As String
    Dim Rh             As String
    Dim PtId           As String
    Dim RFg            As String
    Dim AFg            As String
    Dim PFg             As String
    Dim ExpDt          As String
    Dim Dt             As String
    Dim Tm             As String
    Dim id             As String
    Dim CenterCd       As String
    Dim stscd          As String
    
    Dim donorid        As String
    Dim DonationDt     As String
    Dim donoraccdt     As String
    Dim Available      As Long
    '���������ۼ���....
    Dim ObjDic         As clsDictionary
    Dim DeliveryHold   As String
    Dim strTmp         As String
    Dim Orderptid      As String
    Dim orddt          As String
    Dim ordno          As String
    Dim Ordseq         As String
    Dim FilterFg       As String
    Dim IrradFg        As String
    Dim ordcd          As String
    Dim Ostscd         As String
    Dim MaterialCd     As String
    Dim MateriaQty     As String
    
    Dim Bordcd         As String
    Dim accdt          As String
    Dim accseq         As String
    Dim Method         As String
    
    Dim ii             As Integer
    
    Set ObjDic = New clsDictionary
    Set objSQL = New clsBBSSQLStatement
    Set Rs = New Recordset
    
    ObjDic.Clear
    ObjDic.FieldInialize "ptid,orddt,ordno,ordseq,ordcd,div", "unitqty"
    
    If chkBar.value <> 1 Then
        BldSrc = medGetP(txtBldNo, 1, "-")
        BldYY = medGetP(txtBldNo, 2, "-")
        BldNo = medGetP(txtBldNo, 3, "-")
    Else
        BldSrc = Mid(txtBldNo, 1, 2)
        BldYY = Mid(txtBldNo, 3, 2)
        BldNo = Mid(Mid(txtBldNo, 5), 1, Len(Mid(txtBldNo, 5)) - 2)
    End If
    
    If optVo(0).value = True Then
        Volumn = "320"
    ElseIf optVo(1).value = True Then
        Volumn = "400"
    Else
        Volumn = txtVolumn
    End If
    
    Select Case cboDonorCd.ListIndex
        Case "0":   cboDonorCd.ListIndex = 0: PtId = "":            RFg = "0": AFg = "0": PFg = "0"
        Case "1":   cboDonorCd.ListIndex = 1: PtId = txtReservedID: RFg = "1": AFg = "0": PFg = "0"
        Case "2":   cboDonorCd.ListIndex = 2: PtId = txtReservedID: RFg = "0": AFg = "1": PFg = "0"
        Case "3":   cboDonorCd.ListIndex = 3: PtId = txtReservedID: RFg = "0": AFg = "0": PFg = "1"  ': Method = cboMethod.ListIndex
    End Select
    Ostscd = BBSOrderStatus.stsEnd
    Set Rs = Nothing
    
    donorid = lblDonorID.Caption                                    '������Id
    donoraccdt = Format(tabAccDt.SelectedItem.Caption, PRESENTDATE_FORMAT)  '����������
    DonationDt = Format(dtpDonationDt.value, PRESENTDATE_FORMAT)            '������
    ABO = Mid(lblABO.Caption, 1, 1)                                 '������
    Rh = Mid(lblABO.Caption, 2, 1)                                  'rh
    id = ObjMyUser.EmpId                                            '�����ID
    Dt = Format(GetSystemDate, PRESENTDATE_FORMAT)                      '�������
    Tm = Format(GetSystemDate, PRESENTTIME_FORMAT)                        '����Ͻð�
    CenterCd = ObjSysInfo.BuildingCd                                '�����ڵ�
    ExpDt = DateAdd("d", Available, dtpDonationDt.value)            '
    ExpDt = Format(ExpDt, PRESENTDATE_FORMAT)                               '�����(��ȿ�ϰ� ���)
    Save = objSQL.SetPheresis(DonationDt, BldSrc, BldYY, BldNo, Volumn, donorid, donoraccdt, Dt, id)
    Set ObjDic = Nothing
    Set objSQL = Nothing
End Function

'20010212 �˻��Ƿڳ������� �Űܿ� �ڵ带
Private Sub txtBldNo_Change()
    Dim lngLen As Long
    
    If chkBar.value = 1 Then Exit Sub
    
    
    With txtBldNo
        lngLen = Len(Trim(.Text))
        If lngLen = 2 Then
                .Text = .Text & "-"
                .SelStart = Len(.Text)
        End If
        If lngLen > 2 And lngLen = 5 Then
            .Text = .Text & "-"
            .SelStart = Len(.Text)
        End If
    End With
End Sub

Private Sub txtBldNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txtBldNo_KeyPress(KeyAscii As Integer)
    If chkBar.value = 1 Then Exit Sub
    
    If Len(txtBldNo) <> 3 Or Len(txtBldNo) <> 6 Then
        If KeyAscii = vbKeyInsert Then KeyAscii = 0
    End If
    
    If KeyAscii = vbKeyBack Then
        With txtBldNo
            If .Text = "" Then Exit Sub
            If Mid(.Text, Len(.Text)) = "-" Then
                .Text = Mid(.Text, 1, Len(.Text) - 2)
                .SelStart = Len(.Text)
                KeyAscii = 0
            End If
        End With
    End If
End Sub

Private Sub tabAccdtClickCode(ByVal donorid As String, ByVal donoraccdt As String)
    Dim RsTestReq     As Recordset
    Dim ii            As Integer
    
    
    '�����ڿ� ���ؼ� �ӻ󺴸��� �˻��Ƿڸ� �Ѱ��� ���ܵȴ�.
    If tabAccDt.SelectedItem.Selected Then
        Set objMySQL = New clsBBSSQLStatement
        
        '������ ���������� �д´�.--------------------------------------
        Set RsTestReq = objMySQL.GetDonorAccHistory(donorid, donoraccdt)
        If RsTestReq.EOF Then
'            'dbconn.DisplayErrors
            Set objMySQL = Nothing
            Exit Sub
        End If
        
        If RsTestReq.RecordCount < 1 Then
            MsgBox "���������� ã�� �� �����ϴ�.", vbCritical, "����"
            Set RsTestReq = Nothing
            Set objMySQL = Nothing
            Exit Sub
        End If
        
        '�������� Display-----------------------------------------------
        
        
        Select Case RsTestReq.Fields("donorcd")
            Case "0":   cboDonorCd.ListIndex = 0
                        
            Case "1":   cboDonorCd.ListIndex = 1
            Case "2":   cboDonorCd.ListIndex = 2
            Case "3":   cboDonorCd.ListIndex = 3
            Case Else:  cboDonorCd.ListIndex = -1
        End Select
        
        
        
       'ó���� �����ͼ� ���� ��� ���� �ϴ� �����.
       '=============================================================================
'        fraTest.Visible = False
'        If RsTestReq.Fields("accdt") = "" Then
'            txtAccNo = ""
'        Else
'            txtAccNo = RsTestReq.Fields("accdt") & "-" & RsTestReq.Fields("accseq")
'        End If
'        '�˻���

'        XM_Method txtAccNo
       '=============================================================================
        
        
        txtReservedID = RsTestReq.Fields("reservedid").value & ""
        lblReservedNm.Caption = objMySQL.GetPtntNm(txtReservedID)
        
        If RsTestReq.Fields("donationdt").value & "" <> "" Then
            dtpDonationDt.value = Format(RsTestReq.Fields("donationdt").value & "", "####-##-##")
            For ii = 0 To cboCompo.ListCount
                If medGetP(cboCompo.List(ii), 1, " ") = RsTestReq.Fields("compocd") Then
                    cboCompo.ListIndex = ii
                    Exit For
                End If
            Next
            Select Case RsTestReq.Fields("volumn")
                Case "320": optVo(0).value = True: txtVolumn.Text = "320"
                Case "400": optVo(1).value = True: txtVolumn.Text = "400"
                Case Else:  optVo(2).value = True: txtVolumn = RsTestReq.Fields("volumn")
            End Select
            txtBldNo = RsTestReq.Fields("bldsrc").value & "" & "-" & _
                       RsTestReq.Fields("bldyy").value & "" & "-" & _
                       RsTestReq.Fields("bldno").value & ""
            cmdSave.Enabled = False
            'cmdCancel.Enabled = True
        Else
            txtVolumn.Text = "320"
            cmdSave.Enabled = True
            'cmdCancel.Enabled = False
        End If
                        
        '�߰�����볻��
        'Used_Material donorid, donoraccdt
            
        
        Set RsTestReq = Nothing
        Set objMySQL = Nothing
    End If
    
End Sub

Private Sub XM_Method(ByVal strTmp As String)
'������ȣ�� ������ �ش� ó�濡 ���� �˻� ����� ������´�......
'    Dim objSql As clsBBSSQLStatement
'    Dim accdt  As String
'    Dim accseq As String
'
'    If strTmp = "" Then
'        cboMethod.ListIndex = -1
'    Else
'        Set objSql = New clsBBSSQLStatement
'
'        accdt = medGetP(strTmp, 1, "-")
'        accseq = medGetP(strTmp, 2, "-")
'
''        objSql.setDbConn DBConn
'        cboMethod.ListIndex = objSql.Get_XMethod(accdt, accseq)
'
'        Set objSql = Nothing
'    End If
End Sub
Private Sub QueryInformation(tmpid As String)
'�ӽ�ȯ��id�� ó���ȣ�� ������ �˻������� ��ȸ�Ѵ�.
'�׽�Ʈ �÷� ����
'1:ó���ȣ,2:�˻��,3:�˻��ڵ�,4:��ü,5:�޿�,6:���޿���,7:���ä���Ͻ�
'8:���޿���:9:WorkArea,10:storecd,11:rndfg,12;testdiv,13:multifg,14:spcgrp,15:ordseq
'16:����,17:���ڵ�������,18:�˻簡�ɿ���,19:Location,20:�˻����
    Dim objQueryTest As New clsBBSSQLStatement
    Dim objDicT As New clsDictionary
    Dim objDicD As New clsDictionary
    Dim RsDonorTest As Recordset
    Dim RsDisplay As Recordset
    Dim strTmp As String
    
    Set objMySQL = New clsBBSSQLStatement
'    objMySQL.setDbConn DBConn
    
    objDicT.Clear
    objDicT.FieldInialize "ptid,orddt,ordno,ordseq", "ordcd,spccd,reqdate"
    
    
    Set RsDonorTest = objMySQL.GetDonnorTest(tmpid)
    
    If Not RsDonorTest.EOF Then
        Do Until RsDonorTest.EOF
            objDicT.AddNew Join(Array(RsDonorTest.Fields("ptid"), RsDonorTest.Fields("orddt"), RsDonorTest.Fields("ordno"), _
                                RsDonorTest.Fields("ordseq")), COL_DIV), Join(Array(RsDonorTest.Fields("ordcd"), _
                                RsDonorTest.Fields("spccd"), RsDonorTest.Fields("reqdt") & RsDonorTest.Fields("reqtm")), COL_DIV)
            RsDonorTest.MoveNext
        Loop
    End If
    
    
    If objDicT.RecordCount > 0 Then
        objDicD.Clear
        objDicD.FieldInialize "ptid,orddt,ordno,ordseq", "ordno1,testnm,testcd,spccd,gubyu,stat,reqdt,statfg,workarea," & _
                              "storecd,rndfg,testdiv,multifg,spcgrp,ordseq1,abbrnm5,labelcnt,statflag,location,testlocation"
        objDicT.MoveFirst
        Do Until objDicT.EOF
            strTmp = objDicT.Fields("ordcd") & vbTab & objDicT.Fields("spccd")
            Set RsDisplay = objMySQL.GetTestFindList(strTmp)
            
            With RsDisplay
                objDicD.AddNew Join(Array(objDicT.Fields("ptid"), objDicT.Fields("orddt"), objDicT.Fields("ordno"), objDicT.Fields("ordseq")), COL_DIV), _
                               Join(Array(objDicT.Fields("ordno"), .Fields("testnm"), .Fields("testcd"), .Fields("spccd"), _
                                          "1", "", objDicT.Fields("reqdate"), .Fields("statfg"), .Fields("workarea"), _
                                          .Fields("storecd"), .Fields("rndfg"), .Fields("testdiv"), .Fields("multifg"), _
                                          .Fields("spcgrp"), objDicT.Fields("ordseq"), .Fields("abbrnm5"), _
                                          .Fields("labelcnt"), .Fields("statflags"), "location", "�߾�"), COL_DIV)
            End With
            objDicT.MoveNext
        Loop
    End If
    'ȭ�鿡 ��������......
'    Call TblResult_Display(objDicD)
    '''
    
    Set objDicD = Nothing
End Sub

Private Sub Default_Test(objDefault As clsDictionary)
    Dim objQueryTest As New clsBBSSQLStatement
    Dim objGDic As New clsDictionary
    Dim DefaultTest As Recordset
    Dim strTmp As String
    Dim lngseq As Long
    
    objGDic.Clear
    objGDic.FieldInialize "seq", "ordno1,testnm,testcd,spccd,gubyu,stat,reqdt,statfg,workarea," & _
                          "storecd,rndfg,testdiv,multifg,spcgrp,ordseq1,abbrnm5,labelcnt,statflag,location,testlocation"
    objDefault.MoveFirst
    
    Do Until objDefault.EOF
        
        strTmp = objDefault.Fields("testcd") & vbTab & objDefault.Fields("spccd")
        Set DefaultTest = Nothing
        Set DefaultTest = New Recordset
        DefaultTest.Open objQueryTest.GetDefaultTestList(strTmp), DBConn
        With DefaultTest
            If Not DefaultTest.EOF Then
                lngseq = lngseq + 1
                objGDic.AddNew lngseq, _
                               Join(Array("", .Fields("testnm"), .Fields("testcd"), .Fields("spccd"), _
                                          "1", "", Format(GetSystemDate, "yyyy-MM-dd" & " " & "hh:MM"), .Fields("statfg"), .Fields("workarea"), _
                                          .Fields("storecd"), .Fields("rndfg"), .Fields("testdiv"), .Fields("multifg"), _
                                          .Fields("spcgrp"), "", .Fields("abbrnm5"), _
                                          .Fields("labelcnt"), .Fields("statflags"), "location", "�߾�"), COL_DIV)
            End If
        End With
        objDefault.MoveNext
    Loop
    
    'ȭ�鿡 ��������......
'    Call TblResult_Display(objGDic)
    Set objGDic = Nothing
    Set objQueryTest = Nothing
End Sub

Private Function SaveAll() As Boolean
'    If Not (cboDonorCd.ListIndex = 3) Then
'        SaveAll = SaveDonation
'    Else
        SaveAll = SavePheresis
'    End If
End Function

Private Function SavePheresis() As Boolean
    Dim strOrdDt As String
    Dim strWorkArea As String
    Dim strAccDt As String
    Dim lngAccSeq As Long
    Dim blnSuccess As Boolean
    Dim objSQL As clsBBSSQLStatement
    
    Dim donorid As String
    Dim accdt As String
    
    donorid = lblDonorID.Caption
    accdt = Format(tabAccDt.SelectedItem.Caption, PRESENTDATE_FORMAT)
    
On Error GoTo ErrSave

    '----- Begin Transaction -----
    DBConn.BeginTrans
   
   ' �����԰��� ����
    If Save = False Then GoTo ErrSave
'----- Commit Transaction -----

'    Set objSql = New clsBBSSQLStatement
'    Call objSql.SetDonorStatus(donorid, accdt, DonorStatus.stsDonation, False)
'    Set objSql = Nothing
    
    DBConn.CommitTrans
    SavePheresis = True
    MsgBox "���������� ó���Ǿ����ϴ�.", vbInformation, "����Ȯ��"
    
    Call ClassInitialize
    'Call FormInitialize
    
    Exit Function
    
ErrSave:
'----- Rollback Transaction -----
    DBConn.RollbackTrans
    Call ClassInitialize
    
    SavePheresis = False
    MsgBox Err.Description, vbExclamation
End Function


'% �߻��� ó�浥��Ÿ�� �������� ä������������ �����ϱ� ����
'% ��� ����Ÿ�� Array�� Assign�Ѵ�.
Private Sub ReadyToCollect()
    
    Dim i As Integer
    Dim tmpData() As String
    Dim datDateTime As Date
    
    datDateTime = GetSystemDate
    
    With objMyCollection
    
        .spcyy = Mid(Format(datDateTime, "YYYY"), 4)         '��ü�⵵
        .PtId = objMyOrder.PtId                                    'ȯ��ID
        .ptnm = txtDonorNm
        
        'DonorID, DonorAccDt�� �Ѱ��ش�.
        .donorid = lblDonorID.Caption
        .donoraccdt = Format(tabAccDt.SelectedItem.Caption, PRESENTDATE_FORMAT)
        
        .Sex = Mid(lblSex.Caption, 1, 1)                            '����
        
        .AgeDay = DateDiff("y", medGetP(lblSex.Caption, 2, "/"), datDateTime) 'ȯ���Ϸ�
        .bedindt = ""                                               '�Կ���
        .EntDt = Format(datDateTime, CS_DateDbFormat)         '�Է���
        .EntTm = Format(datDateTime, CS_TimeDbFormat)         '�Է½ð�
        .EntID = ObjMyUser.EmpId                                    '�Է���
        .OrgAccNo = ""                                              '��������ȣ
        .wardid = ""                                                '����ID
        .HosilID = ""                                               '����ID
        .ROOMID = ""                                                '����ID
        .BedID = ""                                                 'ħ��ID
        .coldt = Format(datDateTime, CS_DateDbFormat)         'ä����
        .colid = ObjMyUser.EmpId                                    'ä����
        .OrgBuildCd = ObjSysInfo.BuildingCd                         '** ä���� ����Ǵ� �ǹ��ڵ�
    End With

End Sub

Private Sub ClassInitialize()
    Dim datDateTime  As Date
    
    datDateTime = GetSystemDate
    
    Set objMySQL = Nothing
    Set objMySQL = New clsBBSSQLStatement
'    objMySQL.setDbConn DBConn
    
    Set objMyOrder = Nothing
    Set objMyOrder = New clsDonorBusiOrder
    With objMyOrder
        .DateTime = datDateTime
        .BuildingNo = ObjSysInfo.BuildingNo
'        .setDbConn DBConn
    End With
    
    Set objMyCollection = Nothing
    Set objMyCollection = New clsDonorBusiCollection
    
    With objMyCollection
        .DateTime = datDateTime
        Set .SortList = frmControls.lstList
        Call .InitRtn
    End With
End Sub

Private Function SetPheresisSave() As Boolean
    
    Dim objPHERE As clsBBSADDSQL
        
    On Error GoTo PHERESave_ERROR
    Set objPHERE = New clsBBSADDSQL
    With objPHERE
    
        Select Case chkBar.value
            Case 0
                  .BldSrc = medGetP(txtBldNo, 1, "-")
                  .BldYY = medGetP(txtBldNo, 2, "-")
                  .BldNo = medGetP(txtBldNo, 3, "-")
                  
            Case 1
                  .BldSrc = Mid(txtBldNo, 1, 2)
                  .BldYY = Mid(txtBldNo, 3, 2)
                  .BldNo = Mid(txtBldNo, 5, 6)
        End Select
        .CompoCd = medGetP(cboCompo.Text, 1, " ")
        .Volumn = txtVolumn.Text
        .ABO = Mid(lblABO.Caption, 1, 1)
        .Rh = Mid(lblABO.Caption, 2, 1)
        .PtId = txtReservedID.Text
        .ExcuteID = ObjMyUser.EmpId
        .Available = medGetP(cboCompo.Text, 2, COL_DIV)
        .ExpDt = Format(DateAdd("d", medGetP(cboCompo.Text, 2, COL_DIV), dtpDonationDt.value), PRESENTDATE_FORMAT)
        .RealDt = Format(GetSystemDate, PRESENTDATE_FORMAT)
        .RealTm = Format(GetSystemDate, PRESENTTIME_FORMAT)
        .donorid = lblDonorID.Caption
        .donoraccdt = Format(tabAccDt.SelectedItem.Caption, PRESENTDATE_FORMAT)
        
        DBConn.Execute .SetPheresisInsert401
        DBConn.Execute .SetPheresisUpdate411
        DBConn.Execute .SetPheresisUpdate603
        DBConn.Execute .SetPhereUpdate602
        
    
    
    End With
    
    Set objPHERE = Nothing
    MsgBox "���������� ó���Ǿ����ϴ�.", vbInformation + vbOKOnly, "Pheresis ���"
    SetPheresisSave = True
    Exit Function
    
PHERESave_ERROR:
    DBConn.RollbackTrans
    MsgBox "���������� ó������ �ʾҽ��ϴ�.", vbInformation + vbOKOnly, "Pheresis ��Ͽ���"
    Set objPHERE = Nothing
    

End Function

'2001-11-27�߰�
Public Sub CallDonorNmLostFocus()
    Call txtDonorNm_LostFocus
End Sub



