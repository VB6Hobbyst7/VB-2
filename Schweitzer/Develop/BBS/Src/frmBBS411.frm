VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frmBBS411 
   BackColor       =   &H00DBE6E6&
   Caption         =   "�˻��Ƿ�"
   ClientHeight    =   9255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13920
   Icon            =   "frmBBS411.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   13920
   WindowState     =   2  '�ִ�ȭ
   Begin VB.CommandButton cmdCallBlood 
      BackColor       =   &H00F4F0F2&
      Caption         =   "�������(&N)"
      Height          =   510
      Left            =   2250
      Style           =   1  '�׷���
      TabIndex        =   4
      Tag             =   "15101"
      Top             =   7575
      Width           =   1320
   End
   Begin VB.CommandButton cmdPhersis 
      BackColor       =   &H00C8CEDF&
      Caption         =   "Phersis���(&N)"
      Height          =   510
      Left            =   3585
      Style           =   1  '�׷���
      TabIndex        =   5
      Tag             =   "15101"
      Top             =   7575
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "����(&X)"
      Height          =   510
      Left            =   10875
      Style           =   1  '�׷���
      TabIndex        =   6
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
      TabIndex        =   2
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
      TabIndex        =   1
      Tag             =   "15101"
      Top             =   7575
      Width           =   1320
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00F4F0F2&
      Caption         =   "�˻����"
      Height          =   510
      Left            =   6915
      Style           =   1  '�׷���
      TabIndex        =   3
      Tag             =   "15101"
      Top             =   7575
      Visible         =   0   'False
      Width           =   1320
   End
   Begin MSComctlLib.TabStrip tabAccDt 
      Height          =   315
      Left            =   2280
      TabIndex        =   7
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
      TabIndex        =   9
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
      TabIndex        =   10
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
      TabIndex        =   8
      Top             =   2310
      Width           =   9930
      Begin VB.TextBox txtReservedID 
         Alignment       =   2  '��� ����
         BackColor       =   &H00CFDCDE&
         Height          =   330
         Left            =   4245
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   225
         Width           =   1125
      End
      Begin VB.ComboBox cboDonorCd 
         Appearance      =   0  '���
         Height          =   300
         ItemData        =   "frmBBS411.frx":076A
         Left            =   1035
         List            =   "frmBBS411.frx":077A
         Locked          =   -1  'True
         Style           =   1  '�ܼ� �޺�
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   225
         Width           =   2055
      End
      Begin MedControls1.LisLabel lblReservedNm 
         Height          =   330
         Left            =   5370
         TabIndex        =   35
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
         TabIndex        =   36
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
      Begin MedControls1.LisLabel lbldt 
         Height          =   315
         Index           =   11
         Left            =   3255
         TabIndex        =   37
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
      TabIndex        =   11
      Top             =   2910
      Width           =   9930
      Begin MedControls1.LisLabel lblStsNm 
         Height          =   315
         Left            =   1050
         TabIndex        =   38
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
         TabIndex        =   39
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
         TabIndex        =   40
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
         TabIndex        =   41
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
         TabIndex        =   42
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
         TabIndex        =   43
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
         TabIndex        =   44
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
         TabIndex        =   45
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
         TabIndex        =   46
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
         TabIndex        =   47
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
         TabIndex        =   48
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
         Caption         =   "�˻���"
         Appearance      =   0
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00DBE6E6&
      Height          =   975
      Left            =   2280
      TabIndex        =   16
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
         TabIndex        =   17
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
         TabIndex        =   18
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
         TabIndex        =   19
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
         TabIndex        =   20
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
         TabIndex        =   21
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
         TabIndex        =   22
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
         TabIndex        =   23
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
         TabIndex        =   24
         Top             =   690
         Width           =   210
      End
   End
   Begin MedControls1.LisLabel LisLabel3 
      Height          =   315
      Left            =   2280
      TabIndex        =   25
      Top             =   3525
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   556
      BackColor       =   8388608
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
      Caption         =   "   �� �� �� ��"
      Appearance      =   0
   End
   Begin VB.Frame fraTest 
      BackColor       =   &H00DBE6E6&
      Height          =   3720
      Left            =   2280
      TabIndex        =   12
      Top             =   3765
      Width           =   9930
      Begin MedControls1.LisLabel lblTestChk 
         Height          =   345
         Left            =   30
         TabIndex        =   26
         Top             =   120
         Visible         =   0   'False
         Width           =   7140
         _ExtentX        =   12594
         _ExtentY        =   609
         BackColor       =   12632256
         ForeColor       =   16576
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "�̹� �˻��Ƿڵ� �������Դϴ�."
      End
      Begin FPSpread.vaSpread tblResult 
         Height          =   3150
         Left            =   30
         TabIndex        =   13
         Tag             =   "10114"
         Top             =   495
         Width           =   9765
         _Version        =   196608
         _ExtentX        =   17224
         _ExtentY        =   5556
         _StockProps     =   64
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         EditEnterAction =   5
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FormulaSync     =   0   'False
         GridShowVert    =   0   'False
         MaxCols         =   24
         MaxRows         =   11
         MoveActiveOnFocus=   0   'False
         ProcessTab      =   -1  'True
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         ShadowText      =   0
         SpreadDesigner  =   "frmBBS411.frx":07A8
         StartingColNumber=   2
         VirtualRows     =   24
         VisibleCols     =   5
         VisibleRows     =   11
      End
      Begin MedControls1.LisLabel lblTmpPtId 
         Height          =   315
         Left            =   8190
         TabIndex        =   14
         Top             =   150
         Width           =   1605
         _ExtentX        =   2831
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
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin MSComctlLib.TabStrip tabGroup 
         Height          =   345
         Left            =   30
         TabIndex        =   15
         Top             =   135
         Width           =   7125
         _ExtentX        =   12568
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   315
         Index           =   12
         Left            =   7200
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   150
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
         Caption         =   "�ӽ� ID"
         Appearance      =   0
      End
   End
End
Attribute VB_Name = "frmBBS411"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Enum TblColumn
    tcSEL = 1
    TcName
    tcCODE
    tcQTY
End Enum
Private objMySQL As New clsBBSSQLStatement
Private objMyOrder As New clsDonorBusiOrder
Private objMyCollection As New clsDonorTestCollection
Private objCollect As New clsLISCollectioin

'2001-11-27�߰�
Private strSaveDonorId As String
Private strSaveDonorNm As String


Private Sub FrameInitialize()
    tabAccDt.Tabs.Clear
    tabAccDt.Visible = False
    medClearTable tblResult
    lblTmpPtId.Caption = ""
    txtDonorNm = ""
    lblDonorID.Caption = ""
    lblSex.Caption = ""
    lblABO.Caption = ""
    lblCnt.Caption = ""
    lblTotVol.Caption = ""
    lblDOB.Caption = ""
    lblTestChk.Visible = False
    cmdSave.Enabled = False
    cmdCancel.Enabled = False

End Sub
'2001-11-27�߰�
Private Sub cmdCallBlood_Click()
    frmBBS404.Show
    frmBBS404.txtDonorNm.Text = strSaveDonorNm
    Call frmBBS404.CallDonorNmLostFocus

End Sub

Private Sub cmdCancel_Click()
    Dim donorid As String
    Dim donoraccdt As String
    Dim tmpptid As String
    
    If tabAccDt.SelectedItem Is Nothing Then Exit Sub
    
    donorid = lblDonorID.Caption
    donoraccdt = Format(tabAccDt.SelectedItem.Caption, PRESENTDATE_FORMAT)
    tmpptid = lblTmpPtId.Caption
    
    If objMySQL.SetDonorScreenCancel(donorid, donoraccdt, tmpptid) = True Then
        Call FrameInitialize
    End If
End Sub

Private Sub cmdClear_Click()
    Call FrameInitialize
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub
'2001-11-27�߰�
Private Sub cmdPhersis_Click()
    
    frmBBS412.Show
    frmBBS412.txtDonorNm.Text = strSaveDonorNm
    Call frmBBS412.CallDonorNmLostFocus

End Sub
Private Sub cmdSave_Click()
    
    If Not TEST_FOR_PHERSIS And cboDonorCd.ListIndex = 3 Then
        MsgBox "���������� �˻��Ƿ��ϽǼ� �����ϴ�." & vbCrLf & "Pheresis���ȭ���� ����Ͻʽÿ�", vbInformation + vbOKOnly
        Exit Sub
    End If
        
    If tblResult.DataRowCnt = 0 Then
        MsgBox "�˻��Ƿ��� �׸��� �����ϴ�.", vbInformation, "����Ȯ��"
    Else
        If Save = True Then
        End If
        Call ClassInitialize
    End If
End Sub

Private Function Save() As Boolean
    Dim strOrdDt    As String
    Dim strWorkArea As String
    Dim strAccDt    As String
    Dim lngAccSeq   As Integer
    Dim blnSuccess  As Boolean
    Dim objSQL      As clsBBSSQLStatement

    Dim donorid     As String
    Dim accdt       As String
    Dim SSQL        As String
    Dim ii          As Integer
    
    donorid = lblDonorID.Caption
    If donorid = "" Then
        MsgBox "Donor�� �������� �����ϼ���.", vbInformation + vbOKOnly, Me.Caption
        Exit Function
    End If
    
    accdt = Format(tabAccDt.SelectedItem.Caption, PRESENTDATE_FORMAT)
    
    Call TblSort
    
On Error GoTo ErrOther
    
    'ó�� ��ƾ
    If SaveOrder = False Then GoTo ErrOther
    
    'objCollect.InitRtn
    
    Call ReadyToCollect              'ä���غ�
    
    '���� ������(2001/09/20)
    If objMyCollection.DoCollection = False Then GoTo ErrOther    'ä������

 '----- Begin Transaction -----
    DBConn.BeginTrans
   
On Error GoTo ErrSave

    'ó�泻�� ����
    blnSuccess = objMyOrder.ExecuteSqlStmt
    If blnSuccess = False Then GoTo ErrSave
    '���� ������(2001/09/20)
    
    'ä������ ����
    blnSuccess = objMyCollection.ExecuteSqlStmt
    If blnSuccess = False Then GoTo ErrSave
'==============2001/09/20===================
    'ä�� ��ƾ
'    objCollect.SetTrans = False
'    If objCollect.DoCollection = False Then GoTo ErrSave
'    For ii = 1 To objCollect.ColCount
'        Call objCollect.GetLabNumbers(ii, strWorkArea, strAccdt, lngAccSeq)
'        sSql = objMySQL.SetDonorAccHistoryUpdateByTmpID2(donorid, accdt, lblTmpPtId.Caption)
'        DBConn.Execute sSql
'        sSql = objMySQL.SetTestRequest(donorid, accdt, _
'                            Format(GetSystemDate, PRESENTDATE_FORMAT), ii, strWorkArea, strAccdt, lngAccSeq)
'        DBConn.Execute sSql
'    Next
'============================================
    For ii = 1 To objMyCollection.ColCount
        objMyCollection.GetBarcodeLabel (ii)
    Next

'----- Commit Transaction -----

    Set objSQL = New clsBBSSQLStatement
    If objSQL.SetDonorStatus(donorid, accdt, DonorStatus.stsDonation, False) = False Then GoTo ErrSave
    
    SSQL = objSQL.SetDonorAcc(donorid, accdt)
    DBConn.Execute SSQL
    
    Set objSQL = Nothing
    
    DBConn.CommitTrans
    Save = True
    MsgBox "���������� ó���Ǿ����ϴ�.", vbInformation, "����Ȯ��"
    
    Call FrameInitialize
    Exit Function
    
ErrSave:
'----- Rollback Transaction -----
    DBConn.RollbackTrans
    Save = False
    MsgBox Err.Description, vbExclamation
    Exit Function
    
ErrOther:
    
    MsgBox "���������� ó������ �ʾҽ��ϴ�.", vbInformation, "����Ȯ��"

    Save = False

End Function

Private Function SaveOrder() As Boolean
    Dim i As Long
    Dim lngStartOrdNo As Long
    Dim strTmpPtID As String
    Dim strDonorAccdt As String
    Dim datDateTime As Date
    
    datDateTime = GetSystemDate
    'strTmpPtID = GetPtID
    '������ id�� ���� �ӽ�ȯ��id�� �Ѱ��ش�.
    '20010206
    'strDonorAccdt = Format(tabAccDt.SelectedItem.Caption, PRESENTDATE_FORMAT)
    strTmpPtID = lblTmpPtId.Caption ' GetPtID(strDonorAccdt, lblDonorID.Caption)
    
    If strTmpPtID = "" Or strTmpPtID = "0" Then SaveOrder = False: Exit Function
    
    With objMyOrder
        'Order Class �⺻ ����Ÿ Store
        .PtId = strTmpPtID   '��ȣ �ο� �������� ����
        .orddt = Format(datDateTime, CS_DateDbFormat)
        .Bussdiv = "1"  '�ܷ� 1, ���� 2, ���� 3 ��ü ���� 4
        .bedindt = ""
        .DeptCd = BLOOD_DEPTCD
        .MajDoct = ""
        .wardid = ""
        .HosilID = ""
        .ROOMID = ""
        .OrdDoct = ObjMyUser.EmpId
        .ReceptNo = ""
        .EntID = ObjMyUser.EmpId
        .EntDt = Format(datDateTime, CS_DateDbFormat)
        .EntTm = Format(datDateTime, CS_TimeDbFormat)
        .donefg = "0" 'ó�� '0'
        .OrgAccNo = ""
        .orddiv = "L"
        Call .MoveData(tblResult)                   'Ŭ������ ����Ÿ Move
        If .CreateSqlStmt(lngStartOrdNo) = False Then MsgBox "Createsqlstmt ����": Exit Function  'Database�� ����
        
    End With

    With tblResult
        .Col = 1
        For i = 1 To .DataRowCnt
            .Row = i
            .value = Val(.value) + lngStartOrdNo
        Next
    End With
    
    SaveOrder = True
End Function

'% �߻��� ó�浥��Ÿ�� �������� ä������������ �����ϱ� ����
'% ��� ����Ÿ�� Array�� Assign�Ѵ�.
Private Sub ReadyToCollect()
    

    Dim i As Integer
    Dim tmpData() As String
    Dim datDateTime As Date
    
    datDateTime = GetSystemDate
    
    With objMyCollection

        .spcyy = LIS_BarDiv & Mid(Format(datDateTime, "YYYY"), 4)         '��ü�⵵
       
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
        
    With tblResult
        ReDim tmpData(0 To 17)
        
        For i = 1 To .DataRowCnt
           .Row = i
           .Col = 19:  tmpData(0) = .value                          'Delivery Location
           .Col = 12:  tmpData(1) = .value                          'TestDiv
           .Col = 9:   tmpData(2) = .value                          'WorkArea
           .Col = 4:   tmpData(3) = .value                          'SpcCd
           .Col = 10:  tmpData(4) = .value                          'StoreCd
           .Col = 6:   tmpData(5) = CStr(Val(.value))               'StatFg
           .Col = 7:   tmpData(6) = .value                          'ReqColDate
           
           .Col = 13:  tmpData(7) = .value                          'MultiFg
           .Col = 14:  tmpData(8) = .value                          'SpcGrp
           tmpData(9) = Format(datDateTime, CS_DateDbFormat)        'ó������ ���ä���Ϸ�.. 2000/04/03 by ���̰�
           .Col = 1:   tmpData(10) = .value                         'OrdNo
           .Col = 15:  tmpData(11) = .value                         'OrdSeq
           .Col = 3:   tmpData(12) = .value                         'OrdCd
           tmpData(13) = ObjMyUser.DeptCd                           '�����
           tmpData(14) = ObjMyUser.EmpId                            'ó����
           tmpData(15) = ""                                         '��ġ��
           .Col = 16:  tmpData(16) = .value                         '����
           .Col = 17:  tmpData(17) = .value                         '��������
           Call objMyCollection.AddLabCollect(tmpData(0), tmpData(1), tmpData(2), tmpData(3), tmpData(4), _
                                      tmpData(5), tmpData(6), tmpData(7), tmpData(8), tmpData(9), tmpData(10), _
                                      tmpData(11), tmpData(12), tmpData(13), tmpData(14), tmpData(15), tmpData(16), tmpData(17))
        Next
    End With



End Sub

Private Sub Form_Load()
    Dim objDonorTest As clsDonorTest
    Dim strGroup()   As String
    Dim iCnt         As Long
    Dim i            As Long
    
    
    Set objDonorTest = New clsDonorTest
    iCnt = objDonorTest.GetGroup(strGroup)
    
    tabGroup.Tabs.Clear
    For i = 1 To iCnt
        tabGroup.Tabs.Add , strGroup(i - 1), strGroup(i - 1)
    Next i
    
    Set objDonorTest = Nothing
    
    Call FrameInitialize
    Call ClassInitialize
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set objMySQL = Nothing
    Set objMyOrder = Nothing
    Set objMyCollection = Nothing
    Set objCollect = Nothing
End Sub

Private Sub tabAccDt_Click()
    
    Dim donorid As String
    Dim canEdit As Boolean
    
    donorid = lblDonorID.Caption
    Call tabAccdtClickCode(donorid, Format(tabAccDt.SelectedItem.Caption, PRESENTDATE_FORMAT))
    Call SetDonorStatus(donorid, Format(tabAccDt.SelectedItem.Caption, PRESENTDATE_FORMAT))
    
'    canEdit = GetCanEdit
'    fraDonation.Enabled = canEdit

End Sub

Private Sub tabGroup_Click()
    Dim NewTest       As Recordset
    Dim strGroup      As String
    
    '�˻��Ƿڰ� ���� ���� ȯ�ڿ� ���ؼ��� �˻��׸� �����Ϳ���ϵ� �˻��׸��� �����ش�.
    If tabAccDt.Tabs.Count = 0 Then
        Exit Sub
    End If
    
    strGroup = tabGroup.SelectedItem.Key
    
    Set NewTest = objMySQL.GetTestSpc2(strGroup)
    If Not NewTest.EOF Then
        Dim ObjDic As New clsDictionary
        Dim lngseq As Long

        ObjDic.Clear
        ObjDic.FieldInialize "seq", "testcd,spccd"
        Do Until NewTest.EOF
            lngseq = lngseq + 1
            ObjDic.AddNew lngseq, Join(Array(NewTest.Fields("cdval2").value & "", NewTest.Fields("field1").value & ""), COL_DIV)
            NewTest.MoveNext
        Loop
        lblTestChk.Visible = False
        Call Default_Test(ObjDic)
        Set NewTest = Nothing
        Set ObjDic = Nothing
        cmdSave.Enabled = True
        cmdCancel.Enabled = False
    End If
End Sub



Private Sub tblResult_Click(ByVal Col As Long, ByVal Row As Long)
    Dim ii As Integer
    If lblTestChk.Visible = True Then Exit Sub
    
    If Row = 0 And Col = 6 Then
        With tblResult
            
            For ii = 1 To .DataRowCnt
                .Row = ii
                .Col = 6
                If .CellType = CellTypeCheckBox Then .value = IIf(.value = 0, 1, 0)
            Next
        End With
    End If
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
    '�����ڿ� ���ؼ� ������ ������ ���� ��쿡 ���� ������ �����ش�.

'    objMySQL.setDbConn DBConn
    'Set Rs = objMySQL.GetDonorAccHistory(Trim(lblDonorID.Caption))
    
    
    '���������� ������ ������ �˻��Ƿ��Ҽ� �ְ� ��ȸ(2001/10/04, ��� ��������)
    If TEST_FOR_PHERSIS Then
        Set Rs = objMySQL.GetDonorAccHistory(Trim(lblDonorID.Caption))
    Else
        Set Rs = objMySQL.GetDonorAccdtHistoryDivPheresis(Trim(lblDonorID.Caption))
    End If
    
    If Rs.EOF Then
        MsgBox "�˻��Ƿڴ���� �����ϴ�.", vbInformation + vbOKOnly, "�����ڰ˻��Ƿ�"
        
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

Private Sub tabAccdtClickCode(ByVal donorid As String, ByVal donoraccdt As String)
    Dim RsDonorTest   As Recordset
    Dim RsTestReq     As Recordset
    Dim QueryTest     As Recordset
    Dim NewTest       As Recordset
    Dim ii            As Integer
    
    With tblResult
        .Col = 1: .COL2 = .MaxCols
        .Row = 1: .Row2 = .MaxRows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
    End With
    
    '�����ڿ� ���ؼ� �ӻ󺴸��� �˻��Ƿڸ� �Ѱ��� ���ܵȴ�.
    If tabAccDt.SelectedItem.Selected Then
'        objMySQL.setDbConn DBConn
        
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
        
        
        Select Case RsTestReq.Fields("donorcd").value & ""
            Case "0":   cboDonorCd.ListIndex = 0
            Case "1":   cboDonorCd.ListIndex = 1
            Case "2":   cboDonorCd.ListIndex = 2
            Case "3":   cboDonorCd.ListIndex = 3
            Case Else:  cboDonorCd.ListIndex = -1
        End Select
        lblTmpPtId.Caption = RsTestReq.Fields("tmpid").value & ""
        txtReservedID = RsTestReq.Fields("reservedid").value & ""
        lblReservedNm.Caption = objMySQL.GetPtntNm(txtReservedID)
        
        '�˻��Ƿڳ����� �д´�-----------------------------------------
        Set RsDonorTest = objMySQL.Get_TestHistory(donorid, donoraccdt)
        If RsDonorTest.EOF Then
'            'dbconn.DisplayErrors
            Exit Sub
        End If
        
        
        If RsDonorTest.RecordCount > 0 Then
            '�˻��Ƿڳ����� ��ȸ�Ͽ� �����ش�.
            '�̹� �˻��Ƿڰ� ����� ������ ����������
            
'            If RsTestReq.Fields("donationdt") <> "" Then
                cmdSave.Enabled = False
                cmdCancel.Enabled = True
'            Else
'                cmdSave.Enabled = True
'                cmdCancel.Enabled = False
'            End If
                        
            Set QueryTest = objMySQL.GetDonorTestDt(donorid, donoraccdt)
            Dim strTmpID As String
            lblTmpPtId.Caption = RsTestReq.Fields("tmpid").value & ""
            strTmpID = QueryTest.Fields("tmpid").value & ""
            
            'h7lab102���� �˻��Ƿ� ������ �ҷ��´�.
            lblTestChk.Visible = True
            Call QueryInformation(strTmpID)
            Set QueryTest = Nothing
        Else
            lblTestChk.Visible = False
        End If
        
        
        
''''''''''        If RsDonorTest.RecordCount < 1 Then
''''''''''            '�˻��Ƿڰ� ���� ���� ȯ�ڿ� ���ؼ��� �˻��׸� �����Ϳ���ϵ� �˻��׸���
''''''''''            '������ �����ش�.
''''''''''            Set NewTest = objMySQL.GetTestSpc
''''''''''            If Not NewTest.EOF Then
''''''''''                Dim objdic As New clsDictionary
''''''''''                Dim lngseq As Long
''''''''''
''''''''''                objdic.Clear
''''''''''                objdic.FieldInialize "seq", "testcd,spccd"
''''''''''                Do Until NewTest.EOF
''''''''''                    lngseq = lngseq + 1
''''''''''                    objdic.AddNew lngseq, Join(Array(NewTest.Fields("cdval1").value, NewTest.Fields("field1").value), COL_DIV)
''''''''''                    NewTest.MoveNext
''''''''''                Loop
''''''''''                lblTestChk.Visible = False
''''''''''                Call Default_Test(objdic)
''''''''''                Set NewTest = Nothing
''''''''''                Set objdic = Nothing
''''''''''                cmdSave.Enabled = True
''''''''''                cmdCancel.Enabled = False
''''''''''            End If
''''''''''        Else
''''''''''            '�˻��Ƿڳ����� ��ȸ�Ͽ� �����ش�.
''''''''''            '�̹� �˻��Ƿڰ� ����� ������ ����������
''''''''''
'''''''''''            If RsTestReq.Fields("donationdt") <> "" Then
''''''''''                cmdSave.Enabled = False
''''''''''                cmdCancel.Enabled = True
'''''''''''            Else
'''''''''''                cmdSave.Enabled = True
'''''''''''                cmdCancel.Enabled = False
'''''''''''            End If
''''''''''
''''''''''            Set QueryTest = objMySQL.GetDonorTestDt(donorid, donoraccdt)
''''''''''            Dim strTmpID As String
''''''''''            lblTmpPtId.Caption = RsTestReq.Fields("tmpid").value
''''''''''            strTmpID = QueryTest.Fields("tmpid").value
''''''''''
''''''''''            'h7lab102���� �˻��Ƿ� ������ �ҷ��´�.
''''''''''            lblTestChk.Visible = True
''''''''''            Call QueryInformation(strTmpID)
''''''''''            Set QueryTest = Nothing
''''''''''
''''''''''        End If
        
        
        
        
        
        Set RsDonorTest = Nothing
        Set RsTestReq = Nothing
    End If
    
End Sub

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

Private Sub Default_Test(objDefault As clsDictionary)
    Dim objQueryTest As New clsBBSSQLStatement
    Dim objGDic As New clsDictionary
    Dim DefaultTest As Recordset
    Dim strTmp As String
    Dim lngseq As Long
    
'    objQueryTest.setDbConn DBConn
'SpcNm, " & _
                        "        d.field2 as LabDiv, e.field2 as LabRange, '1' InsurFg " & _
    objGDic.Clear
    objGDic.FieldInialize "seq", "ordno1,testnm,testcd,spccd,gubyu,stat,reqdt,statfg,workarea," & _
                          "storecd,rndfg,testdiv,multifg,spcgrp,ordseq1,abbrnm5,labelcnt,statflag,location,testlocation,spcnm,labdiv,labrange,insurfg"
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
                               Join(Array("", .Fields("testnm").value & "", .Fields("testcd").value & "", .Fields("spccd").value & "", _
                                          "1", "", Format(GetSystemDate, "yyyy-MM-dd" & " " & "hh:MM"), .Fields("statfg").value & "", .Fields("workarea").value & "", _
                                          .Fields("storecd").value & "", .Fields("rndfg").value & "", .Fields("testdiv").value & "", .Fields("multifg").value & "", _
                                          .Fields("spcgrp").value & "", "", .Fields("abbrnm5").value & "", _
                                          .Fields("labelcnt").value & "", .Fields("statflags").value & "", "location", "�߾�", _
                                          .Fields("spcnm").value & "", .Fields("labdiv").value & "", .Fields("labrange").value & "", .Fields("insurfg").value & ""), COL_DIV)
            End If
        End With
        objDefault.MoveNext
    Loop
    
    'ȭ�鿡 ��������......
    Call TblResult_Display(objGDic)
    Set objGDic = Nothing
    Set objQueryTest = Nothing
End Sub

Private Sub TblResult_Display(ObjDic As clsDictionary)
'�׽�Ʈ �÷� ����
'1:ó���ȣ,2:�˻��,3:�˻��ڵ�,4:��ü,5:�޿�,6:���޿���,7:���ä���Ͻ�
'8:���޿���:9:WorkArea,10:storecd,11:rndfg,12;testdiv,13:multifg,14:spcgrp,15:ordseq
'16:����,17:���ڵ�������,18:�˻簡�ɿ���,19:Location,20:�˻����
    Dim ii As Integer
    Dim tmpStatFg As String
    Dim tmpTestFg As String
    
    
    If ObjDic.RecordCount < 1 Then Exit Sub
    With tblResult
        .Row = 1: .Row2 = .MaxRows
        .Col = 1: .COL2 = .MaxCols
        .BlockMode = True
        .Action = ActionClear
        .BlockMode = False
        
        ObjDic.MoveFirst
        Do Until ObjDic.EOF
            If .DataRowCnt = .MaxRows Then
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
            Else
                .Row = .DataRowCnt + 1
            End If
            .Col = 1: .value = ObjDic.Fields("ordno1")
            .Col = 2: .value = ObjDic.Fields("testnm")
            .Col = 3: .value = ObjDic.Fields("testcd")
            .Col = 4: .value = ObjDic.Fields("spccd")
            .Col = 5: .value = ObjDic.Fields("gubyu")
            .Col = 6: .CellType = CellTypeCheckBox: .TypeCheckCenter = True
            If ObjDic.Fields("statfg") = "1" Then
                .value = 1
            Else
                .value = 0
            End If
            

'            If objdic.Fields("statfg") = "1" Then
'                .Col = 6: .CellType = CellTypeCheckBox
'                   .TypeCheckCenter = True
'            Else
'                .Col = 6: .CellType = CellTypeStaticText
'            End If
            If Len(ObjDic.Fields("reqdt")) = 14 Then
                .Col = 7: .value = Format(Mid(ObjDic.Fields("reqdt"), 1, 12), "####-##-## ##:##")
            Else
                .Col = 7: .value = ObjDic.Fields("reqdt")
            End If
            .Col = 8: .value = ObjDic.Fields("statfg")
            .Col = 9: .value = ObjDic.Fields("workarea")
            .Col = 10: .value = ObjDic.Fields("storecd")
            .Col = 11: .value = ObjDic.Fields("rndfg")
            .Col = 12: .value = ObjDic.Fields("testdiv")
            .Col = 13: .value = ObjDic.Fields("multifg")
            .Col = 14: .value = ObjDic.Fields("spcgrp")
            .Col = 15: .value = ObjDic.Fields("ordseq1")
            .Col = 16: .value = ObjDic.Fields("abbrnm5")
            .Col = 17: .value = ObjDic.Fields("labelcnt")
            '
            .Col = 21: .value = ObjDic.Fields("spcnm")
            .Col = 22: .value = ObjDic.Fields("labdiv")
            .Col = 23: .value = ObjDic.Fields("labrange")
            .Col = 24: .value = ObjDic.Fields("insurfg")
            
            tmpStatFg = medGetP(ObjDic.Fields("statflag"), 1, ";")  '�ǹ��� ���ް��� ����
            tmpTestFg = medGetP(ObjDic.Fields("statflag"), 2, ";")  '�ǹ��� �˻簡�� ����
'    '***�ǹ����� ���
'        If ObjSysInfo.UseBuildingInfo = "1" Then
'
'            .Col = enORDSHEET.tcSTATCHK
'            If .value = 1 Then   '���޼���
'                .Col = enORDSHEET.tcSTATFG
'                If .value = "1" Then
'
'                    ' ** �߾�/���̰˻�ǿ��� ���ް˻簡 �߻��ϸ� --> ���޼��ͷ�...
'                    If ObjSysInfo.BuildingCd = CentralLab Or ObjSysInfo.BuildingCd = AneLab Then
'                        .Col = enORDSHEET.tcBUILDCD: .value = EmergencyLab
'                        .Col = enORDSHEET.tcBUILDNM: .value = EmergencyLabNm
'
'                    ' ** �ش�ǹ����� ���ް˻� ������
'                    Else
'                        .Col = enORDSHEET.tcBUILDCD: .value = ObjSysInfo.BuildingCd
'                        .Col = enORDSHEET.tcBUILDNM: .value = ObjSysInfo.BuildingNm
'                    End If
'                    Exit Sub
'
'                Else
'                    ' ** �ش�ǹ����� ���ް˻� �Ұ���...
'                    .Col = enORDSHEET.tcSTATCHK
'                    .CellType = CellTypeStaticText
'                    .Text = ""
'                End If
'            End If
'
'            '** �Ϲݰ˻� ���ɿ���
'            .Col = enORDSHEET.tcTESTFLAG
'
'            ' ** �ش�ǹ����� �Ϲݰ˻� ������
'            If .value = "1" Then
'                .Col = enORDSHEET.tcBUILDCD: .value = ObjSysInfo.BuildingCd
'                .Col = enORDSHEET.tcBUILDNM: .value = ObjSysInfo.BuildingNm
'
'            ' ** �ش�ǹ����� �Ϲݰ˻� �Ұ����� --> �߾Ӱ˻�Ƿ�...
'            Else
'                .Col = enORDSHEET.tcBUILDCD: .value = CentralLab
'                .Col = enORDSHEET.tcBUILDNM: .value = CentralLabNm
'            End If
'
'    '***�ǹ����� ������� ����
'        Else
'            .Col = enORDSHEET.tcBUILDCD: .value = ObjSysInfo.BuildingCd
'            .Col = enORDSHEET.tcBUILDNM: .value = ObjSysInfo.BuildingNm
'        End If
             If ObjSysInfo.UseBuildingInfo = "1" Then
                If ObjDic.Fields("statfg") = "1" Then
                    .Col = 18: .value = Mid(tmpStatFg, ObjSysInfo.BuildingNo, 1)
                    If .value = "1" Then
                        If ObjSysInfo.BuildingCd = "10" Or ObjSysInfo.BuildingCd = "40" Then
                            .Col = 19: .value = "50"
                            .Col = 20: .value = "����"
                        Else
                            .Col = 19: .value = ObjSysInfo.BuildingCd
                            .Col = 20: .value = ObjSysInfo.BuildingNm
                        End If
                    Else
                        If ObjSysInfo.BuildingCd = "20" Or ObjSysInfo.BuildingCd = "30" Then
                            If Mid(tmpStatFg, 5, 1) = "1" Then
                                .Col = 19: .value = "50"
                                .Col = 20: .value = "����"
                            Else
                            End If
                        Else
                            .Col = 18: .value = Mid(tmpTestFg, ObjSysInfo.BuildingNo, 1)
                            If .value = "1" Then
                                .Col = 19: .value = ObjSysInfo.BuildingCd
                                .Col = 20: .value = ObjSysInfo.BuildingNm
                            Else
                                .Col = 19: .value = "10"
                                .Col = 20: .value = "�߾�"
                            End If
                            .Col = 8: .value = "0"
                        End If
                    End If
                Else
                    .Col = 18: .value = Mid(tmpTestFg, ObjSysInfo.BuildingNo, 1)
                    If .value = "1" Then
                        .Col = 19: .value = ObjSysInfo.BuildingCd
                        .Col = 20: .value = ObjSysInfo.BuildingNm
                    Else
                        .Col = 19: .value = "10"
                        .Col = 20: .value = "�߾�"
                    End If
                End If
            Else
                .Col = 19: .value = ObjSysInfo.BuildingCd
                .Col = 20: .value = ObjSysInfo.BuildingNm
            End If
        
            ObjDic.MoveNext
        Loop
    End With
    
            
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
    
'    objMySQL.setDbConn DBConn
    
    objDicT.Clear
    objDicT.FieldInialize "ptid,orddt,ordno,ordseq", "ordcd,spccd,reqdate,statfg"
    
    
    Set RsDonorTest = objMySQL.GetDonnorTest(tmpid)
    
    If Not RsDonorTest.EOF Then
        Do Until RsDonorTest.EOF
            objDicT.AddNew Join(Array(RsDonorTest.Fields("ptid").value & "", RsDonorTest.Fields("orddt").value & "", RsDonorTest.Fields("ordno").value & "", _
                                RsDonorTest.Fields("ordseq").value & ""), COL_DIV), Join(Array(RsDonorTest.Fields("ordcd").value & "", _
                                RsDonorTest.Fields("spccd").value & "", RsDonorTest.Fields("reqdt").value & "" & RsDonorTest.Fields("reqtm").value & "", RsDonorTest.Fields("statfg").value & ""), COL_DIV)
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
            If Not RsDisplay.EOF Then
                With RsDisplay
                    objDicD.AddNew Join(Array(objDicT.Fields("ptid"), objDicT.Fields("orddt"), objDicT.Fields("ordno"), objDicT.Fields("ordseq")), COL_DIV), _
                                   Join(Array(objDicT.Fields("ordno"), .Fields("testnm").value & "", .Fields("testcd").value & "", .Fields("spccd").value & "", _
                                              "1", "", objDicT.Fields("reqdate"), objDicT.Fields("statfg"), .Fields("workarea").value & "", _
                                              .Fields("storecd").value & "", .Fields("rndfg").value & "", .Fields("testdiv").value & "", .Fields("multifg").value & "", _
                                              .Fields("spcgrp").value & "", objDicT.Fields("ordseq"), .Fields("abbrnm5").value & "", _
                                              .Fields("labelcnt").value & "", .Fields("statflags").value & "", "location", "�߾�"), COL_DIV)
                End With
            End If
            
            objDicT.MoveNext
        Loop
    End If
    'ȭ�鿡 ��������......
    Call TblResult_Display(objDicD)
    '''
    
    Set objDicD = Nothing
End Sub

Private Sub TblSort()
    With tblResult
        .SortBy = SortByRow
        .SortKey(1) = 19  'DeliveryLocation
        .SortKey(2) = 7   '���ä��ð�
        .SortKey(3) = 9   'WorkArea
        .SortKey(4) = 4   '��ü�ڵ�
        .SortKey(5) = 10  '��������
        .SortKey(6) = 6   '���޿���
        .SortKey(7) = 3   '�˻��ڵ�
        .SortKeyOrder(1) = SortKeyOrderAscending
        .SortKeyOrder(2) = SortKeyOrderAscending
        .SortKeyOrder(3) = SortKeyOrderAscending
        .SortKeyOrder(4) = SortKeyOrderAscending
        .SortKeyOrder(5) = SortKeyOrderAscending
        .SortKeyOrder(6) = SortKeyOrderAscending
        .SortKeyOrder(7) = SortKeyOrderAscending
        .Col = 1: .COL2 = .MaxCols
        .Row = 0: .Row2 = .MaxRows
        .Action = ActionSort
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
    Set objMyCollection = New clsDonorTestCollection
    
    With objMyCollection
        .DateTime = datDateTime
'        .setDbConn DBConn
        Set .SortList = frmControls.lstList
        Call .InitRtn
    End With
End Sub
'2001-11-27�߰�
Public Sub CallDonorNmLostFocus()
    Call txtDonorNm_LostFocus
End Sub



