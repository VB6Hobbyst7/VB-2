VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm168POCTCol 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  'ũ�� ���� ��ȭ ����
   ClientHeight    =   9435
   ClientLeft      =   60
   ClientTop       =   1500
   ClientWidth     =   14655
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Lis168.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9435
   ScaleWidth      =   14655
   ShowInTaskbar   =   0   'False
   Tag             =   "����ȯ�� �ϰ� ä��"
   Begin VB.ComboBox cboRemark 
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1395
      TabIndex        =   50
      Text            =   "Combo1"
      Top             =   1035
      Visible         =   0   'False
      Width           =   2400
   End
   Begin VB.CommandButton cmdCancle 
      BackColor       =   &H00E0E0E0&
      Caption         =   "�������(&C)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   45
      Style           =   1  '�׷���
      TabIndex        =   49
      Tag             =   "0"
      Top             =   1000
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.CheckBox chkCancle 
      BackColor       =   &H00DBE6E6&
      Caption         =   "�������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3870
      TabIndex        =   48
      Top             =   1050
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.ComboBox cboBarCode 
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1380
      TabIndex        =   47
      Text            =   "Combo1"
      Top             =   1290
      Visible         =   0   'False
      Width           =   1365
   End
   Begin MSComctlLib.ProgressBar pbrPtCnt 
      Height          =   300
      Left            =   3720
      TabIndex        =   46
      Top             =   1560
      Width           =   10725
      _ExtentX        =   18918
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CheckBox chkAction 
      BackColor       =   &H00DBE6E6&
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   7830
      TabIndex        =   45
      Top             =   1050
      Width           =   1155
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00800000&
      Caption         =   "��ü���ܼ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   2040
      TabIndex        =   44
      Top             =   1560
      Width           =   1505
   End
   Begin VB.CommandButton cmdAccept 
      BackColor       =   &H00E0E0E0&
      Caption         =   "�ϰ�����(&A)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   9060
      Style           =   1  '�׷���
      TabIndex        =   43
      Tag             =   "0"
      Top             =   1000
      Width           =   1320
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "����Ʈ���(&P)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   6480
      Style           =   1  '�׷���
      TabIndex        =   42
      Tag             =   "0"
      Top             =   1000
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.CommandButton cmdGetOrders 
      BackColor       =   &H00E0E0E0&
      Caption         =   "��ȸ(&F)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   10440
      Style           =   1  '�׷���
      TabIndex        =   40
      Tag             =   "0"
      Top             =   1000
      Width           =   1320
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   1425
      Left            =   14760
      TabIndex        =   20
      Top             =   120
      Visible         =   0   'False
      Width           =   3015
      Begin VB.CheckBox chkCol 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Ư��ä��ð���ȸ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -120
         TabIndex        =   29
         Top             =   0
         Width           =   1980
      End
      Begin VB.OptionButton optApplyColTm 
         BackColor       =   &H00DBE6E6&
         Caption         =   "���� Row�� ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   1170
         TabIndex        =   24
         Top             =   405
         Width           =   1710
      End
      Begin VB.OptionButton optApplyColTm 
         BackColor       =   &H00DBE6E6&
         Caption         =   "��ü����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   45
         TabIndex        =   23
         Top             =   405
         Width           =   1035
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   21
         Top             =   765
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   344
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         AutoSize        =   -1  'True
         Caption         =   "ä���Ͻ�"
         Appearance      =   0
      End
      Begin MSComCtl2.DTPicker dtpColDtTm 
         Height          =   315
         Left            =   915
         TabIndex        =   22
         Top             =   720
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyy-MM-dd  HH:mm"
         Format          =   85721091
         UpDown          =   -1  'True
         CurrentDate     =   36328.5416666667
      End
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00E0E0E0&
      Caption         =   "�� ��(&X)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   13200
      Style           =   1  '�׷���
      TabIndex        =   2
      Tag             =   "0"
      Top             =   1000
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ȭ������(&C)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   11820
      Style           =   1  '�׷���
      TabIndex        =   1
      Tag             =   "0"
      Top             =   1000
      Width           =   1320
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00E0E0E0&
      Caption         =   "�ϰ�ä�� (&S)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   9060
      Style           =   1  '�׷���
      TabIndex        =   0
      Tag             =   "0"
      Top             =   1000
      Width           =   1320
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   300
      Left            =   0
      TabIndex        =   3
      Top             =   45
      Width           =   14500
      _ExtentX        =   25585
      _ExtentY        =   529
      BackColor       =   8388608
      ForeColor       =   16777215
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
      Caption         =   "���� ����"
      LeftGab         =   100
   End
   Begin MedControls1.LisLabel LisLabel6 
      Height          =   300
      Left            =   0
      TabIndex        =   4
      Top             =   1560
      Width           =   14505
      _ExtentX        =   25585
      _ExtentY        =   529
      BackColor       =   8388608
      ForeColor       =   16777215
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
      Caption         =   "��ü ä�� ����Ʈ"
      LeftGab         =   100
   End
   Begin VB.Frame fraPrtOption 
      BackColor       =   &H00DBE6E6&
      Height          =   2100
      Left            =   14760
      TabIndex        =   5
      Top             =   1680
      Visible         =   0   'False
      Width           =   5130
      Begin VB.CheckBox chkPrintFg 
         BackColor       =   &H00DBE6E6&
         Caption         =   "��¾���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   300
         TabIndex        =   26
         Top             =   315
         Width           =   1305
      End
      Begin VB.CheckBox chkTestdiv 
         BackColor       =   &H00DBE6E6&
         Caption         =   "�˻��ڵ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3480
         TabIndex        =   10
         Top             =   765
         Width           =   1425
      End
      Begin VB.OptionButton optOption 
         BackColor       =   &H00DBE6E6&
         Caption         =   "���ڵ�Lable And ä�� ����Ʈ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   360
         TabIndex        =   9
         Top             =   750
         Width           =   3180
      End
      Begin VB.OptionButton optOption 
         BackColor       =   &H00DBE6E6&
         Caption         =   "���ڵ� Only"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   360
         TabIndex        =   8
         Top             =   1140
         Width           =   3180
      End
      Begin VB.TextBox txtCopy 
         Alignment       =   1  '������ ����
         Height          =   345
         Left            =   3255
         TabIndex        =   7
         Top             =   1515
         Visible         =   0   'False
         Width           =   750
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   360
         Left            =   4020
         TabIndex        =   6
         Top             =   1500
         Visible         =   0   'False
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MedControls1.LisLabel lblColList 
         Height          =   255
         Left            =   855
         TabIndex        =   11
         Top             =   1545
         Visible         =   0   'False
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   450
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   "ä�븮��Ʈ ������"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblPage 
         Height          =   255
         Left            =   4335
         TabIndex        =   12
         Top             =   1575
         Visible         =   0   'False
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   "��"
         Appearance      =   0
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DBE6E6&
      Height          =   600
      Left            =   0
      TabIndex        =   16
      Top             =   390
      Width           =   14500
      Begin VB.CommandButton cmdTestList 
         BackColor       =   &H0098A7A5&
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4755
         Style           =   1  '�׷���
         TabIndex        =   37
         Tag             =   "WardID"
         Top             =   180
         Width           =   360
      End
      Begin VB.TextBox txtTestCd 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3525
         MaxLength       =   9
         TabIndex        =   36
         Top             =   180
         Width           =   1205
      End
      Begin VB.CommandButton cmdWardList 
         BackColor       =   &H0098A7A5&
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   9900
         Style           =   1  '�׷���
         TabIndex        =   18
         Tag             =   "WardID"
         Top             =   180
         Width           =   360
      End
      Begin VB.TextBox txtWardID 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8895
         MaxLength       =   9
         TabIndex        =   17
         Top             =   180
         Width           =   995
      End
      Begin MSComCtl2.DTPicker dtpToTime 
         Height          =   360
         Left            =   820
         TabIndex        =   19
         Top             =   180
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyy-MM-dd"
         Format          =   85721088
         CurrentDate     =   36803
      End
      Begin MedControls1.LisLabel lblWardNm 
         Height          =   360
         Left            =   10275
         TabIndex        =   25
         Top             =   180
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   635
         BackColor       =   13622494
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
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   0
         Left            =   8160
         TabIndex        =   30
         Top             =   180
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   635
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
         Caption         =   "����ID"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   1
         Left            =   105
         TabIndex        =   31
         Top             =   180
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   635
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
         Caption         =   "ó����"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   5
         Left            =   2565
         TabIndex        =   35
         Top             =   180
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   635
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
         Caption         =   "�˻��׸�"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblTestNm 
         Height          =   360
         Left            =   5175
         TabIndex        =   38
         Top             =   180
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   635
         BackColor       =   13622494
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
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   6
         Left            =   12400
         TabIndex        =   39
         Top             =   180
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   635
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
         Caption         =   "��ü�Ǽ�"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblTotalCnT 
         Height          =   360
         Left            =   13400
         TabIndex        =   41
         Top             =   180
         Width           =   1000
         _ExtentX        =   1773
         _ExtentY        =   635
         BackColor       =   13622494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
   End
   Begin Crystal.CrystalReport CReport 
      Left            =   15720
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DBE6E6&
      Height          =   5460
      Left            =   14760
      TabIndex        =   13
      Top             =   3840
      Visible         =   0   'False
      Width           =   5100
      Begin FPSpread.vaSpread tblCount 
         Height          =   5340
         Left            =   2100
         TabIndex        =   14
         Top             =   105
         Width           =   2955
         _Version        =   196608
         _ExtentX        =   5212
         _ExtentY        =   9419
         _StockProps     =   64
         BackColorStyle  =   1
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
         GrayAreaBackColor=   14411494
         MaxCols         =   3
         MaxRows         =   50
         OperationMode   =   1
         ScrollBars      =   2
         ShadowColor     =   14737632
         ShadowDark      =   14737632
         SpreadDesigner  =   "Lis168.frx":08CA
      End
      Begin MedControls1.LisLabel LisLabel3 
         Height          =   255
         Index           =   5
         Left            =   1350
         TabIndex        =   15
         Top             =   1680
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   "��"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblColNm 
         Height          =   330
         Left            =   360
         TabIndex        =   27
         Top             =   750
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   582
         BackColor       =   13752531
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
      Begin MedControls1.LisLabel lblPtCount 
         Height          =   330
         Left            =   360
         TabIndex        =   28
         Top             =   1635
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   582
         BackColor       =   13752531
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
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   3
         Left            =   360
         TabIndex        =   32
         Top             =   375
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   635
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
         Caption         =   "�� ä����"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   4
         Left            =   360
         TabIndex        =   33
         Top             =   1260
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   635
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
         Caption         =   "�� ȯ�ڼ�"
         Appearance      =   0
      End
   End
   Begin FPSpread.vaSpread tblPtList 
      Height          =   7335
      Left            =   0
      TabIndex        =   34
      Top             =   1920
      Width           =   14490
      _Version        =   196608
      _ExtentX        =   25559
      _ExtentY        =   12938
      _StockProps     =   64
      BackColorStyle  =   1
      ColsFrozen      =   4
      DisplayRowHeaders=   0   'False
      EditEnterAction =   5
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   14411494
      MaxCols         =   28
      MaxRows         =   50
      ScrollBarExtMode=   -1  'True
      ScrollBars      =   2
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      SpreadDesigner  =   "Lis168.frx":1511
      TextTip         =   4
      ScrollBarTrack  =   1
   End
End
Attribute VB_Name = "frm168POCTCol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'** ���� :  �ǹ������� OCS���α׷����� �Ѱ��� �μ��ڵ�� �μ������͸� �˻��ؼ�
'           bld_gb�� �����´�.

Option Explicit

'---- Collect
Private objMySql                As clsLISSqlCollection
Private objLISCollect           As clsLISCollectioin
Private objLISQc                As clsLISSqlQc
Private WithEvents objMyList    As clsPopUpList
Attribute objMyList.VB_VarHelpID = -1

Private IsFirst         As Boolean
Private blnCleanFg      As Boolean
Private blnCollectFg    As Boolean             'ä������(�Ѱ��̶��...�Ǹ� True)
Private sWorkDt         As String
Private sWorkTm         As String


Private intPtCount      As Integer
Private intErrCount     As Integer

Public Event LastFormUnload()
Private objCanAcc As New clsLISAccCancel

Private Sub Check1_Click()
    Dim iCnt As Integer
    
    With tblPtList
        For iCnt = 1 To .MaxRows
            .Row = iCnt
            .Col = 1
            If Check1.Value = 1 Then
                .Value = 1
            Else
                .Value = 0
            End If
        Next
    End With
        
End Sub

Private Sub chkAction_Click()
    If chkAction.Value = 0 Then
        cmdSave.Enabled = True
        cmdSave.Visible = True
        cmdAccept.Enabled = False
        cmdAccept.Visible = False
        chkAction.Caption = "�ϰ�ä��"
    Else
        cmdAccept.Enabled = True
        cmdAccept.Visible = True
        cmdSave.Enabled = False
        cmdSave.Visible = False
        chkAction.Caption = "�ϰ�����"
    End If

End Sub

Private Sub chkCancle_Click()
    If chkCancle.Value = 1 Then
        cmdCancle.Enabled = True
        chkAction.Enabled = False
        cmdAccept.Enabled = False
        cmdSave.Enabled = False
    Else
        cmdCancle.Enabled = False
        chkAction.Enabled = True
        cmdAccept.Enabled = True
        cmdSave.Enabled = True
    End If
End Sub

Private Sub chkCol_Click()
    If chkCol.Value = 0 Then
        dtpColDtTm.Value = GetSystemDate
        dtpColDtTm.Enabled = False
    Else
        dtpColDtTm.Enabled = True
    End If
    
End Sub

Private Sub cmdAccept_Click()
    
    Dim i           As Integer
'    Dim SelCount    As Integer
'    Dim CollectCnt  As Integer
    Dim ColSuccess  As Boolean
'    Dim objProgress As clsProgress
    
'    Set objProgress = New clsProgress
    
    ColSuccess = True

    '** ��������:���PCx(��������) ���Կ� ���� ä��-���� ��ƾ �ʿ� (�ܷ�ä���ǿ��� �ٷ� ������ �ϱ� ����)
    '   �߰� By M.G.Choi 2007.04.02
    '---------------------------------------------------------------------------------------------------------------------
'        objProgress.Message = "���� Procedure�� �����ϰ� �ֽ��ϴ�."
        Dim objAccess   As New clsLISAccession
        Dim pWorkArea As String
        Dim pAccDt As String
        Dim pAccSeq As Integer
        
    With tblPtList
        For i = 1 To .MaxRows
            .Row = i
            .Col = 19:  pWorkArea = .Value                             ' WorkArea
            .Col = 24:  pAccDt = .Value                                ' AccDt
            .Col = 25:  pAccSeq = Val(.Value)                          ' AccSeq

'            objProgress.Message = "���� Procedure�� �����ϰ� �ֽ��ϴ�. (" & CStr(i) & "/" & CStr(.MaxRows) & ")"
                    
            ColSuccess = objAccess.DoAccession(pWorkArea, pAccDt, pAccSeq, ObjMyUser.EmpId)
            If Not ColSuccess Then Exit For
'            If objProgress.Value = objProgress.Max Then objProgress.Max = objProgress.Max + 1
'            objProgress.Value = objProgress.Value + 1
            DoEvents
        Next
    End With
'    Set objProgress = Nothing
    Set objAccess = Nothing
    
    '----------------------------------------------------------------------------------------------------------------------
    If Not ColSuccess Then
        MsgBox "ä��ó���� ������ �߻��߽��ϴ� !!"
        MouseDefault  '0
        Exit Sub
    End If

    Call ClearRtn(0)
On Error GoTo Err_Trap
        txtWardID.SetFocus
    Me.MousePointer = 0
Err_Trap:

End Sub

Private Sub cmdCancle_Click()

    Dim i As Integer
    Dim strStsCd As String
    Dim blnCancel As Boolean
    Dim lngCnt As Long
    Dim sFlag   As Long
    Dim strReason As String
    Dim varTmp As Variant
    Dim strPtid, strWorkArea, strAccdt, strAccSeq, strOrdDt, strOrdNo, strOrdSeq As String
    
    strReason = cboRemark.Text
    
    If strReason = "" Then
        MsgBox "��һ����� �Է��ϼ���.", vbInformation, "��һ�������"
        cboRemark.SetFocus
        Exit Sub
    End If
    'ó�����, ä������
    strStsCd = "7"
    sFlag = "1"
    
    With tblPtList
        For i = 1 To .MaxRows
            .Col = 1
            .Row = i
            If .Value = 0 Then
                DBConn.BeginTrans
                .GetText 4, i, varTmp: strPtid = varTmp
                strWorkArea = "14"
                .GetText 24, i, varTmp: strAccdt = varTmp
                .GetText 25, i, varTmp: strAccSeq = varTmp
                .GetText 26, i, varTmp: strOrdDt = varTmp
                .GetText 27, i, varTmp: strOrdNo = varTmp
                .GetText 28, i, varTmp: strOrdSeq = varTmp
                
                blnCancel = objCanAcc.DoCancelAccession_POCT(strPtid, strWorkArea, strAccdt, strAccSeq, _
                                                 strStsCd, ObjSysInfo.EmpId, strReason, strOrdDt, strOrdNo, strOrdSeq, sFlag)
                If Not blnCancel Then GoTo Err_Trap
                DBConn.CommitTrans
            End If
        Next
    End With
    Call cmdGetOrders_Click
    DoEvents
    
    Exit Sub
    
Err_Trap:
    DBConn.RollbackTrans
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub cmdClear_Click()
    Call ClearRtn(1)
    On Error GoTo Err_Trap
    txtWardID.SetFocus
Err_Trap:
End Sub

Private Sub cmdTestList_Click()
'% �˻��ڵ� ����Ʈ�� �˾��Ѵ�.
'    Dim objWard As clsBasisData
    Dim strSQL As String
    Set objLISQc = New clsLISSqlQc

    strSQL = objLISQc.GetTestItem("14", False)

    Set objLISQc = Nothing
    
    Set objMyList = New clsPopUpList
    
    With objMyList
        .Connection = DBConn
        .FormCaption = "�˻��ڵ� ��ȸ"
        .ColumnHeaderText = "�˻��ڵ�;�˻��"
        Call .LoadPopUp(strSQL)
        If .SelectedString <> "" Then
            txtTestCd.Text = medGetP(.SelectedString, 1, ";")
            lblTestNm.Caption = medGetP(.SelectedString, 2, ";")
        End If
    End With
    
    Set objMyList = Nothing

End Sub

Private Sub dtpColDtTm_Change()

    Dim Resp As VbMsgBoxResult
    
    If blnCleanFg Then Exit Sub
    If dtpColDtTm.Value < Now Then
        If dtpColDtTm.Tag = "1" Then
            dtpColDtTm.Tag = "0"
        Else
            Resp = MsgBox("ä��ð��� ����ð����� �����Դϴ�. �����Ͻðڽ��ϱ�?", _
                   vbQuestion + vbYesNo, "ä��ð�����")
            If Resp = vbYes Then
                dtpColDtTm.Tag = "1"
            Else
                dtpColDtTm.Tag = "0"
                dtpColDtTm.Value = GetSystemDate
            End If
        End If
    End If

    With tblPtList
        If optApplyColTm(0).Value Then  '��ü
            .Row = 1: .Row2 = .DataRowCnt
            .Col = 19: .Col2 = 19
            .BlockMode = True
            .Text = Format(dtpColDtTm.Value, "YY-MM-DD")
            .BlockMode = False
            .Col = 20: .Col2 = 20
            .BlockMode = True
            .Text = Format(dtpColDtTm.Value, "HH:MM")
            .BlockMode = False
        Else
            If .ActiveRow > .DataRowCnt Then Exit Sub
            .Row = .ActiveRow
            .Col = 19: .Text = Format(dtpColDtTm.Value, "YY-MM-DD")
            .Col = 20: .Text = Format(dtpColDtTm.Value, "HH:MM")
        End If
    End With

End Sub

Private Sub dtpToTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Activate()
    If Not IsFirst Then Exit Sub
    
    IsFirst = False
    dtpColDtTm.Enabled = False
    txtCopy.Text = 1
    dtpToTime.Value = GetSystemDate
    dtpColDtTm.Value = GetSystemDate
    blnCleanFg = True
    intErrCount = 0
    txtWardID.Text = ""
    
    cmdCancle.Enabled = False
    
On Error GoTo Err_Trap
    txtWardID.SetFocus
    chkPrintFg.Value = 0
    optOption(0).Value = True
    
Err_Trap:
    Resume Next
End Sub

Private Sub Form_Load()
    IsFirst = True
    chkAction.Value = 0
    chkAction.Caption = "�ϰ�ä��"
    
        cmdSave.Enabled = True
        cmdSave.Visible = True
        cmdAccept.Enabled = False
        cmdAccept.Visible = False
        chkAction.Caption = "�ϰ�ä��"
    
    pbrPtCnt.Visible = False
    If P_MornCollection = False Then
'        ChkMornFg.Visible = False
        chkCol.Visible = False
    Else
        chkCol.Visible = False
        optApplyColTm(0).Visible = False
        optApplyColTm(1).Visible = False
    End If
   Set objMySql = New clsLISSqlCollection
   Set objLISCollect = New clsLISCollectioin
   
   cboBarCode.AddItem 0
   cboBarCode.AddItem 500
   cboBarCode.AddItem 1000
   cboBarCode.AddItem 1500
   cboBarCode.AddItem 2000
   cboBarCode.AddItem 2500
   cboBarCode.AddItem 3000
   cboBarCode.ListIndex = 1
   
   cboRemark.AddItem ""
   cboRemark.AddItem "ȯ�ں�����"
   cboRemark.AddItem "ȯ��ä���ź�"
   cboRemark.AddItem "���� ��ȣ��� �ǽ�"
   cboRemark.AddItem "�ߺ�ó��"
   cboRemark.AddItem "���"
   cboRemark.ListIndex = 0

End Sub


'& ��� Option ����
Private Sub chkPrintFg_Click()
    If chkPrintFg.Value = 1 Then
        optOption(0).Value = False
        optOption(1).Value = False
    Else
        optOption(1).Value = True
    End If
End Sub

'% ����
Private Sub cmdExit_Click()
    Unload Me
    Set objMySql = Nothing
    Set objLISCollect = Nothing
    If IsLastForm Then RaiseEvent LastFormUnload
    Unload Me
End Sub

'% �ϰ�ä�� ����
Private Sub cmdSave_Click()
    Dim Resp        As VbMsgBoxResult
    Dim intSelCount As Integer
    Dim sBuildCd    As String
    Dim sBuildNm    As String

    Dim strSavePtId As String
    
    Dim i           As Integer
    
    If tblPtList.DataRowCnt = 0 Then Exit Sub
    
    cmdSave.Enabled = False
    
    blnCollectFg = False
    Set objLISCollect = New clsLISCollectioin

    sWorkDt = Format(GetSystemDate, CS_DateDbFormat)
    sWorkTm = Format(GetSystemDate, CS_TimeDbFormat)

    tblCount.Row = 0
    intErrCount = 0
    intSelCount = 0
    strSavePtId = ""

    Call SetLock(True)

    Me.MousePointer = 11
    
    pbrPtCnt.Visible = True

    With tblPtList
        pbrPtCnt.Visible = True
        pbrPtCnt.Max = .DataRowCnt * 3 * 101
        pbrPtCnt.Min = 0
        lblPtCount.Caption = ""

        For i = 1 To .DataRowCnt
            .Row = i

            '* ���ܹ�ư Check
            .Col = 1: If .Value = 1 Then GoTo Skip

            intSelCount = intSelCount + 1

            '* ä������
'            .Col = 15   'for LIS
'            If Trim(.Value) <> "" Then
'                MsgBox " Call DoCollectionForLIS "
                Call DoCollectionForLIS(i)
'            End If
            
            If pbrPtCnt.Value + 100 >= pbrPtCnt.Max Then pbrPtCnt.Max = pbrPtCnt.Value + 100
            pbrPtCnt.Value = pbrPtCnt.Value + 100
            DoEvents
            .Col = 16
            
            .Col = 17   'for BBS
'            If Trim(.Value) <> "" Then Call DoCollectionForBBS(i)
'
'            If pbrPtCnt.Value + 100 >= pbrPtCnt.Max Then pbrPtCnt.Max = pbrPtCnt.Value + 100
'            pbrPtCnt.Value = pbrPtCnt.Value + 100
'            DoEvents


            '* ȯ�ڼ� Count
            .Row = i: .Col = 3
            If strSavePtId <> Trim(.Text) Then
               lblPtCount.Caption = Val(lblPtCount.Caption) + 1
               strSavePtId = .Text
            End If

            '* ä�� Class Initialize
            objLISCollect.InitRtn
            DoEvents
Skip:
        Next

        'ä����
        lblColNm.Caption = gEmpId

    End With

    If intSelCount = 0 Then
         Screen.MousePointer = vbDefault  '0
         Call cmdClear_Click
         MsgBox "ó���� ����Ÿ�� �����ϴ�..", vbInformation, "Message"
         cmdSave.Enabled = True
         Exit Sub
    End If
    
    If blnCollectFg = True Then
    
        pbrPtCnt.Value = pbrPtCnt.Max
        DoEvents
    
        MouseDefault
    
        If intErrCount > 0 Then
             MsgBox CStr(intErrCount) & "���� ������ �߻��߽��ϴ�.."
        Else
        
             If optOption(0).Value Then
                 Call medClearTable(tblPtList)
                 Resp = MsgBox("��� ���������� ä��ó�� �Ǿ����ϴ�.." & vbCrLf & _
                               "ä�븮��Ʈ�� ���� ����Ͻðڽ��ϱ� ? ", vbYesNo, "ä�븮��Ʈ ���")
                 If Resp = vbYes Then
                     For i = 1 To tblCount.DataRowCnt
                         tblCount.Row = i
                         tblCount.Col = 3:  sBuildCd = tblCount.Value
                         tblCount.Col = 1:  sBuildNm = tblCount.Value
                         Call PrintColList(txtWardID.Text, lblWardNm.Caption, sWorkDt, sWorkTm, sBuildCd, sBuildNm)
                         
                     Next
                 End If
             Else
                 Call MsgBox("��� ���������� ä��ó�� �Ǿ����ϴ�..", vbInformation, "�޼���")
             End If
    
             Call ClearRtn(0)
             On Error GoTo Err_Trap
             txtWardID.SetFocus
        End If
    Else
        Call ClearRtn(0)
On Error GoTo Err_Trap
        txtWardID.SetFocus
    End If
    
    cmdSave.Enabled = True
    
    pbrPtCnt.Visible = False
    Me.MousePointer = 0
Err_Trap:
    pbrPtCnt.Visible = False
    cmdSave.Enabled = True

End Sub

Private Sub SetLock(ByVal blnLock As Boolean)
    'Locking...
    txtWardID.Enabled = Not blnLock
    txtWardID.BackColor = IIf(blnLock, &H8000000F, vbWhite)
    cmdWardList.Enabled = Not blnLock
    dtpToTime.Enabled = Not blnLock
    cmdGetOrders.Enabled = Not blnLock
End Sub

Private Sub BarCode_Print(objDIC As clsDictionary)
    Dim objBar       As clsBarcode
    Dim strBuildNm  As String        '�ǹ��̸�
    Dim strPtid     As String
    Dim strPtnm     As String
    Dim strColDt    As String
    Dim strColTm    As String
    Dim strSpcNo    As String
    Dim strAccSeq   As String         'SpcYy-SpcNo ������ ��ü��ȣ
    Dim HosilId     As String
    Dim strStatFg   As String
    Dim strBarW_H   As String
    
    
    Set objBar = New clsBarcode
    
''    Set objBAR.MyDB = dbconn
    Set objBar.TableInfo = New clsTables
    Set objBar.FieldInfo = New clsFields
    

    strBuildNm = "����"

    objDIC.MoveFirst

    Do Until objDIC.EOF
        strPtid = medGetP(objDIC.GetString, 1, COL_DIV)
        strPtnm = medGetP(objDIC.GetString, 2, COL_DIV)
        strSpcNo = medGetP(objDIC.GetString, 3, COL_DIV)
        strColDt = medGetP(objDIC.GetString, 4, COL_DIV)
        strColDt = Format(Mid(strColDt, 5, 4), "##/##")
        strColTm = Mid(medGetP(objDIC.GetString, 5, COL_DIV), 1, 4)
        strColTm = Format(strColTm, "##:##")
        HosilId = medGetP(objDIC.GetString, 6, COL_DIV)
        strStatFg = medGetP(objDIC.GetString, 7, COL_DIV)
        
        If HosilId <> "" Then
            strBarW_H = txtWardID.Text & "/" & HosilId
        Else
            strBarW_H = txtWardID.Text
        End If
        
        
        '��ü��ȣ ��� : 2001.2.8 �߰�
        strAccSeq = Mid(strSpcNo, 1, 2) & "-" & Format(Mid(strSpcNo, 3), "########0")
        strAccSeq = Format(strAccSeq, String(11, "@"))
        '���ڵ� ���

        objBar.Label_PrintOut _
                        strBuildNm, "XM", "", strAccSeq, strSpcNo, strPtid, _
                        strPtnm, "", "", strStatFg, strBarW_H, _
                        strColDt, strColTm, "", Val(txtCopy)

        objDIC.MoveNext
    Loop
    
    Set objBar = Nothing

End Sub

'& ä�� Ŭ���� MyCollect �� �̿��Ͽ� �ش� ȯ�ڵ��� ó���� ä�������Ѵ�.
Private Sub DoCollectionForLIS(ByVal Row As Long)
    Dim Rs          As Recordset
    
    Dim tmpDate     As String
    Dim tmpTime     As String
    Dim SqlStmt     As String
    
    Dim tmpData()   As String
    Dim tmpDeptCd   As String
    Dim tmpOrdDoct  As String
    Dim tmpMajDoct  As String
    Dim tmpTestCd   As String
    
    Dim sWorkarea   As String
    Dim sAccdt      As String

    Dim sBuildCd    As String
    Dim blnMornCol  As Boolean
    Dim blnSuccess  As Boolean
    
    Dim lngBldRow   As Long
    
    Dim i           As Integer
    Dim j           As Integer
    Dim iAccseq     As Long

    Call objLISCollect.SetWardCol(sWorkDt, sWorkTm, Trim(txtWardID))
    objLISCollect.MornFg = 0 'ChkMornFg.Value      '��ħä������

    ReDim tmpData(0 To 16)
    
    With tblPtList
        .Row = Row
                    tmpData(0) = Mid(Format(Now, "YYYY"), 4)
        .Col = 4:   tmpData(1) = .Value                                     'ȯ��ID
        .Col = 5:   tmpData(2) = .Value                                     'ȯ�ڸ�
        .Col = 7:  tmpData(3) = .Value                                     'ȯ�ڼ���
        .Col = 20:
                    If IsDate(Format(.Value, CS_DateMask)) Then
                        tmpData(4) = DateDiff("y", Format(.Value, CS_DateMask), Now)    'ȯ���Ϸ�
                    Else
                        tmpData(4) = Mid(.Value, 1, 4) & "-01-01"
                        If IsDate(tmpData(4)) Then
                            tmpData(4) = DateDiff("y", tmpData(4), GetSystemDate)
                        Else
                            tmpData(4) = 0
                        End If
                    End If

        .Col = 13:  tmpData(5) = .Value                                 '�Կ���
                    tmpData(6) = Format(Now, CS_DateDbFormat)           '�Է���
                    tmpData(7) = Format(Now, CS_TimeDbFormat)           '�Է½ð�
                    tmpData(8) = ObjSysInfo.EmpId                       '�Է���
                    tmpData(9) = ""                                     '��������ȣ
                    tmpData(10) = Format(Now, CS_DateDbFormat)          'ä����
                    objLISCollect.ColTm = Format(GetSystemDate, "HHMMSS")     'ä����
                    tmpData(11) = ObjSysInfo.EmpId                      'ä����
        .Col = 16:  tmpData(12) = .Value                                '����ID
        .Col = 21:  tmpData(13) = .Value                                '����ID
        .Col = 22:  tmpData(14) = .Value                                'ȣ��ID
                    tmpData(15) = ""                                    'ħ��ID
                    tmpData(16) = ObjSysInfo.BuildingCd                 '** ä���� ����Ǵ� �ǹ��ڵ�
        
        Call objLISCollect.SetColData(tmpData)
        
'        .Col = 22:
           blnMornCol = False
        
        .Col = 15:  tmpDeptCd = .Value                        '�����
        .Col = 17:  tmpOrdDoct = .Value                       'ó����
        .Col = 14:  tmpMajDoct = .Value                       '��ġ��
        .Col = 23:  tmpTestCd = .Value                        'ȣ��ID
    End With


    tmpDate = Format(dtpToTime.Value, CS_DateDbFormat)
    tmpTime = "235959"
       
    
    ' ó�泻�� �˻�
        SqlStmt = SqlReadWardOrderE(objLISCollect.Ptid, tmpDate, tmpTime, , enBussDiv.BussDiv_InPatient, , LIS_ORDDIV, tmpTestCd)
    
    Set Rs = New Recordset
    Rs.Open SqlStmt, DBConn
    
    If Rs.EOF Then
        blnSuccess = False
        GoTo Err_Trap
    End If

    ReDim tmpData(0 To 20)
    With Rs
        
        For i = 1 To .RecordCount
            tmpData(0) = ObjSysInfo.BuildingCd: sBuildCd = tmpData(0)
            tmpData(1) = Trim("" & .Fields("WorkArea").Value)   'WorkArea
            tmpData(2) = Trim("" & .Fields("SpcCd").Value)      'SpcCd
            tmpData(3) = Trim("" & .Fields("StoreCd").Value)    'StoreCd
            tmpData(4) = Trim("" & .Fields("StatFg").Value)
            tmpData(5) = Format("" & Rs.Fields("ReqDt").Value, CS_DateMask) & " " & _
                         Format("" & Rs.Fields("ReqTm").Value, CS_TimeLongMask)        '���ä���Ͻ�
            tmpData(6) = Trim("" & .Fields("TestDiv").Value)    'TestDiv
            tmpData(7) = Trim("" & .Fields("MultiFg").Value)    'MultiFg
            tmpData(8) = Trim("" & .Fields("SpcGrp").Value)     'SpcGrp
            tmpData(9) = Trim("" & .Fields("OrdDt").Value)      'OrdDt
            tmpData(10) = Trim("" & .Fields("OrdNo").Value)     'OrdNo
            tmpData(11) = Trim("" & .Fields("OrdSeq").Value)    'OrdSeq
            tmpData(12) = Trim("" & .Fields("OrdCd").Value)     'OrdCd
            tmpData(13) = tmpDeptCd
            tmpData(14) = tmpOrdDoct
            tmpData(15) = tmpMajDoct
            tmpData(16) = Trim("" & .Fields("AbbrNm5").Value)   'ó�� ����
            tmpData(17) = Trim("" & .Fields("LabelCnt").Value)  '��������
            
'            Call ObjLISComCode.LisItem.KeyChange(tmpData(12))
            tmpData(18) = GetLabDiv(tmpData(12)) ' ObjLISComCode.LisItem.Fields("labdiv")    'LabDiv
            
            Call GetSpcInfo(tmpData(2), tmpData(19), tmpData(20))
'            Call ObjLISComCode.LisSpc.KeyChange(tmpData(2))
'            tmpData(19) = ObjLISComCode.LisSpc.Fields("spcbarnm")    '��ü����
'            tmpData(20) = ObjLISComCode.LisSpc.Fields("labrange")   '�̻���������ȣ����
            
            Call objLISCollect.SetAddLabCollect(tmpData)
            .MoveNext
        Next
    End With

    ' ä�� ����
    
    If Rs.RecordCount > 0 Then
        blnSuccess = objLISCollect.DoCollection(pbrPtCnt)
        blnCollectFg = True
    Else
        GoTo Skip
    End If

Err_Trap:
    If Not blnSuccess Then
        tblPtList.Row = Row
        tblPtList.Col = -1
        tblPtList.ForeColor = vbRed       '������
        intErrCount = intErrCount + 1
    Else
        Dim strBld As String
        
        strBld = GetBuildNm(ObjSysInfo.BuildingCd)
        
        DoEvents
         '* Delivery Location �� Count
         For i = 1 To objLISCollect.ColCount
            Call objLISCollect.GetLabNumbers(i, sWorkarea, sAccdt, iAccseq, sBuildCd)
           
            lngBldRow = 0
            For j = 1 To tblCount.DataRowCnt
                tblCount.Row = j: tblCount.Col = 3
                If tblCount.Value = sBuildCd Then
                    lngBldRow = j
                    Exit For
                End If
            Next

            If lngBldRow = 0 Then lngBldRow = tblCount.DataRowCnt + 1
            tblCount.Row = lngBldRow
            tblCount.Col = 1: tblCount.Text = strBld 'ObjLISComCode.Building.Fields("buildnm")
            tblCount.Col = 2: tblCount.Text = Val(tblCount.Text) + 1
            tblCount.Col = 3: tblCount.Text = sBuildCd
        Next

    End If
Skip:
    Set Rs = Nothing

End Sub
Private Function GetLabDiv(ByVal vTestCd As String) As String
    Dim Rs As Recordset
    Dim strSQL As String
    
    strSQL = " select a.testcd,a.applydt,b.field2 from " & T_LAB001 & " a, " & T_LAB032 & " b "
    strSQL = strSQL & " where " & DBW("b.cdindex=", LC3_WorkArea)
    strSQL = strSQL & " and a.workarea=b.cdval1"
    strSQL = strSQL & " and " & DBW("a.testcd=", vTestCd)
    
    Set Rs = New Recordset
    Rs.Open strSQL, DBConn
    If Rs.EOF = False Then
    GetLabDiv = Rs.Fields("field2").Value & ""
    End If
    Set Rs = Nothing
End Function

'Private Sub GetSpcInfo(ByVal vSpcCd As String, ByRef vSpcAbbrNm As String, _
'                            ByRef vLabRng As String)
'    Dim Rs As Recordset
'    Dim strSQL As String
'
'    strSQL = " select  a.field3 spcabbr, b.field2 labrange,a.field5 spcbarnm  " & _
'            " from " & T_LAB032 & " b, " & T_LAB032 & " a " & _
'            " where " & dbw("a.cdindex =", LC3_Specimen) & _
'            " and " & dbw("a.cdval1=", vSpcCd) & _
'            " and    " & DBJ("b.cdindex ='C217'") & _
'            " and    " & DBJ("b.cdval1  =* a.field2")
'
'    Set Rs = New Recordset
'    Rs.Open strSQL, dbconn
'
'    vSpcAbbrNm = Rs.Fields("spcbarnm").Value & ""
'    vLabRng = Rs.Fields("labrange").Value & ""
'
'    Set Rs = Nothing
'End Sub


'% �������� ���� �Կ����� ȯ�ڵ��� ó���� �˻��Ѵ�.
Private Sub cmdGetOrders_Click()
    Dim objStatus   As jProgressBar.clsProgress
    Dim objProgress As clsProgress
    Dim Rs          As Recordset
    Dim Resp        As VbMsgBoxResult
    Dim i           As Integer
    
    Dim SqlStmt     As String
    Dim tmpDate     As String
    Dim tmpTime     As String
    Dim tmpStatFg   As String

    If chkCancle.Value = 0 Then
        If Trim(txtWardID.Text) = "" Then
            MsgBox "����ID�� �Է��ϼ���.", vbInformation, "��������"
            txtWardID.SetFocus
            Exit Sub
        End If
    End If
    
'    If Trim(txtTestCd.Text) = "" Then
'        MsgBox "�˻��׸��� �Է��ϼ���.", vbInformation, "�˻��׸���"
'        txtTestCd.SetFocus
'        Exit Sub
'    End If

    '2001-11-07 : ������ �����ϰ�ä�� ���� ���� --------------------------------------------------

'    Set objStatus = New jProgressBar.clsProgress
'    With objStatus
'        .Container = Me
'        .Left = LisLabel1.Left
'        .Top = LisLabel1.Top
'        .Width = LisLabel1.Width
'        .Height = 280
'        .Message = "������ �����ϰ� ä�볻���� �����ϰ� �ֽ��ϴ�..."

''        .Choice = True
''        .Appearance = aPlate
''        .SetMyForm Me
''        .XWidth = LisLabel1.Width
''        .XPos = LisLabel1.Left
''        .YPos = LisLabel1.Top
''        .YHeight = 280
''        .ForeColor = &H864B24
''        .Msg = "������ �����ϰ�ä�� ������ �����ϰ� �ֽ��ϴ�..."
''        .Max = 100
''        .Value = 50
'    End With
'
'    Set objLISCollect = New clsLISCollectioin
'    If Not objLISCollect.Archive_WardColData(txtWardID.Text) Then
'        MsgBox "�����ϰ�ä�� ���� Archive�� ������ �߻��߽��ϴ�." & vbCrLf & _
'                "����� Ȥ�� �ӻ󺴸����� �����ٶ��ϴ�. (��" & ObjSysInfo.HelpLine & ")", vbCritical, "�����߻�"
'    '---------------------------------------------------------------------------------------------
'    End If
'    Set objStatus = Nothing
'    Set objLISCollect = Nothing

'    If ChkMornFg.Value = 1 Then
'        Resp = MsgBox("�ӻ󺴸� ��ħä�� �۾��� �����Ͻðڽ��ϱ�?", vbQuestion + vbYesNo, "��ħä��")
'        If Resp = vbNo Then Exit Sub
'    End If
    
    Call TableClear(1)
    
    
    If chkCol.Value = 1 Then
        tmpDate = Format(dtpColDtTm.Value, CS_DateDbFormat)
        tmpTime = Format(dtpColDtTm.Value, CS_TimeDbFormat)
    Else
        tmpDate = Format(dtpToTime.Value, CS_DateDbFormat)
        tmpTime = "235959"
    End If
    
    MouseRunning
    Set objProgress = New clsProgress
    
    With objProgress
        .Container = MainFrm.stsbar
        .Message = Trim(txtWardID.Text) & " ���� ȯ�ڵ��� ó���� �˻����Դϴ�.."
'        .Caption = "�����ϰ�ä��"
'        .Msg = Trim(txtWardID.Text) & " ���� ȯ�ڵ��� ó���� �˻����Դϴ�.."
'        .Mode = 1
    End With

'    If ChkMornFg.Value = 1 Then
'        SqlStmt = objMySql.SqlOrderForMornCol(tmpDate, tmpTime, txtWardID.Text)
'    Else
'        SqlStmt = objMySql.SqlWardOrder(tmpDate, tmpTime, txtWardID.Text)
        If chkCancle.Value = 1 Then
            SqlStmt = Get168SqlWardOrder_Cancle(tmpDate, tmpTime, txtWardID.Text, txtTestCd.Text, 2)    '-- 2��:����/ä�� 3��:W/S�ۼ�
        Else
            SqlStmt = Get168SqlWardOrder(tmpDate, tmpTime, txtWardID.Text, txtTestCd.Text, chkAction.Value)
        End If
'    End If

'    pbrPtCnt.Visible = True

    Set Rs = New Recordset
    Rs.Open SqlStmt, DBConn
    
    If Rs.EOF Then
'        MsgBox "ó�� �˻��� ������ �߻��߽��ϴ�. " &  "����� Ȥ�� �ӻ󺴸����� �����ٶ��ϴ�.", vbExclamation ', "�����߻�"
        GoTo Err_Trap
    End If

    If Not Rs.EOF Then
        Call DisplayOrders(Rs, objProgress)
    End If
    
    'ó�泻�� Display
    cmdSave.Enabled = True
    blnCleanFg = False

    DoEvents

    tblPtList.SetFocus

Err_Trap:
    Set Rs = Nothing
    Set objProgress = Nothing
    
'    pbrPtCnt.Visible = False

    Call MouseDefault

End Sub

Private Sub DisplayOrders(ByVal objRs As Recordset, Optional ByRef objPrgBar As Object = Nothing)

    Dim objGetSql   As clsBBSCollection

    Dim tmpPtId     As String
    Dim tmpStatFg   As String
    Dim tmpSpcCd    As String
    Dim tmpOrdDiv   As String
    Dim i           As Long
    
    
    Set objGetSql = New clsBBSCollection
    lblTotalCnT.Caption = objRs.RecordCount
    With tblPtList
        
        '���α׷����� ó��..
        If Not objPrgBar Is Nothing Then
'            objPrgBar.Min = 0
            objPrgBar.Max = objRs.RecordCount * 100 + 1
'            objPrgBar.Value = 0
'            objPrgBar.Visible = True
            DoEvents
        End If

        .MaxRows = 0
        .MaxRows = objRs.RecordCount  'IIf(objRs.RecordCount < 29, 29, objRs.RecordCount)
        .Row = 1

        intPtCount = 0

        For i = 1 To objRs.RecordCount

            If Not objPrgBar Is Nothing Then objPrgBar.Value = objPrgBar.Value + 50
            DoEvents

            intPtCount = intPtCount + 1
            .Row = intPtCount
            .Col = 2: .Text = "" & objRs.Fields("WardNM").Value             ' ����
            .Col = 3: .Text = "" & objRs.Fields("hosil").Value              ' ȣ��
            .Col = 4: .Text = "" & objRs.Fields("Ptid").Value               ' ȯ��ID
            .Col = 5: .Text = "" & objRs.Fields("ptnm").Value               ' ȯ���̸�
            .Col = 6: .Text = "" & objRs.Fields("age").Value                ' ����
            .Col = 7: .Text = "" & objRs.Fields("Sex").Value                ' ����
'                If IsNumeric(.Text) Then
'                    .Text = Choose((Val(.Text) Mod 2) + 1, "F", "M")
'                End If
            .Col = 8: .Text = "" & objRs.Fields("testnm").Value                    '�˻��Ī
            .Col = 9: .Text = "" & objRs.Fields("MESG").Value                     ' Message
            .Col = 10: .Text = "" & objRs.Fields("orddoctnm").Value                'ó���Ǹ�
            .Col = 11: .Text = "" & objRs.Fields("OrdDtm").Value                   'ó������
            .Col = 12: .Text = "" & objRs.Fields("ReqDTM").Value                   '�������
            .Col = 13: .Text = "" & objRs.Fields("BedInDT").Value                  '�Կ�����
            .Col = 14: .Text = "" & objRs.Fields("majdoct").Value                  'ȣ��ID
            .Col = 15: .Text = "" & objRs.Fields("deptcd").Value                   '��ID
            .Col = 16: .Text = "" & objRs.Fields("wardid").Value                   '����ID
            .Col = 17: .Text = "" & objRs.Fields("orddoct").Value                  'ó����ID
            .Col = 18: .Text = "" & objRs.Fields("spccd").Value                    '��ü�ڵ�
            .Col = 19: .Text = "" & objRs.Fields("workarea").Value                 'WorkArea
            .Col = 20: .Text = "" & objRs.Fields("dob").Value                      '����
            .Col = 21: .Text = "" & objRs.Fields("roomid").Value                   'RoomID
            .Col = 22: .Text = "" & objRs.Fields("hosilid").Value                  'ȣ��ID
            .Col = 23: .Text = "" & objRs.Fields("testcd").Value                   '�˻��ڵ�
            .Col = 24: .Text = "" & objRs.Fields("accdt").Value                    ' AccDT
            .Col = 25: .Text = "" & objRs.Fields("accseq").Value                   ' AccSeq
            .Col = 26: .Text = "" & objRs.Fields("orddt").Value                    ' orddt
            .Col = 27: .Text = "" & objRs.Fields("ordno").Value                    ' ordno
            .Col = 28: .Text = "" & objRs.Fields("ordseq").Value                   ' ordseq

            tmpStatFg = "" & objRs.Fields("StatFg").Value                           '���޿���
            tmpOrdDiv = "" & objRs.Fields("orddiv").Value                           'ó�汸��
            tmpSpcCd = "" & objRs.Fields("SpcCd").Value                             '��ü
            
            
            If chkTestdiv.Value = 1 Then                                            '�˻��ڵ�� ���
                If tmpSpcCd = "" And tmpOrdDiv = BBS_ORDDIV Then tmpSpcCd = "����"
            Else                                                                    '�˻������ ���
                If tmpOrdDiv = LIS_ORDDIV Then
                    Dim tmpSpcNm As String
                    Dim tmpLabRng As String
                    
                    Call GetSpcInfo(tmpSpcCd, tmpSpcNm, tmpLabRng)
                    
                    If tmpSpcNm <> "" Then
                        tmpSpcCd = tmpSpcNm
                    Else
                        tmpSpcCd = objMySql.Get_SpcNm(tmpSpcCd, tmpOrdDiv)
                    End If
                    
'                    If ObjLISComCode.LisSpc.Exists(tmpSpcCd) Then
'                        ObjLISComCode.LisSpc.KeyChange (tmpSpcCd)
'                        tmpSpcCd = ObjLISComCode.LisSpc.Fields("spcbarnm")
'                    Else
'                        tmpSpcCd = objMySql.Get_SpcNm(tmpSpcCd, tmpOrdDiv)
'                    End If
                Else
                    tmpSpcCd = objMySql.Get_SpcNm(tmpSpcCd, tmpOrdDiv)
                End If
                If tmpSpcCd = "" And tmpOrdDiv = BBS_ORDDIV Then
                    tmpSpcCd = "����"
                End If
            End If
'            If tmpStatFg = "1" Then     '���ް�ü
'                .Col = 5
'                If InStr(1, .Text, tmpSpcCd) = 0 Then
'                    .Text = .Text & tmpSpcCd & ", "
'                End If
'
'                .Col = 22: .Text = "0"
'            Else
'                .Col = 6
'
'                If InStr(1, .Text, tmpSpcCd) = 0 Then
'                    .Text = .Text & tmpSpcCd & ", "
'                End If
''                If ChkMornFg.Value = 1 Then
''                    .Col = 22: .Text = "1"
''                Else
''                    .Col = 22: .Text = "0"
''                End If
'            End If

'            Select Case tmpOrdDiv
'            Case LIS_ORDDIV:   '�ӻ�
'                .Col = 15: .ForeColor = vbRed: .Text = "��"     'ó�汸�С��
'            Case BBS_ORDDIV:   '����
'                .Col = 17: .ForeColor = vbRed: .Text = "��"     'ó�汸�С��
'                If objGetSql.Blood_Existence(tmpPtId, Format(GetSystemDate, "yyyyMMdd"), _
'                                            Format(GetSystemDate, "HHmm")) = True Then
'                    .Col = 18: .ForeColor = vbBlue: .Value = "�ű�"
'                Else
'                    .Col = 18: .ForeColor = DCM_Gray: .Value = "����"
'                End If
'
'            End Select
'            .Col = 19: .Text = Format(GetSystemDate, "YY-MM-DD")
'            .Col = 20: .Text = Format(GetSystemDate, "HH:MM")
            objRs.MoveNext
        Next

        If Not objPrgBar Is Nothing Then objPrgBar.Value = objPrgBar.Max
        DoEvents

        pbrPtCnt.Min = 0
        pbrPtCnt.Max = .DataRowCnt * 10
        pbrPtCnt.Value = 0

        
        dtpColDtTm.Value = GetSystemDate '

    End With

    Set objGetSql = Nothing

End Sub

Private Sub GetSpcInfo(ByVal vSpcCd As String, ByRef vSpcAbbr As String, _
                            ByRef vLabRng As String)
    Dim Rs As Recordset
    Dim strSQL As String
    
    strSQL = " select  a.cdval1 spccd, a.field4 spcnm, a.field3 spcabbr, a.field5 spcbarnm,  " & _
            " a.field1 multifg, a.field2 spcgrp, b.field2 labrange " & _
            " from " & T_LAB032 & " b, " & T_LAB032 & " a " & _
            " where " & DBW("a.cdindex =", LC3_Specimen) & _
            " and " & DBW("a.cdval1=", vSpcCd) & _
            " and    " & DBJ("b.cdindex ='C217'") & _
            " and    " & DBJ("b.cdval1  =* a.field2")

    Set Rs = New Recordset
    Rs.Open strSQL, DBConn
    If Rs.EOF = False Then
    vSpcAbbr = Rs.Fields("spcbarnm").Value & ""
    vLabRng = Rs.Fields("labrange").Value & ""
    End If
    Set Rs = Nothing
End Sub

Private Function DupCheck(ByVal pBldNo As String) As Boolean
'�ߺ����� üũ�Ѵ�.

    Dim strClip As String
    Dim strPtid As String
    
    Dim ii As Integer
    
        
    strPtid = pBldNo
    
    With tblPtList

        .Row = 1: .Row2 = .MaxRows
        .Col = 3: .Col2 = 3
        .BlockMode = True
        strClip = .ClipValue
        .BlockMode = False

        If InStr(strClip, strPtid) Then
            DupCheck = True
        Else
            DupCheck = False
        End If
    End With

End Function

' ���ؽð��� ����Ǹ� Clear
Private Sub dtpToTime_Change()

    If Not blnCleanFg Then Call TableClear(1)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call ICSPatientMark
    Set objMySql = Nothing
    Set objLISCollect = Nothing
    Set objMyList = Nothing

End Sub

Private Sub optApplyColTm_Click(Index As Integer)

    Dim Resp As VbMsgBoxResult

    If dtpColDtTm.Value < Now Then
        If dtpColDtTm.Tag = "1" Then
            dtpColDtTm.Tag = "0"
        Else
            Resp = MsgBox("ä��ð��� ����ð����� �����Դϴ�. �����Ͻðڽ��ϱ�?", _
                   vbQuestion + vbYesNo, "ä��ð�����")
            If Resp = vbYes Then
                dtpColDtTm.Tag = "1"
            Else
                dtpColDtTm.Tag = "0"
                dtpColDtTm.Value = Format(GetSystemDate, "YY-MM-DD HH:MM")
            End If
        End If
    End If

    With tblPtList
        If optApplyColTm(0).Value Then  '��ü
            .Row = 1: .Row2 = .DataRowCnt
            .Col = 19: .Col2 = 19
            .BlockMode = True
            .Text = Format(dtpColDtTm.Value, "YY-MM-DD")
            .BlockMode = False
            .Col = 20: .Col2 = 20
            .BlockMode = True
            .Text = Format(dtpColDtTm.Value, "HH:MM")
            .BlockMode = False
        Else
            If .ActiveRow > .DataRowCnt Then Exit Sub
            .Row = .ActiveRow
            .Col = 19: .Text = Format(dtpColDtTm.Value, "YY-MM-DD")
            .Col = 20: .Text = Format(dtpColDtTm.Value, "HH:MM")
            optApplyColTm(1).Value = False
        End If
    End With

End Sub

Private Sub optOption_Click(Index As Integer)

    Select Case Index
    Case 0, 2: txtCopy.Text = 1
                txtCopy.Enabled = True
    Case 1: txtCopy.Text = 0
                txtCopy.Enabled = False
    End Select

End Sub

Private Sub cmdWardList_Click()
'% �����ڵ� ����Ʈ�� �˾��Ѵ�.
'    Dim objWard As clsBasisData
    

    Set objMyList = New clsPopUpList
'    Set objWard = New clsBasisData
    
    With objMyList
        .Connection = DBConn
        .FormCaption = "���� ��ȸ"
        .ColumnHeaderText = "�����ڵ�;������"
        Call .LoadPopUp(GetSQLWardList) ', 2700, Frame2.Left + cmdWardList.Left) ', ObjLISComCode.WardId)
        If .SelectedString <> "" Then
            txtWardID.Text = medGetP(.SelectedString, 1, ";")
            lblWardNm.Caption = medGetP(.SelectedString, 2, ";")
        End If
        ' �������ý� ����Ʈ ȣ�� �߰� :2012-08-21 �½�ȣ
        Call cmdGetOrders_Click
    End With
    
'    Set objWard = Nothing
    Set objMyList = Nothing

End Sub


Private Sub tblPtList_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
    Dim Rs          As Recordset
    Dim tmpToolTip  As String
    
    Dim strSQL      As String
    Dim strPtid     As String
    Dim strOrdDate  As String
    Dim strOrdDiv   As String
    Dim strWardId   As String
    Dim strBBSORDCd As String
    Dim strLISORDCd As String

    If Row = 0 Then Exit Sub

    tmpToolTip = vbCrLf

    With tblPtList
        .Row = Row

        .Col = 2: If Trim(.Value) = "" Then Exit Sub

        .Col = 4: tmpToolTip = tmpToolTip & "  " & .Value & vbCrLf & vbCrLf    'ȯ�ڸ�
        .Col = 5: tmpToolTip = tmpToolTip & "  ���ް�ü : " & .Value & vbCrLf  '���ް�ü
        .Col = 6: tmpToolTip = tmpToolTip & "  �Ϲݰ�ü : " & .Value & vbCrLf  '�Ϲݰ�ü
        
        '-- ToolTip �߰����� : �˻��׸� Display
        ' - ȯ��ID
        .Col = 3: strPtid = Trim(.Value)
        strOrdDate = Format(dtpToTime.Value, CS_DateDbFormat)
        strWardId = Trim(txtWardID.Text)
        
        strSQL = objMySql.WardMn_ORDCD(strPtid, strOrdDate, strWardId)
        
        Set Rs = New Recordset
        Rs.Open strSQL, DBConn
        
        If Rs.BOF = False Then
            Do Until Rs.EOF = True
                strOrdDiv = Trim(Rs.Fields("orddiv").Value & "")
                
                '��굿������ �غκ����� ���� �ҷ����ϱ�....��������.

               Select Case strOrdDiv
                   Case "B"
                       strBBSORDCd = strBBSORDCd & Rs.Fields("abbrnm5").Value & "" & "," '�������� �˻��׸�
                       
                   Case "L"
                       strLISORDCd = strLISORDCd & Rs.Fields("abbrnm5").Value & "" & "," '�ӻ󺴸� �˻��׸�
               End Select
        
                Rs.MoveNext
            Loop
        End If
        
        If strBBSORDCd <> "" Then
            tmpToolTip = tmpToolTip & "  �������� : " & strBBSORDCd & vbCrLf  '�������� �˻��׸�
        ElseIf strLISORDCd <> "" Then
            tmpToolTip = tmpToolTip & "  �ӻ󺴸� : " & strLISORDCd & vbCrLf  '�ӻ󺴸� �˻��׸�
        End If
        
        MultiLine = 1
        TipText = tmpToolTip
        TipWidth = 5000
        .TextTipDelay = 1000
        Call .SetTextTipAppearance("����ü", 9, False, False, &HEEFDF2, &H996666)
        ShowTip = True
    End With
    
    Set Rs = Nothing
End Sub

'% ��� ������ ����Ǹ� Clear
Private Sub txtWardId_Change()
    If Not blnCleanFg Then Call TableClear(1)
End Sub

Private Sub ClearRtn(ByVal intOpt As Integer)
    'Unlocking...
    chkAction.Value = 0
    txtWardID.Enabled = True
    txtWardID.BackColor = &H80000005
    cmdWardList.Enabled = True
    dtpToTime.Enabled = True
    cmdGetOrders.Enabled = True
    cmdSave.Enabled = False
    pbrPtCnt.Visible = False

    sWorkDt = "": sWorkTm = ""
'    txtWardID.Text = ""
'    txtTestCd.Text = ""
    lblWardNm.Caption = ""
    dtpToTime.Value = Format(GetSystemDate, "YYYY-MM-DD hh:mm:ss")
    chkCol.Value = 0
    dtpColDtTm.Value = GetSystemDate
    dtpColDtTm.Enabled = False
    dtpColDtTm.Tag = "0"
    pbrPtCnt.Value = 0
    chkPrintFg = 0
    optOption(0).Value = True
    optApplyColTm(0).Value = True
    intErrCount = 0
    Call TableClear(intOpt)
End Sub


'% Table���� Clear�Ѵ�
Private Sub TableClear(ByVal intOpt As Integer)
    tblPtList.MaxRows = 0
    tblPtList.MaxRows = 50
    If intOpt = 1 Then
        lblColNm.Caption = ""
        lblPtCount.Caption = ""
        tblCount.MaxRows = 0
        tblCount.MaxRows = 50
        blnCleanFg = True
    End If
End Sub

'% ���� ID
Private Sub txtWardId_GotFocus()

    With txtWardID
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub txtWardId_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If objMyList Is Nothing Then Call cmdWardList_Click
    End If
End Sub


Private Sub txtWardId_KeyPress(KeyAscii As Integer)

    On Error GoTo Err_Trap

    KeyAscii = Asc(UCase(Chr(KeyAscii)))

    If KeyAscii = vbKeyReturn Then
        If txtWardID.Text = "" Then
            lblWardNm.Caption = ""
            Exit Sub
        Else
'            Dim objWard As clsBasisData
            Dim Rs As Recordset
            Dim strWard As String
            
'            Set objWard = New clsBasisData
            Set Rs = New Recordset
            
            strWard = GetSQLWard(txtWardID.Text)
            
            Rs.Open strWard, DBConn
            
            If Rs.EOF = False Then
                ObjSysInfo.BuildingCd = Rs.Fields("bldgb").Value & ""
                ObjSysInfo.BuildingNm = Rs.Fields("bldnm").Value & ""
                ObjSysInfo.BuildingNo = Rs.Fields("bldno").Value & ""
                txtWardID.Tag = txtWardID.Text
            Else
                MsgBox "���� �ڵ带 Ȯ���ϼ���.", vbInformation
                txtWardID.Text = ""
                lblWardNm.Caption = ""
                txtWardID.SetFocus
                Call txtWardId_KeyDown(vbKeyDown, 0)
            End If
            Set Rs = Nothing
'            Set objWard = Nothing

'            With ObjLISComCode.WardId
'                If .Exists(txtWardID.Text) Then
'                    Call .KeyChange(txtWardID.Text)
'                    lblWardNm.Caption = .Fields("WardNm")
'                    objsysinfo.BuildingCd = .Tags("bldgb")
'                    objsysinfo.BuildingNm = .Tags("bldnm")
'                    objsysinfo.BuildingNo = .Tags("bldno")
'                    dtpToTime.SetFocus
'                Else
'                    MsgBox "���� �ڵ带 Ȯ���ϼ���..", vbInformation, "�ڵ��Է¿���"
'                    txtWardID.Text = ""
'                    lblWardNm.Caption = ""
'                    txtWardID.SetFocus
'                    Call txtWardId_KeyDown(vbKeyDown, 0)
'                    Exit Sub
'                End If
'            End With
        End If
    End If
    Exit Sub

Err_Trap:
    Resume Next

End Sub

Private Sub PrintColList(ByVal pWardId As String, ByVal pWardNm As String, _
                         ByVal pWorkDt As String, ByVal pWorkTm As String, _
                         ByVal pBuildCd As String, ByVal pBuildNm As String)

    Dim MyReport    As clsWardColList
    Dim strTitleNm  As String
    
    strTitleNm = "���� ä�� ����Ʈ"
    
    Set MyReport = New clsWardColList
    
    With MyReport
        .WardId = pWardId
        .WardNm = pWardNm
        .WorkDt = pWorkDt
        .WorkTm = pWorkTm
        .BuildCd = pBuildCd
        .BuildNm = pBuildNm
        .TestDiv = chkTestdiv.Value
        .TitleNm = strTitleNm
        .SetCrpt CReport
        Call .Print_ColList
    End With

    Set MyReport = Nothing

End Sub


Public Sub Call_WardId_KeyPress()

    Call txtWardId_KeyPress(vbKeyReturn)

End Sub

Public Sub Call_cmdGetOrders_click()

     Call cmdGetOrders_Click

End Sub


Private Sub txtWardId_LostFocus()

On Error GoTo Err_Trap

    If ActiveControl.Name = cmdWardList.Name Then Exit Sub
    If ActiveControl.Name = cmdClear.Name Then Exit Sub
    If ActiveControl.Name = cmdExit.Name Then Exit Sub
    If txtWardID.Text = "" Then
        lblWardNm.Caption = ""
        Exit Sub
    Else
        Call txtWardId_KeyPress(vbKeyReturn)
    End If
    Exit Sub
Err_Trap:
    Resume Next

End Sub

Public Function Get168SqlWardOrder(ByVal ReqDt As String, ByVal ReqTm As String, ByVal WardId As String, ByVal TestCd As String, ByVal pAction As Integer) As String

Dim strTemp As String
Dim strAction As Integer

    strTemp = ""
    If TestCd <> "" Then strTemp = " and c.ordcd = '" & TestCd & "' "
'    strAction = enStsCd.StsCd_LIS_Order
'    If pAction = 1 Then strAction = StsCd_LIS_Collection
' orddtm���� ����
' to_char(a.ordtime, 'hh24miss') as ordtm,  to_char(a.orddate, 'yyyy-mm-dd hh:MM:ss') as orddtm,
' => to_char(a.ordtime, 'yyyy-mm-dd hh:MM:ss') as orddtm,  to_char(a.orddate, 'yyyy-mm-dd hh:MM:ss') as ordtm,
'    Get168SqlWardOrder = " select g.deptnm || ' / ' || h.deptnm as Wardnm, i.roomno1 as hosil,  a.patno as PtId, a.patname as PtNm, TRUNC(MONTHS_BETWEEN(SYSDATE, a.birtdate)/12)  as AGE, " & _
'                         " decode ( " & F_SEX2("a") & ", 1, 'M', 2, 'F', 3,'M', 4,'F', 7,'M', 8,'F')  as Sex, c.MESG, d.testnm, f.username as orddoctnm, b.orddtm as orddtm, " & _
'                         " to_char(to_date(b.reqdt,'yyyymmdd'), 'YYYY-MM-DD')|| ' '|| to_char(to_date(b.reqtm, 'hh24miss'), 'hh:mm:ss') as reqdtm, " & _
'                         " to_char(to_date(b.bedindt,'yyyymmdd'), 'YYYYMMDD') as bedindt, b.majdoct, b.deptcd , b.wardid, b.orddoct, c.spccd, c.statfg, b.orddiv, to_char(a.birtdate,'yyyyMMdd')  as DoB,  b.roomid , b.hosilid, d.testcd, b.donefg, c.stscd, c.workarea, c.accdt, c.accseq " & _
'                    " FROM (" & _
'           " select a.patno as ptid,  to_char(a.orddate, 'yyyymmdd') as orddt,  a.ordseqno as ordno,  to_char(a.ordtime, 'hh24miss') as ordtm,  to_char(a.orddate, 'yyyy-mm-dd hh:MM:ss') as orddtm, " & _
'                " decode(a.patsect, 'I', '2', '1') as bussdiv ,  to_char(a.meddate, 'yyyymmdd') as BEDINDT ," & " to_char(a.hopedate, 'yyyymmdd') as REQDT, " & _
'                " DECODE(to_char(a.hopedate, 'hh24miss') , '000000', to_char(a.ordtime, 'hh24miss'), to_char(a.hopedate, 'hh24miss')) as REQTM, " & _
'                " a.meddept as DEPTCD,  a.orddr   as ORDDOCT, a.chadr as MAJDOCT,  a.editid as ENTID,  to_char(a.editdate, 'yyyymmdd') as ENTDT, " & _
'                " to_char(a.editdate, 'hh24miss') as ENTTM,  'L' as ORDDIV,  a.repeatfg as REPEATFG,  a.orgaccno as ORGACCNO,  a.sporddiv, " & _
'                " to_char(a.donefg) as donefg,  null as RECEPTNO, a.wardno as WARDID,  a.roomno as ROOMID,  a.roomno as HOSILID,  to_char(0) as ORDFG,  null as BEDID,  null as HOSCD,  null as TRUSTDT " & _
'             " from MDEXMORT a " & _
'             " where a.ordgrp = 'C1'  and substr(a.slipcd, 1, 1) in ('L', 'N', 'P') " & _
'             " and " & DBW("a.wardno", WardId, 2) & " AND  a.hopedate = to_date('" & ReqDt & "', 'yyyyMMdd') "

'decode ( " & F_SEX2("a") ���� 2017.05.23 �½�ȣ
    Get168SqlWardOrder = " select g.deptnm || ' / ' || h.deptnm as Wardnm, i.roomno1 as hosil,  a.patno as PtId, a.patname as PtNm, TRUNC(MONTHS_BETWEEN(SYSDATE, a.birtdate)/12)  as AGE, " & _
                         " decode ( " & F_SEX2 & ", 1, 'M', 2, 'F', 3,'M', 4,'F', 7,'M', 8,'F')  as Sex, c.MESG, d.testnm, f.username as orddoctnm, b.orddtm as orddtm, " & _
                         " to_char(to_date(b.reqdt,'yyyymmdd'), 'YYYY-MM-DD')|| ' '|| to_char(to_date(b.reqtm, 'hh24miss'), 'hh:mm:ss') as reqdtm, " & _
                         " to_char(to_date(b.bedindt,'yyyymmdd'), 'YYYYMMDD') as bedindt, b.majdoct, b.deptcd , b.wardid, b.orddoct, c.spccd, c.statfg, b.orddiv, to_char(a.birtdate,'yyyyMMdd')  as DoB,  b.roomid , b.hosilid, d.testcd, b.donefg, c.stscd, c.workarea, c.accdt, c.accseq,c.orddt,c.ordno,c.ordseq " & _
                    " FROM (" & _
           " select a.patno as ptid,  to_char(a.orddate, 'yyyymmdd') as orddt,  a.ordseqno as ordno,  to_char(a.ordtime, 'yyyy-mm-dd hh:MM:ss') as orddtm,  to_char(a.orddate, 'yyyy-mm-dd hh:MM:ss') as ordtm, " & _
                " decode(a.patsect, 'I', '2', '1') as bussdiv ,  to_char(a.meddate, 'yyyymmdd') as BEDINDT ," & " to_char(a.hopedate, 'yyyymmdd') as REQDT, " & _
                " DECODE(to_char(a.hopedate, 'hh24miss') , '000000', to_char(a.ordtime, 'hh24miss'), to_char(a.hopedate, 'hh24miss')) as REQTM, " & _
                " a.meddept as DEPTCD,  a.orddr   as ORDDOCT, a.chadr as MAJDOCT,  a.editid as ENTID,  to_char(a.editdate, 'yyyymmdd') as ENTDT, " & _
                " to_char(a.editdate, 'hh24miss') as ENTTM,  'L' as ORDDIV,  a.repeatfg as REPEATFG,  a.orgaccno as ORGACCNO,  a.sporddiv, " & _
                " to_char(a.donefg) as donefg,  null as RECEPTNO, a.wardno as WARDID,  a.roomno as ROOMID,  a.roomno as HOSILID,  to_char(0) as ORDFG,  null as BEDID,  null as HOSCD,  null as TRUSTDT " & _
             " from MDEXMORT a " & _
             " where a.ordgrp = 'C1'  and substr(a.slipcd, 1, 1) in ('L', 'N', 'P') " & _
             " and " & DBW("a.wardno", WardId, 2) & " AND  a.hopedate = to_date('" & ReqDt & "', 'yyyyMMdd') "

    Get168SqlWardOrder = Get168SqlWardOrder & _
                "  ) b, " & _
                "( select a.patno as PTID, to_char(a.orddate,'yyyymmdd') as ORDDT, a.ordseqno as ORDNO, 1 as ORDSEQ, a.ordcd as ORDCD, a.spccode1 as SPCCD, a.storecd as STORECD, decode(a.discyn,'X','1','Y','1',null) as DCFG, " & _
                " null as DCDT, null as DCNO, null as DCID, a.attrcd as ATTRCD, to_char(a.rsltdate,'yyyymmdd') as EXAMDT, to_char(a.rsltdate,'hh24miss') as EXAMTM, a.cofmdr as EXAMDOCT, " & _
                " to_char(a.acptdate,'yyyymmdd')as RCVDT, to_char(a.acptdate,'hh24miss') as RCVTM, to_char(a.stscd) as STSCD, decode(a.eryn,'Y','1',null) as STATFG, to_char(a.donefg) as DONEFG, a.instype as INSDIV, " & _
                " a.workarea as WORKAREA, a.accdt as ACCDT, to_number(a.accseq) as ACCSEQ, null as RECEPTNO, null as OCSORDNO, null as OCSORDSEQ, null as OCSORDPUMOK, a.remark as MESG, a.addfg as ADDFG, " & _
                " decode(a.rcpstat,'Y',to_char(a.rcpdate,'yyyymmdd'),null) as PAYDT, null as UNITQTY, a.volume as VOLUME, null as IRRADFG, null as FILTERFG, a.newtestdiv as NEWTESTDIV, a.workfg as WORKFG, null as WRKDIV, a.ospcno as OSPCNO, a.slipcd as SLIPCD, null as INOUT, null as INOUTSEQ " & _
                " from MDEXMORT a where a.ordgrp = 'C1' and substr(a.slipcd,1,1) in ('L','N','P') AND  a.hopedate = to_date('" & ReqDt & "', 'yyyyMMdd') " & _
                " ) c, " & T_HIS001 & " a, " & T_LAB001 & " d, " & T_HIS005 & " f, " & T_HIS003 & " g, " & T_HIS003 & " h, ORAA1.APIPDLST i " & _
                " WHERE b.reqdt='" & ReqDt & "'" & _
                " AND " & DBW("b.orddiv<>", "Z") & " AND    " & DBW("b.bussdiv", enBussDiv.BussDiv_InPatient, 2) & _
                " AND a." & F_PTID & " = b.ptid AND c.ptid = b.ptid AND c.orddt = b.orddt AND c.ordno = b.ordno AND " & DBW("c.stscd =", pAction) & _
                " AND ( c.dcfg = '' or c.dcfg is null ) " & strTemp & " and c.ordcd = d.testcd and f.userid = b.orddoct and g.dpcd = b.deptcd and h.dpcd = i.wardno and d.workarea = '14'" & _
                " AND a.patno = i.patno " & _
                " AND i.stayyn = 'Y' " & _
                " Order  By WardId,hosil, PtId, c.statfg, c.spccd "

                
'                " Order  By WardId,hosilid, PtId, c.statfg, c.spccd "

End Function

Public Function Get168SqlWardOrder_Cancle(ByVal ReqDt As String, ByVal ReqTm As String, ByVal WardId As String, ByVal TestCd As String, ByVal pAction As Integer) As String

Dim strTemp As String
Dim strAction As Integer
Dim strSQL  As String

    strTemp = ""
    If TestCd <> "" Then strTemp = " and c.ordcd = '" & TestCd & "' "
'    strAction = enStsCd.StsCd_LIS_Order
'    If pAction = 1 Then strAction = StsCd_LIS_Collection
' orddtm���� ����
' to_char(a.ordtime, 'hh24miss') as ordtm,  to_char(a.orddate, 'yyyy-mm-dd hh:MM:ss') as orddtm,
' => to_char(a.ordtime, 'yyyy-mm-dd hh:MM:ss') as orddtm,  to_char(a.orddate, 'yyyy-mm-dd hh:MM:ss') as ordtm,
'    Get168SqlWardOrder = " select g.deptnm || ' / ' || h.deptnm as Wardnm, i.roomno1 as hosil,  a.patno as PtId, a.patname as PtNm, TRUNC(MONTHS_BETWEEN(SYSDATE, a.birtdate)/12)  as AGE, " & _
'                         " decode ( " & F_SEX2("a") & ", 1, 'M', 2, 'F', 3,'M', 4,'F', 7,'M', 8,'F')  as Sex, c.MESG, d.testnm, f.username as orddoctnm, b.orddtm as orddtm, " & _
'                         " to_char(to_date(b.reqdt,'yyyymmdd'), 'YYYY-MM-DD')|| ' '|| to_char(to_date(b.reqtm, 'hh24miss'), 'hh:mm:ss') as reqdtm, " & _
'                         " to_char(to_date(b.bedindt,'yyyymmdd'), 'YYYYMMDD') as bedindt, b.majdoct, b.deptcd , b.wardid, b.orddoct, c.spccd, c.statfg, b.orddiv, to_char(a.birtdate,'yyyyMMdd')  as DoB,  b.roomid , b.hosilid, d.testcd, b.donefg, c.stscd, c.workarea, c.accdt, c.accseq " & _
'                    " FROM (" & _
'           " select a.patno as ptid,  to_char(a.orddate, 'yyyymmdd') as orddt,  a.ordseqno as ordno,  to_char(a.ordtime, 'hh24miss') as ordtm,  to_char(a.orddate, 'yyyy-mm-dd hh:MM:ss') as orddtm, " & _
'                " decode(a.patsect, 'I', '2', '1') as bussdiv ,  to_char(a.meddate, 'yyyymmdd') as BEDINDT ," & " to_char(a.hopedate, 'yyyymmdd') as REQDT, " & _
'                " DECODE(to_char(a.hopedate, 'hh24miss') , '000000', to_char(a.ordtime, 'hh24miss'), to_char(a.hopedate, 'hh24miss')) as REQTM, " & _
'                " a.meddept as DEPTCD,  a.orddr   as ORDDOCT, a.chadr as MAJDOCT,  a.editid as ENTID,  to_char(a.editdate, 'yyyymmdd') as ENTDT, " & _
'                " to_char(a.editdate, 'hh24miss') as ENTTM,  'L' as ORDDIV,  a.repeatfg as REPEATFG,  a.orgaccno as ORGACCNO,  a.sporddiv, " & _
'                " to_char(a.donefg) as donefg,  null as RECEPTNO, a.wardno as WARDID,  a.roomno as ROOMID,  a.roomno as HOSILID,  to_char(0) as ORDFG,  null as BEDID,  null as HOSCD,  null as TRUSTDT " & _
'             " from MDEXMORT a " & _
'             " where a.ordgrp = 'C1'  and substr(a.slipcd, 1, 1) in ('L', 'N', 'P') " & _
'             " and " & DBW("a.wardno", WardId, 2) & " AND  a.hopedate = to_date('" & ReqDt & "', 'yyyyMMdd') "

'    strSQL = " (select g.deptnm || ' / ' || h.deptnm as Wardnm, i.roomno1 as hosil,  a.patno as PtId, a.patname as PtNm, TRUNC(MONTHS_BETWEEN(SYSDATE, a.birtdate)/12)  as AGE, " & _
'                         " decode ( " & F_SEX2("a") & ", 1, 'M', 2, 'F', 3,'M', 4,'F', 7,'M', 8,'F')  as Sex, c.MESG, d.testnm, f.username as orddoctnm, b.orddtm as orddtm, " & _
'                         " to_char(to_date(b.reqdt,'yyyymmdd'), 'YYYY-MM-DD')|| ' '|| to_char(to_date(b.reqtm, 'hh24miss'), 'hh:mm:ss') as reqdtm, " & _
'                         " to_char(to_date(b.bedindt,'yyyymmdd'), 'YYYYMMDD') as bedindt, b.majdoct, b.deptcd , b.wardid, b.orddoct, c.spccd, c.statfg, b.orddiv, to_char(a.birtdate,'yyyyMMdd')  as DoB,  b.roomid , b.hosilid, d.testcd, b.donefg, c.stscd, c.workarea, c.accdt, c.accseq,c.orddt,c.ordno,c.ordseq " & _
'                    " FROM (" & _
'           " select a.patno as ptid,  to_char(a.orddate, 'yyyymmdd') as orddt,  a.ordseqno as ordno,  to_char(a.ordtime, 'yyyy-mm-dd hh:MM:ss') as orddtm,  to_char(a.orddate, 'yyyy-mm-dd hh:MM:ss') as ordtm, " & _
'                " decode(a.patsect, 'I', '2', '1') as bussdiv ,  to_char(a.meddate, 'yyyymmdd') as BEDINDT ," & " to_char(a.hopedate, 'yyyymmdd') as REQDT, " & _
'                " DECODE(to_char(a.hopedate, 'hh24miss') , '000000', to_char(a.ordtime, 'hh24miss'), to_char(a.hopedate, 'hh24miss')) as REQTM, " & _
'                " a.meddept as DEPTCD,  a.orddr   as ORDDOCT, a.chadr as MAJDOCT,  a.editid as ENTID,  to_char(a.editdate, 'yyyymmdd') as ENTDT, " & _
'                " to_char(a.editdate, 'hh24miss') as ENTTM,  'L' as ORDDIV,  a.repeatfg as REPEATFG,  a.orgaccno as ORGACCNO,  a.sporddiv, " & _
'                " to_char(a.donefg) as donefg,  null as RECEPTNO, a.wardno as WARDID,  a.roomno as ROOMID,  a.roomno as HOSILID,  to_char(0) as ORDFG,  null as BEDID,  null as HOSCD,  null as TRUSTDT " & _
'             " from MDEXMORT a " & _
'             " where a.ordgrp = 'C1'  and substr(a.slipcd, 1, 1) in ('L', 'N', 'P') " & _
'             " AND  a.hopedate = to_date('" & ReqDt & "', 'yyyyMMdd') "
'             '" and " & DBW("a.wardno", WardId, 2) & " AND  a.hopedate = to_date('" & ReqDt & "', 'yyyyMMdd') "
'
'    strSQL = strSQL & _
'                "  ) b, " & _
'                "( select a.patno as PTID, to_char(a.orddate,'yyyymmdd') as ORDDT, a.ordseqno as ORDNO, 1 as ORDSEQ, a.ordcd as ORDCD, a.spccode1 as SPCCD, a.storecd as STORECD, decode(a.discyn,'X','1','Y','1',null) as DCFG, " & _
'                " null as DCDT, null as DCNO, null as DCID, a.attrcd as ATTRCD, to_char(a.rsltdate,'yyyymmdd') as EXAMDT, to_char(a.rsltdate,'hh24miss') as EXAMTM, a.cofmdr as EXAMDOCT, " & _
'                " to_char(a.acptdate,'yyyymmdd')as RCVDT, to_char(a.acptdate,'hh24miss') as RCVTM, to_char(a.stscd) as STSCD, decode(a.eryn,'Y','1',null) as STATFG, to_char(a.donefg) as DONEFG, a.instype as INSDIV, " & _
'                " a.workarea as WORKAREA, a.accdt as ACCDT, to_number(a.accseq) as ACCSEQ, null as RECEPTNO, null as OCSORDNO, null as OCSORDSEQ, null as OCSORDPUMOK, a.remark as MESG, a.addfg as ADDFG, " & _
'                " decode(a.rcpstat,'Y',to_char(a.rcpdate,'yyyymmdd'),null) as PAYDT, null as UNITQTY, a.volume as VOLUME, null as IRRADFG, null as FILTERFG, a.newtestdiv as NEWTESTDIV, a.workfg as WORKFG, null as WRKDIV, a.ospcno as OSPCNO, a.slipcd as SLIPCD, null as INOUT, null as INOUTSEQ " & _
'                " from MDEXMORT a where a.ordgrp = 'C1' and substr(a.slipcd,1,1) in ('L','N','P') AND  a.hopedate = to_date('" & ReqDt & "', 'yyyyMMdd') " & _
'                " ) c, " & T_HIS001 & " a, " & T_LAB001 & " d, " & T_HIS005 & " f, " & T_HIS003 & " g, " & T_HIS003 & " h, ORAA1.APIPDLST i " & _
'                " WHERE b.reqdt='" & ReqDt & "'" & _
'                " AND " & DBW("b.orddiv<>", "Z") & " AND    " & DBW("b.bussdiv", enBussDiv.BussDiv_InPatient, 2) & _
'                " AND a." & F_PTID & " = b.ptid AND c.ptid = b.ptid AND c.orddt = b.orddt AND c.ordno = b.ordno AND " & DBW("c.stscd =", pAction) & _
'                " AND ( c.dcfg = '' or c.dcfg is null ) " & strTemp & " and c.ordcd = d.testcd and f.userid = b.orddoct and g.dpcd = b.deptcd and h.dpcd = i.wardno and d.workarea = '14'" & _
'                " AND a.patno = i.patno " & _
'                " AND i.stayyn = 'Y' " & _
'                " Order  By WardId,hosil, PtId, c.statfg, c.spccd) a "
'
''                " Order  By WardId,hosilid, PtId, c.statfg, c.spccd "
'
'    Get168SqlWardOrder_Cancle = " select a.* from " & strSQL & ", s2lab301 b "
'    Get168SqlWardOrder_Cancle = Get168SqlWardOrder_Cancle & " where a.workarea = b.workarea and a.accdt = b.accdt and a.accseq=b.accseq "

    Get168SqlWardOrder_Cancle = " select g.deptnm || ' / ' || h.deptnm as Wardnm, i.roomno1 as hosil,  a.patno as PtId, a.patname as PtNm, TRUNC(MONTHS_BETWEEN(SYSDATE, a.birtdate)/12)  as AGE, " & _
                         " decode ( " & F_SEX2("a") & ", 1, 'M', 2, 'F', 3,'M', 4,'F', 7,'M', 8,'F')  as Sex, c.MESG, d.testnm, f.username as orddoctnm, b.orddtm as orddtm, " & _
                         " to_char(to_date(b.reqdt,'yyyymmdd'), 'YYYY-MM-DD')|| ' '|| to_char(to_date(b.reqtm, 'hh24miss'), 'hh:mm:ss') as reqdtm, " & _
                         " to_char(to_date(b.bedindt,'yyyymmdd'), 'YYYYMMDD') as bedindt, b.majdoct, b.deptcd , b.wardid, b.orddoct, c.spccd, c.statfg, b.orddiv, to_char(a.birtdate,'yyyyMMdd')  as DoB,  b.roomid , b.hosilid, d.testcd, b.donefg, c.stscd, c.workarea, c.accdt, c.accseq,c.orddt,c.ordno,c.ordseq " & _
                    " FROM (" & _
           " select a.patno as ptid,  to_char(a.orddate, 'yyyymmdd') as orddt,  a.ordseqno as ordno,  to_char(a.ordtime, 'yyyy-mm-dd hh:MM:ss') as orddtm,  to_char(a.orddate, 'yyyy-mm-dd hh:MM:ss') as ordtm, " & _
                " decode(a.patsect, 'I', '2', '1') as bussdiv ,  to_char(a.meddate, 'yyyymmdd') as BEDINDT ," & " to_char(a.hopedate, 'yyyymmdd') as REQDT, " & _
                " DECODE(to_char(a.hopedate, 'hh24miss') , '000000', to_char(a.ordtime, 'hh24miss'), to_char(a.hopedate, 'hh24miss')) as REQTM, " & _
                " a.meddept as DEPTCD,  a.orddr   as ORDDOCT, a.chadr as MAJDOCT,  a.editid as ENTID,  to_char(a.editdate, 'yyyymmdd') as ENTDT, " & _
                " to_char(a.editdate, 'hh24miss') as ENTTM,  'L' as ORDDIV,  a.repeatfg as REPEATFG,  a.orgaccno as ORGACCNO,  a.sporddiv, " & _
                " to_char(a.donefg) as donefg,  null as RECEPTNO, a.wardno as WARDID,  a.roomno as ROOMID,  a.roomno as HOSILID,  to_char(0) as ORDFG,  null as BEDID,  null as HOSCD,  null as TRUSTDT " & _
             " from MDEXMORT a " & _
             " where a.ordgrp = 'C1'  and substr(a.slipcd, 1, 1) in ('L', 'N', 'P') " & _
             " AND  a.hopedate = to_date('" & ReqDt & "', 'yyyyMMdd') "
             '" and " & DBW("a.wardno", WardId, 2) & " AND  a.hopedate = to_date('" & ReqDt & "', 'yyyyMMdd') "

    Get168SqlWardOrder_Cancle = Get168SqlWardOrder_Cancle & _
                "  ) b, " & _
                "( select a.patno as PTID, to_char(a.orddate,'yyyymmdd') as ORDDT, a.ordseqno as ORDNO, 1 as ORDSEQ, a.ordcd as ORDCD, a.spccode1 as SPCCD, a.storecd as STORECD, decode(a.discyn,'X','1','Y','1',null) as DCFG, " & _
                " null as DCDT, null as DCNO, null as DCID, a.attrcd as ATTRCD, to_char(a.rsltdate,'yyyymmdd') as EXAMDT, to_char(a.rsltdate,'hh24miss') as EXAMTM, a.cofmdr as EXAMDOCT, " & _
                " to_char(a.acptdate,'yyyymmdd')as RCVDT, to_char(a.acptdate,'hh24miss') as RCVTM, to_char(a.stscd) as STSCD, decode(a.eryn,'Y','1',null) as STATFG, to_char(a.donefg) as DONEFG, a.instype as INSDIV, " & _
                " a.workarea as WORKAREA, a.accdt as ACCDT, to_number(a.accseq) as ACCSEQ, null as RECEPTNO, null as OCSORDNO, null as OCSORDSEQ, null as OCSORDPUMOK, a.remark as MESG, a.addfg as ADDFG, " & _
                " decode(a.rcpstat,'Y',to_char(a.rcpdate,'yyyymmdd'),null) as PAYDT, null as UNITQTY, a.volume as VOLUME, null as IRRADFG, null as FILTERFG, a.newtestdiv as NEWTESTDIV, a.workfg as WORKFG, null as WRKDIV, a.ospcno as OSPCNO, a.slipcd as SLIPCD, null as INOUT, null as INOUTSEQ " & _
                " from MDEXMORT a where a.ordgrp = 'C1' and substr(a.slipcd,1,1) in ('L','N','P') AND  a.hopedate = to_date('" & ReqDt & "', 'yyyyMMdd') " & _
                " ) c, " & T_HIS001 & " a, " & T_LAB001 & " d, " & T_HIS005 & " f, " & T_HIS003 & " g, " & T_HIS003 & " h, ORAA1.APIPDLST i " & _
                " WHERE b.reqdt='" & ReqDt & "'" & _
                " AND " & DBW("b.orddiv<>", "Z") & " AND    " & DBW("b.bussdiv", enBussDiv.BussDiv_InPatient, 2) & _
                " AND a." & F_PTID & " = b.ptid AND c.ptid = b.ptid AND c.orddt = b.orddt AND c.ordno = b.ordno AND " & DBW("c.stscd =", pAction) & _
                " AND ( c.dcfg = '' or c.dcfg is null ) " & strTemp & " and c.ordcd = d.testcd and f.userid = b.orddoct and g.dpcd = b.deptcd and h.dpcd = i.wardno and d.workarea = '14'" & _
                " AND a.patno = i.patno " & _
                " AND i.stayyn = 'Y' " & _
                " Order  By WardId,hosil, PtId, c.statfg, c.spccd "
                
'                " Order  By WardId,hosilid, PtId, c.statfg, c.spccd "

End Function

Public Function SqlReadWardOrderE(ByVal Ptid As String, ByVal ReqDt As String, ByVal ReqTm As String, _
                                 Optional ByVal StatFg As String = "", Optional ByVal BussDiv As String = "", _
                                 Optional ByVal ReceptNo As String = "", Optional ByVal OrdDiv As String = "", Optional ByVal ordTest As String = "") As String
 
    Dim tmpStr      As String
    Dim tmpStr1     As String
    Dim tmpStr2     As String
    Dim strSQL(2)   As String

    tmpStr = "": tmpStr1 = "": tmpStr2 = ""
    
    If StatFg <> "" Then tmpStr = tmpStr & " AND   " & DBW("b.statfg = ", StatFg)       '���޿���
    If BussDiv <> "" Then tmpStr = tmpStr & " AND  " & DBW("a.bussdiv = ", BussDiv)     '�ܷ���������
    If ordTest <> "" Then tmpStr2 = " AND  " & DBW("c.testcd = ", ordTest)       'Ư���˻��׸�
    
    If BussDiv = enBussDiv.BussDiv_OutPatient Then  '�ܷ�
        tmpStr = tmpStr & " AND  " & DBW("a.orddt = ", ReqDt)
    Else
'        tmpStr1 = " AND a.reqdt||a.reqtm<='" & ReqDt & ReqTm & "'"
        tmpStr1 = " AND a.reqdt = '" & ReqDt & "'" & " AND a.reqtm <= '" & ReqTm & "'"
    End If
    
    If ReceptNo <> "" Then tmpStr = tmpStr & " AND (b.paydt<>'' or b.paydt is not null)  "

'

    strSQL(1) = " SELECT c.testnm, c.abbrnm5, c.testdiv, c.workarea, b.spccd, f.storecd, b.statfg, b.paydt, a.reqdt" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "a.reqtm as ColTm, " & _
                "        d.field3 as SpcNm, d.field5 as SpcNm5, d.field1 as MultiFg, d.field2 as SpcGrp, b.orddt, b.ordno, b.ordseq, b.ordcd, b.mesg, " & _
                "        a.ordtm, a.reqdt, a.reqtm, a.orddoct, a.majdoct, a.receptno, a.orddiv,  a.deptcd,  f.statflags, " & _
                         FUNC_CONVERT("num", "f.labelcnt") & " as labelcnt, a.bedindt as BedInDt, a.wardid as WardId, a.roomid as RoomId, a.hosilid,  " & _
                "        '' as bedid, '' as fzfg " & _
                " FROM " & T_LAB001 & " c, " & T_LAB032 & " d, " & _
                           T_LAB004 & " f, " & T_LAB102 & " b, " & T_LAB101 & " a " & _
                " WHERE " & DBW("a.ptid = ", Ptid) & _
                " AND   " & DBW("a.donefg =", enStsCd.StsCd_LIS_Order) & _
                " AND  " & DBW("a.orddiv = ", LIS_ORDDIV) & tmpStr1 & _
                " AND    b.ptid  = a.ptid " & _
                " AND    b.orddt = a.orddt " & _
                " AND    b.ordno = a.ordno " & _
                " AND   " & DBW(" b.donefg =", enStsCd.StsCd_LIS_Order) & tmpStr & _
                " AND   (b.dcfg = '' or b.dcfg is null) " & _
                " AND    c.testcd  = b.ordcd " & _
                " AND    c.applydt = (SELECT max(applydt) FROM " & T_LAB001 & " WHERE testcd = c.testcd AND applydt <= '" & Format(Now, CS_DateDbFormat) & "') " & _
                " AND  " & DBJ(DBW("d.cdindex = ", LC3_Specimen)) & _
                " AND  " & DBJ("d.cdval1 =* b.spccd") & _
                " AND    f.testcd = b.ordcd AND f.spccd = b.spccd " & tmpStr2 & _
                " AND    f.applydt = (SELECT max(applydt) FROM " & T_LAB004 & " WHERE testcd = f.testcd  AND     spccd = f.spccd ) "

    '�������� : Phersis�˻�� ����(testdiv <> '3')...
    strSQL(2) = " SELECT c.testnm, c.abbrnm5, '3' as testdiv, 'XM' as workarea, b.spccd, '' as storecd, b.statfg, b.paydt, a.reqdt" & FUNC_CONCAT & "' '" & FUNC_CONCAT & "a.reqtm as ColTm, " & _
                "        '����' as SpcNm, '����' as SpcNm5, '' as MultiFg, '' as SpcGrp, b.orddt, b.ordno, b.ordseq, b.ordcd, b.mesg, " & _
                "        a.ordtm, a.reqdt, a.reqtm, a.orddoct, a.majdoct, a.receptno, a.orddiv, a.deptcd,  '' as statflags, " & _
                "        1 as labelcnt, a.bedindt as BedInDt, a.wardid as WardId, a.roomid as RoomId, a.hosilid,  " & _
                "        '' as bedid, '' as fzfg " & _
                " FROM " & T_BBS001 & " c, " & _
                           T_LAB102 & " b, " & T_LAB101 & " a " & _
                " WHERE " & DBW("a.ptid = ", Ptid) & _
                " AND   " & DBW("a.donefg =", enStsCd.StsCd_LIS_Order) & _
                " AND   " & DBW("a.orddiv = ", BBS_ORDDIV) & tmpStr1 & _
                " AND    b.ptid  = a.ptid " & _
                " AND    b.orddt = a.orddt " & _
                " AND    b.ordno = a.ordno " & _
                " AND   " & DBW(" b.donefg =", enStsCd.StsCd_LIS_Order) & tmpStr & _
                " AND   (b.dcfg = '' or b.dcfg is null) " & _
                " AND    c.testcd = b.ordcd " & _
                " AND  " & DBW("c.testdiv <> ", enTestDiv.TST_AboTest) & _
                " AND    c.applydt = (SELECT max(applydt) FROM " & T_BBS001 & " WHERE testcd = c.testcd AND applydt <= '" & Format(Now, CS_DateDbFormat) & "') "
 
    Select Case OrdDiv
        Case "B": SqlReadWardOrderE = strSQL(2)
        Case "L": SqlReadWardOrderE = strSQL(1)
        Case "W":
            If P_IncludeBBSSystem Then
                SqlReadWardOrderE = strSQL(1) & " UNION ALL " & strSQL(2)
            Else
                SqlReadWardOrderE = strSQL(1)
            End If
    End Select

    SqlReadWardOrderE = SqlReadWardOrderE & " Order By ColTm, orddt, ordno, ordcd "                          '<< D/C ó�� ���� >>
      
     
End Function


Public Function GetCollects2ORD101_Query(ByVal IDX As Integer, Optional ByVal pQuery As String) As String
Dim strSQL   As String
Dim strSQLT1 As String
    If IDX = 1 Then strSQLT1 = " and a.patno = '" & pQuery & "' "
    If IDX = 2 Then strSQLT1 = " and a.orddate  = to_date('" & pQuery & "', 'yyyymmdd') "
    
    strSQL = " select a.patno as ptid,  to_char(a.orddate, 'yyyymmdd') as orddt,  a.ordseqno as ordno,  to_char(a.ordtime, 'hh24miss') as ordtm,  decode(a.patsect, 'I', '2', '1') as bussdiv ,  to_char(a.meddate, 'yyyymmdd') as BEDINDT ,  to_char(a.hopedate, 'yyyymmdd') as REQDT,  DECODE(to_char(a.hopedate, 'hh24miss'), '000000', to_char(a.ordtime, 'hh24miss'), to_char(a.hopedate, 'hh24miss')) as REQTM,  a.meddept as DEPCD,  a.orddr   as ORDDOCT, a.chadr as MAJDOCT,  a.editid as ENTID,  to_char(a.editdate, 'yyyymmdd') as ENTDT,  to_char(a.editdate, 'hh24miss') as ENTTM,  'L' as ORDDIV,  a.repeatfg as REPEATFG,  a.orgaccno as ORGACCNO,  a.sporddiv   ,  to_char(a.donefg)   ,  null as RECEPTNO,  a.wardno as WARDID,  a.roomno as ROOMID,  a.roomno as HOSILID,  to_char(0) as ORDFG,  null as BEDID,  null as HOSCD,  null as TRUSTDT " & _
               " from mdexmort a" & _
              " where a.ordgrp = 'C1'  and substr(a.slipcd, 1, 1) in ('L', 'N', 'P')  and to_char(a.meddate, 'yyyymmdd')  > '20090101'" & strSQLT1 & _
   " union all select a.patno as ptid,  to_char(a.orddate, 'yyyymmdd') as orddt,  a.ordseqno as ordno,  to_char(a.ordtime, 'hh24miss') as ordtm,  decode(a.patsect, 'I', '2', '1') as bussdiv ,  to_char(a.meddate, 'yyyymmdd') as BEDINDT ,  to_char(a.orddate, 'yyyymmdd')  as REQDT,  to_char(a.orddate, 'hh24miss')                                                                                     as REQTM,  a.meddept as DEPCD,  a.orddr   as ORDDOCT, a.chadr as MAJDOCT,  a.editid as ENTID,  to_char(a.editdate, 'yyyymmdd') as ENTDT,  to_char(a.editdate, 'hh24miss') as ENTTM,  'B' as ORDDIV,  a.repeatfg as REPEATFG,  a.orgaccno as ORGACCNO,  a.sporddiv   ,  to_char(a.donefg)   ,  null as RECEPTNO,  a.wardno as WARDID,  a.roomno as ROOMID,  a.roomno as HOSILID,  to_char(0) as ORDFG,  null as BEDID,  null as HOSCD,  null as TRUSTDT " & _
               " from mdbldort a " & _
              " where to_char(a.meddate, 'yyyymmdd')  > '20090101' " & strSQLT1 & _
   " union all select a.patno as ptid,  to_char(a.orddate, 'yyyymmdd') as orddt,  a.ordseqno as ordno,  to_char(a.orddate, 'hh24miss') as ordtm,  '1'                              as bussdiv ,  to_char(a.admacpt, 'yyyymmdd') as BEDINDT ,  to_char(a.ordtime, 'yyyymmdd')  as REQDT,  to_char(a.ordtime, 'hh24miss')                                                                                     as REQTM,  a.meddept as DEPCD,  null      as ORDDOCT, null    as MAJDOCT,  a.entid  as ENTID,  to_char(a.entdate, 'yyyymmdd')  as ENTDT,  to_char(a.entdate, 'hh24miss')  as ENTTM,  'L' as ORDDIV,  a.repeatfg as REPEATFG,  a.orgaccno as ORGACCNO,  a.sporddiv   ,  to_char(a.donefg)   ,  null as RECEPTNO,  a.wardno as WARDID,  null     as ROOMID,  null     as HOSILID,  to_char(0) as ORDFG,  null as BEDID,  null as HOSCD,  null as TRUSTDT " & _
               " from su2examt a " & _
              " where a.ordgrp = 'C1'  and substr(a.slipcode, 1, 1) in ('L', 'N', 'P') " & strSQLT1 & _
   " union all select a.patno as ptid,  to_char(a.orddate, 'yyyymmdd') as orddt,  a.ordseqno as ordno,  to_char(a.orddate, 'hh24miss') as ordtm,  '1'                              as bussdiv ,  to_char(a.admacpt, 'yyyymmdd') as BEDINDT ,  to_char(a.ordtime, 'yyyymmdd')  as REQDT,  to_char(a.ordtime, 'hh24miss')                                                                                     as REQTM,  a.meddept as DEPCD,  null      as ORDDOCT, null    as MAJDOCT,  a.entid  as ENTID,  to_char(a.entdate, 'yyyymmdd')  as ENTDT,  to_char(a.entdate, 'hh24miss')  as ENTTM,  'L' as ORDDIV,  a.repeatfg as REPEATFG,  a.orgaccno as ORGACCNO,  a.sporddiv   ,  to_char(a.donefg)   ,  null as RECEPTNO,  a.wardno as WARDID,  null     as ROOMID,  null     as HOSILID,  to_char(0) as ORDFG,  null as BEDID,  null as HOSCD,  null as TRUSTDT " & _
               " from sg2examt a " & _
              " where a.ordgrp = 'C1'  and substr(a.slipcode, 1, 1) in ('L', 'N', 'P') " & strSQLT1

    GetCollects2ORD101_Query = strSQL
    
End Function




