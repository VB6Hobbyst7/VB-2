VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frm161WardCollect 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  'ũ�� ���� ��ȭ ����
   ClientHeight    =   9345
   ClientLeft      =   -330
   ClientTop       =   405
   ClientWidth     =   14640
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
   Icon            =   "Lis161.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9345
   ScaleWidth      =   14640
   ShowInTaskbar   =   0   'False
   Tag             =   "����ȯ�� �ϰ� ä��"
   Begin VB.CheckBox ChkMornFg 
      BackColor       =   &H00800000&
      Caption         =   "�ӻ󺴸� ��ħä��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   4155
      TabIndex        =   32
      Top             =   75
      Width           =   1980
   End
   Begin MedControls1.LisLabel LisLabel7 
      Height          =   300
      Left            =   6195
      TabIndex        =   25
      Top             =   45
      Width           =   3000
      _ExtentX        =   5292
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
      Caption         =   "ä�� �Ͻ�"
      LeftGab         =   100
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   1425
      Left            =   6195
      TabIndex        =   26
      Top             =   270
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
         Left            =   60
         TabIndex        =   36
         Top             =   255
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
         TabIndex        =   30
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
         TabIndex        =   29
         Top             =   405
         Width           =   1035
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   27
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
         TabIndex        =   28
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
         Format          =   38141955
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
      Left            =   13140
      Style           =   1  '�׷���
      TabIndex        =   3
      Tag             =   "0"
      Top             =   8535
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
      TabIndex        =   2
      Tag             =   "0"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ä   �� (&S)"
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
      Left            =   10500
      Style           =   1  '�׷���
      TabIndex        =   1
      Tag             =   "0"
      Top             =   8535
      Width           =   1320
   End
   Begin MSComctlLib.ProgressBar pbrPtCnt 
      Height          =   150
      Left            =   9450
      TabIndex        =   0
      Top             =   2790
      Width           =   4845
      _ExtentX        =   8546
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   300
      Left            =   75
      TabIndex        =   4
      Top             =   45
      Width           =   6075
      _ExtentX        =   10716
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
   Begin FPSpread.vaSpread tblPtList 
      Height          =   6990
      Left            =   75
      TabIndex        =   5
      Top             =   2010
      Width           =   9090
      _Version        =   196608
      _ExtentX        =   16034
      _ExtentY        =   12330
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
      MaxCols         =   23
      MaxRows         =   50
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      SpreadDesigner  =   "Lis161.frx":08CA
      TextTip         =   4
      ScrollBarTrack  =   3
   End
   Begin MedControls1.LisLabel LisLabel6 
      Height          =   300
      Left            =   75
      TabIndex        =   6
      Top             =   1695
      Width           =   9105
      _ExtentX        =   16060
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
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   300
      Left            =   9330
      TabIndex        =   7
      Top             =   45
      Width           =   5115
      _ExtentX        =   9022
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
      Caption         =   "��� �ɼ�"
      LeftGab         =   100
   End
   Begin MedControls1.LisLabel LisLabel5 
      Height          =   300
      Left            =   9345
      TabIndex        =   16
      Top             =   2370
      Width           =   5115
      _ExtentX        =   9022
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
      Caption         =   "���� ��Ȳ"
      LeftGab         =   100
   End
   Begin VB.Frame fraPrtOption 
      BackColor       =   &H00DBE6E6&
      Height          =   2100
      Left            =   9330
      TabIndex        =   8
      Top             =   270
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
         TabIndex        =   33
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
         Top             =   1140
         Width           =   3180
      End
      Begin VB.TextBox txtCopy 
         Alignment       =   1  '������ ����
         Height          =   345
         Left            =   3255
         TabIndex        =   10
         Top             =   1515
         Visible         =   0   'False
         Width           =   750
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   360
         Left            =   4020
         TabIndex        =   9
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
         TabIndex        =   14
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
         TabIndex        =   15
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
      Height          =   1425
      Left            =   75
      TabIndex        =   20
      Top             =   270
      Width           =   6090
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
         Left            =   4665
         Style           =   1  '�׷���
         TabIndex        =   23
         Tag             =   "0"
         Top             =   675
         Width           =   1320
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
         Left            =   2295
         Style           =   1  '�׷���
         TabIndex        =   22
         Tag             =   "WardID"
         Top             =   270
         Width           =   360
      End
      Begin VB.TextBox txtWardID 
         Appearance      =   0  '���
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   885
         MaxLength       =   9
         TabIndex        =   21
         Top             =   270
         Width           =   1395
      End
      Begin MSComCtl2.DTPicker dtpToTime 
         Height          =   375
         Left            =   885
         TabIndex        =   24
         Top             =   750
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   661
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
         Format          =   38141952
         CurrentDate     =   36328
      End
      Begin MedControls1.LisLabel lblWardNm 
         Height          =   315
         Left            =   2700
         TabIndex        =   31
         Top             =   285
         Width           =   2835
         _ExtentX        =   5001
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
         BorderStyle     =   0
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   0
         Left            =   105
         TabIndex        =   37
         Top             =   270
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
         TabIndex        =   38
         Top             =   750
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
   End
   Begin Crystal.CrystalReport CReport 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DBE6E6&
      Height          =   5460
      Left            =   9360
      TabIndex        =   17
      Top             =   3030
      Width           =   5100
      Begin FPSpread.vaSpread tblCount 
         Height          =   5340
         Left            =   2100
         TabIndex        =   18
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
         SpreadDesigner  =   "Lis161.frx":13BC
      End
      Begin MedControls1.LisLabel LisLabel3 
         Height          =   255
         Index           =   5
         Left            =   1350
         TabIndex        =   19
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
         TabIndex        =   34
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
         TabIndex        =   35
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
         TabIndex        =   39
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
         TabIndex        =   40
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
   Begin VB.Shape Shape1 
      BackStyle       =   1  '�������� ����
      BorderColor     =   &H00808080&
      FillColor       =   &H00D8DEDA&
      FillStyle       =   0  '�ܻ�
      Height          =   330
      Index           =   1
      Left            =   9345
      Shape           =   4  '�ձ� �簢��
      Top             =   2700
      Width           =   5100
   End
End
Attribute VB_Name = "frm161WardCollect"
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


Private Sub chkCol_Click()
    If chkCol.Value = 0 Then
        dtpColDtTm.Value = GetSystemDate
        dtpColDtTm.Enabled = False
    Else
        dtpColDtTm.Enabled = True
    End If
    
End Sub

Private Sub cmdClear_Click()
    Call ClearRtn(1)
    On Error GoTo Err_Trap
    txtWardID.SetFocus
Err_Trap:
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
    
On Error GoTo Err_Trap
    txtWardID.SetFocus
    chkPrintFg.Value = 0
    optOption(0).Value = True
    
Err_Trap:
    Resume Next
End Sub

Private Sub Form_Load()
    IsFirst = True

    If P_MornCollection = False Then
        ChkMornFg.Visible = False
        chkCol.Visible = False
    Else
        chkCol.Visible = False
        optApplyColTm(0).Visible = False
        optApplyColTm(1).Visible = False
    End If
   Set objMySql = New clsLISSqlCollection
   Set objLISCollect = New clsLISCollectioin
   
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
            .Col = 15   'for LIS
            If Trim(.Value) <> "" Then Call DoCollectionForLIS(i)
            
            If pbrPtCnt.Value + 100 >= pbrPtCnt.Max Then pbrPtCnt.Max = pbrPtCnt.Value + 100
            pbrPtCnt.Value = pbrPtCnt.Value + 100
            DoEvents
            .Col = 16
            
            .Col = 17   'for BBS
            If Trim(.Value) <> "" Then Call DoCollectionForBBS(i)
            
            If pbrPtCnt.Value + 100 >= pbrPtCnt.Max Then pbrPtCnt.Max = pbrPtCnt.Value + 100
            pbrPtCnt.Value = pbrPtCnt.Value + 100
            DoEvents


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
         cmdSave.Enabled = True
         Call cmdClear_Click
         MsgBox "ó���� ����Ÿ�� �����ϴ�..", vbInformation, "Message"
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

End Sub

Private Sub SetLock(ByVal blnLock As Boolean)
    'Locking...
    txtWardID.Enabled = Not blnLock
    txtWardID.BackColor = IIf(blnLock, &H8000000F, vbWhite)
    cmdWardList.Enabled = Not blnLock
    dtpToTime.Enabled = Not blnLock
    cmdGetOrders.Enabled = Not blnLock
End Sub

Private Sub DoCollectionForBBS(ByVal Row As Long)
    Dim objDIC          As clsDictionary
    Dim objBBSCollect   As clsBBSCollection

    Dim strPtid     As String       'ȯ��id
    Dim strPtnm     As String       'ȯ�ڸ�
    Dim strColID    As String      'ä����
    Dim strColDt    As String      'ä����
    Dim strColTm    As String      'ä���Ͻ�
    Dim strBuildCd  As String
    Dim strHosilid  As String
    Dim strStatFg   As String
    Dim lngErCnt    As Long
    Dim lngGcnt     As Long
    Dim lngBldRow   As Long
    Dim j           As Long
    
    Set objDIC = New clsDictionary
    Set objBBSCollect = New clsBBSCollection
    
    strBuildCd = ObjSysInfo.BuildingCd
    
'    Dim objBld As clsBasisData
    Dim strBld As String
    
'    Set objBld = New clsBasisData
    strBld = GetBuildNm(strBuildCd)
'    Set objBld = Nothing
    
    strColID = ObjSysInfo.EmpId
    Call objBBSCollect.SetWardCol(txtWardID.Text, sWorkDt, sWorkTm)

    With tblPtList
        .Row = Row
        .Col = 3: strPtid = .Value
        .Col = 4: strPtnm = .Value
        .Col = 5
        If .Value = "��" Then   '����
            lngErCnt = lngErCnt + 1
        Else
            lngGcnt = lngGcnt + 1
        End If
        .Col = 23:  strStatFg = IIf(.Value = "1", "1", "")
        .Col = 12:  strHosilid = Trim(.Value)
        .Col = 19:  strColDt = Format(.Text, "YYYYMMDD")
        .Col = 20:  strColTm = Format(.Text, "HHMMss")
        objDIC.Clear
        objDIC.FieldInialize "ptid", "ptnm,coldt,coltm,colid,bussdiv,buildcd,hosilid,statfg"

        objDIC.AddNew strPtid, Join(Array(strPtnm, strColDt, strColTm, strColID, _
                                    enBussDiv.BussDiv_InPatient, strBuildCd, strHosilid, strStatFg), COL_DIV)
        
            
        If objDIC.RecordCount > 0 Then
            objBBSCollect.WardId = txtWardID.Text
            If objBBSCollect.Set_Collect(objDIC, strBuildCd, , True) Then     '�ϰ�ä����������
'                Call ObjLISComCode.Building.KeyChange(strBuildCd)
                
                lngBldRow = 0
                For j = 1 To tblCount.DataRowCnt
                    tblCount.Row = j: tblCount.Col = 3
                    If tblCount.Value = strBuildCd Then
                        lngBldRow = j
                        Exit For
                    End If
                Next

                If lngBldRow = 0 Then lngBldRow = tblCount.DataRowCnt + 1
                tblCount.Row = lngBldRow
                tblCount.Col = 1: tblCount.Text = strBld 'ObjLISComCode.Building.Fields("buildnm")
                tblCount.Col = 2: tblCount.Text = Val(tblCount.Text) + 1
                tblCount.Col = 3: tblCount.Text = strBuildCd

                Dim objBar As New clsDictionary

                Set objBar = objBBSCollect.BldDic
                If objBar.RecordCount > 0 Then
                    BarCode_Print objBar
                    blnCollectFg = True
                End If
            End If
        End If
    End With

    Set objBBSCollect = Nothing
    Set objDIC = Nothing
    Set objBar = Nothing
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
    objLISCollect.MornFg = ChkMornFg.Value      '��ħä������

    ReDim tmpData(0 To 16)
    
    With tblPtList
        .Row = Row
                    tmpData(0) = Mid(Format(Now, "YYYY"), 4)
        .Col = 3:   tmpData(1) = .Value                                     'ȯ��ID
        .Col = 4:   tmpData(2) = .Value                                                 'ȯ�ڸ�
        .Col = 14:  tmpData(3) = .Value                                                 'ȯ�ڼ���
        .Col = 7:
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
        .Col = 8:   tmpData(5) = .Value                                 '�Կ���
                    tmpData(6) = Format(Now, CS_DateDbFormat)           '�Է���
                    tmpData(7) = Format(Now, CS_TimeDbFormat)           '�Է½ð�
                    tmpData(8) = ObjSysInfo.EmpId                       '�Է���
                    tmpData(9) = ""                                     '��������ȣ
                    tmpData(10) = Format(Now, CS_DateDbFormat)          'ä����
                    objLISCollect.ColTm = Format(GetSystemDate, "HHMMSS")     'ä����
                    tmpData(11) = ObjSysInfo.EmpId                      'ä����
        .Col = 2:   tmpData(12) = .Value                                '����ID
        .Col = 12:  tmpData(13) = .Value                                '����ID
        .Col = 13:  tmpData(14) = .Value                                'ȣ��ID
                    tmpData(15) = ""                                    'ħ��ID
                    tmpData(16) = ObjSysInfo.BuildingCd                 '** ä���� ����Ǵ� �ǹ��ڵ�
        
        Call objLISCollect.SetColData(tmpData)
        
        .Col = 22:  blnMornCol = Choose(Val(.Text) + 1, False, True)
        
        .Col = 9:   tmpDeptCd = .Value                        '�����
        .Col = 10:  tmpOrdDoct = .Value                       'ó����
        .Col = 11:  tmpMajDoct = .Value                       '��ġ��
    End With


    tmpDate = Format(dtpToTime.Value, CS_DateDbFormat)
    tmpTime = "235959"
       
    
    ' ó�泻�� �˻�
    If blnMornCol Then
        SqlStmt = objMySql.SqlReadOrderForMornCol(objLISCollect.Ptid, tmpDate, tmpTime)
    Else
        SqlStmt = objMySql.SqlReadWardOrder(objLISCollect.Ptid, tmpDate, tmpTime, , _
                                            enBussDiv.BussDiv_InPatient, , LIS_ORDDIV)
    End If
    
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
            tmpData(4) = Trim(.Fields("StatFg").Value)
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
'        Dim objBld As clsBasisData
        Dim strBld As String
        
'        Set objBld = New clsBasisData
        strBld = GetBuildNm(ObjSysInfo.BuildingCd)
'        Set objBld = Nothing
        
        DoEvents
         '* Delivery Location �� Count
         For i = 1 To objLISCollect.ColCount
            Call objLISCollect.GetLabNumbers(i, sWorkarea, sAccdt, iAccseq, sBuildCd)
'            Call ObjLISComCode.Building.KeyChange(sBuildCd)
           
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

    If Trim(txtWardID.Text) = "" Then
        MsgBox "����ID�� �Է��ϼ���.", vbInformation, "��������"
        txtWardID.SetFocus
        Exit Sub
    End If
    
    '2001-11-07 : ������ �����ϰ�ä�� ���� ���� --------------------------------------------------
    
    Set objStatus = New jProgressBar.clsProgress
    With objStatus
        .Container = Me
        .Left = LisLabel1.Left
        .Top = LisLabel1.Top
        .Width = LisLabel1.Width
        .Height = 280
        .Message = "������ �����ϰ� ä�볻���� �����ϰ� �ֽ��ϴ�..."
'        .Choice = True
'        .Appearance = aPlate
'        .SetMyForm Me
'        .XWidth = LisLabel1.Width
'        .XPos = LisLabel1.Left
'        .YPos = LisLabel1.Top
'        .YHeight = 280
'        .ForeColor = &H864B24
'        .Msg = "������ �����ϰ�ä�� ������ �����ϰ� �ֽ��ϴ�..."
'        .Max = 100
'        .Value = 50
    End With

    Set objLISCollect = New clsLISCollectioin
    If Not objLISCollect.Archive_WardColData(txtWardID.Text) Then
        MsgBox "�����ϰ�ä�� ���� Archive�� ������ �߻��߽��ϴ�." & vbCrLf & _
                "����� Ȥ�� �ӻ󺴸����� �����ٶ��ϴ�. (��" & ObjSysInfo.HelpLine & ")", vbCritical, "�����߻�"
    '---------------------------------------------------------------------------------------------
    End If
    Set objStatus = Nothing
    Set objLISCollect = Nothing
    
    If ChkMornFg.Value = 1 Then
        Resp = MsgBox("�ӻ󺴸� ��ħä�� �۾��� �����Ͻðڽ��ϱ�?", vbQuestion + vbYesNo, "��ħä��")
        If Resp = vbNo Then Exit Sub
    End If
    
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

    If ChkMornFg.Value = 1 Then
        SqlStmt = objMySql.SqlOrderForMornCol(tmpDate, tmpTime, txtWardID.Text)
    Else
        SqlStmt = objMySql.SqlWardOrder(tmpDate, tmpTime, txtWardID.Text)
    End If
    
    Set Rs = New Recordset
    Rs.Open SqlStmt, DBConn
    
    If Rs.EOF Then
'        MsgBox "ó�� �˻��� ������ �߻��߽��ϴ�. " & _
'               "����� Ȥ�� �ӻ󺴸����� �����ٶ��ϴ�.", vbExclamation ', "�����߻�"
        GoTo Err_Trap
    End If

    If Not Rs.EOF Then
        Call DisplayOrders(Rs, objProgress)
    End If
    
    Call BBSAddSpecimenChk(Format(GetSystemDate, "yyyymmdd"), Trim(txtWardID.Text))

    'ó�泻�� Display
    cmdSave.Enabled = True
    blnCleanFg = False

    DoEvents

    tblPtList.SetFocus

Err_Trap:
    Set Rs = Nothing
    Set objProgress = Nothing

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
        .MaxRows = IIf(objRs.RecordCount < 29, 29, objRs.RecordCount)
        .Row = 1

        intPtCount = 0

        For i = 1 To objRs.RecordCount

            If tmpPtId <> Trim(objRs.Fields("PtId")) Then

                If Not objPrgBar Is Nothing Then objPrgBar.Value = objPrgBar.Value + 50
                DoEvents

                intPtCount = intPtCount + 1
                .Row = intPtCount
                .Col = 2: .Text = "" & objRs.Fields("WardId").Value                 '����ID
                .Col = 3: .Text = "" & objRs.Fields("PtId").Value                   'ȯ��ID
                .Col = 4: .Text = "" & objRs.Fields("PtNm").Value                   '����
                .Col = 7: .Text = "" & objRs.Fields("DOB").Value                    '�������
                .Col = 8: .Text = "" & objRs.Fields("BedInDt").Value                '�Կ���
                
                .Col = 14:
                
                .Text = Trim("" & objRs.Fields("Sex").Value)
                If IsNumeric(.Text) Then
                    .Text = Choose((Val(.Text) Mod 2) + 1, "F", "M")
                End If
                
                
                tmpPtId = "" & objRs.Fields("PtId").Value
            End If

            .Col = 9: .Text = "" & objRs.Fields("DeptCd").Value                     '�����
            .Col = 10: .Text = "" & objRs.Fields("OrdDoct").Value                   'ó����
            .Col = 11: .Text = "" & objRs.Fields("MajDoct").Value                   '��ġ��
            .Col = 12: .Text = "" & objRs.Fields("HosilId").Value                   '����ID
            .Col = 13: .Text = "" & objRs.Fields("RoomId").Value                    'ȣ��ID

            tmpStatFg = "" & objRs.Fields("StatFg").Value                           '���޿���
            tmpOrdDiv = "" & objRs.Fields("orddiv").Value                           'ó�汸��
            tmpSpcCd = "" & objRs.Fields("SpcCd").Value                             '��ü
            
            If tmpOrdDiv = BBS_ORDDIV Then .Col = 23: .Value = tmpStatFg
            
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
            If tmpStatFg = "1" Then     '���ް�ü
                .Col = 5
                If InStr(1, .Text, tmpSpcCd) = 0 Then
                    .Text = .Text & tmpSpcCd & ", "
                End If

                .Col = 22: .Text = "0"
            Else
                .Col = 6

                If InStr(1, .Text, tmpSpcCd) = 0 Then
                    .Text = .Text & tmpSpcCd & ", "
                End If
                If ChkMornFg.Value = 1 Then
                    .Col = 22: .Text = "1"
                Else
                    .Col = 22: .Text = "0"
                End If
            End If

            Select Case tmpOrdDiv
            Case LIS_ORDDIV:   '�ӻ�
                .Col = 15: .ForeColor = vbRed: .Text = "��"     'ó�汸�С��
            Case BBS_ORDDIV:   '����
                .Col = 17: .ForeColor = vbRed: .Text = "��"     'ó�汸�С��
                If objGetSql.Blood_Existence(tmpPtId, Format(GetSystemDate, "yyyyMMdd"), _
                                            Format(GetSystemDate, "HHmm")) = True Then
                    .Col = 18: .ForeColor = vbBlue: .Value = "�ű�"
                Else
                    .Col = 18: .ForeColor = DCM_Gray: .Value = "����"
                End If

            End Select
            .Col = 19: .Text = Format(GetSystemDate, "YY-MM-DD")
            .Col = 20: .Text = Format(GetSystemDate, "HH:MM")
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

Private Function BBSAddSpecimenChk(ByVal OrdDt As String, WardId As String) As Boolean
'���������� ä��������߿� ��ü �߰� ����ڰ� ���ԵǾ� �ִ��� �Ǵ��ؼ� �����ش�.
'��ü �߰� ����ڴ� �̹� ������ ȯ�ڸ� �������� �ҷ��´�.
'�߰���û���� ������ ���� ��¥�� �������� �۰ų� ���� �͸��� ������� �Ѵ�.

    Dim objGetSql   As clsBBSCollection
    Dim Rs          As Recordset
    Dim strErChk    As String
    Dim strPtid     As String
    Dim strColDt    As String
    Dim strColTm    As String
    Dim cnt         As Integer


    BBSAddSpecimenChk = True
    strColDt = Format(GetSystemDate, "yyyy-mm-dd")
    strColTm = Format(GetSystemDate, "HH:mm")
    
    Set Rs = New Recordset
    Set objGetSql = New clsBBSCollection
    
    Set Rs = objGetSql.Get_SpcAdd(UCase(WardId))

    If Not Rs.EOF Then
        With tblPtList
            Do Until Rs.EOF
                If DupCheck("" & Rs.Fields("ptid").Value) = False Then
                    If .DataRowCnt <= .MaxRows Then .MaxRows = .MaxRows + 1
                    .Row = .DataRowCnt + 1
                    .ForeColor = vbBlue
                    .Col = 2: .Value = WardId   'lblWardNm.Caption
                    .Col = 3: .Value = "" & Rs.Fields("ptid").Value: strPtid = Trim("" & Rs.Fields("ptid").Value)
                    .Col = 4: .Value = "" & Rs.Fields("ptnm").Value
                    strErChk = objGetSql.ER_Chk(strPtid, "" & Rs.Fields("orddt").Value)
                    .Col = 5: .Value = IIf(strErChk = "1", "��", "")
                    .Col = 6: .Value = IIf(strErChk = "0", "��", "")
                    .Col = 7: .Value = "" & Rs.Fields("dob").Value
                    .Col = 8: .Value = "" & Rs.Fields("bedindt").Value
                    .Col = 14: .Text = Choose((Val("" & Rs.Fields("Sex")) Mod 2) + 1, "F", "M") '����
                    Select Case "" & Rs.Fields("orddiv").Value
                    Case "L":   '�ӻ�
                        .Col = 15: .ForeColor = vbRed: .Text = "��"     'ó�汸�С��
                    Case "B":   '����
                        .Col = 17: .ForeColor = vbRed: .Text = "��"     'ó�汸�С��
                    End Select
                    .Col = 18: .Value = "�߰�"

                    .Col = 19: .Text = Format(dtpColDtTm.Value, "YY-MM-DD")
                    .Col = 20: .Value = Format(dtpColDtTm.Value, "HH:MM:SS")
                    .Col = 9: .Text = "" & Rs.Fields("DeptCd").Value       '�����
                    .Col = 10: .Text = "" & Rs.Fields("OrdDoct").Value     'ó����
                    .Col = 11: .Text = "" & Rs.Fields("MajDoct").Value     '��ġ��
                    .Col = 12: .Text = "" & Rs.Fields("RoomId").Value      '����ID
                    .Col = 13: .Text = "" & Rs.Fields("HosilId").Value     'ȣ��ID
                    cnt = cnt + 1
                Else
                    '�߰�ä����, �Ϲ�ä���� ���ÿ� �߻��Ѱ��
                    .Col = 21: .Value = "*"
                End If
                Rs.MoveNext
            Loop
        End With
    Else
        BBSAddSpecimenChk = False
    End If

    If cnt = 0 Then BBSAddSpecimenChk = False

    Set Rs = Nothing
    Set objGetSql = Nothing

End Function

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
        Call .LoadPopUp(GetSQLWard) ', 2700, Frame2.Left + cmdWardList.Left) ', ObjLISComCode.WardId)
        If .SelectedString <> "" Then
            txtWardID.Text = medGetP(.SelectedString, 1, ";")
            lblWardNm.Caption = medGetP(.SelectedString, 2, ";")
        End If
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
    txtWardID.Enabled = True
    txtWardID.BackColor = &H80000005
    cmdWardList.Enabled = True
    dtpToTime.Enabled = True
    cmdGetOrders.Enabled = True
    cmdSave.Enabled = False

    sWorkDt = "": sWorkTm = ""
    txtWardID.Text = ""
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
    
    strTitleNm = IIf(ChkMornFg.Value = 0, "���� ä�� ����Ʈ", "������ ��ħä�� ����Ʈ")
    
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
