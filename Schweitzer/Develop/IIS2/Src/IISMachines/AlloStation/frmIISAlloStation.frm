VERSION 5.00
Object = "{C8094403-41E7-4EF2-826E-2A56177BCC48}#1.0#0"; "MDIControls.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmIISAlloStation 
   BackColor       =   &H00DBE6E6&
   Caption         =   "AlloStation"
   ClientHeight    =   9180
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   15120
   BeginProperty Font 
      Name            =   "����"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   15120
   StartUpPosition =   3  'Windows �⺻��
   Begin FPSpread.vaSpread vasINPrint 
      Height          =   4845
      Left            =   3840
      TabIndex        =   37
      Top             =   1770
      Visible         =   0   'False
      Width           =   5835
      _Version        =   393216
      _ExtentX        =   10292
      _ExtentY        =   8546
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      DisplayColHeaders=   0   'False
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GridShowHoriz   =   0   'False
      GridShowVert    =   0   'False
      GridSolid       =   0   'False
      MaxCols         =   12
      MaxRows         =   78
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      ScrollBarMaxAlign=   0   'False
      ScrollBars      =   0
      ScrollBarShowMax=   0   'False
      SpreadDesigner  =   "frmIISAlloStation.frx":0000
      UserResize      =   0
   End
   Begin FPSpread.vaSpread vaSpread2 
      Height          =   2685
      Left            =   16020
      TabIndex        =   57
      Top             =   5460
      Visible         =   0   'False
      Width           =   3915
      _Version        =   393216
      _ExtentX        =   6906
      _ExtentY        =   4736
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      DisplayColHeaders=   0   'False
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GridShowHoriz   =   0   'False
      GridShowVert    =   0   'False
      GridSolid       =   0   'False
      MaxCols         =   13
      MaxRows         =   76
      Protect         =   0   'False
      ScrollBarMaxAlign=   0   'False
      ScrollBars      =   0
      ScrollBarShowMax=   0   'False
      SpreadDesigner  =   "frmIISAlloStation.frx":B2CF
      UserResize      =   0
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame2"
      Height          =   2325
      Left            =   16560
      TabIndex        =   51
      Top             =   1740
      Visible         =   0   'False
      Width           =   6555
      Begin VB.TextBox txtRcv 
         Height          =   915
         Left            =   480
         MultiLine       =   -1  'True
         TabIndex        =   54
         Top             =   870
         Width           =   4605
      End
      Begin VB.TextBox txtSend 
         Height          =   315
         Left            =   390
         TabIndex        =   53
         Top             =   450
         Width           =   4575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   555
         Left            =   5160
         TabIndex        =   52
         Top             =   1020
         Width           =   1185
      End
      Begin VB.Label lblFilenm 
         Height          =   315
         Left            =   2610
         TabIndex        =   55
         Top             =   1860
         Width           =   2445
      End
   End
   Begin VB.Timer tmrAllo 
      Left            =   3060
      Top             =   -60
   End
   Begin VB.TextBox txtBarNo 
      Alignment       =   2  '��� ����
      Appearance      =   0  '���
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4950
      TabIndex        =   47
      Text            =   "123456789011"
      Top             =   120
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   2685
      Left            =   4980
      TabIndex        =   41
      Top             =   1380
      Visible         =   0   'False
      Width           =   4485
      Begin VB.TextBox txtWorkNo 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4620
         MaxLength       =   9
         TabIndex        =   43
         Top             =   150
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtFileNm 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3480
         TabIndex        =   42
         Text            =   "import"
         Top             =   330
         Visible         =   0   'False
         Width           =   1785
      End
      Begin MSComCtl2.DTPicker dtpFromDt 
         Height          =   330
         Left            =   300
         TabIndex        =   44
         Top             =   660
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   582
         _Version        =   393216
         Format          =   21364737
         CurrentDate     =   38330
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "Start WorkNo : "
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   3060
         TabIndex        =   46
         Top             =   225
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "���ϸ� : "
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   2610
         TabIndex        =   45
         Top             =   405
         Visible         =   0   'False
         Width           =   810
      End
   End
   Begin FPSpread.vaSpread vasFDPrint 
      Height          =   2685
      Left            =   11070
      TabIndex        =   38
      Top             =   3960
      Visible         =   0   'False
      Width           =   3915
      _Version        =   393216
      _ExtentX        =   6906
      _ExtentY        =   4736
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      DisplayColHeaders=   0   'False
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GridShowHoriz   =   0   'False
      GridShowVert    =   0   'False
      GridSolid       =   0   'False
      MaxCols         =   12
      MaxRows         =   78
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      ScrollBarMaxAlign=   0   'False
      ScrollBars      =   0
      ScrollBarShowMax=   0   'False
      SpreadDesigner  =   "frmIISAlloStation.frx":16148
      UserResize      =   0
   End
   Begin VB.FileListBox FileAllo 
      Height          =   870
      Left            =   2250
      Pattern         =   "*.patlist"
      TabIndex        =   33
      Top             =   8070
      Visible         =   0   'False
      Width           =   2805
   End
   Begin FPSpread.vaSpread tblReady 
      Height          =   3270
      Left            =   105
      TabIndex        =   21
      Top             =   945
      Width           =   6345
      _Version        =   393216
      _ExtentX        =   11192
      _ExtentY        =   5768
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
      MaxCols         =   6
      MaxRows         =   10
      OperationMode   =   2
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   13697023
      SpreadDesigner  =   "frmIISAlloStation.frx":219A9
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00DBE6E6&
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10230
      Style           =   1  '�׷���
      TabIndex        =   31
      Top             =   8580
      Width           =   1185
   End
   Begin MSComCtl2.DTPicker dtpFrDate 
      Height          =   315
      Left            =   1170
      TabIndex        =   29
      Top             =   540
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      _Version        =   393216
      Format          =   21364737
      CurrentDate     =   40270
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "��ȸ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4380
      TabIndex        =   28
      Top             =   540
      Width           =   855
   End
   Begin VB.CommandButton cmdGetRslt 
      BackColor       =   &H00DBE6E6&
      Caption         =   "��� ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9000
      Style           =   1  '�׷���
      TabIndex        =   27
      Top             =   8580
      Width           =   1185
   End
   Begin VB.Timer tmrResult 
      Left            =   5520
      Top             =   8520
   End
   Begin MSComDlg.CommonDialog AlloFile 
      Left            =   6000
      Top             =   8490
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdMakeWS 
      BackColor       =   &H00DBE6E6&
      Caption         =   "����Ʈ ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7770
      Style           =   1  '�׷���
      TabIndex        =   26
      Top             =   8580
      Visible         =   0   'False
      Width           =   1185
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   375
      Left            =   6548
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   107
      Width           =   8580
      _ExtentX        =   15134
      _ExtentY        =   661
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "�� ȯ������"
      Appearance      =   0
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00DBE6E6&
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
      Height          =   495
      Left            =   12698
      Style           =   1  '�׷���
      TabIndex        =   1
      Top             =   8567
      Width           =   1215
   End
   Begin VB.CommandButton cmdAlarm 
      BackColor       =   &H00DBE6E6&
      Caption         =   "&Alarm"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11483
      Style           =   1  '�׷���
      TabIndex        =   0
      Top             =   8567
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00DBE6E6&
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
      Height          =   495
      Left            =   13913
      Style           =   1  '�׷���
      TabIndex        =   2
      Top             =   8567
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   1290
      Left            =   6548
      TabIndex        =   4
      Top             =   407
      Width           =   8595
      Begin MedControls1.LisLabel lblPtId 
         Height          =   315
         Left            =   1245
         TabIndex        =   5
         Top             =   165
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   556
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
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   "00000001"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblDoctNm 
         Height          =   315
         Left            =   3930
         TabIndex        =   6
         Top             =   165
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   556
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
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   "�̻��"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblStatFg 
         Height          =   315
         Left            =   6795
         TabIndex        =   7
         Top             =   165
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
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
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   "����"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblName 
         Height          =   315
         Left            =   1245
         TabIndex        =   8
         Top             =   525
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   556
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
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   "�̻�� �Ʊ�"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblDeptNm 
         Height          =   315
         Left            =   3930
         TabIndex        =   9
         Top             =   525
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   556
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
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   "������"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblSpcNm 
         Height          =   315
         Left            =   6795
         TabIndex        =   10
         Top             =   525
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
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
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   "Blood"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblSexAge 
         Height          =   315
         Left            =   1245
         TabIndex        =   11
         Top             =   885
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   556
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
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   "���� / 29"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblWardNm 
         Height          =   315
         Left            =   3930
         TabIndex        =   12
         Top             =   885
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   556
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
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   "65����"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblPnlNm 
         Height          =   315
         Left            =   6795
         TabIndex        =   35
         Top             =   900
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
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
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin VB.Label lblGeneral 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "PANEL :"
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
         Index           =   5
         Left            =   5760
         TabIndex        =   36
         Top             =   975
         Width           =   630
      End
      Begin VB.Label lblGeneral 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "�� ü �� :"
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
         Index           =   4
         Left            =   5760
         TabIndex        =   20
         Top             =   600
         Width           =   900
      End
      Begin VB.Label lblGeneral 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "���޿��� :"
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
         Index           =   3
         Left            =   5760
         TabIndex        =   19
         Top             =   240
         Width           =   900
      End
      Begin VB.Label lblGeneral 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "��  �� : "
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
         Index           =   2
         Left            =   3105
         TabIndex        =   18
         Top             =   975
         Width           =   810
      End
      Begin VB.Label lblGeneral 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "����� :"
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
         Left            =   3105
         TabIndex        =   17
         Top             =   600
         Width           =   720
      End
      Begin VB.Label lblGeneral 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "ó���� :"
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
         Index           =   0
         Left            =   3105
         TabIndex        =   16
         Top             =   240
         Width           =   720
      End
      Begin VB.Label lblLotNo 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "����/���� :"
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
         Left            =   150
         TabIndex        =   15
         Top             =   975
         Width           =   990
      End
      Begin VB.Label lblLevel 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "��     �� :"
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
         Left            =   150
         TabIndex        =   14
         Top             =   600
         Width           =   990
      End
      Begin VB.Label lblControl 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "ȯ  �� ID :"
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
         Left            =   150
         TabIndex        =   13
         Top             =   240
         Width           =   990
      End
   End
   Begin MSCommLib.MSComm MSComm 
      Left            =   6578
      Top             =   8432
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin FPSpread.vaSpread tblComplete 
      Height          =   4110
      Left            =   90
      TabIndex        =   22
      Top             =   4725
      Width           =   6345
      _Version        =   393216
      _ExtentX        =   11192
      _ExtentY        =   7250
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
      MaxCols         =   14
      MaxRows         =   14
      OperationMode   =   2
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   13697023
      SpreadDesigner  =   "frmIISAlloStation.frx":21F32
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   375
      Left            =   105
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   105
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   661
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "�� �˻��� ����Ʈ"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel3 
      Height          =   405
      Left            =   105
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   4305
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   714
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "�� �˻�Ϸ� ����Ʈ"
      Appearance      =   0
   End
   Begin FPSpread.vaSpread tblResult 
      Height          =   6690
      Left            =   6555
      TabIndex        =   25
      Top             =   1710
      Width           =   8580
      _Version        =   393216
      _ExtentX        =   15134
      _ExtentY        =   11800
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
      MaxCols         =   8
      MaxRows         =   22
      OperationMode   =   2
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   13697023
      SpreadDesigner  =   "frmIISAlloStation.frx":227E6
   End
   Begin MSComCtl2.DTPicker dtpToDate 
      Height          =   315
      Left            =   2790
      TabIndex        =   30
      Top             =   540
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      _Version        =   393216
      Format          =   21364737
      CurrentDate     =   40270
   End
   Begin FPSpread.vaSpread tblexcel 
      Height          =   675
      Left            =   5610
      TabIndex        =   32
      Top             =   8610
      Visible         =   0   'False
      Width           =   675
      _Version        =   393216
      _ExtentX        =   1191
      _ExtentY        =   1191
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
      SpreadDesigner  =   "frmIISAlloStation.frx":22EF3
   End
   Begin MDIControls.MDIActiveX MDIActiveX 
      Left            =   7230
      Top             =   8550
      _ExtentX        =   847
      _ExtentY        =   794
   End
   Begin FPSpread.vaSpread vasExcel 
      Height          =   4605
      Left            =   900
      TabIndex        =   34
      Top             =   1830
      Visible         =   0   'False
      Width           =   9615
      _Version        =   393216
      _ExtentX        =   16960
      _ExtentY        =   8123
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
      SpreadDesigner  =   "frmIISAlloStation.frx":2314A
   End
   Begin MSComCtl2.DTPicker dtpResult 
      Height          =   315
      Left            =   5280
      TabIndex        =   49
      Top             =   510
      Visible         =   0   'False
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      _Version        =   393216
      Format          =   21364737
      CurrentDate     =   40270
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   2685
      Left            =   19980
      TabIndex        =   56
      Top             =   5460
      Visible         =   0   'False
      Width           =   3915
      _Version        =   393216
      _ExtentX        =   6906
      _ExtentY        =   4736
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      DisplayColHeaders=   0   'False
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GridShowHoriz   =   0   'False
      GridShowVert    =   0   'False
      GridSolid       =   0   'False
      MaxCols         =   13
      MaxRows         =   76
      Protect         =   0   'False
      ScrollBarMaxAlign=   0   'False
      ScrollBars      =   0
      ScrollBarShowMax=   0   'False
      SpreadDesigner  =   "frmIISAlloStation.frx":27609
      UserResize      =   0
   End
   Begin FPSpread.vaSpread vaSpread3 
      Height          =   3885
      Left            =   16050
      TabIndex        =   58
      Top             =   120
      Visible         =   0   'False
      Width           =   2685
      _Version        =   393216
      _ExtentX        =   4736
      _ExtentY        =   6853
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      DisplayColHeaders=   0   'False
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GridShowHoriz   =   0   'False
      GridShowVert    =   0   'False
      GridSolid       =   0   'False
      MaxCols         =   13
      MaxRows         =   78
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      ScrollBarMaxAlign=   0   'False
      ScrollBars      =   0
      ScrollBarShowMax=   0   'False
      SpreadDesigner  =   "frmIISAlloStation.frx":32A5E
      UserResize      =   0
   End
   Begin FPSpread.vaSpread vaSpread4 
      Height          =   2685
      Left            =   19680
      TabIndex        =   39
      Top             =   870
      Visible         =   0   'False
      Width           =   3915
      _Version        =   393216
      _ExtentX        =   6906
      _ExtentY        =   4736
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      DisplayColHeaders=   0   'False
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GridShowHoriz   =   0   'False
      GridShowVert    =   0   'False
      GridSolid       =   0   'False
      MaxCols         =   13
      MaxRows         =   78
      Protect         =   0   'False
      ScrollBarMaxAlign=   0   'False
      ScrollBars      =   0
      ScrollBarShowMax=   0   'False
      SpreadDesigner  =   "frmIISAlloStation.frx":3CE0C
      UserResize      =   0
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "�˻����� : "
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   5520
      TabIndex        =   50
      Top             =   420
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "���ڵ��ȣ : "
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   3660
      TabIndex        =   48
      Top             =   195
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00DBE6E6&
      Caption         =   "�� ��������"
      Height          =   180
      Left            =   90
      TabIndex        =   40
      Top             =   615
      Width           =   960
   End
End
Attribute VB_Name = "frmIISAlloStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------'
'   ���ϸ�  : frmIISMEDIWISS.frm
'   �ۼ���  : ������
'   ��  ��  : MEDIWISS �����
'   �ۼ���  : 2016-12-22
'   ��  ��  : 1.0.0
'   ��  ��  :
'       1. ���ֿ�������
'-----------------------------------------------------------------------------'

Option Explicit

'## tblReady�� Column Enum
Private Enum TReadyEnum
    ccNo = 1
    ccBarNo = 2
    ccAccNo = 3
    ccPtId = 4
    ccName = 5
End Enum

'## tblComplete�� Column Enum
Private Enum TCompleteEnum
    ccNo = 1:           ccBarNo = 2
    ccAccNo = 3:        ccPtId = 4
    ccName = 5:         ccSexAge = 6
    ccDoctNm = 7:       ccDeptNm = 8
    ccWardNm = 9:       ccStatFg = 10
    ccSpcNm = 11:       ccQcFg = 12
    ccSendCnt = 13:     ccResult = 14
End Enum

'## tblResult�� Column Enum
Private Enum TResultEnum
    ccTestNm = 1
    ccEqpResult = 2
    ccLISResult = 3
    ccUnit = 4
    ccHLDiv = 5
    ccDPDiv = 6
    ccRef = 7
    ccClass = 8
    '-- 2015.08.28 �߰�
    ccIntBase = 9
End Enum

'## Clear Enum
Private Enum ClearEnum
    ccAll = 1
    ccLabel = 2
End Enum

'## Popup Menu ID
Private Const DELETE    As Long = 1
Private Const DELETEALL As Long = 2

Private WithEvents mIntLib  As clsIISInterface   '�������̽� Ŭ����
Attribute mIntLib.VB_VarHelpID = -1
Private WithEvents mPopup   As clsIISPopup       '�˾��޴�
Attribute mPopup.VB_VarHelpID = -1

Private mIntErrors  As clsIISIntErrors           '�������̽� ���� �÷���
'Private mOrder      As clsIISIntOrder           '�������� Ŭ����

Private mEqpCd  As String   '����ڵ�
Private mEqpKey As String   '���Ű

'�ӽ� - �����
Dim strTransData    As Variant

Dim AdoCn           As ADODB.Connection
Dim AdoRS           As ADODB.Recordset

Private Const AM As Variant = -0.696669715055768
Private Const AN As Variant = 3.57268581287044
Private Const BM As Variant = -0.58804147563697
Private Const BN As Variant = 3.68672614195763
Private Const CM As Variant = -0.424608277908711
Private Const Cn As Variant = 3.85830192881032

Private Const VarA As Variant = 0
Private Const VarD As Variant = 150

Private gBarNo  As String

Private Const ALG_CMNT1 As Variant = "Ư���� �˷����� �Ұ��� ������ ����." & vbNewLine & _
                                     "(�˷����� �˻���, �������� �����Դϴ�.)"
                                     
Private Const ALG_CMNT2 As Variant = "Total IgE�� �����̳�, Specific IgE�� �Ϻ� �˷����� ���߹�����" & vbNewLine & _
                                     "�׿����� �缺������ ����. �Ϻ� �׿����� ������ �ǰ�, ���ѵ� ������" & vbNewLine & _
                                     "���� �缺���� ������ �˷����� ���߹����� ���� ���� ���." & vbNewLine & _
                                     "(������ �˷����� �׿��� �˷������� ������ �� �����Ƿ� ������," & vbNewLine & _
                                     "���ŷ� � ���� �ùٸ� ��ó�� ���� �������� ����� ��õ�մϴ�)"

Private Const ALG_CMNT3 As Variant = "�˻縦 ������ Specific IgE �� �׿��� ���ؼ� ���� ����� ���̳�, ����" & vbNewLine & _
                                     "�˷�������ȯ�� ���ϰ� �ǽɵ� ��� �˻縦 ������ �׸� �̿���" & vbNewLine & _
                                     "�߰����� Specific IgE �˻縦 �ǰ���." & vbNewLine & _
                                     "(Total IgE �� ������ ����˻簡 ����� Specific IgE�� ���� ������" & vbNewLine & _
                                     "�ƴϹǷ� �˷����� ��ȯ�� �ǽɵ� ��� �߰����� Specific IgE�˻縦" & vbNewLine & _
                                     "���մϴ�.)"

Private Const ALG_CMNT4 As Variant = "Total IgE�� Specific IgE �˻��� �缺�� ���̸� ȯ���� ����, ������," & vbNewLine & _
                                     "������ �Ұ��� Ÿ���ϰ� ������ �����ϸ� �˷����� ���� �׿����� ������" & vbNewLine & _
                                     "�Ǹ�, ���� ������ ȯ���� ��� �������� �Ǵ��� �Ǿ� ���� �˷����� ��ȯ" & vbNewLine & _
                                     "������ ���� ���ɼ��� �����Ƿ� ���ǰ� �����." & vbNewLine & _
                                     "(������ �˷����� �׿��� �˷����� ������ �� �����Ƿ� ������," & vbNewLine & _
                                     "���ŷ� � ���� �ùٸ� ��ó�� ���� �������� ����� ��õ�մϴ�.)"
Dim strBuffer       As String




Public Property Let EqpCd(ByVal vData As String)
    mEqpCd = vData
End Property

Public Property Let EqpKey(ByVal vData As String)
    mEqpKey = vData
End Property


Public Function ReadFileBinary(ByVal strFileName As String) As String

On Error GoTo errHandler
    Dim fsT, tFilePath As String

    Set fsT = CreateObject("ADODB.Stream")

    fsT.Type = 2
 
    fsT.Charset = "utf-8"
 
    fsT.Open
'    fsT.Type = adTypeBinary
'    fsT.Type = adTypeText
    fsT.LoadFromFile strFileName
   
    Dim strText As String
    strText = ""
    Do Until fsT.EOS
        strText = strText & fsT.ReadText(adReadLine) & vbLf     ' �ٹٲ� �߰�
    Loop
 

    fsT.Close

    ReadFileBinary = strText
    GoTo finish
 
errHandler:

    MsgBox (Err.Description)
    Exit Function

finish:

End Function

Private Sub OpenExcel()

    Dim strFile As String
    Dim i, iCnt As Integer
    Dim strTemp As String
    Dim varTmp  As Variant
    Dim xlApp As New Excel.Application
    Dim xlSheet As Excel.Worksheet
    Dim strPath As String
    Dim strDestFile As String
    Dim STM As ADODB.Stream
    
    AlloFile.DialogTitle = "�������� ����"
    AlloFile.InitDir = GetMEDIWISSConfig("ExportPath")
    AlloFile.ShowOpen
    
    If Len(AlloFile.FileName) > 0 Then
        xlApp.Workbooks.Open AlloFile.FileName
        strPath = AlloFile.FileName
    Else
        Exit Sub
    End If
    
    
    
    Dim strBuffer As String
    Dim strBuf      As String
    
    Open AlloFile.FileName For Input As #3

    strBuffer = ""

    strBuf = ReadFileBinary(AlloFile.FileName)
    
    Close #3
    
    'lngBufLen = Len(strBuf)
    
    
    
    Set xlSheet = xlApp.Worksheets("export")

'    Set xlSheet = xlApp.Worksheets("sheet1")

    With vasExcel
        .Action = ActionClear
        For iCnt = 1 To .MaxRows
            For i = 1 To .MaxCols
                'If xlSheet.Cells(iCnt, i) <> "" Then
                If Trim(Format(xlSheet.Cells(iCnt, 1), "####-##")) = "" Then
                    xlApp.Workbooks.Close
                    xlApp.Quit

                    Set xlSheet = Nothing
                    GoTo RST
                End If
                
'                Select Case i
'                    Case 1
'                        vasExcel.SetText i, iCnt, 1
'                    Case 2
'                        vasExcel.SetText i, iCnt, Trim(Format(xlSheet.Cells(iCnt + 3, i), "####-##"))
''                    Case 3
''                        vasExcel.SetText i, iCnt, Trim(xlSheet.Cells(iCnt + 3, i) & "��")
'                    Case 6
'                        vasExcel.SetText i, iCnt, Trim(Format(xlSheet.Cells(iCnt + 3, i), "######-#######"))
'                    Case Else
                        vasExcel.SetText i, iCnt, Trim(xlSheet.Cells(iCnt, i))
'                End Select
                'End If
            Next
        Next iCnt
    End With

RST:
   ' xlApp.Workbooks(strPath).Close
    xlApp.Quit
    
    Set xlSheet = Nothing
    
''    '��� ���� �̸��� ����
''    strDestFile = App.path & "\Log\" & Format(Now, "yyyymmdd-hhmm")
''    '������ ��� ����
''    FileCopy strPath, strDestFile
''
''    Kill strSrcfile
    'FileMEDIWISS.Refresh

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



    lngTotColCnt = 0
    lngTotRowCnt = 0
    
    
    '���� ����
    AlloFile.Filter = "Excel(*.xlsx)|*.xlsx|Excel(*.xls)|*.xls"
    AlloFile.Action = 1
    
    
    If AlloFile.FileTitle = "" Then
        Exit Sub
    End If
    
    xlApp.Workbooks.Open (Trim(AlloFile.FileName))
    
    lngSCnt = xlApp.Worksheets.Count
    
    '-- ��ü ��ũ��Ʈ �ҷ�����ͼ� '�ӽ�.txt' ���Ϸ� ����
    For i = 1 To lngSCnt
        Set XLappWS = xlApp.Worksheets(i)
        XLappWS.Activate
        lngSColCnt(i) = XLappWS.UsedRange.Columns.Count
        xlApp.DisplayAlerts = False
    
        '''xlApp.ActiveWorkbook.SaveAs App.Path & "\" & Trim(i) & ".txt", xlText, "", "", False, False '==>2000 + 2003 ����
'        xlApp.ActiveWorkbook.SaveAs "C:\" & Trim(i) & ".txt", xlText, "", "", False, False '==>2000 + 2003 ����
        xlApp.ActiveWorkbook.SaveAs "C:\" & Trim(i) + 10 & ".txt", xlText, "", "", False, False '==>2000 + 2003 ����
        
        
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
'    With vasExcel
'        .Col = 1: .Col2 = .MaxCols
'        .Row = 2: .Row2 = .DataRowCnt
'        .SortBy = 0
'        .SortKey(1) = 6       '����Ű ����ȣ
'        .SortKey(2) = 2       '����Ű ����ȣ
'
'        .SortKeyOrder(1) = SortKeyOrderAscending
'        .SortKeyOrder(2) = SortKeyOrderAscending
'
'        .Action = ActionSort
'    End With


'Dim SortKeys, SortKeyOrder As Variant
'
'    SortKeys = Array(6, 2)
'    SortKeyOrder = Array(6, 2)
'    ' Sort data in first five columns and rows by column 1 and 3
'    vasExcel.Sort 6, 2, 2, vasExcel.MaxRows, SS_SORT_BY_ROW, SortKeys, SortKeyOrder

End Sub

Public Sub FileUp(PropNo As Long, FilePath As String, FileName As String)
    Dim Cn As ADODB.Connection
    Dim Cmd As ADODB.Command
    Dim Object As ADODB.Parameter
'    Dim Fso As FileSystemObject
    Dim Chunk() As Byte
    Dim Sql As String
    Dim FileType As String  '����Ȯ���ڸ� ����.
    Dim Fd As Integer       '�����ڵ�
    Dim Flen As Long
    Dim szConn As String
    
    On Error GoTo ErrLog
    
    Dim Rs As ADODB.Recordset
    
    'Set Fso = Nothing
    Set Object = Nothing
    Set Cmd = Nothing
    Set Cn = Nothing
    
    
    'Up�� �ѱ����������� ���Ͽ� ODBC���� Oracle ODBC�� ������ ��
    'OleDB For ODBC�� ����Ͽ� �����Ѵ�.
    szConn = "Provider=MSDAORA.1;Data Source=MEDIWISS"

'        "Dbq=" & Mid(Dialog1.FileName, 1, InStrRev(Dialog1.FileName, "\")) & ";" & _

szConn = "Driver={Microsoft Text Driver (*.txt; *.csv)};" & _
        "Dbq=C:\;" & _
        "Extensions=asc,csv,tab,txt;       "




    Set Cn = New Connection 'Ŀ�ؼǰ�ü ����
    Set Cmd = New Command   'Ŀ�ǵ�ĳü ����
'
    Cn.ConnectionString = szConn
    
    Cn.Open
    
    Set Rs = New ADODB.Recordset
    
    '�ش��ϴ� PropNo�� ������ Update
    Sql = "SELECT * FROM export.csv;"
    Rs.Open Sql, Cn ', adOpenDynamic, adLockReadOnly ', adCmdText

            Dim i As Integer
            
    Do Until Rs.EOF
        For i = 1 To 20
        Debug.Print Rs.Fields(i).Value
        Next
        Rs.MoveNext
        
    Loop
    
    Rs.Close
    
    '����Ȯ���ڴ� ��ħǥ�� �����Ͽ� �ִ� 5�ڸ��� ����Ѵ�.
    FileType = Mid(FileName, InStrRev(FileName, "."))
    FileType = Trim(Left(FileType, 5))
    
    With Cmd
        .ActiveConnection = Cn
        .CommandText = Sql
        .CommandType = adCmdText

        .Parameters.Append .CreateParameter("@file_size", adInteger, adParamInput, , FileLen(FilePath & FileName))
        .Parameters.Append .CreateParameter("@file_type", adVarChar, adParamInput, 5, FileType)
        
    End With
    
    '������ �������Ϸ� �����Ͽ� �д´�.
    Fd = FreeFile

    Open FilePath & FileName For Binary Access Read As Fd

    Flen = LOF(Fd)

    '����ũ�Ⱑ 0�϶�
    If Flen = 0 Then
        Close Fd
        Set Cmd = Nothing
        Set Cn = Nothing

        MsgBox "Error while opening the file"
        Exit Sub
    End If

    Set Object = Cmd.CreateParameter("object", adLongVarBinary, adParamInput, Flen + 100)

    ReDim Chunk(1 To Flen)
    Get Fd, , Chunk()

    Object.AppendChunk Chunk()
    Cmd.Parameters.Append Object

    Close Fd
        
    Cmd.Execute
        
    Cn.Close
    
'    Set Fso = New FileSystemObject
    
'    '������ ������ BLOB ���̺�� �ø��� �����Ѵ�.
'    If Fso.FileExists(FilePath & FileName) Then
'        Fso.DeleteFile (FilePath & FileName)
'    End If
    
'    Set Fso = Nothing
    Set Object = Nothing
    Set Cmd = Nothing
    Set Cn = Nothing
    
    Exit Sub
    
ErrLog:
    'WriteEventLog "FileUp(" & PropNo & "," & FilePath & "," & FileName & ") " & vbCrLf & Err.Number & "  " & Err.Description
'    Set Fso = Nothing
    Set Object = Nothing
    Set Cmd = Nothing
    Set Cn = Nothing
End Sub

 


Private Sub cmdGetRslt_Click()

    
    tmrAllo.Enabled = False
    
    If Not Set_DbConnect_Jet Then
        MsgBox "AlloScan �����ͺ��̽� ��θ� Ȯ���ϼ���", vbOKOnly + vbCritical, Me.Caption
'        End
    Else
'        tmrResult.Interval = 1000 '60000
'        tmrResult.Enabled = True
    End If
    
    Call Get_SearchListIN
    Call Get_SearchListFD
    
    tmrAllo.Enabled = True
    
    
End Sub

'   Result List Recordset
Public Function Get_ResultList_IN() As ADODB.Recordset
    Dim strSql      As String

On Error GoTo ErrorTrap

             strSql = "SELECT * "
    strSql = strSql & "  FROM UNITS "
    'strSql = strSql & " WHERE EXAMDATE = '" & Format(dtpResult.Value, "yyyy-mm-dd") & "' "
    strSql = strSql & " WHERE EXAMDATE >= '" & Format(dtpFrDate.Value, "yyyy-mm-dd") & "' "
    strSql = strSql & "   AND EXAMDATE <= '" & Format(dtpToDate.Value, "yyyy-mm-dd") & "' "
    strSql = strSql & "   AND PANEL IN ('Inhalant','Standard') "
    strSql = strSql & " ORDER BY EXAMDATE,PATIENTID, PANEL "
    
    Set AdoRS = New ADODB.Recordset
    If Get_Recordset(AdoCn, strSql, AdoRS, "") Then
        Set Get_ResultList_IN = AdoRS
    Else
        Set Get_ResultList_IN = Nothing
    End If
    
    Set AdoRS = Nothing

Exit Function

ErrorTrap:
    Set AdoRS = Nothing

End Function

Private Sub Get_SearchListIN()
    Dim intRow      As Long
    Dim strAge      As String
    Dim strTransDt  As String
    Dim strHMsg     As String
    Dim strDMsg     As String
    
    Dim objIntInfo   As clsIISIntInfo    '�������̽� ��ü���� Ŭ����
    Dim objIntNms    As clsIISIntNms     '��� �˻��׸� �÷��� Ŭ����
    Dim objBuffer    As clsIISBuffer     '���� Ŭ����
    
    Dim vWorkNo      As Variant  'Spread�� WorkNo
    Dim vBarNo       As Variant  'Spread�� ���ڵ��ȣ
    Dim strRcvBuf    As String   '������ Data
    Dim strType      As String   '������ Record Type
    Dim strBarNo     As String   '������ BarNO
    Dim strWorkNo    As String   '������ WorkNo
    Dim strIntResult As String   '������ �˻���
    Dim strIntBase   As String   '������ ������ �˻��
    Dim strResult    As String   'LIS �˻���
    Dim strClass     As String   'Class �������
    Dim strClassA    As String   'Class �������
    Dim strMaxClass  As String   'Class �������
    
    Dim strTemp      As String
    Dim i            As Long
    Dim blnSameBar   As Boolean
    Dim intCnt       As Integer
    Dim strTIgE      As String   'Total IgE
    Dim strSIgE      As String   'Ŭ������ ���� ���� ��
    Dim strRemark    As String
    Dim SIgE         As String
    
    Dim Y, Y1, Y2, Y3, X1, X2
    
    Set objIntNms = mIntLib.IntNms
    
    On Error Resume Next
    
    Set AdoRS = Get_ResultList_IN
    
    If Not AdoRS.BOF Then
        intRow = 1: strTransDt = ""
        Do Until AdoRS.EOF
            strBarNo = AdoRS.Fields("PATIENTID").Value
            'If strBarNo = "170012010479" Then Stop
            
            Set objIntInfo = New clsIISIntInfo
            With objIntInfo
                'If UCase(Mid(AdoRS.Fields("PANEL").Value, 1, 2)) = "IN" Then
                    .SpcPos = "IN"
                'ElseIf UCase(Mid(AdoRS.Fields("PANEL").Value, 1, 2)) = "FO" Then
                '    .SpcPos = "FD"
                'Else
                '    .SpcPos = strIntBase
                'End If
                '.OrgBarNo = strBarNo
                '.BarNo = Mid(strBarNo, 1, 10)
                .BarNo = strBarNo
            End With
            
            Call GetOrder(strBarNo)
            
            'If strBarNo = "C150151603" Then Stop
            strSIgE = "0"
            strMaxClass = "0"
            strClassA = "0"
            strIntBase = ""
            
            For intCnt = 1 To 32
                If intCnt = 1 Then
                    'If UCase(Trim(AdoRS.Fields("PANEL").Value & "")) = "FOOD" Then
                    '    strIntBase = "FD"
                    'ElseIf UCase(Trim(AdoRS.Fields("PANEL").Value & "")) = "INHALANT" Then
                        strIntBase = "IN"
                    'ElseIf UCase(Trim(AdoRS.Fields("PANEL").Value & "")) = "ATOPY" Then
                    '    strIntBase = "AT"
                    'Else
                    '    strIntBase = ""
                    'End If
                End If
                
                strIntResult = ""
                strTemp = AdoRS.Fields("BANDVAL" & intCnt) & ""
                If strTemp <> "" Then
                    strIntBase = objIntInfo.SpcPos & mGetP(strTemp, 3, "|")
                    'strIntBase = UCase(strIntBase)
                    strIntResult = mGetP(strTemp, 1, "|")
                End If
                
'
''-- Class �߰�
'If strIntBase = "FD|IgE" Or strIntBase = "IN|IgE" Then
'    If strIntResult = "<100" Then
'        strClass = "1.0"
'    ElseIf strIntResult = "100-200" Then
'        strClass = "2.0"
'    ElseIf strIntResult = ">200" Then
'        strClass = "3.0"
'    End If
'Else
'    If intCnt + 2 <= UBound(varTmp) Then
'        strClass = varTmp(intCnt + 2)
'        strClass = Format(strClass, "0.0")
'    End If
'End If
'
''-- ������� Class���� ����
'strResult = strClass & " Class" & "(" & strResult & ")"
                
                If strIntBase = "INtIgE" Or strIntBase = "FDtIgE" Then
                    If IsNumeric(strIntResult) Then
                        If strIntResult > 100 Then
                            strIntResult = ">100"
                            strClass = "P" '"P ����" '������
                            strTIgE = "P"
                        Else ' strIntResult <= 100 Then
                            strIntResult = "=<100"
                            strClass = "N" '"N ����" '����
                            strTIgE = "N"
                        End If
                    Else
                        If Trim(strIntResult) = "> 100" Then
                            strIntResult = ">100"
                            strClass = "P" '"P ����" '������
                            strTIgE = "P"
                        ElseIf Trim(strIntResult) = "�� 100" Then
                            strIntResult = "��100"
                            strClass = "N" '"N ����"
                            strTIgE = "N"
                        Else
                            strIntResult = "��100"
                            strClass = "N"
                            strTIgE = "N"
                        End If
                    End If
                Else
                    If strIntResult > strSIgE Then
                        strSIgE = strIntResult
                    End If
                    
                    If IsNumeric(strIntResult) Then
                        If strIntResult < 0.35 Then
                            'strClass = "����"
                            strClass = "0"
                            strClassA = "0"
                            strIntResult = "0.00"
                        ElseIf strIntResult >= 0.35 And strIntResult < 0.7 Then
                            'strClass = "����"
                            strClass = "1"
                            strClassA = "1"
                            strIntResult = Format(strIntResult, "#0.#0")
                        ElseIf strIntResult >= 0.7 And strIntResult < 3.5 Then
                            'strClass = "�߰�"
                            strClass = "2"
                            strClassA = "2"
                            strIntResult = Format(strIntResult, "#0.#0")
                        ElseIf strIntResult >= 3.5 And strIntResult < 17.5 Then
                            'strClass = "�߰�/����"
                            strClass = "3"
                            strClassA = "3"
                            strIntResult = Format(strIntResult, "#0.#0")
                        ElseIf strIntResult >= 17.5 And strIntResult < 50 Then
                            'strClass = "����"
                            strClass = "4"
                            strClassA = "4"
                            strIntResult = Format(strIntResult, "#0.#0")
                        ElseIf strIntResult >= 50 And strIntResult < 100 Then
                            'strClass = "�ſ� ����"
                            strClass = "5"
                            strClassA = "5"
                            strIntResult = Format(strIntResult, "#0.#0")
                        ElseIf strIntResult >= 100 Then
                            'strClass = "���� ����"
                            strClass = "6"
                            strClassA = "6"
                            strIntResult = ">=100"
                        End If
                    Else
                         strClass = mGetP(strTemp, 2, "|")
                         strClassA = mGetP(strTemp, 2, "|")
                    End If
                    
                    If strClassA > strMaxClass Then
                        strMaxClass = strClassA
                    End If
                    
                End If
                If strIntResult = "0." Then strIntResult = "0.00"

                'strIntResult = strIntResult '& " :" & strClass
                strResult = strIntResult
                
                
                '-- ������� Class���� ����
                strResult = strClass & " Class" & "(" & strResult & ")"

                If objIntNms.ExistIntBase(strIntBase) Then
                    Call objIntInfo.IntResults.Add(strIntBase, objIntNms.GetIntNm(strIntBase), strResult, strResult, strClass)
                    mIntLib.State = "R"
                End If
            Next
            
            '-- ��ǥ�ڵ忡 �Ұ�����
            strIntBase = Mid(strIntBase, 1, 2)
            
            If strTIgE = "N" Then
                If strMaxClass < 2 Then
                    strIntResult = "����"
                    strResult = "����"
                    strClass = ""
                    'strClass = strClass & "|" & ALG_CMNT1
                ElseIf strMaxClass >= 2 Then
                    strIntResult = "��÷����"
                    strResult = "��÷����"
                    strClass = "*"
                    'strClass = strClass & "|" & ALG_CMNT21
                End If
            ElseIf strTIgE = "P" Then
                If strMaxClass < 2 Then
                    strIntResult = "��÷����"
                    strResult = "��÷����"
                    strClass = "*"
                    'strClass = strClass & "|" & ALG_CMNT3
                ElseIf strMaxClass >= 2 Then
                    strIntResult = "��÷����"
                    strResult = "��÷����"
                    strClass = "*"
                    'strClass = strClass & "|" & ALG_CMNT4
                End If
            End If
            
            If objIntNms.ExistIntBase(strIntBase) Then
                Call objIntInfo.IntResults.Add(strIntBase, objIntNms.GetIntNm(strIntBase), strIntResult, strResult, strClass)
                mIntLib.State = "R"
            End If

            '## DB�� �������
            If mIntLib.State = "R" Then
                Call SaveServer(objIntInfo, "IN")
                Set objIntInfo = Nothing
                mIntLib.State = ""
            End If
            
            intRow = intRow + 1
            AdoRS.MoveNext
        Loop
    End If
    
    Set AdoRS = Nothing

End Sub

'   Result List Recordset
Public Function Get_ResultList_FD() As ADODB.Recordset
    Dim strSql      As String

On Error GoTo ErrorTrap

             strSql = "SELECT * "
    strSql = strSql & "  FROM UNITS "
    'strSql = strSql & " WHERE EXAMDATE = '" & Format(dtpResult.Value, "yyyy-mm-dd") & "' "
    strSql = strSql & " WHERE EXAMDATE >= '" & Format(dtpFrDate.Value, "yyyy-mm-dd") & "' "
    strSql = strSql & "   AND EXAMDATE <= '" & Format(dtpToDate.Value, "yyyy-mm-dd") & "' "
    strSql = strSql & "   AND PANEL IN ('Food','Standard') "
    strSql = strSql & " ORDER BY EXAMDATE,PATIENTID, PANEL "
    
    Set AdoRS = New ADODB.Recordset
    If Get_Recordset(AdoCn, strSql, AdoRS, "") Then
        Set Get_ResultList_FD = AdoRS
    Else
        Set Get_ResultList_FD = Nothing
    End If
    
    Set AdoRS = Nothing

Exit Function

ErrorTrap:
    Set AdoRS = Nothing

End Function


Private Sub Get_SearchListFD()
    Dim intRow      As Long
    Dim strAge      As String
    Dim strTransDt  As String
    Dim strHMsg     As String
    Dim strDMsg     As String
    
    Dim objIntInfo   As clsIISIntInfo    '�������̽� ��ü���� Ŭ����
    Dim objIntNms    As clsIISIntNms     '��� �˻��׸� �÷��� Ŭ����
    Dim objBuffer    As clsIISBuffer     '���� Ŭ����
    
    Dim vWorkNo      As Variant  'Spread�� WorkNo
    Dim vBarNo       As Variant  'Spread�� ���ڵ��ȣ
    Dim strRcvBuf    As String   '������ Data
    Dim strType      As String   '������ Record Type
    Dim strBarNo     As String   '������ BarNO
    Dim strWorkNo    As String   '������ WorkNo
    Dim strIntResult As String   '������ �˻���
    Dim strIntBase   As String   '������ ������ �˻��
    Dim strResult    As String   'LIS �˻���
    Dim strClass     As String   'Class �������
    Dim strClassA    As String   'Class �������
    Dim strMaxClass  As String   'Class �������
    
    Dim strTemp      As String
    Dim i            As Long
    Dim blnSameBar   As Boolean
    Dim intCnt       As Integer
    Dim strTIgE      As String   'Total IgE
    Dim strSIgE      As String   'Ŭ������ ���� ���� ��
    Dim strRemark    As String
    Dim SIgE         As String
    Dim strPrevBarNo As String
    
    Dim Y, Y1, Y2, Y3, X1, X2
    
    Set objIntNms = mIntLib.IntNms
    
    On Error Resume Next
    
    Set AdoRS = Get_ResultList_FD
    
    If Not AdoRS.BOF Then
        intRow = 1: strTransDt = ""
        Do Until AdoRS.EOF
            strBarNo = AdoRS.Fields("PATIENTID").Value
            Set objIntInfo = New clsIISIntInfo
            With objIntInfo
                'If UCase(Mid(AdoRS.Fields("PANEL").Value, 1, 2)) = "IN" Then
                '    .SpcPos = "IN"
                'ElseIf UCase(Mid(AdoRS.Fields("PANEL").Value, 1, 2)) = "FO" Then
                    .SpcPos = "FD"
                'Else
                '    .SpcPos = strIntBase
                'End If
                '.OrgBarNo = strBarNo
                '.BarNo = Mid(strBarNo, 1, 10)
                .BarNo = strBarNo
            End With
            
            Call GetOrder(strBarNo)
            
            'If strBarNo = "170012224203" Then Stop
            
            'If strPrevBarNo <> "" And strPrevBarNo <> strBarNo Then
                strSIgE = "0"
                strMaxClass = "0"
                strClassA = "0"
                strIntBase = ""
           ' End If
            
            For intCnt = 1 To 32
                If intCnt = 1 Then
                    'If UCase(Trim(AdoRS.Fields("PANEL").Value & "")) = "FOOD" Then
                        strIntBase = "FD"
                    'ElseIf UCase(Trim(AdoRS.Fields("PANEL").Value & "")) = "INHALANT" Then
                    '    strIntBase = "IN"
                    'ElseIf UCase(Trim(AdoRS.Fields("PANEL").Value & "")) = "ATOPY" Then
                    '    strIntBase = "AT"
                    'Else
                    '    strIntBase = ""
                    'End If
                End If
                
                strIntResult = ""
                strTemp = AdoRS.Fields("BANDVAL" & intCnt) & ""
                If strTemp <> "" Then
                    strIntBase = objIntInfo.SpcPos & mGetP(strTemp, 3, "|")
                    strIntResult = mGetP(strTemp, 1, "|")
                End If

                If strIntBase = "INtIgE" Or strIntBase = "FDtIgE" Then
'                    If strIntResult > 100 Then
'                        strIntResult = ">100"
'                        strClass = "P" '"P ����" '������
'                        strTIgE = "P"
'                    Else ' strIntResult <= 100 Then
'                        strIntResult = "=<100"
'                        strClass = "N" '"N ����" '����
'                        strTIgE = "N"
'                    End If
                    If IsNumeric(strIntResult) Then
                        If strIntResult > 100 Then
                            strIntResult = ">100"
                            strClass = "P" '"P ����" '������
                            strTIgE = "P"
                        Else ' strIntResult <= 100 Then
                            strIntResult = "=<100"
                            strClass = "N" '"N ����" '����
                            strTIgE = "N"
                        End If
                    Else
                        If Trim(strIntResult) = "> 100" Then
                            strIntResult = ">100"
                            strClass = "P" '"P ����" '������
                            strTIgE = "P"
                        ElseIf Trim(strIntResult) = "�� 100" Then
                            strIntResult = "��100"
                            strClass = "N" '"N ����"
                            strTIgE = "N"
                        Else
                            strIntResult = "��100"
                            strClass = "N"
                            strTIgE = "N"
                        End If
                    End If
                
                Else
                    If strIntResult > strSIgE Then
                        strSIgE = strIntResult
                    End If
                    
                    If IsNumeric(strIntResult) Then
                        If strIntResult < 0.35 Then
                            'strClass = "����"
                            strClass = "0"
                            strClassA = "0"
                            strIntResult = "0.00"
                        ElseIf strIntResult >= 0.35 And strIntResult < 0.7 Then
                            'strClass = "����"
                            strClass = "1"
                            strClassA = "1"
                            strIntResult = Format(strIntResult, "#0.#0")
                        ElseIf strIntResult >= 0.7 And strIntResult < 3.5 Then
                            'strClass = "�߰�"
                            strClass = "2"
                            strClassA = "2"
                            strIntResult = Format(strIntResult, "#0.#0")
                        ElseIf strIntResult >= 3.5 And strIntResult < 17.5 Then
                            'strClass = "�߰�/����"
                            strClass = "3"
                            strClassA = "3"
                            strIntResult = Format(strIntResult, "#0.#0")
                        ElseIf strIntResult >= 17.5 And strIntResult < 50 Then
                            'strClass = "����"
                            strClass = "4"
                            strClassA = "4"
                            strIntResult = Format(strIntResult, "#0.#0")
                        ElseIf strIntResult >= 50 And strIntResult < 100 Then
                            'strClass = "�ſ� ����"
                            strClass = "5"
                            strClassA = "5"
                            strIntResult = Format(strIntResult, "#0.#0")
                        ElseIf strIntResult >= 100 Then
                            'strClass = "���� ����"
                            strClass = "6"
                            strClassA = "6"
                            strIntResult = ">=100"
                        End If
                    Else
                        strClass = mGetP(strTemp, 2, "|")
                        strClassA = mGetP(strTemp, 2, "|")
                    End If
                    If strClassA > strMaxClass Then
                        strMaxClass = strClassA
                    End If
                    
                End If
                If strIntResult = "0." Then strIntResult = "0.00"

                'strIntResult = strIntResult '& " :" & strClass
                strResult = strIntResult
                
                strResult = strClass & " Class" & "(" & strResult & ")"
                
                If objIntNms.ExistIntBase(strIntBase) Then
                    Call objIntInfo.IntResults.Add(strIntBase, objIntNms.GetIntNm(strIntBase), strResult, strResult, strClass)
                    mIntLib.State = "R"
                End If
            Next
            
            '-- ��ǥ�ڵ忡 �Ұ�����
            strIntBase = Mid(strIntBase, 1, 2)
            
            If strTIgE = "N" Then
                If strMaxClass < 2 Then
                    strIntResult = "����"
                    strResult = "����"
                    strClass = ""
                    strClass = strClass & "|" & ALG_CMNT1
                ElseIf strMaxClass >= 2 Then
                    strIntResult = "��÷����"
                    strResult = "��÷����"
                    strClass = "*"
                    strClass = strClass & "|" & ALG_CMNT2
                End If
            ElseIf strTIgE = "P" Then
                If strMaxClass < 2 Then
                    strIntResult = "��÷����"
                    strResult = "��÷����"
                    strClass = "*"
                    strClass = strClass & "|" & ALG_CMNT3
                ElseIf strMaxClass >= 2 Then
                    strIntResult = "��÷����"
                    strResult = "��÷����"
                    strClass = "*"
                    strClass = strClass & "|" & ALG_CMNT4
                End If
            End If
            
            strResult = strClass & " Class" & "(" & strResult & ")"
            
            If objIntNms.ExistIntBase(strIntBase) Then
                Call objIntInfo.IntResults.Add(strIntBase, objIntNms.GetIntNm(strIntBase), strIntResult, strResult, strClass)
                mIntLib.State = "R"
            End If

            '## DB�� �������
            If mIntLib.State = "R" Then
                Call SaveServer(objIntInfo, "FD")
                Set objIntInfo = Nothing
                mIntLib.State = ""
            End If
            
            intRow = intRow + 1
            AdoRS.MoveNext
        Loop
    End If
    
    Set AdoRS = Nothing

End Sub

Private Sub cmdMakeWS_Click()

    Dim objAccInfo As clsIISAccInfo     '�������� Ŭ����
    Dim objResult   As clsIISResult     '������� Ŭ����
    Dim objQCResult As clsIISQCResult   'QC������� Ŭ����
    Dim mLogOn      As clsIISLogOn
    Dim strAlloFile As String
    Dim lngFIleNum  As Long
    Dim strInFo     As String
    Dim strOldInFo  As String
    Dim iCnt        As Integer
    Dim varTmp      As Variant
    Dim OrderPath   As String
    Dim i           As Integer
    Dim strOrgBarNo    As String
    Dim strBarNo    As String
    Dim strPtID     As String
    Dim strDeptNm   As String
    Dim strPtNm     As String
'    Dim strBarNo    As String
    Dim varBuffer   As Variant
    Dim intCnt As Integer
    Dim strBuf      As String
    Dim Rs          As ADODB.Recordset
    Dim Rs1 As New Recordset
    Dim strKey As String
    Dim strTemp As String
    Dim j As Integer
    
    varBuffer = Split(strBuffer, vbCr)
    
    'OrderPath = OrderPath & Format(Now, "YYMMDDHHMMSS") & ".patlist"
    
    '.FileName = OrderPath & "PatList " & Format(Now, "yyyy-mm-dd") & ".patlist"
    '.FileName = OrderPath & Format(Now, "YYMMDDHHMMSS") & ".patlist"
    AlloFile.FileName = lblFilenm.Caption
    AlloFile.CancelError = True
    j = 0
    For intCnt = 0 To UBound(varBuffer)
        strBuf = varBuffer(intCnt)
        If Mid(strBuf, 1, 2) = "--" Or Mid(strBuf, 1, 2) = "NA" Or Mid(strBuf, 1, 2) = "TY" Or Mid(strBuf, 1, 2) = "ID" Then
    
        Else
            j = j + 1
            
            strOrgBarNo = Mid(strBuf, 1, 12)
            
            If Len(strBuf) > 13 Then  'Or strBuf = ""
                Close #lngFIleNum
                strBuffer = ""
                Exit Sub
            End If
            
            If j = 1 Then
                If Len(Dir(AlloFile.FileName)) Then
                     Close #lngFIleNum
                     Kill AlloFile.FileName
                End If
                lngFIleNum = FreeFile
                        
                Open AlloFile.FileName For Append As #lngFIleNum
            
                '-- �������ϸ� ���� : PatList 2012-01-16.patlist
                OrderPath = GetAlloConfig("OrderPath")
            
                Print #lngFIleNum, "-----------------------------------------------------------------------------"
                Print #lngFIleNum, "NAME: AdvanSure"
                Print #lngFIleNum, "TYPE: Patient List Files"
                Print #lngFIleNum, "-----------------------------------------------------------------------------"
                Print #lngFIleNum, "ID " & vbTab & " NAME " & vbTab & " PANEL " & vbTab & " A " & vbTab & " B " & vbTab & " AGE " & vbTab & " GENDER " & vbTab & " ADDR " & vbTab & _
                                   " CONTACT " & vbTab & " BIRTH " & vbTab & " RRN " & vbTab & " CLIENT " & vbTab & " INSPECTOR " & vbTab & " HOSPITAL " & vbTab & " EXAMDATE"
            End If
                        
            'strOrgBarNo = Mid(strBuf, 1, 12)
            Set objAccInfo = New clsIISAccInfo
            
            'Set Rs = objAccInfo.GetTargetSpcsBar_Allergy(mEqpCd, Mid(strOrgBarNo, 1, 10), "03")
            
            Call GetOrder(strOrgBarNo)

      
        End If
    Next
    
    
    With AlloFile

        For iCnt = 1 To tblReady.MaxRows
            tblReady.GetText 3, iCnt, varTmp: strPtID = varTmp
            tblReady.GetText 4, iCnt, varTmp: strDeptNm = varTmp
            tblReady.GetText 5, iCnt, varTmp: strPtNm = varTmp
            tblReady.GetText 2, iCnt, varTmp: strBarNo = varTmp
            If varTmp = "" Then Exit For
                
            tblReady.GetText 1, iCnt, varTmp
            'varTmp = Split(varTmp, "/")
            'For i = 0 To UBound(varTmp)
                '-- ȣ���
                If varTmp = "IN" Then
                    strInFo = "1"
                    Print #lngFIleNum, strBarNo & vbTab & strPtNm & vbTab & strInFo & vbTab & "1" & vbTab & "1" & vbTab & "" & vbTab & "" & vbTab & vbTab & vbTab & vbTab & "" & vbTab & strDeptNm & vbTab & "" & vbTab & vbTab & Format(Now, "yyyy-mm-dd")
                '-- ����
                ElseIf varTmp = "FD" Then
                    strInFo = "2"
                    Print #lngFIleNum, strBarNo & vbTab & strPtNm & vbTab & strInFo & vbTab & "1" & vbTab & "1" & vbTab & "" & vbTab & "" & vbTab & vbTab & vbTab & vbTab & "" & vbTab & strDeptNm & vbTab & "" & vbTab & vbTab & Format(Now, "yyyy-mm-dd")
                '-- ������
                'ElseIf varTmp(i) = "AT" Then
                '    strInFo = "3"
                '    Print #lngFIleNum, strBarNo & vbTab & strPtNm & vbTab & strInFo & vbTab & "1" & vbTab & "0" & vbTab & "" & vbTab & "" & vbTab & vbTab & vbTab & vbTab & "" & vbTab & strDeptNm & vbTab & "" & vbTab & vbTab & Format(Now, "yyyy-mm-dd")
                End If
        
                'If strOldInFo <> strInFo Then
                '    Print #lngFIleNum, strBarNo & vbTab & strPtNm & vbTab & strInFo & vbTab & "1" & vbTab & "1" & vbTab & "" & vbTab & "" & vbTab & vbTab & vbTab & vbTab & "" & vbTab & strDeptNm & vbTab & "" & vbTab & vbTab & Format(Now, "yyyy-mm-dd")
                'End If
        
                'strOldInFo = strInFo
                
                tblReady.SetText 1, iCnt, varTmp & " ��"
                
            'Next
        Next
        
        'MsgBox "���� ���� ���� �Ϸ�", vbOKOnly + vbInformation, Me.Caption
        
        
    End With
    
    Close #lngFIleNum
    strBuffer = ""
    
End Sub

Private Function GetMEDIWISSConfig(ByVal strConfigNm As String) As String

Dim strFileName As String
Dim strReturnedString As String

    strFileName = App.Path & "\MEDIWISS.ini"
    
    strReturnedString = String(1024, " ")
    GetPrivateProfileString "MEDIWISS", strConfigNm, "", strReturnedString, Len(strReturnedString), strFileName
    strReturnedString = Trim(Replace(strReturnedString, Chr(0), Chr(32), 1, -1, vbBinaryCompare))
    GetMEDIWISSConfig = strReturnedString
    
End Function

Private Sub cmdPrint_Click()
    Dim strQcFg     As String   'QC����
    Dim strResult   As String   'LIS ���
    Dim strTemp     As String
    Dim i           As Long
    
    Dim strPrtData(100) As String
    Dim intRow      As Integer
    Dim intCol      As Integer
    Dim intCnt      As Integer
    Dim strPanel    As String
    Dim strValue    As String
    Dim strClass    As String
    Dim iDestRow    As Integer
    
    Erase strPrtData
    intCnt = 0
    
    With vasINPrint
        Call .SetText(3, 2, ""): Call .SetText(7, 2, ""): Call .SetText(11, 2, "")
        Call .SetText(3, 3, ""): Call .SetText(7, 3, ""): Call .SetText(11, 3, "")
        Call .SetText(3, 4, ""): Call .SetText(7, 4, ""): Call .SetText(11, 4, "")
        For i = 6 To 67
            Call .SetText(4, i, ""): Call .SetText(7, i, ""): Call .SetText(10, i, ""): Call .SetText(12, i, "")
        Next
    End With
    
    With vasFDPrint
        Call .SetText(3, 2, ""): Call .SetText(7, 2, ""): Call .SetText(11, 2, "")
        Call .SetText(3, 3, ""): Call .SetText(7, 3, ""): Call .SetText(11, 3, "")
        Call .SetText(3, 4, ""): Call .SetText(7, 4, ""): Call .SetText(11, 4, "")
        For i = 6 To 67
            Call .SetText(4, i, ""): Call .SetText(7, i, ""): Call .SetText(10, i, ""): Call .SetText(12, i, "")
        Next
    End With
    
    
    If lblPtId.Caption <> "" And lblPnlNm.Caption <> "" Then
        strPrtData(0) = lblPtId.Caption
        strPrtData(1) = lblName.Caption
        strPrtData(2) = Format(Now, "yyyy-mm-dd")
        strPrtData(3) = mGetP(lblSexAge.Caption, 2, "/")
        'strPrtData(4) = IIf(mGetP(lblSexAge.Caption, 1, "/") = "M", "����", "����")
        strPrtData(4) = IIf(InStr(lblSexAge.Caption, "M") > 0, "����", "����")
        strPrtData(5) = lblPnlNm.Caption
        strPrtData(6) = lblDeptNm.Caption
        
        'strPrtData(7) = "�ֿ�ȯ"
        '2020-12-31 ������ ����� ����
        'strPrtData(8) = "���μ�" '"������"
        
        strPrtData(7) = GetAlloConfig("Pathologist")    ' "�ֿ�ȯ"
        strPrtData(8) = GetAlloConfig("Doctor")         ' "���μ�"
        
        strPanel = IIf(Trim(lblPnlNm.Caption) = "INHALANT", "IN", "FD")
        strPanel = IIf(Trim(lblPnlNm.Caption) = "FOOD", "FD", "IN")
        
        Call vasINPrint.SetText(3, 71, Format(Now, "yyyy-mm-dd"))
        Call vasFDPrint.SetText(3, 71, Format(Now, "yyyy-mm-dd"))
        
        For intRow = tblComplete.ActiveRow - 1 To tblComplete.DataRowCnt
            tblComplete.Row = intRow
            tblComplete.Col = TCompleteEnum.ccPtId
            If Trim(tblComplete.Text) = Trim(lblPtId.Caption) Then
                tblComplete.Col = TCompleteEnum.ccNo
'                If Trim(tblComplete.Text) = strPanel Then ' 2017.05.10�½�ȣ
                    intCnt = intCnt + 1
                    If strPanel = "IN" Then
                        With vasINPrint
                            If intCnt = 1 Then
                                Call .SetText(3, 2, strPrtData(0)): Call .SetText(7, 2, strPrtData(3)): Call .SetText(11, 2, strPrtData(6))
                                Call .SetText(3, 3, strPrtData(1)): Call .SetText(7, 3, strPrtData(4)): Call .SetText(11, 3, strPrtData(7))
                                Call .SetText(3, 4, strPrtData(2)): Call .SetText(7, 4, lblPnlNm.Caption): Call .SetText(11, 4, strPrtData(8))
                            End If
                            For i = TCompleteEnum.ccResult To tblComplete.DataColCnt
                                tblComplete.Col = i:   strTemp = tblComplete.Text
                                Debug.Print UCase(mGetP(strTemp, TResultEnum.ccIntBase, DIV))
                                Select Case UCase(mGetP(strTemp, TResultEnum.ccIntBase, DIV))
                                    ' 1KO
'                                    Case "IN|IgE":      iDestRow = 6
'                                    Case "IN|F14":      iDestRow = 7
'                                    Case "IN|F2":       iDestRow = 8
'                                    Case "IN|F1":       iDestRow = 9
'                                    Case "IN|F23":      iDestRow = 10
'                                    Case "IN|F24":      iDestRow = 11
'                                    Case "IN|F95":      iDestRow = 12
'                                    Case "IN|T35":      iDestRow = 13
'                                    Case "IN|T15":      iDestRow = 14
'                                    Case "IN|T2_T3":    iDestRow = 15
'                                    Case "IN|T12":      iDestRow = 16
'                                    Case "IN|F17":      iDestRow = 17
'                                    Case "IN|T17":      iDestRow = 18
'                                    Case "IN|T7":       iDestRow = 19
'                                    Case "IN|T14":      iDestRow = 20
'                                    Case "IN|T1_T11":   iDestRow = 21
'                                    Case "IN|G2":       iDestRow = 22
'                                    Case "IN|G3":       iDestRow = 23
'                                    Case "IN|G6":       iDestRow = 24
'                                    Case "IN|G12":      iDestRow = 25
'                                    Case "IN|w12":      iDestRow = 26
'                                    Case "IN|I1":       iDestRow = 27
'                                    Case "IN|D72":      iDestRow = 28
'                                    Case "IN|G9":       iDestRow = 29
'                                    Case "IN|T225":     iDestRow = 30
'                                    Case "IN|F244":     iDestRow = 31
'                                    Case "IN|CCDx":     iDestRow = 32
'                                    Case "IN|F84":      iDestRow = 33
'                                    Case "IN|F313":     iDestRow = 34
'                                    Case "IN|I81":      iDestRow = 35
'                                    '2KO
'                                    Case "IN|W14":      iDestRow = 36
'                                    Case "IN|W11":      iDestRow = 37
'                                    Case "IN|W8":       iDestRow = 38
'                                    Case "IN|W6":       iDestRow = 39
'                                    Case "IN|W2":       iDestRow = 40
'                                    Case "IN|M6":       iDestRow = 41
'                                    Case "IN|M3":       iDestRow = 42
'                                    Case "IN|M2":       iDestRow = 43
'                                    Case "IN|M1":       iDestRow = 44
'                                    Case "IN|E1":       iDestRow = 45
'                                    Case "IN|E5":       iDestRow = 46
'                                    Case "IN|I6":       iDestRow = 47
'                                    Case "IN|HX":       iDestRow = 48
'                                    Case "IN|D2":       iDestRow = 49
'                                    Case "IN|D1":       iDestRow = 50
'                                    Case "IN|G1":       iDestRow = 51
'                                    Case "IN|G7":       iDestRow = 52
'                                    Case "IN|T16":      iDestRow = 53
'                                    Case "IN|W7":       iDestRow = 54
'                                    Case "IN|W22sc":    iDestRow = 55
'                                    Case "IN|F206":     iDestRow = 56
'                                    Case "IN|F299":     iDestRow = 57
'                                    Case "IN|Fx21":     iDestRow = 58
'                                    Case "IN|F92":      iDestRow = 59
'                                    Case "IN|F37":      iDestRow = 60
'                                    Case "IN|F91":      iDestRow = 61
'                                    Case "IN|F35":      iDestRow = 62
'                                    Case "IN|F93":      iDestRow = 63
'                                    Case "IN|K82":      iDestRow = 64
'                                    Case "IN|E81":      iDestRow = 65
                                    
                                    Case "INTIGE":      iDestRow = 6
                                    Case "IND1":        iDestRow = 7
                                    Case "IND2":        iDestRow = 8
                                    Case "IND72":       iDestRow = 9
                                    Case "INE1":        iDestRow = 10
                                    Case "INE2":        iDestRow = 11
                                    Case "INF1":        iDestRow = 12
                                    Case "INF2":        iDestRow = 13
                                    Case "INF8":        iDestRow = 14
                                    Case "INF10":       iDestRow = 15
                                    Case "INF14":       iDestRow = 16
                                    Case "INF23":       iDestRow = 17
                                    Case "INF24":       iDestRow = 18
                                    Case "INF35":       iDestRow = 19
                                    Case "INF49":       iDestRow = 20
                                    Case "INF93":       iDestRow = 21
                                    Case "INF95":       iDestRow = 22
                                    Case "INF206":      iDestRow = 23
                                    Case "INO214":      iDestRow = 24
                                    Case "ING12":       iDestRow = 25
                                    Case "INH1":        iDestRow = 26
                                    Case "INI6":        iDestRow = 27
                                    Case "INM2":        iDestRow = 28
                                    Case "INM3":        iDestRow = 29
                                    Case "INM6":        iDestRow = 30
                                    Case "INT2":        iDestRow = 31
                                    Case "INT3":        iDestRow = 32
                                    Case "INT7":        iDestRow = 33
                                    Case "INW2":        iDestRow = 34
                                    Case "INW6":        iDestRow = 35
                                    Case "INW22":       iDestRow = 36
                                    Case "IND70":       iDestRow = 37
                                    Case "INE3":        iDestRow = 38
                                    Case "INE6":        iDestRow = 39
                                    Case "INE81":       iDestRow = 40
                                    Case "INE82":       iDestRow = 41
                                    Case "INE84":       iDestRow = 42
                                    Case "INT4":        iDestRow = 43
                                    Case "ING1":        iDestRow = 44
                                    Case "ING2":        iDestRow = 45
                                    Case "ING3":        iDestRow = 46
                                    Case "ING6":        iDestRow = 47
                                    Case "ING7":        iDestRow = 48
                                    Case "ING9":        iDestRow = 49
                                    Case "INI1":        iDestRow = 50
                                    Case "INI3":        iDestRow = 51
                                    Case "INK82":       iDestRow = 52
                                    Case "INM1":        iDestRow = 53
                                    Case "INT11":       iDestRow = 54
                                    Case "INT12":       iDestRow = 55
                                    Case "INT14":       iDestRow = 56
                                    Case "INT15":       iDestRow = 57
                                    Case "INT16":       iDestRow = 58
                                    Case "INT17":       iDestRow = 59
                                    Case "INT19":       iDestRow = 60
                                    Case "INT222":      iDestRow = 61
                                    Case "INW7":        iDestRow = 62
                                    Case "INW8":        iDestRow = 63
                                    Case "INW9":        iDestRow = 64
                                    Case "INW11":       iDestRow = 65
                                    Case "INW12":       iDestRow = 66
                                    Case "INW14":       iDestRow = 67


                                    
                                    
                                    
                                End Select
                                
                                strValue = ""
                                strClass = ""
                                If Trim(strTemp) <> "" Then
                                    strValue = mGetP(strTemp, TResultEnum.ccLISResult, DIV)
                                    strValue = Mid(strValue, InStr(strValue, "(") + 1)
                                    strValue = Mid(strValue, 1, Len(strValue) - 1)
                                    strClass = mGetP(strTemp, TResultEnum.ccLISResult, DIV)
                                    strClass = mGetP(strClass, 1, " ")
                                End If
                                Call .SetText(10, iDestRow, strValue)
                                Call .SetText(12, iDestRow, strClass)
                                If IsNumeric(strClass) Then
                                    If CCur(strClass) >= 2 Then
                                        Call .SetText(4, iDestRow, "*")
                                        Call .SetText(7, iDestRow, "*")
                                    End If
                                End If
                            Next i
                        End With
                    ElseIf strPanel = "FD" Then
                        With vasFDPrint
                                
                            If intCnt = 1 Then
                                Call .SetText(3, 2, strPrtData(0)): Call .SetText(7, 2, strPrtData(3)): Call .SetText(11, 2, strPrtData(6))
                                Call .SetText(3, 3, strPrtData(1)): Call .SetText(7, 3, strPrtData(4)): Call .SetText(11, 3, strPrtData(7))
                                Call .SetText(3, 4, strPrtData(2)): Call .SetText(7, 4, lblPnlNm.Caption): Call .SetText(11, 4, strPrtData(8))
                            End If
                            For i = TCompleteEnum.ccResult To tblComplete.DataColCnt
                                tblComplete.Col = i:   strTemp = tblComplete.Text
                                Select Case UCase(mGetP(strTemp, TResultEnum.ccIntBase, DIV))
                                    
                                
                                    Case "FDTIGE":       iDestRow = 6
                                    Case "FDD1":        iDestRow = 7
                                    Case "FDD2":        iDestRow = 8
                                    Case "FDD72":       iDestRow = 9
                                    Case "FDE1":        iDestRow = 10
                                    Case "FDE2":        iDestRow = 11
                                    Case "FDF1":        iDestRow = 12
                                    Case "FDF2":        iDestRow = 13
                                    Case "FDF8":        iDestRow = 14
                                    Case "FDF10":       iDestRow = 15
                                    Case "FDF14":       iDestRow = 16
                                    Case "FDF23":       iDestRow = 17
                                    Case "FDF24":       iDestRow = 18
                                    Case "FDF35":       iDestRow = 19
                                    Case "FDF49":       iDestRow = 20
                                    Case "FDF93":       iDestRow = 21
                                    Case "FDF95":       iDestRow = 22
                                    Case "FDF206":      iDestRow = 23
                                    Case "FDO214":      iDestRow = 24
                                    Case "FDG12":       iDestRow = 25
                                    Case "FDH1":        iDestRow = 26
                                    Case "FDI6":        iDestRow = 27
                                    Case "FDM2":        iDestRow = 28
                                    Case "FDM3":        iDestRow = 29
                                    Case "FDM6":        iDestRow = 30
                                    Case "FDT2":        iDestRow = 31
                                    Case "FDT3":        iDestRow = 32
                                    Case "FDT7":        iDestRow = 33
                                    Case "FDW2":        iDestRow = 34
                                    Case "FDW6":        iDestRow = 35
                                    Case "FDW22":       iDestRow = 36
                                    Case "FDF26":       iDestRow = 37
                                    Case "FDF27":       iDestRow = 38
                                    Case "FDF81":       iDestRow = 39
                                    Case "FDF83":       iDestRow = 40
                                    Case "FDO211":      iDestRow = 41
                                    Case "FDF25":       iDestRow = 42
                                    Case "FDF84":       iDestRow = 43
                                    Case "FDF91":       iDestRow = 44
                                    Case "FDF92":       iDestRow = 45
                                    Case "FDFX":        iDestRow = 46
                                    Case "FDF13":       iDestRow = 47
                                    Case "FDF256":      iDestRow = 48
                                    Case "FDF299":      iDestRow = 49
                                    Case "FDF4":        iDestRow = 50
                                    Case "FDF6":        iDestRow = 51
                                    Case "FDF9":        iDestRow = 52
                                    Case "FDF11":       iDestRow = 53
                                    Case "FDF47":       iDestRow = 54
                                    Case "FDF48":       iDestRow = 55
                                    Case "FDF85":       iDestRow = 56
                                    Case "FDF244":      iDestRow = 57
                                    Case "FDF3":        iDestRow = 58
                                    Case "FDF37":       iDestRow = 59
                                    Case "FDF40":       iDestRow = 60
                                    Case "FDF41":       iDestRow = 61
                                    Case "FDF207":      iDestRow = 62
                                    Case "FDF258":      iDestRow = 63
                                    Case "FDF313":      iDestRow = 64
                                    Case "FDF45":       iDestRow = 65
                                    Case "FDF212":      iDestRow = 66
                                    Case "FDM5":        iDestRow = 67
                                    
                                End Select
                                
                                strValue = ""
                                strClass = ""
                                If Trim(strTemp) <> "" Then
                                    strValue = mGetP(strTemp, TResultEnum.ccLISResult, DIV)
                                    strValue = Mid(strValue, InStr(strValue, "(") + 1)
                                    strValue = Mid(strValue, 1, Len(strValue) - 1)
                                    strClass = mGetP(strTemp, TResultEnum.ccLISResult, DIV)
                                    strClass = mGetP(strClass, 1, " ")
                                End If
                                Call .SetText(10, iDestRow, strValue)
                                Call .SetText(12, iDestRow, strClass)
                                If IsNumeric(strClass) Then
                                    If CCur(strClass) >= 2 Then
                                        Call .SetText(4, iDestRow, "*")
                                        Call .SetText(7, iDestRow, "*")
                                    End If
                                End If
                            Next i
                        End With
                    
'                    End If
                End If
            End If
            If intCnt = 2 Then
                Exit For
            End If
        Next
        
        If intCnt = 2 Then
            If strPanel = "IN" Then
                vasINPrint.PrintOrientation = PrintOrientationPortrait '�������
                vasINPrint.PrintBorder = False
                vasINPrint.Action = ActionPrint
            ElseIf strPanel = "FD" Then
                vasFDPrint.PrintOrientation = PrintOrientationPortrait '�������
                vasINPrint.PrintBorder = False
                vasFDPrint.Action = ActionPrint  '13
            End If
        End If
    
    End If

End Sub


'Private Sub mnuSaveExL_Click()
'    Dim sFile As String, FileNamed As String
'    Dim irow As Long, icol As Long
'
'With AlloFile
'        .DialogTitle = "Save as Excel"
'        .FileName = ""
'        .CancelError = False
'
'        .Filter = "Text Files (*.xls)|*.xls"
'        .ShowSave
'        If Len(.FileName) = 0 Then
'            Exit Sub
'        End If
'        sFile = .FileName
'    End With
'    If InStr(sFile, ".xls") = 0 Then
'    sFile = sFile & ".xls"
'
'End If
'
'With tblexcel 'myExcelFile
'
'    FileNamed = sFile
'    .CreateFile FileNamed
'
'    .PrintGridLines = True
'
'    .SetFont "Arial", 10, xlsNoFormat              'font0
'    .SetFont "Arial", 10, xlsBold                  'font1
'    .SetFont "Arial", 10, xlsBold + xlsUnderline   'font2
'    .SetFont "Courier", 12, xlsItalic              'font3
'
'    For irow = 1 To GrdSheet.Rows - 1
'        GrdSheet.Row = irow
'        For icol = 1 To GrdSheet.Cols - 1
'            GrdSheet.Col = icol
'            .SetFont GrdSheet.CellFontName, 10, xlsNoFormat
'            .WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, irow, icol, GrdSheet.Text
'            If Formula(GrdSheet.Row, GrdSheet.Col) > "" Then
'                .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsHidden, irow, icol, Formula(GrdSheet.Row, GrdSheet.Col)
'            End If
'        Next
'    Next
'
'    .CloseFile
'    Close
'
'
'End With
'
'Close
'
'    MsgBox "Excel BIFF Spreadsheet created." & vbCrLf & "Filename: " & FileNamed, vbInformation + vbOKOnly, "Excel Class"
'End Sub

''Private Sub cmdExcel_Click()
''    Call Excel_Save
''End Sub
'Private Sub Excel_Save()
'
'    If tblResult.MaxRows = 1 Then Exit Sub
'    AlloFile.FileName = ""
'
'    AlloFile.Filter = "Excel(*.xls)|*.xls"
'    AlloFile.ShowSave
'
'    If AlloFile.FileName <> "" Then
'
'       Call Excel(AlloFile.FileName)
'    End If
'
'End Sub
'Private Sub Excel(File_Name As String)
'     Dim Ex_App As Object
''     dim Ex_Book  as
'
'    On Error GoTo Err
'
'     Screen.MousePointer = vbHourglass
'
'     Set Ex_App = CreateObject("Excel.Application")
''     Set Ex_Book = Ex_App.Workbooks.Add(1)
''     Set Ex_Sheet = Ex_Book.Worksheets(1)
''
''     Ex_App.ScreenUpdating = False
''     Ex_App.DisplayAlerts = False
'
'     Call EXCEL_DRAW
'
'
'     Ex_App.DisplayAlerts = True
'     Ex_App.ScreenUpdating = True
'
''     Ex_Book.SaveAs File_Name
'
''     Ex_Book.Close
'
'     MsgBox "Excel Complete!!"
'
'     Screen.MousePointer = vbDefault
'     Ex_App.Quit
''     Set Ex_Sheet = Nothing
''     Set Ex_Book = Nothing
''     Set Ex_App = Nothing
'
'     Exit Sub
'
'Err:
'
'    MsgBox "Excel Cancel!!"
'
'    Screen.MousePointer = vbDefault
''    Ex_App.DisplayAlerts = False
''    Ex_Book.Close
''    Ex_App.Quit
''
''     Set Ex_Sheet = Nothing
''     Set Ex_Book = Nothing
''     Set Ex_App = Nothing
'End Sub

'Private Sub EXCEL_DRAW()
'
'    Dim Title As Variant
'    Dim location As Variant
'    Dim COL_Location() As Variant
'    Dim i As Integer
'
'
'
'
'    With Ex_App
'
'        Title = Array("KIDO IP LIST", _
'                       "����", "����", "USE IP", "USE ID", "USE NAME", "PASSWORD", "USE MAIL", "GROUP ID", "GROUP PWD")
'
'         location = Array("A1:I2", _
'                          "A3:A3", "B3:B3", "C3:C3", "D3:D3", "E3:E3", _
'                          "F3:F3", "G3:G3", "H3:H3", "I3:I3")
'
'
'         COL_Location = Array("A", _
'                                "B", "C", "D", "E", "F", "G", "H", "I")
'
'
'         '�� ������
'         .Range("A1").ColumnWidth = 18
'         .Range("B1").ColumnWidth = 18
'         .Range("C1").ColumnWidth = 18
'         .Range("D1").ColumnWidth = 18
'         .Range("E1").ColumnWidth = 18
'         .Range("F1").ColumnWidth = 18
'         .Range("G1").ColumnWidth = 18
'         .Range("H1").ColumnWidth = 18
'         .Range("I1").ColumnWidth = 18
'
'          ' title �� �� ��ġ��
'         For i = 0 To UBound(location)
'            Call Cell_Draw(location(i), Title, i)
'         Next i
'
'        With .Range("A2:i2")
'           .Interior.Color = vbYellow
'        End With
'
'         'DATA ���̱� (Reason�� �����)
'         For i = 1 To grdM.Rows - 1
'
'            ' �� ���� <����>
'            '''''''''''''''''''''''''''''''''''''''''''''''''''''
'            If i = 1 Then
'                First_Value = ""
'                Last_Value = ""
'
'                first_row = 0
'                last_row = 0
'
'                First_Value = tblResult.TextMatrix(i, tblResult.ColIndex("COM")) ''�ش� LINE
'                f = i
'                L = 0
'            End If
'
'            If i > 0 Then
'
'                If First_Value = tblResult.TextMatrix(i, tblResult.ColIndex("COM")) Then
'
'                    first_row = f
'                    L = L + 1
'                    last_row = L
'
'
'                    '�ʱ�ȭ
'                    f = i
'                    L = i
'                    First_Value = tblResult.TextMatrix(f, tblResult.ColIndex("COM"))
'
'                End If
'            End If
'
'
'          .Cells(i + 3, "A") = tblResult.TextMatrix(i, tblResult.ColIndex("COM"))
'          .Cells(i + 3, "B") = tblResult.TextMatrix(i, tblResult.ColIndex("KIND"))
'          .Cells(i + 3, "C") = tblResult.TextMatrix(i, tblResult.ColIndex("USEIP"))
'          .Cells(i + 3, "D") = tblResult.TextMatrix(i, tblResult.ColIndex("USEID"))
'          .Cells(i + 3, "E") = tblResult.TextMatrix(i, tblResult.ColIndex("USENAME"))
'          .Cells(i + 3, "F") = tblResult.TextMatrix(i, tblResult.ColIndex("PASSWORD"))
'          .Cells(i + 3, "G") = tblResult.TextMatrix(i, tblResult.ColIndex("EMAIL"))
'          .Cells(i + 3, "H") = tblResult.TextMatrix(i, tblResult.ColIndex("GROUPID"))
'          .Cells(i + 3, "I") = tblResult.TextMatrix(i, tblResult.ColIndex("GROUPPWD"))
'         Next
'
'
'        '�⺻ FONT, FONTSIZE ���ϱ�
'        With .Range("A4:i" & CStr(i + 3))
'            .Font.Name = "����ü"
'            .Font.Size = "9"
'        End With
'
'        '�⺻�����ϱ�
'        With .Range("A4:i" & CStr(i + 3))
'            .VerticalAlignment = xlCenter
'            .HorizontalAlignment = xlCenter
'        End With
'
''        .Range("A2:i2").Activate
''        .ActiveWindow.FreezePanes = True
'
'        End With
'      End Sub
'      '���� ä����
'Private Sub Cell_Draw(Location_Name As Variant, HeadName As Variant, ARR_xl As Integer)
'    With Ex_App
'
'        With .Range(Location_Name)
'              .Select
'              .VerticalAlignment = xlCenter
'              .WrapText = False
'              .Orientation = 0
'              .AddIndent = True
'              .ShrinkToFit = True
'              .MergeCells = True
'              .Value = HeadName(ARR_xl)
'              Select Case ARR_xl
'              Case 0
'                    .HorizontalAlignment = xlCenter
'              Case 2 ' right
'                    .HorizontalAlignment = xlCenter
'
'         End With
'
'
'         With .Selection.Font
'               .Name = "����ü"
'               Select Case ARR_xl
'               Case 0
'                .Size = 20
'                .Underline = xlUnderlineStyleSingle  ' ���ڹؿ� ���ٱ߱�
'               Case Else
'
'                .Size = 8
'               End Select
'               .FontStyle = "Bold"
'
'         End With
'
'   End With
'End Sub

'Private Sub cmdSearch_Click()
'
''    Call GetOrderByAccNo
'
'    Dim Rs          As ADODB.Recordset
'    Dim objAccInfo  As clsIISAccInfo    '�������� Ŭ����
'    Dim strEqpCd    As String           '����ڵ�
'    Dim strFromDt   As String           'From Date
'    Dim strToDt     As String           'To Date
'    Dim strKey      As String           'Spread�� Ű(SpcYy+SpcNo)
'    Dim strSpcYy    As String           '���ڵ��ȣ(����)
'    Dim strSpcNo    As String           '���ڵ��ȣ(����)
'    Dim strTemp     As String
'    Dim vBarNo      As String
'
'    'strEqpCd = Trim$(txtEqpCd.Text)
'    strFromDt = Format$(dtpFrDate.Value, "YYYYMMDD")
'    strToDt = Format$(dtpToDate.Value, "YYYYMMDD")
'    'If strEqpCd = "" Then
'    '    MsgBox "��� �����ϼ���.", vbInformation, "����"
'    '    Exit Sub
'    'End If
'
'    Me.MousePointer = vbHourglass
'    Call mTblClear(tblReady)
'
'On Error GoTo Errors
'    Set objAccInfo = New clsIISAccInfo
'    Set Rs = objAccInfo.GetTargetSpcs(mEqpCd, strFromDt, strToDt)
'
'    If Not (Rs.BOF Or Rs.EOF) Then
'        With tblReady
'            tblReady.Visible = False
'            Do Until Rs.EOF
'                strSpcYy = Rs.Fields("SPCYY").Value
'                strSpcNo = Rs.Fields("SPCNO").Value
'
'                vBarNo = Format$(strSpcYy, String$(SPCYYLEN, "0")) & Format$(strSpcNo, String$(SPCNOLEN, "0"))
'                vBarNo = Format$(vBarNo, String$(SPCLEN, "#"))
'
'                Call GetOrder(vBarNo)
''''                strSpcYy = Rs.Fields("SPCYY").Value
''''                strSpcNo = Rs.Fields("SPCNO").Value
''''                strKey = strSpcYy & strSpcNo
''''                If strTemp <> strKey Then
''''                    '## �ٸ� ���ڵ��ȣ �϶��� ������� ǥ��
''''                    If .MaxRows <= .DataRowCnt Then
''''                        .MaxRows = .MaxRows + 1
''''                        .Row = .MaxRows
''''                    Else
''''                        .Row = .DataRowCnt + 1
''''                    End If
''''
''''                    .Col = TReadyEnum.ccNo:      .Value = .Row
''''                    .Col = TReadyEnum.ccPtId:    .Value = Rs.Fields("PTID").Value & ""
''''                    .Col = TReadyEnum.ccName:    .Value = Rs.Fields("NAME").Value & ""
''''                    .Col = TReadyEnum.ccAccNo:   .Value = Rs.Fields("WORKAREA").Value & "-" & _
''''                                                          Mid$(Rs.Fields("ACCDT").Value, 3) & "-" & _
''''                                                          Rs.Fields("ACCSEQ").Value
''''
''''                    vBarNo = Format$(strSpcYy, String$(SPCYYLEN, "0")) & Format$(strSpcNo, String$(SPCNOLEN, "0"))
''''                    vBarNo = Format$(vBarNo, String$(SPCLEN, "#"))
''''
''''                    .Col = TReadyEnum.ccBarNo:   .Value = vBarNo
'''''                    .Col = TReadyEnum.ccSexAge:  .Value = Rs.Fields("SEX").Value & "" & "/" & _
'''''                                                          mGetAge(Mid$(Rs.Fields("SSN").Value & "", 1, 6))
'''''                    .Col = TReadyEnum.ccStatFg:  .Value = IIf(Rs.Fields("STATFG").Value & "" = "1", "Y", "")
'''''                    .Col = TReadyEnum.ccWardId:  .Value = Rs.Fields("WARDID").Value & ""
'''''                    .Col = TReadyEnum.ccDept:    .Value = Rs.Fields("DEPTCD").Value & ""
'''''                    .Col = TReadyEnum.ccSpcNm:   .Value = Rs.Fields("SPCNM").Value & ""
'''''                    .Col = TReadyEnum.ccTestNms: .Value = Rs.Fields("TESTNM").Value & ""
'''''                    .Col = TReadyEnum.ccRcvNm:   .Value = Rs.Fields("RCVNM").Value & ""
'''''                    .Col = TReadyEnum.ccRcvDt:   .Value = Format$(Rs.Fields("RCVDT").Value & "", "####-##-##") & " " & _
''''                                                          Mid$(Rs.Fields("RCVTM").Value & "", 1, 2) & ":" & _
''''                                                          Mid$(Rs.Fields("RCVTM").Value & "", 3, 2)
''''
''''                    '## 1.2.3:  (2005-06-14)
''''                    '   - ó�渮��ũ�� ��ȸ�ϵ��� ����
'''''                    .Col = TReadyEnum.ccRmk:     .Value = Rs.Fields("MESG").Value & ""
''''                    strTemp = strKey
''''                Else
''''                    '## ���� ���ڵ��ȣ �϶��� �˻�� ǥ��
'''''                    .Col = TReadyEnum.ccTestNms
'''''                    .Value = .Value & "," & Rs.Fields("TESTNM").Value & ""
''''                End If
'                Rs.MoveNext
'            Loop
'            tblReady.Visible = True
'
''            lblCnt.Caption = CStr(.DataRowCnt)
'        End With
'    End If
'
'    Rs.Close
'    Set Rs = Nothing
'    Set objAccInfo = Nothing
'    Me.MousePointer = vbDefault
'    Exit Sub
'
'Errors:
'    Set Rs = Nothing
'    Set objAccInfo = Nothing
'    Me.MousePointer = vbDefault
'    MsgBox Err.Description, vbCritical, "����"
'End Sub

Private Sub cmdSearch_Click()
    Dim varTemp As Variant
    Dim i       As Integer
    Dim strBarcode  As String
    
    
    varTemp = Get_NewResult
    varTemp = Split(varTemp, "/")
    
    For i = 0 To UBound(varTemp)
        If varTemp(i) <> "" Then
            
            strBarcode = varTemp(i)
            
'        strSpcYy = Rs.Fields("SPCYY").Value
'        strSpcYy = Mid$(strSpcYy, 1, SPCYYLEN)
'        lngSpcNo = Rs.Fields("SPCNO").Value
'        lngSpcNo = Mid$(lngSpcNo, 1, SPCNOLEN)
'
'        vBarNo = Format$(strSpcYy, String$(SPCYYLEN, "0")) & Format$(lngSpcNo, String$(SPCNOLEN, "0"))
'
'        vBarNo = Format$(vBarNo, String$(SPCLEN, "#"))
            strBarcode = strBarcode & CheckDisit(strBarcode)
            
            Call GetOrder(strBarcode)
        End If
    Next
    
    
End Sub

'   Newest Result Recordset
Public Function Get_NewResult() As String

    Dim Rs          As ADODB.Recordset
    Dim objAccInfo  As clsIISAccInfo    '�������� Ŭ����
    Dim strEqpCd    As String           '����ڵ�
    Dim strFromDt   As String           'From Date
    Dim strToDt     As String           'To Date
    Dim strKey      As String           'Spread�� Ű(SpcYy+SpcNo)
    Dim strSpcYy    As String           '���ڵ��ȣ(����)
    Dim strSpcNo    As String           '���ڵ��ȣ(����)
    Dim strTemp     As String
    Dim varTemp     As Variant
    Dim i           As Integer
    
    strFromDt = Format$(dtpFrDate.Value, "YYYYMMDD")
    strToDt = Format$(dtpToDate.Value, "YYYYMMDD")
    
    Me.MousePointer = vbHourglass
    Call mTblClear(tblReady)
    strTemp = ""
    
On Error GoTo Errors
    Set objAccInfo = New clsIISAccInfo
    Set Rs = objAccInfo.GetTargetSpcs_Allergy(mEqpCd, strFromDt, strToDt)
    If Not (Rs.BOF Or Rs.EOF) Then
        With tblReady
            Do Until Rs.EOF
                strSpcYy = Rs.Fields("SPCYY").Value
                strSpcNo = Rs.Fields("SPCNO").Value
        
                strKey = Format$(strSpcYy, String$(SPCYYLEN, "0")) & Format$(strSpcNo, String$(SPCNOLEN, "0"))
                
                'strKey = strSpcYy & strSpcNo
                If strTemp <> strKey Then
                    '## �ٸ� ���ڵ��ȣ �϶��� ������� ǥ��
                    varTemp = varTemp & strKey & "/"
                    
                Else
                    '## ���� ���ڵ��ȣ �϶��� �˻�� ǥ��
'                    .Col = TReadyEnum.ccTestNms
'                    .Value = .Value & "," & Rs.Fields("TESTNM").Value & ""
                End If
                strTemp = strKey
                Rs.MoveNext
            Loop
            
        End With
    End If

    Rs.Close
    Set Rs = Nothing
    Set objAccInfo = Nothing
    Me.MousePointer = vbDefault
    
    Get_NewResult = varTemp
    
    Exit Function
    
Errors:
    Set Rs = Nothing
    Set objAccInfo = Nothing
    Me.MousePointer = vbDefault
    MsgBox Err.Description, vbCritical, "����"
End Function

Private Sub Form_Activate()
    MainFrm.lblMenuNm = Me.Caption
'    Me.MDIActiveX.WindowState = ccMaximize
End Sub

'   Access DB Connect
Public Function Set_DbConnect_Jet() As Boolean


    Dim DB_Name         As String
    Dim UserName        As String
    Dim Password        As String
    Dim blnWinNTAuth    As Boolean
    Dim strSrcfile      As String
    Dim strDestFile     As String
    
    Set AdoCn = New ADODB.Connection

On Error GoTo ConnectError

    'DB_Name = "C:\Program Files (x86)\LG Life Sciences\AdvanSure AlloStationSmart\DB\DBUnit.mdb" 'GetAlloConfig("MDBPath")      'C:\Program Files\LG Life Sciences\AdvanSure AlloStationSmart\DB
    DB_Name = GetAlloConfig("MDBPath")
    UserName = GetAlloConfig("USERNAME")    '"admin"
    Password = GetAlloConfig("PASSWORD")    '"reader_admin"

    If (DB_Name = "") Or (UserName = "") Then
        Set_DbConnect_Jet = False
        Set AdoCn = Nothing
        Exit Function
    End If
        
    With AdoCn
        .ConnectionTimeout = 25
        .CursorLocation = adUseClient
        .Provider = "Microsoft.Jet.OLEDB.4.0"
        .Properties("Mode").Value = adModeReadWrite
        .Properties("Persist Security Info").Value = False
        .Properties("Data Source").Value = DB_Name
        .Properties("User ID").Value = UserName
        .Properties("Jet OLEDB:Database Password").Value = Password
        .Properties("Jet OLEDB:Compact Without Replica Repair").Value = True
        .Open
    End With

    Set_DbConnect_Jet = True
    
 Exit Function

ConnectError:
    '   ����ó��
    MsgBox "   Error No. : " & Err.Number & vbCrLf & _
           " Description : " & Err.Description & vbCrLf & _
           "      Source : " & Err.Source & vbCrLf & vbCrLf _
           , vbCritical, " DB Open Error"

    If AdoCn.State <> adStateOpen Then
        Set_DbConnect_Jet = False
        Set AdoCn = Nothing
    End If
    
End Function



Private Sub Form_Load()
    Dim strInterval As String
    Dim strPath     As String
    
    Me.Caption = mEqpKey
    
    Me.MousePointer = vbHourglass
    
    Set mIntErrors = New clsIISIntErrors
    Set mIntLib = New clsIISInterface
    
    Call CtlClear
    Call mIntLib.SetConfig(mEqpCd, mEqpKey)
    DoEvents
    
    txtFileNm.Text = "AlloScan_" & Format(Now, "yyyymmdd")
    
    dtpFrDate.Value = Now
    dtpToDate.Value = Now
    dtpResult.Value = Now
    
    strInterval = GetAlloConfig("Interval")
    strPath = GetAlloConfig("Path")
    
    tmrAllo.Interval = CLng(strInterval) '30000
    tmrAllo.Enabled = True
    Call WinExec("C:\TEMP\HostInterface.exe", 0)
    FileAllo.Path = strPath '"C:\TEMP"
        
    DoEvents
    
    
    Me.MousePointer = vbDefault
    
End Sub

Private Function GetAlloConfig(ByVal strConfigNm As String) As String

Dim strFileName As String
Dim strReturnedString As String

    strFileName = App.Path & "\allo.ini"
    
    strReturnedString = String(1024, " ")
    GetPrivateProfileString "ALLOSCAN", strConfigNm, "", strReturnedString, Len(strReturnedString), strFileName
    strReturnedString = Trim(Replace(strReturnedString, Chr(0), Chr(32), 1, -1, vbBinaryCompare))
    GetAlloConfig = strReturnedString
    
End Function


Private Sub Form_Deactivate()
'    Me.MDIActiveX.WindowState = ccMinimize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mIntLib = Nothing
    Set mIntErrors = Nothing
    Set frmIISAlloStation = Nothing
End Sub



Private Sub cmdAlarm_Click()
    Dim objErrorShow As clsIISErrorShow     '������ ǥ�� Ŭ����
    
    Set objErrorShow = New clsIISErrorShow
    Call objErrorShow.ShowErrors(mIntErrors)
    Set objErrorShow = Nothing
    
    '## ������ ������ ��ư���� �������, ������ ��� ������
    cmdAlarm.BackColor = IIf(mIntErrors.Count = 0, &HF4F0F2, vbRed)
    
    '## 1.0.1: �̻��(2005-02-22)
    '   - Alarmâ�� ������ ��Ŀ���� txtBarNo�� �̵�
'    txtBarNo.SetFocus

End Sub

Private Sub cmdClear_Click()
    Call CtlClear
    Call mIntLib.AccInfos.RemoveAll
    
'    txtBarNo.SetFocus
End Sub

Private Sub cmdExit_Click()
    
'    With MSComm
'        '## �̹� ��Ʈ�� �������
'        If .PortOpen Then
'            If MsgBox(Me.Caption & " ���� ����Ǿ� �ֽ��ϴ�." & vbNewLine & vbNewLine & _
'                      Me.Caption & " �������̽��� �����Ͻðڽ��ϱ�?", vbYesNo + vbCritical) = vbYes Then
'                Unload Me
'            End If
'        End If
'
'    End With
    
    Unload Me
End Sub



Private Sub tmrAllo_Timer()
    Dim intIdx      As Integer
    Dim strSrcfile  As String
    Dim strDestFile As String
'    Dim strBuffer   As String
    Dim strtmpBuf   As String
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long
    Dim intCnt      As Integer
    Dim strBuf      As String
    Dim TextLine

    FileAllo.Refresh
    
    DoEvents
    
    For intIdx = 0 To FileAllo.ListCount - 1
        FileAllo.ListIndex = intIdx
        
        'FileAllo.FileName = ALLO_201606081241.patlist
        If Mid(FileAllo.FileName, 6, 8) <> Format(Now, "yyyymmdd") Then
            GoTo RST
        End If
        
        If Right(FileAllo.Path, 1) = "\" Then
            strSrcfile = FileAllo.Path & FileAllo.FileName     ' ���� ���� �̸��� �����մϴ�.
        Else
            strSrcfile = FileAllo.Path & "\" & FileAllo.FileName    ' ���� ���� �̸��� �����մϴ�.
        End If
        
        
        Open strSrcfile For Input As #3

        strBuffer = ""
        strBuf = ""
        
        Do While Not EOF(3)
            Line Input #3, TextLine ' ������ ������ ���� �о���Դϴ�.
            strBuf = strBuf & TextLine & vbCr
        Loop

        Close #3
        
        '��� ���� �̸��� ����
        strDestFile = App.Path & "\Log\" & FileAllo.FileName
        '������ ��� ����
        FileCopy strSrcfile, strDestFile
        
        
        lblFilenm.Caption = strSrcfile
        FileAllo.Refresh
        
        lngBufLen = Len(strBuf)
        
        strBuffer = strBuf
        
        If Len(strBuffer) > 300 Then
            Call cmdMakeWS_Click
            strBuffer = ""
        End If
RST:
    Next
End Sub

Private Sub tmrResult_Timer()
'    Get_SearchList
End Sub

Private Sub Get_SearchList()
    Dim intRow      As Long
    Dim strAge      As String
    Dim strTransDt  As String
    Dim strHMsg     As String
    Dim strDMsg     As String
    
    Dim objIntInfo   As clsIISIntInfo    '�������̽� ��ü���� Ŭ����
    Dim objIntNms    As clsIISIntNms     '��� �˻��׸� �÷��� Ŭ����
    Dim objBuffer    As clsIISBuffer     '���� Ŭ����
    
    Dim vWorkNo      As Variant  'Spread�� WorkNo
    Dim vBarNo       As Variant  'Spread�� ���ڵ��ȣ
    Dim strRcvBuf    As String   '������ Data
    Dim strType      As String   '������ Record Type
    Dim strBarNo     As String   '������ BarNO
    Dim strWorkNo    As String   '������ WorkNo
    Dim strIntResult As String   '������ �˻���
    Dim strIntBase   As String   '������ ������ �˻��
    Dim strResult    As String   'LIS �˻���
    Dim strClass     As String   'Class �������
    Dim strTemp      As String
    Dim i            As Long
'    Dim intRow       As Integer
    Dim blnSameBar   As Boolean
    Dim intCnt As Integer
    
    Dim Y, Y1, Y2, Y3, X1, X2
    
    Set objIntNms = mIntLib.IntNms
    

    Set AdoRS = Get_ResultList

    If Not AdoRS.BOF Then
'        tblComplete.MaxRows = AdoRS.RecordCount
        intRow = 1: strTransDt = ""
        Do Until AdoRS.EOF
'           Call EditRcvData
            strBarNo = AdoRS.Fields("PATIENTID").Value
            Set objIntInfo = New clsIISIntInfo
            With objIntInfo
                .BarNo = strBarNo
                .SpcPos = strWorkNo
            End With

            For intCnt = 1 To 42
                If intCnt = 1 Then
                    If UCase(Trim(AdoRS.Fields("STRIPPANEL_A").Value & "")) = "FOOD" Then
                        strIntBase = "FD"
                    ElseIf UCase(Trim(AdoRS.Fields("STRIPPANEL_A").Value & "")) = "INHALANT" Then
                        strIntBase = "IN"
                    End If
                ElseIf intCnt = 22 Then
                    If UCase(Trim(AdoRS.Fields("STRIPPANEL_B").Value & "")) = "FOOD" Then
                        strIntBase = "FD"
                    ElseIf UCase(Trim(AdoRS.Fields("STRIPPANEL_A").Value & "")) = "INHALANT" Then
                        strIntBase = "IN"
                    End If
                End If
                
                
                strIntBase = Mid(strIntBase, 1, 2) & Format(intCnt, "00")
                
                If intCnt <= 21 Then
                    strResult = "BANDVAL_A" & intCnt + 1
                Else
                    strResult = "BANDVAL_B" & intCnt - 20
                End If
                
                strIntResult = AdoRS.Fields(strResult).Value
                
'                On Error Resume Next
                If Mid(strIntBase, 1, 2) = "FD" Then
                    Select Case intCnt
                    Case 9, 16, 22, 26, 27, 30, 37, 40 '-- �Լ�A
                        Y = strIntResult
                        Y1 = (Y - VarD) / (VarA - VarD)
                        Y2 = Y1 / (1 - Y1)
                        If Y2 >= 0 Then
                            Y3 = Log(Y2)
                        Else
                            Y3 = Log(Abs(Y2))
                        End If
                        X1 = (Y3 - AN) / AM
                        X2 = Exp(X1)
                        strIntResult = X2
                    Case 1, 3, 4, 6, 7, 11, 12, 15, 18, 20, 24, 39 '-- �Լ�B
                        Y = strIntResult
                        Y1 = (Y - VarD) / (VarA - VarD)
                        Y2 = Y1 / (1 - Y1)
                        'Y3 = Log(Y2)
                        If Y2 >= 0 Then
                            Y3 = Log(Y2)
                        Else
                            Y3 = Log(Abs(Y2))
                        End If
                        X1 = (Y3 - BN) / BM
                        X2 = Exp(X1)
                        strIntResult = X2
                    Case 2, 5, 8, 10, 13, 14, 17, 19, 21, 23, 25, 28, 29, 31, 32, 33, 34, 35, 36, 38, 41, 42 '-- �Լ�C
                        Y = strIntResult
                        Y1 = (Y - VarD) / (VarA - VarD)
                        Y2 = Y1 / (1 - Y1)
                        'Y3 = Log(Y2)
                        If Y2 >= 0 Then
                            Y3 = Log(Y2)
                        Else
                            Y3 = Log(Abs(Y2))
                        End If
                        
                        X1 = (Y3 - Cn) / CM
                        X2 = Exp(X1)
                        strIntResult = X2
                    End Select
                    

                ElseIf Mid(strIntBase, 1, 2) = "IN" Then
                    Select Case intCnt
                    Case 13, 18, 21, 22, 24, 25, 26, 29, 37, 39, 40 '-- �Լ�A
                        Y = strIntResult
                        Y1 = (Y - VarD) / (VarA - VarD)
                        Y2 = Y1 / (1 - Y1)
                        Y3 = Log(Abs(Y2))
                        X1 = (Y3 - AN) / AM
                        X2 = Exp(X1)
                        strIntResult = X2
                    Case 1, 3, 5, 6, 7, 9, 12, 14, 17, 19, 30, 38 '-- �Լ�B
                        Y = strIntResult
                        Y1 = (Y - VarD) / (VarA - VarD)
                        Y2 = Y1 / (1 - Y1)
                        Y3 = Log(Abs(Y2))
                        X1 = (Y3 - BN) / BM
                        X2 = Exp(X1)
                        strIntResult = X2
                    Case 2, 4, 8, 10, 11, 15, 16, 20, 23, 27, 28, 31, 32, 33, 34, 35, 36, 41, 42 '-- �Լ�C
                        Y = strIntResult
                        Y1 = (Y - VarD) / (VarA - VarD)
                        Y2 = Y1 / (1 - Y1)
                        Y3 = Log(Abs(Y2))
                        X1 = (Y3 - Cn) / CM
                        X2 = Exp(X1)
                        strIntResult = X2
                    End Select
                End If

                If strIntBase = "FD01" Or strIntBase = "IN01" Then
                    If strIntResult > 100 Then
                        strIntResult = ">100"
                        strClass = "P ����" '������
                    ElseIf strIntResult <= 100 Then
                        strIntResult = "=<100"
                        strClass = "N ����" '����
                    End If
                Else
                    If strIntResult < 0.35 Then
                        strClass = "0"
                        '-- 2010.04.07 �ܴ� ����� ������ �䱸����
                        '-- Class 0 �� "0.00" ���� ġȯ�Ѵ�.
                        'strIntResult = Format(strIntResult, "#0.#0")
                        strIntResult = "0.00"
                    ElseIf strIntResult >= 0.35 And strIntResult < 0.7 Then
                        strClass = "1" '"*"
                        strIntResult = Format(strIntResult, "#0.#0")
                    ElseIf strIntResult >= 0.7 And strIntResult < 3.5 Then
                        strClass = "2" '"**"
                        strIntResult = Format(strIntResult, "#0.#0")
                    ElseIf strIntResult >= 3.5 And strIntResult < 17.5 Then
                        strClass = "3" '"***"
                        strIntResult = Format(strIntResult, "#0.#0")
                    ElseIf strIntResult >= 17.5 And strIntResult < 50 Then
                        strClass = "4" '"****"
                        strIntResult = Format(strIntResult, "#0.#0")
                    ElseIf strIntResult >= 50 And strIntResult < 100 Then
                        strClass = "5" '"*****"
                        strIntResult = Format(strIntResult, "#0.#0")
                    ElseIf strIntResult >= 100 Then
                        strClass = "6" '"******"
                        strIntResult = ">=100"
                    Else
                        strIntResult = "0.00"
                        strClass = ""
                    End If
                End If
                If strIntResult = "0." Then strIntResult = "0.00"

                strIntResult = strIntResult & "  " & strClass
                strResult = strIntResult
'                If strIntBase = "FD01" Or strIntBase = "IN01" Then
'                '    strIntResult = strIntResult & " " & strClass
'                Else
'                    strIntResult = strIntResult & " " & strClass
'                End If
'                x >= 100
'                50 =< x <100
'                17.5 =< x < 50
'                3.5 =< x < 17.5
'                0.7 =< x < 3.5
'                0.35 =< x < 0.7
'                x < 0.35

                '## ������� "?", ">", "<" ���ԵǾ� ������ ������ ǥ��
                'If IsNumeric(strIntResult) Then
                '    strResult = strIntResult
                'Else
                '    strResult = IISERROR
                'End If

                If objIntNms.ExistIntBase(strIntBase) Then
                    Call objIntInfo.IntResults.Add(strIntBase, objIntNms.GetIntNm(strIntBase), _
                         strIntResult, strResult, strClass)
                End If
            Next
            
            Call SaveServer(objIntInfo)
            
            intRow = intRow + 1
            AdoRS.MoveNext
        Loop
    End If
    
    Set AdoRS = Nothing
    tmrResult.Enabled = False
End Sub

'   Result List Recordset
Public Function Get_ResultList() As ADODB.Recordset
    Dim strSql      As String

On Error GoTo ErrorTrap

             strSql = "SELECT * "
    strSql = strSql & "  FROM RESULTS "
'    strSql = strSql & " WHERE EXAMDATE = '" & Format(Now, "yyyy-mm-01") & "' "
'    strSql = strSql & " WHERE PATIENTID = '10000189895' "
    
    Set AdoRS = New ADODB.Recordset
    If Get_Recordset(AdoCn, strSql, AdoRS, "") Then
        Set Get_ResultList = AdoRS
        'blnRS = True
    Else
        Set Get_ResultList = Nothing
        'blnRS = False
    End If
    
    Set AdoRS = Nothing

Exit Function

ErrorTrap:
    Set AdoRS = Nothing
'    blnRS = False

End Function

'   Record Set Open
Public Function Get_Recordset(ByVal AdoCn As ADODB.Connection, ByVal strSql As String, _
                             ByVal AdoRS As ADODB.Recordset, _
                             Optional Call_Name As String, _
                             Optional Cursor_Location As ADODB.CursorLocationEnum = adUseClient, _
                             Optional Cursor_Type As ADODB.CursorTypeEnum = adOpenStatic, _
                             Optional Lock_Type As ADODB.LockTypeEnum = adLockPessimistic) As Boolean

On Error GoTo DBOpenRsError
    
    With AdoRS
        .CursorLocation = Cursor_Location
        .Source = strSql
        .ActiveConnection = AdoCn
        .CursorType = Cursor_Type
        .LockType = Lock_Type
        .Open
    End With
    
    Get_Recordset = True

Exit Function

DBOpenRsError:
    Set AdoRS = Nothing
    Get_Recordset = False

End Function

Private Sub txtBarNo_GotFocus()
    With txtBarNo
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Public Sub txtBarNo_KeyDown(KeyCode As Integer, Shift As Integer)
    '## �ش� ���ڵ��ȣ�� ���� �������� ��ȸ
    If KeyCode = vbKeyReturn Then
        Me.MousePointer = vbHourglass
        Call GetOrder(Trim(txtBarNo.Text))
        Me.MousePointer = vbDefault
    End If
End Sub

Private Sub txtBarNo_KeyPress(KeyAscii As Integer)
    '## ���ڸ� �Է�
    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub tblReady_Click(ByVal Col As Long, ByVal Row As Long)
'''    Dim objAccInfo  As clsIISAccInfo    '�������� Ŭ����
'''    Dim vBarNo      As Variant          'Spread�� ���ڵ��ȣ
'''    Dim strSpcYy    As String           '��ü����
'''    Dim lngSpcNo    As Long             '��ü��ȣ
'''
'''    If Row = 0 Then Exit Sub
'''
''''    Call CtlClear(ccLabel)
''''    Call mTblClear(tblResult)
'''    Call tblReady.GetText(TReadyEnum.ccBarNo, Row, vBarNo)
'''    If vBarNo = "" Then Exit Sub
'''
''''    Call GetOrder(vBarNo)
'''
'''    Set objAccInfo = mIntLib.GetAccInfo(vBarNo)
'''    If Not (objAccInfo Is Nothing) Then
'''        '## tblReady, tblresult, Label�� ����ǥ��
''''        Call SetReady(objAccInfo)
'''        Call SetLabel(objAccInfo)
'''        Call SetResult(objAccInfo)
'''
'''
''''        Call SetOrderWS(objAccInfo)
'''
'''        Set objAccInfo = Nothing
'''    End If
'''    txtBarNo.Text = "": txtBarNo.SetFocus
'''
''''    strSpcYy = Mid$(vBarNo, 1, SPCYYLEN)
''''    lngSpcNo = CLng(Mid$(vBarNo, SPCYYLEN + 1, SPCNOLEN))
''''    Set objAccInfo = mIntLib.AccInfos(strSpcYy, lngSpcNo)
'''
'''
'''    '## tblResult, Label�� ����ǥ��
''''    Call SetLabel(objAccInfo)
''''    Call SetResult(objAccInfo)
'''
'''    Set objAccInfo = Nothing

    Dim objAccInfo  As clsIISAccInfo    '�������� Ŭ����
    Dim vBarNo      As Variant          'Spread�� ���ڵ��ȣ
    Dim strSpcYy    As String           '��ü����
    Dim lngSpcNo    As Long             '��ü��ȣ
    
    If Row = 0 Then Exit Sub
    
    Call CtlClear(ccLabel)
    Call mTblClear(tblResult)
    Call tblReady.GetText(TReadyEnum.ccBarNo, Row, vBarNo)
    If vBarNo = "" Then Exit Sub
    
    strSpcYy = Mid$(vBarNo, 1, SPCYYLEN)
    lngSpcNo = CLng(Mid$(vBarNo, SPCYYLEN + 1, SPCNOLEN))
    Set objAccInfo = mIntLib.AccInfos(strSpcYy, lngSpcNo)
    
    '## tblResult, Label�� ����ǥ��
    Call SetLabel(objAccInfo)
    Call SetResult(objAccInfo)
    
    Set objAccInfo = Nothing


End Sub

Private Sub tblReady_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    Set mPopup = New clsIISPopup
    With mPopup
        .AddMenu DELETE, "Delete"
        .AddMenu DELETEALL, "Delete All"
        .PopupMenus Me.hWnd
    End With
    Set mPopup = Nothing
End Sub

Private Sub tblComplete_Click(ByVal Col As Long, ByVal Row As Long)
    Dim strQcFg     As String   'QC����
    Dim strResult   As String   'LIS ���
    Dim strTemp     As String
    Dim i           As Long
    
    Call CtlClear(ccLabel)
    Call mTblClear(tblResult)
    With tblComplete
        .Row = Row
        
        .Col = TCompleteEnum.ccQcFg:    strQcFg = .Text
        If strQcFg = "0" Then
            .Col = TCompleteEnum.ccPtId:    lblPtId.Caption = .Text
            .Col = TCompleteEnum.ccName:    lblName.Caption = .Text
            .Col = TCompleteEnum.ccSexAge:  lblSexAge.Caption = .Text
            .Col = TCompleteEnum.ccDoctNm:  lblDoctNm.Caption = .Text
            .Col = TCompleteEnum.ccDeptNm:  lblDeptNm.Caption = .Text
            .Col = TCompleteEnum.ccWardNm:  lblWardNm.Caption = .Text
            .Col = TCompleteEnum.ccStatFg:  lblStatFg.Caption = .Text
            .Col = TCompleteEnum.ccSpcNm:   lblSpcNm.Caption = .Text
            '-- �߰�
            .Col = TCompleteEnum.ccNo:      lblPnlNm.Caption = IIf(.Text = "IN", "INHALANT", "FOOD")

        ElseIf strQcFg = "1" Then
            .Col = TCompleteEnum.ccPtId:    lblPtId.Caption = .Text
            .Col = TCompleteEnum.ccName:    lblName.Caption = .Text
            .Col = TCompleteEnum.ccSexAge:  lblSexAge.Caption = .Text
        End If
        
        For i = TCompleteEnum.ccResult To .DataColCnt
            .Col = i:   strTemp = .Text
            
            '## 1.0.3: �̻��(2005-06-24)
            '   - ȭ��ǥ�� ���׼���
            If tblResult.MaxRows <= tblResult.DataRowCnt Then
                tblResult.MaxRows = tblResult.MaxRows + 1
                tblResult.Row = tblResult.MaxRows
            Else
                tblResult.Row = tblResult.DataRowCnt + 1
            End If
            
            tblResult.Col = TResultEnum.ccTestNm:       tblResult.Text = mGetP(strTemp, TResultEnum.ccTestNm, DIV)
            tblResult.Col = TResultEnum.ccUnit:         tblResult.Text = mGetP(strTemp, TResultEnum.ccUnit, DIV)
            tblResult.Col = TResultEnum.ccHLDiv:        tblResult.Text = mGetP(strTemp, TResultEnum.ccHLDiv, DIV)
            tblResult.Col = TResultEnum.ccDPDiv:        tblResult.Text = mGetP(strTemp, TResultEnum.ccDPDiv, DIV)
            tblResult.Col = TResultEnum.ccRef:          tblResult.Text = mGetP(strTemp, TResultEnum.ccRef, DIV)
            tblResult.Col = TResultEnum.ccEqpResult:    tblResult.Text = mGetP(strTemp, TResultEnum.ccEqpResult, DIV)
            tblResult.Col = TResultEnum.ccClass:        tblResult.Text = mGetP(strTemp, TResultEnum.ccClass, DIV)
            tblResult.Col = TResultEnum.ccLISResult
                strResult = mGetP(strTemp, TResultEnum.ccLISResult, DIV)
                tblResult.Text = strResult
                If strResult = IISERROR Then
                    tblResult.ForeColor = vbRed
                Else
                    tblResult.ForeColor = vbBlack
                End If
        Next i
        
        'Call AllergyPrint(Col, Row)
        
    End With

End Sub



'Private Sub AllergyPrint(ByVal sPtid As Long, ByVal sPanel As Long)
'    Dim strPanel    As String
'    Dim strBarNo    As String
'    Dim strResult   As String   'LIS ���
'    Dim strTemp     As String
'    Dim i           As Long
'    Dim j           As Integer
'
'
'    'Call mTblClear(vasPrint)
'    With tblComplete
'        .Row = Row
'        .Col = TCompleteEnum.ccNo:    strPanel = .Text
'        If strPanel = "IN" Then
'            strPanel = "INHALANT"
'        ElseIf strPanel = "FD" Then
'            strPanel = "FOOD"
'        End If
'        .Col = TCompleteEnum.ccBarNo:    strBarNo = .Text
'        If strBarNo <> "" Then
'            .Col = TCompleteEnum.ccPtId:    lblPtId.Caption = .Text
'            .Col = TCompleteEnum.ccName:    lblName.Caption = .Text
'            .Col = TCompleteEnum.ccSexAge:  lblSexAge.Caption = .Text
'            .Col = TCompleteEnum.ccDoctNm:  lblDoctNm.Caption = .Text
'            .Col = TCompleteEnum.ccDeptNm:  lblDeptNm.Caption = .Text
'            .Col = TCompleteEnum.ccWardNm:  lblWardNm.Caption = .Text
'            .Col = TCompleteEnum.ccStatFg:  lblStatFg.Caption = .Text
'            .Col = TCompleteEnum.ccSpcNm:   lblSpcNm.Caption = .Text
'        End If
'
'        '-- �˷��� Information
'        Call vasPrint.SetText(1, 4, "ID:")
'        Call vasPrint.SetText(2, 4, strBarNo)
'        Call vasPrint.SetText(3, 4, "AGE:")
'        Call vasPrint.SetText(4, 4, mGetP(lblSexAge.Caption, 2, "/"))
'        Call vasPrint.SetText(5, 4, "�Ƿڰ�:")
'        Call vasPrint.SetText(6, 4, lblWardNm.Caption)
'
'        Call vasPrint.SetText(1, 5, "����:")
'        Call vasPrint.SetText(2, 5, lblName.Caption)
'        Call vasPrint.SetText(3, 5, "SEX:")
'        Call vasPrint.SetText(4, 5, IIf(mGetP(lblSexAge.Caption, 1, "/") = "M", "����", "����"))
'        Call vasPrint.SetText(5, 5, "�˻���:")
'        Call vasPrint.SetText(6, 5, lblDoctNm.Caption)
'
'        Call vasPrint.SetText(1, 6, "�˻���:")
'        Call vasPrint.SetText(2, 6, lblName.Caption)
'        Call vasPrint.SetText(3, 6, "PANEL:")
'        Call vasPrint.SetText(4, 6, strPanel)
'        Call vasPrint.SetText(5, 6, "Ȯ����:")
'        Call vasPrint.SetText(6, 6, lblDoctNm.Caption)
'
'        '-- �˷��� Panel
'        If strPanel = "INHALANT" Then
'            Call vasPrint.SetText(3, 10, "INHALANT 1KO PANEL")
'        ElseIf strPanel = "FOOD" Then
'            Call vasPrint.SetText(3, 10, "FOOD 3KO PANEL")
'        End If
'
'        '-- �˷��� ���
'        Call vasPrint.SetText(2, 12, "Control")
'        Call vasPrint.SetText(3, 12, "[POSITIVE]")
'        j = 1
'        For i = TCompleteEnum.ccResult To .DataColCnt
'            .Col = i:   strTemp = .Text
'            Call vasPrint.SetText(1, j + 13, CStr(j))
'            Call vasPrint.SetText(2, j + 13, mGetP(strTemp, TResultEnum.ccTestNm, DIV))
'            Call vasPrint.SetText(3, j + 13, mGetP(strTemp, TResultEnum.ccEqpResult, DIV))
'            Call vasPrint.SetText(4, j + 13, mGetP(strTemp, TResultEnum.ccEqpResult, DIV))
'            Call vasPrint.SetText(5, j + 13, mGetP(strTemp, TResultEnum.ccEqpResult, DIV))
'            Call vasPrint.SetText(6, j + 13, mGetP(strTemp, TResultEnum.ccEqpResult, DIV))
'            j = j + 1
'            'vasPrint.Col = TResultEnum.ccTestNm:       vasPrint.Text = mGetP(strTemp, TResultEnum.ccTestNm, DIV)
'            'vasPrint.Col = TResultEnum.ccUnit:         vasPrint.Text = mGetP(strTemp, TResultEnum.ccUnit, DIV)
'            'vasPrint.Col = TResultEnum.ccHLDiv:        vasPrint.Text = mGetP(strTemp, TResultEnum.ccHLDiv, DIV)
'            'vasPrint.Col = TResultEnum.ccDPDiv:        vasPrint.Text = mGetP(strTemp, TResultEnum.ccDPDiv, DIV)
'            'vasPrint.Col = TResultEnum.ccRef:          vasPrint.Text = mGetP(strTemp, TResultEnum.ccRef, DIV)
'            'vasPrint.Col = TResultEnum.ccEqpResult:    vasPrint.Text = mGetP(strTemp, TResultEnum.ccEqpResult, DIV)
'            'vasPrint.Col = TResultEnum.ccClass:        vasPrint.Text = mGetP(strTemp, TResultEnum.ccClass, DIV)
'            'vasPrint.Col = TResultEnum.ccLISResult
'        Next i
'    End With
'
'End Sub

Private Sub MSComm_OnComm()
    Dim EVMsg As String
    Dim ERMsg As String
    Dim Ret   As Long
    
    '�ӽ� - �����
    If strTransData <> "" Then GoTo TransRst
    
    Select Case MSComm.CommEvent
        Case comEvReceive
            Dim Buffer      As Variant
            Dim BufChar     As String
            Dim lngBufLen   As Long
            Dim i           As Long
            
            Buffer = MSComm.Input
'�ӽ� - �����
TransRst:
            Buffer = strTransData
            
            Call mIntLib.WriteLog(Buffer, ccEqp)
            
            lngBufLen = Len(Buffer)
            For i = 1 To lngBufLen
                BufChar = Mid$(Buffer, i, 1)

                Select Case mIntLib.Phase
                    Case 1
                        'Debug.Print "----> BufChar " & BufChar
                        Select Case Asc(BufChar)
                            Case 5  'ENQ()
                                '-- �ӽ� - ����� ����ũ ����
                                'Debug.Print "----> ENQ ����"
'                                MSComm.Output = Chr(6) 'ACK()
                                '-- �ӽ� - ����� ����ũ ��
                                mIntLib.Phase = 2
                                mIntLib.BufCnt = 1
                        End Select
                    Case 2
                        Select Case Asc(BufChar)
                            Case 2
                                'Debug.Print "----> STX ����"
                                mIntLib.ClearBuffer
                            Case 23 'ETB()
                                'Debug.Print "----> ETB "
'                                mIntLib.phase = 3
                            Case 3  'ETX()
                                'Debug.Print "----> ETX "
                                mIntLib.Phase = 3
                            Case Asc(vbCr)
                                'Debug.Print "----> vbCr "
                                '-- �ӽ� - ����� ����ũ ����
'                                MSComm.Output = Chr(6)
                                '-- �ӽ� - ����� ����ũ ��
                                mIntLib.BufCnt = mIntLib.BufCnt + 1
                                mIntLib.Phase = 3
                            Case Else
                                Call mIntLib.AddBuffer(BufChar)
                        End Select
                    Case 3
                        Select Case Asc(BufChar)
                            Case 3 'ETX
                                'Debug.Print "----> ETX "
'                                mIntLib.Phase = 4
                            Case 2
                                'Debug.Print "----> STX "
                                mIntLib.Phase = 2
                            Case 4
                                'Debug.Print "----> EOT "
                                Call EditRcvData
                                '-- �ӽ� - ����� ����
                                mIntLib.Phase = 1
                                '-- �ӽ� - ����� ��
                                
                                '-- �ӽ� - ����� ����ũ ����
                                'Debug.Print "----> EditData "
'                                If mIntLib.State = "Q" Then
'                                    MSComm.Output = Chr(5)
'                                    mIntLib.SndPhase = 1
'                                    mIntLib.FrameNo = 1
'                                End If
''                                MSComm.Output = Chr(6) 'ACK()
'                                mIntLib.Phase = 4
                                '-- �ӽ� - ����� ����ũ ��
                        End Select
                    Case 4
                        Select Case Asc(BufChar)
                            Case 6
                                'Debug.Print "----> ACK "
                                If mIntLib.State = "Q" Then
                                    'Call SendOrdData
                                    'Debug.Print "----> SendData "
                                End If
                            Case 5
                                'Debug.Print "----> EOT "
                                MSComm.Output = Chr(6)
                                mIntLib.Phase = 2
                            Case 21
                                'Debug.Print "----> NAK "
                                MSComm.Output = Chr(5)
                            Case 4
                                'Debug.Print "----> EOT "
                                mIntLib.Phase = 1
                        End Select
                End Select
            Next i
        Case comEvSend
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

    If Len(EVMsg$) Then
        StatusBar.Panels(2).Text = EVMsg$
    ElseIf Len(ERMsg$) Then
        StatusBar.Panels(2).Text = ERMsg$
    End If
End Sub

'-----------------------------------------------------------------------------'
'   ��� : ���κ� ������ ������ ����
'-----------------------------------------------------------------------------'
Private Sub EditRcvData()
    Dim objIntInfo   As clsIISIntInfo    '�������̽� ��ü���� Ŭ����
    Dim objIntNms    As clsIISIntNms     '��� �˻��׸� �÷��� Ŭ����
    Dim objBuffer    As clsIISBuffer     '���� Ŭ����
    
    Dim vWorkNo      As Variant  'Spread�� WorkNo
    Dim vBarNo       As Variant  'Spread�� ���ڵ��ȣ
    Dim strRcvBuf    As String   '������ Data
    Dim strType      As String   '������ Record Type
    Dim strBarNo     As String   '������ BarNO
    Dim strWorkNo    As String   '������ WorkNo
    Dim strIntResult As String   '������ �˻���
    Dim strIntBase   As String   '������ ������ �˻��
    Dim strResult    As String   'LIS �˻���
    Dim strTemp      As String
    Dim i            As Long
    Dim intRow       As Integer
    Dim blnSameBar   As Boolean
    Dim strTest      As String
    Dim strClass     As String
    Dim varTmp       As Variant
    Dim intCnt       As Integer
    Dim strPart      As String
    Dim strState     As String
    
    strState = ""
'On Error Resume Next

    Set objIntNms = mIntLib.IntNms
    For Each objBuffer In mIntLib.Buffers
        strRcvBuf = objBuffer.Buffers
        varTmp = strRcvBuf
        'varTmp = Split(varTmp, "|")
        varTmp = Split(varTmp, ",")
        
        'For intCnt = 0 To UBound(varTmp) - 1
        For intCnt = 0 To UBound(varTmp) - 1
            If intCnt = 1 Then
                strPart = varTmp(intCnt)
                If strPart = "Panel 30 KO Inhalant A" Then
                    strPart = "IN|"
                ElseIf strPart = "Panel 30 KO Inhalant B" Then
                    strPart = "IN|"
                ElseIf strPart = "Panel 30 KO Food  A" Then
                    strPart = "FD|"
                ElseIf strPart = "Panel 30 KO Food  B" Then
                    strPart = "FD|"
                End If
            
                'MsgBox strPart
            End If
            
            'If intCnt = 8 Then
            If intCnt = 9 Then
                strBarNo = varTmp(intCnt)
                strBarNo = Mid(strBarNo, 1, 11)
                
                'strBarNo = "15001375841"
                'strBarNo = "15001358055"
                'MsgBox strBarNo
                Set objIntInfo = New clsIISIntInfo
                With objIntInfo
                    .BarNo = strBarNo
                    .SpcPos = Mid(strPart, 1, 2)
                End With
            End If
            
            'If intCnt > 8 And (InStr(intCnt, 4) > 0 Or InStr(intCnt, 9) > 0) Then
            If intCnt > 9 Then
                strIntBase = varTmp(intCnt)
                '������ ���� ��������
                'If strIntBase = "I81" Then Stop
                
                
                'Hx HX
                If UCase(strIntBase) = "HX" Then
                    strIntBase = "HX"
                End If
                'Debug.Print strIntBase
                If strIntBase = "Total_IgE" Then
                    strIntBase = "IgE"
                End If
                strIntBase = strPart & strIntBase
                
                
                strIntResult = varTmp(intCnt + 1)
                strResult = strIntResult
'                If IsNumeric(strIntResult) Then
'                    strResult = strIntResult
'                Else
'                    strResult = IISERROR
'                End If
                
                '-- Class �߰�
                If strIntBase = "FD|IgE" Or strIntBase = "IN|IgE" Then
                    If strIntResult = "<100" Then
                        strClass = "1.0"
                    ElseIf strIntResult = "100-200" Then
                        strClass = "2.0"
                    ElseIf strIntResult = ">200" Then
                        strClass = "3.0"
                    End If
                Else
                    If intCnt + 2 <= UBound(varTmp) Then
                        strClass = varTmp(intCnt + 2)
                        strClass = Format(strClass, "0.0")
                    End If
                End If

                '-- ������� Class���� ����
                strResult = strClass & " Class" & "(" & strResult & ")"

                If strIntResult <> "" Then
                    If objIntNms.ExistIntBase(strIntBase) Then
                        Call objIntInfo.IntResults.Add(strIntBase, objIntNms.GetIntNm(strIntBase), _
                             strResult, strResult, strClass)
                        strState = "R"
                    End If
                End If
                
            End If

        Next
        strPart = ""
        strBarNo = ""
        
        'MsgBox strState
        If strState = "R" Then
            Call SaveServer(objIntInfo)
        End If

    Next
    
    Set objIntNms = Nothing
    Set objBuffer = Nothing
End Sub

'-----------------------------------------------------------------------------'
'   ��� : �������, �������, ȭ��ǥ��
'   �μ� :
'       - pIntInfo : �������̽� ��ü���� Ŭ����
'-----------------------------------------------------------------------------'
Private Sub SaveServer(ByVal pIntInfo As clsIISIntInfo, Optional strPanel As String)
    Dim objAccInfo  As clsIISAccInfo    '�������� Ŭ����
    Dim vBarNo      As Variant 'Spread�� ���ڵ��ȣ
    Dim strBarNo    As String  '���ڵ��ȣ
    Dim strSpcYy    As String  '��ü����
    Dim lngSpcNo    As Long    '��ü��ȣ
    Dim i           As Long
    
    Me.MousePointer = vbHourglass
    
    strBarNo = pIntInfo.BarNo
    
    '## �������
    If mIntLib.CheckResult(pIntInfo) = -1 Then
        '## ���������� ������ ���ǥ��
        Call SetComplete1(pIntInfo)
        Call tblComplete_Click(1, tblComplete.DataRowCnt)
        Me.MousePointer = vbDefault
        Exit Sub
    Else
        '## ���������� ������ ���ǥ��
        strSpcYy = Mid$(strBarNo, 1, SPCYYLEN)
        lngSpcNo = CLng(Mid$(strBarNo, SPCYYLEN + 1, SPCNOLEN))
        Set objAccInfo = mIntLib.AccInfos(strSpcYy, lngSpcNo)
        
        Call SetComplete2(objAccInfo)
        
        tblComplete.Col = TCompleteEnum.ccNo
        tblComplete.Text = strPanel
        
        Call tblComplete_Click(1, tblComplete.DataRowCnt)
        Set objAccInfo = Nothing
        
        '## ClientDb, Server�� �������
        Call mIntLib.SaveClientDb(strSpcYy, lngSpcNo)
        Call mIntLib.SaveResult(strSpcYy, lngSpcNo)
        Call mIntLib.Remove(strSpcYy, lngSpcNo)
        StatusBar.Panels(2).Text = "��ü��ȣ:" & strBarNo & " �� ���������� ������� �߽��ϴ�."
    End If
    
    '## tblReady���� ���۵� ��ü����
    If mIntLib.BarPos = ccPC Then
        With tblReady
            For i = 1 To .DataRowCnt
                Call .GetText(TReadyEnum.ccBarNo, i, vBarNo)
                If CStr(vBarNo) = strBarNo Then
                    Call .DeleteRows(i, 1)
                    Exit For
                End If
            Next i
        End With
    End If

    Me.MousePointer = vbDefault
End Sub

'-----------------------------------------------------------------------------'
'   ��� : �ش� ���ڵ��ȣ�� ���� �������� ��ȸ, tblReady, tblResult�� ǥ��
'   �μ� :
'       - pBarNo : ���ڵ��ȣ
'-----------------------------------------------------------------------------'
Private Sub GetOrder(ByVal pBarNo As String)
    Dim objAccInfo As clsIISAccInfo     '�������� Ŭ����
    
    If pBarNo = "" Then Exit Sub
    
    gBarNo = pBarNo
    
    Set objAccInfo = mIntLib.GetAccInfo(pBarNo)
    If Not (objAccInfo Is Nothing) Then
        '## tblReady, tblresult, Label�� ����ǥ��
        Call SetReady(objAccInfo)
        Call SetLabel(objAccInfo)
        Call SetResult(objAccInfo)
        
        
'        Call SetOrderWS(objAccInfo)
        
        Set objAccInfo = Nothing
    End If
    txtBarNo.Text = ""
    'txtBarNo.SetFocus
End Sub

'-----------------------------------------------------------------------------'
'   ��� : �ش� ���ڵ��ȣ�� ���� �������� ��ȸ, tblReady, tblResult�� ǥ��
'   �μ� :
'       - pBarNo : ���ڵ��ȣ
'-----------------------------------------------------------------------------'
'Private Sub GetOrderByAccNo()
'    Dim objAccInfo As clsIISAccInfo     '�������� Ŭ����
'    Dim Rs   As Recordset
'    Dim vBarNo      As Variant  'Spread�� ���ڵ��ȣ
'    Dim strSpcYy    As String   '��ü����
'    Dim lngSpcNo    As Long     '��ü��ȣ
'
'    Set Rs = New Recordset
'    Set objAccInfo = New clsIISAccInfo
'    'rsBarcode = objAccInfo.GetTargetSpcs(mEqpCd, Format(dtpFrDate.Value, "yyyymmdd"), Format(dtpToDate.Value, "yyyymmdd"))
'
''    If pBarNo = "" Then Exit Sub
'
'    Set Rs = objAccInfo.GetTargetSpcs(mEqpCd, Format(dtpFrDate.Value, "yyyymmdd"), Format(dtpToDate.Value, "yyyymmdd"))
'    Do Until Rs.EOF
'        strSpcYy = Rs.Fields("SPCYY").Value
'        strSpcYy = Mid$(strSpcYy, 1, SPCYYLEN)
'        lngSpcNo = Rs.Fields("SPCNO").Value
'        lngSpcNo = Mid$(lngSpcNo, 1, SPCNOLEN)
'
'        vBarNo = Format$(strSpcYy, String$(SPCYYLEN, "0")) & Format$(lngSpcNo, String$(SPCNOLEN, "0"))
'
'        vBarNo = Format$(vBarNo, String$(SPCLEN, "#"))
'
'        Set objAccInfo = mIntLib.GetAccInfo(vBarNo)
'
'        If Not (objAccInfo Is Nothing) Then
'            '## tblReady, tblresult, Label�� ����ǥ��
'            Call SetReady(objAccInfo)
'            Call SetLabel(objAccInfo)
'            Call SetResult(objAccInfo)
'
'
'            Call SetOrderWS(objAccInfo)
'
'            Set objAccInfo = Nothing
'        End If
'        Rs.MoveNext
'    Loop
'
'    Rs.Close
'    Set Rs = Nothing
''    txtBarNo.Text = "": txtBarNo.SetFocus
'
'End Sub

'-----------------------------------------------------------------------------'
'   ��� : tblReady�� ����ǥ��
'   �μ� :
'       - pAccInfo : �������� Ŭ����
'-----------------------------------------------------------------------------'
Private Sub SetReady(ByVal pAccInfo As clsIISAccInfo)
    Dim lngWorkNo As Long   'WorkNo
    Dim objResult   As clsIISResult     '������� Ŭ����
    Dim blnFood     As Boolean
    Dim blnIn       As Boolean
    
    blnFood = False
    blnIn = False
    
    
    With tblReady
        If .MaxRows <= .DataRowCnt Then
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
        Else
            .Row = .DataRowCnt + 1
        End If
        
        .Col = TReadyEnum.ccNo:     .Text = CStr(lngWorkNo)
        .Col = TReadyEnum.ccBarNo:  .Text = gBarNo 'pAccInfo.GetBarNo
        .Col = TReadyEnum.ccAccNo:  .Text = mGetAccNo(pAccInfo.Workarea, pAccInfo.AccDt, pAccInfo.AccSeq)
        
        If pAccInfo.QcFg = "0" Then         '## �Ϲݰ�ü
            .Col = TReadyEnum.ccPtId:   .Text = pAccInfo.PtId
            .Col = TReadyEnum.ccName:   .Text = pAccInfo.Name
        ElseIf pAccInfo.QcFg = "1" Then     '## QC��ü
            .Col = TReadyEnum.ccPtId:   .Text = pAccInfo.CtrlCd
            .Col = TReadyEnum.ccName:   .Text = pAccInfo.LevelCd
        End If
        
        .Col = 6:  .Text = pAccInfo.Sex & "|" & pAccInfo.Ssn
        
        For Each objResult In pAccInfo.Results
            'Debug.Print objResult.TestNm & ":" & objResult.Testcd
            'If InStr(UCase(objResult.TestNm), "FOOD") > 0 Then
            If objResult.TestCd = "MASTF26" Then
                .Col = TReadyEnum.ccNo:     .Text = "FD"
                blnFood = True
                Exit For
            'ElseIf InStr(UCase(objResult.TestNm), "INHAL") > 0 Then
            ElseIf objResult.TestCd = "MASQE3" Then
                .Col = TReadyEnum.ccNo:     .Text = "IN"
                blnIn = True
                Exit For
            End If
        Next
        
        '-- �ٸ� ó���� �ִ��� Ȯ��(Food�� IN�� IN�̸� Food�� ã�´�.
        If blnFood = True Then
            For Each objResult In pAccInfo.Results
                'If InStr(UCase(objResult.TestNm), "INHALANT") > 0 Then
                If objResult.TestCd = "MASQE3" Then
                    If .MaxRows <= .DataRowCnt Then
                        .MaxRows = .MaxRows + 1
                        .Row = .MaxRows
                    Else
                        .Row = .DataRowCnt + 1
                    End If
                    
                    .Col = TReadyEnum.ccNo:     .Text = CStr(lngWorkNo)
                    .Col = TReadyEnum.ccBarNo:  .Text = gBarNo 'pAccInfo.GetBarNo
                    .Col = TReadyEnum.ccAccNo:  .Text = mGetAccNo(pAccInfo.Workarea, pAccInfo.AccDt, pAccInfo.AccSeq)
                    
                    If pAccInfo.QcFg = "0" Then         '## �Ϲݰ�ü
                        .Col = TReadyEnum.ccPtId:   .Text = pAccInfo.PtId
                        .Col = TReadyEnum.ccName:   .Text = pAccInfo.Name
                    ElseIf pAccInfo.QcFg = "1" Then     '## QC��ü
                        .Col = TReadyEnum.ccPtId:   .Text = pAccInfo.CtrlCd
                        .Col = TReadyEnum.ccName:   .Text = pAccInfo.LevelCd
                    End If
                    
                    .Col = 6:  .Text = pAccInfo.Sex & "|" & pAccInfo.Ssn
                    .Col = TReadyEnum.ccNo:     .Text = "IN"
                    Exit For
                End If
            Next
        End If
        
        If blnIn = True Then
            For Each objResult In pAccInfo.Results
                'If InStr(UCase(objResult.TestNm), "FOOD") > 0 Then
                If objResult.TestCd = "MASTF26" Then
                    If .MaxRows <= .DataRowCnt Then
                        .MaxRows = .MaxRows + 1
                        .Row = .MaxRows
                    Else
                        .Row = .DataRowCnt + 1
                    End If
                    
                    .Col = TReadyEnum.ccNo:     .Text = CStr(lngWorkNo)
                    .Col = TReadyEnum.ccBarNo:  .Text = gBarNo 'pAccInfo.GetBarNo
                    .Col = TReadyEnum.ccAccNo:  .Text = mGetAccNo(pAccInfo.Workarea, pAccInfo.AccDt, pAccInfo.AccSeq)
                    
                    If pAccInfo.QcFg = "0" Then         '## �Ϲݰ�ü
                        .Col = TReadyEnum.ccPtId:   .Text = pAccInfo.PtId
                        .Col = TReadyEnum.ccName:   .Text = pAccInfo.Name
                    ElseIf pAccInfo.QcFg = "1" Then     '## QC��ü
                        .Col = TReadyEnum.ccPtId:   .Text = pAccInfo.CtrlCd
                        .Col = TReadyEnum.ccName:   .Text = pAccInfo.LevelCd
                    End If
                    
                    .Col = 6:  .Text = pAccInfo.Sex & "|" & pAccInfo.Ssn
                    .Col = TReadyEnum.ccNo:     .Text = "FD"
                    Exit For
                End If
            Next
        End If

        
        Call .SetActiveCell(1, .DataRowCnt)
    End With
End Sub

'-----------------------------------------------------------------------------'
'   ��� : tblComplete�� ����ǥ�� (���������� ������)
'   �μ� :
'       - pIntInfo : �������̽� ��ü���� Ŭ����
'-----------------------------------------------------------------------------'
Private Sub SetComplete1(ByVal pIntInfo As clsIISIntInfo)
    Dim objIntResult As clsIISIntResult     '�������̽� ��� Ŭ����
    Dim i            As Long
    
    With tblComplete
        If .MaxRows <= .DataRowCnt Then
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
        Else
            .Row = .DataRowCnt + 1
        End If

        .Col = TCompleteEnum.ccNo:      .Text = pIntInfo.SpcPos
        .Col = TCompleteEnum.ccBarNo:   .Text = pIntInfo.BarNo
        .Col = TCompleteEnum.ccSendCnt: .Text = pIntInfo.IntResults.Count
        
        For Each objIntResult In pIntInfo.IntResults
            If .MaxCols <= .DataColCnt Then
                .MaxCols = .MaxCols + 1
            End If
            .Col = TCompleteEnum.ccResult + i
            .ColHidden = True
            
            '## 1.0.4: �̻��(2005-06-29)
            '   - ��������� LIS����� ǥ�õǴ� ���׼���
            .Text = objIntResult.IntNm & DIV & objIntResult.IntResult & DIV & DIV & DIV & DIV & DIV & DIV & objIntResult.Info
            i = i + 1
        Next
        Set objIntResult = Nothing
        
        Call .SetActiveCell(TCompleteEnum.ccNo, .DataRowCnt)
    End With
End Sub

'-----------------------------------------------------------------------------'
'   ��� : tblComplete�� ����ǥ�� (���������� ������)
'   �μ� :
'       - pAccInfo : �������� Ŭ����
'-----------------------------------------------------------------------------'
Private Sub SetComplete2(ByVal pAccInfo As clsIISAccInfo)
    Dim objResult   As clsIISResult     '������� Ŭ����
    Dim objQCResult As clsIISQCResult   'QC������� Ŭ����
    Dim i           As Long
    
    With tblComplete
        If .MaxRows <= .DataRowCnt Then
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
        Else
            .Row = .DataRowCnt + 1
        End If

        .Col = TCompleteEnum.ccNo:      .Text = pAccInfo.SpcPos
        .Col = TCompleteEnum.ccBarNo:   .Text = pAccInfo.GetBarNo
        .Col = TCompleteEnum.ccAccNo:   .Text = mGetAccNo(pAccInfo.Workarea, pAccInfo.AccDt, pAccInfo.AccSeq)
        
        If pAccInfo.QcFg = "0" Then         '## �Ϲݰ�ü
            .Col = TCompleteEnum.ccPtId:    .Text = pAccInfo.PtId
            .Col = TCompleteEnum.ccName:    .Text = pAccInfo.Name
            .Col = TCompleteEnum.ccSexAge:  .Text = pAccInfo.Sex & " / " & mGetAge(Mid$(pAccInfo.Ssn, 1, 6))
            .Col = TCompleteEnum.ccDoctNm:  .Text = pAccInfo.OrdDoctNm
            .Col = TCompleteEnum.ccDeptNm:  .Text = pAccInfo.DeptNm
            .Col = TCompleteEnum.ccWardNm:  .Text = pAccInfo.WardNm
            .Col = TCompleteEnum.ccStatFg:  .Text = IIf(pAccInfo.StatFg = "1", "Y", "N")
            .Col = TCompleteEnum.ccSpcNm:   .Text = pAccInfo.SpcNm
            .Col = TCompleteEnum.ccQcFg:    .Text = pAccInfo.QcFg
            .Col = TCompleteEnum.ccSendCnt: .Text = pAccInfo.SendCnt
            
            For Each objResult In pAccInfo.Results
                If .MaxCols <= .DataColCnt Then
                    .MaxCols = .MaxCols + 1
                End If
                .Col = TCompleteEnum.ccResult + i
                .ColHidden = True
                If objResult.IntResult <> "" Then
                    .Text = objResult.IntNm.IntNm & DIV & objResult.IntResult & DIV & objResult.RstCd & _
                            DIV & objResult.Unit & DIV & objResult.HLDiv & DIV & objResult.DPDiv & _
                            DIV & mGetRef(objResult.Ref.RefFrVal, objResult.Ref.RefToVal) & DIV & DIV & objResult.IntNm.IntBase & DIV
                            '-- 2015.08.28�߰� & objResult.IntNm.IntBase & DIV
                    i = i + 1
                End If
                If Mid(objResult.IntNm.IntNm, 1, 3) = "(I)" Then
                    .Col = TCompleteEnum.ccNo:      .Text = "IN"
                Else
                    .Col = TCompleteEnum.ccNo:      .Text = "FD"
                End If
            Next
            Set objResult = Nothing
        ElseIf pAccInfo.QcFg = "1" Then     '## QC��ü
            .Col = TCompleteEnum.ccPtId:    .Text = pAccInfo.CtrlCd
            .Col = TCompleteEnum.ccName:    .Text = pAccInfo.LevelCd
            .Col = TCompleteEnum.ccSexAge:  .Text = pAccInfo.LotNo
            .Col = TCompleteEnum.ccQcFg:    .Text = pAccInfo.QcFg
            .Col = TCompleteEnum.ccSendCnt: .Text = pAccInfo.SendCnt
            
            For Each objQCResult In pAccInfo.QCResults
                If .MaxCols <= .DataColCnt Then
                    .MaxCols = .MaxCols + 1
                End If
                .Col = TCompleteEnum.ccResult + i
                .ColHidden = True
                
                .Text = objQCResult.IntNm.IntNm & DIV & objQCResult.IntResult & DIV & _
                        objQCResult.RstCd & DIV & objQCResult.Unit & DIV & objQCResult.RADiv & _
                        DIV & DIV & DIV
                i = i + 1
            Next
            Set objQCResult = Nothing
        End If
        Call .SetActiveCell(TCompleteEnum.ccNo, .DataRowCnt)
    End With
End Sub

'-----------------------------------------------------------------------------'
'   ��� : tblResult ����ǥ��
'   �μ� :
'       - pAccInfo : �������� Ŭ����
'-----------------------------------------------------------------------------'
Private Sub SetResult(ByVal pAccInfo As clsIISAccInfo)
    Dim objResult   As clsIISResult     '������� Ŭ����
    Dim objQCResult As clsIISQCResult   'QC������� Ŭ����
    
    Call mTblClear(tblResult)
    If pAccInfo.QcFg = "0" Then         '## �Ϲݰ�ü
        For Each objResult In pAccInfo.Results
            With tblResult
                If .MaxRows <= .DataRowCnt Then
                    .MaxRows = .MaxRows + 1
                    .Row = .MaxRows
                Else
                    .Row = .DataRowCnt + 1
                End If
                
                .Col = TResultEnum.ccTestNm:    .Text = objResult.IntNm.IntNm
                .Col = TResultEnum.ccEqpResult: .Text = objResult.Result
                .Col = TResultEnum.ccLISResult: .Text = objResult.RstCd
                .Col = TResultEnum.ccUnit:      .Text = objResult.Unit
                .Col = TResultEnum.ccHLDiv:     .Text = objResult.HLDiv
                .Col = TResultEnum.ccDPDiv:     .Text = objResult.DPDiv
                .Col = TResultEnum.ccRef:       .Text = mGetRef(objResult.Ref.RefFrVal, objResult.Ref.RefToVal)
            End With
        Next
    ElseIf pAccInfo.QcFg = "1" Then     '## QC��ü
        For Each objQCResult In pAccInfo.QCResults
            With tblResult
                If .MaxRows <= .DataRowCnt Then
                    .MaxRows = .MaxRows + 1
                    .Row = .MaxRows
                Else
                    .Row = .DataRowCnt + 1
                End If
                
                .Col = TResultEnum.ccTestNm:    .Text = objQCResult.IntNm.IntNm
                .Col = TResultEnum.ccEqpResult: .Text = objQCResult.Result
                .Col = TResultEnum.ccLISResult: .Text = objQCResult.RstCd
                .Col = TResultEnum.ccHLDiv:     .Text = objQCResult.RADiv
            End With
        Next
    End If
    
    Set objResult = Nothing
    Set objQCResult = Nothing
End Sub

'-----------------------------------------------------------------------------'
'   ��� : AlloScan ��ũ����Ʈ ���� ����ǥ��
'   �μ� :
'       - pAccInfo : �������� Ŭ����
'-----------------------------------------------------------------------------'
Private Sub SetOrderWS(ByVal pAccInfo As clsIISAccInfo)
    Dim objResult   As clsIISResult     '������� Ŭ����
    Dim objQCResult As clsIISQCResult   'QC������� Ŭ����
    Dim mLogOn As clsIISLogOn
    
    Dim strAlloFile As String
    Dim lngFIleNum  As Long
    Dim strInFo     As String
    Dim strOldInFo  As String
    Dim blnNewFlag  As Boolean
    
    Set mLogOn = New clsIISLogOn
    
    
'    If pAccInfo.QcFg = "0" Then         '## �Ϲݰ�ü
            With AlloFile
                .CancelError = True
                .FileName = "C:\RAPID\IMPORT\" & Trim(txtFileNm.Text) & ".asc"
'                .ShowSave
                If Len(Dir(.FileName)) Then
                    Kill .FileName
                    blnNewFlag = True
                Else
                    blnNewFlag = False
                End If
                
                lngFIleNum = FreeFile
                
                Open .FileName For Append As #lngFIleNum
                If blnNewFlag = False Then
'[JOBLIST]
'JOBName;1/01/2009;S01-45329578234;Panel 2KO v80 UK.RDF.TST;Last01;First01;1/2/2008;MALE;
'JOBName;1/01/2009;S02-64325984359;Panel 2KO v80 UK.RDF.TST;Last01;First01;12/31/2007;MALE;
'JOBName;1/01/2009;S03-76534954562;Panel 2KO v80 UK.RDF.TST;Last03;First03;1/2/2008;MALE;
'JOBName;1/01/2009;S04-64632134814;Panel 2KO v80 UK.RDF.TST;Last04;First04;1/2/2008;MALE;
'JOBName;1/01/2009;S05-16324873488;Panel 2KO v80 UK.RDF.TST;Last05;First05;1/2/2008;MALE;
'[EOF]
                    
                    
                    Print #lngFIleNum, "[JOBLIST]"
'                    Print #lngFIleNum, "JOBName;" & Format("m/dd/yyyy") & ";" & pAccInfo.GetBarNo & ";" & "Panel 2KO v80 UK.RDF.TST"
                End If

                For Each objResult In pAccInfo.Results
                    If InStr(objResult.TestNm, "Food") > 0 Then
                        strInFo = "1"
                    ElseIf InStr(objResult.TestNm, "Inhalant") > 0 Then
                        strInFo = "2"
                    End If

                    If strOldInFo <> strInFo Then
                        If strInFo = "1" Then
                            Print #lngFIleNum, "JOBName;" & Format(Now, "m/dd/yyyy") & ";" & pAccInfo.GetBarNo & ";Panel 1KO v80 UK.TST"
                            Print #lngFIleNum, "JOBName;" & Format(Now, "m/dd/yyyy") & ";" & pAccInfo.GetBarNo & ";Panel 2KO v80 UK.TST"
                        Else
                            Print #lngFIleNum, "JOBName;" & Format(Now, "m/dd/yyyy") & ";" & pAccInfo.GetBarNo & ";Panel 3KO v80 UK.TST"
                            Print #lngFIleNum, "JOBName;" & Format(Now, "m/dd/yyyy") & ";" & pAccInfo.GetBarNo & ";Panel 4KO v80 UK.TST"
                        End If
                    End If

                    strOldInFo = strInFo
                Next
            End With
            
        Print #lngFIleNum, "[EOF]"
        Close #lngFIleNum

'    ElseIf pAccInfo.QcFg = "1" Then     '## QC��ü
'        For Each objQCResult In pAccInfo.QCResults
'            With tblResult
'                If .MaxRows <= .DataRowCnt Then
'                    .MaxRows = .MaxRows + 1
'                    .Row = .MaxRows
'                Else
'                    .Row = .DataRowCnt + 1
'                End If
'
'                .Col = TResultEnum.ccTestNm:    .Text = objQCResult.IntNm.IntNm
'                .Col = TResultEnum.ccEqpResult: .Text = objQCResult.Result
'                .Col = TResultEnum.ccLISResult: .Text = objQCResult.RstCd
'                .Col = TResultEnum.ccHLDiv:     .Text = objQCResult.RADiv
'            End With
'        Next
'    End If
    
    Set objResult = Nothing
    Set objQCResult = Nothing
End Sub

'-----------------------------------------------------------------------------'
'   ��� : Label�� ȯ������, �������� ǥ��
'   �μ� :
'       - pAccInfo : �������� Ŭ����
'-----------------------------------------------------------------------------'
Private Sub SetLabel(ByVal pAccInfo As clsIISAccInfo)
    Call CtlClear(ccLabel)
    
    If pAccInfo.QcFg = "0" Then         '## �Ϲݰ�ü
        Call LabelShow("0")
        lblPtId.Caption = pAccInfo.PtId
        lblName.Caption = pAccInfo.Name
        lblSexAge.Caption = pAccInfo.Sex & " / " & mGetAge(Mid$(pAccInfo.Ssn, 1, 6))
        lblDoctNm.Caption = pAccInfo.OrdDoctNm
        lblDeptNm.Caption = pAccInfo.DeptNm
        lblWardNm.Caption = pAccInfo.WardNm
        lblStatFg.Caption = IIf(pAccInfo.StatFg = "1", "Y", "N")
        lblSpcNm.Caption = pAccInfo.SpcNm
        'lblPnlNm.Caption = ""
    ElseIf pAccInfo.QcFg = "1" Then     '## QC��ü
        Call LabelShow("1")
        lblPtId.Caption = pAccInfo.CtrlCd
        lblName.Caption = pAccInfo.LevelCd
        lblSexAge.Caption = pAccInfo.LotNo
    End If
End Sub

'-----------------------------------------------------------------------------'
'   ��� : ��ü������ ���� Label�� �ٸ��� ǥ��
'   �μ� :
'       - pQcFg : 0(�Ϲݰ�ü), 1(QC��ü)
'-----------------------------------------------------------------------------'
Private Sub LabelShow(ByVal pQcFg As String)
    Dim i As Long
    
    If pQcFg = "0" Then         '## �Ϲݰ�ü
        lblControl.Caption = "ȯ  �� ID :"
        lblLevel.Caption = "��     �� :"
        lblLotNo.Caption = "����/���� :"
        For i = 0 To lblGeneral.Count - 1
            lblGeneral(i).Visible = True
        Next i
        
        lblDoctNm.Visible = True:   lblDeptNm.Visible = True
        lblWardNm.Visible = True:   lblStatFg.Visible = True
        lblSpcNm.Visible = True
    ElseIf pQcFg = "1" Then     '## QC��ü
        lblControl.Caption = "Control :"
        lblLevel.Caption = "Level   :"
        lblLotNo.Caption = "Lot No  :"
        For i = 0 To lblGeneral.Count - 1
            lblGeneral(i).Visible = False
        Next i
        
        lblDoctNm.Visible = False:   lblDeptNm.Visible = False
        lblWardNm.Visible = False:   lblStatFg.Visible = False
        lblSpcNm.Visible = False
    End If
End Sub

'-----------------------------------------------------------------------------'
'   ��� : ����� ������ȸ, ��Ʈ Open
'-----------------------------------------------------------------------------'
Private Sub GetEqpComm()
    Dim objComm     As clsIISEqpComm    '��ż��� Ŭ����
    Dim strErrMsg   As String           '�����޽���
    
    '## ��ż��� ������ȸ
    Set objComm = mIntLib.GetEqpComm
    If objComm Is Nothing Then Exit Sub
    
    With objComm
        MSComm.CommPort = .Port
        MSComm.Settings = .GetSettings
    End With
    Set objComm = Nothing

On Error GoTo Errors
    '## ��Ʈ Open
    With MSComm
        '## �̹� ��Ʈ�� �������
        If .PortOpen Then
            strErrMsg = mEqpCd & " ����� �����Ʈ�� �̹� �����ֽ��ϴ�."
            Error.SetLog App.EXEName, "frmIISABL835", "GetEqpComm", strErrMsg, Now
            Call mIntLib_EqpError("E004")
            Exit Sub
        End If
        
        .RThreshold = 1
        .SThreshold = 1
        .RTSEnable = True
        .PortOpen = True
    End With
    
    '## �������� ���������� ����
    Call mIntLib.DelHistoryData
    Exit Sub
    
Errors:
    '## �ٸ� ��ġ���� ��Ʈ�� ����ϴ� ���
    If Err.Number = 8005 Then
        strErrMsg = mEqpCd & " ��� ������ ��Ʈ�� �̹� ������Դϴ�."
        Error.SetLog App.EXEName, "frmIISABL835", "GetEqpComm", strErrMsg, Now
        Call mIntLib_EqpError("E005")
    End If
End Sub

'-----------------------------------------------------------------------------'
'   ��� : ��Ʈ�� �ʱ�ȭ
'-----------------------------------------------------------------------------'
Private Sub CtlClear(Optional ByVal pFlag As ClearEnum = ccAll)
    lblPtId.Caption = "":       lblName.Caption = ""
    lblSexAge.Caption = "":     lblDoctNm.Caption = ""
    lblDeptNm.Caption = "":     lblWardNm.Caption = ""
    lblStatFg.Caption = "":     lblSpcNm.Caption = ""
    lblPnlNm.Caption = ""
    If pFlag = ccAll Then
        txtBarNo.Text = "":         Call mTblClear(tblResult)
        Call mTblClear(tblReady):   Call mTblClear(tblComplete)
    End If
End Sub

'------------------------------------------------------------------'
'   ��� : ����� ���� ����ó��
'------------------------------------------------------------------'
Private Sub mIntLib_EqpError(ByVal pCode As String)
    cmdAlarm.BackColor = vbRed
    Call mIntErrors.Add(pCode, mEqpCd, mEqpKey)
End Sub

'------------------------------------------------------------------'
'   ��� : ��ü���� ����ó��1
'------------------------------------------------------------------'
Private Sub mIntLib_SpcError(ByVal pCode As String, ByVal pBarNo As String)
    cmdAlarm.BackColor = vbRed
    Call mIntErrors.Add(pCode, mEqpCd, mEqpKey, pBarNo)
End Sub

'------------------------------------------------------------------'
'   ��� : ��ü���� ����ó��2
'------------------------------------------------------------------'
Private Sub mIntLib_SpcErrorX(ByVal pCode As String, ByVal pBarNo As String, ByVal pPtId As String, ByVal pName As String)
    cmdAlarm.BackColor = vbRed
    Call mIntErrors.Add(pCode, mEqpCd, mEqpKey, pBarNo, pPtId, pName)
End Sub

'------------------------------------------------------------------'
'   ��� : Popup �޴� Click �̺�Ʈ
'------------------------------------------------------------------'
Private Sub mPopup_Click(ByVal vMenuID As Long)
    Dim vBarNo      As Variant  'Spread�� ���ڵ��ȣ
    Dim strSpcYy    As String   '��ü����
    Dim lngSpcNo    As Long     '��ü��ȣ
    
    Select Case vMenuID
        Case DELETE     '## Delete
            With tblReady
                Call .GetText(TReadyEnum.ccBarNo, .ActiveRow, vBarNo)
                If vBarNo <> "" Then
                    strSpcYy = Mid$(vBarNo, 1, SPCYYLEN)
                    lngSpcNo = CLng(Mid$(vBarNo, SPCYYLEN + 1, SPCNOLEN))
                    Call mIntLib.AccInfos.Remove(strSpcYy, lngSpcNo)
                    Call .DeleteRows(.ActiveRow, 1)
                End If
            End With
        Case DELETEALL  '## Delete All
            Call mIntLib.AccInfos.RemoveAll
            Call mTblClear(tblReady)
    End Select
End Sub

Private Sub txtFileNm_DblClick()
    If vasINPrint.Visible = True Then
        vasINPrint.Visible = False
    Else
        vasINPrint.Visible = True
    End If
End Sub

Private Sub txtWorkNo_GotFocus()
    With txtWorkNo
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtWorkNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtWorkNo_KeyPress(KeyAscii As Integer)
    '## ���ڸ� �Է�
    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub
