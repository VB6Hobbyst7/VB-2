VERSION 5.00
Object = "{9167B9A7-D5FA-11D2-86CA-00104BD5476F}#5.0#0"; "DRctl1.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmBBS303N 
   BackColor       =   &H00DBE6E6&
   Caption         =   "Blood Batch Delivery"
   ClientHeight    =   9285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14655
   Icon            =   "frmBBS303N.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9285
   ScaleWidth      =   14655
   WindowState     =   2  '�ִ�ȭ
   Begin MedControls1.LisLabel LisLabel5 
      Height          =   225
      Left            =   60
      TabIndex        =   44
      Top             =   15
      Width           =   4740
      _ExtentX        =   8361
      _ExtentY        =   397
      BackColor       =   8388608
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "ȯ������"
      Appearance      =   0
   End
   Begin VB.Frame Frame2 
      Height          =   900
      Left            =   45
      TabIndex        =   45
      Top             =   135
      Width           =   4770
      Begin VB.TextBox txtCompID 
         Appearance      =   0  '���
         Height          =   360
         Left            =   1095
         TabIndex        =   46
         Top             =   165
         Width           =   1125
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   7
         Left            =   45
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   165
         Width           =   1035
         _ExtentX        =   1826
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
         Caption         =   "���ȯ��"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblCompNm 
         Height          =   360
         Left            =   2235
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   165
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   635
         BackColor       =   14411494
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
         Caption         =   "������"
         Appearance      =   0
      End
      Begin DRcontrol1.DrLabel lblSexAge 
         Height          =   300
         Left            =   1095
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   540
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   529
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
         Border          =   1
         Caption         =   ""
      End
      Begin VB.Label lblABO 
         Alignment       =   2  '��� ����
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "AB(AB)+"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   3450
         TabIndex        =   50
         Top             =   360
         Width           =   1170
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  '���� ����
         Height          =   675
         Left            =   3315
         TabIndex        =   51
         Top             =   165
         Width           =   1380
      End
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   225
      Left            =   4860
      TabIndex        =   29
      Top             =   15
      Width           =   4170
      _ExtentX        =   7355
      _ExtentY        =   397
      BackColor       =   8388608
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "���׹�ȣ"
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   885
      Left            =   4860
      TabIndex        =   30
      Top             =   150
      Width           =   4185
      Begin VB.CommandButton cmdBldNo 
         BackColor       =   &H00DEDBDD&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   3765
         MousePointer    =   14  'ȭ��ǥ�� ����ǥ
         Style           =   1  '�׷���
         TabIndex        =   34
         Top             =   150
         Width           =   350
      End
      Begin VB.CheckBox chkBarCode 
         BackColor       =   &H00DBE6E6&
         Caption         =   "���ڵ��Է�"
         Height          =   195
         Left            =   60
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   225
         Width           =   1575
      End
      Begin VB.TextBox txtBldNo 
         Appearance      =   0  '���
         Height          =   345
         Left            =   1650
         TabIndex        =   32
         Top             =   150
         Width           =   2115
      End
      Begin VB.ComboBox cboCompo 
         Height          =   300
         Left            =   1650
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   510
         Width           =   2475
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   0
         Left            =   60
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   510
         Width           =   1560
         _ExtentX        =   2752
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
         Caption         =   "Component"
         Appearance      =   0
      End
   End
   Begin VB.TextBox txtRcvId 
      Appearance      =   0  '���
      Height          =   360
      Left            =   5940
      TabIndex        =   28
      Top             =   1035
      Width           =   1110
   End
   Begin VB.CommandButton cmdRcvId 
      BackColor       =   &H00DEDBDD&
      Caption         =   "..."
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
      Left            =   7050
      MousePointer    =   14  'ȭ��ǥ�� ����ǥ
      Style           =   1  '�׷���
      TabIndex        =   27
      Top             =   1035
      Width           =   350
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "����(&X)"
      CausesValidation=   0   'False
      Height          =   510
      Left            =   13140
      Style           =   1  '�׷���
      TabIndex        =   26
      Tag             =   "128"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "ȭ������(&C)"
      CausesValidation=   0   'False
      Height          =   510
      Left            =   11820
      Style           =   1  '�׷���
      TabIndex        =   25
      Tag             =   "124"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "���(&S)"
      Height          =   510
      Left            =   9060
      Style           =   1  '�׷���
      TabIndex        =   24
      Tag             =   "15101"
      Top             =   885
      Width           =   1320
   End
   Begin VB.CommandButton cmdReprint 
      BackColor       =   &H00F4F0F2&
      Caption         =   "�����(&P)"
      Height          =   510
      Left            =   11700
      Style           =   1  '�׷���
      TabIndex        =   23
      Tag             =   "15101"
      Top             =   885
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.CommandButton cmdTransStatPaper 
      BackColor       =   &H00F4F0F2&
      Caption         =   "���޿�û��"
      Height          =   510
      Left            =   13020
      Style           =   1  '�׷���
      TabIndex        =   22
      Tag             =   "15101"
      Top             =   885
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.CommandButton cmdFOpen 
      BackColor       =   &H00F4F0F2&
      Caption         =   "Filter ���"
      Height          =   510
      Left            =   10380
      Style           =   1  '�׷���
      TabIndex        =   21
      Tag             =   "15101"
      Top             =   885
      Width           =   1320
   End
   Begin VB.Frame fraFilter 
      BackColor       =   &H00F4F0F2&
      Caption         =   "Frame2"
      Height          =   5115
      Left            =   4590
      TabIndex        =   0
      Top             =   1815
      Visible         =   0   'False
      Width           =   6615
      Begin VB.TextBox txtPtid 
         Appearance      =   0  '���
         Height          =   360
         Left            =   1470
         TabIndex        =   6
         Top             =   720
         Width           =   1200
      End
      Begin VB.TextBox txtQty 
         Appearance      =   0  '���
         Height          =   360
         Left            =   1470
         TabIndex        =   5
         Top             =   1095
         Width           =   1200
      End
      Begin VB.CommandButton cmdFSave 
         BackColor       =   &H00F4F0F2&
         Caption         =   "�������"
         Height          =   510
         Left            =   1620
         Style           =   1  '�׷���
         TabIndex        =   3
         Tag             =   "124"
         Top             =   4500
         Width           =   1320
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00F4F0F2&
         Caption         =   "�ݱ�"
         Height          =   510
         Left            =   2940
         Style           =   1  '�׷���
         TabIndex        =   2
         Tag             =   "128"
         Top             =   4500
         Width           =   1320
      End
      Begin VB.CommandButton cmdFQuery 
         BackColor       =   &H00F4F0F2&
         Caption         =   "��ȸ(&Q)"
         Height          =   1125
         Left            =   4920
         Style           =   1  '�׷���
         TabIndex        =   1
         Tag             =   "124"
         Top             =   705
         Width           =   1650
      End
      Begin MSComCtl2.DTPicker dtpFromDate 
         Height          =   375
         Left            =   1485
         TabIndex        =   4
         Top             =   345
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   661
         _Version        =   393216
         Format          =   24576001
         CurrentDate     =   38170
      End
      Begin MedControls1.LisLabel LisLabel3 
         Height          =   315
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   6600
         _ExtentX        =   11642
         _ExtentY        =   556
         BackColor       =   8388608
         ForeColor       =   -2147483634
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   "FILTER ���"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   1
         Left            =   45
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   720
         Width           =   1410
         _ExtentX        =   2487
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
         Caption         =   "ȯ��ID"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   2
         Left            =   45
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1095
         Width           =   1410
         _ExtentX        =   2487
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
         Caption         =   "Fileter����"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblPtNm 
         Height          =   360
         Left            =   2685
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   720
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   635
         BackColor       =   14411494
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
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   3
         Left            =   2685
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1095
         Width           =   1020
         _ExtentX        =   1799
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
         Caption         =   "������"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblDelCnt 
         Height          =   360
         Left            =   3720
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1095
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   635
         BackColor       =   14411494
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
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   4
         Left            =   45
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1470
         Width           =   1410
         _ExtentX        =   2487
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
         Caption         =   "��ȯ����"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   5
         Left            =   2685
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1470
         Width           =   1020
         _ExtentX        =   1799
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
         Caption         =   "������"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblExpCnt 
         Height          =   360
         Left            =   3720
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1470
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   635
         BackColor       =   14411494
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
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   360
         Index           =   6
         Left            =   45
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   345
         Width           =   1410
         _ExtentX        =   2487
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
         Caption         =   "��ȸ�Ⱓ"
         Appearance      =   0
      End
      Begin MSComCtl2.DTPicker dtpToDate 
         Height          =   375
         Left            =   2985
         TabIndex        =   17
         Top             =   345
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   661
         _Version        =   393216
         Format          =   24576001
         CurrentDate     =   38170
      End
      Begin FPSpread.vaSpread tblFilter 
         Height          =   2580
         Left            =   30
         TabIndex        =   18
         Top             =   1845
         Width           =   6540
         _Version        =   196608
         _ExtentX        =   11536
         _ExtentY        =   4551
         _StockProps     =   64
         BackColorStyle  =   1
         ColHeaderDisplay=   0
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   15265518
         GridColor       =   16703181
         GridShowVert    =   0   'False
         MaxCols         =   11
         MaxRows         =   7
         ScrollBars      =   2
         ShadowColor     =   14737632
         ShadowDark      =   14737632
         ShadowText      =   0
         SpreadDesigner  =   "frmBBS303N.frx":076A
         TextTip         =   2
      End
      Begin MedControls1.LisLabel lblRetCnt 
         Height          =   360
         Left            =   1470
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1470
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   635
         BackColor       =   14411494
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
      End
      Begin VB.Label Label1 
         BackColor       =   &H00F4F0F2&
         Caption         =   "~"
         Height          =   255
         Left            =   2805
         TabIndex        =   20
         Top             =   420
         Width           =   360
      End
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   360
      Left            =   75
      TabIndex        =   36
      Top             =   1425
      Width           =   14385
      _ExtentX        =   25374
      _ExtentY        =   635
      BackColor       =   8388608
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "���� ����Ʈ"
      Appearance      =   0
   End
   Begin FPSpread.vaSpread tblBldList 
      Height          =   6615
      Left            =   75
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   1800
      Width           =   14385
      _Version        =   196608
      _ExtentX        =   25374
      _ExtentY        =   11668
      _StockProps     =   64
      BackColorStyle  =   1
      ButtonDrawMode  =   1
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
      MaxCols         =   32
      MaxRows         =   20
      ScrollBars      =   2
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      SpreadDesigner  =   "frmBBS303N.frx":0CB4
      TextTip         =   4
   End
   Begin MedControls1.LisLabel lblRcvNm 
      Height          =   360
      Left            =   7425
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   1035
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   635
      BackColor       =   14411494
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
   End
   Begin MedControls1.LisLabel lblCompoNm 
      Height          =   360
      Left            =   12015
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   635
      BackColor       =   14411494
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
      Appearance      =   0
   End
   Begin MedControls1.LisLabel lblCompoCd 
      Height          =   360
      Left            =   13785
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   635
      BackColor       =   14411494
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
      Appearance      =   0
   End
   Begin MedControls1.LisLabel lblDeliveryNm 
      Height          =   360
      Left            =   1125
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   1035
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   635
      BackColor       =   14411494
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
      Caption         =   "������"
      Appearance      =   0
   End
   Begin Crystal.CrystalReport CReport 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   360
      Index           =   17
      Left            =   4860
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   1035
      Width           =   1035
      _ExtentX        =   1826
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
      Caption         =   "������ "
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   360
      Index           =   18
      Left            =   45
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   1035
      Width           =   1035
      _ExtentX        =   1826
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
      Caption         =   "�����"
      Appearance      =   0
   End
End
Attribute VB_Name = "frmBBS303N"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const mode = 0
Private Enum TblColumn
    tcCOMPONM = 1
    tcBldNo
    tcABO
    tcVol
    tcAVAIL
    
    tcFilter
    tcRT
    tcDELDT
    tcPTID
    tcPTNM
    
    tcRMK
    tcABOP
    tcRTP
    tcORDDT
    tcTEST
    
    tcUNIT
    tcRCVNM
    tcCOMPOCD
    tcORDNO
    tcORDSEQ
    
    tcACCDT
    tcACCSEQ
    tcRSTSEQ
    tcORDCD
    tcNewTest
    
    tcDUP
    tcRmk2
    tcDEPT
    tcSEX
    tcPtRmk
    
    tcBuss
    tcOcsOrdno
End Enum

'Private WithEvents mnuPopup As Menu
'Private WithEvents mnuDelete As Menu
'Private WithEvents mnuSave  As Menu
Private WithEvents objPop As clsPopupMenu
Attribute objPop.VB_VarHelpID = -1
Private Const MENU_DEL& = 1
Private Const MENU_SAVE& = 2
Private Const MENU_SEP& = 99
Private WithEvents objMyList As clsPopUpList
Attribute objMyList.VB_VarHelpID = -1
Private WithEvents objBldList As clsPopUpList
Attribute objBldList.VB_VarHelpID = -1

Private sABOType    As String
Private sRHType     As String


Private Sub cmdBldNo_Click()
'    With frmBloodFind
'        .mode = mode
'
'        .Show vbModal
'        If .isSelected = True Then
'            If chkBarCode.value = 1 Then
'                txtBldNo.Text = .BldSrc & .BldYY & Format(.BldNo, "00000#") '& "10"
'            Else
'                txtBldNo.Text = .BldSrc & "-" & .BldYY & "-" & Format(.BldNo, "00000#")
'            End If
'            txtBldNoLostFocus
'            cboCompo.ListIndex = medComboFind(cboCompo, .Compo)
'        End If
'    End With
'    Set frmBloodFind = Nothing

'iTmx.Text = .Fields("bldsrc").value & "-" & .Fields("bldyy").value & "-" & .Fields("bldno").value
'iTmx.SubItems(1) = .Fields("ptid").value
'iTmx.SubItems(2) = GetPtNm(.Fields("ptid").value & "")
'iTmx.SubItems(3) = .Fields("compocd").value & " " & .Fields("componm").value


    Dim objSql As clsBldDelivery
    Dim strSQL As String
    
    Set objSql = New clsBldDelivery
    
    Set objBldList = New clsPopUpList
    With objBldList
        .Connection = DBConn
        .ColumnHeaderText = "���׹�ȣ;ȯ��ID;ȯ�ڸ�;����"
        .ColumnHeaderWidth = "1230.236;915.0237;929.7639;2085.166"
        .FormHeight = 4125
        .FormWidth = 5715
        .FormCaption = "��� ��� ����Ʈ"
        .SortColumn = 3
        .ColumnHeaderAlign = "0;2;2;0"
        .SqlStmt = objSql.GetSQLBloodList(mode)
        .AutoGap = True
        .LoadPopUp
    End With
    
    Set objBldList = Nothing
    Set objSql = Nothing
End Sub

Private Sub cmdClose_Click()
    fraFilter.Visible = False
End Sub

Private Sub cmdFOpen_Click()
    Call FilterClear
    
    fraFilter.Visible = True
    fraFilter.ZOrder 0
    txtPtId.SetFocus
End Sub

Private Sub FilterClear()
    txtPtId.Text = "": lblDelCnt.Caption = "": lblRetCnt.Caption = "": lblExpCnt.Caption = ""
    txtQty.Text = ""
    lblPtNm.Caption = ""
    tblFilter.MaxRows = 0
    dtpFromDate.value = DateAdd("d", -3, GetSystemDate)
    dtpToDate.value = GetSystemDate
End Sub

'Filteró����ȸ
Private Sub cmdFQuery_Click()
    Dim RS          As Recordset
    Dim strPtid     As String
    Dim strFDate    As String
    Dim strTDate    As String
    Dim strFilter   As String
    Dim SSQL        As String
    Dim ii As Integer
    
    strPtid = txtPtId.Text
    If strPtid = "" Then
        MsgBox "ȯ��ID�� �Է��� ��ȸ�ϼ���.", vbInformation + vbOKOnly, "Info"
        Exit Sub
    End If
    
    strFDate = Format(dtpFromDate.value, PRESENTDATE_FORMAT)
    strTDate = Format(dtpToDate.value, PRESENTDATE_FORMAT)
    
    SSQL = "SELECT  CDVAL1 FROM " & T_COM003 & " WHERE " & DBW("CDINDEX=", BC2_MATERIAL)
    
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    If Not RS.EOF Then
        Do Until RS.EOF
            strFilter = strFilter & "'" & RS.Fields("cdval1").value & "" & "',"
            RS.MoveNext
        Loop
        If strFilter <> "" Then
            strFilter = Mid(strFilter, 1, Len(strFilter) - 1)
        Else
            MsgBox "FILTER CODE �� �������� �ʾҽ��ϴ�.", vbInformation + vbOKOnly, "Info"
            Set RS = Nothing
            Exit Sub
        End If
    Else
        MsgBox "FILTER CODE �� �������� �ʾҽ��ϴ�.", vbInformation + vbOKOnly, "Info"
        Set RS = Nothing
        Exit Sub
    End If
    
    SSQL = " SELECT  b.ptid,b.orddt,b.ordno,b.ordseq,b.ordcd,b.ocsordno,c.testnm,b.unitqty,a.bussdiv,a.deptcd " & _
           " FROM " & T_BBS001 & " c," & T_LAB102 & " b," & T_LAB101 & " a" & _
           " WHERE " & DBW("a.ptid=", strPtid) & _
           " AND " & DBW("a.orddt>=", strFDate) & _
           " AND " & DBW("a.orddt<=", strTDate) & _
           " AND " & DBW("a.orddiv=", C_WORKAREA) & _
           " AND " & DBW("a.donefg>=", enStsCd.StsCd_LIS_Accession) & _
           " AND a.ptid=b.ptid and a.orddt=b.orddt and a.ordno=b.ordno " & _
           " AND b.ordcd in(" & strFilter & ") " & _
           " AND b.ordcd=c.testcd " & _
           " AND (c.expdt='' or c.expdt is null)" & _
            " ORDER BY orddt desc,testcd"
    
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    With tblFilter
        .MaxRows = 0
        If Not RS.EOF Then
            Do Until RS.EOF
                If .DataRowCnt + 1 > .MaxRows Then .MaxRows = .MaxRows + 1
                .Row = .DataRowCnt + 1
                .RowHeight(.Row) = 13.3
                .Col = 1: .CellType = CellTypeCheckBox
                          .TypeHAlign = TypeHAlignCenter
                .Col = 2: .value = RS.Fields("ordcd").value & ""
                .Col = 3: .value = RS.Fields("testnm").value & ""
                .Col = 4: .value = RS.Fields("unitQty").value & ""
                .Col = 5: .value = RS.Fields("orddt").value & ""
                .Col = 6: .value = RS.Fields("ordno").value & ""
                .Col = 7: .value = RS.Fields("ordseq").value & ""
                .Col = 8: .value = RS.Fields("ocsordno").value & ""
                .Col = 9: .value = RS.Fields("bussdiv").value & ""
                .Col = 10: .value = RS.Fields("deptcd").value & ""
                .Col = 11: .value = Format(RS.Fields("orddt").value & "", "####-##-##")
                
                RS.MoveNext
            Loop
            
            For ii = 1 To .DataRowCnt
                Call tblFilter_Click(1, ii)
            Next
            Call tblFilter_Click(1, ii)
        End If
    End With
    
    Set RS = Nothing
    
End Sub

Private Sub cmdFSave_Click()
    Dim objBldDelivery As clsBldDelivery
    Dim RS              As Recordset
    Dim strWorkArea     As String
    Dim strAccDt        As String
    Dim strAccSeq       As String
    Dim strACnt         As String
    Dim strDCnt         As String
    
    Dim strPtid         As String
    Dim strOrdDt        As String
    Dim strOrdNo        As String
    Dim strOrdSeq       As String
    Dim strOrdCd        As String
    Dim strBussdiv      As String
    Dim strOrder_key    As String
    Dim strOcsOrdNo     As String
    Dim strTestCd       As String
    Dim strTestNm       As String
    Dim strEntdt        As String
    Dim strEntTm        As String
    Dim strEntID        As String
    Dim strStsCd        As String
    Dim strDeptCd       As String
    Dim strTmp          As String
    Dim blnDelivery     As Boolean
    Dim SSQL           As String
    Dim AK_Chk         As String
    
    Dim ii             As Integer
    Dim jj             As Integer
    
    
    If Val(txtQty.Text) < 1 Then Exit Sub
    
    strPtid = txtPtId.Text
    strEntdt = Format(GetSystemDate, PRESENTDATE_FORMAT)
    strEntTm = Format(GetSystemDate, PRESENTTIME_FORMAT)
    strEntID = ObjSysInfo.EmpId
    strStsCd = BBSBloodStatus.stsDELIVERY
    
    Set objBldDelivery = New clsBldDelivery
    
On Error GoTo Errors
    DBConn.BeginTrans
    
    With tblFilter
        For ii = 1 To .DataRowCnt
            .Row = ii: .Col = 1
            
            If .CellType = CellTypeCheckBox And .value = "1" Then
                
                .Col = 2: strTestCd = .value
                .Col = 3: strTestNm = .value
                .Col = 5: strOrdDt = .value
                .Col = 6: strOrdNo = .value
                .Col = 7: strOrdSeq = .value
                .Col = 8: strOcsOrdNo = .value
                .Col = 9: strBussdiv = .value
                
                .Col = 10: strDeptCd = .value
                
                
                strTmp = MsgBox("�˻��ڵ� : " & strTestCd & "[" & strTestNm & "]" & vbCrLf & _
                                "����       : " & txtQty.Text & " �� ���͸� ����Ͻðٽ��ϱ�?", vbYesNo + vbInformation, "Info")
                
                If strTmp = vbYes Then
                    If Not (Val(txtQty.Text) > Val(txtQty.tag) Or Val(txtQty.Text) < 1) Then
                        '���� ���������� ������� �ٽ� ã�� �ʿ� ����(����)
'                        SSQL = " SELECT PATIENT_NO, BOJO " & _
'                               " FROM  med_pmpa.OPD_MAST " & _
'                               " WHERE PATIENT_NO = '" & txtPtId.Text & "'" & _
'                               "   AND ACT_DATE = To_Date('" & Format(strOrdDt, "####-##-##") & "', 'YYYY-MM-DD')" & _
'                               "   AND BOJO IN ('01', '11') "
'                        Set RS = Nothing
'                        Set RS = New Recordset
'
'                        RS.Open SSQL, DBConn
'
'                        If Not RS.EOF Then
'                            AK_Chk = "AK"
'                        End If
'
'                        Set RS = Nothing
                        For jj = 1 To txtQty.Text
                            '���� ���������� ���� �ۼ��� ���ص� ��
'                            '�������� �ۼ�
'                            SSQL = objBldDelivery.Filter_SuNapSQL(strOcsOrdNo, strBussdiv, strTestCd, strTestCd, "1", "", strTestCd, strOrder_key)
'
'                            If strBussdiv = "2" Or (strDeptCd = "ER" Or AK_Chk = "AK") Then
'                                If SSQL <> "" Then DBConn.Execute SSQL
'                            End If
                            
                            '������� ���� �ۼ�
                            Dim RsF  As Recordset
                            Dim strSQL As String
                            Dim strWorkSeq As String
                            
                            Set RsF = Nothing
                            Set RsF = New Recordset
                            
                            strSQL = " select max(workseq) as MaxWS from " & T_BBS304
                            
                            RsF.Open strSQL, DBConn
                            
                            If RsF.EOF Then
                                strWorkSeq = 1
                            Else
                                strWorkSeq = Val(RsF.Fields("MaxWS").value & "") + 1
                            End If
                            
                            Set RsF = Nothing
                            
                            SSQL = " INSERT INTO " & T_BBS304 & "(workseq,ptid,orddt,ordno,ordseq,ordcd,stscd,entdt,enttm,entid) values(" & _
                                 DBV("workseq", strWorkSeq, 1) & DBV("ptid", strPtid, 1) & DBV("orddt", strOrdDt, 1) & _
                                 DBV("ordno", strOrdNo, 1) & DBV("ordseq", strOrdSeq, 1) & DBV("ordcd", strTestCd, 1) & _
                                 DBV("stscd", strStsCd, 1) & DBV("entdt", strEntdt, 1) & DBV("enttm", strEntTm, 1) & _
                                 DBV("entid", strEntID) & ")"
                            DBConn.Execute SSQL
                            '�ٵ� ������Ʈ
'                            SSQL = " update " & T_LAB102 & " set " & DBW("stscd", BBSOrderStatus.stsINPROCESS, 3) & DBW("donefg", BBSOrderStatus.stsINPROCESS, 2) & _
'                                   " where " & DBW("ptid=", strPtID) & _
'                                   " and " & DBW("orddt=", strOrdDt) & _
'                                   " and " & DBW("ordno=", strOrdNo) & _
'                                   " and " & DBW("ordseq=", strOrdSeq)
'                            DBConn.Execute SSQL

                            '���� ����������
'���� ���� ������
'                            SSQL = " update mdbldort set " & DBW("stscd", BBSOrderStatus.stsINPROCESS, 3) & DBW("donefg", BBSOrderStatus.stsINPROCESS, 2) & _
'                                   " where " & DBW("patno=", strPtid) & _
'                                   " and orddate=to_date(" & strOrdDt & " ,'yyyymmdd') " & _
'                                   " and " & DBW("ordseqno=", strOrdNo)
'���� ���� ����
'OCS�� ���¸� R�� ����
                            SSQL = " update mdbldort set " & DBW("stscd", BBSOrderStatus.stsINPROCESS, 3) & DBW("donefg", BBSOrderStatus.stsINPROCESS, 2) & ", procstat='R'" & _
                                   " where " & DBW("patno=", strPtid) & _
                                   " and orddate=to_date(" & strOrdDt & " ,'yyyymmdd') " & _
                                   " and " & DBW("ordseqno=", strOrdNo)
                            DBConn.Execute SSQL
                            
                            '������ȣ�� �Ǽ� ������Ʈ
                            SSQL = " select a.workarea,a.accdt,a.accseq ,b.assigncnt,b.deliverycnt " & _
                                   " from " & T_LAB102 & " a," & T_BBS203 & " b" & _
                                   " where " & DBW("a.ptid=", strPtid) & _
                                   " and " & DBW("a.orddt=", strOrdDt) & _
                                   " and " & DBW("a.ordno=", strOrdNo) & _
                                   " and " & DBW("a.ordseq=", strOrdSeq) & _
                                   " and a.workarea=b.workarea and a.accdt=b.accdt and a.accseq=b.accseq"
                            
                            Set RS = Nothing
                            Set RS = New Recordset
                            RS.Open SSQL, DBConn
                            
                            strDCnt = "1": strACnt = "1"
                            If Not RS.EOF Then
                                strWorkArea = RS.Fields("workarea").value & ""
                                strAccDt = RS.Fields("accdt").value & ""
                                strAccSeq = RS.Fields("accseq").value & ""
                                strDCnt = Val(strDCnt) + Val(RS.Fields("deliverycnt").value & "")
                                strACnt = Val(strACnt) + Val(RS.Fields("assigncnt").value & "")
                                 
                                SSQL = " update " & T_BBS203 & " set " & DBW("deliverycnt", strDCnt, 3) & DBW("assigncnt", strACnt, 2) & _
                                       " where " & DBW("workarea=", strWorkArea) & _
                                       " and " & DBW("accdt=", strAccDt) & _
                                       " and " & DBW("accseq=", strAccSeq)
                                DBConn.Execute SSQL
                            End If
                        Next
                    Else
                        MsgBox "�������� Ȯ���ϼ���", vbInformation + vbOKOnly, "Info"
                        txtQty.SetFocus
                        GoTo Skip
                        Exit For
                    End If
                    blnDelivery = True
                End If
            End If
        Next
    End With
    
    If blnDelivery = True Then MsgBox "���������� ���Ǿ����ϴ�.", vbInformation
    
    Call FilterClear
Skip:
    DBConn.CommitTrans
    Set objBldDelivery = Nothing
    Exit Sub
Errors:
    DBConn.RollbackTrans
    Set objBldDelivery = Nothing
End Sub

Private Sub objBldList_SelectedItem(ByVal pSelectedItem As String)
    If chkBarCode.value = 1 Then
        txtBldNo.Text = Replace(medGetP(pSelectedItem, 1, ";"), "-", "")
    Else
        txtBldNo.Text = medGetP(pSelectedItem, 1, ";")
    End If
    txtBldNoLostFocus
'    cboCompo.ListIndex = medComboFind(cboCompo, medGetP(pSelectedItem, 4, ";"))
End Sub

Private Sub objPop_Click(ByVal vMenuID As Long)
    Select Case vMenuID
        Case MENU_DEL
            With tblBldList
                .Row = .ActiveRow
                .Action = ActionDeleteRow
                .RowHeight(-1) = 12.95
            End With
        Case MENU_SAVE
            Dim strPtid As String
            Dim strPtnm As String
            Dim strOrdDt As String
            
            With tblBldList
                .Row = .ActiveRow:  .Col = TblColumn.tcPTID: strPtid = .value
                                    .Col = TblColumn.tcPTNM: strPtnm = .value
                                    .Col = TblColumn.tcORDDT: strOrdDt = .value
            End With
            If strPtid = "" Then Exit Sub
            DoEvents
            
            Call cmdFOpen_Click
            
            txtPtId.Text = strPtid: lblPtNm.Caption = strPtnm
            dtpToDate.value = strOrdDt
            dtpFromDate.value = DateAdd("d", -3, dtpToDate.value)
            
            If cmdFQuery.Enabled Then cmdFQuery_Click
    End Select
End Sub

Private Sub tblFilter_Click(ByVal Col As Long, ByVal Row As Long)
    Dim RS          As Recordset
    Dim SSQL        As String
    Dim strPtid     As String
    Dim strOrdDt    As String
    Dim strOrdNo    As String
    Dim strOrdSeq   As String
    Dim strUnitQty  As String
    Dim ii          As Integer
    
    If Row < 1 Then Exit Sub
    
    strPtid = txtPtId.Text
    
    With tblFilter
        .Row = Row: .Col = Col
        If .value = "" Then Exit Sub
        .Col = 4: strUnitQty = .value
        .Col = 5: strOrdDt = .value
        .Col = 6: strOrdNo = .value
        .Col = 7: strOrdSeq = .value
    End With
    
    txtQty.Text = strUnitQty
    lblDelCnt.Caption = "":  lblRetCnt.Caption = "": lblExpCnt.Caption = ""
    
    '���Ǽ�
    SSQL = " SELECT stscd,count(stscd) as cnt " & _
           " FROM " & T_BBS304 & _
           " WHERE " & DBW("ptid=", strPtid) & _
           " AND " & DBW("orddt=", strOrdDt) & _
           " AND " & DBW("ordno=", strOrdNo) & _
           " AND " & DBW("ordseq=", strOrdSeq) & _
           " GROUP BY stscd"
           
    
    Set RS = Nothing
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    If Not RS.EOF Then
        lblDelCnt.Caption = RS.Fields("cnt").value & ""
    End If
    
    SSQL = " SELECT count(expfg) as expcnt from " & T_BBS304 & _
           " WHERE " & DBW("ptid=", strPtid) & _
           " AND " & DBW("orddt=", strOrdDt) & _
           " AND " & DBW("ordno=", strOrdNo) & _
           " AND " & DBW("ordseq=", strOrdSeq) & _
           " AND " & DBW("expfg=", "1") & _
           " GROUP BY expfg"
           
    Set RS = Nothing
    Set RS = New Recordset
    
    RS.Open SSQL, DBConn
    
    If Not RS.EOF Then
        lblExpCnt.Caption = RS.Fields("expcnt").value & ""
    End If
    
     SSQL = " SELECT count(retfg) as retcnt from " & T_BBS304 & _
           " WHERE " & DBW("ptid=", strPtid) & _
           " AND " & DBW("orddt=", strOrdDt) & _
           " AND " & DBW("ordno=", strOrdNo) & _
           " AND " & DBW("ordseq=", strOrdSeq) & _
           " AND " & DBW("retfg=", "1") & _
           " GROUP BY retfg"
           
    Set RS = Nothing
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    If Not RS.EOF Then
        lblRetCnt.Caption = RS.Fields("retcnt").value & ""
    End If
    
    txtQty.Text = Val(txtQty.Text) - (Val(lblDelCnt.Caption) - (Val(lblExpCnt.Caption) + Val(lblRetCnt.Caption)))
    txtQty.SelStart = 0
    txtQty.SelLength = Len(txtQty.Text)
    txtQty.SetFocus
    txtQty.tag = txtQty.Text
    If txtQty.Text = "0" Then
        
        With tblFilter
            
            .Row = Row: .Col = 1
'            If .CellType <> CellTypeStaticText Then
'                MsgBox "�̹� ���� �����Դϴ�.", vbInformation + vbOKOnly, "Info"
'            End If
            .CellType = CellTypeStaticText: .value = "��": .FontBold = True: .ForeColor = DCM_LightRed
        End With
    End If
    
    Set RS = Nothing
End Sub

Private Sub Form_Activate()
    txtCompID.SetFocus
End Sub

Private Sub tblBldList_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    Dim strFilter As String
    
    If Row < 1 Then Exit Sub
    
    With tblBldList
        .Row = Row
        .Col = 1
        If .value = "" Then Exit Sub
        .Action = ActionActiveCell
        .Col = TblColumn.tcFilter: strFilter = .value
    End With
    
    Set objPop = New clsPopupMenu
    With objPop
        .AddMenu MENU_DEL, "DELETE"
        
        If strFilter = "1" Then
            .AddMenu MENU_SEP, "-", eSEPARATOR
            .AddMenu MENU_SAVE, "FILTER ���"
        End If
        
        .PopupMenus Me.hwnd
    End With
    Set objPop = Nothing
End Sub

Private Sub tblBldList_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
    Dim strTmp  As String
    Dim strTmp1 As String
    Dim strtip  As String
    
    If Row < 1 Then Exit Sub
    
    With tblBldList
        .Row = Row
        .Col = TblColumn.tcRmk2:  strTmp = .value
        .Col = TblColumn.tcPtRmk: strTmp1 = .value
        If strTmp = "" And strTmp1 = "" Then Exit Sub
        Call .SetTextTipAppearance("����ü", 9, False, False, &HFFFFC0, vbBlack)
        If strTmp1 <> "" Then
            strtip = strtip & " [ȯ��Ư�̻���]" & vbCrLf & vbCrLf & strTmp1 & vbNewLine
        End If
        If strTmp <> "" Then
            strtip = strtip & " [�� �� �� �� ]" & vbCrLf & vbCrLf & strTmp
        End If
        TipWidth = 5000
        MultiLine = 1
        TipText = vbNewLine & strtip & vbNewLine
        ShowTip = True
        
    End With
    
End Sub

Private Sub txtBldNo_Change()
    
    If chkBarCode.value = 1 Then Exit Sub
    Dim lngLen As Long
    
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
    
    If lblCompoCd.Caption <> "" Then
        lblCompoCd.Caption = ""
        lblCompoNm.Caption = ""
        cboCompo.Clear
    End If
End Sub

Private Sub txtBldNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
    
    If chkBarCode.value = 1 Then Exit Sub
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

Private Sub txtBldNo_GotFocus()
    SendKeys "{Home}+{End}"
    txtBldNo.tag = txtBldNo
End Sub

Private Sub txtBldNo_LostFocus()
    If chkBarCode.value <> 1 Then
        If Len(Trim(txtBldNo.Text)) <= 6 Then Exit Sub
    Else
        If Len(txtBldNo.Text) < 6 Then Exit Sub
    End If
    If Trim(txtCompID.Text) = "" Then
        MsgBox "���ȯ�ڹ�ȣ�� �Է� �ϼ���!", vbCritical, "Ȯ��"
        txtBldNo.Text = ""
        txtCompID.SetFocus
        Exit Sub
    End If
    If txtBldNo.Text = "" Then Exit Sub
    If txtBldNo.tag = txtBldNo Then Exit Sub
    
    Me.MousePointer = 11
    '--------- �ڷ���ȸ ----------
    Call txtBldNoLostFocus
    txtBldNo.Text = ""
    Me.MousePointer = 0
End Sub
Private Sub cmdClear_Click()
    Clear
    txtCompID.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
    Set frmBBS303N = Nothing
End Sub

Private Sub cmdRcvId_Click()
    Dim ii      As Integer
    
    Set objMyList = New clsPopUpList
    
    With objMyList
        .Connection = DBConn
        .FormCaption = "������ȸ": .ColumnHeaderText = "���;������"
        txtRcvId.Text = "": lblRcvNm.Caption = ""
        Call .LoadPopUp(GetSQLHisEmpList) ', 2350, 7650)
        If .SelectedString <> "" Then
            txtRcvId.Text = medGetP(.SelectedString, 1, ";")
            lblRcvNm.Caption = medGetP(.SelectedString, 2, ";")
        End If
    End With
    
    With tblBldList
        For ii = 1 To .DataRowCnt
            .Row = ii
            .Col = TblColumn.tcRCVNM: .value = lblRcvNm.Caption
        Next
    End With
    
    Set objMyList = Nothing
End Sub

'Private Sub SetHisEmpToLisEmp(ByVal vEmpId As String, ByVal vEmpNm As String)
''HIS�� ���������� LIS�� �Է��Ѵ�. (���������Ϳ� ����� �����Ϳ� ����)
'    Dim strSQL As String
'    Dim RS As Recordset
'
'    On Error GoTo ErrTrap
'
'    'LIS ���� �����Ϳ� �Է�
'    strSQL = " select * from s2com006"
'    strSQL = strSQL & " where empid='" & vEmpId & "'"
'
'    DBConn.BeginTrans
'
'    Set RS = New Recordset
'    RS.Open strSQL, DBConn
'
'    If RS.EOF Then 'lis ���� �����Ϳ� ���� ���
'        strSQL = " insert into s2com006"
'        strSQL = strSQL & " (empid, empnm) values"
'        strSQL = strSQL & " ('" & vEmpId & "','" & vEmpNm & "')"
'
'        DBConn.Execute strSQL
'    End If
'
'    'LIS ����� �����Ϳ� �Է�
'    strSQL = " select * from s2com010"
'    strSQL = strSQL & " where loginid=''"
'
'    Set RS = Nothing
'    Set RS = New Recordset
'
'    RS.Open strSQL, DBConn
'
'    If RS.EOF Then 'LIS ����� �����Ϳ� ���� ���
'        strSQL = " insert into s2com010"
'        strSQL = strSQL & " (loginid, loginpass,empid,logindesc, groupid) values"
'        strSQL = strSQL & " ('" & vEmpId & "','2','" & vEmpId & "','" & vEmpNm & "','G002')"
'
'        DBConn.Execute strSQL
'    End If
'
'    Set RS = Nothing
'
'    DBConn.CommitTrans
'
'    Exit Sub
'
'ErrTrap:
'    DBConn.RollbackTrans
'
'End Sub

Private Sub Form_Load()
    Call Clear
    chkBarCode.value = 1
    If BLOOD_DEL_USED Then
        cmdRePrint.Visible = True
        cmdTransStatPaper.Visible = True
    End If
End Sub

Private Function Search_PtInfo() As Boolean
    Dim objPt   As clsPtInformation
    Dim RS      As Recordset
    Dim ii      As Long
    Dim strLng  As String
    
    
    tblFilter.MaxRows = 0
    lblRetCnt.Caption = "": txtQty.Text = "": lblDelCnt.Caption = "": lblExpCnt.Caption = ""
    
    If txtPtId.Text = "" Then
        lblPtNm.Caption = ""
        Search_PtInfo = True
    Else
        For ii = 1 To Val(BBS_PTID_LENGTH) - 1
            strLng = strLng & "0"
        Next ii

        If Len(Trim(txtPtId.Text)) <> BBS_PTID_LENGTH Then
            txtPtId.Text = Format(txtPtId.Text, strLng & "#")
        End If

        Set objPt = New clsPtInformation
        Set RS = Nothing
        Set RS = New Recordset
         
        RS.Open objPt.Get_Ptid(txtPtId), DBConn
        
        If RS.EOF = False Then
            With objPt
                .BedPt_Chk txtPtId.Text, Format(GetSystemDate, PRESENTDATE_FORMAT)
                If .PtDiv = "BED" Then
                    lblPtNm.Caption = .ptnm
                Else
                    lblPtNm.Caption = .ptnm
                End If
            End With
            Search_PtInfo = True
        Else
            MsgBox "�ش�Ǵ� ȯ�ڰ� �����ϴ�. Ȯ���� ��ȸ�ϼ���.", vbInformation + vbOKOnly, Me.Caption
            txtPtId = ""
            lblPtNm.Caption = ""
            Search_PtInfo = False
        End If
        Set RS = Nothing
        Set objPt = Nothing
    End If
    If Search_PtInfo Then cmdFQuery.SetFocus
End Function

Private Sub txtCompID_GotFocus()
    txtCompID.SelStart = 0
    txtCompID.SelLength = Len(txtCompID)
End Sub

Private Sub txtCompID_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txtCompID_LostFocus()
    Dim ii      As Integer
    Dim strLng  As String
    
    If txtCompID = "" Then
        Clear
        Exit Sub
    End If
        
    For ii = 1 To Val(BBS_PTID_LENGTH) - 1
        strLng = strLng & "0"
    Next ii
    txtCompID.Text = Format(txtCompID.Text, strLng & "#")
    txtCompID.Text = txtCompID.Text
    
    If CompPtInfo = False Then
        txtCompID.SetFocus
    Else
        txtBldNo.SetFocus
    End If
End Sub

Private Function CompPtInfo() As Boolean
    Dim objPtInfo As clsPtInformation
    Dim DrRS      As Recordset
    Dim ii        As Long
    Dim strLng    As String
    
    lblCompNm.Caption = "": lblSexAge.Caption = ""
    If txtCompID.Text = "" Then CompPtInfo = True: Exit Function
    
    Set objPtInfo = New clsPtInformation
    Set DrRS = New Recordset
    
    DrRS.Open objPtInfo.Get_Ptid(txtCompID.Text), DBConn
    If DrRS.EOF = True Then
        MsgBox "�ش�Ǵ� ȯ�ڰ� �����ϴ�. Ȯ���� ��ȸ�ϼ���.", vbInformation + vbOKOnly, Me.Caption
        CompPtInfo = False
    Else
        With objPtInfo
            .BedPt_Chk txtCompID.Text, Format(GetSystemDate, PRESENTDATE_FORMAT)
            If .PtDiv = "BED" Then
                txtCompID.Text = .Ptid
                lblCompNm.Caption = .ptnm
                lblSexAge.Caption = .Sex & " / " & .Age
            Else
                txtCompID.Text = .Ptid
                lblCompNm.Caption = .ptnm
                lblSexAge.Caption = .Sex & " / " & .Age
            End If
        End With
        
        CompPtInfo = True
    End If
    
    Set DrRS = Nothing
    Set objPtInfo = Nothing
    
    '** �ش� ȯ���� �������� ��������
    Call DetailSearch(Trim(txtCompID.Text))
    
End Function

Private Sub txtPtId_Change()
    If lblPtNm.Caption <> "" Then Exit Sub
    
    lblDelCnt.Caption = "": lblRetCnt.Caption = "": lblExpCnt.Caption = ""
    txtQty.Text = ""
    lblPtNm.Caption = ""
    tblFilter.MaxRows = 0
    dtpFromDate.value = DateAdd("d", -3, GetSystemDate)
    dtpToDate.value = GetSystemDate
End Sub

Private Sub txtPtId_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call txtPtId_LostFocus
    End If
End Sub

Private Sub txtPtId_LostFocus()
    If lblPtNm.Caption <> "" Then Exit Sub
    Call Search_PtInfo
End Sub

Private Sub txtRcvId_Change()
    Dim i As Long
    
    If lblRcvNm.Caption <> "" Then
        lblRcvNm.Caption = ""
        
        With tblBldList
            For i = 1 To .DataRowCnt
                .Row = i
                .Col = TblColumn.tcRCVNM: .value = ""
            Next
        End With
    End If
End Sub

Private Sub txtRcvId_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtRcvId_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtRcvId_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtRcvId_Validate(Cancel As Boolean)
    Dim strEmpNm As String
    Dim i As Long
    
'    If txtBldNo.Text = "" Then Exit Sub
    If txtRcvId.Text = "" Then Exit Sub
    
    strEmpNm = GetEmpNm(txtRcvId.Text)
    
    If strEmpNm = "" Then
        Cancel = True
        MsgBox "��ϵ��� ���� ������Դϴ�.", vbExclamation
    Else
        lblRcvNm.Caption = strEmpNm
        
        With tblBldList
            For i = 1 To .DataRowCnt
                .Row = i
                .Col = TblColumn.tcRCVNM: .value = lblRcvNm.Caption
            Next
        End With
    End If
    
    If Cancel Then SendKeys "{Home}+{End}"
End Sub


'Private Sub txtRcvId_LostFocus()
''    If txtRcvId.tag = txtRcvId Then Exit Sub
'    Call Query_RcvNm
''    txtRcvId.tag = txtRcvId
'End Sub
'
'Private Sub Query_RcvNm()
'    Dim name As String
'    Dim ii   As Integer
'
'    If txtRcvId.Text = "" Then
'        lblRcvNm.Caption = ""
'        Exit Sub
'    End If
'
''OCS ���� �����Ϳ��� ��ȸ�� �㿡 LIS�� ����� ������, ���������Ϳ� insert ����
''    name = GetEmpNm(txtRcvId.Text)
'
'    name = GetEmpNm(txtRcvId.Text)
'
'    If name = "" Then txtRcvId.Text = ""
'    lblRcvNm.Caption = name
'
''    If name <> "" Then
''        Call SetHisEmpToLisEmp(txtRcvId.Text, lblRcvNm.Caption)
''    End If
'
'
'End Sub

Private Sub txtBldNoLostFocus()
    Dim DrRS           As Recordset
    Dim objBldDelivery As clsBldDelivery
    Dim BldSrc  As String
    Dim BldYY   As String
    Dim BldNo   As String
    
    Dim i As Long
    
    
    If chkBarCode.value = 1 Then
        BldSrc = Mid(txtBldNo, 1, 2)
        BldYY = Mid(txtBldNo, 3, 2)
'        BldNo = Format(Mid(Mid(txtBldNo, 5), 1, Len(Mid(txtBldNo, 5)) - 2), "######")
        BldNo = Format(Mid(txtBldNo.Text, 5, 6), "00000#")
    Else
        BldSrc = medGetP(txtBldNo, 1, "-")
        BldYY = medGetP(txtBldNo, 2, "-")
        BldNo = Format(medGetP(txtBldNo, 3, "-"), "######")
    End If
    
    
    If BldSrc = "" Or BldYY = "" Or BldNo = "" Then Exit Sub
    If Trim(txtCompID.Text) = "" Then
        MsgBox "���ȯ�ڹ�ȣ�� �Է� �ϼ���!", vbCritical, "�Է�Ȯ��"
        Exit Sub
    End If
    
    Set objBldDelivery = New clsBldDelivery
    
    Set DrRS = objBldDelivery.GetBloodCompoList(BldSrc, BldYY, BldNo, mode)
    
    If DrRS Is Nothing Then
        Set objBldDelivery = Nothing
        Exit Sub
    End If
    
    With DrRS
        cboCompo.Clear
        If .RecordCount = 1 Then
            cboCompo.AddItem .Fields("compocd").value & "" & COL_DIV & .Fields("componm").value & ""
            lblCompoCd.Caption = .Fields("compocd").value & ""
            lblCompoNm.Caption = .Fields("componm").value & ""
            cboCompo.ListIndex = 0
        ElseIf .RecordCount > 1 Then
            For i = 1 To .RecordCount
                cboCompo.AddItem .Fields("compocd").value & "" & COL_DIV & .Fields("componm").value & ""
                .MoveNext
            Next i
        Else
            MsgBox "�̹� ���Ǿ��ų�, �غ���� ���� �����Դϴ�", vbCritical, Me.Caption
            txtBldNo = ""
            txtBldNo.SetFocus
        End If
    End With
    
    Set DrRS = Nothing
    Set objBldDelivery = Nothing
End Sub

Private Sub DetailSearch(Ptid As String)
    Dim ObjABO  As New clsABO
    Dim strRh   As String
    Dim strTmp  As String
    Dim ii      As Integer
    
    With ObjABO
        .Ptid = Ptid
        If .GetABO = True Then
            lblABO.Caption = .ABO & .Rh
            sABOType = .ABO
            sRHType = .Rh
        Else
            lblABO.Caption = ""
            sABOType = ""
            sRHType = ""
        End If
    End With
    
    Set ObjABO = Nothing
    
    If Len(lblABO.Caption) > 3 Then
        lblABO.Caption = medGetP(lblABO.Caption, 1, "(") & medGetP(lblABO.Caption, 2, ")")
    Else
        For ii = 1 To Len(lblABO.Caption)
            If Mid(lblABO.Caption, ii, 1) = "+" Or Mid(lblABO.Caption, ii, 1) = "-" Then
                strRh = Mid(lblABO.Caption, ii, 1)
            Else
                strTmp = strTmp & Mid(lblABO.Caption, ii, 1)
            End If
        Next ii
        lblABO.Caption = strTmp & strRh
    End If
            
End Sub

Private Function DupCheck(ByVal pBldNo As String) As Boolean
'�ߺ����� üũ�Ѵ�.
    Dim strClip As String
    
    With tblBldList
        .Row = 1: .Row2 = .DataRowCnt
        .Col = 1: .COL2 = TblColumn.tcDUP
        .BlockMode = True
        strClip = .ClipValue
        .BlockMode = False
        
        If InStr(strClip, pBldNo) Then
            DupCheck = True
        Else
            DupCheck = False
        End If
    End With
    
End Function

Private Sub cboCompo_Click()
    Dim BldSrc As String
    Dim BldYY  As String
    Dim BldNo  As String
    
    If chkBarCode.value = 1 Then
        BldSrc = Mid(txtBldNo, 1, 2)
        BldYY = Mid(txtBldNo, 3, 2)
'        BldNo = Format(Mid(Mid(txtBldNo, 5), 1, Len(Mid(txtBldNo, 5)) - 2), "######")
        BldNo = Format(Mid(txtBldNo, 5, 6), "00000#")
    Else
        BldSrc = medGetP(txtBldNo.Text, 1, "-")
        BldYY = medGetP(txtBldNo.Text, 2, "-")
        BldNo = Format(medGetP(txtBldNo.Text, 3, "-"), "00000#")
    End If
    lblCompoCd.Caption = medGetP(cboCompo.Text, 1, COL_DIV)
    lblCompoNm.Caption = medGetP(cboCompo.Text, 2, COL_DIV)
    
    If DupCheck(BldSrc & BldYY & Format(BldNo, "00000#") & COL_DIV & lblCompoCd.Caption) = True Then
        MsgBox "����� ������� �����Դϴ�.", vbInformation + vbOKOnly, "�������"
        txtBldNo.SetFocus
      '  Call Clear
        Exit Sub
    End If
    
    Call SetBloodInfo(BldSrc, BldYY, BldNo, lblCompoCd.Caption)
End Sub

Private Sub SetBloodInfo(ByVal BldSrc As String, ByVal BldYY As String, ByVal BldNo As String, ByVal CompoCd As String)
    Dim objBldDelivery As clsBldDelivery
    Dim objSql         As clsCrossMatching
    Dim ObjABO         As clsABO
    Dim RS             As Recordset
    
    Dim strFilter      As String
    Dim strRT          As String
    Dim strDUP         As String
    Dim ADt            As String
    Dim sPtid          As String
    Dim strTmp         As String
    Dim strSSN         As String
    Dim strABO         As String
    
    Dim ii              As Integer
    
    DoEvents
    
    strTmp = ""
    Set objBldDelivery = New clsBldDelivery
    Set RS = objBldDelivery.GetBloodInfo(BldSrc, BldYY, BldNo, CompoCd, mode)
    
    If RS Is Nothing Then
        Set objBldDelivery = Nothing
        Exit Sub
    End If
    
    Set ObjABO = New clsABO
    Set objSql = New clsCrossMatching
        
    With tblBldList
        .ReDraw = False
        If RS.RecordCount < 1 Then
            MsgBox "������ ã�� �� �����ϴ�", vbCritical, Me.Caption
            Call Clear
        Else
            .RowHeight(-1) = 12.95
            
            If .MaxRows < .DataRowCnt + 1 Then
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
            Else
                .Row = .DataRowCnt + 1
            End If
            
            '** ���ȯ�ڿ� ���ؼ� Ʋ���� ���������� ó�� �� By 2007.12.27 M.G.Choi
            .Row = .Row: .Row2 = .Row
            .Col = 1: .COL2 = .MaxCols
            .BlockMode = True
            strABO = RS.Fields("abo").value & "" & RS.Fields("rh").value & ""
            If strABO <> lblABO.Caption Then
                MsgBox "���ȯ���� �������� ���� ���� �ʽ��ϴ�.", vbCritical, "Ȯ���ʿ�"
                .FontBold = True
                .ForeColor = DCM_LightRed
            End If
            .BlockMode = False
            
            If Trim(txtCompID.Text) <> RS.Fields("ptid").value & "" Then
                MsgBox "�ش������� ���ȯ�ڿ� Ʋ���ϴ�.", vbCritical, "Ȯ���ʿ�"
                txtBldNo.Text = ""
                txtBldNo.SetFocus
                txtBldNo.SelStart = 0
                txtBldNo.SelLength = Len(txtCompID)
                cboCompo.Clear
                With tblBldList
                    For ii = 1 To .MaxRows
                        .Row = ii
                        .Col = TblColumn.tcRTP
                        .CellType = CellTypeStaticText
                        .Col = TblColumn.tcFilter
                        .CellType = CellTypeStaticText
                    Next
                End With
                Call medClearTable(tblBldList, False, False)
                GoTo Skip
'                .FontBold = True
'                .ForeColor = DCM_LightRed
            End If
            
            .Col = TblColumn.tcBldNo: .value = BldSrc & "-" & BldYY & "-" & Format(BldNo, "00000#")
            strDUP = Replace(.value, "-", "")
            .Col = TblColumn.tcCOMPONM: .value = lblCompoNm.Caption
            .Col = TblColumn.tcABO: .value = RS.Fields("abo").value & "" & RS.Fields("rh").value & ""
            .Col = TblColumn.tcVol: .value = RS.Fields("volumn").value & "" & "cc"
            
            ADt = Format(DateAdd("d", Val(RS.Fields("available").value & "") - 1, Format(RS.Fields("coldt").value & "", "####-##-##")), "YYYY-MM-dd")
            
            .Col = TblColumn.tcAVAIL:   .value = ADt
                        
            .Col = TblColumn.tcFilter:  .CellType = CellTypeCheckBox: .TypeHAlign = TypeHAlignCenter
                        
            strRT = RS.Fields("irrfg").value & ""
            .Col = TblColumn.tcRT:      .value = IIf(strRT = "1", "��", ""): .ForeColor = DCM_LightRed: .TypeHAlign = TypeHAlignCenter
            
            .Col = TblColumn.tcDELDT:   .value = Format(GetSystemDate, "yyyy-mm-dd")
            .Col = TblColumn.tcPTID:    .value = RS.Fields("ptid").value & ""
                                        sPtid = RS.Fields("ptid").value & ""
            .Col = TblColumn.tcPTNM:    .value = GetPtNm(sPtid)
                                        strTmp = objSql.GetptidRmk(sPtid)
                                        If strTmp <> "" Then
                                            .ForeColor = DCM_LightRed: .FontBold = True
                                            .Col = TblColumn.tcPtRmk: .value = strTmp
                                        Else
                                            .ForeColor = vbBlack: .FontBold = False
                                            .Col = TblColumn.tcPtRmk: .value = strTmp
                                        End If
            
            .Col = TblColumn.tcRTP:     .CellType = CellTypeCheckBox: .TypeHAlign = TypeHAlignCenter
            If strRT = "1" Then .value = 1
            
            .Col = TblColumn.tcORDDT:   .value = Format(RS.Fields("orddt").value & "", "####-##-##")
            .Col = TblColumn.tcTEST:    .value = RS.Fields("testnm").value
            .Col = TblColumn.tcUNIT:    .value = RS.Fields("unitqty").value & ""
            
            .Col = TblColumn.tcORDNO:   .value = RS.Fields("ordno").value & ""
            .Col = TblColumn.tcORDSEQ:  .value = RS.Fields("ordseq").value & ""
            .Col = TblColumn.tcACCDT:   .value = RS.Fields("accdt").value & ""
            .Col = TblColumn.tcACCSEQ:  .value = RS.Fields("accseq").value & ""
            .Col = TblColumn.tcRSTSEQ:  .value = RS.Fields("rstseq").value & ""
            .Col = TblColumn.tcCOMPOCD: .value = lblCompoCd.Caption
            .Col = TblColumn.tcRCVNM:   .value = lblRcvNm.Caption
            
            .Col = TblColumn.tcRMK:     .value = IIf(RS.Fields("rmk").value & "" <> "", "Y", ""): .ForeColor = DCM_Red
            .Col = TblColumn.tcRmk2:    .value = RS.Fields("rmk").value & ""
            'ó���ڵ�

            .Col = TblColumn.tcORDCD: .value = SetNewTest(BldSrc, BldYY, BldNo, CompoCd, RS.Fields("volumn").value & "", _
                                                                  RS.Fields("testdiv").value & "")
            
            '�ߺ�üũ
            .Col = TblColumn.tcDUP: .value = strDUP & COL_DIV & lblCompoCd.Caption
            ObjABO.Ptid = RS.Fields("ptid").value & ""
            If ObjABO.GetABO = True Then
                .Col = TblColumn.tcABOP: .value = ObjABO.ABO & ObjABO.Rh
            Else
                .Col = TblColumn.tcABOP: .value = ""
            End If
            
            .Col = TblColumn.tcDEPT:    .value = RS.Fields("wardid").value & "" & COL_DIV & RS.Fields("hosilid").value & "" & COL_DIV & RS.Fields("deptcd").value & ""
'            Call getbbs_ptinfo(sPtid, strSSN)
            .Col = TblColumn.tcSEX:     .value = GetSSN(sPtid) 'medGetP(strSSN, 1, COL_DIV) & "/" & medGetP(strSSN, 3, COL_DIV)
            .Col = TblColumn.tcBuss:    .value = RS.Fields("bussdiv").value & ""
            .Col = TblColumn.tcOcsOrdno: .value = RS.Fields("ocsordno").value & ""
            
            '���Ͱ� ó���� �ƴ� �÷������� �Ѿ�ö� ���
'            .Col = TblColumn.tcFilter: .value = Rs.Fields("filterfg").value & ""
            '���Ͱ� ó������ �����ð�쿡 ����
            Call GetFilterOrd(.Row)
            
            txtBldNo.SetFocus
            txtBldNo.SelStart = 0
            txtBldNo.SelLength = Len(txtBldNo)
        End If
        
Skip:
        
        .ReDraw = True
    End With
    
    Set RS = Nothing
    Set objSql = Nothing
    Set ObjABO = Nothing
    Set objBldDelivery = Nothing
End Sub

Private Function GetFilterOrd(ByVal vRow As Long) As String
'���Ͱ� ó������ �Ѿ���°�쿡 ���� �÷��� üũ���ش�.
    Dim RS As Recordset
    Dim strSQL As String
    Dim strFilter As String
    Dim strPtid As String
    Dim strOrdDt As String
        
    With tblBldList
        .Row = vRow
        
        .Col = TblColumn.tcPTID: strPtid = .value
        .Col = TblColumn.tcORDDT: strOrdDt = Replace(.value, "-", "")
    End With
        
    strSQL = "SELECT  CDVAL1 FROM " & T_COM003 & _
             " WHERE " & DBW("CDINDEX=", BC2_MATERIAL)
        
    Set RS = New Recordset
    RS.Open strSQL, DBConn
        
    If RS.EOF Then
        tblBldList.Row = vRow
        tblBldList.Col = TblColumn.tcFilter
        tblBldList.CellType = CellTypeStaticText
        tblBldList.value = ""
        Set RS = Nothing
        Exit Function
    End If
            
    Do Until RS.EOF
        strFilter = strFilter & "'" & RS.Fields("cdval1").value & "" & "',"
        RS.MoveNext
    Loop
    
    If strFilter <> "" Then
        strFilter = Mid(strFilter, 1, Len(strFilter) - 1)
    End If
    
    Set RS = Nothing
    
    strSQL = " SELECT  b.ptid,b.orddt,b.ordno,b.ordseq,b.ordcd,b.ocsordno,c.testnm,b.unitqty,a.bussdiv,a.deptcd " & _
             " FROM " & T_BBS001 & " c," & T_LAB102 & " b," & T_LAB101 & " a" & _
             " WHERE " & DBW("a.ptid=", strPtid) & _
             " AND " & DBW("a.orddt=", strOrdDt) & _
             " AND " & DBW("a.orddiv=", C_WORKAREA) & _
             " AND " & DBW("a.donefg=", enStsCd.StsCd_LIS_Accession) & _
             " AND a.ptid=b.ptid and a.orddt=b.orddt and a.ordno=b.ordno " & _
             " AND b.ordcd in(" & strFilter & ") " & _
             " AND b.ordcd=c.testcd " & _
             " AND (c.expdt='' or c.expdt is null)" & _
             " ORDER BY orddt desc,testcd "
             
    Set RS = New Recordset
    
    RS.Open strSQL, DBConn
    
    If RS.EOF Then
        tblBldList.Row = vRow
        tblBldList.Col = TblColumn.tcFilter
        tblBldList.CellType = CellTypeStaticText
        tblBldList.value = ""
    Else
        tblBldList.Row = vRow
        tblBldList.Col = TblColumn.tcFilter
        tblBldList.CellType = CellTypeCheckBox: tblBldList.TypeHAlign = TypeHAlignCenter
        tblBldList.value = 1
        tblBldList.Lock = True
    End If
    
    Set RS = Nothing
End Function

Private Function GetSSN(ByVal vPtID As String) As String
    Dim objSql As New clsPatient
    Dim RS As New Recordset

    RS.Open objSql.GetSQLPt(vPtID), DBConn

    GetSSN = RS.Fields("ssn").value & ""

    Set RS = Nothing
    Set objSql = Nothing
End Function

Private Sub cmdSave_Click()
    
    If tblBldList.DataRowCnt = 0 Then Exit Sub
    
    If txtRcvId.Text = "" Or lblRcvNm.Caption = "" Then
        MsgBox "�����ڸ� �����Ͻ��� ������� �Ͻʽÿ�.", vbInformation + vbOKOnly, "�����ڼ���"
        Exit Sub
    End If
    
    If BldDelivery = True Then
        MsgBox "���������� ��� ó���Ǿ����ϴ�.", vbInformation + vbOKOnly, "�����"
        
       If BLOOD_DEL_USED Then Call TransFusionPrint
        
        Call Clear
        txtBldNo.SetFocus
    Else
        MsgBox "���� �����߻�", vbInformation + vbOKOnly, "����Ͽ���"
    End If
    
End Sub
Private Sub cmdRePrint_Click()
    Call TransFusionPrint
End Sub
Private Function VfyNm(ByVal sBloodNo As String) As String
    Dim BldSrc As String
    Dim BldYY  As String
    Dim BldNo  As String
    Dim SSQL   As String
    Dim RS     As Recordset
    
    
    BldSrc = medGetP(sBloodNo, 1, "-")
    BldYY = medGetP(sBloodNo, 2, "-")
    BldNo = medGetP(sBloodNo, 3, "-")
    
    SSQL = " select b.empnm from " & T_COM006 & " b," & T_BBS302 & " a " & _
           " where " & _
                     DBW("a.bldsrc=", BldSrc) & _
           " and " & DBW("a.bldyy=", BldYY) & _
           " and " & DBW("a.bldno=", BldNo) & _
           " and a.vfyid=b.empid"
           
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    If Not RS.EOF Then
        VfyNm = RS.Fields("empnm").value & ""
    Else
        VfyNm = lblRcvNm.Caption
    End If
    
    Set RS = Nothing
    
    
End Function


Private Sub TransFusionPrint()
    
    Dim ii        As Integer
    Dim strBldNo  As String
    Dim strTestNm As String
    Dim strDelDt  As String
    Dim strDelNm  As String
    Dim strRcvNm  As String
    Dim strTmp    As String
    Dim strPtnm   As String
    Dim strPtid   As String
    Dim strABO    As String
    Dim strDept   As String     '����-ȣ��
    Dim strDeptCd As String     '�������
    Dim strDeptNm As String     '������
    Dim strSEX    As String
    Dim intFNum As Integer
    Dim strRfile As String
    Dim strRptPath As String
    
    Dim kk         As Integer
    
    'strDelNm = lblDeliveryNm.Caption
    strRcvNm = lblRcvNm.Caption
    strDelDt = Format(GetSystemDate, "YYYY-MM-DD")
    

    With tblBldList
        If .DataRowCnt < 11 Then
            kk = 10
        Else
            kk = ((.DataRowCnt \ 10) + 1) * 10
        End If
        
        For ii = 1 To kk
            .Row = ii
            .Col = TblColumn.tcBldNo: strBldNo = .value
            strDelNm = VfyNm(strBldNo)
            .Col = TblColumn.tcCOMPONM: strTestNm = .value
            .Col = TblColumn.tcVol:
            If .value <> "" Then strTestNm = strTestNm & Mid(.value, 1, Len(.value) - 2)
            
            If ii < 2 Then
                .Col = TblColumn.tcABO: strABO = .value
                .Col = TblColumn.tcPTID: strPtid = .value
                .Col = TblColumn.tcPTNM: strPtnm = .value
                .Col = TblColumn.tcDEPT:
                       strDept = medGetP(.value, 1, COL_DIV)
                       If strDept <> "" Then
                            If medGetP(.value, 2, COL_DIV) <> "" Then
                                strDept = strDept & "-" & medGetP(.value, 2, COL_DIV)
                            End If
                        End If
                        strDeptCd = medGetP(.value, 3, COL_DIV)
                        
                        If strDeptCd <> "" Then
'                            ObjComCode.DeptCd.Exists (strDeptCd)
'                            Call ObjComCode.DeptCd.KeyChange(strDeptCd)
                            strDeptCd = GetDeptNm(strDeptCd) 'ObjComCode.DeptCd.Fields("deptnm")
                        End If
                        
                        strDeptNm = medGetP(.value, 1, COL_DIV)
                        If strDeptNm <> "" Then
'                            ObjComCode.wardid.Exists (strDeptNM)
'                            Call ObjComCode.wardid.KeyChange(strDeptNM)
                            strDeptNm = GetWardNm(strDeptNm) 'ObjComCode.wardid.Fields("wardnm")
                        End If
                        
                .Col = TblColumn.tcSEX: strSEX = .value
            End If
            If ii > .DataRowCnt Then
                strDelDt = "": strDelNm = "": strRcvNm = ""
            End If
            strTmp = strTmp & strDelDt & vbTab
            strTmp = strTmp & strTestNm & vbTab
            strTmp = strTmp & strBldNo & vbTab
            strTmp = strTmp & strDelNm & vbTab
            strTmp = strTmp & strRcvNm & vbCr
        Next
        strTmp = Mid(strTmp, 1, Len(strTmp) - 1)
    End With
    
    
    strRfile = InstallDir & "BBS\RPT" & "\CrystalReport.txt"
    strRptPath = InstallDir & "BBS\RPT" & "\frmBBS303.rpt"
    
    
    intFNum = FreeFile
    Open strRfile For Output As #intFNum
    Print #intFNum, strTmp
    Close #intFNum
    
    With CReport
            .ParameterFields(0) = "Ptnm;" & strPtnm & ";TRUE"
            .ParameterFields(1) = "Ptid;" & strPtid & ";TRUE"
            .ParameterFields(2) = "ABO;" & strABO & ";TRUE"
            .ParameterFields(3) = "Dept;" & strDept & ";TRUE"
            .ParameterFields(4) = "Sex;" & strSEX & ";TRUE"
            .ParameterFields(5) = "DeptCd;" & strDeptCd & ";TRUE"
            .ParameterFields(6) = "DeptNm;" & strDeptNm & ";TRUE"
        .ReportFileName = strRptPath
        .RetrieveDataFiles
        .WindowState = crptMaximized
'        .Destination = crptToWindow
        .Destination = crptToPrinter


        .Action = 1
        .Reset
    End With
    Me.MousePointer = 0

End Sub

Private Function BldDelivery() As Boolean
    Dim objBldDelivery As clsBldDelivery
    Dim objPrgBar      As New clsProgress
    Dim RS             As Recordset
    Dim SSQL           As String
    Dim strPtid        As String
    Dim strBussdiv     As String
    Dim blnOutPatient  As Boolean
    
    Dim tmpStr         As String
    Dim ii             As Integer
    
    Dim today          As Date
    
    '----------------------------------------------------------------------------
    '���� String���� �ѱ��
    '----------------------------------------------------------------------------
    'bldsrc,bldyy,bldno,compocd,deliverydt,deliveryseq,deliverytm,deliveryid
    'rcvid,ptid,orddt,ordno,ordseq,rstseq,ordcd,localcd,irrafg,filter,retfg,expfg
    '----------------------------------------------------------------------------
    
    On Error GoTo SAVE_ERROR
    
    DBConn.BeginTrans
    objPrgBar.Container = medMain.stsBar
        
    today = GetSystemDate
    Set objBldDelivery = New clsBldDelivery

    With tblBldList
        objPrgBar.Max = .DataRowCnt
        
        For ii = 1 To .DataRowCnt
            .Row = ii
            .Col = TblColumn.tcBldNo:   tmpStr = medGetP(.value, 1, "-") & COL_DIV & medGetP(.value, 2, "-") & COL_DIV & Format(medGetP(.value, 3, "-"), "000000") & COL_DIV
            .Col = TblColumn.tcCOMPOCD: tmpStr = tmpStr & .value & COL_DIV
            .Col = TblColumn.tcDELDT:   tmpStr = tmpStr & Replace(.value, "-", "") & COL_DIV & _
                                                     "" & COL_DIV & Format(today, "HHMMSS") & COL_DIV & ObjMyUser.EmpId & COL_DIV & _
                                               txtRcvId & COL_DIV & C_WORKAREA & COL_DIV
            .Col = TblColumn.tcACCDT:   tmpStr = tmpStr & .value & COL_DIV
            .Col = TblColumn.tcACCSEQ:  tmpStr = tmpStr & .value & COL_DIV
            .Col = TblColumn.tcRSTSEQ:  tmpStr = tmpStr & .value & COL_DIV
            .Col = TblColumn.tcORDCD:   tmpStr = tmpStr & .value & COL_DIV & "" & COL_DIV
            .Col = TblColumn.tcRTP:     tmpStr = tmpStr & IIf(.value = "1", "1", "") & COL_DIV ' & "" & COL_DIV & "" & COL_DIV & "" & COL_DIV
            .Col = TblColumn.tcFilter:  tmpStr = tmpStr & IIf(.value = "1", "1", "") & COL_DIV & "" & COL_DIV & "" & COL_DIV
            .Col = TblColumn.tcPTID:    tmpStr = tmpStr & .value & COL_DIV: strPtid = .value
            .Col = TblColumn.tcORDDT:   tmpStr = tmpStr & Replace(.value, "-", "") & COL_DIV
            .Col = TblColumn.tcORDNO:   tmpStr = tmpStr & .value & COL_DIV
            .Col = TblColumn.tcORDSEQ:  tmpStr = tmpStr & .value & COL_DIV
            .Col = TblColumn.tcBuss:    tmpStr = tmpStr & .value & COL_DIV: strBussdiv = .value
            .Col = TblColumn.tcOcsOrdno: tmpStr = tmpStr & .value
            
            
            If strBussdiv = enBussDiv.BussDiv_InPatient Then
                SSQL = " select * from " & T_HIS002 & _
                     " where " & DBW(F_PTID, strPtid, 2) & _
                     " and ( " & F_BEDOUTDT & " ='' or " & F_BEDOUTDT & " is null) "
                
                Set RS = Nothing
                Set RS = New Recordset
                RS.Open SSQL, DBConn
                
                If RS.EOF Then  '�Կ�ȯ��
                    blnOutPatient = True
                End If
                Set RS = Nothing
            End If
            
            '�Կ�ó���ϰ�� ���� �Կ�ȯ�ڿ� ���ؼ��� ����۾��� �Ѵ�.
            If blnOutPatient = False Then
                BldDelivery = objBldDelivery.BldDelivery(tmpStr)
            Else
                MsgBox "���� ����� ȯ���Դϴ�.���ó���� ����Ͻ��� ����۾��� �����ϼ���.", vbInformation + vbOKOnly, "info"
            End If
            
            If BldDelivery = False Then
                GoTo SAVE_ERROR
            End If
            objPrgBar.value = ii
        Next
    End With
    
    DBConn.CommitTrans
    
    BldDelivery = True
    Set objPrgBar = Nothing
    Set objBldDelivery = Nothing
    Exit Function
            
SAVE_ERROR:
    DBConn.RollbackTrans
    BldDelivery = False
    Set objPrgBar = Nothing
    Set objBldDelivery = Nothing
    If Err.Number > 0 Then
        MsgBox Err.Description, vbExclamation
    End If
End Function

Private Function SetNewTest(ByVal BldSrc As String, ByVal BldYY As String, ByVal BldNo As String, _
                            ByVal CompoCd As String, ByVal volume As String, ByVal TestDiv As String) As String

    Dim Cnt As Long
    Dim aryOrdCd() As String
    Dim today As Date
    Dim objBldDelivery As clsBldDelivery
    Dim i As Long
    
    today = GetSystemDate
    
    Set objBldDelivery = New clsBldDelivery
    Cnt = objBldDelivery.GetOrdCd(BldSrc, BldYY, BldNo, CompoCd, Format(today, PRESENTDATE_FORMAT), volume, TestDiv, aryOrdCd)
    Set objBldDelivery = Nothing
    
    If Cnt > 0 Then
        SetNewTest = medGetP(aryOrdCd(0), 1, " ")
    End If
    
'    lstNewTest.Clear
'    If cnt > 0 Then
'        For i = 1 To cnt
'            lstNewTest.AddItem aryOrdCd(i - 1)
'        Next i
'        onPgm = True
'        If lstNewTest.ListCount = 1 Then lstNewTest.Selected(0) = True
'        onPgm = False
'    End If
End Function

Private Sub Clear()
    Dim ii As Integer
    
    txtRcvId.Text = ""
    txtBldNo.Text = ""
    txtCompID.Text = ""
    
    lblRcvNm.Caption = ""
    lblCompoCd.Caption = ""
    lblCompoNm.Caption = ""
    lblCompNm.Caption = ""
    lblCompNm.Caption = ""
    lblSexAge.Caption = ""
    lblABO.Caption = ""
    
    lblDeliveryNm.Visible = True
    lblDeliveryNm.Caption = ObjSysInfo.EmpNm
    cboCompo.Clear
    With tblBldList
        For ii = 1 To .MaxRows
            .Row = ii
            .Col = TblColumn.tcRTP
            .CellType = CellTypeStaticText
            .Col = TblColumn.tcFilter
            .CellType = CellTypeStaticText
        Next
    End With
    Call medClearTable(tblBldList, False, False)
    
End Sub
Private Sub cmdTransStatPaper_Click()
    Call PrintIntionlize
    Call PrintBloodStat
End Sub

Private Sub PrintIntionlize()
    PrtLeft = 5
    LineSpace = 6
    lngCurYPos = 20
    Printer.Font = "����ü"
    Printer.FontSize = 9
    Printer.Orientation = vbPRORPortrait '/* ����
    Printer.ScaleMode = vbMillimeters
    Twidth = Printer.ScaleWidth
    LastLineYpos = Printer.ScaleHeight             '����������Y��ġ
End Sub

Private Sub PrintBloodStat()
    Dim lngX1 As Long
    Dim lngX2 As Long
    Dim lngX3 As Long

    lngX1 = 25
    lngX2 = 85
    lngX3 = 145
    
    Printer.FontSize = 20: Printer.FontBold = True
    Call Print_Setting("���޼�����û��", PrtLeft, lngCurYPos, Twidth, "C", "C", False)
    Printer.FontSize = 10: Printer.FontBold = False
    lngCurYPos = lngCurYPos + 30
    
    Call Print_Setting("ȯ�ڼ��� : ", lngX1, LineSpace, , , "C", False)
    Call Print_Setting("��Ϲ�ȣ : ", lngX2, LineSpace, , , "C", False)
    Call Print_Setting("�������� : ", lngX3, LineSpace, , "L", "C")
    Call Print_Setting("��    �� : ", lngX1, LineSpace, , , "C", False)
    Call Print_Setting("�� �� �� : ", lngX2, LineSpace, , , "C")
    Call Print_Setting("�� �� �� : ", lngX1, LineSpace, Twidth, "L", "C")
    
    Call Print_Setting("��ġ�� ���� : ", lngX1, LineSpace, Twidth, "L", "C", False)
    Call Print_Setting("�� �� ó : ", lngX2, LineSpace, Twidth, "L", "C")
    
    Call Print_Setting("������û���׼��� : ", lngX1, LineSpace, Twidth, "L", "C")
    Printer.FontBold = True
    Call Print_Setting("���׹�ȣ : ", lngX1, LineSpace, Twidth, "L", "C")
    Printer.FontBold = False
    lngCurYPos = lngCurYPos + 10
    
    
    Call Print_Setting("��ġ�� : ", lngX1, LineSpace, Twidth, "L", "C")
    Call Print_Setting(" ������ ��� ȯ�ڿ� ���� ���޼����� �� �ʿ��Ͽ� �ʼ� ������ �˻簡 �Ϸ���� ��������", _
                        lngX1, LineSpace, , "L", "C")
    
    Call Print_Setting("�� �ұ��ϰ� ���� ���׼����� �����Ͽ� �ٰ��� ������ å���Ͽ� ��û�մϴ�. " _
                       , lngX1, LineSpace, Twidth, "L", "C")
                       
    lngCurYPos = lngCurYPos + 5
    
    Call Print_Setting("�� O�� ����(ABO, Rh typing & crossmatching ����", _
                        lngX1, LineSpace, Twidth, "L", "C")
    Call Print_Setting("�� ABO, Rh typing �˻����, & crossmatching ����", _
                        lngX1, LineSpace, Twidth, "L", "C")
    Call Print_Setting("�� ABO, Rh typing & immedite bromelin  crossmatching(15�мҿ�)", _
                        lngX1, LineSpace, Twidth, "L", "C")
    Call Print_Setting("�� ABO, Rh typing & bromelin crossmatching(30�мҿ�)", _
                        lngX1, LineSpace, Twidth, "L", "C")
    lngCurYPos = lngCurYPos + 5
    
    Call Print_Setting("    20  ��    ��    ��" & "                            �� ġ ��             (��)", lngX1, LineSpace, , "L", "C")
    lngCurYPos = lngCurYPos + 5
    
    Call Print_Setting("ȯ��/��ȣ��  : ", lngX1, LineSpace, , "L", "C")
    Call Print_Setting(" ������ ��� ȯ�ڿ� ���� ���޼����� �� �ʿ��Ͽ� �ʼ����� ���� �˻簡 �Ϸ���� ����", _
                         lngX1, LineSpace, , "L", "C")
    Call Print_Setting("������ ������ �����ϸ�, �̿� ���õǾ� ���ߵɼ� �ִ� ���ۿ뿡 ���Ͽ� �ƹ��� ���Ǹ� ��", _
                         lngX1, LineSpace, , "L", "C")
    Call Print_Setting("������ �ʰڽ��ϴ�.", lngX1, LineSpace, , "L", "C")
    
    lngCurYPos = lngCurYPos + 10
    Call Print_Setting("    20  ��    ��    ��" & "                         ȯ��/��ȣ��             (��)", lngX1, LineSpace, , "L", "C")
    lngCurYPos = lngCurYPos + 10
    
    Call Print_Setting("* �������࿡���� ���� ���� ��� �Ŀ��� crossmatching �˻縦 �������� ��Ģ���� �մϴ�.", _
                        lngX1, LineSpace, , "L", "C")
    Call Print_Setting("* ��ü�� �ݵ�� ���� �ֽñ� �ٶ��ϴ�.", lngX1, LineSpace, , "L", "C")
    
    lngCurYPos = lngCurYPos + 30
    
    Printer.FontBold = True
    Call Print_Setting("Department of Clinical Pathology, Gil Medical Center", PrtLeft, LineSpace, Twidth, "C", "C")
    
    Call Print_Setting("������ȭ : (032)460-3938", PrtLeft, LineSpace, Twidth, "C", "C")
    
    Call Print_Setting("405-760 ��õ������ ������ ������ 1198���� ��õ�ǰ����б溴�� ���ܰ˻����а� ��������", PrtLeft, LineSpace, Twidth, "C", "C", False)
    
    Printer.FontBold = False
    
    Printer.EndDoc
    
End Sub


