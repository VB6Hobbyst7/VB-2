VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{9167B9A7-D5FA-11D2-86CA-00104BD5476F}#5.0#0"; "DRCTL1.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frmBBS101 
   BackColor       =   &H00DBE6E6&
   Caption         =   "ó����"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14580
   Icon            =   "frmBBS101.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   14580
   WindowState     =   2  '�ִ�ȭ
   Begin VB.ListBox lstReason 
      Height          =   2940
      Left            =   10695
      TabIndex        =   52
      Top             =   3180
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.ListBox lstOcd 
      Height          =   3660
      Left            =   10560
      TabIndex        =   53
      Top             =   3600
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.ListBox lstTestCd 
      Height          =   2940
      Left            =   4500
      TabIndex        =   51
      Top             =   3180
      Visible         =   0   'False
      Width           =   4155
   End
   Begin MedControls1.LisLabel lblReaction 
      Height          =   315
      Left            =   2400
      TabIndex        =   49
      Top             =   45
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      BackColor       =   12640511
      ForeColor       =   255
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
      Caption         =   "Reaction"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel lblInfection 
      Height          =   315
      Left            =   1980
      TabIndex        =   50
      Top             =   45
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   556
      BackColor       =   12640511
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "@"
      Appearance      =   0
   End
   Begin VB.PictureBox picReqDtTm 
      Height          =   390
      Left            =   9120
      ScaleHeight     =   330
      ScaleWidth      =   2055
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   2100
      Visible         =   0   'False
      Width           =   2115
      Begin MSComCtl2.DTPicker dtpReqDt 
         Height          =   315
         Left            =   0
         TabIndex        =   47
         Top             =   0
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   62128131
         CurrentDate     =   36868
      End
      Begin MSComCtl2.DTPicker dtpReqTm 
         Height          =   315
         Left            =   1260
         TabIndex        =   48
         Top             =   0
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "HH:mm"
         Format          =   62128131
         UpDown          =   -1  'True
         CurrentDate     =   36868
      End
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   315
      Index           =   1
      Left            =   75
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   45
      Width           =   14235
      _ExtentX        =   25109
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
      Caption         =   "  ȯ�� �⺻ ����"
      Appearance      =   0
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DBE6E6&
      Height          =   1095
      Left            =   75
      TabIndex        =   32
      Top             =   300
      Width           =   14235
      Begin VB.CommandButton cmdPopUp 
         BackColor       =   &H00C7D8D8&
         Caption         =   "..."
         Height          =   315
         Index           =   1
         Left            =   9345
         Style           =   1  '�׷���
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   240
         Width           =   360
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   7
         Left            =   4995
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   660
         Width           =   1050
         _ExtentX        =   1852
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
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   8
         Left            =   11070
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   660
         Width           =   1050
         _ExtentX        =   1852
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
         Caption         =   "����"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   5
         Left            =   7305
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   240
         Width           =   1050
         _ExtentX        =   1852
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
         Caption         =   "ó����"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   6
         Left            =   7305
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   660
         Width           =   1050
         _ExtentX        =   1852
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
         Caption         =   "�����"
         Appearance      =   0
      End
      Begin VB.Frame fraWard 
         BackColor       =   &H00DBE6E6&
         BorderStyle     =   0  '����
         Height          =   375
         Left            =   11025
         TabIndex        =   42
         Top             =   600
         Width           =   2955
         Begin VB.TextBox txtWardID 
            Alignment       =   2  '��� ����
            Appearance      =   0  '���
            Height          =   315
            Index           =   2
            Left            =   2385
            TabIndex        =   8
            Text            =   "7123456"
            Top             =   60
            Width           =   495
         End
         Begin VB.TextBox txtWardID 
            Alignment       =   2  '��� ����
            Appearance      =   0  '���
            Height          =   315
            Index           =   1
            Left            =   1845
            TabIndex        =   7
            Text            =   "7123456"
            Top             =   60
            Width           =   495
         End
         Begin VB.CommandButton cmdPopUp 
            BackColor       =   &H00C7D8D8&
            Caption         =   "..."
            Height          =   315
            Index           =   3
            Left            =   660
            Style           =   1  '�׷���
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   60
            Width           =   360
         End
         Begin VB.TextBox txtWardID 
            Alignment       =   2  '��� ����
            Appearance      =   0  '���
            Height          =   315
            Index           =   0
            Left            =   1125
            TabIndex        =   6
            Top             =   60
            Width           =   675
         End
      End
      Begin VB.CommandButton cmdPopUp 
         BackColor       =   &H00C7D8D8&
         Caption         =   "..."
         Height          =   315
         Index           =   2
         Left            =   9345
         Style           =   1  '�׷���
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   660
         Width           =   360
      End
      Begin VB.TextBox txtDeptcd 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         Height          =   315
         Left            =   8385
         TabIndex        =   5
         Top             =   660
         Width           =   915
      End
      Begin VB.TextBox txtOrderDoctor 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         Height          =   315
         Left            =   8385
         TabIndex        =   4
         Top             =   240
         Width           =   915
      End
      Begin VB.TextBox txtPtid 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         Height          =   315
         Left            =   1155
         TabIndex        =   0
         Text            =   "7123456"
         ToolTipText     =   "ȯ��ID�� �Է��Ͻÿ�"
         Top             =   240
         Width           =   1245
      End
      Begin VB.CommandButton cmdPtId 
         BackColor       =   &H00C7D8D8&
         Caption         =   "..."
         Height          =   315
         Left            =   2415
         Style           =   1  '�׷���
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   240
         Width           =   360
      End
      Begin DRcontrol1.DrLabel lblPtNm 
         Height          =   315
         Left            =   1155
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   660
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
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
      Begin DRcontrol1.DrLabel DrLabel2 
         Height          =   330
         Left            =   3900
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   240
         Width           =   3315
         _ExtentX        =   5847
         _ExtentY        =   582
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
         TextPosition    =   4
         Caption         =   ""
         Begin VB.OptionButton optOrderDiv 
            BackColor       =   &H00DBE6E6&
            Caption         =   "����"
            Enabled         =   0   'False
            Height          =   180
            Index           =   2
            Left            =   2310
            TabIndex        =   3
            Top             =   75
            Width           =   735
         End
         Begin VB.OptionButton optOrderDiv 
            BackColor       =   &H00DBE6E6&
            Caption         =   "�ܷ�"
            Height          =   180
            Index           =   0
            Left            =   270
            TabIndex        =   1
            Top             =   75
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.OptionButton optOrderDiv 
            BackColor       =   &H00DBE6E6&
            Caption         =   "�Կ�"
            Height          =   180
            Index           =   1
            Left            =   1290
            TabIndex        =   2
            Top             =   75
            Width           =   735
         End
      End
      Begin DRcontrol1.DrLabel lblOrderDoctor 
         Height          =   315
         Left            =   9765
         TabIndex        =   38
         Top             =   240
         Width           =   4140
         _ExtentX        =   7303
         _ExtentY        =   556
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
      Begin DRcontrol1.DrLabel lblSexAge 
         Height          =   315
         Left            =   3900
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   660
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
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
      Begin DRcontrol1.DrLabel lblDob 
         Height          =   315
         Left            =   6060
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   660
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
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
      Begin DRcontrol1.DrLabel lblDeptNm 
         Height          =   315
         Left            =   9765
         TabIndex        =   41
         Top             =   660
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
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
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   0
         Left            =   75
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   240
         Width           =   1050
         _ExtentX        =   1852
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
         Caption         =   "ȯ��ID"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   2
         Left            =   90
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   660
         Width           =   1050
         _ExtentX        =   1852
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
         Caption         =   "����"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   3
         Left            =   2820
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   240
         Width           =   1050
         _ExtentX        =   1852
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
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   4
         Left            =   2820
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   660
         Width           =   1050
         _ExtentX        =   1852
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
         Caption         =   "����/����"
         Appearance      =   0
      End
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   315
      Left            =   75
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1485
      Width           =   4350
      _ExtentX        =   7673
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
      Caption         =   "  ó�� �ڵ� ����Ʈ"
      Appearance      =   0
   End
   Begin VB.CommandButton cmdOrder 
      BackColor       =   &H00F4F0F2&
      Caption         =   "ó �� (&O)"
      Height          =   510
      Left            =   9180
      Style           =   1  '�׷���
      TabIndex        =   22
      Tag             =   "0"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdOCollection 
      BackColor       =   &H00F4F0F2&
      Caption         =   "ó��,����(&S)"
      Height          =   510
      Left            =   10500
      Style           =   1  '�׷���
      TabIndex        =   24
      Tag             =   "0"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "ȭ������(&C)"
      Height          =   510
      Left            =   11820
      Style           =   1  '�׷���
      TabIndex        =   26
      Tag             =   "0"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "�� ��(&X)"
      Height          =   510
      Left            =   13140
      Style           =   1  '�׷���
      TabIndex        =   28
      Tag             =   "0"
      Top             =   8535
      Width           =   1320
   End
   Begin DRcontrol1.DrFrame fraOrderInformation 
      Height          =   915
      Left            =   4485
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1500
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   1614
      BorderStyle     =   0   'False
      Appearance      =   0
      Title           =   " �� ó�� ����"
      TitlePos        =   1
      BackColor       =   14411494
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmdRight 
         BackColor       =   &H00C7D8D8&
         Caption         =   "�˻��߰�"
         Height          =   375
         Left            =   180
         Style           =   1  '�׷���
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   420
         Width           =   1440
      End
      Begin VB.CheckBox chkStat 
         BackColor       =   &H00DBE6E6&
         Caption         =   "���޿���"
         Height          =   195
         Left            =   8040
         TabIndex        =   13
         Top             =   540
         Width           =   1095
      End
      Begin VB.TextBox txtReceiptNo 
         Appearance      =   0  '���
         Height          =   315
         Left            =   4260
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   480
         Width           =   3375
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   9
         Left            =   3195
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   480
         Width           =   1050
         _ExtentX        =   1852
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
         Caption         =   "��������ȣ"
         Appearance      =   0
      End
   End
   Begin FPSpread.vaSpread tblTestList 
      Height          =   4500
      Left            =   4485
      TabIndex        =   9
      Top             =   2505
      Width           =   9945
      _Version        =   196608
      _ExtentX        =   17542
      _ExtentY        =   7938
      _StockProps     =   64
      BackColorStyle  =   1
      ButtonDrawMode  =   1
      DisplayRowHeaders=   0   'False
      EditModePermanent=   -1  'True
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
      MaxCols         =   12
      MaxRows         =   18
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      SpreadDesigner  =   "frmBBS101.frx":076A
      UserResize      =   0
   End
   Begin MSComctlLib.ListView lvwOrderList 
      Height          =   7170
      Left            =   75
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1845
      Width           =   4350
      _ExtentX        =   7673
      _ExtentY        =   12647
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ó���"
         Object.Width           =   5644
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ó���ڵ�"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "pheresis����"
         Object.Width           =   0
      EndProperty
   End
   Begin MedControls1.LisLabel LisLabel3 
      Height          =   315
      Left            =   4485
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   7080
      Width           =   9960
      _ExtentX        =   17568
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
      Caption         =   "  ��ü ����"
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   1035
      Left            =   4485
      TabIndex        =   15
      Top             =   7320
      Width           =   9975
      Begin VB.ComboBox cboLeg 
         Height          =   300
         ItemData        =   "frmBBS101.frx":1BA3
         Left            =   3060
         List            =   "frmBBS101.frx":1BA5
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   54
         Top             =   540
         Width           =   1050
      End
      Begin VB.TextBox txtRowNo 
         Alignment       =   2  '��� ����
         Height          =   315
         Left            =   4140
         TabIndex        =   18
         Top             =   540
         Width           =   675
      End
      Begin VB.TextBox txtColNo 
         Alignment       =   2  '��� ����
         Height          =   315
         Left            =   4800
         TabIndex        =   20
         Top             =   540
         Width           =   675
      End
      Begin VB.CheckBox chkSPos 
         BackColor       =   &H00DBE6E6&
         Caption         =   "��ü������� �ڵ�����"
         Height          =   555
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   1395
      End
      Begin MedControls1.LisLabel lblDestCenterNm 
         Height          =   315
         Left            =   8940
         TabIndex        =   29
         Top             =   240
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   556
         BackColor       =   12648447
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
      Begin MedControls1.LisLabel lblDestCenterCd 
         Height          =   315
         Left            =   8340
         TabIndex        =   31
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         BackColor       =   12648447
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
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "�˻� ����"
         Height          =   180
         Left            =   7380
         TabIndex        =   30
         Top             =   300
         Width           =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "��ü�������"
         Height          =   180
         Left            =   1800
         TabIndex        =   23
         Top             =   450
         Width           =   1080
      End
      Begin VB.Label Label5 
         Alignment       =   2  '��� ����
         BackStyle       =   0  '����
         BorderStyle     =   1  '���� ����
         Caption         =   "Rack"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3060
         TabIndex        =   21
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label Label6 
         Alignment       =   2  '��� ����
         BackStyle       =   0  '����
         BorderStyle     =   1  '���� ����
         Caption         =   "Row"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4140
         TabIndex        =   19
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label7 
         Alignment       =   2  '��� ����
         BackStyle       =   0  '����
         BorderStyle     =   1  '���� ����
         Caption         =   "Col"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4800
         TabIndex        =   17
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label8 
         BackStyle       =   0  '����
         BorderStyle     =   1  '���� ����
         Height          =   600
         Left            =   1620
         TabIndex        =   25
         Top             =   240
         Width           =   1440
      End
   End
End
Attribute VB_Name = "frmBBS101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum TblColumn
    tcTESTNM = 1
    tcTESTCD
    tcQTY
    tcREQDTTM
    tcREASONNM
    tcOPNM
    tcIRR
    tcFilter
    tcREASONCD
    tcOPCD
    tcISOP
    tcISPHER
End Enum

Private WithEvents GetPtInfo    As frmPtInfo
Attribute GetPtInfo.VB_VarHelpID = -1
'Private WithEvents objPtInfo As clsPatient
Private WithEvents objListPop   As clsPopUpList
Attribute objListPop.VB_VarHelpID = -1
'Private WithEvents mnuPopup     As Menu
'Private WithEvents mnuDelete    As Menu
Private WithEvents objPop As clsPopupMenu
Attribute objPop.VB_VarHelpID = -1
Private Const MENU_DEL& = 1
Private BarSpcNum As String


Private Sub chkSPos_Click()
    
    If chkSPos.value = 1 Then
        txtRowNo = ""
        txtColNo = ""
        txtRowNo.Locked = True
        txtColNo.Locked = True
        txtRowNo.BackColor = Me.BackColor
        txtColNo.BackColor = Me.BackColor
    Else
        txtRowNo.Locked = False
        txtColNo.Locked = False
        txtRowNo.BackColor = RGB(255, 255, 255)
        txtColNo.BackColor = RGB(255, 255, 255)
    End If
End Sub

Private Sub cmdClear_Click()

    Call Clear   '�ʱ�ȭ
    txtPtId.Enabled = True
    txtPtId = "": txtPtId.SetFocus
End Sub

Private Sub cmdExit_Click()     '����
    Unload Me
End Sub

Private Sub BarCode_Print()
    Dim objSQL     As clsBBSCollection
    Dim objBar     As clsBarcode
    Dim strBuildNm As String        '�ǹ��̸�
    Dim strW_Dept  As String
    Dim strColDt   As String
    Dim strColTm   As String
    Dim strAccSeq  As String         'SpcYy-SpcNo ������ ��ü��ȣ

    '���ڵ�
    Set objBar = New clsBarcode
    Set objSQL = New clsBBSCollection
    
'    Set objBar.MyDB = dbconn
    Set objBar.TableInfo = New clsTables
    Set objBar.FieldInfo = New clsFields
    
    strW_Dept = txtWardId(0)
    
    If strW_Dept = "" Then strW_Dept = txtDeptcd
    
    If strW_Dept <> "" Then
        If txtWardId(1) <> "" Then strW_Dept = strW_Dept & "/" & txtWardId(1)
    End If
    
    strColDt = Format(Mid(Format(GetSystemDate, PRESENTDATE_FORMAT), 5, 4), "00/00")
    strColTm = Format(GetSystemDate, "HH:MM")

    '��ü��ȣ ��� : 2001.2.8 �߰�
    strAccSeq = Mid(BarSpcNum, 1, 2) & "-" & Format(Mid(BarSpcNum, 3), "000000000")
    strAccSeq = Format(strAccSeq, String(11, "@"))
    strBuildNm = BBSName
    
    '���ڵ� ���
    objBar.Label_PrintOut strBuildNm, "XM", "", strAccSeq, BarSpcNum, txtPtId.Text, lblPtNm.Caption, "", "", _
                          chkStat.value, strW_Dept, strColDt, strColTm, "", 1
                          ', IIf(chkStat.value = 1, True, False)
   
    Set objBar = Nothing
    Set objSQL = Nothing
End Sub

Private Sub cmdOCollection_Click()  'ó��/���� ���ÿ�
    Dim AccessOk As Boolean
    
    If tblTestList.DataRowCnt = 0 Then Exit Sub
    
    
    '---------- ���� ��� Check -------------------------------------------------------------
    If chkSPos.value = 0 Then
        If cboLeg.Text = "" Or txtRowNo = "" Or txtColNo = "" Then
            MsgBox "��������� �Է��� �����Ǿ����ϴ�.", vbCritical + vbOKOnly, "������� ����"
            Exit Sub
        End If
    Else
        If cboLeg.ListIndex < 0 Then
            MsgBox "������� �ڵ��ο��� ��� �ݵ�� Leg�� �����ϼž� �մϴ�.", vbCritical + vbOKOnly, "������� Leg����"
            Exit Sub
        End If
    End If
    
    '---------- �˻� �ǹ� Check -------------------------------------------------------------
    If lblDestCenterCd.Caption = ObjSysInfo.BuildingCd Then
        AccessOk = True
    Else
        If MsgBox("�� �ǹ����� �˻��� �� �����ϴ�. ó���ϸ� �Ͻð����ϱ�?", vbYesNo + vbDefaultButton2 + vbQuestion, Me.Caption) = vbNo Then
            Exit Sub
        Else
            AccessOk = False
        End If
    End If
    
    
    '--------- ó��,ä��,���� ���� ----------------------------------------------------------
    If OnlyOrderSaver(AccessOk) = False Then
        MsgBox "ó�� ���� Error �Դϴ�.", vbCritical + vbOKOnly, "ó��/����"
        Exit Sub
    End If
    
    
    '--------- ä���Ǿ����� ���ڵ带 ����Ѵ�. ----------------------------------------------
    If BarSpcNum <> "" Then
        BarCode_Print
    Else
        'MsgBox "�̹� ��ü�� �����Ͽ� ���ڵ尡 ��µ��� �ʽ��ϴ�.", vbInformation + vbOKOnly, "���ڵ����"
    End If
    
    
    '--------- ������ ó�� ------------------------------------------------------------------
    Call Clear
    txtPtId = ""
    txtPtId.SetFocus
End Sub

Private Sub cmdOrder_Click()        'ó�游
    If tblTestList.DataRowCnt = 0 Then Exit Sub
    
    If OnlyOrderSaver(False) = True Then
        Clear
        txtPtId = ""
        txtPtId.SetFocus
    Else
        MsgBox "ó�� Error �Դϴ�.", vbCritical + vbOKOnly, Me.Caption
    End If

End Sub


Private Function Valadiation_Check() As Boolean
    Dim i As Integer
    
    If txtPtId.Text = "" Then
        MsgBox "ȯ�������� �����Ǿ����ϴ�.", vbCritical + vbOKOnly, "�Է���������"
        Exit Function
    End If
    
    If txtDeptcd.Text = "" Then
        MsgBox "����� ������ �����Ǿ����ϴ�.", vbCritical + vbOKOnly, "�Է���������"
        Exit Function
    End If
    
    If optOrderDiv(1).value = True Then
        If txtWardId(0) = "" Then
            MsgBox "���������� �����Ǿ����ϴ�.", vbCritical + vbOKOnly, "�Է���������"
            Exit Function
        End If
    End If
    
    With tblTestList
        For i = 1 To .DataRowCnt
            .Row = i
            .Col = TblColumn.tcREQDTTM
            If .value = "" Then
                MsgBox "������������ �����Ǿ����ϴ�..", vbCritical + vbOKOnly, "�Է���������"
                Exit Function
            End If
            .Col = TblColumn.tcREASONNM
            If .value = "" Then
                MsgBox "���������� �����Ǿ����ϴ�.", vbCritical + vbOKOnly, "�Է���������"
                Exit Function
            End If
        Next
    End With
    Valadiation_Check = True
End Function

'#������/��������/�������� sort(������ ó������� ���� ��ȣ�� �ֱ����ؼ�)
Private Sub OrderSort()
    With tblTestList
        .ReDraw = False
        .SortBy = SortByRow
        .SortKey(1) = TblColumn.tcREQDTTM
        .SortKey(2) = TblColumn.tcREASONCD
        .SortKey(3) = TblColumn.tcOPCD

        .SortKeyOrder(1) = SortKeyOrderAscending
        .SortKeyOrder(2) = SortKeyOrderAscending
        .SortKeyOrder(3) = SortKeyOrderAscending

        .Col = 1:  .COL2 = .MaxCols
        .Row = 1:  .Row2 = .MaxRows
        .BlockMode = True
        .Action = 25
        .BlockMode = False
        .ReDraw = True
    End With
End Sub

'#ȭ����� ��� �˾�����Ʈ
Private Sub cmdPopUp_Click(Index As Integer)

    If Index <> 0 Then If txtPtId.Text = "" Then Exit Sub
    
    Set objListPop = New clsPopUpList
    With objListPop
        .Connection = DBConn
'        .BackColor = Me.BackColor
        Select Case Index
            'ȯ�ڰ˻�
            Case 0:
                GetPtInfo.Show 1
'                Set objPtInfo = New clsPatient
'                objPtInfo.LoadSearchForm
            Case 1
            '��ġ���ڵ� �ҷ�����
                .FormCaption = "��ġ�� ��ȸ": .ColumnHeaderText = "�ڵ�;�ڵ��"
                txtOrderDoctor.Text = "": lblOrderDoctor.Caption = ""
'                .Width = .Width + 700
                Call .LoadPopUp(GetSQLDoctList) ', 2850, 7650)
                If .SelectedString <> "" Then
                    txtOrderDoctor.Text = medGetP(.SelectedString, 1, ";")
                    lblOrderDoctor.Caption = medGetP(.SelectedString, 2, ";")
                End If
                
            Case 2
                .FormCaption = "�������ȸ": .ColumnHeaderText = "�ڵ�;�ڵ��"
'                .Width = .Width + 300:   .ColSize(0) = 1000
                txtDeptcd.Text = "": lblDeptNm.Caption = ""
                Call .LoadPopUp(GetSQLDeptList) ', 2350, 7650) ', ObjBBSComCode.DeptCd)
                If .SelectedString <> "" Then
                    txtDeptcd.Text = medGetP(.SelectedString, 1, ";")
                    lblDeptNm.Caption = medGetP(.SelectedString, 2, ";")
                    If fraWard.Enabled = True Then
                        txtWardId(0).SetFocus
                    Else
                        tblTestList.SetFocus
                    End If
                End If
            Case 3
                .FormCaption = "������ȸ": .ColumnHeaderText = "�ڵ�;�ڵ��"
'                .Width = .Width + 300:   .ColSize(0) = 1000
                txtWardId(0).Text = ""
                Call .LoadPopUp(GetSQLWardList) ', 2350, 7650) ', ObjBBSComCode.wardid)
                If .SelectedString <> "" Then
                    txtWardId(0).Text = medGetP(.SelectedString, 1, ";")
                    txtWardId(1).SetFocus
                Else
                    txtWardId(0).SetFocus
                End If
                Call SetDestCenter
        End Select
    End With
    Set objListPop = Nothing
    
End Sub

Private Sub cmdPtId_Click()
    GetPtInfo.Show 1
'    Set objPtInfo = New clsPatient
'    Call objPtInfo.LoadSearchForm
    
End Sub

Private Sub dtpReqDt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dtpReqTm.SetFocus   '������ ������.
    End If
End Sub

Private Sub dtpReqTm_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        With tblTestList
            .Row = .ActiveRow
            .Col = TblColumn.tcREQDTTM
            .value = Format(dtpReqDt, "YYYY-MM-DD") & " " & Format(dtpReqTm, "HH:MM")
            
            .Row = .ActiveRow
            .Col = TblColumn.tcREASONNM
            .Action = ActionActiveCell
            
            .SetFocus
            
            picReqDtTm.Visible = False
        End With
    End If
End Sub

Private Sub Form_Activate()
    medMain.lblSubMenu.Caption = Me.Caption
End Sub

Private Sub QueryTestCode()
    Dim iTmx        As ListItem
    Dim objBBSsql   As clsGetSqlStatement
    Dim objPrgBar   As clsProgress
    Dim i           As Long
    Dim Rs          As Recordset
    Dim rs1         As Recordset
    Dim rs2         As Recordset
    
    lvwOrderList.ListItems.Clear
    lstTestCd.Clear
    
    Set objBBSsql = New clsGetSqlStatement
    Set objPrgBar = New clsProgress
'    Set objPrgBar.StatusBar = medMain.stsBar
    objPrgBar.Container = MainFrm.stsBar
    
    Set Rs = objBBSsql.Get_TransOrderList
    Set rs1 = objBBSsql.GEtReasonList
    Set rs2 = objBBSsql.GetOPCodeList
    
    
    objPrgBar.Min = 1
    objPrgBar.Max = Rs.RecordCount + rs1.RecordCount + rs2.RecordCount
    
    '�˻��׸�
    With Rs
        For i = 1 To .RecordCount
        
            objPrgBar.value = i
        
            Set iTmx = lvwOrderList.ListItems.Add()
            iTmx.Text = .Fields("testnm").value & ""
            iTmx.SubItems(1) = .Fields("testcd").value & ""
            iTmx.SubItems(2) = .Fields("testdiv").value & ""
            
            lstTestCd.AddItem .Fields("testcd").value & "" & vbTab & _
                              .Fields("testnm").value & "" & vbTab & _
                              .Fields("testdiv").value & "" & vbTab & _
                              "0"
            Rs.MoveNext
        Next i
    End With
    
    '��������
    With rs1
        For i = 1 To .RecordCount
            
            objPrgBar.value = i
            
            lstReason.AddItem .Fields("field1").value & "" & vbTab & _
                              .Fields("cdval1").value & "" & vbTab & _
                              .Fields("field2").value & ""
            .MoveNext
        Next i
    End With
    
    '�����ڵ�
    With rs2
        For i = 1 To .RecordCount
            
            objPrgBar.value = i
            
            lstOcd.AddItem Trim(Trim0(.Fields("onm").value & "")) & vbTab & _
                           Trim(.Fields("ocd").value & "")
            .MoveNext
        Next i
    End With

    Set Rs = Nothing
    Set rs1 = Nothing
    Set rs2 = Nothing
    
    Set objBBSsql = Nothing
    Set objPrgBar = Nothing
End Sub

Private Sub Form_Load()
    Set GetPtInfo = New frmPtInfo
    
    Me.MousePointer = 11
    chkSPos.value = 1
    dtpReqDt = GetSystemDate
    dtpReqTm = GetSystemDate
    
    
    Clear
    txtPtId = ""
    
    Call medClearTable(tblTestList)

    Dim objAccess As clsBBSAccess
    Dim Rs        As Recordset
    
    Set objAccess = New clsBBSAccess
    
    With objAccess
        Set Rs = New Recordset
        Rs.Open .Get_LegPos(ObjSysInfo.BuildingCd), DBConn
        If Rs.EOF = False Then
            cboLeg.Clear
            Do Until Rs.EOF = True
                cboLeg.AddItem Rs.Fields("legcd").value & ""
                Rs.MoveNext
            Loop
            cboLeg.ListIndex = 0
        End If
        
    End With
    Set Rs = Nothing
    Set objAccess = Nothing
    
    
    optOrderDiv(0).value = True
    fraWard.Enabled = False
    txtWardId(0).BackColor = Me.BackColor
    txtWardId(1).BackColor = Me.BackColor
    txtWardId(2).BackColor = Me.BackColor
    
    Me.Show
    DoEvents

    Call QueryTestCode
    Me.MousePointer = 0
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set GetPtInfo = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call ICSPatientMark
End Sub

Private Sub GetPtInfo_Click(ByVal isSELECT As Boolean, ByVal ptInfo As S2BBS_Library.clsPtInformation)
    If isSELECT = False Then Exit Sub
    Call Clear
    With ptInfo
        If .PtDiv = "BED" Then
            optOrderDiv(1).value = True

            txtPtId = .PtId
            lblPtNm.Caption = .ptnm
            lblSexAge.Caption = .Sex & " / " & .Age
            lblDob.Caption = .DOB
            lblOrderDoctor.Caption = .DoctNm
            lblDeptNm.Caption = .DeptNm

            txtDeptcd = .DeptCd
            txtOrderDoctor = .MajDoct
            txtWardId(0) = .wardid
            txtWardId(1) = .HosilID
            txtWardId(2) = .BedID

            fraWard.Enabled = True
        Else
            optOrderDiv(0).value = True
            txtPtId = .PtId
            lblPtNm.Caption = .ptnm
            lblSexAge.Caption = .Sex & " / " & .Age
            lblDob.Caption = .DOB
            fraWard.Enabled = False
        End If
        Call SetDestCenter

        SendKeys "{tab}"
    End With

End Sub

Private Sub DetailSearch()
    Dim objinfection    As clsInfection
    Dim objReaction     As clsReaction
    
    Set objinfection = New clsInfection
    Set objReaction = New clsReaction
    
    With objinfection
        .PtId = txtPtId
        .GetInfection
        If .Infection = True Then
            lblInfection.Visible = True
        Else
            lblInfection.Visible = False
        End If
    End With
    With objReaction
        .PtId = txtPtId
        .GetReaction
        If .Reaction = True Then
            lblReaction.Visible = True
        Else
            lblReaction.Visible = False
        End If
    End With
    Set objReaction = Nothing
    Set objinfection = Nothing
End Sub

Private Sub Clear()
    medClearTable tblTestList

    txtOrderDoctor = ""
    txtDeptcd = ""
    txtWardId(0) = ""
    txtWardId(1) = ""
    txtWardId(2) = ""
    txtRowNo = ""
    txtColNo = ""
    
    lblPtNm.Caption = ""
    lblSexAge.Caption = ""
    lblDob.Caption = ""
    lblDeptNm.Caption = ""
    lblOrderDoctor.Caption = ""
    lblInfection.Visible = False
    lblReaction.Visible = False
    chkStat.value = 0
    
    lblDestCenterCd.Caption = ""
    lblDestCenterNm.Caption = ""
End Sub

Private Sub cmdRight_Click()
    Dim Row             As Long
    
    Dim strTmptestnm    As String
    Dim strTmpTestcd    As String
    Dim strTmpTestDiv   As String
    Dim reasonNm        As String
    Dim ReasonCd        As String
    Dim opDiv           As String
    Dim opNm            As String
    Dim opCd            As String
    Dim i               As Integer
    Dim iTmx            As ListItem
    
    
    reasonNm = medGetP(lstReason.Text, 1, vbTab)
    ReasonCd = medGetP(lstReason.Text, 2, vbTab)
    opDiv = medGetP(lstReason.Text, 3, vbTab)
    opNm = medGetP(lstOcd.Text, 1, vbTab)
    opCd = medGetP(lstOcd.Text, 2, vbTab)
    
    
    With tblTestList
        For Each iTmx In lvwOrderList.ListItems
            If iTmx.Checked = True Then
                strTmptestnm = iTmx.Text
                strTmpTestcd = iTmx.SubItems(1)
                strTmpTestDiv = iTmx.SubItems(2)
                
                Row = .DataRowCnt + 1
                .Row = Row
                If .Row > .MaxRows Then .MaxRows = .MaxRows + 1
                
                
                .Col = TblColumn.tcTESTNM:          .value = strTmptestnm
                .Col = TblColumn.tcTESTCD:          .value = strTmpTestcd
                .Col = TblColumn.tcQTY:             .value = 1
                .Col = TblColumn.tcREQDTTM:         .value = Format(dtpReqDt, "YYYY-MM-DD") & " " & Format(dtpReqTm, "HH:MM")
                .Col = TblColumn.tcREASONNM:        .value = Trim(reasonNm)
                .Col = TblColumn.tcREASONCD:        .value = Trim(ReasonCd)
                
                .Col = TblColumn.tcISOP:            .value = opDiv
                If opDiv = "1" Then
                    .Col = TblColumn.tcREASONNM:    .value = Trim(reasonNm) & "-" & Trim(opNm)
                    .Col = TblColumn.tcOPCD:        .value = Trim(opCd)
                Else
                    .Col = TblColumn.tcOPCD:        .value = ""
                End If
                
                .Col = TblColumn.tcISPHER:          .value = strTmpTestDiv
                
                iTmx.Checked = False
            End If
        Next iTmx
    End With
End Sub

Private Sub lstTestCd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strTmp      As String
    Dim astrTmp()   As String
    
    strTmp = lstTestCd.Text
    If strTmp <> "" Then
        astrTmp = Split(strTmp, vbTab)
        
        With tblTestList
            .Row = .ActiveRow
            If astrTmp(3) = "0" Then
                .Col = TblColumn.tcTESTCD: .value = astrTmp(0)
                .Col = TblColumn.tcTESTNM: .value = astrTmp(1)
                .Col = TblColumn.tcISPHER: .value = astrTmp(2)
            Else
                .Col = TblColumn.tcTESTCD: .value = astrTmp(0)
                .Col = TblColumn.tcTESTNM: .value = astrTmp(1)
                .Col = TblColumn.tcISPHER: .value = astrTmp(2)
            End If
            .Col = TblColumn.tcQTY
            .Action = ActionActiveCell
        End With
        
    End If
    
    lstTestCd.Visible = False
End Sub

Private Sub lvwOrderList_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Item.Selected = True
End Sub

'Private Sub mnuDelete_Click()
'    With tblTestList
'        .Row = .ActiveRow
'        .Action = ActionDeleteRow
'    End With
'End Sub

Private Sub SetDestCenter()
    Dim objOrder        As clsBBSOrder
    Dim Orgcentercd     As String
    Dim Destcentercd    As String
    
    If optOrderDiv(0).value = True Then
        '�ܷ�--------------------------
        lblDestCenterCd.Caption = ObjSysInfo.BuildingCd
        lblDestCenterNm.Caption = ObjSysInfo.BuildingNm
    Else
        '����--------------------------
        If txtWardId(0) = "" Then
            lblDestCenterCd.Caption = ""
            lblDestCenterNm.Caption = ""
        Else
            Set objOrder = New clsBBSOrder
            Orgcentercd = objOrder.Get_Build(txtWardId(0))
            If Orgcentercd <> "" Then
                If chkStat.value = 1 Then
                    Destcentercd = medGetP(objOrder.Get_TestBuild(Orgcentercd), 2, COL_DIV)
                Else
                    Destcentercd = medGetP(objOrder.Get_TestBuild(Orgcentercd), 1, COL_DIV)
                End If
            Else
                Destcentercd = ""
            End If
            Set objOrder = Nothing
            
            lblDestCenterCd.Caption = Destcentercd
            lblDestCenterNm.Caption = GetCenterNm(Destcentercd)
        End If
    End If
End Sub

'Private Sub objListPop_SelectedItem(ByVal pSelectedItem As String)
'
'End Sub

Private Sub objPop_Click(ByVal vMenuID As Long)
    Select Case vMenuID
        Case MENU_DEL
            With tblTestList
                .Row = .ActiveRow
                .Action = ActionDeleteRow
            End With
    End Select
End Sub

'Private Sub objPtInfo_SelectedId(ByVal vPtID As String)
'    If vPtID = "" Then GoTo NoData
'
'    Call Clear
'    With objPtInfo
'        If .GETPatient(vPtID) = False Then GoTo NoData
'
'        If .INADMISSION Then
'            optOrderDiv(1).value = True
'
'            txtPtId = .PtId
'            lblPtNm.Caption = .ptnm
'            lblSexAge.Caption = .Sex & " / " & .Age
'            lblDob.Caption = .DOB
'            lblOrderDoctor.Caption = .DoctNm
'            lblDeptNm.Caption = .DeptNm
'
'            txtDeptcd = .DeptCd
'            txtOrderDoctor = .MajDoct
'            txtWardId(0) = .wardid
'            txtWardId(1) = .ROOMID
'            txtWardId(2) = .BedID
'
'            fraWard.Enabled = True
'        Else
'            optOrderDiv(0).value = True
'            txtPtId = .PtId
'            lblPtNm.Caption = .ptnm
'            lblSexAge.Caption = .Sex & " / " & .Age
'            lblDob.Caption = .DOB
'            fraWard.Enabled = False
'        End If
'
'        Call SetDestCenter
'
'        SendKeys "{tab}"
''        If .PtDiv = "BED" Then
''            optOrderDiv(1).value = True
''
''            txtPtid = .PtId
''            lblPtNm.Caption = .ptnm
''            lblSexAge.Caption = .Sex & " / " & .Age
''            lblDob.Caption = .DOB
''            lblOrderDoctor.Caption = .DoctNm
''            lblDeptNm.Caption = .DeptNm
''
''            txtDeptcd = .DeptCd
''            txtOrderDoctor = .MajDoct
''            txtWardID(0) = .wardid
''            txtWardID(1) = .HosilID
''            txtWardID(2) = .BedID
''
''            fraWard.Enabled = True
''        Else
''            optOrderDiv(0).value = True
''            txtPtid = .PtId
''            lblPtNm.Caption = .ptnm
''            lblSexAge.Caption = .Sex & " / " & .Age
''            lblDob.Caption = .DOB
''            fraWard.Enabled = False
''        End If
''        Call SetDestCenter
''
''        SendKeys "{tab}"
'    End With
'
'NoData:
'    Set objPtInfo = Nothing
'End Sub

Private Sub optOrderDiv_Click(Index As Integer)
    If Index = 1 Then
        '����--------------------------------------
        Call SetDestCenter
        
        fraWard.Enabled = True
        txtWardId(0).BackColor = RGB(255, 255, 255)
        txtWardId(1).BackColor = RGB(255, 255, 255)
        txtWardId(2).BackColor = RGB(255, 255, 255)
    Else
        '�ַ�--------------------------------------
        Call SetDestCenter
        
        fraWard.Enabled = False
        txtWardId(0).BackColor = Me.BackColor
        txtWardId(1).BackColor = Me.BackColor
        txtWardId(2).BackColor = Me.BackColor
    End If
End Sub

Private Function Search_Dept() As Boolean
    If txtDeptcd.Text = "" Then Search_Dept = True: Exit Function
    
    Search_Dept = True
    lblDeptNm.Caption = GetDeptNm(UCase(txtDeptcd.Text))
    If lblDeptNm.Caption = "" Then
        MsgBox "�ش�Ǵ� �ڷᰡ �����ϴ�.Ȯ���� �Է��ϼ���.", vbInformation + vbOKOnly, Me.Caption
        txtDeptcd = ""
        lblDeptNm.Caption = ""
        Search_Dept = False
    End If
    
End Function

Private Function Search_Doctor() As Boolean
    
    If txtOrderDoctor = "" Then Search_Doctor = True: Exit Function
    Search_Doctor = True
    lblOrderDoctor.Caption = GetDoctNm(txtOrderDoctor.Text)
    
    If lblOrderDoctor.Caption = "" Then
        MsgBox "�ش�Ǵ� �ڷᰡ �����ϴ�.Ȯ���� �Է��ϼ���.", vbInformation + vbOKOnly, Me.Caption
        txtOrderDoctor = ""
        lblOrderDoctor.Caption = ""
        Search_Doctor = False
    End If

End Function

Private Sub Search_Ward()
    
    txtWardId(1).Text = ""
    txtWardId(2).Text = ""
    
    If txtWardId(0) <> "" Then
        txtWardId(0) = UCase(txtWardId(0))
        If GetWardNm(txtWardId(0).Text) = "" Then
            MsgBox "�ش�Ǵ� �ڷᰡ �����ϴ�. Ȯ���� �Է��ϼ���.", vbInformation + vbOKOnly, "�����Է�"
            Exit Sub
        End If
        Call SetDestCenter
    End If
End Sub

Private Sub optOrderDiv_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub tblTestList_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    If Row < 1 Then Exit Sub
    
    With tblTestList
        .Col = Col
        .Row = Row
        .Action = ActionActiveCell
    End With
    Set objPop = New clsPopupMenu
    With objPop
        .AddMenu MENU_DEL, "DELETE"
        .PopupMenus Me.hwnd
    End With
    Set objPop = Nothing
'    Set mnuPopup = frmControls.mnuPopup
'    Set mnuDelete = frmControls.mnuSub
'    mnuDelete.Caption = "Delete"
'    PopupMenu mnuPopup
'
'    Set mnuPopup = Nothing
'    Set mnuDelete = Nothing
End Sub

Private Sub txtDeptCd_GotFocus()
    txtDeptcd.tag = txtDeptcd
End Sub

Private Sub txtDeptcd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Search_Dept = True Then
            txtDeptcd.tag = txtDeptcd
            SendKeys "{TAB}"
        End If
    End If
End Sub

Private Sub txtDeptcd_LostFocus()
    If Screen.ActiveForm.ActiveControl.name = "cmdExit" Then Exit Sub
    If Screen.ActiveForm.ActiveControl.name = "cmdClear" Then Exit Sub
    
    If txtDeptcd.tag = txtDeptcd Then Exit Sub
    
    If Search_Dept = False Then txtDeptcd.SetFocus
End Sub

'Private Sub txtLegcd_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
'End Sub

Private Sub txtOrderDoctor_GotFocus()
    txtOrderDoctor.tag = txtOrderDoctor
End Sub

Private Sub txtOrderDoctor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Search_Doctor = True Then
            txtOrderDoctor.tag = txtOrderDoctor
            SendKeys "{TAB}"
        End If
    End If
End Sub

Private Sub txtOrderDoctor_LostFocus()
'Screen.ActiveForm.ActiveControl.name
    If ActiveControl.name = "cmdExit" Then Exit Sub
    If ActiveControl.name = "cmdClear" Then Exit Sub
    
    If txtOrderDoctor.tag = txtOrderDoctor Then Exit Sub
    
    If Search_Doctor = False Then txtOrderDoctor.SetFocus
End Sub

Private Sub txtPtId_GotFocus()
    txtPtId.tag = txtPtId
    
    txtPtId.SelStart = 0
    txtPtId.SelLength = Len(txtPtId)
    Exit Sub
End Sub

Private Sub txtPtid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Search_PtInfo = True Then
            SendKeys "{TAB}"
            txtPtId.tag = txtPtId
        End If
    End If
End Sub


Private Sub txtPtId_LostFocus()
    If Screen.ActiveForm.name <> Me.name Then Exit Sub
    
    If Screen.ActiveForm.ActiveControl.name = "cmdExit" Then Exit Sub
    If Screen.ActiveForm.ActiveControl.name = "cmdClear" Then Exit Sub
    
    If txtPtId.tag = txtPtId Then Exit Sub
    
    If Search_PtInfo = False Then txtPtId.SetFocus
End Sub


Private Function Search_PtInfo() As Boolean
    Dim objPtInfo As clsPtInformation
    Dim DrRS      As Recordset
    Dim ii        As Long
    Dim strLng    As String
    
    Call Clear
    If txtPtId = "" Then Search_PtInfo = True: Exit Function
    
    For ii = 1 To Val(BBS_PTID_LENGTH) - 1
        strLng = strLng & "0"
    Next ii

    If Len(Trim(txtPtId.Text)) <> BBS_PTID_LENGTH Then
        txtPtId.Text = Format(txtPtId.Text, strLng & "#")
    End If
    
    Set objPtInfo = New clsPtInformation
    Set DrRS = New Recordset
    
    DrRS.Open objPtInfo.Get_Ptid(txtPtId.Text), DBConn
    If DrRS.EOF = True Then
        MsgBox "�ش�Ǵ� ȯ�ڰ� �����ϴ�. Ȯ���� ��ȸ�ϼ���.", vbInformation + vbOKOnly, Me.Caption
        Search_PtInfo = False
    Else
        With objPtInfo
            .BedPt_Chk txtPtId.Text, Format(GetSystemDate, PRESENTDATE_FORMAT)
            If .PtDiv = "BED" Then
                optOrderDiv(1).value = True
                txtPtId = .PtId
                lblPtNm.Caption = .ptnm
                lblSexAge.Caption = .Sex & " / " & .Age
                lblDob.Caption = .DOB
                txtOrderDoctor = .MajDoct
                lblOrderDoctor.Caption = .DoctNm
                lblDeptNm.Caption = .DeptNm
                txtDeptcd = .DeptCd
                txtWardId(0) = .wardid
                txtWardId(1) = .HosilID
                txtWardId(2) = .BedID
            Else
                optOrderDiv(0).value = True
                txtPtId = .PtId
                lblPtNm.Caption = .ptnm
                lblSexAge.Caption = .Sex & " / " & .Age
                lblDob.Caption = .DOB
            End If
            Call SetDestCenter
        End With
        
        '�������� Ȯ��
        Call ICSPatientMark(txtPtId.Text, enICSNum.BBS_ALL)
        
        
        Dim objinfection As New clsInfection
        Dim objReaction As New clsReaction
        
        With objinfection
            .PtId = txtPtId
            .GetInfection
            lblInfection.Caption = IIf(.Infection = True, "*", "")
        End With
        With objReaction
            .PtId = txtPtId
            .GetReaction
            lblReaction.Caption = IIf(.Reaction = True, "��", "��")
            
        End With
        Set objReaction = Nothing
        Set objinfection = Nothing
        
        Search_PtInfo = True
    End If
    Set DrRS = Nothing
    Set objPtInfo = Nothing
    
End Function

Private Sub txtReceiptNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Len(txtReceiptNo) > 12 Then
            txtReceiptNo = ""
            txtReceiptNo.SetFocus
        End If
    End If
End Sub

Private Sub txtReceiptNo_KeyPress(KeyAscii As Integer)
    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And (KeyAscii <> vbKeyBack) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub txtWardID_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If Index = 0 Then Call cmdPopUp_Click(3)
    End If
    If KeyCode = 13 Then SendKeys "{tab}"
End Sub
Private Sub txtWardId_LostFocus(Index As Integer)
    If Index <> 0 Then Exit Sub
    If txtWardId(0) = "" Then
        txtWardId(1) = ""
        txtWardId(2) = ""
    Else
        Call Search_Ward
    End If
End Sub
Private Sub txtRowNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    Call OnlyNum(txtRowNo, KeyAscii)
End Sub

Private Sub txtColNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    Call OnlyNum(txtColNo, KeyAscii)
End Sub
Private Sub OnlyNum(TxtBox As TextBox, KeyAscii As Integer)
    '�ؽ�Ʈ�ڽ��� ���ڸ� �Է��Ҽ� ����....
    Dim i As Byte
   
    Select Case KeyAscii
      Case 8, 9, 13, 48 To 57 '8: BackSpace,45: -,46: . ,47: / , 48: 0~57:9
        '���
      Case 45  '-(���� ��ȣ�� ó����)
        If TxtBox.SelStart <> 0 Then KeyAscii = 0
      Case 46  '.(�Ҽ����� �ѹ���)
        If TxtBox.SelLength = 0 Then
          For i = 1 To Len(TxtBox)
            If Mid(TxtBox, i, 1) = "." Then KeyAscii = 0
          Next i
        Else
          For i = 1 To TxtBox.SelStart
            If Mid(TxtBox, i, 1) = "." Then KeyAscii = 0
          Next i
          For i = TxtBox.SelStart + TxtBox.SelLength + 1 To Len(TxtBox)
            If Mid(TxtBox, i, 1) = "." Then KeyAscii = 0
          Next i
        End If
      Case Else
        KeyAscii = 0
    End Select
 End Sub


'----------------
'���� �Լ�
'----------------
Private Function OnlyOrderSaver(ByVal TF As Boolean) As Boolean
    '----------------------------------------------------------
    'TF = False : ó���ϸ� ó��
    '     True  : ó��, ä��, ���� ó��
    '----------------------------------------------------------
    Dim objOrderSave As clsBBSOrder
    Dim objOrder    As clsDictionary       'ó�����
    Dim objBody     As clsDictionary       'ó��ٵ�
    
    '����� ����
    Dim PtId      As String
    Dim Bussdiv   As String
    Dim bedindt   As String
    Dim reqdt     As String
    Dim reqtm     As String
    Dim DeptCd    As String
    Dim OrdDoct   As String
    Dim EntID     As String
    Dim wardid    As String
    Dim BedID     As String
    Dim HosilID   As String
    Dim OCd       As String
    Dim ReasonCd  As String
    
    '�ٵ� ����
    Dim ordcd     As String
    Dim unitqty   As String
    Dim IrradFg   As String
    Dim FilterFg  As String
    Dim Phere     As String
    Dim TestDiv   As String
    Dim orddiv    As String
    
    '�ٵ������ �������� ����
    Dim StatFg    As String
    Dim ReceiptNo As String
    
    'ó������ä���� ����
    Dim SavePos   As String
    Dim Leg       As String
    Dim Row       As String
    Dim Col       As String
    Dim colid     As String
    Dim coldt     As String
    Dim ColTm     As String
    
    Dim ii        As Integer
    Dim kk        As Integer
    
    Dim strBkReqDtTm    As String
    Dim strBkReason     As String
    Dim strBkOpCd       As String
    Dim strBkOrdDiv     As String
    
    
    '�Է°� Check---------------------------------------------------------------------------------
    If Valadiation_Check = False Then Exit Function
    Call OrderSort
    
    '�޸� �Ҵ�----------------------------------------------------------------------------------
    Set objOrderSave = New clsBBSOrder
    Set objOrder = New clsDictionary
    Set objBody = New clsDictionary
    
    
    
    'ó�� �ش��� Dictionary �ʱ�ȭ----------------------------------------------------------------
    objOrder.FieldInialize "ptid", _
                           "bussdiv,bedindt,reqdt,reqtm,deptcd,orddoct,entid,receiptno," & _
                           "wardid,bedid,hosilid,ocd,reasoncd,statfg,savepos,leg,row,col," & _
                           "colid,coldt,coltm,buildcd,orddiv"

    'ó�� �ٵ�� Dictionary �ʱ�ȭ----------------------------------------------------------------
    objBody.FieldInialize "seq", "ordcd,unitqty,irradfg,filterfg,phere,testdiv,orddiv"



    '�ܷ�,������ ���� �Է°� ó��-----------------------------------------------------------------
    If optOrderDiv(0).value = True Then
        Bussdiv = BBSBUSSDIV.stsNotBed
        bedindt = ""
        wardid = ""
        BedID = ""
        HosilID = ""
    Else
        Bussdiv = BBSBUSSDIV.stsBed
        bedindt = ""
        wardid = txtWardId(0)
        BedID = txtWardId(1)
        HosilID = txtWardId(2)
    End If
    
    '��ü ������� ���� ����� ���� �Է°� ó��---------------------------------------------------
    If chkSPos.value = 1 Then
        SavePos = "1"
        Leg = cboLeg.Text
        Row = ""
        Col = ""
    Else
        SavePos = "0"
        Leg = cboLeg.Text
        Row = txtRowNo
        Col = txtColNo
    End If
    
    
    
    '�Է°� ó��----------------------------------------------------------------------------------
    StatFg = chkStat.value
    coldt = Format(GetSystemDate, PRESENTDATE_FORMAT)
    ColTm = Format(GetSystemDate, PRESENTTIME_FORMAT)
    DeptCd = txtDeptcd
    OrdDoct = txtOrderDoctor
    colid = ObjSysInfo.EmpId
    EntID = ObjSysInfo.EmpId
    
    
    '��������ȣ,�������� �Է°� ó��--------------------------------------------------------------
    ReceiptNo = txtReceiptNo
    
    
    'ó�� Spread�� �����鼭, �ٵ� Dictionary�� �� Setting-----------------------------------------
    strBkReqDtTm = ""
    strBkReason = ""
    strBkOpCd = ""
    strBkOrdDiv = ""
    
    DBConn.BeginTrans
    
    With tblTestList
        For ii = 1 To .DataRowCnt
            .Row = ii
            
            .Col = TblColumn.tcTESTCD:      ordcd = .value
            .Col = TblColumn.tcQTY:         unitqty = IIf(.value = "", "", .value)
            .Col = TblColumn.tcIRR:         IrradFg = IIf(.value = 1, "1", "0")
            .Col = TblColumn.tcFilter:      FilterFg = IIf(.value = 1, "1", "0")
            .Col = TblColumn.tcISPHER:      Phere = .value
            .Col = TblColumn.tcREQDTTM:     reqdt = Format(.value, PRESENTDATE_FORMAT)
                                            reqtm = Format(.value, PRESENTTIME_FORMAT)
            .Col = TblColumn.tcREASONCD:    ReasonCd = .value
            .Col = TblColumn.tcOPCD:        OCd = .value
            
            TestDiv = Get_Testdiv(ordcd)
            orddiv = Get_OrdDiv(ordcd)
            
            If ii = 1 Then
                kk = kk + 1
                objOrder.AddNew txtPtId.Text, _
                                Join(Array(Bussdiv, bedindt, reqdt, reqtm, DeptCd, OrdDoct, EntID, ReceiptNo, _
                                           wardid, BedID, HosilID, OCd, ReasonCd, StatFg, _
                                           SavePos, Leg, Row, Col, colid, coldt, ColTm, ObjSysInfo.BuildingCd, orddiv), _
                                     COL_DIV)
                objBody.AddNew kk, Join(Array(ordcd, unitqty, IrradFg, FilterFg, Phere, TestDiv), COL_DIV)
                
                strBkReqDtTm = reqdt & reqtm
                strBkReason = ReasonCd
                strBkOpCd = OCd
            Else
                If strBkReqDtTm <> (reqdt & reqtm) Or _
                   strBkReason <> ReasonCd Or _
                   strBkOpCd <> OCd Or _
                   strBkOrdDiv <> orddiv Then
                    '
                    ' ���� ó��
                    '
                    If Not objOrderSave.Set_Order(objOrder, objBody) Then GoTo OnlyOrderSaver_error
                    
                    Set objOrder = Nothing
                    Set objBody = Nothing
                    
                    Set objOrder = New clsDictionary
                    Set objBody = New clsDictionary
                    
                    kk = 1
                    'ó�� �ش��� Dictionary �ʱ�ȭ----------------------------------------------------------------
                    objOrder.FieldInialize "ptid", _
                                           "bussdiv,bedindt,reqdt,reqtm,deptcd,orddoct,entid,receiptno," & _
                                           "wardid,bedid,hosilid,ocd,reasoncd,statfg,savepos,leg,row,col," & _
                                           "colid,coldt,coltm,buildcd,orddiv"
                
                    'ó�� �ٵ�� Dictionary �ʱ�ȭ----------------------------------------------------------------
                    objBody.FieldInialize "seq", "ordcd,unitqty,irradfg,filterfg,phere,testdiv"
                    
                    
                    objOrder.AddNew txtPtId.Text, _
                                    Join(Array(Bussdiv, bedindt, reqdt, reqtm, DeptCd, OrdDoct, EntID, ReceiptNo, _
                                               wardid, BedID, HosilID, OCd, ReasonCd, StatFg, _
                                               SavePos, Leg, Row, Col, colid, coldt, ColTm, ObjSysInfo.BuildingCd, orddiv), _
                                         COL_DIV)
                    objBody.AddNew kk, Join(Array(ordcd, unitqty, IrradFg, FilterFg, Phere, TestDiv), COL_DIV)
                    
                    strBkReqDtTm = reqdt & reqtm
                    strBkReason = ReasonCd
                    strBkOpCd = OCd
                    strBkOrdDiv = orddiv
                Else
                    kk = kk + 1
                    objBody.AddNew kk, Join(Array(ordcd, unitqty, IrradFg, FilterFg, Phere, TestDiv), COL_DIV)
                End If
            End If
        Next
    End With
    
    '
    ' ���� ó��
    '
    If Not objOrderSave.Set_Order(objOrder, objBody, TF) Then GoTo OnlyOrderSaver_error
                        
    DBConn.CommitTrans
    OnlyOrderSaver = True

    '���ڵ� ����� ���Ͼ�
    If TF = True Then BarSpcNum = objOrderSave.SpcNum
    
    '�޸� ����----------------------------------------------------------------------------------
    Set objOrderSave = Nothing
    Set objOrder = Nothing
    Set objBody = Nothing

    Exit Function
    
OnlyOrderSaver_error:
    DBConn.RollbackTrans
    OnlyOrderSaver = False
    Set objOrderSave = Nothing
    Set objOrder = Nothing
    Set objBody = Nothing
    MsgBox Err.Description, vbExclamation
End Function

Private Function Get_OrdDiv(ByVal TestCd As String) As String
    Dim objOrder As New clsBBSOrder
    
    Get_OrdDiv = objOrder.Get_OrdDiv(TestCd)
    
    If Get_OrdDiv = "" Then Get_OrdDiv = "B"
    
    Set objOrder = Nothing
End Function

Private Function Get_Testdiv(ByVal TestCd As String) As String
    Dim objOrder As New clsBBSOrder
    
    
    Get_Testdiv = objOrder.Get_Testdiv(TestCd)
    
    If Get_Testdiv = "" Then Get_Testdiv = "N"
    
    Set objOrder = Nothing
End Function


'#�����ڵ� ����Ʈ �ڽ�
Private Sub lstOcd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim tmpStr As String
    Dim tmpField1 As String
    Dim tmpField2 As String
    
    If Button <> 1 Then Exit Sub

    tmpStr = lstOcd.List(lstOcd.ListIndex)
    With tblTestList
        .SetFocus
        
        tmpField1 = Trim(medShift(tmpStr, Chr(9)))
        tmpField2 = medShift(tmpStr, Chr(9))
        Call medShift(tmpStr, Chr(9))
        
        .Row = .ActiveRow

        .Col = TblColumn.tcOPNM: .value = Trim(tmpField1)       ' ������
        .Col = TblColumn.tcOPCD: .value = Trim(tmpField2)       ' �����ڵ�
        
        .Row = .ActiveRow + 1
        .Col = TblColumn.tcTESTNM
        .Action = ActionActiveCell
        
        lstOcd.Visible = False
    End With

End Sub

''#�������� ����Ʈ�ڽ�
Private Sub lstReason_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim tmpStr As String
    Dim tmpField1 As String
    Dim tmpField2 As String
    Dim tmpField3 As String
    Dim Wdt As Long, Hgt As Long
    Dim X1 As Long
    Dim Y1 As Long
    Dim Ret As Boolean
    Dim Cancel As Boolean

    If Button <> 1 Then Exit Sub

    tmpStr = lstReason.List(lstReason.ListIndex)
    With tblTestList
        .SetFocus
        
        tmpField1 = Trim(medShift(tmpStr, Chr(9)))
        tmpField2 = Trim(medShift(tmpStr, Chr(9)))
        tmpField3 = Trim(medShift(tmpStr, Chr(9)))

        .Row = .ActiveRow
        .Col = TblColumn.tcREASONNM:  .value = Trim(tmpField1)  ' ��������
        .Col = TblColumn.tcREASONCD:  .value = Trim(tmpField2)  ' ���������ڵ�
        .Col = TblColumn.tcISOP:      .value = Trim(tmpField3)  ' ��������
        
        If tmpField3 = "1" Then
            .Row = .ActiveRow
            .Col = TblColumn.tcOPNM
            .Action = ActionActiveCell
        Else
            .Row = .ActiveRow + 1
            .Col = TblColumn.tcTESTNM
            .Action = ActionActiveCell
        End If
        
        lstReason.Visible = False

    End With
End Sub



Private Sub tblTestList_EditChange(ByVal Col As Long, ByVal Row As Long)
    Dim tmpIndex    As Integer
    Dim tmpStr      As String
    Dim Wdt         As Long
    Dim Hgt         As Long
    Dim X           As Long
    Dim Y           As Long
    Dim Ret         As Boolean
    Dim strInput    As String
    
    With tblTestList
        .Col = Col
        .Row = Row
        strInput = .value

         Select Case Col
            '�˻��׸�
            Case TblColumn.tcTESTNM
                lstTestCd.Visible = True
                tmpIndex = medListFind(lstTestCd, strInput)
                tmpStr = lstTestCd.List(tmpIndex)
                ' ���ڰ� �Էµɶ����� ����� ã��

                If tmpIndex = -1 Or UCase(tmpStr) <> UCase(Trim(.value)) Then
                   Ret = .GetCellPos(Col, Row, X, Y, Wdt, Hgt)
                   Y = Y + Hgt
                   If .Height - Y < lstTestCd.Height Or Y < 0 Then
                      Ret = .GetCellPos(Col, Row, X, Y, Wdt, Hgt)
                      lstTestCd.Top = .Top + Y - lstTestCd.Height + medMain.picMain.Height + 950

                      lstTestCd.Left = .Left + X
                   Else
                      lstTestCd.Left = .Left + X
                      lstTestCd.Top = .Top + Y
                   End If
                   If tmpIndex >= 0 Then
                      lstTestCd.ListIndex = tmpIndex
                      If tmpIndex - lstTestCd.TopIndex > 10 Then lstTestCd.TopIndex = tmpIndex
                   End If
                   lstTestCd.Visible = True
                   lstTestCd.ZOrder 0
                Else
                   lstTestCd.ListIndex = tmpIndex
                   Call lstTestCd_MouseDown(1, 0, 0, 0)
                End If
            '��������
            Case TblColumn.tcREASONNM
                lstReason.Visible = True
                tmpIndex = medListFind(lstReason, strInput)
                tmpStr = lstReason.List(tmpIndex)
                ' ���ڰ� �Էµɶ����� ����� ã��

                If tmpIndex = -1 Or UCase(tmpStr) <> UCase(Trim(.value)) Then
                   Ret = .GetCellPos(Col, Row, X, Y, Wdt, Hgt)
                   Y = Y + Hgt
                   If .Height - Y < lstOcd.Height Or Y < 0 Then
                      Ret = .GetCellPos(Col, Row, X, Y, Wdt, Hgt)
                      lstReason.Top = .Top + Y - lstReason.Height + medMain.picMain.Height + 950

                      lstReason.Left = .Left + X
                   Else
                      lstReason.Left = .Left + X
                      lstReason.Top = .Top + Y
                   End If
                   If tmpIndex >= 0 Then
                      lstReason.ListIndex = tmpIndex
                      If tmpIndex - lstReason.TopIndex > 10 Then lstReason.TopIndex = tmpIndex
                   End If
                   lstReason.Visible = True
                   lstReason.ZOrder 0
                Else
                   lstReason.ListIndex = tmpIndex
                   Call lstReason_MouseDown(1, 0, 0, 0)
                End If
            '�����ڵ�
            Case TblColumn.tcOPNM
                lstOcd.Visible = True
                tmpIndex = medListFind(lstOcd, strInput)
                tmpStr = lstOcd.List(tmpIndex)
                ' ���ڰ� �Էµɶ����� ����� ã��

                If tmpIndex = -1 Or UCase(tmpStr) <> UCase(Trim(.value)) Then
                   Ret = .GetCellPos(Col, Row, X, Y, Wdt, Hgt)
                   Y = Y + Hgt
                   If .Height - Y < lstOcd.Height Or Y < 0 Then
                      Ret = .GetCellPos(Col, Row, X, Y, Wdt, Hgt)
                      lstOcd.Top = .Top + Y - lstOcd.Height + medMain.picMain.Height + 950

                      lstOcd.Left = .Left + X - (lstOcd.Width - Wdt)
                   Else
                      lstOcd.Left = .Left + X - (lstOcd.Width - Wdt)
                      lstOcd.Top = .Top + Y
                   End If
                   If tmpIndex >= 0 Then
                      lstOcd.ListIndex = tmpIndex
                      If tmpIndex - lstOcd.TopIndex > 10 Then lstOcd.TopIndex = tmpIndex
                   End If
                   lstOcd.Visible = True
                   lstOcd.ZOrder 0
                Else
                   lstOcd.ListIndex = tmpIndex
                   Call lstOcd_MouseDown(1, 0, 0, 0)
                End If
                
         End Select

    End With

End Sub



Private Sub tblTestList_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim tmpTestCd As Variant
    Dim tmpSpcCd As Variant
    Dim tmpDate As Variant
    Dim Wdt As Long, Hgt As Long
    Dim X As Long, Y As Long
    Dim Ret As Boolean
    Dim strTest As String
    Dim strSpc As String
    Dim tmpValue As String
    Dim i As Long
    
    If NewCol < 0 Or NewRow < 0 Then
        lstTestCd.Visible = False
        Exit Sub
    End If

    Select Case NewCol
        Case TblColumn.tcREQDTTM
            lstTestCd.Visible = False
            lstReason.Visible = False
            lstOcd.Visible = False
            With tblTestList
                Ret = .GetCellPos(NewCol, NewRow, X, Y, Wdt, Hgt)
                If Y > 0 Then
                    picReqDtTm.Top = .Top + Y + Hgt
                    picReqDtTm.Left = .Left + X
                Else
                    picReqDtTm.Top = .Top + Y + Hgt - picReqDtTm.Height
                    picReqDtTm.Left = .Left + X
                End If
                
                .Row = NewRow
                .Col = NewCol
                tmpValue = .value
                If tmpValue <> "" Then
                    dtpReqDt = Mid(tmpValue, 1, 10)
                    dtpReqTm = Mid(tmpValue, 12, 5)
                Else
                End If
                picReqDtTm.Visible = True
                picReqDtTm.ZOrder 0
                dtpReqDt.SetFocus
            End With
        Case Else
            lstTestCd.Visible = False
            picReqDtTm.Visible = False
            lstReason.Visible = False
            lstOcd.Visible = False
    End Select
End Sub


Private Sub tblTestList_GotFocus()
    Dim Cancel As Boolean

    With tblTestList
        Call tblTestList_LeaveCell(.ActiveCol, .ActiveRow, .ActiveCol, .ActiveRow, Cancel)
    End With

End Sub
'
'% ���̺� ����Ʈ�� �� �ְ� �Ʒ�ȭ��ǥŰ�� ������ ��� ��Ŀ�� �̵�
Private Sub tblTestList_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim Cancel As Boolean
    Dim tmpValue As String
    
    '�����Է��� enterŰ
    With tblTestList
        If KeyCode = vbKeyReturn Then
            If .ActiveCol = TblColumn.tcQTY Then
                .Row = .ActiveRow
                .Col = .ActiveCol
                tmpValue = .value
                If Val(tmpValue) > 0 Then
                    .Row = .ActiveRow
                    .Col = TblColumn.tcREQDTTM
                    .Action = ActionActiveCell
                    Call tblTestList_LeaveCell(TblColumn.tcQTY, .ActiveRow, TblColumn.tcREQDTTM, .ActiveRow, Cancel)
                End If
                Exit Sub
            End If
        End If
    End With

    '�˻��׸�
    With lstTestCd
        If .Visible Then
            Select Case KeyCode
                Case vbKeyDown, vbKeyPageDown:
                    If .ListCount - 1 > .ListIndex Then .ListIndex = .ListIndex + 1
                    KeyCode = 0
                Case vbKeyUp, vbKeyPageUp:
                    If .ListIndex > 0 Then .ListIndex = .ListIndex - 1
                    KeyCode = 0
                Case vbKeyEscape:
                    .Visible = False
                Case vbKeyLeft, vbKeyRight:
                    .Visible = False
                Case vbKeyReturn:
                    Call lstTestCd_MouseDown(1, 0, 0, 0)
                    KeyCode = 0
            End Select
        End If
    End With

    '��������
    With lstReason
        If .Visible Then
            Select Case KeyCode
                Case vbKeyDown, vbKeyPageDown:
                    If .ListCount - 1 > .ListIndex Then .ListIndex = .ListIndex + 1
                    KeyCode = 0
                Case vbKeyUp, vbKeyPageUp:
                    If .ListIndex > 0 Then .ListIndex = .ListIndex - 1
                    KeyCode = 0
                Case vbKeyEscape:
                    .Visible = False
                Case vbKeyLeft, vbKeyRight:
                    .Visible = False
                Case vbKeyReturn:
                    Call lstReason_MouseDown(1, 0, 0, 0)
                    KeyCode = 0
            End Select
        End If
    End With

    '�����ڵ�
    With lstOcd
        If .Visible Then
            Select Case KeyCode
                Case vbKeyDown, vbKeyPageDown:
                    If .ListCount - 1 > .ListIndex Then .ListIndex = .ListIndex + 1
                    KeyCode = 0
                Case vbKeyUp, vbKeyPageUp:
                    If .ListIndex > 0 Then .ListIndex = .ListIndex - 1
                    KeyCode = 0
                Case vbKeyEscape:
                    .Visible = False
                Case vbKeyLeft, vbKeyRight:
                    .Visible = False
                Case vbKeyReturn:
                    Call lstOcd_MouseDown(1, 0, 0, 0)
                    KeyCode = 0
            End Select
        End If
    End With
End Sub

