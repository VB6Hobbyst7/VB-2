VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEQ_�˻������� 
   Caption         =   "�˻�������"
   ClientHeight    =   9375
   ClientLeft      =   3885
   ClientTop       =   5220
   ClientWidth     =   13200
   BeginProperty Font 
      Name            =   "����ü"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEQ_�˻�������.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9375
   ScaleWidth      =   13200
   Begin VB.Frame fraChange 
      Caption         =   "��ü��ȣ ����"
      Height          =   2355
      Left            =   13140
      TabIndex        =   58
      Top             =   600
      Visible         =   0   'False
      Width           =   2355
      Begin VB.CommandButton cmdDel 
         Caption         =   "����"
         Height          =   315
         Left            =   1260
         TabIndex        =   67
         Top             =   180
         Width           =   915
      End
      Begin VB.CommandButton cmdChangeCancel 
         Caption         =   "���"
         Height          =   375
         Left            =   1260
         TabIndex        =   64
         Top             =   1860
         Width           =   915
      End
      Begin VB.CommandButton cmdChange 
         Caption         =   "����"
         Height          =   375
         Left            =   240
         TabIndex        =   63
         Top             =   1860
         Width           =   915
      End
      Begin VB.TextBox txtAfterBarcode 
         Height          =   375
         Left            =   240
         TabIndex        =   62
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox txtBeforBarcode 
         Height          =   375
         Left            =   240
         TabIndex        =   61
         Top             =   540
         Width           =   1935
      End
      Begin VB.Label lblRowNo 
         Caption         =   "Row��ȣ"
         Height          =   195
         Left            =   1320
         TabIndex        =   66
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "After"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   60
         Top             =   1080
         Width           =   675
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Before"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   59
         Top             =   300
         Width           =   795
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  '����
      Height          =   375
      Left            =   600
      TabIndex        =   56
      Top             =   3180
      Width           =   435
      Begin VB.Label lblcheck 
         Caption         =   "���ۿ���     0"
         Height          =   375
         Left            =   10
         TabIndex        =   57
         Top             =   0
         Width           =   435
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "���(&C)"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6660
      Style           =   1  '�׷���
      TabIndex        =   11
      Top             =   2760
      Width           =   915
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFC0C0&
      Caption         =   "����(&S)"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11220
      Style           =   1  '�׷���
      TabIndex        =   13
      Top             =   60
      Width           =   915
   End
   Begin VB.ComboBox cboSENDFLAG 
      Height          =   300
      Left            =   4020
      Style           =   2  '��Ӵٿ� ���
      TabIndex        =   9
      Top             =   2760
      Width           =   1575
   End
   Begin VB.ComboBox cboSTATEFLAG 
      Height          =   300
      Left            =   4020
      Style           =   2  '��Ӵٿ� ���
      TabIndex        =   8
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox txtPATNM 
      Height          =   315
      Left            =   960
      TabIndex        =   7
      Text            =   "1234567890"
      Top             =   2760
      Width           =   1035
   End
   Begin VB.TextBox txtPATNO 
      Height          =   315
      Left            =   960
      TabIndex        =   6
      Text            =   "1234567890"
      Top             =   2400
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Caption         =   "[��������]"
      Height          =   915
      Left            =   60
      TabIndex        =   23
      Top             =   1020
      Width           =   7515
      Begin VB.OptionButton optDateSection 
         Caption         =   "ó������"
         Height          =   180
         Index           =   3
         Left            =   6360
         TabIndex        =   34
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optDateSection 
         Caption         =   "�˻�����������"
         Height          =   180
         Index           =   2
         Left            =   4500
         TabIndex        =   2
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton optDateSection 
         Caption         =   "�˻�����������"
         Height          =   180
         Index           =   1
         Left            =   2640
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton optDateSection 
         Caption         =   "�˻�ó����������"
         Height          =   180
         Index           =   0
         Left            =   780
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker dtpDateFrom 
         Height          =   315
         Left            =   780
         TabIndex        =   3
         Top             =   540
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   21364737
         CurrentDate     =   40820
      End
      Begin MSComCtl2.DTPicker dtpDateTo 
         Height          =   315
         Left            =   2340
         TabIndex        =   4
         Top             =   540
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   21364737
         CurrentDate     =   40820
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '����
         Caption         =   "�Ⱓ"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   26
         Top             =   600
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "~"
         Height          =   180
         Index           =   8
         Left            =   2160
         TabIndex        =   25
         Top             =   600
         Width           =   90
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '����
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.TextBox txtBARCD 
      Height          =   315
      Left            =   960
      TabIndex        =   5
      Text            =   "201101011234567"
      Top             =   2040
      Width           =   1575
   End
   Begin FPSpread.vaSpread sprDResult 
      Height          =   6195
      Left            =   7620
      TabIndex        =   12
      Top             =   3120
      Width           =   5475
      _Version        =   393216
      _ExtentX        =   9657
      _ExtentY        =   10927
      _StockProps     =   64
      BackColorStyle  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   14
      MaxRows         =   501
      SpreadDesigner  =   "frmEQ_�˻�������.frx":263A
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "�ݱ�(&Q)"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12180
      Style           =   1  '�׷���
      TabIndex        =   14
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "��ȸ(&V)"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5700
      Style           =   1  '�׷���
      TabIndex        =   10
      Top             =   2760
      Width           =   915
   End
   Begin MSComctlLib.ProgressBar barStatus 
      Height          =   75
      Left            =   60
      TabIndex        =   15
      Top             =   600
      Width           =   13035
      _ExtentX        =   22992
      _ExtentY        =   132
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin FPSpread.vaSpread sprLResult 
      Height          =   6195
      Left            =   60
      TabIndex        =   33
      Top             =   3120
      Width           =   7515
      _Version        =   393216
      _ExtentX        =   13256
      _ExtentY        =   10927
      _StockProps     =   64
      BackColorStyle  =   1
      ButtonDrawMode  =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   16
      MaxRows         =   20
      SpreadDesigner  =   "frmEQ_�˻�������.frx":432D
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   2055
      Left            =   2400
      TabIndex        =   65
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '����
      Caption         =   "Sample No"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   23
      Left            =   7680
      TabIndex        =   55
      Top             =   1620
      Width           =   915
   End
   Begin VB.Label lblSAMPLENO 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "1234567890"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   8640
      TabIndex        =   54
      Top             =   1620
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '����
      Caption         =   "(Like)"
      Height          =   180
      Index           =   22
      Left            =   2040
      TabIndex        =   53
      Top             =   2820
      Width           =   540
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '����
      Caption         =   "���� ��ȣ"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   21
      Left            =   11040
      TabIndex        =   52
      Top             =   2100
      Width           =   915
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '����
      Caption         =   "������ ��"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   18
      Left            =   11040
      TabIndex        =   51
      Top             =   2340
      Width           =   915
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '����
      Caption         =   "����/����"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   17
      Left            =   11040
      TabIndex        =   50
      Top             =   2580
      Width           =   915
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '����
      Caption         =   "ó�� ����"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   15
      Left            =   7680
      TabIndex        =   49
      Top             =   2100
      Width           =   915
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '����
      Caption         =   "ó�� ����"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   14
      Left            =   11040
      TabIndex        =   48
      Top             =   1620
      Width           =   915
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '����
      Caption         =   "ó�� ����"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   13
      Left            =   11040
      TabIndex        =   47
      Top             =   1860
      Width           =   915
   End
   Begin VB.Label lblEXDT 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "1234567890"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   8640
      TabIndex        =   46
      Top             =   2100
      Width           =   900
   End
   Begin VB.Label lblPATNO 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "1234567890"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   12000
      TabIndex        =   45
      Top             =   2100
      Width           =   900
   End
   Begin VB.Label lblPATNM 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "1234567890"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   12000
      TabIndex        =   44
      Top             =   2340
      Width           =   900
   End
   Begin VB.Label lblSEXAGE 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "1234567890"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   12000
      TabIndex        =   43
      Top             =   2580
      Width           =   900
   End
   Begin VB.Label lblORDDT 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "1234567890"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   12000
      TabIndex        =   42
      Top             =   1620
      Width           =   900
   End
   Begin VB.Label lblORDGB 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "1234567890"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   12000
      TabIndex        =   41
      Top             =   1860
      Width           =   900
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '����
      Caption         =   "��� ����"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   12
      Left            =   7680
      TabIndex        =   40
      Top             =   2340
      Width           =   915
   End
   Begin VB.Label lblRCDT 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "1234567890"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   8640
      TabIndex        =   39
      Top             =   2340
      Width           =   900
   End
   Begin VB.Label lblSDDT 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "1234567890"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   8640
      TabIndex        =   38
      Top             =   2580
      Width           =   900
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '����
      Caption         =   "��� ����"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   11
      Left            =   7680
      TabIndex        =   37
      Top             =   2580
      Width           =   915
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '����
      Caption         =   "Rack/Pos"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   16
      Left            =   7680
      TabIndex        =   36
      Top             =   1860
      Width           =   915
   End
   Begin VB.Label lblDISKNOPOSNO 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "1234567890"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   8640
      TabIndex        =   35
      Top             =   1860
      Width           =   900
   End
   Begin VB.Label lblEXSEQ 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   12000
      TabIndex        =   32
      Top             =   1080
      Width           =   120
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '����
      Caption         =   "�˻� ȸ��"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   20
      Left            =   11040
      TabIndex        =   31
      Top             =   1080
      Width           =   915
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '����
      Caption         =   "��ü ��ȣ"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   19
      Left            =   7680
      TabIndex        =   30
      Top             =   1080
      Width           =   915
   End
   Begin VB.Label lblBARCD 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "1234567890"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   8640
      TabIndex        =   29
      Top             =   1080
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "��ü��ȣ�� ��������"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   10
      Left            =   7740
      TabIndex        =   28
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "��ü��ȣ�� �˻���"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   9
      Left            =   7740
      TabIndex        =   27
      Top             =   2820
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '����
      Caption         =   "���ۻ���"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   3180
      TabIndex        =   22
      Top             =   2820
      Width           =   795
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '����
      Caption         =   "�������"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   3180
      TabIndex        =   21
      Top             =   2460
      Width           =   795
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '����
      Caption         =   "�����ڸ�"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   20
      Top             =   2820
      Width           =   795
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '����
      Caption         =   "���Ϲ�ȣ"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   19
      Top             =   2460
      Width           =   795
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '����
      Caption         =   "��ü��ȣ"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   18
      Top             =   2100
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "�˻縮��Ʈ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   17
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lbl���� 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "�˻�������"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   180
      TabIndex        =   16
      Top             =   60
      Width           =   2160
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808000&
      BackStyle       =   1  '�������� ����
      BorderColor     =   &H00000000&
      FillColor       =   &H0000C000&
      FillStyle       =   5  '���� �밢��
      Height          =   495
      Index           =   3
      Left            =   60
      Shape           =   4  '�ձ� �簢��
      Top             =   60
      Width           =   2595
   End
   Begin VB.Shape shpDResult 
      BackColor       =   &H00808000&
      BackStyle       =   1  '�������� ����
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFC0C0&
      FillStyle       =   5  '���� �밢��
      Height          =   255
      Left            =   7620
      Shape           =   4  '�ձ� �簢��
      Top             =   2820
      Width           =   5475
   End
   Begin VB.Shape shpDInfo 
      BackColor       =   &H00808000&
      BackStyle       =   1  '�������� ����
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFC0C0&
      FillStyle       =   5  '���� �밢��
      Height          =   255
      Left            =   7620
      Shape           =   4  '�ձ� �簢��
      Top             =   720
      Width           =   5475
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808000&
      BackStyle       =   1  '�������� ����
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFC0C0&
      FillStyle       =   5  '���� �밢��
      Height          =   255
      Index           =   0
      Left            =   60
      Shape           =   4  '�ձ� �簢��
      Top             =   720
      Width           =   7515
   End
End
Attribute VB_Name = "frmEQ_�˻�������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lngMeHeight     As Long '/Me.Height�� �ʱⰪ
Dim lngMeWidth      As Long '/Me.Width�� �ʱⰪ5

Private Type ConWhere   ' ����� ���� ������ ����ϴ�.
   Nm       As String
   Left     As Long
   Top      As Long
   Width    As Long
   Height   As Long
End Type
Dim CW()    As ConWhere

Public Sub SUB_MM_CANCEL()
    barStatus.Max = 100
    barStatus.Value = 100
    
    txtBARCD = ""
    txtPATNO = ""
    txtPATNM = ""
    cboSTATEFLAG.ListIndex = -1
    cboSENDFLAG.ListIndex = -1
    
    If sprLResult.MaxRows > 0 Then sprLResult.MaxRows = 0
    
    Call SUB_MM_KEY_CLEAR("1")
End Sub

Public Function FUNC_MM_DELETE() As Boolean
'    FUNC_MM_DELETE = False
'
'    Dim intActCol    As Integer
'    Dim intActRow    As Integer
'
'    '/1.���� ���� Check
'    If sprVIEW.ActiveRow = 0 Then MsgBox "������ ������ �����Ͻʽÿ�", vbInformation, "Ȯ��": Exit Function
'
'    '/2.���� ����
'    If MsgBox("���˻��ڵ� : " & GET_CELL(sprVIEW, 1, sprVIEW.ActiveRow) & vbCrLf & _
'              "���˻��   : " & GET_CELL(sprVIEW, 2, sprVIEW.ActiveRow) & vbCrLf & vbCrLf & _
'              "�� �ڷḦ �����ϰڽ��ϱ�?", vbQuestion + vbOKCancel, "��������") = vbCancel Then Exit Function
'
'    '/3.Process
'    If ConnDB_LOC = False Then Exit Function
'
'    ADC_LOC.BeginTrans
'
'    If sprVIEW.IsBlockSelected Then
'        intActCol = sprVIEW.SelBlockCol
'        intActRow = sprVIEW.SelBlockRow
'    Else
'        intActCol = sprVIEW.ActiveCol
'        intActRow = sprVIEW.ActiveRow
'    End If
'    If sprVIEW.IsBlockSelected Then
'        For intX = sprVIEW.SelBlockRow To sprVIEW.SelBlockRow2
'            gstrQuy = "DELETE FROM EQ_MST "
'            gstrQuy = gstrQuy & vbCrLf & " WHERE EQCD = '" & GET_CELL(sprVIEW, 1, intX) & "' "
'            If RunSQL_LOC(gstrQuy) = False Then ADC_LOC.RollbackTrans: Call CloseDB_LOC: Exit Function
'        Next intX
'    Else
'        gstrQuy = "DELETE FROM EQ_MST "
'        gstrQuy = gstrQuy & vbCrLf & " WHERE EQCD = '" & GET_CELL(sprVIEW, 1, sprVIEW.ActiveRow) & "' "
'        If RunSQL_LOC(gstrQuy) = False Then ADC_LOC.RollbackTrans: Call CloseDB_LOC: Exit Function
'    End If
'
'    ADC_LOC.CommitTrans
'
'    Call CloseDB_LOC
'
'    FUNC_MM_DELETE = True
'
'    MsgBox "�����Ǿ����ϴ�!", vbInformation, "Ȯ��"
'
'    '/4.ȭ��ó��
'    Call FUNC_MM_VIEW_LIST
'    sprVIEW.Col = intActCol
'    sprVIEW.Row = intActRow
'    sprVIEW.Action = ActionActiveCell
End Function

Private Sub SUB_MM_INITIAL()
    '/Form Resize�� ���� ��Ʈ�� �ʱⰪ �б�
    For intX = 0 To Me.Count - 1
        Select Case True
            Case TypeOf Me.Controls(intX) Is Line
            Case TypeOf Me.Controls(intX) Is CommonDialog
            Case Else
                ReDim Preserve CW(intX)
                
                CW(intX).Nm = Me.Controls(intX).Name
                CW(intX).Left = Me.Controls(intX).Left
                CW(intX).Top = Me.Controls(intX).Top
                CW(intX).Width = Me.Controls(intX).Width
                CW(intX).Height = Me.Controls(intX).Height
        End Select
    Next intX
    
    '/Form Resize�� ���� �ʱⰪ ����
    lngMeHeight = 9855
    lngMeWidth = 13275
    
    '/ȭ�� ��� ��ġ
    Me.Height = lngMeHeight
    Me.Width = lngMeWidth
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    '''Me.Show
    
    GoSub ADD_ITEM
    
    optDateSection(0).Value = True  '/��������/����:ó������
    dtpDateFrom.Value = Date        '/�ⰣFrom
    dtpDateTo.Value = Date          '/�ⰣTo
    
    Call SUB_MM_CANCEL
Exit Sub

'/------------------------------------------------------------------------------------------/

ADD_ITEM:
    '/���������� (0.ó��, 1.���)
    cboSTATEFLAG.AddItem ""
    cboSTATEFLAG.AddItem "0.ó��"
    cboSTATEFLAG.AddItem "1.���"
    
    '/HIS ���� FLAG (0.���, 1.�Ϸ�)
    cboSENDFLAG.AddItem ""
    cboSENDFLAG.AddItem "0.���"
    cboSENDFLAG.AddItem "1.�Ϸ�"
Return
End Sub

Public Sub SUB_MM_INPUT()
    gstrInputUpdate = "1" '/1.Input, 2.Update
    gstrInputUpdateYN = False

    frmEQ����_���˻��ڵ����_�Է�.Show vbModal

    If gstrInputUpdateYN = True Then
        Call FUNC_MM_VIEW_LIST
    End If
End Sub

Private Sub SUB_MM_KEY_CLEAR(ArgSection As String) '/ArgSection: 1.�˻縮��Ʈ, 2.��ü��ȣ��
    If ArgSection = "1" Then
        If sprLResult.MaxRows > 0 Then sprLResult.MaxRows = 0 '/�˻縮��Ʈ
    End If
    
    lblBARCD = ""       '/��ü��ȣ
    lblEXSEQ = ""       '/�˻�ȸ��
    lblDISKNOPOSNO = "" '/Rack/Pos
    lblEXDT = ""        '/�˻�ó����������
    lblRCDT = ""        '/�˻�����������
    lblSDDT = ""        '/�˻�����������
    lblPATNO = ""       '/���Ϲ�ȣ
    lblPATNM = ""       '/�����ڸ�
    lblORDDT = ""       '/ó������
    lblSEXAGE = ""      '/����/����
    lblORDGB = ""       '/��/�ܱ���
    
    If sprDResult.MaxRows > 0 Then sprDResult.MaxRows = 0 '/��ü��ȣ�� �˻���
        
End Sub

Public Sub SUB_MM_UPDATE()
'    Dim intActCol    As Integer
'    Dim intActRow    As Integer
'
'    If sprVIEW.ActiveRow = 0 Then MsgBox "������ ����� �����Ͻʽÿ�!", vbInformation, "Ȯ��": Exit Sub
'
'    gstrInputUpdate = "2" '/1.Input, 2.Update
'    gstrInputUpdateYN = False
'    gstrArgTemp1 = GET_CELL(sprVIEW, 1, sprVIEW.ActiveRow)
'
'    frmEQ����_���˻��ڵ����_�Է�.Show vbModal
'
'    If gstrInputUpdateYN = True Then
'        intActCol = sprVIEW.ActiveCol
'        intActRow = sprVIEW.ActiveRow
'
'        Call FUNC_MM_VIEW_LIST
'
'        sprVIEW.Col = intActCol
'        sprVIEW.Row = intActRow
'        sprVIEW.Action = ActionActiveCell
'    End If
End Sub

Public Function FUNC_MM_VIEW_LIST() As Boolean
    FUNC_MM_VIEW_LIST = False
    
On Error GoTo RTN_ERR
    
    Call SUB_MM_KEY_CLEAR("1")
    
    If ConnDB_LOC = False Then Exit Function
    
    With sprLResult
        gstrQuy = "SELECT BARCD, EXSEQ, SAMPLENO, DISKNO, POSNO, "
        gstrQuy = gstrQuy & vbCrLf & "       MAX(STATEFLAG) AS STATEFLAG, "
        gstrQuy = gstrQuy & vbCrLf & "       MAX(SENDFLAG)  AS SENDFLAG, "
        gstrQuy = gstrQuy & vbCrLf & "       MAX(EXDT+' '+EXTM) AS EXDT, "
        gstrQuy = gstrQuy & vbCrLf & "       MAX(RCDT+' '+RCTM) AS RCDT, "
        gstrQuy = gstrQuy & vbCrLf & "       MAX(SDDT+' '+SDTM) AS SDDT, "
        gstrQuy = gstrQuy & vbCrLf & "       MAX(ORDDT)     AS ORDDT, "
        gstrQuy = gstrQuy & vbCrLf & "       MAX(ORDGB)     AS ORDGB, "
        gstrQuy = gstrQuy & vbCrLf & "       MAX(PATNO)     AS PATNO, "
        gstrQuy = gstrQuy & vbCrLf & "       MAX(PATNM)     AS PATNM, "
        gstrQuy = gstrQuy & vbCrLf & "       MAX(PATSEX)    AS PATSEX, "
        gstrQuy = gstrQuy & vbCrLf & "       MAX(PATAGE)    AS PATAGE  "
        gstrQuy = gstrQuy & vbCrLf & "  FROM PAT_RES "
        Select Case True
            Case optDateSection(0).Value '/�˻�ó����������
                gstrQuy = gstrQuy & vbCrLf & " WHERE EXDT >= '" & Format(dtpDateFrom.Value, "YYYYMMDD") & "' "
                gstrQuy = gstrQuy & vbCrLf & "   AND EXDT <= '" & Format(dtpDateTo.Value, "YYYYMMDD") & "' "
            
            Case optDateSection(1).Value '/�˻�����������
                gstrQuy = gstrQuy & vbCrLf & " WHERE RCDT >= '" & Format(dtpDateFrom.Value, "YYYYMMDD") & "' "
                gstrQuy = gstrQuy & vbCrLf & "   AND RCDT <= '" & Format(dtpDateTo.Value, "YYYYMMDD") & "' "
            
            Case optDateSection(2).Value '/�˻�����������
                gstrQuy = gstrQuy & vbCrLf & " WHERE SDDT >= '" & Format(dtpDateFrom.Value, "YYYYMMDD") & "' "
                gstrQuy = gstrQuy & vbCrLf & "   AND SDDT <= '" & Format(dtpDateTo.Value, "YYYYMMDD") & "' "
            
            Case optDateSection(3).Value '/ó������
                gstrQuy = gstrQuy & vbCrLf & " WHERE ORDDT >= '" & Format(dtpDateFrom.Value, "YYYYMMDD") & "' "
                gstrQuy = gstrQuy & vbCrLf & "   AND ORDDT <= '" & Format(dtpDateTo.Value, "YYYYMMDD") & "' "
        
        End Select
        
        '/��ü��ȣ
        If Trim(txtBARCD) <> "" Then
            gstrQuy = gstrQuy & vbCrLf & "   AND BARCD = '" & Trim(txtBARCD) & "' "
        End If
        
        '/���Ϲ�ȣ
        If Trim(txtPATNO) <> "" Then
            gstrQuy = gstrQuy & vbCrLf & "   AND PATNO = '" & Trim(txtPATNO) & "' "
        End If
        
        '/�����ڸ�
        If Trim(txtPATNM) <> "" Then
            gstrQuy = gstrQuy & vbCrLf & "   AND PATNM LIKE '%" & Trim(txtPATNM) & "%' "
        End If
        
        '/���������� (0:ó��, 1:���)
        If Trim(cboSTATEFLAG) <> "" Then
            gstrQuy = gstrQuy & vbCrLf & "   AND STATEFLAG = '" & Trim(Left(cboSTATEFLAG, 1)) & "' "
        End If
        
        '/HIS ���� FLAG (0:���, 1:�Ϸ�)
        If Trim(cboSENDFLAG) <> "" Then
            gstrQuy = gstrQuy & vbCrLf & "   AND SENDFLAG = '" & Trim(Left(cboSENDFLAG, 1)) & "' "
        End If
        
        gstrQuy = gstrQuy & vbCrLf & " GROUP BY BARCD, EXSEQ, SAMPLENO, DISKNO, POSNO "
        gstrQuy = gstrQuy & vbCrLf & " ORDER BY BARCD, EXSEQ, SAMPLENO, DISKNO, POSNO "
        If ReadSQL_LOC(gstrQuy, ADR_LOC) = False Then Call CloseDB_LOC: Exit Function
        
        If Not ADR_LOC Is Nothing Then
            .MaxRows = ARC_LOC
            barStatus.Max = ARC_LOC
            intX = 0
            
            Do Until ADR_LOC.EOF
                intX = intX + 1: .Row = intX: barStatus.Value = intX
                
                .Col = 2: .Text = Trim(ADR_LOC!BARCD & "")     '/��ü��ȣ(Barcode)
                .Col = 3: .Text = Trim(ADR_LOC!PATNM & "") '/�����ڸ�
                .Col = 4: .Text = Trim(ADR_LOC!EXSEQ & "")     '/��ü��ȣ(Barcode)�� �˻�ȸ��
                .Col = 5: .Text = Trim(ADR_LOC!SAMPLENO & "")  '/Sample No
                .Col = 6: .Text = Trim(ADR_LOC!DISKNO & "")    '/��ũ��ȣ or ����ȣ
                .Col = 7: .Text = Trim(ADR_LOC!POSNO & "")     '/��ġ��ȣ
                
                .Col = 8                                        '/���������� (0:ó��, 1:���)
                Select Case Trim(ADR_LOC!STATEFLAG & "")
                    Case "0": .Text = "ó��"
                    Case "1": .Text = "���"
                End Select
                
                .Col = 9                                        '/HIS ���� FLAG (0:���, 1:�Ϸ�)
                Select Case Trim(ADR_LOC!SENDFLAG & "")
                    Case "0": .Text = "���"
                    Case "1": .Text = "�Ϸ�"
                End Select
                
                .Col = 10: '/�˻�ó����������
                If Trim(ADR_LOC!EXDT & "") <> "" Then
                    .Text = Format(Left(Trim(ADR_LOC!EXDT & ""), 8), "@@@@-@@-@@") & " " & Format(Mid(Trim(ADR_LOC!EXDT & ""), 10), "@@:@@:@@")
                End If
                .Col = 11 '/�˻�����������
                If Trim(ADR_LOC!RCDT & "") <> "" Then
                    .Text = Format(Left(Trim(ADR_LOC!RCDT & ""), 8), "@@@@-@@-@@") & " " & Format(Mid(Trim(ADR_LOC!RCDT & ""), 10), "@@:@@:@@")
                End If
                .Col = 12 '/�˻�����������
                If Trim(ADR_LOC!SDDT & "") <> "" Then
                    .Text = Format(Left(Trim(ADR_LOC!SDDT & ""), 8), "@@@@-@@-@@") & " " & Format(Mid(Trim(ADR_LOC!SDDT & ""), 10), "@@:@@:@@")
                End If
                .Col = 13 '/ó������
                If Trim(ADR_LOC!ORDDT & "") <> "" Then
                    .Text = Format(Trim(ADR_LOC!ORDDT & ""), "@@@@-@@-@@")
                End If
                
                .Col = 14
                Select Case Trim(ADR_LOC!ORDGB & "") '/ó������(O.�ܷ�, I.�Կ�, G.�ǰ�����)
                    Case "O": .Text = "�ܷ�"
                    Case "I": .Text = "�Կ�"
                    Case "G": .Text = "����"
                End Select
                
                .Col = 15: .Text = Trim(ADR_LOC!PATNO & "") '/���Ϲ�ȣ
                
                If Trim(ADR_LOC!PATSEX & "") <> "" Or Trim(ADR_LOC!PATAGE & "") <> "" Then
                    .Col = 16: .Text = Trim(ADR_LOC!PATSEX & "") & "/" & Trim(ADR_LOC!PATAGE & "") '/Sex/Age
                End If
                
                If .MaxTextRowHeight(intX) > 13.3 Then .RowHeight(intX) = .MaxTextRowHeight(intX)
                
                ADR_LOC.MoveNext
            Loop
            ADR_LOC.Close: Set ADR_LOC = Nothing
        Else
            MsgBox "�ڷᰡ �����ϴ�.", vbInformation, "Ȯ��"
        End If
    End With

    Call CloseDB_LOC

    FUNC_MM_VIEW_LIST = True
    
Exit Function

'/----------------------------------------------------------------------------------------------------/

RTN_ERR:

End Function

Public Function FUNC_MM_VIEW_RSLT(argBARCD As String, argEXSEQ As String) As Boolean
    FUNC_MM_VIEW_RSLT = False
    
On Error GoTo RTN_ERR
    
    If ConnDB_LOC = False Then Exit Function
    
    With sprDResult
        gstrQuy = "SELECT B.EQCD, B.EQNM, "
        gstrQuy = gstrQuy & vbCrLf & " A.EXAMCD, A.Result, A.EQRESULT,"
        gstrQuy = gstrQuy & vbCrLf & " A.AFLAG, A.PFLAG, A.DFLAG, "
        gstrQuy = gstrQuy & vbCrLf & " A.STATEFLAG, A.SENDFLAG, "
        gstrQuy = gstrQuy & vbCrLf & " A.ORDDT, A.EXDT, A.RCDT, A.SDDT, A.EXTM, A.RCTM, A.SDTM "
        gstrQuy = gstrQuy & vbCrLf & "  FROM PAT_RES A, EQ_MST B "
        gstrQuy = gstrQuy & vbCrLf & " WHERE A.EQCD  = B.EQCD "
        gstrQuy = gstrQuy & vbCrLf & "   AND A.BARCD = '" & Trim(argBARCD) & "' "
        gstrQuy = gstrQuy & vbCrLf & "   AND A.EXSEQ =  " & Val(argEXSEQ) & " "
        gstrQuy = gstrQuy & vbCrLf & " ORDER BY B.EQSEQ "
        If ReadSQL_LOC(gstrQuy, ADR_LOC) = False Then Call CloseDB_LOC: Exit Function
        
        If Not ADR_LOC Is Nothing Then
            .MaxRows = ARC_LOC
            barStatus.Max = ARC_LOC
            intX = 0
            
            Do Until ADR_LOC.EOF
                intX = intX + 1: .Row = intX: barStatus.Value = intX
                
                .Col = 1:  .Text = Trim(ADR_LOC!EQCD & "")     '/���˻��ڵ�
                .Col = 2:  .Text = Trim(ADR_LOC!EXAMCD & "")     '/�˻��ڵ�
                .Col = 3:  .Text = Trim(ADR_LOC!EQNM & "")     '/�˻���
                .Col = 4:  .Text = Trim(ADR_LOC!Result & "")     '/�˻���
                .Col = 5:  .Text = Trim(ADR_LOC!EQRESULT & "")     '/�����
                .Col = 6:  .Text = Trim(ADR_LOC!AFLAG & "")     '/R
                .Col = 7:  .Text = Trim(ADR_LOC!PFLAG & "")     '/P
                .Col = 8:  .Text = Trim(ADR_LOC!DFLAG & "")     '/D
                
                .Col = 9                                        '/���������� (0:ó��, 1:���)
                Select Case Trim(ADR_LOC!STATEFLAG & "")
                    Case "0": .Text = "ó��"
                    Case "1": .Text = "���"
                End Select
                
                .Col = 10                                        '/HIS ���� FLAG (0:���, 1:�Ϸ�)
                Select Case Trim(ADR_LOC!SENDFLAG & "")
                    Case "0": .Text = "���"
                    Case "1": .Text = "�Ϸ�"
                End Select
                .Col = 11: .Text = Trim(ADR_LOC!ORDDT & "")     '/ó������
                .Col = 12: .Text = Trim(ADR_LOC!EXDT & "") & " " & Trim(ADR_LOC!EXTM & "") '/�˻�ó�������Ͻ�
                .Col = 13: .Text = Trim(ADR_LOC!RCDT & "") & " " & Trim(ADR_LOC!RCTM & "") '/�˻��������Ͻ�
                .Col = 14: .Text = Trim(ADR_LOC!SDDT & "") & " " & Trim(ADR_LOC!SDTM & "") '/�˻��������Ͻ�
                
                If .MaxTextRowHeight(intX) > 13.3 Then .RowHeight(intX) = .MaxTextRowHeight(intX)
                
                ADR_LOC.MoveNext
            Loop
            ADR_LOC.Close: Set ADR_LOC = Nothing
        Else
            MsgBox "�ڷᰡ �����ϴ�.", vbInformation, "Ȯ��"
        End If
    End With

    Call CloseDB_LOC

    FUNC_MM_VIEW_RSLT = True
    
Exit Function

'/----------------------------------------------------------------------------------------------------/

RTN_ERR:

End Function

Private Sub cboSENDFLAG_Click()
    Call SUB_MM_KEY_CLEAR("1")
End Sub

Private Sub cboSENDFLAG_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cboSTATEFLAG_Click()
    Call SUB_MM_KEY_CLEAR("1")
End Sub

Private Sub cboSTATEFLAG_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cmdChange_Click()
    txtBeforBarcode.Enabled = True
    Call FUNC_BARCODE_CHANGE(txtBeforBarcode, txtAfterBarcode, lblRowNo)
    fraChange.Visible = False
    Call cmdView_Click
End Sub

Private Sub cmdChangeCancel_Click()
    txtBeforBarcode.Enabled = True
    fraChange.Visible = False
End Sub

Private Sub cmdDel_Click()
    '/�������ڵ��ȣ�� ���� ������ ����
    gtypPAT_RES.SAMPLENO = GET_CELL(sprLResult, 5, lblRowNo)
    Call FUNC_LOC_DELETE_PAT_RES(GET_CELL(sprLResult, 2, lblRowNo), GET_CELL(sprLResult, 4, lblRowNo))
    Call sprLResult.DeleteRows(lblRowNo, 1)
    sprLResult.MaxRows = sprLResult.MaxRows - 1
    '/�������ڵ��ȣ�� ���� ������ ����
    fraChange.Visible = False
    
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim intRow      As Integer
    
    For intRow = sprLResult.MaxRows To 1 Step -1
        sprLResult.Col = 1
        sprLResult.Row = intRow
        
        If sprLResult.Value = 1 Then
            gtypPAT_RES.BARCD = GET_CELL(sprLResult, 2, intRow)
            gtypPAT_RES.EXSEQ = GET_CELL(sprLResult, 4, intRow)
            gtypPAT_RES.SAMPLENO = GET_CELL(sprLResult, 5, intRow)
            gtypPAT_RES.DISKNO = GET_CELL(sprLResult, 6, intRow)
            gtypPAT_RES.POSNO = GET_CELL(sprLResult, 7, intRow)
            If FUNC_HIS_SAVE_MANUAL(intRow) = True Then
                sprLResult.Col = 1
                sprLResult.Row = intRow
                sprLResult.Value = 0
                Call sprLResult.DeleteRows(intRow, 1)
                sprLResult.MaxRows = sprLResult.MaxRows - 1
            Else
                Call SET_CELL(sprLResult, 9, intRow, "����")
            End If
        End If
        
    Next intRow

End Sub

Private Sub cmdView_Click()
    Call FUNC_MM_VIEW_LIST
    If sprLResult.MaxRows > 0 Then sprLResult.SetFocus
End Sub

Private Sub dtpDateFrom_Change()
    Call SUB_MM_KEY_CLEAR("1")
End Sub

Private Sub dtpDateFrom_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub dtpDateTo_Change()
    Call SUB_MM_KEY_CLEAR("1")
End Sub

Private Sub dtpDateTo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If fraChange.Visible = True And KeyAscii = vbKeyEscape Then fraChange.Visible = False
End Sub

Private Sub Form_Load()
    Call SUB_MM_INITIAL
    cboSENDFLAG.ListIndex = 1
'''    Call FUNC_MM_VIEW_LIST
End Sub

Private Sub Form_Resize()
    Dim intCnt  As Integer
    
On Error Resume Next
    '/object.Move Left, Top, Width, Height
    '/(((Me.Height - lngMeHeight) / 3) * 2) : ���̰� �þ�� ��ü 3��, �����λ� �ش� ��ü ���� �þ ��ü�� 2��
    For intCnt = 0 To UBound(CW)
        Select Case CW(intCnt).Nm
            Case cmdSave.Name:      cmdSave.Move CW(intCnt).Left + (Me.Width - lngMeWidth), CW(intCnt).Top, CW(intCnt).Width, CW(intCnt).Height
            Case cmdQuit.Name:      cmdQuit.Move CW(intCnt).Left + (Me.Width - lngMeWidth), CW(intCnt).Top, CW(intCnt).Width, CW(intCnt).Height
            Case barStatus.Name: barStatus.Move CW(intCnt).Left, CW(intCnt).Top, CW(intCnt).Width + (Me.Width - lngMeWidth), CW(intCnt).Height
            Case sprLResult.Name:   sprLResult.Move CW(intCnt).Left, CW(intCnt).Top, CW(intCnt).Width, CW(intCnt).Height + (Me.Height - lngMeHeight)
            Case shpDInfo.Name: shpDInfo.Move CW(intCnt).Left, CW(intCnt).Top, CW(intCnt).Width + (Me.Width - lngMeWidth), CW(intCnt).Height
            Case shpDResult.Name: shpDResult.Move CW(intCnt).Left, CW(intCnt).Top, CW(intCnt).Width + (Me.Width - lngMeWidth), CW(intCnt).Height
            Case sprDResult.Name:   sprDResult.Move CW(intCnt).Left, CW(intCnt).Top, CW(intCnt).Width + (Me.Width - lngMeWidth), CW(intCnt).Height + (Me.Height - lngMeHeight)
        End Select
    Next intCnt
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CloseDB_LOC
    Set frmEQ����_���˻��ڵ����_��ȸ = Nothing
End Sub

Private Sub lblcheck_Click()
    '/�������� ��ü ���� ����
    Dim intRow  As Integer
    
    For intRow = 1 To sprLResult.MaxRows
        
        If Trim(Mid(lblcheck.Caption, 5)) = 1 Then
            sprLResult.Row = intRow
            sprLResult.Col = 1
            sprLResult.Value = 0
        ElseIf Trim(Mid(lblcheck.Caption, 5)) = 0 Then
            sprLResult.Row = intRow
            sprLResult.Col = 1
            sprLResult.Value = 1
        End If
    Next intRow
    
    If Trim(Mid(lblcheck.Caption, 5)) = 1 Then
        lblcheck.Caption = Mid(lblcheck.Caption, 1, 4) & "     0"
    Else
        lblcheck.Caption = Mid(lblcheck.Caption, 1, 4) & "     1"
    End If
End Sub

Private Sub optDateSection_Click(Index As Integer)
    Call SUB_MM_KEY_CLEAR("1")
End Sub

Private Sub sprLResult_Click(ByVal Col As Long, ByVal Row As Long)
        With sprLResult
        If Row <> 0 Then Exit Sub
        If Col = 0 Then Exit Sub
        
        .Col = -1
        .Row = 1
        .Col2 = -1
        .Row2 = .MaxRows
        .BlockMode = True
        .SortBy = SortByRow
        
        .SortKey(1) = Col
        If Val(Mid(sprLResult.Tag, 2)) = Col Then
            If Left(sprLResult.Tag, 1) = "A" Then
                .SortKeyOrder(1) = SortKeyOrderDescending
                sprLResult.Tag = "D" & CStr(Col)
            Else
                .SortKeyOrder(1) = SortKeyOrderAscending
                sprLResult.Tag = "A" & CStr(Col)
            End If
        Else
            .SortKeyOrder(1) = SortKeyOrderAscending
            sprLResult.Tag = "A" & CStr(Col)
        End If
        
        .Action = ActionSort
        .BlockMode = False
    End With
End Sub

Private Sub sprLResult_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    If Row < 1 Or Col = 1 Then Exit Sub
    If fraChange.Visible = False Then fraChange.Visible = True
        fraChange.Left = 3240
        fraChange.Top = 3600
        lblRowNo = Row
        txtBeforBarcode = GET_CELL(sprLResult, 2, Row)
        txtBeforBarcode.Enabled = False
        txtAfterBarcode = ""
        sprLResult.Action = ActionActiveCell
        'sprLResult.ActiveRow = Row
End Sub

Private Sub sprLResult_DblClick(ByVal Col As Long, ByVal Row As Long)
    If Col < 2 Then Exit Sub
    If Row < 1 Then Exit Sub
    
    Call SUB_MM_KEY_CLEAR("2")

    lblBARCD = GET_CELL(sprLResult, 2, Row)     '/��ü��ȣ
    lblPATNM = GET_CELL(sprLResult, 3, Row)    '/�����ڸ�
    lblEXSEQ = GET_CELL(sprLResult, 4, Row)     '/�˻�ȸ��
    lblSAMPLENO = GET_CELL(sprLResult, 5, Row)  '/Sample No
    lblDISKNOPOSNO = GET_CELL(sprLResult, 6, Row) & "/" & GET_CELL(sprLResult, 7, Row) '/Rack/Pos
    lblEXDT = GET_CELL(sprLResult, 10, Row)      '/�˻�ó����������
    lblRCDT = GET_CELL(sprLResult, 11, Row)     '/�˻�����������
    lblSDDT = GET_CELL(sprLResult, 12, Row)     '/�˻�����������
    lblORDDT = GET_CELL(sprLResult, 13, Row)    '/ó������
    lblORDGB = GET_CELL(sprLResult, 14, Row)    '/��/�ܱ���
    lblPATNO = GET_CELL(sprLResult, 15, Row)    '/���Ϲ�ȣ
    lblSEXAGE = GET_CELL(sprLResult, 16, Row)   '/����/����
    
    Call FUNC_MM_VIEW_RSLT(GET_CELL(sprLResult, 2, Row), GET_CELL(sprLResult, 4, Row))
    sprLResult.Row = Row
    sprLResult.Col = 1
    If sprLResult.Value = 1 Then
        sprLResult.Value = 0
    Else
        sprLResult.Value = 1
    End If
    
End Sub

Private Sub sprLResult_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call sprLResult_DblClick(sprLResult.ActiveCol, sprLResult.ActiveRow)
End Sub

Private Sub optDateSection_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtBARCD_Change()
    Call SUB_MM_KEY_CLEAR("1")
End Sub

Private Sub txtBARCD_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub txtBARCD_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtPATNM_Change()
    Call SUB_MM_KEY_CLEAR("1")
End Sub

Private Sub txtPATNM_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub txtPATNM_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtPATNO_Change()
    Call SUB_MM_KEY_CLEAR("1")
End Sub

Private Sub txtPATNO_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub txtPATNO_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Public Function FUNC_HIS_SAVE_MANUAL(argRow As Integer) As Boolean
    
    FUNC_HIS_SAVE_MANUAL = False
On Error GoTo RTN_ERR
    If ConnDB_LOC = False Then Exit Function
                       gstrQuy = "SELECT * "
    gstrQuy = gstrQuy & vbCrLf & "  FROM PAT_RES "
    gstrQuy = gstrQuy & vbCrLf & " WHERE BARCD     = '" & gtypPAT_RES.BARCD & "' "
    gstrQuy = gstrQuy & vbCrLf & "   AND EXSEQ     =  " & Val(gtypPAT_RES.EXSEQ) & " "
    gstrQuy = gstrQuy & vbCrLf & "   AND SAMPLENO  = '" & gtypPAT_RES.SAMPLENO & "' "
    gstrQuy = gstrQuy & vbCrLf & "   AND DISKNO    = '" & gtypPAT_RES.DISKNO & "' "
    gstrQuy = gstrQuy & vbCrLf & "   AND POSNO     = '" & gtypPAT_RES.POSNO & "' "
    If ReadSQL_LOC(gstrQuy, ADR_LOC) = False Then Call CloseDB_LOC: End
    
    If Not ADR_LOC Is Nothing Then
        Do Until ADR_LOC.EOF
            gtypPAT_RES.Result = Trim(ADR_LOC!Result & "")      '/���
            gtypPAT_RES.EQCD = Trim(ADR_LOC!EQCD & "")          '/���˻��ڵ�
            gtypPAT_RES.EXAMCD = Trim(ADR_LOC!EXAMCD & "")            '/ó��˻��ڵ�
            gtypPAT_RES.DFLAG = Trim(ADR_LOC!DFLAG & "")
            gtypPAT_RES.PFLAG = Trim(ADR_LOC!PFLAG & "")
            gtypPAT_RES.AFLAG = Trim(ADR_LOC!AFLAG & "")
            
            If gtypPAT_RES.EXAMCD <> "" Then  'gtypPAT_RES.Result <> "" And
                If FUNC_HIS_SAVE = True Then  '/HIS�� ��� ����
                    Call FUNC_LOC_SAVE_SEND(gtypPAT_RES.BARCD, gtypPAT_RES.EXSEQ, gtypPAT_RES.EQCD, gtypPAT_RES.SAMPLENO, gtypPAT_RES.DISKNO, gtypPAT_RES.POSNO, "1") '/HIS�� ��� ����
                    Call SET_CELL(frmEQ_�˻�������.sprLResult, 8, argRow, IIf(gtypPAT_RES.SENDFLAG = "1", "�Ϸ�", "���"))
                    FUNC_HIS_SAVE_MANUAL = True
                Else
                    FUNC_HIS_SAVE_MANUAL = False
                End If
            Else
                Call FUNC_LOC_SAVE_SEND(gtypPAT_RES.BARCD, gtypPAT_RES.EXSEQ, gtypPAT_RES.EQCD, gtypPAT_RES.SAMPLENO, gtypPAT_RES.DISKNO, gtypPAT_RES.POSNO, "1") '/HIS�� ��� ����
                Call SET_CELL(frmEQ_�˻�������.sprLResult, 8, argRow, IIf(gtypPAT_RES.SENDFLAG = "1", "�Ϸ�", "���"))
            End If
            ADR_LOC.MoveNext
        Loop
        
    End If
    
    Call CloseDB_LOC
Exit Function
    
'/----------------------------------------------------------------------------------------------------/
    
RTN_ERR:
    FUNC_HIS_SAVE_MANUAL = False
End Function

Public Function FUNC_BARCODE_CHANGE(argBF_BARCD As String, argAF_BARCD As String, argRow As Integer) As Boolean
    '/�� �˻� �� ��Ī�� ���
    Dim intCol      As Integer
    Dim strEXSEQ    As String
    Dim EXCD_LIST   As String
    FUNC_BARCODE_CHANGE = False

    gtypPAT_RES.BARCD = GET_CELL(sprLResult, 2, argRow)      '/��ü��ȣ
    gtypPAT_RES.PATNM = GET_CELL(sprLResult, 3, argRow)    '/�����ڸ�
    gtypPAT_RES.EXSEQ = GET_CELL(sprLResult, 4, argRow)     '/�˻�ȸ��
    gtypPAT_RES.SAMPLENO = GET_CELL(sprLResult, 5, argRow)   '/Sample No
    gtypPAT_RES.DISKNO = GET_CELL(sprLResult, 6, argRow)
    gtypPAT_RES.POSNO = GET_CELL(sprLResult, 7, argRow) '/Rack/Pos
    gtypPAT_RES.EXDT = GET_CELL(sprLResult, 10, argRow)      '/�˻�ó����������
    gtypPAT_RES.RCDT = GET_CELL(sprLResult, 11, argRow)     '/�˻�����������
    gtypPAT_RES.SDDT = GET_CELL(sprLResult, 12, argRow)     '/�˻�����������
    gtypPAT_RES.ORDDT = GET_CELL(sprLResult, 13, argRow)    '/ó������
    gtypPAT_RES.ORDGB = GET_CELL(sprLResult, 14, argRow)    '/��/�ܱ���
    gtypPAT_RES.PATNO = GET_CELL(sprLResult, 15, argRow)    '/���Ϲ�ȣ

    '/ó����������,�����������,����������� ��ȯ
    gtypPAT_RES.EXDT = Format(Mid(GET_CELL(sprLResult, 10, argRow), 1, 10), "YYYYMMDD")
    gtypPAT_RES.EXTM = Format(Mid(GET_CELL(sprLResult, 10, argRow), 12), "HHMMSS")
    gtypPAT_RES.RCDT = Format(Mid(GET_CELL(sprLResult, 11, argRow), 1, 10), "YYYYMMDD")
    gtypPAT_RES.RCTM = Format(Mid(GET_CELL(sprLResult, 11, argRow), 12), "HHMMSS")
    gtypPAT_RES.SDDT = Format(Mid(GET_CELL(sprLResult, 12, argRow), 1, 10), "YYYYMMDD")
    gtypPAT_RES.SDTM = Format(Mid(GET_CELL(sprLResult, 12, argRow), 12), "HHMMSS")
    

    If ConnDB_LOC = True Then
        '/����ڵ庰 ó���ڵ� ��������
                 gstrQuy = "SELECT * "
        gstrQuy = gstrQuy & vbCrLf & "  FROM PAT_RES "
        gstrQuy = gstrQuy & vbCrLf & " WHERE BARCD      = '" & gtypPAT_RES.BARCD & "' "
        gstrQuy = gstrQuy & vbCrLf & "   AND SAMPLENO   = '" & gtypPAT_RES.SAMPLENO & "' "
        gstrQuy = gstrQuy & vbCrLf & "   AND DISKNO     = '" & gtypPAT_RES.DISKNO & "' "
        gstrQuy = gstrQuy & vbCrLf & "   AND POSNO      = '" & gtypPAT_RES.POSNO & "' "
        gstrQuy = gstrQuy & vbCrLf & "   AND EXSEQ      =  " & gtypPAT_RES.EXSEQ & " "
        gstrQuy = gstrQuy & vbCrLf & "   AND PATNO      = '" & gtypPAT_RES.PATNO & "' "
        gstrQuy = gstrQuy & vbCrLf & "   AND EXDT       = '" & gtypPAT_RES.EXDT & "' "
        gstrQuy = gstrQuy & vbCrLf & "   AND EXTM       = '" & gtypPAT_RES.EXTM & "' "
        gstrQuy = gstrQuy & vbCrLf & "   AND RCDT       = '" & gtypPAT_RES.RCDT & "' "
        gstrQuy = gstrQuy & vbCrLf & "   AND RCTM       = '" & gtypPAT_RES.RCTM & "' "
        gstrQuy = gstrQuy & vbCrLf & "   AND SDDT       = '" & gtypPAT_RES.SDDT & "' "
        gstrQuy = gstrQuy & vbCrLf & "   AND SDTM       = '" & gtypPAT_RES.SDTM & "' "
        gstrQuy = gstrQuy & vbCrLf & "   AND ORDDT      = '" & gtypPAT_RES.ORDDT & "' "
        gstrQuy = gstrQuy & vbCrLf & "   AND ORDGB      = '" & gtypPAT_RES.ORDGB & "' "
        gstrQuy = gstrQuy & vbCrLf & " ORDER BY EXSEQ DESC "

        If ReadSQL_LOC(gstrQuy, ADR_LOC_BAR) = False Then Call CloseDB_LOC: End
        If Not ADR_LOC_BAR Is Nothing Then
            gtypPAT_RES.EXDT = Format(Now, "YYYYMMDD")
            gtypPAT_RES.EXTM = Format(Now, "HHMMSS")
            gtypPAT_RES.RCDT = Format(Now, "YYYYMMDD")      '/RCDT(�˻�����������(YYYYMMDD) �Ƿ���� ->HIEQ)
            gtypPAT_RES.RCTM = Format(Now, "HHMMSS")        '/RCTM(�˻������Žð�(24HHMMSS) �Ƿ���� ->HIEQ)
            
            Do Until ADR_LOC_BAR.EOF
                gtypPAT_RES.BARCD = Trim(argAF_BARCD)
                gtypPAT_RES.EQCD = Trim(ADR_LOC_BAR!EQCD & "")
                gtypPAT_RES.EQRESULT = Trim(ADR_LOC_BAR!EQRESULT & "")
                
                Call FUNC_HIS_PATIENT
                Call FUNC_HIS_ORDER_VIEW
                
                If IsNumeric(gtypPAT_RES.EQRESULT) = True Then
                    Call FUNC_HIS_RESULT_JUDGMENT '/��� ����
                Else
                    gtypPAT_RES.Result = gtypPAT_RES.EQRESULT
                End If
                
                gtypPAT_RES.STATEFLAG = "1"                     '/STATEFLAG(���������� (0:ó��, 1:���))
                gtypPAT_RES.SENDFLAG = "0"
                
                
                
                '/�˻�SEQ ã��
                If strEXSEQ <> "Y" Then
                    gtypPAT_RES.EXSEQ = FUNC_GET_EXSEQ(gtypPAT_RES.BARCD) '/��ü��ȣ(Barcode)�� �˻�ȸ��
                    strEXSEQ = "Y"
                End If
                
                Call SET_CELL(frmEQ_�˻�������.sprLResult, 8, argRow, IIf(gtypPAT_RES.SENDFLAG = "1", "�Ϸ�", "���"))
                
                
                

                Call FUNC_LOC_SAVE_PAT_RES                                                                      '/����� ���ڵ��ȣ�� ���� ������ �Է�
                
                gtypPAT_RES.EXAMCD = ""
                ADR_LOC_BAR.MoveNext
            Loop
        End If

        Call CloseDB_LOC

    End If
    '/�������ڵ��ȣ�� ���� ������ ����
    Call FUNC_LOC_DELETE_PAT_RES(GET_CELL(sprLResult, 2, argRow), GET_CELL(sprLResult, 4, argRow))
    '/�������ڵ��ȣ�� ���� ������ ����
    '/���� ���������� ���ڵ�, ó������,�������,������� ���� �� �ð��� ����
    Call SET_CELL(sprLResult, 2, argRow, gtypPAT_RES.BARCD)
    Call SET_CELL(sprLResult, 3, argRow, gtypPAT_RES.PATNM)
    Call SET_CELL(sprLResult, 4, argRow, gtypPAT_RES.EXSEQ)
    Call SET_CELL(sprLResult, 10, argRow, gtypPAT_RES.EXDT)
    Call SET_CELL(sprLResult, 10, argRow, gtypPAT_RES.EXTM)
    Call SET_CELL(sprLResult, 11, argRow, gtypPAT_RES.RCDT)
    Call SET_CELL(sprLResult, 11, argRow, gtypPAT_RES.RCTM)
    Call SET_CELL(sprLResult, 12, argRow, gtypPAT_RES.SDDT)
    Call SET_CELL(sprLResult, 12, argRow, gtypPAT_RES.SDTM)
    '/���� ���������� ���ڵ�, ó������,�������,������� ���� �� �ð��� ����
    
    gtypPAT_RES.BARCD = ""
    gtypPAT_RES.EXSEQ = ""            '/EXSEQ(��ü��ȣ(Barcode)�� �˻�ȸ��)
    gtypPAT_RES.EQCD = ""             '/EQCD(���˻��ڵ�)
    gtypPAT_RES.EXAMCD = ""           '/EXAMCD(ó���ڵ�(HIS or LIS�� �˻��ڵ�))
    gtypPAT_RES.EXDT = ""             '/EXDT(�˻�ó����������(YYYYMMDD) HIEQ->�Ƿ����)
    gtypPAT_RES.EXTM = ""             '/EXTM(�˻�ó�����۽ð�(24HHMMSS) HIEQ->�Ƿ����)
    gtypPAT_RES.RCDT = ""             '/RCDT(�˻�����������(YYYYMMDD) �Ƿ���� ->HIEQ)
    gtypPAT_RES.RCTM = ""             '/RCTM(�˻������Žð�(24HHMMSS) �Ƿ���� ->HIEQ)
    gtypPAT_RES.SDDT = ""             '/SDDT(�˻�����������(YYYYMMDD) HIEQ->HIS)
    gtypPAT_RES.SDTM = ""             '/SDTM(�˻������۽ð�(24HHMMSS) HIEQ->HIS)
    gtypPAT_RES.Result = ""           '/RESULT(�˻���(������ ���))
    gtypPAT_RES.EQRESULT = ""         '/EQRESULT(�����ð��)
    gtypPAT_RES.AFLAG = ""            '/AFLAG(Abnormal(��������ġ ���� (H)High or (L)Low �� ǥ��))
    gtypPAT_RES.PFLAG = ""            '/PFLAG(Panic)
    gtypPAT_RES.DFLAG = ""            '/DFLAG(Delta)
    gtypPAT_RES.SAMPLENO = ""         '/Sample No(AU2700, Uriscan � ���)
    gtypPAT_RES.DISKNO = ""           '/DISKNO(��ũ��ȣ or ����ȣ)
    gtypPAT_RES.POSNO = ""            '/POSNO(��ġ��ȣ)
    gtypPAT_RES.ORDDT = ""            '/ORDDT(ó������)
    gtypPAT_RES.ORDGB = ""            '/ORDGB(ó������(O.�ܷ�, I.�Կ�, G.�ǰ�����))
    gtypPAT_RES.PATNO = ""            '/PATNO(���Ϲ�ȣ)
    gtypPAT_RES.PATNM = ""            '/PATNM(�����ڸ�)
    gtypPAT_RES.PATSEX = ""           '/PATSEX(����)
    gtypPAT_RES.PATAGE = ""           '/PATAGE(����)
    gtypPAT_RES.SENDFLAG = ""         '/SENDFLAG(HIS ���� FLAG (0:���, 1:�Ϸ�))
    gtypPAT_RES.STATEFLAG = ""        '/STATEFLAG(���������� (0:ó��, 1:���))

    FUNC_BARCODE_CHANGE = True
    
    
End Function

