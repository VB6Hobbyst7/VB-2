VERSION 5.00
Object = "{8996B0A4-D7BE-101B-8650-00AA003A5593}#4.0#0"; "Cfx4032.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmCounsel_1 
   BorderStyle     =   0  '����
   Caption         =   "Form1"
   ClientHeight    =   9270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13515
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9270
   ScaleWidth      =   13515
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdInputTreat 
      Caption         =   "ó���Է�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11040
      TabIndex        =   67
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdPreTreat 
      Caption         =   "��ȸó�氪 �ҷ�����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   10410
      TabIndex        =   61
      Top             =   1860
      Width           =   1785
   End
   Begin VB.CheckBox chkUserCalory 
      Appearance      =   0  '���
      BackColor       =   &H80000005&
      Caption         =   "�Ĵܿ��������Է��ϱ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   7890
      TabIndex        =   60
      Top             =   4440
      Width           =   1965
   End
   Begin VB.TextBox txtUserCalory 
      Alignment       =   1  '������ ����
      Height          =   285
      Left            =   9870
      TabIndex        =   58
      Top             =   4380
      Width           =   495
   End
   Begin VB.TextBox txtMemo 
      Appearance      =   0  '���
      Height          =   2415
      Left            =   7890
      MultiLine       =   -1  'True
      TabIndex        =   57
      Text            =   "frmCounsel_1.frx":0000
      Top             =   5670
      Width           =   3255
   End
   Begin VB.TextBox txtNotice 
      Appearance      =   0  '���
      Height          =   2415
      Left            =   7890
      MultiLine       =   -1  'True
      TabIndex        =   56
      Text            =   "frmCounsel_1.frx":0006
      Top             =   5670
      Width           =   3255
   End
   Begin VB.ComboBox cmbProgram 
      Height          =   300
      Left            =   9210
      Style           =   2  '��Ӵٿ� ���
      TabIndex        =   21
      Top             =   1500
      Width           =   1485
   End
   Begin VB.CommandButton cmdSub 
      Caption         =   "��ü����"
      Height          =   315
      Index           =   9
      Left            =   6210
      TabIndex        =   20
      Top             =   7560
      Width           =   1245
   End
   Begin VB.CommandButton cmdSub 
      Caption         =   "��ü����"
      Height          =   315
      Index           =   8
      Left            =   6210
      TabIndex        =   19
      Top             =   7260
      Width           =   1245
   End
   Begin VB.CommandButton cmdSub 
      Caption         =   "��ü����"
      Height          =   315
      Index           =   7
      Left            =   6210
      TabIndex        =   18
      Top             =   6960
      Width           =   1245
   End
   Begin VB.CommandButton cmdSub 
      Caption         =   "��ü����"
      Height          =   315
      Index           =   6
      Left            =   6210
      TabIndex        =   17
      Top             =   6660
      Width           =   1245
   End
   Begin VB.CommandButton cmdSub 
      Caption         =   "��ü����"
      Height          =   315
      Index           =   5
      Left            =   6210
      TabIndex        =   16
      Top             =   6360
      Width           =   1245
   End
   Begin VB.CommandButton cmdSub 
      Caption         =   "��ü����"
      Height          =   315
      Index           =   4
      Left            =   6210
      TabIndex        =   15
      Top             =   6060
      Width           =   1245
   End
   Begin VB.CommandButton cmdSub 
      Caption         =   "R  M  R"
      Height          =   315
      Index           =   3
      Left            =   6210
      TabIndex        =   14
      Top             =   5760
      Width           =   1245
   End
   Begin VB.CommandButton cmdSub 
      Caption         =   "W  H  R"
      Height          =   315
      Index           =   2
      Left            =   6210
      TabIndex        =   13
      Top             =   5460
      Width           =   1245
   End
   Begin VB.CommandButton cmdSub 
      Caption         =   "�㸮�ѷ�"
      Height          =   315
      Index           =   1
      Left            =   6210
      TabIndex        =   12
      Top             =   5160
      Width           =   1245
   End
   Begin VB.CommandButton cmdSub 
      Caption         =   "ü      ��"
      Height          =   315
      Index           =   0
      Left            =   6210
      TabIndex        =   11
      Top             =   4860
      Width           =   1245
   End
   Begin MSFlexGridLib.MSFlexGrid grdTable 
      Height          =   3345
      Left            =   750
      TabIndex        =   9
      Top             =   4830
      Width           =   5445
      _ExtentX        =   9604
      _ExtentY        =   5900
      _Version        =   393216
      Rows            =   1
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorBkg    =   16777215
      BorderStyle     =   0
      Appearance      =   0
   End
   Begin ChartfxLibCtl.ChartFX Chart 
      Height          =   3345
      Left            =   750
      TabIndex        =   10
      Top             =   4860
      Width           =   5475
      _cx             =   9657
      _cy             =   5900
      Build           =   20
      TypeMask        =   109576194
      Axis(0).Max     =   90
      Axis(2).Format  =   5
      Axis(2).Format  =   5
      RGB2DBk         =   16777215
      nColors         =   16
      Colors          =   "frmCounsel_1.frx":000C
      nSer            =   1
      NumSer          =   1
      _Data_          =   "frmCounsel_1.frx":00AC
   End
   Begin FPSpread.vaSpread sprTreat 
      Height          =   1875
      Left            =   7890
      TabIndex        =   1
      Top             =   6210
      Width           =   1650
      _Version        =   196608
      _ExtentX        =   2910
      _ExtentY        =   3307
      _StockProps     =   64
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "frmCounsel_1.frx":01B9
   End
   Begin FPSpread.vaSpread sprPrint 
      Height          =   1875
      Left            =   9660
      TabIndex        =   0
      Top             =   6210
      Width           =   1455
      _Version        =   196608
      _ExtentX        =   2566
      _ExtentY        =   3307
      _StockProps     =   64
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "frmCounsel_1.frx":0341
   End
   Begin VB.Image imgNew 
      Height          =   795
      Left            =   11370
      Picture         =   "frmCounsel_1.frx":04C9
      Top             =   4905
      Width           =   795
   End
   Begin VB.Label lblDispContent 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "��ȸ��"
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
      Left            =   10185
      TabIndex        =   66
      Top             =   1065
      Width           =   585
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "���� ó�泻��"
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
      Left            =   8895
      TabIndex        =   65
      Top             =   1065
      Width           =   1245
   End
   Begin VB.Label lblDispDate 
      BackStyle       =   0  '����
      Caption         =   "2005-01-01"
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
      Left            =   7785
      TabIndex        =   64
      Top             =   1065
      Width           =   1125
   End
   Begin VB.Image imgValuation 
      Height          =   330
      Left            =   9720
      Picture         =   "frmCounsel_1.frx":0B9A
      Top             =   4830
      Width           =   1185
   End
   Begin VB.Label lblTreatCal 
      BackStyle       =   0  '����
      Caption         =   "1,500 kcal"
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   225
      Left            =   9600
      TabIndex        =   63
      Top             =   4170
      Width           =   1185
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '����
      Caption         =   "ó���� �Ĵܿ��� = "
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8160
      TabIndex        =   62
      Top             =   4170
      Width           =   1395
   End
   Begin VB.Image imgSelectEx 
      Height          =   330
      Left            =   10800
      Picture         =   "frmCounsel_1.frx":1E97
      Top             =   4260
      Width           =   1185
   End
   Begin VB.Image TopImage 
      Height          =   960
      Left            =   -30
      Picture         =   "frmCounsel_1.frx":29CD
      Top             =   50
      Width           =   13140
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '����
      Caption         =   "kcal"
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   10350
      TabIndex        =   59
      Top             =   4440
      Width           =   405
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000004&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  '�ܻ�
      Height          =   615
      Left            =   7800
      Shape           =   4  '�ձ� �簢��
      Top             =   4110
      Width           =   4305
   End
   Begin VB.Image imgTreat 
      Height          =   330
      Left            =   6240
      Picture         =   "frmCounsel_1.frx":4553
      Top             =   7860
      Width           =   1170
   End
   Begin VB.Image imgSelection 
      Height          =   285
      Left            =   6210
      Top             =   7890
      Width           =   1245
   End
   Begin VB.Label lblObInfo 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "31 %"
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   2010
      TabIndex        =   55
      Top             =   3330
      Width           =   600
   End
   Begin VB.Label lblAnaerobic 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "100��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   3
      Left            =   10590
      TabIndex        =   54
      Top             =   3870
      Width           =   615
   End
   Begin VB.Label lblAnaerobic 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "100��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   2
      Left            =   10590
      TabIndex        =   53
      Top             =   3540
      Width           =   615
   End
   Begin VB.Label lblAnaerobic 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "100��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   1
      Left            =   10590
      TabIndex        =   52
      Top             =   3210
      Width           =   615
   End
   Begin VB.Label lblAnaerobic 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "100��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   0
      Left            =   10590
      TabIndex        =   51
      Top             =   2910
      Width           =   615
   End
   Begin VB.Label lblAerobic 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "47��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   9870
      TabIndex        =   50
      Top             =   3870
      Width           =   675
   End
   Begin VB.Label lblAerobic 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "47��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   9870
      TabIndex        =   49
      Top             =   3540
      Width           =   675
   End
   Begin VB.Label lblAerobic 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "47��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   9870
      TabIndex        =   48
      Top             =   3210
      Width           =   675
   End
   Begin VB.Label lblAerobic 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "47��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   9870
      TabIndex        =   47
      Top             =   2910
      Width           =   675
   End
   Begin VB.Image imgExCal 
      Height          =   300
      Index           =   3
      Left            =   9180
      Picture         =   "frmCounsel_1.frx":5199
      Top             =   3795
      Width           =   645
   End
   Begin VB.Image imgExCal 
      Height          =   300
      Index           =   2
      Left            =   9180
      Picture         =   "frmCounsel_1.frx":57D9
      Top             =   3480
      Width           =   645
   End
   Begin VB.Image imgExCal 
      Height          =   300
      Index           =   1
      Left            =   9180
      Picture         =   "frmCounsel_1.frx":5DE6
      Top             =   3165
      Width           =   645
   End
   Begin VB.Image imgExCal 
      Height          =   300
      Index           =   0
      Left            =   9180
      Picture         =   "frmCounsel_1.frx":6423
      Top             =   2850
      Width           =   645
   End
   Begin VB.Image imgExDay 
      Height          =   300
      Index           =   3
      Left            =   8730
      Picture         =   "frmCounsel_1.frx":6A55
      Top             =   3795
      Width           =   435
   End
   Begin VB.Image imgExDay 
      Height          =   300
      Index           =   2
      Left            =   8730
      Picture         =   "frmCounsel_1.frx":6F44
      Top             =   3480
      Width           =   435
   End
   Begin VB.Image imgExDay 
      Height          =   300
      Index           =   1
      Left            =   8730
      Picture         =   "frmCounsel_1.frx":742F
      Top             =   3165
      Width           =   435
   End
   Begin VB.Image imgExDay 
      Height          =   300
      Index           =   0
      Left            =   8730
      Picture         =   "frmCounsel_1.frx":7926
      Top             =   2850
      Width           =   435
   End
   Begin VB.Label lblLossWeight 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "2.2 kg/��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   11250
      TabIndex        =   46
      Top             =   3390
      Width           =   825
   End
   Begin VB.Image imgPrint 
      Height          =   780
      Left            =   11370
      Picture         =   "frmCounsel_1.frx":7E1A
      Top             =   7350
      Width           =   750
   End
   Begin VB.Image imgDel 
      Height          =   795
      Left            =   11370
      Picture         =   "frmCounsel_1.frx":8EA5
      Top             =   6525
      Width           =   795
   End
   Begin VB.Image imgSave 
      Height          =   795
      Left            =   11370
      Picture         =   "frmCounsel_1.frx":9DF8
      Top             =   5715
      Width           =   795
   End
   Begin VB.Image imgTab 
      Height          =   345
      Index           =   2
      Left            =   10230
      Picture         =   "frmCounsel_1.frx":ACB5
      Top             =   5250
      Width           =   885
   End
   Begin VB.Image imgTab 
      Height          =   345
      Index           =   1
      Left            =   9360
      Picture         =   "frmCounsel_1.frx":B360
      Top             =   5250
      Width           =   885
   End
   Begin VB.Image imgTab 
      Height          =   345
      Index           =   0
      Left            =   7800
      Picture         =   "frmCounsel_1.frx":BA6E
      Top             =   5250
      Width           =   1560
   End
   Begin VB.Image imgGoReserve 
      Height          =   330
      Left            =   11010
      Picture         =   "frmCounsel_1.frx":C899
      Top             =   1440
      Width           =   1185
   End
   Begin VB.Image imgAppend 
      Height          =   315
      Index           =   3
      Left            =   5940
      Top             =   2970
      Width           =   1365
   End
   Begin VB.Image imgAppend 
      Height          =   345
      Index           =   2
      Left            =   5940
      Top             =   2550
      Width           =   1365
   End
   Begin VB.Image imgAppend 
      Height          =   315
      Index           =   1
      Left            =   5940
      Top             =   2160
      Width           =   1365
   End
   Begin VB.Image imgAppend 
      Height          =   345
      Index           =   0
      Left            =   5940
      Top             =   1740
      Width           =   1365
   End
   Begin VB.Label lblMin 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "59 kg"
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   4620
      TabIndex        =   45
      Top             =   3930
      Width           =   795
   End
   Begin VB.Label lblMin 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "59 kg"
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   4620
      TabIndex        =   44
      Top             =   3630
      Width           =   795
   End
   Begin VB.Label lblMin 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "59 kg"
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   4620
      TabIndex        =   43
      Top             =   3330
      Width           =   795
   End
   Begin VB.Label lblMin 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "59 kg"
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   4620
      TabIndex        =   42
      Top             =   3030
      Width           =   795
   End
   Begin VB.Label lblMin 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "59 kg"
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   4620
      TabIndex        =   41
      Top             =   2730
      Width           =   795
   End
   Begin VB.Label lblMin 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "59 kg"
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   4620
      TabIndex        =   40
      Top             =   2430
      Width           =   795
   End
   Begin VB.Label lblMin 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "59 kg"
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   4620
      TabIndex        =   39
      Top             =   2130
      Width           =   795
   End
   Begin VB.Label lblMin 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "59 kg"
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   4620
      TabIndex        =   38
      Top             =   1800
      Width           =   795
   End
   Begin VB.Label lblMax 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "65 kg"
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   3720
      TabIndex        =   37
      Top             =   3930
      Width           =   825
   End
   Begin VB.Label lblMax 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "65 kg"
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   3720
      TabIndex        =   36
      Top             =   3630
      Width           =   825
   End
   Begin VB.Label lblMax 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "65 kg"
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   3720
      TabIndex        =   35
      Top             =   3330
      Width           =   825
   End
   Begin VB.Label lblMax 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "65 kg"
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   3720
      TabIndex        =   34
      Top             =   3030
      Width           =   825
   End
   Begin VB.Label lblMax 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "65 kg"
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   3720
      TabIndex        =   33
      Top             =   2730
      Width           =   825
   End
   Begin VB.Label lblMax 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "65 kg"
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   3720
      TabIndex        =   32
      Top             =   2430
      Width           =   825
   End
   Begin VB.Label lblMax 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "65 kg"
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3720
      TabIndex        =   31
      Top             =   2130
      Width           =   825
   End
   Begin VB.Label lblMax 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "65 kg"
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   3720
      TabIndex        =   30
      Top             =   1800
      Width           =   825
   End
   Begin VB.Label lblUpDown 
      BackStyle       =   0  '����
      Caption         =   "5 kg"
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   3030
      TabIndex        =   29
      Top             =   3930
      Width           =   585
   End
   Begin VB.Label lblUpDown 
      BackStyle       =   0  '����
      Caption         =   "5 kg"
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   3030
      TabIndex        =   28
      Top             =   3630
      Width           =   585
   End
   Begin VB.Label lblUpDown 
      BackStyle       =   0  '����
      Caption         =   "5 kg"
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   3030
      TabIndex        =   27
      Top             =   3330
      Width           =   585
   End
   Begin VB.Label lblUpDown 
      BackStyle       =   0  '����
      Caption         =   "5 kg"
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   3030
      TabIndex        =   26
      Top             =   3030
      Width           =   585
   End
   Begin VB.Label lblUpDown 
      BackStyle       =   0  '����
      Caption         =   "5 kg"
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   3030
      TabIndex        =   25
      Top             =   2730
      Width           =   585
   End
   Begin VB.Label lblUpDown 
      BackStyle       =   0  '����
      Caption         =   "5 kg"
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   3030
      TabIndex        =   24
      Top             =   2430
      Width           =   585
   End
   Begin VB.Label lblUpDown 
      BackStyle       =   0  '����
      Caption         =   "5 kg"
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3030
      TabIndex        =   23
      Top             =   2130
      Width           =   585
   End
   Begin VB.Label lblUpDown 
      BackStyle       =   0  '����
      Caption         =   "5 kg"
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   3030
      TabIndex        =   22
      Top             =   1800
      Width           =   585
   End
   Begin VB.Image imgUpDown 
      Height          =   105
      Index           =   7
      Left            =   2790
      Picture         =   "frmCounsel_1.frx":D3BD
      Top             =   3975
      Width           =   150
   End
   Begin VB.Image imgUpDown 
      Height          =   105
      Index           =   6
      Left            =   2790
      Picture         =   "frmCounsel_1.frx":D722
      Top             =   3675
      Width           =   150
   End
   Begin VB.Image imgUpDown 
      Height          =   105
      Index           =   5
      Left            =   2790
      Picture         =   "frmCounsel_1.frx":DA87
      Top             =   3375
      Width           =   150
   End
   Begin VB.Image imgUpDown 
      Height          =   105
      Index           =   4
      Left            =   2790
      Picture         =   "frmCounsel_1.frx":DDEC
      Top             =   3060
      Width           =   150
   End
   Begin VB.Image imgUpDown 
      Height          =   105
      Index           =   3
      Left            =   2790
      Picture         =   "frmCounsel_1.frx":E151
      Top             =   2745
      Width           =   150
   End
   Begin VB.Image imgUpDown 
      Height          =   105
      Index           =   2
      Left            =   2790
      Picture         =   "frmCounsel_1.frx":E4B6
      Top             =   2445
      Width           =   150
   End
   Begin VB.Image imgUpDown 
      Height          =   105
      Index           =   1
      Left            =   2790
      Picture         =   "frmCounsel_1.frx":E81B
      Top             =   2160
      Width           =   150
   End
   Begin VB.Image imgDietCal 
      Height          =   300
      Index           =   5
      Left            =   7800
      Picture         =   "frmCounsel_1.frx":EB80
      Top             =   3795
      Width           =   900
   End
   Begin VB.Image imgDietCal 
      Height          =   300
      Index           =   4
      Left            =   7800
      Picture         =   "frmCounsel_1.frx":F257
      Top             =   3480
      Width           =   900
   End
   Begin VB.Image imgDietCal 
      Height          =   300
      Index           =   3
      Left            =   7800
      Picture         =   "frmCounsel_1.frx":F8D8
      Top             =   3165
      Width           =   900
   End
   Begin VB.Image imgDietCal 
      Height          =   300
      Index           =   2
      Left            =   12450
      Picture         =   "frmCounsel_1.frx":FFA4
      Top             =   2850
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Image imgDietCal 
      Height          =   300
      Index           =   1
      Left            =   7800
      Picture         =   "frmCounsel_1.frx":1063A
      Top             =   2850
      Width           =   900
   End
   Begin VB.Image imgUpDown 
      Height          =   105
      Index           =   0
      Left            =   2790
      Picture         =   "frmCounsel_1.frx":10CCC
      Top             =   1830
      Width           =   150
   End
   Begin VB.Image imgDietCal 
      Height          =   300
      Index           =   0
      Left            =   12450
      Picture         =   "frmCounsel_1.frx":11031
      Top             =   2490
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Image imgViewHistory 
      Height          =   345
      Index           =   2
      Left            =   3840
      Picture         =   "frmCounsel_1.frx":116B9
      Top             =   4410
      Width           =   1575
   End
   Begin VB.Image imgViewHistory 
      Height          =   345
      Index           =   1
      Left            =   2250
      Picture         =   "frmCounsel_1.frx":12061
      Top             =   4410
      Width           =   1575
   End
   Begin VB.Image imgViewHistory 
      Height          =   345
      Index           =   0
      Left            =   690
      Picture         =   "frmCounsel_1.frx":12A98
      Top             =   4410
      Width           =   1575
   End
   Begin VB.Label lblObInfo 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "65 kg"
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   2040
      TabIndex        =   8
      Top             =   1800
      Width           =   600
   End
   Begin VB.Label lblObInfo 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "87 cm"
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   2040
      TabIndex        =   7
      Top             =   2430
      Width           =   600
   End
   Begin VB.Label lblObInfo 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "0.93"
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   2040
      TabIndex        =   6
      Top             =   2730
      Width           =   600
   End
   Begin VB.Label lblObInfo 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "120 %"
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   2040
      TabIndex        =   5
      Top             =   2130
      Width           =   600
   End
   Begin VB.Label lblObInfo 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "27"
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   2040
      TabIndex        =   4
      Top             =   3030
      Width           =   600
   End
   Begin VB.Label lblObInfo 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "1,280"
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   2010
      TabIndex        =   3
      Top             =   3630
      Width           =   600
   End
   Begin VB.Label lblObInfo 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "2,400"
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   2010
      TabIndex        =   2
      Top             =   3930
      Width           =   600
   End
End
Attribute VB_Name = "frmCounsel_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'���̾�Ʈ ���� ���ϱ� DietCal (O)
'*** �� �ε��
'1) ȯ�漳������ �ҷ��´�. ���� �Է��ؾ� �ϴ� ����(���߿��� �ʼ��Է°���)
'*** ����
'3) �����Ѵ�.(BodyData, BioChem, Treat, TreatData, TreatPrint
Option Explicit
Private Const IMG_ME As String = "\Back\Counsel\01\Me\"
Private Const IMG_EX As String = "\Back\Counsel\01\Ex\"
'+---------------------------------------------------------------------------------+
'| ��� > ó��/�񸸵���� > �Ļ簨��Į�θ�
'+---------------------------------------------------------------------------------+
Private Const IMG_M0Y As String = "\Back\Counsel\01\Me\125-green.jpg"
Private Const IMG_M0N As String = "\Back\Counsel\01\Me\125-gray.jpg"
Private Const IMG_M1Y As String = "\Back\Counsel\01\Me\250-green.jpg"
Private Const IMG_M1N As String = "\Back\Counsel\01\Me\250-gray.jpg"
Private Const IMG_M2Y As String = "\Back\Counsel\01\Me\375-green.jpg"
Private Const IMG_M2N As String = "\Back\Counsel\01\Me\375-gray.jpg"
Private Const IMG_M3Y As String = "\Back\Counsel\01\Me\500-green.jpg"
Private Const IMG_M3N As String = "\Back\Counsel\01\Me\500-gray.jpg"
Private Const IMG_M4Y As String = "\Back\Counsel\01\Me\750-green.jpg"
Private Const IMG_M4N As String = "\Back\Counsel\01\Me\750-gray.jpg"
Private Const IMG_M5Y As String = "\Back\Counsel\01\Me\1000-green.jpg"
Private Const IMG_M5N As String = "\Back\Counsel\01\Me\1000-gray.jpg"
'+---------------------------------------------------------------------------------+
'| ��� > ó��/�񸸵���� > ��ϼ�
'+---------------------------------------------------------------------------------+
Private Const IMG_EX3DY As String = "\Back\Counsel\01\Ex\3��-green.jpg"
Private Const IMG_EX3DN As String = "\Back\Counsel\01\Ex\3��-gray.jpg"
Private Const IMG_EX4DY As String = "\Back\Counsel\01\Ex\4��-green.jpg"
Private Const IMG_EX4DN As String = "\Back\Counsel\01\Ex\4��-gray.jpg"
Private Const IMG_EX5DY As String = "\Back\Counsel\01\Ex\5��-green.jpg"
Private Const IMG_EX5DN As String = "\Back\Counsel\01\Ex\5��-gray.jpg"
Private Const IMG_EX6DY As String = "\Back\Counsel\01\Ex\6��-green.jpg"
Private Const IMG_EX6DN As String = "\Back\Counsel\01\Ex\6��-gray.jpg"
'+---------------------------------------------------------------------------------+
'| ��� > ó��/�񸸵���� > ��Ҹ�Į�θ�
'+---------------------------------------------------------------------------------+
Private Const IMG_EX0Y As String = "\Back\Counsel\01\Ex\200-green.jpg"
Private Const IMG_EX0N As String = "\Back\Counsel\01\Ex\200-gray.jpg"
Private Const IMG_EX1Y As String = "\Back\Counsel\01\Ex\300-green.jpg"
Private Const IMG_EX1N As String = "\Back\Counsel\01\Ex\300-gray.jpg"
Private Const IMG_EX2Y As String = "\Back\Counsel\01\Ex\400-green.jpg"
Private Const IMG_EX2N As String = "\Back\Counsel\01\Ex\400-gray.jpg"
Private Const IMG_EX3Y As String = "\Back\Counsel\01\Ex\500-green.jpg"
Private Const IMG_EX3N As String = "\Back\Counsel\01\Ex\500-gray.jpg"
'+---------------------------------------------------------------------------------+
'| ��� > ó��/�񸸵���� > ����ȭ��ǥ(���ǽ�)
'+---------------------------------------------------------------------------------+
Private Const IMG_UP As String = "\Back\Counsel\01\icon-red.jpg"
Private Const IMG_DOWN As String = "\Back\Counsel\01\icon-blue.jpg"

Private Const IMG_RES_ON As String = "\Back\Counsel\01\���ຯ�� on.jpg"
Private Const IMG_RES_OFF As String = "\Back\Counsel\01\���ຯ�� off.jpg"
Private Const IMG_SELEX_ON As String = "\Back\Counsel\01\������� on.jpg"
Private Const IMG_SELEX_OFF As String = "\Back\Counsel\01\������� off.jpg"
Private Const IMG_TREAT_ON As String = "\Back\Counsel\01\ġ���̷º��� on.jpg"
Private Const IMG_TREAT_OFF As String = "\Back\Counsel\01\ġ���̷º��� off.jpg"
'+---------------------------------------------------------------------------------+
'| ��� > ó��/�񸸵���� > ��, ����, ����, ���
'+---------------------------------------------------------------------------------+
Private Const PATH01 As String = "\Back\Counsel\01\"
Private Const IMG_LTAB1_ON As String = "��ȭ��ǥ on.jpg"
Private Const IMG_LTAB1_OFF As String = "��ȭ��ǥ off.jpg"
Private Const IMG_LTAB2_ON As String = "��ȭ���׷��� on.jpg"
Private Const IMG_LTAB2_OFF As String = "��ȭ���׷��� off.jpg"
Private Const IMG_LTAB3_ON As String = "ó�泻����ȸ on.jpg"
Private Const IMG_LTAB3_OFF As String = "ó�泻����ȸ off.jpg"
Private Const IMG_RTAB1_ON As String = "ġ������ on.jpg"
Private Const IMG_RTAB1_OFF As String = "ġ������ off.jpg"
Private Const IMG_RTAB2_ON As String = "notice on.jpg"
Private Const IMG_RTAB2_OFF As String = "notice off.jpg"
Private Const IMG_RTAB3_ON As String = "memo on.jpg"
Private Const IMG_RTAB3_OFF As String = "memo off.jpg"
Private Const IMG_SAVE_ON As String = "save on.jpg"
Private Const IMG_SAVE_OFF As String = "save off.jpg"
Private Const IMG_DEL_ON As String = "delete on.jpg"
Private Const IMG_DEL_OFF As String = "delete off.jpg"
Private Const IMG_PRINT_ON As String = "�񸸵��� on.jpg"
Private Const IMG_PRINT_OFF As String = "�񸸵��� off.jpg"

Private Const IMG_NEW_ON As String = "new_on.jpg"
Private Const IMG_NEW_OFF As String = "new_off.jpg"


'+---------------------------------------------------------------------------------+
'| �� > 2Fon
'+---------------------------------------------------------------------------------+
Private Const IMG_��_ON As String = "\Img\��Ÿ\�� on.jpg"
Private Const IMG_��_OFF As String = "\Img\��Ÿ\�� off.jpg"

Private Type mCustomer
    strCustomName As String
    strJuminNum As String
    strSex As String
    intAge As Integer
    strChPgCode As String       'ġ�����α׷� �ڵ�
End Type
Private Type mTreatCal
    intDietCal As Integer
    intExCal As Integer
    intExDay As Integer
    sngTreatCal As Single
    sngLossWeight As Single
    intUserCal As Single
End Type
'ü��
'�񸸵�
'�㸮�ѷ�
'WHR
'BMI
'ü�����
'RMR
'TEE
Private Type mObesity
    sngWeight As Single
    sngObesityRate As Single
    sngWaist As Single
    sngWHR As Single
    sngBMI As Single
    sngChFatRate As Single
    sngRMR As Single
    sngTEE As Single
End Type
Private typCustomer As mCustomer
Private typTreatCal As mTreatCal
Private typObesity As mObesity

Public intMode As Integer             '�ű��Է¸��=1, ��ȸ/�������=2
Private glngBodyDataNum As Long       '��ü���� ����
Public glngTreatNum As Long           '���� ����
Private glngCompDataNum As Long       '���� ����

Private gintBottomButton As Integer
Private gintLeftButton As Integer
'�ó������ ����
Public gintMain As Integer    '������ �Ѱ�
Public gintSub1 As Integer    '������ �װ�
Public gintSub2 As Integer
Public gintSub3 As Integer
Public gintSub4 As Integer

Public intNowExCalory As Integer   '������ ���� �� ������ ó��� �Į�θ�

'��� ����
Dim crxApplication As New CRAXDRT.Application
Dim crxReport As CRAXDRT.Report
Dim crxDatabaseTable As CRAXDRT.DatabaseTable
Dim crxFormula As CRAXDRT.FormulaFieldDefinition
Dim strServer As String, strDBName As String, strUID As String, strPWD As String

'�Ļ簨��Į�θ�, ��ϼ�, ��Һ�Į�θ� Ŭ���� �ε��� ���
Dim idxDietCal As Integer, idxExDay As Integer, idxExCal As Integer


'2005-01-23 ������ ó����Ʈ���� �Է°���/�Ұ������� ����.
Private Sub EnabledInput(Optional bEnable As Boolean = True)
Dim ctrl As Control, i As Integer
    If bEnable Then
        lblDispContent = "�Է���"
        cmdInputTreat.Caption = "ó���Է�"
    Else
        lblDispContent = "��ȸ��"
        cmdInputTreat.Caption = "ó�����"
    End If
    cmbProgram.Enabled = bEnable
    imgGoReserve.Enabled = bEnable
    cmdPreTreat.Enabled = bEnable
    If bEnable Then
        If chkUserCalory.Value = vbChecked Then
            For Each ctrl In imgDietCal
                ctrl.Enabled = False
            Next
            txtUserCalory.Enabled = True
        Else
            For Each ctrl In imgDietCal
                ctrl.Enabled = True
            Next
            txtUserCalory.Enabled = False
        End If
    Else
        For Each ctrl In imgDietCal
            ctrl.Enabled = bEnable
        Next
        txtUserCalory.Enabled = bEnable
    End If
    For Each ctrl In imgExDay
        ctrl.Enabled = bEnable
    Next
    For Each ctrl In imgExCal
        ctrl.Enabled = bEnable
    Next
    chkUserCalory.Enabled = bEnable
    'imgSelectEx.Enabled = bEnable
    With sprTreat
        For i = 1 To .MaxRows
            .Row = i
            .Col = 2: .Lock = (Not bEnable)
            .Col = 3: .Lock = (Not bEnable)
        Next
    End With
    With sprPrint
        For i = 1 To .MaxRows
            .Row = i
            .Col = 2: .Lock = (Not bEnable)
        Next
    End With
    txtMemo.Enabled = bEnable
    txtNotice.Enabled = bEnable
    gbIsEnable = bEnable
End Sub

Private Sub chkUserCalory_Click()
    Dim i As Integer
    Dim sngMinus As Single
    
     AdminYn 11
    If AccessYn = False Then
        MsgBox "���ٱ����� �����ϴ�. �����ڿ��� �����Ͻʽÿ�.", vbOKOnly + vbInformation, "���ٱ��� ����"
        Exit Sub
    End If
    
    With typTreatCal
    If chkUserCalory.Value = 0 Then
        For i = 0 To 5
            imgDietCal(i).Enabled = True
        Next i
        lblTreatCal.Caption = Format(typObesity.sngTEE, "#,###") & " kcal"
        If .intDietCal >= 0 And .intExCal >= 0 And .intExDay >= 0 Then
            .sngLossWeight = ((.intDietCal * 7) + (.intExCal * .intExDay)) / 7700 * 4
            lblLossWeight.Caption = Format(.sngLossWeight, "0.0") & " kg/��"
        End If
        .intUserCal = 0
        txtUserCalory.BackColor = FRM_GRAY
        txtUserCalory.Enabled = False
    Else
        For i = 0 To 5
            Set imgDietCal(i).Picture = LoadPicture(App.Path & IMG_ME & i & "gray.jpg")
            imgDietCal(i).Enabled = False
        Next i
        .intDietCal = 0
        If txtUserCalory.Text <> "" And IsNumeric(txtUserCalory.Text) = True Then
            .intUserCal = txtUserCalory.Text
        Else
            .intUserCal = typObesity.sngTEE
        End If
        .sngTreatCal = .intUserCal
        lblTreatCal.Caption = Format(.sngTreatCal, "#,###") & " kcal"
    '+=========================================
    '+ ����ü�� ���ϴ� ��
    '+-----------------------------------------
        sngMinus = typObesity.sngTEE - .intUserCal
        If sngMinus >= 0 And .intExCal >= 0 And .intExDay >= 0 Then
            .sngLossWeight = ((sngMinus * 7) + (.intExCal * .intExDay)) / 7700 * 4
            lblLossWeight.Caption = Format(.sngLossWeight, "0.0") & " kg/��"
        End If
        txtUserCalory.BackColor = vbWhite
        txtUserCalory.Enabled = True
    End If
    End With
End Sub

Private Sub cmdInputTreat_Click()
    '���� frmBottom.dtpUserDay.Value�� �ش��ϴ� ó���� ���ٸ�
    '���� �Է��Ѵ�.
    '���� frmBottom.dtpUserDay.Value�� �ش��ϴ� ó���� �ִٸ�
    '���� �����͸� �����Ѵ�.
    '���� �����͸� �����Ҷ� �ش� ���ڿ� �������� ó���� ������ �����Ƿ�
    '�ش� ��¥�� ó���ȣ�� ������ �����͸� �����Ѵ�.
    
    If cmdInputTreat.Caption = "ó���Է�" Then
        lblDispDate.Caption = Format(gdatUserDay, "YYYY-MM-DD")
        intMode = 1
        Call EnabledInput(True)
    Else
        intMode = 2
        Call EnabledInput(True)
    End If
    cmdInputTreat.Visible = False
    Dim qrySelect As String
'    qrySelect = "SELECT TOP 1 * FROM Treat a LEFT JOIN "
'
'    Call EnabledInput(Not cmbProgram.Enabled)
End Sub

Private Sub cmdPreTreat_Click()
'���� �ֱ� ó���� ���̿���, �����, ��ϼ� �����ͼ� üũ
    Dim qrySelect As String, rValue As Variant
    Dim i As Integer, Index As Integer
    
    Set clsSelect = New clsSelect
    qrySelect = "SELECT TOP 1 TreatNum, DietCalory, ExDay, ExCalory, TreatCalory, UserCalory "
    qrySelect = qrySelect & "FROM Treat "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & " AND NOT TreatCalory IS NULL"
    qrySelect = qrySelect & " AND (NOT DietCalory IS NULL OR NOT ExDay IS NULL OR NOT ExCalory IS NULL "
    qrySelect = qrySelect & " OR NOT UserCalory IS NULL)"
    qrySelect = qrySelect & " ORDER BY TreatDay DESC;"
    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue) Then
        With typTreatCal
            .intDietCal = Is_Null(rValue(1, 0), 0)
            .intExDay = Is_Null(rValue(2, 0), 0)
            .intExCal = Is_Null(rValue(3, 0), 0)
            .sngTreatCal = Is_Null(rValue(4, 0), 0)
            lblTreatCal.Caption = Format(.sngTreatCal, "#,###") & " kcal"
            '���̾�Ʈ ����
            If Not IsNull(rValue(5, 0)) Then
                chkUserCalory.Value = 1
                .intUserCal = rValue(5, 0)
                txtUserCalory.Text = .intUserCal
                For i = 0 To 5
                    Set imgDietCal(i).Picture = LoadPicture(App.Path & IMG_ME & i & "gray.jpg")
                    imgDietCal(i).Enabled = False
                Next i
            Else
                chkUserCalory.Value = 0
                .intUserCal = 0
                txtUserCalory.Text = ""
                For i = 0 To 5
                    Set imgDietCal(i).Picture = LoadPicture(App.Path & IMG_ME & i & "gray.jpg")
                    imgDietCal(i).Enabled = True
                Next i
                Select Case .intDietCal
                    Case 125: Index = 0
                    Case 250: Index = 1
                    Case 375: Index = 2
                    Case 500: Index = 3
                    Case 750: Index = 4
                    Case 1000: Index = 5
                    Case Else: Index = 6
                End Select
                If Index <> 6 Then
                    Set imgDietCal(Index).Picture = LoadPicture(App.Path & IMG_ME & Index & "green.jpg")
                End If
            End If
            '��ϼ�
            For i = 0 To 3
                Set imgExDay(i).Picture = LoadPicture(App.Path & IMG_EX & i + 3 & "��-gray.jpg")
            Next i
            If .intExDay >= 3 And .intExDay <= 6 Then
                Set imgExDay(.intExDay - 3).Picture = LoadPicture(App.Path & IMG_EX & .intExDay & "��-green.jpg")
            End If
            '�Į�θ�
            For i = 0 To 3
                Set imgExCal(i).Picture = LoadPicture(App.Path & IMG_EX & i + 2 & "00-gray.jpg")
            Next i
            If .intExCal >= 200 And .intExCal <= 500 Then
                Set imgExCal((.intExCal / 100) - 2).Picture = LoadPicture(App.Path & IMG_EX & (.intExCal / 100) & "00-green.jpg")
            End If
            '����ü��
            ' ����ü���� ���� ü���� ����ؼ� �ٽ� ����Ѵ�.
'            lblLossWeight.Caption = Format(.sngLossWeight, "0.0") & " kg/��"
            If .intUserCal = 0 Then
                .sngTreatCal = typObesity.sngTEE - .intDietCal
                If .intDietCal >= 0 And .intExCal >= 0 And .intExDay >= 0 Then
                    .sngLossWeight = ((.intDietCal * 7) + (.intExCal * .intExDay)) / 7700 * 4
                    lblLossWeight.Caption = Format(.sngLossWeight, "0.0") & " kg/��"
                End If
            Else
                .sngTreatCal = typObesity.sngTEE - .intUserCal
                If .intUserCal >= 0 And .intExCal >= 0 And .intExDay >= 0 Then
                    .sngLossWeight = ((.intUserCal * 7) + (.intExCal * .intExDay)) / 7700 * 4
                    lblLossWeight.Caption = Format(.sngLossWeight, "0.0") & " kg/��"
                End If
            End If
        End With
    Else
        MsgBox "������ ó���� ������ �����ϴ�.", vbOKOnly + vbExclamation
    End If
    Set clsSelect = Nothing
End Sub

Public Sub Form_Load()
    Dim i As Integer
'1) ��ü����,Ķ����,��ȭ�м�ġ���� �Է��� �� �ֵ��� �������� ����
'2) ����ȯ���� ��� ����� ������ ������ óġ�� ��¹��� �̸� �����ְ�, �񸸵��� ���ؼ��� �����ش�
'3) ������ ����� �������� ǥ,�׷���,ó���� �� �� �ְ� �Ѵ�.
On Error GoTo ShowErr
'On Error Resume Next
'���� �ߴ� ��ġ �� �׷��� ==================================================
    Set Me.Picture = LoadPicture(App.Path & "\Back\Counsel\01\���_ó�� back.jpg")
    Me.Top = FRM_TOP
    Me.Left = FRM_LEFT
    Me.Width = FRM_WIDTH
    Me.Height = FRM_HEIGHT
    Me.BackColor = vbWhite
    Set imgGoReserve.Picture = LoadPicture(App.Path & IMG_RES_ON)
    Set imgSelectEx.Picture = LoadPicture(App.Path & IMG_SELEX_ON)
' frame1,2,3
' frame4,5,6
'///////////////////////////////////////////////////////////////////////////
    If LoadCustomerInfo = False Then
'        Exit Sub
    End If

    gintBottomButton = 0
    gintLeftButton = 0
    glngBodyDataNum = 0
    glngTreatNum = 0
    glngCompDataNum = 0
    
    gintMain = 0
    gintSub1 = 0: gintSub2 = 0: gintSub3 = 0: gintSub4 = 0

    '�񸸰��� �׸� �ʱ�ȭ/ġ�����α׷� �ʱ�ȭ �� ���� ȭ�� ��Ʈ�� �ʱ�ȭ
    '----------------------------------------------------------------------
    '2005-01-27 ������
    '�������� ��ȸ����� ǥ���ϴ� ��Ʈ�� �ʱ�ȭ
    'ġ�����α׷� �޺��ڽ�, ������Է� ó��Į�α�, Notice/Memo��Ʈ�� �ʱ�ȭ
    '----------------------------------------------------------------------
    Call InitialControl
    'ó���Է��ϴ� ��������(ġ������/��¹�) �ʱ�ȭ
    '======================
    Call InitialSpread4 'ġ������
    Call InitialSpread5 '��¹�
    '======================
    
    '��ȭ��(ǥ) ���� ���������� ��.(���� �Ʒ�)
    Call imgViewHistory_Click(0)

    'imgViewHistory_Click(0)���� Call Bottom0(0)�� �θ� Ȥ�� �����ص� �ɼ��� ���� Ȯ���� ����.
    '2004-12-09 ������ imgViewHistory_Click(0)���� Bottom0(0)�� ȣ����.
    'Call Bottom0(0)
    
    'ġ������/��¹�(������ �ϴ�) ���� ���������� ��.
    Call imgTab_Click(0)

'2005-01-27 ������
'==========> ������� ��Ʈ�� �ʱ�ȭ/���⼭���� ������ �����ֱ�

    '���� �ֱ� �Է°����� �����ֱ�
    
    '���� ��� �񸸰��� ������ (ü��/�񸸵�/�㸮�ѷ�/WHR/BMI/ü�����/RMR/TEE)�����ֱ�
    Call ShowObesityRecord
    
    
    '������ ��� Į�θ�ó�� ������ �ʱ�ȭ
    Call InitialTreatCalory
    intMode = 1
    cmdPreTreat.Visible = True
    '���������� Treat�� ������ �ش� ���ᵥ���� �ҷ�����
    Call ShowTodayTreat
    
    '���̾�Ʈ�� �ϵ��� �ϴ� ó��Į�θ� ��Ʈ�� �ʱ�ȭ
    Call InitialDietCalory
    
    '���� �ֱ� ó���� ����� ��������
    Call GetLatestExItem
    
    Exit Sub
ShowErr:
    '2004-12-23 ������ �αױ��
    'WriteLog "Form_Load", "frmCounsel_1", Err.Number, Err.Description
    MsgBox Err.Number & Err.Description
End Sub

'���ó�¥�� ó���� ������ �����ֱ�
Private Sub ShowTodayTreat()
    Dim qrySelect As String, rValue As Variant
    
    qrySelect = "SELECT TreatNum, DietCalory, TreatCalory, UserCalory FROM Treat WHERE CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & " AND TreatDay='" & Format(gdatUserDay, "YYYYMMDD") & "'"
    qrySelect = qrySelect & " ORDER BY TreatNum DESC "
    Set clsSelect = New clsSelect
    rValue = clsSelect.Query(qrySelect)
    Set clsSelect = Nothing
    If Not IsNull(rValue) Then
        glngTreatNum = CLng(rValue(0, 0))
        intMode = 2
        If (IsNull(rValue(1, 0)) Or IsNull(rValue(2, 0))) And (IsNull(rValue(2, 0)) Or IsNull(rValue(3, 0))) Then
            cmdPreTreat.Visible = True
        Else
            cmdPreTreat.Visible = False
        End If
        
        '2004-12-09 ������ �̳�� �� �̸� �Դٸ� ���ٸ� �ϸ鼭 �����͸� �����ִ°���? �Ѥ�;
        '==========
        '������ ���(?) ġ�᳻�� �����ֱ�
        Call ShowTreat(glngTreatNum)
        '�������ϴ� ġ������ ����� ������ �����ֱ�
        Call ShowTreatData(glngTreatNum)
        '�������ϴ� ��¹��� ����� ������ �����ֱ�
        Call ShowTreatPrint(glngTreatNum)
        '==========
        
        Call EnabledInput(False)
        lblDispDate.Caption = Format(gdatUserDay, "YYYY-MM-DD")
        lblDispContent.Caption = "��ȸ��"
        cmdInputTreat.Visible = True
    Else
        Call imgNew_MouseUp(0, 0, 0, 0)
    End If
    
End Sub

Private Sub GetLatestExItem()
'���� �ֱ��� ó��� �����(������1, ������4) ��������
    Dim qrySelect As String, rValue As Variant
    
    qrySelect = "SELECT TOP 1 TreatNum, main, sub1, sub2, sub3, sub4 "
    qrySelect = qrySelect & "FROM Treat "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & " AND NOT main IS NULL"
    qrySelect = qrySelect & " ORDER BY TreatDay DESC;"
    Set clsSelect = New clsSelect
    rValue = clsSelect.Query(qrySelect)
    Set clsSelect = Nothing
    If Not IsNull(rValue) Then
        gintMain = Is_Null(rValue(1, 0), 0)
        gintSub1 = Is_Null(rValue(2, 0), 0)
        gintSub2 = Is_Null(rValue(3, 0), 0)
        gintSub3 = Is_Null(rValue(4, 0), 0)
        gintSub4 = Is_Null(rValue(5, 0), 0)
    Else
        gintMain = 0
        gintSub1 = 0
        gintSub2 = 0
        gintSub3 = 0
        gintSub4 = 0
    End If
End Sub

Private Sub cmbProgram_Change()
    Dim qrySelect As String, rValue As Variant

    Set clsSelect = New clsSelect
    AdminYn 10
    If AccessYn = False Then
        Exit Sub
    End If

    qrySelect = "SELECT ChPgCode FROM ChPgUpdate WHERE ChPgName='" & Trim(cmbProgram.Text) & "';"
    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue) Then
        typCustomer.strChPgCode = Trim(rValue(0, 0))
    Else
        typCustomer.strChPgCode = 0
    End If
    Set clsSelect = Nothing

    '�ش� ġ�����α׷���� �����Ѵ�.
End Sub

Private Sub cmbProgram_Click()
    AdminYn 10
    If AccessYn = False Then
'        MsgBox "���ٱ����� �����ϴ�. �����ڿ��� �����Ͻʽÿ�.", vbOKOnly + vbInformation, "���ٱ��� ����"
        Exit Sub
    End If
    Call cmbProgram_Change
End Sub

'��ȭ��(ǥ),��ȭ��(�׷���),ó�泻����ȸ ���� �����ư�� Ŭ���̺�Ʈ
Private Sub cmdSub_Click(Index As Integer)
    gintLeftButton = Index

    Select Case gintBottomButton
        Case 0  '��ȭ��(ǥ)
            Call Bottom0(gintLeftButton)
        Case 1  '��ȭ��(�׷���)
            Call Bottom1(gintLeftButton)
        Case 2  'ó�泻����ȸ
            Call Bottom2(gintLeftButton)
    End Select
End Sub

'��ȭ��(ǥ) ���� �����ư���� Ŭ���̺�Ʈ
Private Sub Bottom0(intIndex As Integer)
'0~2
    Dim i As Integer, j As Integer
    Dim qrySelect As String, rValue As Variant
    
On Error GoTo Err
    grdTable.Visible = True
    Chart.Visible = False
    'ü����, ��ü�ѷ�, �Ǻεβ�, ��ȭ�а˻��ư�� ���̰� ������ ����
    '=============
    For i = 0 To 3
        cmdSub(i).Visible = True
    Next i
    For i = 4 To 9
        cmdSub(i).Visible = False
    Next i
    
    cmdSub(0).Caption = "ü   ��   ��"
    cmdSub(1).Caption = "�� ü �� ��"
    cmdSub(2).Caption = "�� �� �� ��"
    cmdSub(3).Caption = "��ȭ�� �˻�"
    '=============
    
    Select Case intIndex
        Case 0   '����,ü��,VO2,RMR+ü����
            Set clsSelect = New clsSelect
            qrySelect = "SELECT InputName, InputField "
            qrySelect = qrySelect & "FROM C_InputData WHERE InputKind='B' "
            qrySelect = qrySelect & "ORDER BY InputOrder ASC;"
            
            rValue = clsSelect.Query(qrySelect)
            Set clsSelect = Nothing
            
            If Not IsNull(rValue) Then
                With grdTable
                    .Clear
                    .Rows = 3
                    .RowHeight(-1) = 300
                    .RowHeight(1) = 0
                    .Cols = UBound(rValue, 2) + 5 + 4 '+1 �������
                    .SelectionMode = flexSelectionByRow
                    
                    .TextMatrix(0, 0) = ""          '����
                    .ColWidth(0) = 300
                    .TextMatrix(0, 1) = "������"
                    .ColWidth(1) = 1000
                    For i = 2 To .Cols - 1
                        .ColWidth(i) = 1000
                    Next i
                    'InputOrder�� ������� �����ش�.
                    '��ü����, Ȱ��������
                    .TextMatrix(0, 2) = "����"
                    .TextMatrix(1, 2) = "Height"
                    .TextMatrix(0, 3) = "ü��"
                    .TextMatrix(1, 3) = "Weight"
                    .TextMatrix(0, 4) = "VO2"
                    .TextMatrix(1, 4) = "VO2"
                    .TextMatrix(0, 5) = "����RMR"
                    .TextMatrix(1, 5) = "inRMR"
                    .TextMatrix(0, 6) = "����RMR"
                    .TextMatrix(1, 6) = "RMR"
                    For i = 0 To UBound(rValue, 2)
                        .TextMatrix(0, i + 7) = Trim(rValue(0, i))
                        .TextMatrix(1, i + 7) = Trim(rValue(1, i))
                    Next i
                    .TextMatrix(1, i + 7) = "Treat.TreatNum"
                    .ColAlignment(-1) = flexAlignCenterCenter
                    .ColWidth(i + 7) = 0
                End With
                Erase rValue
                '�����̸� �׳� �Ѿ���� ������ ���
                Call ShowTable_Measure
            End If
        Case 1    '��ü�ѷ�
            Set clsSelect = New clsSelect
            qrySelect = "SELECT InputName, InputField "
            qrySelect = qrySelect & "FROM C_InputData WHERE InputKind='C' "
            qrySelect = qrySelect & "ORDER BY InputOrder ASC;"
            
            rValue = clsSelect.Query(qrySelect)
            Set clsSelect = Nothing
            
            If Not IsNull(rValue) Then
                With grdTable
                    .Clear
                    .Rows = 3
                    .RowHeight(-1) = 300
                    .RowHeight(1) = 0
                    .Cols = UBound(rValue, 2) + 4                     '+1 �������
                    .SelectionMode = flexSelectionByRow
                    .TextMatrix(0, 0) = ""          '����
                    .ColWidth(0) = 300
                    .TextMatrix(0, 1) = "������"
                    .ColWidth(1) = 1000
                    For i = 2 To .Cols - 1
                        .ColWidth(i) = 1290
                    Next i
                    For i = 0 To UBound(rValue, 2)
                        .TextMatrix(0, i + 2) = Trim(rValue(0, i))
                        .TextMatrix(1, i + 2) = Trim(rValue(1, i))
                    Next i
                    .TextMatrix(1, i + 2) = "Treat.TreatNum"
                    .ColAlignment(-1) = flexAlignCenterCenter
                    .ColWidth(i + 2) = 0
                End With
                Erase rValue
                Call ShowTable_Circumference
            End If
        Case 2     '�Ǻεβ�
            Set clsSelect = New clsSelect
            qrySelect = "SELECT InputName, InputField "
            qrySelect = qrySelect & "FROM C_InputData WHERE InputKind='S' "
            qrySelect = qrySelect & "ORDER BY InputOrder ASC;"
            
            rValue = clsSelect.Query(qrySelect)
            Set clsSelect = Nothing
            
            If Not IsNull(rValue) Then
                With grdTable
                    .Clear
                    .Rows = 3
                    .RowHeight(-1) = 300
                    .RowHeight(1) = 0
                    .Cols = UBound(rValue, 2) + 4                     '+1 �������
                    .SelectionMode = flexSelectionByRow
                    .TextMatrix(0, 0) = ""          '����
                    .ColWidth(0) = 300
                    .TextMatrix(0, 1) = "������"
                    .ColWidth(1) = 1000
                    For i = 2 To .Cols - 1
                        .ColWidth(i) = 1290
                    Next i
                    For i = 0 To UBound(rValue, 2)
                        .TextMatrix(0, i + 2) = Trim(rValue(0, i))
                        .TextMatrix(1, i + 2) = Trim(rValue(1, i))
                    Next i
                    .TextMatrix(1, i + 2) = "Treat.TreatNum"
                    .ColAlignment(-1) = flexAlignCenterCenter
                    .ColWidth(i + 2) = 0
                End With
                Erase rValue
'                Call ShowTable_Caliper
                Call ShowTable_Circumference
            End If
        Case 3     '��ȭ�а˻�
            Set clsSelect = New clsSelect
            qrySelect = "SELECT InputName, InputField "
            qrySelect = qrySelect & "FROM C_InputData WHERE InputKind='L' "
            qrySelect = qrySelect & "ORDER BY InputOrder ASC;"
            
            rValue = clsSelect.Query(qrySelect)
            Set clsSelect = Nothing
            
            If Not IsNull(rValue) Then
                With grdTable
                    .Clear
                    .Rows = 3
                    .RowHeight(-1) = 300
                    .RowHeight(1) = 0
                    .Cols = UBound(rValue, 2) + 4                     '+1 �������
                    .SelectionMode = flexSelectionByRow
                    .TextMatrix(0, 0) = ""          '����
                    .ColWidth(0) = 300
                    .TextMatrix(0, 1) = "������"
                    .ColWidth(1) = 1000
                    For i = 2 To .Cols - 1
'                        .ColWidth(i) = 1000
                        .ColWidth(i) = 1290
                    Next i
                    For i = 0 To UBound(rValue, 2)
                        .TextMatrix(0, i + 2) = Trim(rValue(0, i))
                        .TextMatrix(1, i + 2) = Trim(rValue(1, i))
                    Next i
                    .TextMatrix(1, i + 2) = "Treat.TreatNum"
                    .ColAlignment(-1) = flexAlignCenterCenter
                    .ColWidth(i + 2) = 0
                End With
                Erase rValue
                Call ShowTable_BioChem
            End If
    End Select
    Exit Sub
Err:
    '2004-12-23 ������ �αױ��
    'WriteLog "Bottom0", "frmCounsel_1", Err.Number, Err.Description
    MsgBox Err.Number & Err.Description
End Sub

'���� �ϴ��� ��ȭ��(ǥ)�� �������忡 ü���е��� �Է��Ѵ�.
Private Sub ShowTable_Measure()
    Dim qrySelect As String, rValue As Variant
    Dim i As Integer, j As Integer

    Set clsSelect = New clsSelect

    With grdTable
    qrySelect = "SELECT TreatDay"
    For i = 1 To .Cols - 2
        qrySelect = qrySelect & "," & Trim(.TextMatrix(1, i + 1))
    Next i
    qrySelect = qrySelect & " FROM BodyData INNER JOIN Treat ON "
    qrySelect = qrySelect & "BodyData.TreatNum=Treat.TreatNum "
    qrySelect = qrySelect & "WHERE BodyData.CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & " ORDER BY TreatDay ASC;"

    rValue = clsSelect.Query(qrySelect)

    If Not IsNull(rValue) Then
        .Rows = .Rows + UBound(rValue, 2)
        For i = 0 To UBound(rValue, 2)
            .TextMatrix(i + 2, 0) = i + 1
            For j = 0 To .Cols - 2
                If Not IsNull(rValue(j, i)) Then
                    .TextMatrix(i + 2, j + 1) = rValue(j, i)
                Else
                    .TextMatrix(i + 2, j + 1) = "-"
                End If
            Next j
' 2004-12-09 ������ ó�泯¥�� �Ʒ� �۾���¥�� ������ �������� ���� ������ ���̴µ� �̷�ƾ�� ����� Ż���� ����.
' Ȥ�� Ÿ�� �Ѵٸ� �׶� �ٽ� �ѹ� �Ѥ�; �Ʒ� ���ǹ����� ��ġ

'=========================================================
' 2005-01-26 ������ ���⼭ �� �������� ���� ����????
' �������� ������ư�� Ŭ��������츸...
'���� ��¥ �����͸� �����ֱ� ���ؼ��ϴ� ���ΰ�?
'=========================================================
            
            If Trim(rValue(0, i)) = Format(gdatUserDay, "yyyy-MM-dd") Then
'            If rValue(0, i) = gdatUserDay Then
                .Row = i + 2: .Col = 0: .ColSel = .Cols - 1
                glngTreatNum = .TextMatrix(i + 2, .Cols - 1)
                '��ȸ,���� ���
               ' intMode = 2

                Call ShowTreat(glngTreatNum)
                Call ShowTreatData(glngTreatNum)
                Call ShowTreatPrint(glngTreatNum)
            End If
        Next i
    End If
    End With

    Set clsSelect = Nothing
End Sub

Private Sub ShowTable_BioChem()
    Dim qrySelect As String, rValue As Variant
    Dim qrySelect1 As String
    Dim i As Integer, j As Integer, k As Integer
    Dim intCount As Integer
    
    Set clsSelect = New clsSelect
    
    intCount = grdTable.Cols - 3
    
    qrySelect = "SELECT TreatDay, "
    For i = 1 To intCount - 1
        qrySelect = qrySelect & "a" & i & ","
    Next i
    qrySelect = qrySelect & "a" & i & " FROM Treat INNER JOIN ("
    qrySelect = qrySelect & "SELECT TreatNum, "
    For i = 1 To intCount - 1
        qrySelect = qrySelect & "SUM(t" & i & ") AS a" & grdTable.TextMatrix(1, i + 1) & ","
    Next i
    qrySelect = qrySelect & "SUM(t" & i & ") AS a" & grdTable.TextMatrix(1, i + 1) & " FROM ("
    For i = 1 To intCount
        qrySelect1 = "SELECT Treat.TreatNum, "
        If i > 0 Then
            For j = 0 To i - 1
                qrySelect1 = qrySelect1 & "0 AS t" & j & ","
            Next j
        End If
        If i <= intCount Then
            qrySelect1 = qrySelect1 & "CASE BioChemCode "
            qrySelect1 = qrySelect1 & "WHEN " & grdTable.TextMatrix(1, i + 1) & " THEN BioChemSu ELSE 0 END AS t" & i & ","
            For k = i + 1 To intCount
                qrySelect1 = qrySelect1 & "0 AS t" & k & ","
            Next k
        End If
        qrySelect = qrySelect & Left(qrySelect1, Len(qrySelect1) - 1)
        qrySelect = qrySelect & " FROM BodyData INNER JOIN Treat ON BodyData.TreatNum=Treat.TreatNum "
        qrySelect = qrySelect & "LEFT JOIN BioChemData ON Treat.TreatNum=BioChemData.TreatNum "
        qrySelect = qrySelect & "WHERE Treat.CustomerNum=" & glngCustomerNum
        qrySelect = qrySelect & " GROUP BY Treat.TreatNum, BioChemCode, BioChemSu "
        If i < intCount Then
            qrySelect = qrySelect & "UNION ALL "
        End If
    Next i
    qrySelect = qrySelect & ") a GROUP BY TreatNum) b ON Treat.TreatNum=b.TreatNum"
    qrySelect = qrySelect & " ORDER BY TreatDay"
    
    rValue = clsSelect.Query(qrySelect)
    
    If Not IsNull(rValue) Then
    With grdTable
        .Rows = UBound(rValue, 2) + 3
        For i = 0 To UBound(rValue, 2)
            .TextMatrix(i + 2, 0) = i + 1
            For j = 0 To intCount
                If rValue(j, i) = 0 Then
                    .TextMatrix(i + 2, j + 1) = "-"
                Else
                    .TextMatrix(i + 2, j + 1) = rValue(j, i)
                End If
            Next j
        Next i
        .ColAlignment(-1) = flexAlignCenterCenter
    End With
    End If
    
    Set clsSelect = Nothing
End Sub

Private Sub ShowTable_TreatCode()
    Dim qrySelect As String, rValue As Variant
    Dim qrySelect1 As String
    Dim i As Integer, j As Integer, k As Integer
    Dim intCount As Integer

    Set clsSelect = New clsSelect
    
    '���� ���� ������ ó���� �ϳ��� ���� ���(TreatCode�� ���� ���)���� �׳� �Ѿ��
    If grdTable.TextMatrix(1, 1) = "" Then
        Exit Sub
    End If
    
'    '�ϴ� ���� ����Ǿ� �ִ� TreatCode�� �� �ҷ��ø���.
    intCount = grdTable.Cols - 3

'2005-01-23 ������ ���� �̸� �����ұ�? �Ѥ�;
'    qrySelect = "SELECT TreatNum, TreatDay, "
'    For i = 0 To intCount - 1
'        qrySelect = qrySelect & "SUM(t" & i & ") AS a" & grdTable.TextMatrix(1, i + 1) & ","
'    Next i
'    qrySelect = qrySelect & "SUM(t" & i & ") AS a" & grdTable.TextMatrix(1, i + 1) & " FROM ("
'    For i = 0 To intCount
'        qrySelect1 = "SELECT Treat.TreatNum, TreatDay, "
'        If i > 0 Then
'            For j = 0 To i - 1
'                qrySelect1 = qrySelect1 & "0 AS t" & j & ","
'            Next j
'        End If
'        If i <= intCount Then
'            qrySelect1 = qrySelect1 & "CASE TreatCode "
'            qrySelect1 = qrySelect1 & "WHEN " & grdTable.TextMatrix(1, i + 1) & " THEN 1 ELSE 0 END AS t" & i & ","
'            For k = i + 1 To intCount
'                qrySelect1 = qrySelect1 & "0 AS t" & k & ","
'            Next k
'        End If
'        qrySelect = qrySelect & Left(qrySelect1, Len(qrySelect1) - 1)
'        qrySelect = qrySelect & " FROM Treat LEFT JOIN TreatData "
'        qrySelect = qrySelect & "ON Treat.TreatNum=TreatData.TreatNum "
'        qrySelect = qrySelect & "WHERE Treat.CustomerNum=" & glngCustomerNum
'        qrySelect = qrySelect & " GROUP BY Treat.TreatNum, TreatDay, TreatCode "
'        If i < intCount Then
'            qrySelect = qrySelect & "UNION ALL "
'        End If
'    Next i
'    qrySelect = qrySelect & ") a GROUP BY TreatNum, TreatDay"

    
    
    qrySelect = "SELECT TreatNum, TreatDay, "
    For i = 0 To intCount - 1
        qrySelect = qrySelect & "SUM(t" & i & ") AS a" & grdTable.TextMatrix(1, i + 1) & ","
    Next i
    qrySelect = qrySelect & "SUM(t" & i & ") AS a" & grdTable.TextMatrix(1, i + 1) & " FROM ("
        
        
    '��������
    qrySelect1 = "SELECT Treat.TreatNum, TreatDay, "
    For i = 0 To intCount
        qrySelect1 = qrySelect1 & " CASE TreatCode WHEN " & grdTable.TextMatrix(1, i + 1) & " THEN 1 ELSE 0 END AS t" & i & ","
    Next i
    'qrySelect = qrySelect & "SELECT Treat.TreatNum, TreatDay, "
    qrySelect = qrySelect & Left(qrySelect1, Len(qrySelect1) - 1)
    qrySelect = qrySelect & " FROM Treat LEFT JOIN TreatData "
    qrySelect = qrySelect & "ON Treat.TreatNum=TreatData.TreatNum "
    qrySelect = qrySelect & "WHERE Treat.CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & "AND TreatCalory IS NOT NULL "
    qrySelect = qrySelect & ") a GROUP BY TreatNum, TreatDay"
    qrySelect = qrySelect & " ORDER BY TreatDay "
    
    rValue = clsSelect.Query(qrySelect)

    If Not IsNull(rValue) Then
    With grdTable
        .Rows = UBound(rValue, 2) + 3
        For i = 0 To UBound(rValue, 2)
            .TextMatrix(i + 2, 0) = rValue(1, i)
            For j = 1 To intCount + 1
                If CInt(rValue(j + 1, i)) = 1 Then
                    .TextMatrix(i + 2, j) = "O"
                Else
                    .TextMatrix(i + 2, j) = "X"
                End If
            Next j
            .TextMatrix(i + 2, intCount + 2) = rValue(0, i)
        Next i
        .ColAlignment(-1) = flexAlignCenterCenter
    End With
    End If

    Set clsSelect = Nothing
End Sub

Private Sub ShowTable_Print()
    Dim qrySelect As String, rValue As Variant
    Dim qrySelect1 As String
    Dim i As Integer, j As Integer, k As Integer
    Dim intCount As Integer

    Set clsSelect = New clsSelect
    '���� ���� ������ ó���� �ϳ��� ���� ��쿡�� �׳� �Ѿ��
    If grdTable.TextMatrix(1, 1) = "" Then
        Exit Sub
    End If
    
    intCount = grdTable.Cols - 3

'    qrySelect = "SELECT TreatNum, TreatDay, "
'    For i = 0 To intCount - 1
'        qrySelect = qrySelect & "SUM(t" & i & ") AS a" & grdTable.TextMatrix(1, i + 1) & ","
'    Next i
'    qrySelect = qrySelect & "SUM(t" & i & ") AS a" & grdTable.TextMatrix(1, i + 1) & " FROM ("
'    For i = 0 To intCount
'        qrySelect1 = "SELECT Treat.TreatNum, TreatDay, "
'        If i > 0 Then
'            For j = 0 To i - 1
'                qrySelect1 = qrySelect1 & "0 AS t" & j & ","
'            Next j
'        End If
'        If i <= intCount Then
'            qrySelect1 = qrySelect1 & "CASE PrintoutNum "
'            qrySelect1 = qrySelect1 & "WHEN " & grdTable.TextMatrix(1, i + 1) & " THEN 1 ELSE 0 END AS t" & i & ","
'            For k = i + 1 To intCount
'                qrySelect1 = qrySelect1 & "0 AS t" & k & ","
'            Next k
'        End If
'        qrySelect = qrySelect & Left(qrySelect1, Len(qrySelect1) - 1)
'        qrySelect = qrySelect & " FROM Treat LEFT JOIN TreatPrint "
'        qrySelect = qrySelect & "ON Treat.TreatNum=TreatPrint.TreatNum "
'        qrySelect = qrySelect & "WHERE Treat.CustomerNum=" & glngCustomerNum
'        qrySelect = qrySelect & " GROUP BY Treat.TreatNum, TreatDay, PrintoutNum "
'        If i < intCount Then
'            qrySelect = qrySelect & "UNION ALL "
'        End If
'    Next i
'    qrySelect = qrySelect & ") a GROUP BY TreatNum, TreatDay"



    qrySelect = "SELECT TreatNum, TreatDay, "
    For i = 0 To intCount - 1
        qrySelect = qrySelect & "SUM(t" & i & ") AS a" & grdTable.TextMatrix(1, i + 1) & ","
    Next i
    qrySelect = qrySelect & "SUM(t" & i & ") AS a" & grdTable.TextMatrix(1, i + 1) & " FROM ("
        
        
    '��������
    qrySelect1 = "SELECT Treat.TreatNum, TreatDay, "
    For i = 0 To intCount
        qrySelect1 = qrySelect1 & " CASE PrintoutNum WHEN " & grdTable.TextMatrix(1, i + 1) & " THEN 1 ELSE 0 END AS t" & i & ","
    Next i
    'qrySelect = qrySelect & "SELECT Treat.TreatNum, TreatDay, "
    qrySelect = qrySelect & Left(qrySelect1, Len(qrySelect1) - 1)
    qrySelect = qrySelect & " FROM Treat LEFT JOIN TreatPrint "
    qrySelect = qrySelect & "ON Treat.TreatNum=TreatPrint.TreatNum "
    qrySelect = qrySelect & "WHERE Treat.CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & "AND TreatCalory IS NOT NULL "
    qrySelect = qrySelect & ") a GROUP BY TreatNum, TreatDay"
    qrySelect = qrySelect & " ORDER BY TreatDay "


    rValue = clsSelect.Query(qrySelect)

    If Not IsNull(rValue) Then
    With grdTable
        .Rows = UBound(rValue, 2) + 3
        For i = 0 To UBound(rValue, 2)
            .TextMatrix(i + 2, 0) = rValue(1, i)
            For j = 1 To intCount + 1
                If CInt(rValue(j + 1, i)) = 1 Then
                    .TextMatrix(i + 2, j) = "O"
                Else
                    .TextMatrix(i + 2, j) = "X"
                End If
            Next j
            .TextMatrix(i + 2, intCount + 2) = rValue(0, i)
        Next i
        .ColAlignment(-1) = flexAlignCenterCenter
    End With
    End If

    Set clsSelect = Nothing
End Sub

'��ȭ��(�׷���) ���� �����ư���� Ŭ���̺�Ʈ
Private Sub Bottom1(intIndex As Integer)
    Dim i As Integer
    Dim qrySelect As String, rValue As Variant
    Dim sngMin As Single, sngMax As Single
    Dim strTitle As String

On Error GoTo Err
    grdTable.Visible = False
    For i = 0 To 9
        cmdSub(i).Visible = True
    Next i

    ' *** ��ȭ�� �׷��� �׸�
    ' ü��, �㸮�ѷ�, WHR, RMR, ü�����, ������, Į�θ�(�Ļ��ϱ�)
    cmdSub(0).Caption = "ü      ��"
    cmdSub(1).Caption = "�㸮�ѷ�"
    cmdSub(2).Caption = "��ϵѷ�"
    cmdSub(3).Caption = "W  H  R"
    cmdSub(4).Caption = "V  O  2"
    cmdSub(5).Caption = "R  M  R"
    cmdSub(6).Caption = "ü�����"
    cmdSub(7).Caption = "�� �� ��"
    cmdSub(8).Caption = "Į�θ�(�Ļ�)"
    cmdSub(9).Caption = "Į�θ�(�)"

    Call InitialChart
    ' *** ��ȭ�� �׷��� �׸�
    ' ü��, �㸮�ѷ�, WHR, ü�����, ������, �޽Ĵ�緮, Į�θ�(�Ļ��ϱ�), Į�θ�(��ϱ�)
    ' --> ü��, �㸮�ѷ�, WHR, RMR, ü�����, ������, Į�θ�(�Ļ��ϱ�)
    Dim sngStep As Single
    Select Case intIndex
        Case 0   'ü��
            qrySelect = "SELECT TreatDay, Weight FROM BodyData "
            qrySelect = qrySelect & "INNER JOIN Treat ON BodyData.TreatNum=Treat.TreatNum "
            qrySelect = qrySelect & "WHERE BodyData.CustomerNum=" & glngCustomerNum
            qrySelect = qrySelect & " AND Weight > 0 ORDER BY TreatDay ASC;"
            sngMin = MinValue("Weight") - 1
            sngMax = MaxValue("Weight") + 1
            sngStep = 2
            strTitle = "ü�� ��ȭ"
            Chart.Axis(AXIS_Y).Decimals = 1
        Case 1   '�㸮�ѷ�
            qrySelect = "SELECT TreatDay, Waist FROM BodyData "
            qrySelect = qrySelect & "INNER JOIN Treat ON BodyData.TreatNum=Treat.TreatNum "
            qrySelect = qrySelect & "WHERE BodyData.CustomerNum=" & glngCustomerNum
            qrySelect = qrySelect & " AND Waist IS NOT NULL ORDER BY TreatDay ASC;"
            sngMin = MinValue("Waist") - 5
            sngMax = MaxValue("Waist") + 5
            sngStep = 5
            strTitle = "�㸮�ѷ� ��ȭ"
            Chart.Axis(AXIS_Y).Decimals = 1
        Case 2   '��ϵѷ�
            qrySelect = "SELECT TreatDay, UpperArm FROM BodyData "
            qrySelect = qrySelect & "INNER JOIN Treat ON BodyData.TreatNum=Treat.TreatNum "
            qrySelect = qrySelect & "WHERE BodyData.CustomerNum=" & glngCustomerNum
            qrySelect = qrySelect & " AND UpperArm IS NOT NULL ORDER BY TreatDay ASC;"
            sngMin = MinValue("UpperArm") - 3
            sngMax = MaxValue("UpperArm") + 3
            sngStep = 3
            strTitle = "��ϵѷ� ��ȭ"
            Chart.Axis(AXIS_Y).Decimals = 1
        Case 3   'WHR
            qrySelect = "SELECT TreatDay, WHR FROM BodyData "
            qrySelect = qrySelect & "INNER JOIN Treat ON BodyData.TreatNum=Treat.TreatNum "
            qrySelect = qrySelect & "WHERE BodyData.CustomerNum=" & glngCustomerNum
            qrySelect = qrySelect & " AND WHR IS NOT NULL ORDER BY TreatDay ASC;"
            sngMin = MinValue("WHR") - 0.1
            sngMax = MaxValue("WHR") + 0.1
            sngStep = 0.1
            strTitle = "WHR ��ȭ"
            Chart.Axis(AXIS_Y).Decimals = 2
        Case 4    'VO2
            qrySelect = "SELECT TreatDay, VO2 FROM BodyData "
            qrySelect = qrySelect & "INNER JOIN Treat ON BodyData.TreatNum=Treat.TreatNum "
            qrySelect = qrySelect & "WHERE BodyData.CustomerNum=" & glngCustomerNum
            qrySelect = qrySelect & " AND VO2 IS NOT NULL ORDER BY TreatDay ASC;"
            sngMin = MinValue("VO2") - 10
            sngMax = MaxValue("VO2") + 10
            sngStep = 10
            strTitle = "VO2 ��ȭ"
            Chart.Axis(AXIS_Y).Decimals = 0
        Case 5    'RMR
            qrySelect = "SELECT TreatDay, RMR FROM BodyData "
            qrySelect = qrySelect & "INNER JOIN Treat ON BodyData.TreatNum=Treat.TreatNum "
            qrySelect = qrySelect & "WHERE BodyData.CustomerNum=" & glngCustomerNum
            qrySelect = qrySelect & " AND RMR IS NOT NULL ORDER BY TreatDay ASC;"
            sngMin = MinValue("RMR") - 50
            sngMax = MaxValue("RMR") + 50
            sngStep = 100
            strTitle = "RMR ��ȭ"
            Chart.Axis(AXIS_Y).Decimals = 0
        Case 6   'ü�����
            qrySelect = "SELECT TreatDay, ChFatRate FROM BodyData "
            qrySelect = qrySelect & "INNER JOIN Treat ON BodyData.TreatNum=Treat.TreatNum "
            qrySelect = qrySelect & "WHERE BodyData.CustomerNum=" & glngCustomerNum
            qrySelect = qrySelect & " AND ChFatRate IS NOT NULL ORDER BY TreatDay ASC;"
            sngMin = MinValue("ChFatRate") - 2
            sngMax = MaxValue("ChFatRate") + 2
            sngStep = 2
            strTitle = "ü����� ��ȭ"
            Chart.Axis(AXIS_Y).Decimals = 1
        Case 7   '������
            qrySelect = "SELECT TreatDay, Muscle FROM BodyData "
            qrySelect = qrySelect & "INNER JOIN Treat ON BodyData.TreatNum=Treat.TreatNum "
            qrySelect = qrySelect & "WHERE BodyData.CustomerNum=" & glngCustomerNum
            qrySelect = qrySelect & " AND Muscle IS NOT NULL ORDER BY TreatDay ASC;"
            sngMin = MinValue("Muscle") - 2
            sngMax = MaxValue("Muscle") + 2
            sngStep = 2
            strTitle = "������ ��ȭ"
            Chart.Axis(AXIS_Y).Decimals = 1
        Case 8   'Į�θ�(�Ļ��ϱ�)
            '�Ļ��ϱ� �Է��� ���� ��� ��Į�θ��� ������
            '�ִ밪, �ּҰ� ���� ���Ѵ�.
            qrySelect = "SELECT MAX(a), MIN(a) FROM ( "
            qrySelect = qrySelect & "SELECT SUM(MealCalory) AS a FROM DietDiary "
            qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
            qrySelect = qrySelect & " GROUP BY MealDate) total;"
            Set clsSelect = New clsSelect
            rValue = clsSelect.Query(qrySelect)
            If Not IsNull(rValue) Then
                If IsNull(rValue(0, 0)) And IsNull(rValue(1, 0)) Then
                    MsgBox "ǥ���� �Էµ����Ͱ� �����ϴ�.", vbExclamation, "��ȭ�� �׷���"
                    Chart.Visible = False
                    Exit Sub
                End If
                sngMax = CInt(rValue(0, 0)) + 10
                sngMin = CInt(rValue(1, 0)) - 10
            Else
                sngMin = 500
                sngMax = 3000
            End If
            Set clsSelect = Nothing
            
            qrySelect = "SELECT MealDate, SUM(MealCalory) FROM DietDiary "
            qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
            qrySelect = qrySelect & " GROUP BY MealDate ORDER BY MealDate ASC;"
            sngStep = 500
            strTitle = "�Ļ緮(Į�θ�) ��ȭ"
             Chart.Axis(AXIS_Y).Decimals = 0
       Case 9   'Į�θ�(��ϱ�)
            qrySelect = "SELECT MAX(a), MIN(a) FROM ( "
            qrySelect = qrySelect & "SELECT SUM(BurnCalories) AS a FROM Sportsdiary "
            qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
            qrySelect = qrySelect & " GROUP BY PlayDay) total;"
            Set clsSelect = New clsSelect
            rValue = clsSelect.Query(qrySelect)
            If Not IsNull(rValue) Then
                If IsNull(rValue(0, 0)) And IsNull(rValue(1, 0)) Then
                    MsgBox "ǥ���� �Էµ����Ͱ� �����ϴ�.", vbExclamation, "��ȭ�� �׷���"
                    Chart.Visible = False
                    Exit Sub
                End If
                sngMax = CInt(rValue(0, 0)) + 10
                sngMin = CInt(rValue(1, 0)) - 10
            Else
                sngMin = 100
                sngMax = 1500
            End If
            Set clsSelect = Nothing
            
            qrySelect = "SELECT PlayDay, SUM(BurnCalories) FROM Sportsdiary "
            qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
            qrySelect = qrySelect & " GROUP BY PlayDay ORDER BY PlayDay ASC;"
            sngStep = 200
            strTitle = "���(Į�θ�) ��ȭ"
            Chart.Axis(AXIS_Y).Decimals = 0
    End Select
    Set clsSelect = New clsSelect

    rValue = clsSelect.Query(qrySelect)

    If Not IsNull(rValue) Then
        Chart.Visible = True
        Chart.Title(CHART_TOPTIT) = strTitle
        Chart.OpenDataEx COD_VALUES, 0, COD_UNKNOWN
        Chart.Axis(AXIS_Y).Min = sngMin
        Chart.Axis(AXIS_Y).Max = sngMax
        For i = 0 To UBound(rValue, 2)
            Chart.ValueEx(0, i) = rValue(1, i)
            Chart.Axis(AXIS_X).Label(i) = Format(Is_Null(rValue(0, i), ""), "M/D")
        Next i
        Chart.CloseData COD_VALUES
    Else
        MsgBox "ǥ���� �Էµ����Ͱ� �����ϴ�.", vbExclamation, "��ȭ�� �׷���"
        Chart.Visible = False
    End If

    Set clsSelect = Nothing
    Exit Sub
Err:
    '2004-12-23 ������ �αױ��
    'WriteLog "Bottom1", "frmCounsel_1", Err.Number, Err.Description
    MsgBox Err.Number & Err.Description
End Sub

Private Function MinValue(strField As String) As Single
    Dim qrySelect As String, rMin As Variant

    Set clsSelect = New clsSelect

    If strField = "RMR" Then
        qrySelect = "SELECT MIN(rmr) FROM ( "
        If typCustomer.intAge >= ADULT_AGE Then
            qrySelect = qrySelect & "SELECT CASE AdBasicDsa "
        Else
            qrySelect = qrySelect & "SELECT CASE BaBasicDsa "
        End If
        qrySelect = qrySelect & "WHEN 8 THEN inRMR WHEN 9 THEN etcRMR ELSE RMR END AS rmr "
        qrySelect = qrySelect & "FROM BodyData INNER JOIN CompData "
        qrySelect = qrySelect & "ON BodyData.CompDataNum=CompData.CompDataNum "
        qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
        qrySelect = qrySelect & ") a"
    Else
        qrySelect = "SELECT MIN(" & strField & ") FROM BodyData WHERE CustomerNum=" & glngCustomerNum
    End If
    
    rMin = clsSelect.Query(qrySelect)
    If Not IsNull(rMin(0, 0)) Then
        MinValue = CSng(rMin(0, 0))
        Erase rMin
    Else
        MinValue = 0
    End If
    Set clsSelect = Nothing
End Function

Private Function MaxValue(strField As String) As Single
    Dim qrySelect As String, rMax As Variant

    Set clsSelect = New clsSelect
    
    If strField = "RMR" Then
        qrySelect = "SELECT MAX(rmr) FROM ( "
        If typCustomer.intAge >= ADULT_AGE Then
            qrySelect = qrySelect & "SELECT CASE AdBasicDsa "
        Else
            qrySelect = qrySelect & "SELECT CASE BaBasicDsa "
        End If
        qrySelect = qrySelect & "WHEN 8 THEN inRMR WHEN 9 THEN etcRMR ELSE RMR END AS rmr "
        qrySelect = qrySelect & "FROM BodyData INNER JOIN CompData "
        qrySelect = qrySelect & "ON BodyData.CompDataNum=CompData.CompDataNum "
        qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
        qrySelect = qrySelect & ") a"
    Else
        qrySelect = "SELECT MAX(" & strField & ") FROM BodyData WHERE CustomerNum=" & glngCustomerNum
    End If
    
    rMax = clsSelect.Query(qrySelect)
    If Not IsNull(rMax(0, 0)) Then
        MaxValue = CSng(rMax(0, 0))
        Erase rMax
    Else
        MaxValue = 0
    End If
    
    Set clsSelect = Nothing
End Function

'ó�泻����ȸ ���� �����ư���� Ŭ���̺�Ʈ
Private Sub Bottom2(intIndex As Integer)
    Dim i As Integer
    Dim qrySelect As String, rValue As Variant

    grdTable.Visible = True
    Chart.Visible = False
    cmdSub(0).Visible = True
    cmdSub(1).Visible = True
    For i = 2 To 9
        cmdSub(i).Visible = False
    Next i
    cmdSub(0).Caption = "ó      ġ"
    cmdSub(1).Caption = "��  ��  ��"

    Select Case intIndex
        Case 0
            'óġ ���̺� ó�泻����ȸ ���̺�� �ʱ�ȭ
            With grdTable
                Set clsSelect = New clsSelect
                .Clear
                .Rows = 3
                .RowHeight(1) = 0
                .TextMatrix(0, 0) = "������"
                qrySelect = "SELECT TreatName, TreatCode FROM TreatCode;"
                rValue = clsSelect.Query(qrySelect)

                If Not IsNull(rValue) Then
                    .Cols = UBound(rValue, 2) + 3
                    For i = 0 To UBound(rValue, 2)
                        .TextMatrix(0, i + 1) = Trim(rValue(0, i))
                        .TextMatrix(1, i + 1) = Trim(rValue(1, i))
                    Next i
                Else
                    .Cols = 2
                    .Rows = 3
                End If
                For i = 0 To .Cols - 2
                    .ColWidth(i) = 1000
                Next i
                .ColWidth(.Cols - 1) = 0
                Set clsSelect = Nothing
            End With
            
            'óġ������ �����ֱ�
            Call ShowTable_TreatCode
        Case 1
             'óġ ���̺� ��¹�������ȸ ���̺�� �ʱ�ȭ
            With grdTable
                Set clsSelect = New clsSelect
                .Clear
                .Rows = 3
                .RowHeight(1) = 0
                .TextMatrix(0, 0) = "������"
                qrySelect = "SELECT PrintoutName, PrintoutNum FROM Printout;"
                rValue = clsSelect.Query(qrySelect)

                If Not IsNull(rValue) Then
                    .Cols = UBound(rValue, 2) + 3
                    For i = 0 To UBound(rValue, 2)
                        .TextMatrix(0, i + 1) = Trim(rValue(0, i))
                        .TextMatrix(1, i + 1) = Trim(rValue(1, i))
                    Next i
                End If
                For i = 0 To .Cols - 2
                    .ColWidth(i) = 1000
                Next i
                .ColWidth(.Cols - 1) = 0
                Set clsSelect = Nothing
            End With
            Call ShowTable_Print
    End Select
End Sub

Private Function LoadCustomerInfo() As Boolean
    Dim qrySelect As String, rValue As Variant
On Error GoTo SelErr
    Set clsSelect = New clsSelect

    'ȯ�ڱ⺻����
    qrySelect = "SELECT CustomName, JuminNum, Age, Sex, ChPgCode FROM CustomerInfo WHERE CustomerNum=" & glngCustomerNum

    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue) Then
        Me.Enabled = True
        With typCustomer
            .strCustomName = Trim(rValue(0, 0))
            .strJuminNum = Trim(rValue(1, 0))
            .intAge = CInt(rValue(2, 0))
            .strSex = Trim(rValue(3, 0))
            .strChPgCode = Is_Null(rValue(4, 0), 0)
        End With
    Else
        MsgBox "��ü����/ó���� �Է��� ȯ�ڸ� �����Ͻʽÿ�." & vbNewLine & vbNewLine & "ó�� �湮�� ȯ���̸� 'ȯ�ڵ��'�� ���� �Ͻʽÿ�.", vbOKOnly + vbCritical
        LoadCustomerInfo = False
        Me.Enabled = False
        Exit Function
    End If
    
    LoadCustomerInfo = True
    Set clsSelect = Nothing
    Exit Function
SelErr:
    '2004-12-23 ������ �αױ��
    'WriteLog "LoadCustomerInfo", "frmCounsel_1", Err.Number, Err.Description
    LoadCustomerInfo = False
End Function

Private Sub InitialControl()
    Dim rValue As Variant
    Dim i As Integer

    Set clsSelect = New clsSelect

    '�񸸰��� �׸� �ʱ�ȭ
    For i = 0 To 7
        lblObInfo(i).Caption = ""                   '����
        Set imgUpDown(i).Picture = LoadPicture("")  '��ȸ��� �̹���
        lblUpDown(i).Caption = ""                   '��ȸ��� ��
        lblMax(i).Caption = ""                      '�ִ�ġ
        lblMin(i).Caption = ""                      '�ּ�ġ
    Next i

    'ġ�����α׷� �ʱ�ȭ
    rValue = clsSelect.Query("SELECT ChPgName, ChPgCode FROM ChPgUpdate;")

    If Not IsNull(rValue) Then
        cmbProgram.Clear
        For i = 0 To UBound(rValue, 2)
            cmbProgram.AddItem Trim(rValue(0, i))
            If Trim(rValue(1, i)) = Trim(typCustomer.strChPgCode) Then
                cmbProgram.ListIndex = i
            End If
        Next i
    End If
    
    txtUserCalory.Text = ""
    'Memo / Notice
    txtNotice.Text = ""
    txtMemo.Text = ""

    Set clsSelect = Nothing
End Sub

'Į�θ�ó�� ������ �ʱ�ȭ
Private Sub InitialTreatCalory()
    Dim i As Integer
    Dim intExTime As Integer
'200 / 300 / 400 / 500
'�����ȱ� : 0.093 / �ٷ¿ : 0.105
    typTreatCal.intDietCal = 0
    typTreatCal.intExCal = 0
    typTreatCal.intExDay = 0
    typTreatCal.intUserCal = 0
    typTreatCal.sngLossWeight = 0
    If typObesity.sngWeight = 0 Then
        For i = 0 To 3
            lblAerobic(i).Caption = ""
            lblAnaerobic(i).Caption = ""
        Next i
        lblLossWeight.Caption = ""
        Exit Sub
    End If
    '����ҿ�� �ٷ¿ �����ֱ�
    For i = 0 To 3
        intExTime = ((i + 2) * 100) / (typObesity.sngWeight * 0.093)
        lblAerobic(i).Caption = intExTime & "��"
        
        intExTime = ((i + 2) * 100) / (typObesity.sngWeight * 0.105)
        lblAnaerobic(i).Caption = intExTime & "��"
    Next i
    
    For i = 0 To 5
        Set imgDietCal(i).Picture = LoadPicture(App.Path & IMG_ME & i & "gray.jpg")
    Next i
    For i = 0 To 3
        Set imgExDay(i).Picture = LoadPicture(App.Path & IMG_EX & i + 3 & "��-gray.jpg")
    Next i
    For i = 0 To 3
        Set imgExCal(i).Picture = LoadPicture(App.Path & IMG_EX & i + 2 & "00-gray.jpg")
    Next i
    lblLossWeight.Caption = ""
    If typTreatCal.sngTreatCal <> 0 Then    'ó���� �Ĵܿ���
        lblTreatCal.Caption = Format(typTreatCal.sngTreatCal, "#,###") & " kcal"
    End If
    Call chkUserCalory_Click
End Sub

Private Sub InitialSpread4()
'ó���Է��ϴ� �������� �ʱ�ȭ(ġ������/��¹� ȭ���� ġ������)
'�ϴ� ��� ó���� �ҷ�����
'����Ǿ� �ִ� ó���� ���� �ÿ��� �װͿ� üũ
    Dim qrySelect As String
    Dim rValue As Variant, i As Integer, j As Integer

    With sprTreat
        .EditEnterAction = EditEnterActionDown
        .EditModePermanent = True
        
        .GrayAreaBackColor = vbWhite
        .BackColor = &HCEF7E7
        .GridColor = vbWhite
        .Font.Size = 7
        .RowHeight(-1) = 13
        .MaxCols = 4
        .ColWidth(1) = 9
        .ColWidth(2) = 2
        .ColWidth(3) = 2
        .ColWidth(4) = 0
        .Col = 2: .Row = -1
        .TypeHAlign = TypeHAlignCenter
        .TypeVAlign = TypeVAlignCenter
        .DisplayColHeaders = False
        .DisplayRowHeaders = False
        .ScrollBars = ScrollBarsNone

        Set clsSelect = New clsSelect
        rValue = clsSelect.Query("SELECT TreatName, TreatCode FROM TreatCode;")
        If Not IsNull(rValue) Then
            .MaxRows = UBound(rValue, 2) + 1
            If .MaxRows > 6 Then
                .ScrollBars = ScrollBarsVertical
                .ColWidth(1) = 8.5
                .Width = 1800
            Else
                .ScrollBars = ScrollBarsNone
                .ColWidth(1) = 9
                .Width = 1650
            End If
            .Col = 1
            For i = 0 To UBound(rValue, 2)
                .Row = i + 1
                .Col = 1: .Text = Trim(rValue(0, i)): .Lock = True
                .Col = 2: .CellType = CellTypeCheckBox: .Value = False
                .TypeCheckCenter = True
                .TypeCheckType = TypeCheckTypeNormal
                .Col = 3: .CellType = CellTypeCheckBox: .Value = False
                .Col = 4: .Text = Trim(rValue(1, i))
            Next i
        Else
            .MaxRows = 1
            .Row = 1: .Col = 1: .Text = ""
            .Col = 2: .CellType = CellTypeCheckBox: .Value = False
            .Col = 3: .CellType = CellTypeCheckBox: .Value = False
            .Col = 4: .Text = ""
            Exit Sub
        End If

        '����Ǿ� �ִ��� Ȯ���ϰ� �����ÿ��� �ش� ġ����� üũ�Ѵ�
        '���� ���� �޸��Ѵ�.
        qrySelect = "SELECT ReserveDetail.ReDetailNum, ReserveTreat.TreatCode "
        qrySelect = qrySelect & "FROM ReserveDetail INNER JOIN "
        qrySelect = qrySelect & "Reserve ON ReserveDetail.ReserveNum = Reserve.ReserveNum "
        qrySelect = qrySelect & "INNER JOIN ReserveTreat ON "
        qrySelect = qrySelect & "ReserveDetail.ReDetailNum = ReserveTreat.ReDetailNum "
        qrySelect = qrySelect & "WHERE ReserveDetail.ReserveDay='" & Format(gdatUserDay, "YYYYMMDD") & "' "
        qrySelect = qrySelect & "AND Reserve.CustomerNum = " & glngCustomerNum

        rValue = clsSelect.Query(qrySelect)
        If Not IsNull(rValue) Then
            For i = 0 To UBound(rValue, 2)
                For j = 1 To .MaxRows
                    .Row = j: .Col = 1
                    If Trim(.Text) = Trim(rValue(1, i)) Then
                        .Col = 2: .CellType = CellTypeCheckBox: .Value = True
                        Exit For
                    End If
                Next j
            Next i
        End If
        Set clsSelect = Nothing
        .Lock = True
    End With
End Sub

Private Sub InitialSpread5()
'ó���Է��ϴ� �������� �ʱ�ȭ(ġ������/��¹� ȭ���� ��¹�)
'�ϴ� ��� ��¹��� �ҷ�����
'����Ǿ� �ִ� ��¹��� ���� �ÿ��� �װͿ� üũ
    Dim qrySelect As String
    Dim rValue As Variant, i As Integer, j As Integer

    With sprPrint
        .EditEnterAction = EditEnterActionDown
        .EditModePermanent = True

        .GrayAreaBackColor = vbWhite
        .BackColor = &HCEF7E7
        .GridColor = vbWhite
        .Font.Size = 7
        .RowHeight(-1) = 13
        .MaxCols = 3
        .ColWidth(1) = 10
        .ColWidth(2) = 2
        .ColWidth(3) = 0
        .DisplayColHeaders = False
        .DisplayRowHeaders = False
        .ScrollBars = ScrollBarsNone

        Set clsSelect = New clsSelect
        rValue = clsSelect.Query("SELECT PrintoutName, PrintoutNum FROM PrintOut;")
        If Not IsNull(rValue) Then
            .MaxRows = UBound(rValue, 2) + 1
            If .MaxRows > 6 Then
                .ScrollBars = ScrollBarsVertical
                .ColWidth(1) = 9.5
                .Width = 1650
            Else
                .ScrollBars = ScrollBarsNone
                .ColWidth(1) = 10
                .Width = 1600
            End If
            .Col = 1
            For i = 0 To UBound(rValue, 2)
                .Row = i + 1
                .Col = 1: .Text = Trim(rValue(0, i)): .Lock = True
                .Col = 2: .CellType = CellTypeCheckBox: .Value = False
                .Col = 3: .Text = Trim(rValue(1, i))
            Next i
        Else
            .MaxRows = 1
            .Row = 1: .Col = 1: .Text = ""
            .Col = 2: .CellType = CellTypeCheckBox: .Value = False
            .Col = 3: .Text = ""
            Exit Sub
        End If

        '����Ǿ� �ִ� ��¹����� �ϴ� üũ�ϱ�
        '����Ǿ� �ִ��� Ȯ���ϰ� �����ÿ��� �ش� ��¹����� üũ�Ѵ�
        qrySelect = "SELECT ReserveDetail.ReDetailNum, ReservePrint.PrintCode "
        qrySelect = qrySelect & "FROM ReserveDetail INNER JOIN "
        qrySelect = qrySelect & "Reserve ON ReserveDetail.ReserveNum=Reserve.ReserveNum "
        qrySelect = qrySelect & "INNER JOIN ReservePrint ON ReserveDetail.ReDetailNum=ReservePrint.ReDetailNum "
        qrySelect = qrySelect & "WHERE ReserveDetail.ReserveDay='" & Format(gdatUserDay, "YYYYMMDD") & "' "
        qrySelect = qrySelect & "AND Reserve.CustomerNum=" & glngCustomerNum
        
        rValue = clsSelect.Query(qrySelect)
        If Not IsNull(rValue) Then
            For i = 0 To UBound(rValue, 2)
                For j = 1 To .MaxRows
                    .Row = j: .Col = 1
                    If Trim(.Text) = Trim(rValue(1, i)) Then
                        .Col = 2: .CellType = CellTypeCheckBox: .Value = True
                        Exit For
                    End If
                Next j
            Next i
        End If
        Set clsSelect = Nothing
    End With
End Sub

Private Sub InitialChart()
    With Chart
        .Gallery = LINES
        .Chart3D = False
        .MarkerShape = MK_RECT
        .MarkerSize = 2
        .AxesStyle = CAS_FLATFRAME
        .Axis(0).Grid = True

        ' Color Settings
        .Border = False
        .RGBBk = vbWhite

        ' Layout Settings
        .LegendBox = False
        .SerLegBox = False
        .ToolBar = False
        .PointLabels = True
        .MultipleColors = False
        
        .AllowDrag = False
        .AllowEdit = False
        .AllowResize = False
    End With
End Sub

Private Function InputValues(sngValue As Single) As String
    Dim strQuery As String

    If sngValue = 0 Then
        strQuery = "NULL"
    Else
        strQuery = CStr(sngValue)
    End If
    InputValues = strQuery
End Function

Private Sub ShowTreatData(lngTreatNum As Long)
    Dim qrySelect As String, rValue As Variant
    Dim i As Integer, j As Integer

    Set clsSelect = New clsSelect
    For j = 0 To sprTreat.MaxRows
        sprTreat.Row = j
        sprTreat.Col = 2
        sprTreat.Value = 0
        sprTreat.Col = 3
        sprTreat.Value = 0
    Next j

    qrySelect = "SELECT TreatCode, Execution FROM TreatData WHERE TreatNum=" & lngTreatNum
    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue) Then
        For i = 0 To UBound(rValue, 2)
            For j = 0 To sprTreat.MaxRows
                sprTreat.Col = 4
                sprTreat.Row = j
                If Trim(sprTreat.Text) = Trim(rValue(0, i)) Then
                    sprTreat.Col = 2
                    sprTreat.Value = 1
                    sprTreat.Col = 3: sprTreat.Lock = False
                    If Trim(rValue(1, i)) = "Y" Then
                        sprTreat.Col = 3
                        sprTreat.Value = 1
                    End If
                    Exit For
                End If
            Next j
        Next i
    End If

    Set clsSelect = Nothing
End Sub

Private Sub ShowTable_Circumference()
'�߰��� ����ڰ� �߰��� �Է°��� ������...?
'BioChemData���� �ش簪�� �����´�. TreatNum����..
'�ٵ� ������ ��� ���߳�..�߰��� ����� �߰����� ������..
    Dim qrySelect As String, rValue As Variant
    Dim rUValue As Variant, intUEnd As Integer, intUCount As Integer
    Dim i As Integer, j As Integer
    
    Set clsSelect = New clsSelect
    
    With grdTable
    intUCount = 0
    qrySelect = "SELECT TreatDay"
    For i = 1 To .Cols - 2
        qrySelect = qrySelect & "," & Trim(.TextMatrix(1, i + 1))
        If IsNumeric(.TextMatrix(1, i + 1)) = True Then
            intUEnd = i + 1
            intUCount = intUCount + 1
        End If
    Next i
    qrySelect = qrySelect & " FROM BodyData INNER JOIN Treat ON "
    qrySelect = qrySelect & "BodyData.TreatNum=Treat.TreatNum "
    qrySelect = qrySelect & "WHERE BodyData.CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & " ORDER BY TreatDay ASC;"
    
    rValue = clsSelect.Query(qrySelect)
    
    If Not IsNull(rValue) Then
        .Rows = .Rows + UBound(rValue, 2)
        For i = 0 To UBound(rValue, 2)
            .TextMatrix(i + 2, 0) = i + 1
            For j = 0 To .Cols - 2
                If Not IsNull(rValue(j, i)) Then
                    .TextMatrix(i + 2, j + 1) = Trim(rValue(j, i))
                Else
                    .TextMatrix(i + 2, j + 1) = "-"
                End If
            Next j
            '����� �߰� �Է°�
            If intUCount > 0 Then
                For j = (intUEnd - intUCount + 1) To intUEnd
                    qrySelect = "SELECT BioChemSu FROM BioChemData "
                    qrySelect = qrySelect & "WHERE BioChemCode='" & Trim(.TextMatrix(1, j))
                    qrySelect = qrySelect & "' AND TreatNum=" & Trim(.TextMatrix(i + 2, .Cols - 1))
                    rUValue = clsSelect.Query(qrySelect)
                    If Not IsNull(rUValue) Then
                        .TextMatrix(i + 2, j) = Trim(rUValue(0, 0))
                    Else
                        .TextMatrix(i + 2, j) = "-"
                    End If
                Next j
            End If
            If rValue(0, i) = gdatUserDay Then
                .Row = i + 2: .Col = 0: .ColSel = .Cols - 1
                glngTreatNum = .TextMatrix(i + 2, .Cols - 1)
            End If
        Next i
    End If
    End With
    
    Set clsSelect = Nothing
    
End Sub

Private Sub ShowTreatPrint(lngTreatNum As Long)
    Dim qrySelect As String, rValue As Variant
    Dim i As Integer, j As Integer

    Set clsSelect = New clsSelect
    For j = 0 To sprPrint.MaxRows
        sprPrint.Row = j
        sprPrint.Col = 2
        sprPrint.Value = 0
    Next j

    qrySelect = "SELECT PrintoutNum FROM TreatPrint WHERE TreatNum=" & lngTreatNum
    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue) Then
        For i = 0 To UBound(rValue, 2)
            For j = 0 To sprPrint.MaxRows
                sprPrint.Col = 3
                sprPrint.Row = j
                If Trim(sprPrint.Text) = Trim(rValue(0, i)) Then
                    sprPrint.Col = 2
                    sprPrint.Value = 1
                    Exit For
                End If
            Next j
        Next i
    End If

    Set clsSelect = Nothing
End Sub

Private Function SaveTreat() As Long
    Dim qryInsert As String
    Dim lngTreatNum As Long

On Error GoTo InsertErr
    If txtUserCalory.Text = "" Then
        typTreatCal.intUserCal = 0
    Else
        If IsNumeric(txtUserCalory.Text) = True Then
            If CInt(txtUserCalory.Text) < 1000 Or CInt(txtUserCalory.Text) > 3500 Then
                MsgBox "�Ĵܿ����� 1000kcal�̻� 3500kcal���Ϸ� �Է��Ͻʽÿ�.", vbOKOnly + vbExclamation
                txtUserCalory.SetFocus
                SaveTreat = 0
                Exit Function
            Else
                typTreatCal.intUserCal = CInt(txtUserCalory.Text)
            End If
        Else
            typTreatCal.intUserCal = 0
        End If
    End If
    
    lngTreatNum = WhatisCode("Treat", "TreatNum")

    qryInsert = "INSERT INTO Treat("
    qryInsert = qryInsert & "TreatNum,"
    qryInsert = qryInsert & "CustomerNum,"
    qryInsert = qryInsert & "TreatDay, "
    qryInsert = qryInsert & "LossWeight,"
    qryInsert = qryInsert & "ExDay,"
    qryInsert = qryInsert & "ExCalory,"
    qryInsert = qryInsert & "DietCalory, "
    qryInsert = qryInsert & "TreatCalory,"
    qryInsert = qryInsert & "Notice,"
    qryInsert = qryInsert & "Memo,"
    qryInsert = qryInsert & "main,"
    qryInsert = qryInsert & "sub1,"
    qryInsert = qryInsert & "sub2,"
    qryInsert = qryInsert & "sub3,"
    qryInsert = qryInsert & "sub4,"
    qryInsert = qryInsert & "UserCalory"
    qryInsert = qryInsert & ") "
    qryInsert = qryInsert & "VALUES(" & lngTreatNum
    qryInsert = qryInsert & "," & glngCustomerNum
    qryInsert = qryInsert & ",'" & Format(gdatUserDay, "YYYYMMDD") & "'"
    qryInsert = qryInsert & "," & InputValues(typTreatCal.sngLossWeight)
    qryInsert = qryInsert & "," & InputValues(CSng(typTreatCal.intExDay))
    qryInsert = qryInsert & "," & InputValues(CSng(typTreatCal.intExCal))
    qryInsert = qryInsert & "," & InputValues(CSng(typTreatCal.intDietCal))
    qryInsert = qryInsert & "," & InputValues(typTreatCal.sngTreatCal)
    qryInsert = qryInsert & ",'" & Trim(txtNotice.Text) & "'"
    qryInsert = qryInsert & ",'" & Trim(txtMemo.Text) & "'"
    qryInsert = qryInsert & "," & InputValues(CSng(gintMain))
    qryInsert = qryInsert & "," & InputValues(CSng(gintSub1))
    qryInsert = qryInsert & "," & InputValues(CSng(gintSub2))
    qryInsert = qryInsert & "," & InputValues(CSng(gintSub3))
    qryInsert = qryInsert & "," & InputValues(CSng(gintSub4))
    qryInsert = qryInsert & "," & InputValues(CInt(typTreatCal.intUserCal))
    qryInsert = qryInsert & ");"
    modSql.AdoExcuteSql qryInsert

    SaveTreat = lngTreatNum
    Exit Function
InsertErr:
    '2004-12-23 ������ �αױ��
    'WriteLog "SaveTreat", "frmCounsel_1", Err.Number, Err.Description
    SaveTreat = 0
End Function

Private Function UpdateTreat(lngTreatNum As Long) As Boolean
    Dim qryUpdate As String

On Error GoTo UpdateErr
    If txtUserCalory.Text = "" Then
        typTreatCal.intUserCal = 0
    Else
        If IsNumeric(txtUserCalory.Text) = True Then
            If CInt(txtUserCalory.Text) < 1000 Or CInt(txtUserCalory.Text) > 3500 Then
                MsgBox "�Ĵܿ����� 1000kcal�̻� 3500kcal���Ϸ� �Է��Ͻʽÿ�.", vbOKOnly + vbExclamation
                txtUserCalory.SetFocus
                UpdateTreat = False
                Exit Function
            Else
                typTreatCal.intUserCal = CInt(txtUserCalory.Text)
            End If
        Else
            typTreatCal.intUserCal = 0
        End If
    End If

    qryUpdate = "UPDATE Treat SET "
    qryUpdate = qryUpdate & "LossWeight=" & InputValues(typTreatCal.sngLossWeight)
    qryUpdate = qryUpdate & ", ExDay=" & InputValues(CSng(typTreatCal.intExDay))
    qryUpdate = qryUpdate & ", ExCalory=" & InputValues(CSng(typTreatCal.intExCal))
    qryUpdate = qryUpdate & ", DietCalory=" & InputValues(CSng(typTreatCal.intDietCal))
    qryUpdate = qryUpdate & ", TreatCalory=" & InputValues(typTreatCal.sngTreatCal)
    qryUpdate = qryUpdate & ", Notice='" & Trim(txtNotice.Text) & "'"
    qryUpdate = qryUpdate & ", Memo='" & Trim(txtMemo.Text) & "'"
    qryUpdate = qryUpdate & ", main=" & InputValues(CSng(gintMain))
    qryUpdate = qryUpdate & ", sub1=" & InputValues(CSng(gintSub1))
    qryUpdate = qryUpdate & ", sub2=" & InputValues(CSng(gintSub2))
    qryUpdate = qryUpdate & ", sub3=" & InputValues(CSng(gintSub3))
    qryUpdate = qryUpdate & ", sub4=" & InputValues(CSng(gintSub4))
    
    If chkUserCalory.Value = 1 Then
        qryUpdate = qryUpdate & ", UserCalory=" & InputValues(typTreatCal.intUserCal)
    Else
        qryUpdate = qryUpdate & ", UserCalory=Null"
    End If
        qryUpdate = qryUpdate & " WHERE TreatNum=" & lngTreatNum & ";"

        modSql.AdoExcuteSql (qryUpdate)
        UpdateTreat = True
    Exit Function
UpdateErr:
    '2004-12-23 ������ �αױ��
    'WriteLog "UpdateTreat", "frmCounsel_1", Err.Number, Err.Description
    UpdateTreat = False
End Function

Private Function SaveTreatData() As Boolean
    Dim qryInsert As String, qryDelete As String
    Dim lngTreatDataNum As Long, i As Integer

On Error GoTo InsertErr
    qryDelete = "DELETE FROM TreatData WHERE TreatNum=" & glngTreatNum
    modSql.AdoExcuteSql (qryDelete)

    lngTreatDataNum = WhatisCode("TreatData", "TreatDataNum")
    With sprTreat
        For i = 1 To .MaxRows
            .Row = i
            .Col = 2
            If .Value <> 0 Then
                qryInsert = "INSERT INTO TreatData(TreatDataNum, ReDetailNum, TreatNum, TreatCode, Execution) "
                qryInsert = qryInsert & "VALUES(" & lngTreatDataNum
                qryInsert = qryInsert & ",0"    '******* �պ���
                qryInsert = qryInsert & "," & glngTreatNum
                .Col = 4
                qryInsert = qryInsert & "," & Trim(.Text)
                .Col = 3
                If .Value = 1 Then
                    qryInsert = qryInsert & ",'Y');"
                Else
                    qryInsert = qryInsert & ",'N');"
                End If

                modSql.AdoExcuteSql (qryInsert)
                lngTreatDataNum = lngTreatDataNum + 1
            End If
        Next i
    End With
    SaveTreatData = True
    Exit Function
InsertErr:
    '2004-12-23 ������ �αױ��
    'WriteLog "SaveTreatData", "frmCounsel_1", Err.Number, Err.Description
    SaveTreatData = False

End Function

Private Function SaveTreatPrint() As Boolean
    Dim qryInsert As String, qryDelete As String
    Dim lngTreatPrintNum As Long, i As Integer
On Error GoTo InsertErr
    qryDelete = "DELETE FROM TreatPrint WHERE TreatNum=" & glngTreatNum
    modSql.AdoExcuteSql (qryDelete)

    lngTreatPrintNum = WhatisCode("TreatPrint", "TreatPrintNum")
    With sprPrint
        For i = 1 To .MaxRows
            .Row = i
            .Col = 2
            If .Value <> 0 Then
                qryInsert = "INSERT INTO TreatPrint(TreatPrintNum, ReDetailNum, TreatNum, PrintoutNum) "
                qryInsert = qryInsert & "VALUES(" & lngTreatPrintNum
                qryInsert = qryInsert & ",0"      '******* ���̼պ���
                qryInsert = qryInsert & "," & glngTreatNum
                .Col = 3
                qryInsert = qryInsert & "," & Trim(.Text) & ");"

                modSql.AdoExcuteSql (qryInsert)
                lngTreatPrintNum = lngTreatPrintNum + 1
            End If
        Next i
    End With
    SaveTreatPrint = True
    Exit Function
InsertErr:
    '2004-12-23 ������ �αױ��
    'WriteLog "SaveTreatPrint", "frmCounsel_1", Err.Number, Err.Description
    SaveTreatPrint = False

End Function

Private Function DeleteTreat(lngTreatNum As Long) As Boolean
    Dim qryDelete As String
    
On Error GoTo DelErr
    'TreatData  ����
    qryDelete = "DELETE FROM TreatData WHERE TreatNum=" & lngTreatNum
    modSql.AdoExcuteSql (qryDelete)
    
    'TreatPrint ����
    qryDelete = "DELETE FROM TreatPrint WHERE TreatNum=" & lngTreatNum
    modSql.AdoExcuteSql (qryDelete)
    
    'BodyData ����
    qryDelete = "DELETE FROM BodyData WHERE TreatNum=" & lngTreatNum
    modSql.AdoExcuteSql (qryDelete)

    'Treat ����
    qryDelete = "DELETE FROM Treat WHERE TreatNum=" & lngTreatNum
    modSql.AdoExcuteSql (qryDelete)
    
    DeleteTreat = True
    Exit Function
DelErr:
    '2004-12-23 ������ �αױ��
    'WriteLog "DeleteTreat", "frmCounsel_1", Err.Number, Err.Description
    DeleteTreat = False
End Function

Public Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub grdTable_Click()
    If grdTable.Row <= 1 Then
        Exit Sub
    End If

    If grdTable.TextMatrix(grdTable.Row, grdTable.Cols - 1) = "" Or IsNumeric(grdTable.TextMatrix(grdTable.Row, grdTable.Cols - 1)) = False Then
        glngTreatNum = 0
        Exit Sub
    End If

    glngTreatNum = grdTable.TextMatrix(grdTable.Row, grdTable.Cols - 1)
    If glngTreatNum <> 0 Then
        '��ȸ,���� ���
        intMode = 2
        cmdPreTreat.Visible = False
        
        Call ShowTreat(glngTreatNum)
        Call ShowTreatData(glngTreatNum)
        Call ShowTreatPrint(glngTreatNum)
        
        cmdInputTreat.Visible = True
        Call EnabledInput(False)
    End If
    
End Sub

Private Sub ShowTreat(lngTreatNum As Long)
    Dim qrySelect As String, rValue As Variant
    Dim i As Integer, Index As Integer
    Dim sngWeight As Single, intExTime As Integer

    Set clsSelect = New clsSelect

    qrySelect = "SELECT LossWeight, ExCalory, DietCalory, TreatCalory, ExDay, Memo "
    qrySelect = qrySelect & ",main, sub1, sub2, sub3, sub4, UserCalory "
    qrySelect = qrySelect & ",Weight "
    qrySelect = qrySelect & "FROM Treat LEFT JOIN BodyData "
    qrySelect = qrySelect & "ON Treat.TreatNum=BodyData.TreatNum "
    qrySelect = qrySelect & "WHERE Treat.TreatNum=" & lngTreatNum
    
    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue(0, 0)) Then
        typTreatCal.sngLossWeight = Is_Null(rValue(0, 0), 0)
        typTreatCal.intExCal = Is_Null(rValue(1, 0), 0)
        typTreatCal.intDietCal = Is_Null(rValue(2, 0), 0)
        typTreatCal.sngTreatCal = Is_Null(rValue(3, 0), typObesity.sngTEE)
        typTreatCal.intExDay = Is_Null(rValue(4, 0), 0)
        typTreatCal.intUserCal = Is_Null(rValue(11, 0), 0)
        
        txtMemo.Text = Is_Null(rValue(5, 0), "")
        
        gintMain = Is_Null(rValue(6, 0), 0)
        gintSub1 = Is_Null(rValue(7, 0), 0)
        gintSub2 = Is_Null(rValue(8, 0), 0)
        gintSub3 = Is_Null(rValue(9, 0), 0)
        gintSub4 = Is_Null(rValue(10, 0), 0)
    Else
        typTreatCal.sngLossWeight = 0
        typTreatCal.intExCal = 0
        typTreatCal.intDietCal = 0
        typTreatCal.sngTreatCal = 0
        typTreatCal.intExDay = 0
        typTreatCal.intUserCal = 0
        
        txtMemo.Text = ""
        
        gintMain = 0
        gintSub1 = 0
        gintSub2 = 0
        gintSub3 = 0
        gintSub4 = 0
    End If
    If Not IsNull(rValue(12, 0)) Then
        sngWeight = rValue(12, 0)
        For i = 0 To 3
            intExTime = ((i + 2) * 100) / (sngWeight * 0.093)
            lblAerobic(i).Caption = intExTime & "��"
            
            intExTime = ((i + 2) * 100) / (sngWeight * 0.105)
            lblAnaerobic(i).Caption = intExTime & "��"
        Next i
    End If
    
    qrySelect = "SELECT Notice "
    qrySelect = qrySelect & "FROM Treat WHERE TreatNum=" & lngTreatNum
    
    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue(0, 0)) Then
        txtNotice.Text = rValue(0, 0)
    Else
        txtNotice.Text = ""
    End If
    
    Set clsSelect = Nothing
    
    '�Ļ�Į�θ�
    With typTreatCal
    
        For i = 0 To 5
            Set imgDietCal(i).Picture = LoadPicture(App.Path & IMG_ME & i & "gray.jpg")
        Next i
        Select Case .intDietCal
            Case 125: Index = 0
            Case 250: Index = 1
            Case 375: Index = 2
            Case 500: Index = 3
            Case 750: Index = 4
            Case 1000: Index = 5
            Case Else: Index = 6
        End Select
        If Index <> 6 Then
            Set imgDietCal(Index).Picture = LoadPicture(App.Path & IMG_ME & Index & "green.jpg")
        End If
        
        '�����ʿ俭���� ���� �����ȵǴ� Į�θ��� ����..�׷� Į�θ��� �ƿ� �Ⱥ��̵��� 2004.04.01
        Select Case typObesity.sngTEE
            Case 1000 To 1249:      '1000~1200
                For i = 0 To 5
                    imgDietCal(i).Visible = False
                Next i
            Case 1250 To 1449:      '1300~1400
                imgDietCal(0).Visible = False
                imgDietCal(1).Visible = True
                For i = 2 To 5
                    imgDietCal(i).Visible = False
                Next i
            Case 1450 To 1749:      '1500~1700
                imgDietCal(0).Visible = False
                imgDietCal(1).Visible = True
                imgDietCal(2).Visible = False
                imgDietCal(3).Visible = True
                imgDietCal(4).Visible = False
                imgDietCal(5).Visible = False
            Case 1750 To 1949:      '1800~1900
                imgDietCal(0).Visible = False
                imgDietCal(1).Visible = True
                imgDietCal(2).Visible = False
                imgDietCal(3).Visible = True
                imgDietCal(4).Visible = True
                imgDietCal(5).Visible = False
            Case 1950 To 3500:      '2000~3500
                imgDietCal(0).Visible = False
                imgDietCal(1).Visible = True
                imgDietCal(2).Visible = False
                For i = 3 To 5
                    imgDietCal(i).Visible = True
                Next i
            Case Else
                For i = 0 To 5
                    imgDietCal(i).Visible = False
                Next i
        End Select
        
        '��ϼ�
        For i = 0 To 3
            Set imgExDay(i).Picture = LoadPicture(App.Path & IMG_EX & i + 3 & "��-gray.jpg")
        Next i
        If .intExDay >= 3 And .intExDay <= 6 Then
            Set imgExDay(.intExDay - 3).Picture = LoadPicture(App.Path & IMG_EX & .intExDay & "��-green.jpg")
        End If
        
        '�Į�θ�
        For i = 0 To 3
            Set imgExCal(i).Picture = LoadPicture(App.Path & IMG_EX & i + 2 & "00-gray.jpg")
        Next i
        If .intExCal >= 200 And .intExCal <= 500 Then
            Set imgExCal((.intExCal / 100) - 2).Picture = LoadPicture(App.Path & IMG_EX & (.intExCal / 100) & "00-green.jpg")
        End If
        
        '����ü��
        lblLossWeight.Caption = Format(.sngLossWeight, "0.0") & " kg/��"
        If .intUserCal = 0 Then
            txtUserCalory.Text = ""
            txtUserCalory.BackColor = FRM_GRAY
            txtUserCalory.Enabled = False
            chkUserCalory.Value = 0
        Else
            For i = 0 To 5
                Set imgDietCal(i).Picture = LoadPicture(App.Path & IMG_ME & i & "gray.jpg")
                imgDietCal(i).Enabled = False
            Next i
            txtUserCalory.Text = .intUserCal
            txtUserCalory.BackColor = vbWhite
            txtUserCalory.Enabled = True
            chkUserCalory.Value = 1
            .sngTreatCal = .intUserCal
        End If
        lblTreatCal.Caption = Format(.sngTreatCal, "#,###") & " kcal"
    End With
End Sub

Private Sub InitialDietCalory()
    Dim i As Integer
    
    Select Case typObesity.sngTEE
        Case 1000 To 1249:      '1000~1200
            For i = 0 To 5
                imgDietCal(i).Visible = False
            Next i
        Case 1250 To 1449:      '1300~1400
            imgDietCal(0).Visible = False
            imgDietCal(1).Visible = True
            For i = 2 To 5
                imgDietCal(i).Visible = False
            Next i
        Case 1450 To 1749:      '1500~1700
            imgDietCal(0).Visible = False
            imgDietCal(1).Visible = True
            imgDietCal(2).Visible = False
            imgDietCal(3).Visible = True
            imgDietCal(4).Visible = False
            imgDietCal(5).Visible = False
        Case 1750 To 1949:      '1800~1900
            imgDietCal(0).Visible = False
            imgDietCal(1).Visible = True
            imgDietCal(2).Visible = False
            imgDietCal(3).Visible = True
            imgDietCal(4).Visible = True
            imgDietCal(5).Visible = False
        Case 1950 To 3500:      '2000~3500
            imgDietCal(0).Visible = False
            imgDietCal(1).Visible = True
            imgDietCal(2).Visible = False
            For i = 3 To 5
                imgDietCal(i).Visible = True
            Next i
        Case Else
            For i = 0 To 5
                imgDietCal(i).Visible = False
            Next i
    End Select
End Sub

'�񸸰��� ������ �����ֱ�
'ü��/�񸸵�/�㸮�ѷ�/WHR/BMI/ü�����/RMR/TEE
' �޽Ĵ�緮 8 : inRMR / 9 : etcRMR / �׹ۿ� : RMR
Private Sub ShowObesityRecord()
    Dim qrySelect As String, rValue As Variant
    Dim sngGab As Single
    Dim i As Integer
    
    Set clsSelect = New clsSelect
    
    qrySelect = "SELECT TOP 2 Weight, ObesityRate, Waist, b.WHR, BMI, ChFatRate,"
    If typCustomer.intAge >= ADULT_AGE Then
        qrySelect = qrySelect & "CASE AdBasicDsa "
    Else
        qrySelect = qrySelect & "CASE BaBasicDsa "
    End If
    qrySelect = qrySelect & "WHEN 8 THEN inRMR WHEN 9 THEN etcRMR ELSE RMR END RMR, TEE "
    qrySelect = qrySelect & "FROM Treat RIGHT JOIN BodyData AS b ON b.TreatNum=Treat.TreatNum "
    qrySelect = qrySelect & "INNER JOIN CompData AS c ON b.CompDataNum=c.CompDataNum "
    qrySelect = qrySelect & "WHERE Treat.CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & " ORDER BY Treat.TreatDay DESC;"
    
    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue) Then
        With typObesity
            .sngWeight = Is_Null(rValue(0, 0), 0)
            .sngObesityRate = Is_Null(rValue(1, 0), 0)
            .sngWaist = Is_Null(rValue(2, 0), 0)
            .sngWHR = Is_Null(rValue(3, 0), 0)
            .sngBMI = Is_Null(rValue(4, 0), 0)
            .sngChFatRate = Is_Null(rValue(5, 0), 0)
            .sngRMR = Is_Null(rValue(6, 0), 0)
            .sngTEE = Is_Null(rValue(7, 0), 0)
        End With
        typTreatCal.sngTreatCal = typObesity.sngTEE
        '1) ����
        For i = 0 To 7
            lblObInfo(i).Caption = Is_Null(rValue(i, 0), "-")
        Next i
        lblObInfo(0).Caption = lblObInfo(0).Caption & "kg"
        lblObInfo(1).Caption = CInt(lblObInfo(1).Caption) & "%"
        If lblObInfo(2).Caption <> "-" Then
            lblObInfo(2).Caption = lblObInfo(2).Caption & "cm"
        End If
        If lblObInfo(3).Caption <> "-" Then
            lblObInfo(3).Caption = Format(lblObInfo(3).Caption, "0.00")
        End If
        If lblObInfo(4).Caption <> "-" Then
            lblObInfo(4).Caption = Format(lblObInfo(4).Caption, "#.0")
        End If
        If lblObInfo(5).Caption <> "-" Then
            lblObInfo(5).Caption = Format(lblObInfo(5).Caption, "#.0") & "%"
        End If
        lblObInfo(6).Caption = Format(lblObInfo(6).Caption, "#,###")
        lblObInfo(7).Caption = Format(lblObInfo(7).Caption, "#,###")
        If UBound(rValue, 2) > 0 Then
            '2) ��ȸ���
            If Not IsNull(rValue(0, 1)) Then    '������
                sngGab = typObesity.sngWeight - rValue(0, 1)
                Call DrawUpDown(sngGab, 0, "kg")
            Else
                Set imgUpDown(0).Picture = LoadPicture("")
                lblUpDown(0).Caption = "-"
            End If
            If Not IsNull(rValue(1, 1)) Then    '�񸸵�
                sngGab = CInt(typObesity.sngObesityRate - rValue(1, 1))
                Call DrawUpDown(sngGab, 1, "%")
            Else
                Set imgUpDown(1).Picture = LoadPicture("")
                lblUpDown(1).Caption = "-"
            End If
            If Not IsNull(rValue(2, 1)) Then    '�㸮
                sngGab = typObesity.sngWaist - rValue(2, 1)
                Call DrawUpDown(sngGab, 2, "cm")
            Else
                Set imgUpDown(2).Picture = LoadPicture("")
                lblUpDown(2).Caption = "-"
            End If
            If Not IsNull(rValue(3, 1)) Then    'WHR
                sngGab = typObesity.sngWHR - rValue(3, 1)
                Call DrawUpDown(sngGab, 3, "", "0.00")
            Else
                Set imgUpDown(3).Picture = LoadPicture("")
                lblUpDown(3).Caption = "-"
            End If
            If Not IsNull(rValue(4, 1)) Then    'BMI
                sngGab = typObesity.sngBMI - rValue(4, 1)
                Call DrawUpDown(sngGab, 4, "")
                lblUpDown(4).Caption = Format(lblUpDown(4).Caption, "0.0")
            Else
                Set imgUpDown(4).Picture = LoadPicture("")
                lblUpDown(4).Caption = "-"
            End If
            If Not IsNull(rValue(5, 1)) Then    'ü������
                sngGab = typObesity.sngChFatRate - rValue(5, 1)
                
                Call DrawUpDown(Format(sngGab, "0.0"), 5, "%")
            Else
                Set imgUpDown(5).Picture = LoadPicture("")
                lblUpDown(5).Caption = "-"
            End If
            If Not IsNull(rValue(6, 1)) Then    'RMR(�޽Ĵ�緮)
                sngGab = typObesity.sngRMR - rValue(6, 1)
                Call DrawUpDown(sngGab, 6, "")
            Else
                Set imgUpDown(6).Picture = LoadPicture("")
                lblUpDown(6).Caption = "-"
            End If
            If Not IsNull(rValue(7, 1)) Then    'TEE (?? ��������?)
                sngGab = typObesity.sngTEE - rValue(7, 1)
                Call DrawUpDown(sngGab, 7, "")
            Else
                Set imgUpDown(7).Picture = LoadPicture("")
                lblUpDown(7).Caption = "-"
            End If
            '3) �ְ�/����
            'ü��
            lblMax(0).Caption = MaxValue("Weight") & "kg"
            lblMin(0).Caption = MinValue("Weight") & "kg"
            '�񸸵�
            lblMax(1).Caption = CInt(MaxValue("ObesityRate")) & "%"
            lblMin(1).Caption = CInt(MinValue("ObesityRate")) & "%"
            '�㸮�ѷ�
            lblMax(2).Caption = MaxValue("Waist") & "cm"
            lblMin(2).Caption = MinValue("Waist") & "cm"
            'WHR
            lblMax(3).Caption = Format(MaxValue("WHR"), "0.00")
            lblMin(3).Caption = Format(MinValue("WHR"), "0.00")
            'BMI
            lblMax(4).Caption = Format(MaxValue("BMI"), "0.0")
            lblMin(4).Caption = Format(MinValue("BMI"), "0.0")
            'ü�����
            lblMax(5).Caption = Format(MaxValue("ChFatRate"), "0.0") & "%"
            lblMin(5).Caption = Format(MinValue("ChFatRate"), "0.0") & "%"
            'RMR
            lblMax(6).Caption = Format(MaxValue("RMR"), "#,###")
            lblMin(6).Caption = Format(MinValue("RMR"), "#,###")
            'TEE
            lblMax(7).Caption = Format(MaxValue("TEE"), "#,###")
            lblMin(7).Caption = Format(MinValue("TEE"), "#,###")
        End If
    Else
        With typObesity
            .sngWeight = 0
            .sngObesityRate = 0
            .sngWaist = 0
            .sngWHR = 0
            .sngBMI = 0
            .sngChFatRate = 0
            .sngRMR = 0
            .sngTEE = 0
        End With
        typTreatCal.sngTreatCal = 0

        For i = 0 To 7
            lblObInfo(i).Caption = ""
            Set imgUpDown(i).Picture = LoadPicture("")
            lblUpDown(i).Caption = ""
            lblMax(i).Caption = ""
            lblMin(i).Caption = ""
        Next i
    End If
    
    Set clsSelect = Nothing
End Sub

Private Sub DrawUpDown(sngGab As Single, i As Integer, strUnit As String, Optional strFormat As String)
    If sngGab < 0 Then       '���� ȭ��ǥ �Ķ��� �۾�
        Set imgUpDown(i).Picture = LoadPicture(App.Path & IMG_DOWN)
        If strFormat <> "" Then
            lblUpDown(i).Caption = Format(sngGab, strFormat) & strUnit
        Else
            lblUpDown(i).Caption = sngGab & strUnit
        End If
    ElseIf sngGab > 0 Then   '���� ȭ��ǥ ������ �۾�
        Set imgUpDown(i).Picture = LoadPicture(App.Path & IMG_UP)
        If strFormat <> "" Then
            lblUpDown(i).Caption = Format(sngGab, strFormat) & strUnit
        Else
            lblUpDown(i).Caption = sngGab & strUnit
        End If
    Else                     '������ ����..��������
        Set imgUpDown(i).Picture = LoadPicture("")
        lblUpDown(i).Caption = "---"
    End If
End Sub

Private Sub imgAppend_Click(Index As Integer)
    frmPop_Additional.intCounselNo = Index
    Call frmPop_Additional.Show
End Sub

Private Sub imgDel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgDel.Picture = LoadPicture(App.Path & PATH01 & IMG_DEL_ON)
End Sub

Private Sub imgDel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgDel.Picture = LoadPicture(App.Path & PATH01 & IMG_DEL_OFF)
    
    If glngTreatNum = 0 Then
        MsgBox "������ ���᳻���� �����Ͻʽÿ�.", vbOKOnly + vbExclamation
        Exit Sub
    End If
    
    '��ü��������� �Ѱ����� ��� ����Ʈ ����
    If grdTable.Rows <= 3 Then
        If MsgBox("������ ������ �����Ͻø� �ڷᰡ �ϳ��� ���� �ʽ��ϴ�." & vbNewLine & "�׷��� �����Ͻðڽ��ϱ�?", vbYesNo + vbQuestion) = vbNo Then
            Exit Sub
        End If
    Else
        If gintBottomButton = 0 Then
            If MsgBox(grdTable.TextMatrix(grdTable.Row, 1) & " ���� ���᳻���� �����Ͻðڽ��ϱ�?" & vbNewLine & vbNewLine & "* ���� : �ش����� �˻��׸� �� Notice ��� �����Ͱ� �����˴ϴ� !", vbYesNo + vbQuestion) = vbNo Then
                Exit Sub
            End If
        ElseIf gintBottomButton = 2 Then
            If MsgBox(grdTable.TextMatrix(grdTable.Row, 0) & " ���� ���᳻���� ��� �����Ͻðڽ��ϱ�?" & vbNewLine & vbNewLine & "* ���� : �ش����� �˻��׸� �� Notice ��� �����Ͱ� �����˴ϴ� !", vbYesNo + vbQuestion) = vbNo Then
                Exit Sub
            End If
        End If
    End If

    'Treat,BodyData,TreatData,TreatPrint ����
    If DeleteTreat(glngTreatNum) = True Then
        Call Form_Load
        MsgBox "�����Ǿ����ϴ�.", vbOKOnly + vbInformation
    Else
        MsgBox "������ �����߽��ϴ�.", vbOKOnly + vbCritical
    End If
End Sub

Private Sub imgDietCal_Click(Index As Integer)
    Dim i As Integer
    
    AdminYn 11
    If AccessYn = False Then
        MsgBox "���ٱ����� �����ϴ�. �����ڿ��� �����Ͻʽÿ�.", vbOKOnly + vbInformation, "���ٱ��� ����"
        Exit Sub
    End If
    
    For i = 0 To 5
        Set imgDietCal(i).Picture = LoadPicture(App.Path & IMG_ME & i & "gray.jpg")
    Next i
    If idxDietCal = Index Then
        With typTreatCal
            If .intDietCal = 0 Then
                Set imgDietCal(Index).Picture = LoadPicture(App.Path & IMG_ME & Index & "green.jpg")
                idxDietCal = Index
            Else
                .intDietCal = 0
                .sngTreatCal = typObesity.sngTEE - .intDietCal
                lblTreatCal.Caption = Format(.sngTreatCal, "#,###") & " kcal"
                If .intDietCal >= 0 And .intExCal >= 0 And .intExDay >= 0 Then
                    .sngLossWeight = ((.intDietCal * 7) + (.intExCal * .intExDay)) / 7700 * 4
                    lblLossWeight.Caption = Format(.sngLossWeight, "0.0") & " kg/��"
                End If
                Exit Sub
            End If
        End With
    Else
        Set imgDietCal(Index).Picture = LoadPicture(App.Path & IMG_ME & Index & "green.jpg")
        idxDietCal = Index
    End If
    
    With typTreatCal
        Select Case Index
            Case 0: .intDietCal = 125
            Case 1: .intDietCal = 250
            Case 2: .intDietCal = 375
            Case 3: .intDietCal = 500
            Case 4: .intDietCal = 750
            Case 5: .intDietCal = 1000
            Case Else: .intDietCal = 0
        End Select
    
'���� says:
'����ü�߱��ϴ½�
'���� says:
'(�Ļ翭��*7 + �����*��ϼ�)/7700 * 4
    
        If .intDietCal >= 0 Then
            .sngTreatCal = typObesity.sngTEE - .intDietCal
            lblTreatCal.Caption = Format(.sngTreatCal, "#,###") & " kcal"
        End If
    '+=========================================
    '+ ����ü�� ���ϴ� ��
    '+-----------------------------------------
        If .intDietCal >= 0 And .intExCal >= 0 And .intExDay >= 0 Then
            .sngLossWeight = ((.intDietCal * 7) + (.intExCal * .intExDay)) / 7700 * 4
            lblLossWeight.Caption = Format(.sngLossWeight, "0.0") & " kg/��"
        End If
    End With
End Sub

Private Sub imgExCal_Click(Index As Integer)
    Dim i As Integer
    
    AdminYn 11
    If AccessYn = False Then
        MsgBox "���ٱ����� �����ϴ�. �����ڿ��� �����Ͻʽÿ�.", vbOKOnly + vbInformation, "���ٱ��� ����"
        Exit Sub
    End If
    
    For i = 0 To 3
        Set imgExCal(i).Picture = LoadPicture(App.Path & IMG_EX & i + 2 & "00-gray.jpg")
    Next i
    If idxExCal = Index Then
        With typTreatCal
            If .intExCal = 0 Then
                Set imgExCal(Index).Picture = LoadPicture(App.Path & IMG_EX & Index + 2 & "00-green.jpg")
                .intExCal = (Index + 2) * 100
                typExProgram.sngExCalory = (Index + 2) * 100
                intNowExCalory = .intExCal
                idxExCal = Index
            Else
                .intExCal = 0
                If chkUserCalory.Value = 0 Then
                    If .intDietCal >= 0 And .intExCal >= 0 And .intExDay >= 0 Then
                        .sngLossWeight = ((.intDietCal * 7) + (.intExCal * .intExDay)) / 7700 * 4
                        lblLossWeight.Caption = Format(.sngLossWeight, "0.0") & " kg/��"
                    End If
                Else
                    If .intUserCal >= 0 And .intExCal >= 0 And .intExCal >= 0 Then
                        .sngLossWeight = ((.intUserCal * 7) + (.intExCal * .intExDay)) / 7700 * 4
                        lblLossWeight.Caption = Format(.sngLossWeight, "0.0") & " kg/��"
                    End If
                End If
                Exit Sub
            End If
        End With
    Else
        Set imgExCal(Index).Picture = LoadPicture(App.Path & IMG_EX & Index + 2 & "00-green.jpg")
        typTreatCal.intExCal = (Index + 2) * 100
        typExProgram.sngExCalory = (Index + 2) * 100
        intNowExCalory = typTreatCal.intExCal
        idxExCal = Index
    End If
    
    With typTreatCal
        .intExCal = (Index + 2) * 100
        intNowExCalory = .intExCal
        
        If chkUserCalory.Value = 0 Then
            If .intDietCal >= 0 And .intExCal >= 0 And .intExDay >= 0 Then
                .sngLossWeight = ((.intDietCal * 7) + (.intExCal * .intExDay)) / 7700 * 4
                lblLossWeight.Caption = Format(.sngLossWeight, "0.0") & " kg/��"
            End If
        Else
            If .intUserCal >= 0 And .intExCal >= 0 And .intExCal >= 0 Then
                .sngLossWeight = ((.intUserCal * 7) + (.intExCal * .intExDay)) / 7700 * 4
                lblLossWeight.Caption = Format(.sngLossWeight, "0.0") & " kg/��"
            End If
        End If
    End With
End Sub

Private Sub imgExDay_Click(Index As Integer)
    Dim i As Integer
    
    AdminYn 11
    If AccessYn = False Then
        MsgBox "���ٱ����� �����ϴ�. �����ڿ��� �����Ͻʽÿ�.", vbOKOnly + vbInformation, "���ٱ��� ����"
        Exit Sub
    End If
    
    For i = 0 To 3
        Set imgExDay(i).Picture = LoadPicture(App.Path & IMG_EX & i + 3 & "��-gray.jpg")
    Next i
    If idxExDay = Index Then
        With typTreatCal
            If .intExDay = 0 Then
                Set imgExDay(Index).Picture = LoadPicture(App.Path & IMG_EX & Index + 3 & "��-green.jpg")
                .intExDay = Index + 3
                idxExDay = Index
            Else
                .intExDay = 0
                If chkUserCalory.Value = 0 Then
                    If .intDietCal >= 0 And .intExCal >= 0 And .intExDay >= 0 Then
                        .sngLossWeight = ((.intDietCal * 7) + (.intExCal * .intExDay)) / 7700 * 4
                        lblLossWeight.Caption = Format(.sngLossWeight, "0.0") & " kg/��"
                    End If
                Else
                    If .intUserCal >= 0 And .intExCal >= 0 And .intExCal >= 0 Then
                        .sngLossWeight = ((.intUserCal * 7) + (.intExCal * .intExDay)) / 7700 * 4
                        lblLossWeight.Caption = Format(.sngLossWeight, "0.0") & " kg/��"
                    End If
                End If
                Exit Sub
            End If
        End With
    Else
        Set imgExDay(Index).Picture = LoadPicture(App.Path & IMG_EX & Index + 3 & "��-green.jpg")
        typTreatCal.intExDay = Index + 3
        idxExDay = Index
    End If
    
    With typTreatCal
        .intExDay = Index + 3
        
        If chkUserCalory.Value = 0 Then
            If .intDietCal >= 0 And .intExCal >= 0 And .intExDay >= 0 Then
                .sngLossWeight = ((.intDietCal * 7) + (.intExCal * .intExDay)) / 7700 * 4
                lblLossWeight.Caption = Format(.sngLossWeight, "0.0") & " kg/��"
            End If
        Else
            If .intUserCal >= 0 And .intExCal >= 0 And .intExCal >= 0 Then
                .sngLossWeight = ((.intUserCal * 7) + (.intExCal * .intExDay)) / 7700 * 4
                lblLossWeight.Caption = Format(.sngLossWeight, "0.0") & " kg/��"
            End If
        End If
    End With
End Sub

Private Sub imgGoReserve_Click()
    AdminYn 10
    If AccessYn = False Then
        MsgBox "���ٱ����� �����ϴ�. �����ڿ��� �����Ͻʽÿ�.", vbOKOnly + vbInformation, "���ٱ��� ����"
        Exit Sub
    End If
   frm_ReserveInsert.Show vbModal
    Call InitialSpread4
    Call InitialSpread5
End Sub

Private Sub imgGoReserve_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgGoReserve.Picture = LoadPicture(App.Path & IMG_RES_OFF)
End Sub

Private Sub imgGoReserve_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgGoReserve.Picture = LoadPicture(App.Path & IMG_RES_ON)
End Sub

'+--------------------------------------------------
'+ �񸸵���
'+--------------------------------------------------
Private Sub Print_Obesity()
    Dim strfilename As String
    Dim qrySelect As String, rValue As Variant
            
On Error GoTo PrintErr
    qrySelect = "SELECT BodyDataNum FROM Treat RIGHT JOIN BodyData "
    qrySelect = qrySelect & "ON Treat.TreatNum=BodyData.TreatNum "
    qrySelect = qrySelect & "WHERE Treat.TreatNum=" & glngTreatNum
    Set clsSelect = New clsSelect
    rValue = clsSelect.Query(qrySelect)
    Set clsSelect = Nothing
    If IsNull(rValue) Then
        MsgBox "�񸸵��򰡸� ����� ������ �����ϴ�. �ٸ� �������� �����Ͻʽÿ�.", vbOKOnly + vbExclamation
        Exit Sub
    End If
    
    Call Prepare_OBMent(typCustomer.strSex)
    If typCustomer.intAge >= ADULT_AGE Then
        strfilename = "\Report\�񸸵�.rpt"
    Else
        If typCustomer.strSex = "M" Then
            strfilename = "\Report\�񸸵�MB.rpt"
        ElseIf typCustomer.strSex = "F" Then
            strfilename = "\Report\�񸸵�FB.rpt"
        End If
    End If
    Set crxReport = crxApplication.OpenReport(App.Path & strfilename)
    crxReport.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
    crxReport.PaperOrientation = crPortrait
    crxReport.RecordSelectionFormula = "{Treat.TreatNum}=" & glngTreatNum
        
    crxReport.Database.Tables(1).SetLogOnInfo strServer, strDBName, strUID, strPWD
    crxReport.PrintOut
   
    MsgBox "����� �Ϸ�Ǿ����ϴ�.", vbOKOnly + vbInformation, "���"
    Exit Sub
PrintErr:
    '2004-12-23 ������ �αױ��
    'WriteLog "Print_Obesity", "frmCounsel_1", Err.Number, Err.Description
    MsgBox "��¿� �����߽��ϴ�." & vbNewLine & Err.Number & Err.Description, vbOKOnly + vbCritical
End Sub

'�񸸵����� ��Ʈ�� �غ��ϴ� �Լ�
Private Sub Prepare_OBMent(strSex As String)
    Dim qrySelect As String
    Dim rValue As Variant
    Dim intSex As Integer

    If strSex = "M" Then
        intSex = 1
    Else
        intSex = 2
    End If
    Set clsSelect = New clsSelect
    
    '�񸸵� ��Ʈ�� �غ��Ѵ�.
    qrySelect = "SELECT ment FROM BodyData a, ObesityMent b "
    qrySelect = qrySelect & "WHERE a.ObesityRate >= b.ObesityRate1 "
    qrySelect = qrySelect & "AND a.ObesityRate < b.ObesityRate2 "
    qrySelect = qrySelect & "AND a.ChFatRate >= b.ChFatRate1 "
    qrySelect = qrySelect & "AND a.ChFatRate < b.ChFatRate2 "
    qrySelect = qrySelect & "AND ISNULL(a.WHR, 0) >= b.WHR1 "
    qrySelect = qrySelect & "AND ISNULL(a.WHR, 0) < b.WHR2 "
    qrySelect = qrySelect & "AND Sex=" & intSex

    qrySelect = qrySelect & " AND a.TreatNum=" & glngTreatNum

    rValue = clsSelect.Query(qrySelect)
        
    If Not IsNull(rValue) Then
        qrySelect = "DELETE FROM RPT_ObMent WHERE CustomerNum=" & glngCustomerNum
        Call modSql.AdoExcuteSql(qrySelect)
        
        qrySelect = "INSERT INTO RPT_ObMent (CustomerNum, Ment) "
        qrySelect = qrySelect & "VALUES(" & glngCustomerNum
        qrySelect = qrySelect & ",'" & rValue(0, 0) & "')"
        Call modSql.AdoExcuteSql(qrySelect)
        
        Erase rValue
    End If
End Sub

Private Sub imgNew_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgNew.Picture = LoadPicture(App.Path & PATH01 & IMG_NEW_OFF)
End Sub

Private Sub imgNew_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgNew.Picture = LoadPicture(App.Path & PATH01 & IMG_NEW_ON)
    lblDispDate.Caption = Format(gdatUserDay, "YYYY-MM-DD")
    intMode = 1
    EnabledInput (True)
    cmdPreTreat.Visible = True
    cmdInputTreat.Visible = False
End Sub

Private Sub imgPrint_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgPrint.Picture = LoadPicture(App.Path & PATH01 & IMG_PRINT_ON)
End Sub

Private Sub imgPrint_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgPrint.Picture = LoadPicture(App.Path & PATH01 & IMG_PRINT_OFF)

    If glngTreatNum = 0 Then
        MsgBox "����� ���� ����� �����ϼ���.", vbOKOnly + vbExclamation
        Exit Sub
    End If

    strServer = ServerName
'2005-01-18 ������ DB��������
    strDBName = DBinfo.DBName
    strUID = DBinfo.DBID
    strPWD = DBinfo.DBPWD
'    strDBName = "Body"
'    strUID = "sa"
'    strPWD = "1111"

''�̸�����
'    CrystalReport1.Destination = crptToWindow
''����Ʈ���
'    CrystalReport1.Destination = crptToPrinter
    Call Print_Obesity
End Sub

Private Sub imgSave_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgSave.Picture = LoadPicture(App.Path & PATH01 & IMG_SAVE_ON)
End Sub

Private Sub imgSave_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' * ���� : ��ü���� or ��ü����+óġ���� or óġ����
'�ԷµǾ� ������ �����ؾ��ϳ�?

    Dim qryUpdate As String
On Error GoTo SaveErr
'Ʈ����� �ɾ�߰���
    Set imgSave.Picture = LoadPicture(App.Path & PATH01 & IMG_SAVE_OFF)
    
    If typTreatCal.sngTreatCal = 0 Then
        MsgBox "Į�θ�ó���� �Ͻ��� �ٽ� �����Ͻʽÿ�.", vbOKOnly + vbExclamation
        Exit Sub
    End If
    
    If intMode = 1 Then    '�ű��Է¸��
        glngTreatNum = SaveTreat
        If glngTreatNum = 0 Then
            MsgBox "���忡 �����߽��ϴ�.", vbOKOnly + vbCritical
            Exit Sub
        Else
            If SaveTreatData = True Then
                If SaveTreatPrint = True Then
                    '�ϴ��� ǥ�� �׷��� ������Ʈ �� �� óġ�������� �����ش�.
            
                    'ġ�����α׷� ����
                    qryUpdate = "UPDATE CustomerInfo SET ChPgCode='" & typCustomer.strChPgCode
                    qryUpdate = qryUpdate & "' WHERE CustomerNum=" & glngCustomerNum
                    modSql.AdoExcuteSql (qryUpdate)
            
                    '�ϴ��� ǥ�� �׷��� ������Ʈ
                    Call imgViewHistory_Click(0)
                    MsgBox "����Ǿ����ϴ�.", vbOKOnly + vbInformation
                Else
                    MsgBox "���忡 �����߽��ϴ�.", vbOKOnly + vbCritical
                    Exit Sub
                End If
            Else
                MsgBox "���忡 �����߽��ϴ�.", vbOKOnly + vbCritical
                Exit Sub
            End If
        End If
    ElseIf intMode = 2 Then    '��ȸ/�������
        If glngTreatNum = 0 Then
            MsgBox "������ ���᳻���� �����Ͻʽÿ�.", vbOKOnly + vbInformation
            Exit Sub
        Else
            If UpdateTreat(glngTreatNum) = True Then
                If SaveTreatData = True Then
                    If SaveTreatPrint = True Then
                        'ġ�����α׷� ����
                        qryUpdate = "UPDATE CustomerInfo SET ChPgCode='" & typCustomer.strChPgCode
                        qryUpdate = qryUpdate & "' WHERE CustomerNum=" & glngCustomerNum
                        modSql.AdoExcuteSql (qryUpdate)
            
                        Call imgViewHistory_Click(0)
                        MsgBox "����Ǿ����ϴ�.", vbOKOnly + vbInformation
                    Else
                        MsgBox "���忡 �����߽��ϴ�.", vbOKOnly + vbCritical
                        Exit Sub
                    End If
                Else
                    MsgBox "���忡 �����߽��ϴ�.", vbOKOnly + vbCritical
                    Exit Sub
                End If
            Else
                MsgBox "���忡 �����߽��ϴ�.", vbOKOnly + vbCritical
                Exit Sub
            End If
        End If
    End If

    Exit Sub
SaveErr:
    '2004-12-23 ������ �αױ��
    'WriteLog "imgSave_MouseUp", "frmCounsel_1", Err.Number, Err.Description
    MsgBox "���忡 �����߽��ϴ�." & vbNewLine & vbNewLine & Err.Description, vbOKOnly + vbCritical
End Sub

Private Sub imgSelectEx_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgSelectEx.Picture = LoadPicture(App.Path & IMG_SELEX_OFF)
End Sub

Private Sub imgSelectEx_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim intProgram As Integer
    
    Call LoadExProgram(glngCustomerNum)
    intProgram = WhatExProgram(typExProgram.intAge, typExProgram.strSex, typExProgram.intObesity, typExProgram.intComplication, typExProgram.strBodyStatus)
    
    Set imgSelectEx.Picture = LoadPicture(App.Path & IMG_SELEX_ON)
    If WhatExConfig = True Then
        If intProgram = 9 Or intProgram = 23 Or intProgram = 37 Then
            MsgBox "������� �����Ͻ� �� �����ϴ�.", vbOKOnly + vbExclamation
            Exit Sub
        Else
            frmPop_ExAerobic.Show vbModal
        End If
    Else
        frmPop_ExAerobic2.Show vbModal
    End If
End Sub

Private Sub imgTreat_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgTreat.Picture = LoadPicture(App.Path & IMG_TREAT_OFF)
End Sub

Private Sub imgTreat_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgTreat.Picture = LoadPicture(App.Path & IMG_TREAT_ON)
    'ġ���̷� â ����
    frm_TreatDetail.Show vbModal
End Sub

'��ȭ��(ǥ),��ȭ��(�׷���),ó�泻����ȸ ���� Ŭ����.
Private Sub imgViewHistory_Click(Index As Integer)
    Dim i As Integer
On Error GoTo Err
    gintBottomButton = Index
    Select Case Index
        Case 0  '��ȭ��(ǥ)
            Set imgViewHistory(0).Picture = LoadPicture(App.Path & PATH01 & IMG_LTAB1_ON)
            Set imgViewHistory(1).Picture = LoadPicture(App.Path & PATH01 & IMG_LTAB2_OFF)
            Set imgViewHistory(2).Picture = LoadPicture(App.Path & PATH01 & IMG_LTAB3_OFF)
            
            Call Bottom0(0)
            imgTreat.Visible = False
        Case 1  '��ȭ��(�׷���)
            Set imgViewHistory(0).Picture = LoadPicture(App.Path & PATH01 & IMG_LTAB1_OFF)
            Set imgViewHistory(1).Picture = LoadPicture(App.Path & PATH01 & IMG_LTAB2_ON)
            Set imgViewHistory(2).Picture = LoadPicture(App.Path & PATH01 & IMG_LTAB3_OFF)
            
            Call Bottom1(0)
            imgTreat.Visible = False
        Case 2  'ó�泻����ȸ
            Set imgViewHistory(0).Picture = LoadPicture(App.Path & PATH01 & IMG_LTAB1_OFF)
            Set imgViewHistory(1).Picture = LoadPicture(App.Path & PATH01 & IMG_LTAB2_OFF)
            Set imgViewHistory(2).Picture = LoadPicture(App.Path & PATH01 & IMG_LTAB3_ON)
            
            Call Bottom2(0)
            imgTreat.Visible = True
    End Select
    Exit Sub
Err:
    '2004-12-23 ������ �αױ��
    'WriteLog "imgViewHistory_Click", "frmCounsel_1", Err.Number, Err.Description
    MsgBox Err.Number & Err.Description
End Sub

Private Sub imgTab_Click(Index As Integer)
    Select Case Index
        Case 0: 'ġ������/��¹� ��
            Set imgTab(0).Picture = LoadPicture(App.Path & PATH01 & IMG_RTAB1_ON)
            Set imgTab(1).Picture = LoadPicture(App.Path & PATH01 & IMG_RTAB2_OFF)
            Set imgTab(2).Picture = LoadPicture(App.Path & PATH01 & IMG_RTAB3_OFF)
            
            sprTreat.Visible = True
            sprPrint.Visible = True
            txtNotice.Visible = False
            txtMemo.Visible = False
        Case 1: 'Notice ��
            AdminYn 14
            If AccessYn = False Then
                MsgBox "���ٱ����� �����ϴ�. �����ڿ��� �����Ͻʽÿ�.", vbOKOnly + vbInformation, "���ٱ��� ����"
                Exit Sub
            End If

            Set imgTab(0).Picture = LoadPicture(App.Path & PATH01 & IMG_RTAB1_OFF)
            Set imgTab(1).Picture = LoadPicture(App.Path & PATH01 & IMG_RTAB2_ON)
            Set imgTab(2).Picture = LoadPicture(App.Path & PATH01 & IMG_RTAB3_OFF)
            
            txtNotice.Visible = True
            sprTreat.Visible = False
            sprPrint.Visible = False
            txtMemo.Visible = False
            
            If txtNotice.Enabled Then txtNotice.SetFocus
        Case 2: 'Memo��
            AdminYn 13
            If AccessYn = False Then
                MsgBox "���ٱ����� �����ϴ�. �����ڿ��� �����Ͻʽÿ�.", vbOKOnly + vbInformation, "���ٱ��� ����"
                Exit Sub
            End If
            
            Set imgTab(0).Picture = LoadPicture(App.Path & PATH01 & IMG_RTAB1_OFF)
            Set imgTab(1).Picture = LoadPicture(App.Path & PATH01 & IMG_RTAB2_OFF)
            Set imgTab(2).Picture = LoadPicture(App.Path & PATH01 & IMG_RTAB3_ON)
            
            txtMemo.Visible = True
            sprTreat.Visible = False
            sprPrint.Visible = False
            txtNotice.Visible = False
            If txtMemo.Enabled Then txtMemo.SetFocus
    End Select
End Sub

Private Sub imgTab_DblClick(Index As Integer)
    If Index = 2 Then
        AdminYn 13
        If AccessYn = False Then
            MsgBox "���ٱ����� �����ϴ�. �����ڿ��� �����Ͻʽÿ�.", vbOKOnly + vbInformation, "���ٱ��� ����"
            Exit Sub
        End If
        frmCounsel_1Pop.Show vbModal
    End If
End Sub

Private Sub sprTreat_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    If Col = 3 Then
        AdminYn 12
        If AccessYn = False Then
            MsgBox "���ٱ����� �����ϴ�. �����ڿ��� �����Ͻʽÿ�.", vbOKOnly + vbInformation, "���ٱ��� ����"
            Exit Sub
        End If
    End If
End Sub

Private Sub txtUserCalory_Change()
    Dim sngMinus As Single
    
    If txtUserCalory.Text = "" Then
        Exit Sub
    End If
    If IsNumeric(txtUserCalory.Text) = False Then
        Exit Sub
    End If

    If CInt(txtUserCalory.Text) > 3500 Then
        MsgBox "�Ĵܿ����� 1000kcal�̻� 3500kcal���Ϸ� �Է��Ͻʽÿ�.", vbOKOnly + vbExclamation
        txtUserCalory.SetFocus
        Exit Sub
    End If
    
    With typTreatCal

        .intUserCal = CInt(txtUserCalory.Text)
        If .intUserCal >= 0 Then

            .sngTreatCal = .intUserCal
            lblTreatCal.Caption = Format(.sngTreatCal, "#,###") & " kcal"
        End If
    '+=========================================
    '+ ����ü�� ���ϴ� ��
    '+-----------------------------------------
        sngMinus = typObesity.sngTEE - .intUserCal
        If sngMinus >= 0 And .intExCal >= 0 And .intExDay >= 0 Then
            .sngLossWeight = ((sngMinus * 7) + (.intExCal * .intExDay)) / 7700 * 4
            lblLossWeight.Caption = Format(.sngLossWeight, "0.0") & " kg/��"
        End If
    End With
End Sub

Private Sub txtUserCalory_LostFocus()
    Dim sngMinus As Single
    
    If txtUserCalory.Text = "" Then
        Exit Sub
    End If
    If IsNumeric(txtUserCalory.Text) = False Then
        Exit Sub
    End If
    With typTreatCal
        .intUserCal = CInt(txtUserCalory.Text)
        If .intUserCal >= 0 Then
            .sngTreatCal = .intUserCal
            lblTreatCal.Caption = Format(.sngTreatCal, "#,###") & " kcal"
        End If
        '+=========================================
        '+ ����ü�� ���ϴ� ��
        '+-----------------------------------------
        sngMinus = typObesity.sngTEE - .intUserCal
        If sngMinus >= 0 And .intExCal >= 0 And .intExDay >= 0 Then
            .sngLossWeight = ((sngMinus * 7) + (.intExCal * .intExDay)) / 7700 * 4
            lblLossWeight.Caption = Format(.sngLossWeight, "0.0") & " kg/��"
        End If
    End With
End Sub

'************ 2Fon ���� �߰�����(�򰡵�� �˾�)
Private Sub imgValuation_Click()
    Set imgTreat.Picture = LoadPicture(App.Path & IMG_��_ON)
    frmPop_Valuation.Show vbModal
End Sub

Private Sub imgValuation_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgValuation.Picture = LoadPicture(App.Path & IMG_��_ON)
End Sub

Private Sub imgValuation_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgValuation.Picture = LoadPicture(App.Path & IMG_��_OFF)
End Sub

