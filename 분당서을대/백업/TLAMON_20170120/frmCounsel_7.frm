VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form frmCounsel_7 
   BorderStyle     =   0  '����
   Caption         =   "Form1"
   ClientHeight    =   9105
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13425
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9105
   ScaleWidth      =   13425
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtCaution 
      Appearance      =   0  '���
      BorderStyle     =   0  '����
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1400
      Left            =   2010
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  '����
      TabIndex        =   20
      Top             =   3210
      Width           =   3000
   End
   Begin MSFlexGridLib.MSFlexGrid grdExImages 
      Height          =   3105
      Index           =   0
      Left            =   720
      TabIndex        =   6
      Top             =   5400
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   5477
      _Version        =   393216
      BackColorBkg    =   16777215
      Appearance      =   0
   End
   Begin MSFlexGridLib.MSFlexGrid grdExImages 
      Height          =   3100
      Index           =   1
      Left            =   720
      TabIndex        =   7
      Top             =   5400
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   5477
      _Version        =   393216
      BackColorBkg    =   16777215
      Appearance      =   0
   End
   Begin MSFlexGridLib.MSFlexGrid grdExImages 
      Height          =   3100
      Index           =   2
      Left            =   720
      TabIndex        =   8
      Top             =   5400
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   5477
      _Version        =   393216
      BackColorBkg    =   16777215
      Appearance      =   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdHealth 
      Height          =   5355
      Left            =   930
      TabIndex        =   18
      Top             =   2340
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   9446
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lblComment 
      Alignment       =   1  '������ ����
      BackStyle       =   0  '����
      Caption         =   "* ǥ�� ����Ŭ���Ͻø� ū �̹����� ���� �� �ֽ��ϴ�."
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF80FF&
      Height          =   195
      Left            =   6480
      TabIndex        =   21
      Top             =   5130
      Width           =   4185
   End
   Begin VB.Image TopImage 
      Height          =   960
      Left            =   -30
      Picture         =   "frmCounsel_7.frx":0000
      Top             =   50
      Width           =   13140
   End
   Begin VB.Label lblActProgram 
      BackStyle       =   0  '����
      Caption         =   "�ⱸ����α׷�"
      Height          =   225
      Left            =   2070
      TabIndex        =   19
      Top             =   1680
      Width           =   3105
   End
   Begin VB.Image imgPage 
      Height          =   345
      Index           =   1
      Left            =   10320
      Top             =   2820
      Width           =   345
   End
   Begin VB.Image imgPage 
      Height          =   345
      Index           =   0
      Left            =   9900
      Top             =   2820
      Width           =   345
   End
   Begin VB.Image imgBaby 
      Height          =   1020
      Left            =   10710
      Picture         =   "frmCounsel_7.frx":1ACE
      Top             =   5370
      Width           =   210
   End
   Begin VB.Image imgShift 
      Height          =   225
      Left            =   11010
      Top             =   8310
      Width           =   285
   End
   Begin VB.Image imgP2 
      Height          =   4830
      Left            =   990
      Top             =   3270
      Width           =   9735
   End
   Begin VB.Image imgP3T 
      Height          =   2070
      Left            =   960
      Top             =   2820
      Width           =   9735
   End
   Begin VB.Image imgP3Tab 
      Height          =   345
      Index           =   1
      Left            =   2640
      Top             =   2490
      Width           =   1665
   End
   Begin VB.Image imgP3Tab 
      Height          =   345
      Index           =   0
      Left            =   960
      Top             =   2490
      Width           =   1665
   End
   Begin VB.Image imgPreg 
      Height          =   330
      Index           =   3
      Left            =   5340
      Top             =   1350
      Width           =   1545
   End
   Begin VB.Image imgPreg 
      Height          =   330
      Index           =   2
      Left            =   3780
      Top             =   1350
      Width           =   1545
   End
   Begin VB.Image imgPreg 
      Height          =   330
      Index           =   1
      Left            =   2220
      Top             =   1350
      Width           =   1545
   End
   Begin VB.Image imgPreg 
      Height          =   330
      Index           =   0
      Left            =   660
      Top             =   1350
      Width           =   1545
   End
   Begin VB.Image imgSlim 
      Height          =   330
      Index           =   2
      Left            =   3780
      Top             =   1770
      Width           =   1545
   End
   Begin VB.Image imgSlim 
      Height          =   330
      Index           =   1
      Left            =   2220
      Top             =   1770
      Width           =   1545
   End
   Begin VB.Image imgSlim 
      Height          =   330
      Index           =   0
      Left            =   660
      Top             =   1770
      Width           =   1545
   End
   Begin VB.Label lblName 
      Alignment       =   1  '������ ����
      BackStyle       =   0  '����
      Caption         =   "�μ���"
      Height          =   225
      Left            =   870
      TabIndex        =   17
      Top             =   1290
      Width           =   585
   End
   Begin VB.Label lblSub 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "������������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   3
      Left            =   9600
      TabIndex        =   16
      Top             =   4230
      Width           =   1000
   End
   Begin VB.Label lblSub 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "������������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   2
      Left            =   8500
      TabIndex        =   15
      Top             =   4230
      Width           =   1000
   End
   Begin VB.Label lblSub 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "������������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   7380
      TabIndex        =   14
      Top             =   4230
      Width           =   1000
   End
   Begin VB.Label lblSub 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "������������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   6250
      TabIndex        =   13
      Top             =   4230
      Width           =   1000
   End
   Begin VB.Image imgSub 
      Height          =   945
      Index           =   3
      Left            =   9630
      Top             =   3210
      Width           =   945
   End
   Begin VB.Image imgSub 
      Height          =   945
      Index           =   2
      Left            =   8520
      Top             =   3210
      Width           =   945
   End
   Begin VB.Image imgSub 
      Height          =   945
      Index           =   1
      Left            =   7410
      Top             =   3210
      Width           =   945
   End
   Begin VB.Image imgSub 
      Height          =   945
      Index           =   0
      Left            =   6300
      Top             =   3210
      Width           =   945
   End
   Begin VB.Label lblMuscle3 
      BackStyle       =   0  '����
      Caption         =   "������Ÿ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   2040
      TabIndex        =   12
      Top             =   3390
      Width           =   7995
   End
   Begin VB.Label lblMuscle2 
      BackStyle       =   0  '����
      Caption         =   "������Ÿ��"
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
      Left            =   2040
      TabIndex        =   11
      Top             =   3000
      Width           =   2565
   End
   Begin VB.Label lblMuscleTime 
      BackStyle       =   0  '����
      Caption         =   "������Ÿ��"
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
      Left            =   2040
      TabIndex        =   10
      Top             =   2640
      Width           =   4000
   End
   Begin VB.Label lblMuscle1 
      BackStyle       =   0  '����
      Caption         =   "������Ÿ��"
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
      Left            =   2040
      TabIndex        =   9
      Top             =   2250
      Width           =   3000
   End
   Begin VB.Label lblIntensity2 
      BackStyle       =   0  '����
      Caption         =   "(�ɹڼ� :"
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
      Left            =   6300
      TabIndex        =   5
      Top             =   2910
      Width           =   3500
   End
   Begin VB.Label lblIntensity 
      BackStyle       =   0  '����
      Caption         =   "�ణ ����ٰ� �������� ������ ��Ѵ�."
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
      Left            =   6300
      TabIndex        =   4
      Top             =   2700
      Width           =   4500
   End
   Begin VB.Label lblTime 
      BackStyle       =   0  '����
      Caption         =   "40 ~ 50 ��"
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
      Left            =   2040
      TabIndex        =   3
      Top             =   2730
      Width           =   3000
   End
   Begin VB.Label lblOften 
      BackStyle       =   0  '����
      Caption         =   "3�� / ��"
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
      Left            =   6270
      TabIndex        =   2
      Top             =   2250
      Width           =   4500
   End
   Begin VB.Label lblMain 
      BackStyle       =   0  '����
      Caption         =   "������Ÿ��"
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
      Left            =   2040
      TabIndex        =   1
      Top             =   2250
      Width           =   2565
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  '����
      Caption         =   "�μ��� ���� 3���� ����� �, 3���� �ٷ¿�� �Ͻʽÿ�."
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   900
      TabIndex        =   0
      Top             =   1290
      Width           =   8115
   End
   Begin VB.Image Image2 
      Height          =   180
      Left            =   690
      Picture         =   "frmCounsel_7.frx":2332
      Stretch         =   -1  'True
      Top             =   1290
      Width           =   165
   End
   Begin VB.Image imgPrint 
      Height          =   1065
      Left            =   11190
      Picture         =   "frmCounsel_7.frx":273D
      Top             =   7140
      Width           =   1065
   End
   Begin VB.Image imgTab 
      Height          =   975
      Index           =   2
      Left            =   10710
      Picture         =   "frmCounsel_7.frx":3DB9
      Top             =   7320
      Width           =   210
   End
   Begin VB.Image imgTab 
      Height          =   975
      Index           =   1
      Left            =   10710
      Picture         =   "frmCounsel_7.frx":4561
      Top             =   6360
      Width           =   210
   End
   Begin VB.Image imgTopTab 
      Height          =   330
      Index           =   1
      Left            =   2250
      Top             =   1650
      Width           =   1575
   End
   Begin VB.Image imgTab 
      Height          =   975
      Index           =   0
      Left            =   10710
      Picture         =   "frmCounsel_7.frx":4D55
      Top             =   5370
      Width           =   210
   End
   Begin VB.Image imgTopTab 
      Height          =   330
      Index           =   0
      Left            =   660
      Top             =   1650
      Width           =   1575
   End
End
Attribute VB_Name = "frmCounsel_7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'+---------------------------------------------------------------------------------+
'| ��� > � > ����̵� / �Ϲ� �� ��ȯ�� �ִ� ���
'+---------------------------------------------------------------------------------+
Private Const IMG_BACK_1 As String = "\Back\Counsel\07\���_����̵� back.jpg"
Private Const IMG_BACK_2 As String = "\Back\Counsel\07\����̵�2 back.jpg"
Private Const IMG_TOPTAB0_ON As String = "\Back\Counsel\07\����ҿon.jpg"
Private Const IMG_TOPTAB0_OFF As String = "\Back\Counsel\07\����ҿoff.jpg"
Private Const IMG_TOPTAB1_ON As String = "\Back\Counsel\07\�ٷ¿on.jpg"
Private Const IMG_TOPTAB1_OFF As String = "\Back\Counsel\07\�ٷ¿off.jpg"
Private Const IMG_TAB0_ON As String = "\Back\Counsel\07\subTab\�ٷ¿1 on.jpg"
Private Const IMG_TAB0_OFF As String = "\Back\Counsel\07\subTab\�ٷ¿1 off.jpg"
Private Const IMG_TAB1_ON As String = "\Back\Counsel\07\subTab\�ٷ¿2 on.jpg"
Private Const IMG_TAB1_OFF As String = "\Back\Counsel\07\subTab\�ٷ¿2 off.jpg"
Private Const IMG_TAB2_ON As String = "\Back\Counsel\07\subTab\�ٷ¿3 on.jpg"
Private Const IMG_TAB2_OFF As String = "\Back\Counsel\07\subTab\�ٷ¿3 off.jpg"
'+---------------------------------------------------------------------------------+
'| ��� > � > ����̵� / �Ҿ�
'+---------------------------------------------------------------------------------+
Private Const IMG_BABY1 As String = "\Back\Counsel\07\subTab\���彺Ʈ��Ī on.jpg"
Private Const IMG_BABY2 As String = "\Back\Counsel\07\subTab\�Ϲݽ�Ʈ��Ī on.jpg"
'+---------------------------------------------------------------------------------+
'| ��� > � > ����̵� / ����
'+---------------------------------------------------------------------------------+
Private Const IMG_PREG_1 As String = "\Back\Counsel\07\����\���_����̵�(����)up.jpg"     '�����
Private Const IMG_PREG_2 As String = "\Back\Counsel\07\����\���_����̵����2��.jpg" '���2~3��
Private Const IMG_PREG_2SUB1 As String = "\Back\Counsel\07\����\���_����̵���� 2�� �׸�1.jpg"
Private Const IMG_PREG_2SUB2 As String = "\Back\Counsel\07\����\���_����̵���� 2�� �׸�2.jpg"
Private Const IMG_PAGE1_ON As String = "\Back\Counsel\07\����\������������ư1 on.jpg"
Private Const IMG_PAGE1_OFF As String = "\Back\Counsel\07\����\������������ư1 off.jpg"
Private Const IMG_PAGE2_ON As String = "\Back\Counsel\07\����\������������ư2 on.jpg"
Private Const IMG_PAGE2_OFF As String = "\Back\Counsel\07\����\������������ư2 off.jpg"
Private Const IMG_PREG_3 As String = "\Back\Counsel\07\����\����4�� back.jpg"                '���4~9��
Private Const IMG_PREG_3TAB1 As String = "\Back\Counsel\07\����\����4�� ��Ʈ��Ī.jpg"   '   > ��Ʈ��Ī
Private Const IMG_PREG_3TAB1ON As String = "\Back\Counsel\07\����\��Ʈ��Īon.jpg"
Private Const IMG_PREG_3TAB1OFF As String = "\Back\Counsel\07\����\��Ʈ��Īoff.jpg"
Private Const IMG_PREG_3TAB2 As String = "\Back\Counsel\07\����\����4�� ����ҿ.jpg" '   > ����ҿ
Private Const IMG_PREG_3TAB2ON As String = "\Back\Counsel\07\����\����ҿon.jpg"
Private Const IMG_PREG_3TAB2OFF As String = "\Back\Counsel\07\����\����ҿoff.jpg"
Private Const IMG_PREG_4 As String = "\Back\Counsel\07\����\���� 10��.jpg"                  '���10~12��

Private Const IMG_PREGTAB1_ON As String = "\Back\Counsel\07\����� on.jpg"
Private Const IMG_PREGTAB1_OFF As String = "\Back\Counsel\07\����� off.jpg"
Private Const IMG_PREGTAB2_ON As String = "\Back\Counsel\07\���2~3���� on.jpg"
Private Const IMG_PREGTAB2_OFF As String = "\Back\Counsel\07\���2~3���� off.jpg"
Private Const IMG_PREGTAB3_ON As String = "\Back\Counsel\07\���4~9���� on.jpg"
Private Const IMG_PREGTAB3_OFF As String = "\Back\Counsel\07\���4~9���� off.jpg"
Private Const IMG_PREGTAB4_ON As String = "\Back\Counsel\07\���10~12�� on.jpg"
Private Const IMG_PREGTAB4_OFF As String = "\Back\Counsel\07\���10~12�� off.jpg"
'+---------------------------------------------------------------------------------+
'| ��� > � > ����̵� / ��ü��
'+---------------------------------------------------------------------------------+
Private Const IMG_SLIM_1 As String = "\Back\Counsel\07\���_����̵� ��ü�� ����(����).jpg"
Private Const IMG_SLIM_2 As String = "\Back\Counsel\07\���_����̵� ��ü�� ����(��ü).jpg"
Private Const IMG_SLIM_3 As String = "\Back\Counsel\07\���_����̵� ��ü�� ����(��ü).jpg"
Private Const IMG_SLIMTAB1_ON As String = "\Back\Counsel\07\�ٷ¿1 ON.jpg"
Private Const IMG_SLIMTAB1_OFF As String = "\Back\Counsel\07\�ٷ¿1 OFF.jpg"
Private Const IMG_SLIMTAB2_ON As String = "\Back\Counsel\07\�ٷ¿2 ON.jpg"
Private Const IMG_SLIMTAB2_OFF As String = "\Back\Counsel\07\�ٷ¿2 OFF.jpg"
Private Const IMG_SLIMTAB3_ON As String = "\Back\Counsel\07\�ٷ¿3 ON.jpg"
Private Const IMG_SLIMTAB3_OFF As String = "\Back\Counsel\07\�ٷ¿3 OFF.jpg"

Private Const PATH07 As String = "\Back\Counsel\07\"
Private Const IMG_PRINT_ON As String = "����̵���� on.jpg"
Private Const IMG_PRINT_OFF As String = "����̵���� off.jpg"
Private Const IMG_USEREX As String = "����ڿ.jpg"
Private intSelProgram As Integer

Public Sub Form_Load()
    Set Me.Picture = LoadPicture(App.Path & "\Back\Counsel\07\����ڿ.jpg")
    Me.Top = FRM_TOP
    Me.Left = FRM_LEFT
    Me.Width = FRM_WIDTH
    Me.Height = FRM_HEIGHT
    Me.BackColor = vbWhite
    
' ȯ�漳���� ���� �ٵ��÷��� ������ ��� ���,  ����� ��� ��쿡 ���� �ٸ��� ����
' �ٵ��÷��� ���� �
    With typExProgram
        .strCustomName = ""
        .intAge = 0
        .strSex = ""
        .strBodyStatus = ""
        .sngHeight = 0
        .sngWeight = 0
        .intObesity = 0
        .intComplication = 0
        .strDietPart = ""
        .sngExCalory = 0
        .intExDay = 0
        .sngBMI = 0
    End With
'1) ������ �Էµ� ���� �ҷ��´�
    Call LoadExProgram(glngCustomerNum)
    lblComment.Visible = False
    If WhatExConfig = False Then
        Call ShowUserEx
        Call VisibleFalse
        Call VisibleFalse_Slim
        Call VisibleFalse_Preg
        imgBaby.Visible = False
'        Set imgPrint.Picture = LoadPicture("")
        Set imgPrint.Picture = LoadPicture(App.Path & PATH07 & IMG_PRINT_OFF)
        Exit Sub
    End If
'///////////////////////////////////////////////////////////////////// Start ~!
'[*] ���࿡ �ʿ��� ���� ���� ��쿡�� �Է��� �޴´�.
'    ����α׷��� ���ؼ��� �� �ʿ�������
'    �ʼ��Է°��� �ƴϾ��� ���� �Է¹��� �� �ִ� ��Ʈ�� �ʿ�
    lblActProgram.Visible = False
    grdHealth.Visible = False
    If typExProgram.sngExCalory = 0 Then
        MsgBox "����̵带 ���÷��� �Į�θ��� ���� ó���Ͻʽÿ�.", vbOKOnly + vbInformation
        Set Me.Picture = LoadPicture("")
        Call VisibleFalse
        Call VisibleFalse_Slim
        Call VisibleFalse_Preg
        imgBaby.Visible = False
        Set imgPrint.Picture = LoadPicture("")
        Exit Sub
    End If

    With typExProgram
    If .intObesity = 0 Then
        MsgBox "����α׷��� �� �� ���� �����̰ų�(5���̸�)" & vbNewLine & vbNewLine & "���ʵ����Ͱ� ����Ȯ�ϹǷ� ����α׷� ���ÿ� �����߽��ϴ�.", vbOKOnly + vbExclamation
        Call VisibleFalse
        Call VisibleFalse_Slim
        Call VisibleFalse_Preg
        imgBaby.Visible = False
        Set imgPrint.Picture = LoadPicture("")
        Exit Sub
    End If
'[*] 1. � ��� �� ������ ���� ������.
    intSelProgram = WhatExProgram(.intAge, .strSex, .intObesity, .intComplication, .strBodyStatus)
    End With
    imgBaby.Visible = False
    imgTopTab(0).Enabled = True
    imgTopTab(1).Enabled = True
    Select Case intSelProgram
        Case 1 To 4
            Call VisibleFalse_Slim
            Call VisibleFalse_Preg
            Call InitialControl
            Call imgTopTab_Click(0)
            imgTopTab(0).Enabled = False
            imgTopTab(1).Enabled = False
            Set imgTopTab(0).Picture = LoadPicture(App.Path & IMG_TOPTAB0_ON)
            Set imgTopTab(1).Picture = LoadPicture("")
            Call DrawAerobic(intSelProgram)    '����ҿ ��� ǥ
            Call DrawAnaerobic(intSelProgram)  '�ٷ¿ ��� ǥ
            Call CompositeForm(intSelProgram)  '�ٷ¿ �ϴ� ǥ
            imgBaby.Visible = True
            Set imgBaby.Picture = LoadPicture(App.Path & IMG_BABY1)
            lblTitle.Visible = True
            Image2.Visible = True
            lblComment.Visible = True
        Case 5 To 6
            Call VisibleFalse_Slim
            Call VisibleFalse_Preg
            Call InitialControl
            Call imgTopTab_Click(0)
            imgTopTab(0).Enabled = False
            imgTopTab(1).Enabled = False
            Set imgTopTab(0).Picture = LoadPicture(App.Path & IMG_TOPTAB0_ON)
            Set imgTopTab(1).Picture = LoadPicture("")
            Call DrawAerobic(intSelProgram)    '����ҿ ��� ǥ
            Call DrawAnaerobic(intSelProgram)  '�ٷ¿ ��� ǥ
            Call CompositeForm(intSelProgram)  '�ٷ¿ �ϴ� ǥ
            imgBaby.Visible = True
            Set imgBaby.Picture = LoadPicture(App.Path & IMG_BABY2)
            lblTitle.Visible = True
            Image2.Visible = True
            lblComment.Visible = True
        Case 7  '����
            Call VisibleFalse
            Call VisibleFalse_Slim
            Call PregSetting
        Case 8   '��ü��
            Call VisibleFalse
            Call VisibleFalse_Preg
            Call SlimSetting
            lblName.Caption = WhatName
            lblName.Visible = True
        Case 9, 23, 37   '��ȯ�� ���� ��� �� �Ϻ�
            '����ҿ ó�� ����
            '�ٷ¿�� ó����
            Call VisibleFalse_Slim
            Call VisibleFalse_Preg
            Call InitialControl
            Call imgTopTab_Click(1)
            imgTopTab(0).Enabled = False
            imgTopTab(1).Enabled = False
            Set imgTopTab(0).Picture = LoadPicture(App.Path & IMG_TOPTAB1_ON)
            Set imgTopTab(1).Picture = LoadPicture("")
            Call imgTab_Click(0)
            lblTitle.Visible = True
            Image2.Visible = True
            lblTitle.Caption = WhatName & " ���� " & typExProgram.intExDay & "�� �ٷ¿�� �Ͻʽÿ�."
            Call DrawAnaerobic(intSelProgram)  '�ٷ¿ ��� ǥ
            Call CompositeForm(intSelProgram)  '�ٷ¿ �ϴ� ǥ
            lblComment.Visible = True
        Case 51     '��ȯ��, ����, �񸸵�����
            '����ҿ ó����
            '�ٷ¿ ó�� ����
            Call VisibleFalse_Slim
            Call VisibleFalse_Preg
            Call InitialControl
            Call imgTopTab_Click(0)
            imgTopTab(0).Enabled = False
            imgTopTab(1).Enabled = False
            Set imgTopTab(0).Picture = LoadPicture(App.Path & IMG_TOPTAB0_ON)
            Set imgTopTab(1).Picture = LoadPicture("")
            Set imgTab(0).Picture = LoadPicture("")
            Set imgTab(1).Picture = LoadPicture("")
            Set imgTab(2).Picture = LoadPicture("")
            grdExImages(0).Visible = False
            grdExImages(1).Visible = False
            grdExImages(2).Visible = False
            lblTitle.Visible = True
            Image2.Visible = True
            lblTitle.Caption = WhatName & " ���� " & typExProgram.intExDay & "�� ����ҿ�� �Ͻʽÿ�."
            Call DrawAerobic(intSelProgram)
        Case Else
            Call VisibleFalse_Slim
            Call VisibleFalse_Preg
            Call InitialControl
            Call imgTopTab_Click(0)
            Call imgTab_Click(0)
            lblTitle.Visible = True
            Image2.Visible = True
            Call DrawAerobic(intSelProgram)    '����ҿ ��� ǥ
            Call DrawAnaerobic(intSelProgram)  '�ٷ¿ ��� ǥ
            Call CompositeForm(intSelProgram)  '�ٷ¿ �ϴ� ǥ
            lblTitle.Visible = True
            Image2.Visible = True
            lblComment.Visible = True
    End Select
    Set imgPrint.Picture = LoadPicture(App.Path & PATH07 & IMG_PRINT_OFF)
End Sub

Private Sub ShowUserEx()
    Dim qrySelect As String, rValue As Variant
    Dim i As Integer
    
    Set Me.Picture = LoadPicture(App.Path & PATH07 & IMG_USEREX)
    lblActProgram.Visible = True
    grdHealth.Visible = True
    Call InitialGrid
    qrySelect = "SELECT ActProgramName, HealthActName, ActSet FROM CustomerInfo INNER JOIN ActProgram "
    qrySelect = qrySelect & "ON CustomerInfo.ActProgram=ActProgram.ActProgramNum "
    qrySelect = qrySelect & "INNER JOIN ActDetail "
    qrySelect = qrySelect & "ON ActProgram.ActProgramNum=ActDetail.ActProgramNum "
    qrySelect = qrySelect & "INNER JOIN HealthAct ON ActDetail.HealthActNo=HealthAct.HealthActNo "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    Set clsSelect = New clsSelect
    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue) Then
        lblActProgram.Caption = Is_Null(rValue(0, 0), "")
        With grdHealth
        .Clear
        .RowS = 0
        For i = 0 To UBound(rValue, 2)
            .RowS = .RowS + 1
            .TextMatrix(i, 0) = i + 1
            .TextMatrix(i, 1) = Trim(Is_Null(rValue(1, i), ""))
            .TextMatrix(i, 2) = Trim(Is_Null(rValue(2, i), ""))
        Next i
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignCenterCenter
        .ColAlignment(2) = flexAlignCenterCenter
        .RowHeight(-1) = 300
        End With
    End If
End Sub

Private Sub InitialGrid()
    With grdHealth
        .Clear
        .SelectionMode = flexSelectionByRow
        .FocusRect = flexFocusNone
        .Appearance = flexFlat
        .BorderStyle = flexBorderNone
        .RowS = 0
        .ColS = 3
        .FixedCols = 0
        .FixedRows = 0
        .BackColorBkg = vbWhite
        
        .ColWidth(0) = 940
        .ColWidth(1) = 3000
        .ColWidth(2) = 2300
        
        '�׸����� �� ����..
        .GridColor = FRM_GRAY
        .GridLineWidth = 3
        
    End With
End Sub

Public Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub grdExImages_DblClick(Index As Integer)
    frmPop_Counsel_7.intProgram = intSelProgram
    frmPop_Counsel_7.Show vbModal
End Sub

Private Sub imgP3Tab_Click(Index As Integer)
    imgP3T.Visible = True
    imgP3Tab(0).Visible = True
    imgP3Tab(1).Visible = True
    Select Case Index
        Case 0:
            Set imgP3T.Picture = LoadPicture(App.Path & IMG_PREG_3TAB1)
            Set imgP3Tab(0).Picture = LoadPicture(App.Path & IMG_PREG_3TAB1ON)
            Set imgP3Tab(1).Picture = LoadPicture(App.Path & IMG_PREG_3TAB2OFF)
        Case 1:
            Set imgP3T.Picture = LoadPicture(App.Path & IMG_PREG_3TAB2)
            Set imgP3Tab(0).Picture = LoadPicture(App.Path & IMG_PREG_3TAB1OFF)
            Set imgP3Tab(1).Picture = LoadPicture(App.Path & IMG_PREG_3TAB2ON)
    End Select
End Sub

Private Sub imgPage_Click(Index As Integer)
    imgPage(0).Visible = True
    imgPage(1).Visible = True
    
    If Index = 0 Then
        Set imgPage(0).Picture = LoadPicture(App.Path & IMG_PAGE1_ON)
        Set imgPage(1).Picture = LoadPicture(App.Path & IMG_PAGE2_OFF)
        Set imgP2.Picture = LoadPicture(App.Path & IMG_PREG_2SUB1)
    Else
        Set imgPage(0).Picture = LoadPicture(App.Path & IMG_PAGE1_OFF)
        Set imgPage(1).Picture = LoadPicture(App.Path & IMG_PAGE2_ON)
        Set imgP2.Picture = LoadPicture(App.Path & IMG_PREG_2SUB2)
    End If
End Sub

Private Sub imgPreg_Click(Index As Integer)
    Select Case Index
        Case 0:     '�����
            Set Me.Picture = LoadPicture(App.Path & IMG_PREG_1)
            
            imgP3T.Visible = False
            imgP3Tab(0).Visible = False
            imgP3Tab(1).Visible = False
            imgP2.Visible = False
            imgPage(0).Visible = False
            imgPage(1).Visible = False

            Set imgPreg(0).Picture = LoadPicture(App.Path & IMG_PREGTAB1_ON)
            Set imgPreg(1).Picture = LoadPicture(App.Path & IMG_PREGTAB2_OFF)
            Set imgPreg(2).Picture = LoadPicture(App.Path & IMG_PREGTAB3_OFF)
            Set imgPreg(3).Picture = LoadPicture(App.Path & IMG_PREGTAB4_OFF)
        Case 1:     '���2~3��
            Set Me.Picture = LoadPicture(App.Path & IMG_PREG_2)
            
            imgP3T.Visible = False
            imgP3Tab(0).Visible = False
            imgP3Tab(1).Visible = False
            imgP2.Visible = True
            Call imgPage_Click(0)
            
            Set imgPreg(0).Picture = LoadPicture(App.Path & IMG_PREGTAB1_OFF)
            Set imgPreg(1).Picture = LoadPicture(App.Path & IMG_PREGTAB2_ON)
            Set imgPreg(2).Picture = LoadPicture(App.Path & IMG_PREGTAB3_OFF)
            Set imgPreg(3).Picture = LoadPicture(App.Path & IMG_PREGTAB4_OFF)
        Case 2:     '���4~9��
            Set Me.Picture = LoadPicture(App.Path & IMG_PREG_3)
            
            imgP2.Visible = False
            imgPage(0).Visible = False
            imgPage(1).Visible = False
            Call imgP3Tab_Click(0)
            
            Set imgPreg(0).Picture = LoadPicture(App.Path & IMG_PREGTAB1_OFF)
            Set imgPreg(1).Picture = LoadPicture(App.Path & IMG_PREGTAB2_OFF)
            Set imgPreg(2).Picture = LoadPicture(App.Path & IMG_PREGTAB3_ON)
            Set imgPreg(3).Picture = LoadPicture(App.Path & IMG_PREGTAB4_OFF)
        Case 3:     '���10~12��
            Set Me.Picture = LoadPicture(App.Path & IMG_PREG_4)
            
            imgP3T.Visible = False
            imgP3Tab(0).Visible = False
            imgP3Tab(1).Visible = False
            imgP2.Visible = False
            imgPage(0).Visible = False
            imgPage(1).Visible = False
            
            Set imgPreg(0).Picture = LoadPicture(App.Path & IMG_PREGTAB1_OFF)
            Set imgPreg(1).Picture = LoadPicture(App.Path & IMG_PREGTAB2_OFF)
            Set imgPreg(2).Picture = LoadPicture(App.Path & IMG_PREGTAB3_OFF)
            Set imgPreg(3).Picture = LoadPicture(App.Path & IMG_PREGTAB4_ON)
    End Select
End Sub

Private Sub imgPrint_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgPrint.Picture = LoadPicture(App.Path & PATH07 & IMG_PRINT_ON)
End Sub

Private Sub imgPrint_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgPrint.Picture = LoadPicture(App.Path & PATH07 & IMG_PRINT_OFF)
    ' ����̵� ���
End Sub

Private Sub imgSlim_Click(Index As Integer)
    Select Case Index
        Case 0:
            Set Me.Picture = LoadPicture(App.Path & IMG_SLIM_1)
    
            Set imgSlim(0).Picture = LoadPicture(App.Path & IMG_SLIMTAB1_ON)
            Set imgSlim(1).Picture = LoadPicture(App.Path & IMG_SLIMTAB2_OFF)
            Set imgSlim(2).Picture = LoadPicture(App.Path & IMG_SLIMTAB3_OFF)
        Case 1:
            Set Me.Picture = LoadPicture(App.Path & IMG_SLIM_2)
            
            Set imgSlim(0).Picture = LoadPicture(App.Path & IMG_SLIMTAB1_OFF)
            Set imgSlim(1).Picture = LoadPicture(App.Path & IMG_SLIMTAB2_ON)
            Set imgSlim(2).Picture = LoadPicture(App.Path & IMG_SLIMTAB3_OFF)
        Case 2:
            Set Me.Picture = LoadPicture(App.Path & IMG_SLIM_3)
            
            Set imgSlim(0).Picture = LoadPicture(App.Path & IMG_SLIMTAB1_OFF)
            Set imgSlim(1).Picture = LoadPicture(App.Path & IMG_SLIMTAB2_OFF)
            Set imgSlim(2).Picture = LoadPicture(App.Path & IMG_SLIMTAB3_ON)
    End Select
End Sub

Private Sub imgTab_Click(Index As Integer)
    Select Case Index
        Case 0:
            Set imgTab(0).Picture = LoadPicture(App.Path & IMG_TAB0_ON)
            Set imgTab(1).Picture = LoadPicture(App.Path & IMG_TAB1_OFF)
            Set imgTab(2).Picture = LoadPicture(App.Path & IMG_TAB2_OFF)
            
            grdExImages(0).Visible = True
            grdExImages(1).Visible = False
            grdExImages(2).Visible = False
        Case 1:
            Set imgTab(0).Picture = LoadPicture(App.Path & IMG_TAB0_OFF)
            Set imgTab(1).Picture = LoadPicture(App.Path & IMG_TAB1_ON)
            Set imgTab(2).Picture = LoadPicture(App.Path & IMG_TAB2_OFF)
            
            grdExImages(0).Visible = False
            grdExImages(1).Visible = True
            grdExImages(2).Visible = False
        Case 2:
            Set imgTab(0).Picture = LoadPicture(App.Path & IMG_TAB0_OFF)
            Set imgTab(1).Picture = LoadPicture(App.Path & IMG_TAB1_OFF)
            Set imgTab(2).Picture = LoadPicture(App.Path & IMG_TAB2_ON)
            
            grdExImages(0).Visible = False
            grdExImages(1).Visible = False
            grdExImages(2).Visible = True
    End Select
End Sub

Private Sub imgTopTab_Click(Index As Integer)
    Dim i As Integer
    imgTopTab(0).Visible = True
    imgTopTab(1).Visible = True
    If Index = 0 Then
        Set Me.Picture = LoadPicture(App.Path & IMG_BACK_1)
        Set imgTopTab(0).Picture = LoadPicture(App.Path & IMG_TOPTAB0_ON)
        Set imgTopTab(1).Picture = LoadPicture(App.Path & IMG_TOPTAB1_OFF)
        lblMuscle1.Visible = False
        lblMuscleTime.Visible = False
        lblMuscle2.Visible = False
        lblMuscle3.Visible = False
        
        lblMain.Visible = True
        lblOften.Visible = True
        lblTime.Visible = True
        lblIntensity.Visible = True
        txtCaution.Visible = True
        For i = 0 To 3
            imgSub(i).Visible = True
            lblSub(i).Visible = True
        Next i
    Else
        Set Me.Picture = LoadPicture(App.Path & IMG_BACK_2)
        Set imgTopTab(0).Picture = LoadPicture(App.Path & IMG_TOPTAB0_OFF)
        Set imgTopTab(1).Picture = LoadPicture(App.Path & IMG_TOPTAB1_ON)
        lblMain.Visible = False
        lblOften.Visible = False
        lblTime.Visible = False
        lblIntensity.Visible = False
        txtCaution.Visible = False
        For i = 0 To 3
            imgSub(i).Visible = False
            lblSub(i).Visible = False
        Next i
        
        lblMuscle1.Visible = True
        lblMuscleTime.Visible = True
        lblMuscle2.Visible = True
        lblMuscle3.Visible = True
    End If
End Sub

'+----------------------------------------------------------------
'+ �ʱ�ȭ
'+----------------------------------------------------------------
Private Sub InitialControl()
    Dim i As Integer
    
    lblTitle.Caption = ""
    
    lblMain.Caption = ""
    lblOften.Caption = ""
    lblTime.Caption = ""
    lblIntensity.Caption = ""
    lblIntensity2.Caption = ""
    txtCaution.Text = ""
    
    For i = 0 To 3
        imgSub(i).Picture = LoadPicture("")
        lblSub(i).Caption = ""
    Next i
    
    lblMuscle1.Caption = ""
    lblMuscleTime.Caption = ""
    lblMuscle2.Caption = ""
    lblMuscle3.Caption = ""
    
    For i = 0 To 2
    With grdExImages(i)
        .Clear
        .ScrollBars = flexScrollBarHorizontal
        .BorderStyle = flexBorderNone
        .GridColor = vbWhite
        .GridLineWidth = 3
        
        .WordWrap = True
        .RowS = 2
        .FixedCols = 0
        .FixedRows = 0
        .RowHeight(0) = 2000
        .RowHeight(1) = 870
    End With
    Next i

End Sub

Private Sub VisibleFalse()
    Dim i As Integer
    
    Image2.Visible = False
    lblTitle.Visible = False
    imgTopTab(0).Visible = False
    imgTopTab(1).Visible = False
    lblMain.Visible = False
    lblOften.Visible = False
    lblTime.Visible = False
    lblIntensity.Visible = False
    lblIntensity2.Visible = False
    txtCaution.Visible = False
    For i = 0 To 3
        imgSub(i).Visible = False
        lblSub(i).Visible = False
    Next i
    lblMuscle1.Visible = False
    lblMuscleTime.Visible = False
    lblMuscle2.Visible = False
    lblMuscle3.Visible = False
    
    For i = 0 To 2
        grdExImages(i).Visible = False
        imgTab(i).Visible = False
    Next i
End Sub

Private Sub VisibleFalse_Slim()
    Dim i As Integer
    
    lblName.Visible = False
    
    For i = 0 To 2
        imgSlim(i).Visible = False
    Next i
End Sub

Private Sub VisibleFalse_Preg()
    Dim i As Integer
    
    For i = 0 To 3
        imgPreg(i).Visible = False
    Next i
    imgP3T.Visible = False
    imgP3Tab(0).Visible = False
    imgP3Tab(1).Visible = False
    imgP2.Visible = False
    imgPage(0).Visible = False
    imgPage(1).Visible = False
End Sub

Private Sub SlimSetting()
    Dim i As Integer
    
    Call imgSlim_Click(0)
    
    For i = 0 To 2
        imgSlim(i).Visible = True
    Next i
End Sub

Private Sub PregSetting()
    Dim i As Integer
    
    Call imgPreg_Click(0)
    
    For i = 0 To 3
        imgPreg(i).Visible = True
    Next i
End Sub

Private Sub CompositeForm(intProgram As Integer)
    Dim i As Integer
    '////////////////// �Ʒ����ʹ� ����� �ٷ¿/��Ʈ��Ī
    Select Case intProgram
        Case 1 To 4
            Call DrawExImages(0, 12, 1)
            imgTab(0).Visible = False
            imgTab(1).Visible = False
            imgTab(2).Visible = False
        Case 5 To 6
            Call DrawExImages(0, 12, 13)
            imgTab(0).Visible = False
            imgTab(1).Visible = False
            imgTab(2).Visible = False
        Case 7                        '����
        
        Case 8                         '��ü��
        Case 9, 16, 23, 30, 37, 44   '��ȯ ����
            Call DrawExImages(0, 8, 119)
            imgTab(0).Visible = True
            imgTab(1).Visible = False
            imgTab(2).Visible = False
       Case 10, 17, 24, 31, 38, 45    '�索��
            Call DrawExImages(0, 9, 41)
            Call DrawExImages(1, 9, 50)
            Call DrawExImages(2, 9, 59)
            imgTab(0).Visible = True
            imgTab(1).Visible = True
            imgTab(2).Visible = True
       Case 11, 18, 25, 32, 39, 46    '��������
            Call DrawExImages(0, 9, 41)
            Call DrawExImages(1, 9, 50)
            Call DrawExImages(2, 9, 59)
            imgTab(0).Visible = True
            imgTab(1).Visible = True
            imgTab(2).Visible = True
        Case 12, 19, 26, 33, 40, 47    '������
            Call DrawExImages(0, 9, 68)
            Call DrawExImages(1, 7, 77)
            imgTab(0).Visible = True
            imgTab(1).Visible = True
            imgTab(2).Visible = False
        Case 13, 20, 27, 34, 41, 48     'ô����ȯ = �����а� ����
            Call DrawExImages(0, 9, 68)
            Call DrawExImages(1, 7, 77)
            imgTab(0).Visible = True
            imgTab(1).Visible = True
            imgTab(2).Visible = False '
        Case 14, 21, 28, 35, 42, 49     '������ȯ
            Call DrawExImages(0, 9, 109)
            imgTab(0).Visible = True
            imgTab(1).Visible = False
            imgTab(2).Visible = False
        Case 15, 22, 29, 36, 43, 50   '��ٰ���
            Call DrawExImages(0, 9, 41)
            Call DrawExImages(1, 9, 50)
            Call DrawExImages(2, 9, 59)
            imgTab(0).Visible = True
            imgTab(1).Visible = True
            imgTab(2).Visible = True
        Case Else
            Call DrawExImages(0, 9, 41)
            Call DrawExImages(1, 9, 50)
            Call DrawExImages(2, 9, 59)
            imgTab(0).Visible = True
            imgTab(1).Visible = True
            imgTab(2).Visible = True
    End Select
End Sub

'����� � : ǥ�� �ٲ��� �� �ִ� � ǥ�� �����ֱ�
Private Sub DrawAerobic(intProgram As Integer)
    Dim qrySelect As String, rValue As Variant, rValue2 As Variant
    Dim sngAerobic As Single, sngAnaerobic As Single
    Dim intExTime As Integer, intLen As Integer, intCol As Integer, strTemp As String
    Dim i As Integer
    Dim sngFactor As Single, sngFactor2 As Single
    
    Set clsSelect = New clsSelect

On Error GoTo ShowErr
    '�ϴ� ó��� ������� �������� �ҷ����µ�, ���� ó��� ���� ���������� �׳� tblExMethod�� ����� �� �ҷ�����
    qrySelect = "SELECT TOP 1 main, sub1, sub2, sub3, sub4 FROM Treat "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & " AND main IS NOT NULL ORDER BY TreatDay DESC;"
    
    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue) Then
        '�����(������)
        lblMain.Caption = WhatExName(CInt(rValue(0, 0)))
        sngFactor = WhatFactor(CInt(rValue(0, 0)))
        
        '��ü����(������)
    Else
        qrySelect = "SELECT main, ExName, factor, sub "
        qrySelect = qrySelect & "FROM tblExMethod a LEFT JOIN tblExercise b "

        qrySelect = qrySelect & "ON a.main=b.ExNo "
        qrySelect = qrySelect & "WHERE gno=" & intProgram
        
        rValue = clsSelect.Query(qrySelect)
        If Not IsNull(rValue) Then
            '�����(������)
            lblMain.Caption = Trim(Is_Null(rValue(1, 0), ""))
            sngFactor = Is_Null(rValue(2, 0), 0)
            
            '��ü����(������)
            '�׿� �   /�� ���е� ������ ������� ������
        End If
    End If

    Set clsSelect = New clsSelect
    qrySelect = "SELECT often, intensity, caution, factor "
    qrySelect = qrySelect & "FROM tblExMethod a LEFT JOIN tblExerciseItem b "
    qrySelect = qrySelect & "ON a.main=b.xno "
    qrySelect = qrySelect & "WHERE gno=" & intProgram

    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue) Then
    
    '��ܿ� ��Ʈ �ֱ�
    If typExProgram.intExDay = 0 Then
        '���
        lblOften.Caption = Trim(Is_Null(rValue(0, 0), ""))
    Else
        If intProgram < 7 Then
            lblTitle.Caption = WhatName & " ���� ����ҿ�� ��Ʈ��Ī�� �����Ͻʽÿ�."
            lblOften.Caption = typExProgram.intExDay & "��/��"
        ElseIf intProgram = 9 Or intProgram = 23 Or intProgram = 37 Then
            lblTitle.Caption = WhatName & " ���� �ٷ¿�� ��Ʈ��Ī�� �����Ͻʽÿ�."
            lblOften.Caption = typExProgram.intExDay & "��/��"
        ElseIf intProgram = 30 Or typExProgram.intComplication = 2 Then
            lblTitle.Caption = WhatName & " ���� ����ҿ�� �ٷ¿�� �����Ͻʽÿ�."
            lblOften.Caption = typExProgram.intExDay & "��/��(�ٷ¿�� ����)"
        ElseIf intProgram = 51 Then
            lblTitle.Caption = WhatName & " ���� ����ҿ�� ��Ʈ��Ī�� �����Ͻʽÿ�."
            lblOften.Caption = typExProgram.intExDay & "��/��"
        Else
            Select Case typExProgram.intExDay
                Case 5:
                    lblTitle.Caption = WhatName & " ���� 3���� ����ҿ�� �ٷ¿�� �����ϰ�, 2���� ����ҿ�� �Ͻʽÿ�."
                    lblOften.Caption = "2��/��"
                Case 6:
                    lblTitle.Caption = WhatName & " ���� 3���� ����ҿ, 3���� �ٷ¿�� �Ͻʽÿ�."
                    lblOften.Caption = "3��/��"
                Case Else
                    lblTitle.Caption = WhatName & " ���� ����ҿ�� �ٷ¿�� �����Ͻʽÿ�."
                    lblOften.Caption = typExProgram.intExDay & "��/��(�ٷ¿�� ����)"
            End Select
        End If
    End If
    
    '�����
    lblIntensity.Caption = Trim(Is_Null(rValue(1, 0), ""))
    '���ǻ���
    txtCaution.Text = Trim(Is_Null(rValue(2, 0), ""))
    Dim intAfterTime As Integer
    
    Select Case intProgram
        Case 16:
            sngAerobic = typExProgram.sngExCalory
            intExTime = Int(sngAerobic / (typExProgram.sngWeight * sngFactor))
            If typExProgram.intExDay = 6 Then
                lblTime.Caption = intExTime & " ��"
            ElseIf typExProgram.intExDay = 5 Then
                intAfterTime = Int((sngAerobic - (typExProgram.sngWeight * 0.105 * 30)) / (typExProgram.sngWeight * sngFactor))
                lblTime.Caption = intExTime & " �� (�ٷ¿�� �����ϴ� ���� " & intAfterTime & " ��)"
            Else
                intAfterTime = Int((sngAerobic - (typExProgram.sngWeight * 0.105 * 30)) / (typExProgram.sngWeight * sngFactor))
                lblTime.Caption = intAfterTime & " ��"
            End If
        Case 30:
            sngAerobic = typExProgram.sngExCalory
            intAfterTime = Int((sngAerobic - (typExProgram.sngWeight * 0.105 * 30)) / (typExProgram.sngWeight * sngFactor))
            lblTime.Caption = intAfterTime & " ��"
            lblOften.Caption = typExProgram.intExDay & "��/��(�ٷ¿�� ����)"
        Case 44:
            sngAerobic = typExProgram.sngExCalory
            intExTime = Int(sngAerobic / (typExProgram.sngWeight * sngFactor))
            If typExProgram.intExDay = 6 Then
                lblTime.Caption = intExTime & " ��"
            ElseIf typExProgram.intExDay = 5 Then
                intAfterTime = Int((sngAerobic - (typExProgram.sngWeight * 0.105 * 30)) / (typExProgram.sngWeight * sngFactor))
                lblTime.Caption = intExTime & " �� (�ٷ¿�� �����ϴ� ���� " & intAfterTime & " ��)"
            Else
                intAfterTime = Int((sngAerobic - (typExProgram.sngWeight * 0.105 * 30)) / (typExProgram.sngWeight * sngFactor))
                lblTime.Caption = intAfterTime & " ��"
            End If
        Case 51:
            sngAerobic = typExProgram.sngExCalory
            intExTime = Int(sngAerobic / (typExProgram.sngWeight * sngFactor))
            lblTime.Caption = intExTime & " ��"
        Case 10, 17, 24, 31, 38, 45:
            sngAerobic = typExProgram.sngExCalory
            intExTime = Int(sngAerobic / (typExProgram.sngWeight * sngFactor))
            If typExProgram.intExDay = 6 Then
                lblTime.Caption = intExTime & " ��"
            ElseIf typExProgram.intExDay = 5 Then
                intAfterTime = Int((sngAerobic - (typExProgram.sngWeight * 0.105 * 30)) / (typExProgram.sngWeight * sngFactor))
                lblTime.Caption = intExTime & " �� (�ٷ¿�� �����ϴ� ���� " & intAfterTime & " ��)"
            Else
                intAfterTime = Int((sngAerobic - (typExProgram.sngWeight * 0.105 * 30)) / (typExProgram.sngWeight * sngFactor))
                lblTime.Caption = intAfterTime & " ��"
            End If
        Case 11, 18, 25, 32, 39, 46    '��������
        '2) �ٷ¿ + ����ҿ(�����ȱ� ����)
        '  -> �ٷ¿ 30��(������)���� �Ҹ��ϴ� Į�θ� ���
            sngAerobic = typExProgram.sngExCalory
            intAfterTime = Int((sngAerobic - (typExProgram.sngWeight * 0.105 * 30)) / (typExProgram.sngWeight * sngFactor))
            lblTime.Caption = intAfterTime & " ��"
            
            lblTitle.Caption = WhatName & " ���� ����ҿ�� �ٷ¿�� �����Ͻʽÿ�."
            lblOften.Caption = typExProgram.intExDay & "��/��(�ٷ¿�� ����)"
        Case 12, 19, 26, 33, 40, 47     '������
            sngAerobic = typExProgram.sngExCalory
            intExTime = Int(sngAerobic / (typExProgram.sngWeight * sngFactor))
            If typExProgram.intExDay = 6 Then
                lblTime.Caption = intExTime & " ��"
            ElseIf typExProgram.intExDay = 5 Then
                intAfterTime = Int((sngAerobic - (typExProgram.sngWeight * 0.105 * 30)) / (typExProgram.sngWeight * sngFactor))
                lblTime.Caption = intExTime & " �� (�ٷ¿�� �����ϴ� ���� " & intAfterTime & " ��)"
            Else
                intAfterTime = Int((sngAerobic - (typExProgram.sngWeight * 0.105 * 30)) / (typExProgram.sngWeight * sngFactor))
                lblTime.Caption = intAfterTime & " ��"
            End If
        Case 13, 20, 27, 34, 41, 48     'ô����ȯ
            sngAerobic = typExProgram.sngExCalory
            intExTime = Int(sngAerobic / (typExProgram.sngWeight * sngFactor))
            If typExProgram.intExDay = 6 Then
                lblTime.Caption = intExTime & " ��"
            ElseIf typExProgram.intExDay = 5 Then
                intAfterTime = Int((sngAerobic - (typExProgram.sngWeight * 0.105 * 30)) / (typExProgram.sngWeight * sngFactor))
                lblTime.Caption = intExTime & " �� (�ٷ¿�� �����ϴ� ���� " & intAfterTime & " ��)"
            Else
                intAfterTime = Int((sngAerobic - (typExProgram.sngWeight * 0.105 * 30)) / (typExProgram.sngWeight * sngFactor))
                lblTime.Caption = intAfterTime & " ��"
            End If
        Case 14, 21, 28, 35, 42, 49     '������ȯ
        '1) ��Ʈ��Ī + ����ҿ
        '  -> ��Ʈ��Ī 20��(������)
            sngAerobic = typExProgram.sngExCalory
            intExTime = Int(sngAerobic / (typExProgram.sngWeight * sngFactor))
            If typExProgram.intExDay = 6 Then
                lblTime.Caption = intExTime & " ��"
            ElseIf typExProgram.intExDay = 5 Then
                intAfterTime = Int((sngAerobic - (typExProgram.sngWeight * 0.105 * 30)) / (typExProgram.sngWeight * sngFactor))
                lblTime.Caption = intExTime & " �� (�ٷ¿�� �����ϴ� ���� " & intAfterTime & " ��)"
            Else
                intAfterTime = Int((sngAerobic - (typExProgram.sngWeight * 0.105 * 30)) / (typExProgram.sngWeight * sngFactor))
                lblTime.Caption = intAfterTime & " ��"
            End If
        Case 15, 22, 29, 36, 43, 50   '��ٰ���
            sngAerobic = typExProgram.sngExCalory
            intExTime = Int(sngAerobic / (typExProgram.sngWeight * sngFactor))
            If typExProgram.intExDay = 6 Then
                lblTime.Caption = intExTime & " ��"
            ElseIf typExProgram.intExDay = 5 Then
                intAfterTime = Int((sngAerobic - (typExProgram.sngWeight * 0.105 * 30)) / (typExProgram.sngWeight * sngFactor))
                lblTime.Caption = intExTime & " �� (�ٷ¿�� �����ϴ� ���� " & intAfterTime & " ��)"
            Else
                intAfterTime = Int((sngAerobic - (typExProgram.sngWeight * 0.105 * 30)) / (typExProgram.sngWeight * sngFactor))
                lblTime.Caption = intAfterTime & " ��"
            End If
        Case Else
            sngAerobic = typExProgram.sngExCalory
            If sngFactor <> 0 Then
            intExTime = Int(sngAerobic / (typExProgram.sngWeight * sngFactor))
            lblTime.Caption = intExTime & " ��"
            End If
    End Select
    
    Set clsSelect = New clsSelect

    '�ϴ� ó��� ������� �������� �ҷ����µ�, ���� ó��� ���� ���������� �׳� tblExMethod�� ����� �� �ҷ�����
    qrySelect = "SELECT TOP 1 main, sub1, sub2, sub3, sub4 FROM Treat "
    qrySelect = qrySelect & "WHERE CustomerNum=" & glngCustomerNum
    qrySelect = qrySelect & " AND main IS NOT NULL ORDER BY TreatDay DESC;"
    
    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue) Then
        '�����(������)
        
        '��ü����(������)
        If Not IsNull(rValue(1, 0)) Then
            Set imgSub(0).Picture = LoadPicture(App.Path & "\Back\Ex\Ys\" & CInt(rValue(1, 0)) & ".jpg")
            lblSub(0).Caption = WhatExName(CInt(rValue(1, 0)))
            sngFactor2 = WhatFactor(CInt(rValue(1, 0)))
            intExTime = Int(sngAerobic / (typExProgram.sngWeight * sngFactor2))
            lblSub(0).Caption = lblSub(0).Caption & vbNewLine & intExTime & " ��"
        End If
        If Not IsNull(rValue(2, 0)) Then
            Set imgSub(1).Picture = LoadPicture(App.Path & "\Back\Ex\Ys\" & CInt(rValue(2, 0)) & ".jpg")
            lblSub(1).Caption = WhatExName(CInt(rValue(2, 0)))
            sngFactor2 = WhatFactor(CInt(rValue(2, 0)))
            intExTime = Int(sngAerobic / (typExProgram.sngWeight * sngFactor2))
            lblSub(1).Caption = lblSub(1).Caption & vbNewLine & intExTime & " ��"
        End If
        If Not IsNull(rValue(3, 0)) Then
            Set imgSub(2).Picture = LoadPicture(App.Path & "\Back\Ex\Ys\" & CInt(rValue(3, 0)) & ".jpg")
            lblSub(2).Caption = WhatExName(CInt(rValue(3, 0)))
            sngFactor2 = WhatFactor(CInt(rValue(3, 0)))
            intExTime = Int(sngAerobic / (typExProgram.sngWeight * sngFactor2))
            lblSub(2).Caption = lblSub(2).Caption & vbNewLine & intExTime & " ��"
        End If
        If Not IsNull(rValue(4, 0)) Then
            Set imgSub(3).Picture = LoadPicture(App.Path & "\Back\Ex\Ys\" & CInt(rValue(4, 0)) & ".jpg")
            lblSub(3).Caption = WhatExName(CInt(rValue(4, 0)))
            sngFactor2 = WhatFactor(CInt(rValue(4, 0)))
            intExTime = Int(sngAerobic / (typExProgram.sngWeight * sngFactor2))
            lblSub(3).Caption = lblSub(3).Caption & vbNewLine & intExTime & " ��"
        End If
    Else
        qrySelect = "SELECT main, ExName, factor, sub "
        qrySelect = qrySelect & "FROM tblExMethod a LEFT JOIN tblExercise b "
        qrySelect = qrySelect & "ON a.main=b.ExNo "
        qrySelect = qrySelect & "WHERE gno=" & intProgram
        
        rValue = clsSelect.Query(qrySelect)
        If Not IsNull(rValue) Then
            '�����(������)
            lblMain.Caption = Trim(Is_Null(rValue(1, 0), ""))
            sngFactor = Is_Null(rValue(2, 0), 0)
            
            '��ü����(������)
            '�׿� �   /�� ���е� ������ ������� ������
            intLen = Len(Trim(Is_Null(rValue(3, 0), ""))) + 1
            intLen = intLen / 3
            intCol = 0
            For i = 0 To intLen - 1
                strTemp = CInt(Mid(Trim(rValue(3, 0)), (i * 3) + 1, 2))
                qrySelect = "SELECT ExName, Factor FROM tblExercise WHERE ExNo=" & strTemp
                rValue2 = clsSelect.Query(qrySelect)
                If Not IsNull(rValue2) Then
                    lblSub(intCol).Caption = Trim(rValue2(0, 0))
                        If rValue2(1, 0) <> 0 Then
                            intExTime = Int(sngAerobic / (typExProgram.sngWeight * CSng(rValue2(1, 0))))
                            If intExTime <= 20 Then
                                intExTime = 20
                            End If
                        Else
                            intExTime = 20
                        End If
                        lblSub(intCol).Caption = lblSub(intCol).Caption & vbNewLine & intExTime & " ��"
                End If
                Set imgSub(intCol).Picture = LoadPicture(App.Path & "\Back\Ex\Ys\" & strTemp & ".jpg")
                intCol = intCol + 1
                If intCol = 4 Then
                    Exit Sub
                End If
            Next i
        End If
    End If
    
    End If
    
    Exit Sub
ShowErr:
    '2004-12-23 ������ �αױ��
    'WriteLog "DrawAerobic", "frmCounsel_7", Err.Number, Err.Description
    If Err.Number = 53 Then   '�̹��� ������ ���� ���
        Err.Clear
        Resume Next
    Else
        MsgBox Err.Number & Err.Description
    End If
End Sub

'����� � : �ٷ¿ ��� ǥ
Private Sub DrawAnaerobic(intProgram As Integer)
    Dim qrySelect As String, rValue As Variant
    Dim intExTime As Integer
    
    Set clsSelect = New clsSelect
    qrySelect = "SELECT muscle1, muscle2, muscle3 FROM tblExMethod "
    qrySelect = qrySelect & "WHERE gno=" & intProgram
    
    rValue = clsSelect.Query(qrySelect)
    If Not IsNull(rValue) Then
        lblMuscle1.Caption = Is_Null(rValue(0, 0), "")
       '��ð�
        With typExProgram
        Select Case intProgram
            Case 9:
                intExTime = Int(.sngExCalory / (.sngWeight * 0.105))
                lblMuscleTime.Caption = intExTime & " ��"
            Case 23, 37:
                intExTime = Int(.sngExCalory - 30 * 0.105 * .sngWeight) / (.sngWeight * 0.093)
                lblMuscleTime.Caption = "30 �� (�ٷ¿ �� �����ȱ� " & intExTime & " ��)"
            Case 16:
                If .intExDay = 6 Then
                    intExTime = Int(.sngExCalory - 30 * 0.105 * .sngWeight) / (.sngWeight * 0.093)
                    lblMuscleTime.Caption = "30 �� (�ٷ¿ �� �����ȱ� " & intExTime & " ��)"
                Else
                    lblMuscleTime.Caption = "30 ��"
                End If
            Case 30:
                lblMuscleTime.Caption = "30 ��"
            Case 44:
                If .intExDay = 6 Then
                    intExTime = Int(.sngExCalory / (.sngWeight * 0.105))
                    lblMuscleTime.Caption = intExTime & " ��"
                Else
                    lblMuscleTime.Caption = "30 ��"
                End If
            Case 51:
            
            Case 10, 17, 24, 31, 38, 45:
                If .intExDay = 6 Then
                    intExTime = Int(.sngExCalory - 30 * 0.105 * .sngWeight) / (.sngWeight * 0.093)
                    lblMuscleTime.Caption = "30 �� (�ٷ¿ �� �����ȱ� " & intExTime & " ��)"
                Else
                    lblMuscleTime.Caption = "30 ��"
                End If
            Case 11, 18, 25, 32, 39, 46:
                lblMuscleTime.Caption = "30 ��"
            Case 12, 19, 26, 33, 40, 47:
                If .intExDay = 6 Then
                    intExTime = Int(.sngExCalory / (.sngWeight * 0.105))
                    lblMuscleTime.Caption = intExTime & " ��"
                Else
                    lblMuscleTime.Caption = "30 ��"
                End If
            Case 13, 20, 27, 34, 41, 48:
                If .intExDay = 6 Then
                    intExTime = Int(.sngExCalory - 30 * 0.105 * .sngWeight) / (.sngWeight * 0.093)
                    lblMuscleTime.Caption = "30 �� (�ٷ¿ �� �����ȱ� " & intExTime & " ��)"
                Else
                    lblMuscleTime.Caption = "30 ��"
                End If
            Case 14, 21, 28, 35, 42, 49:
                If .intExDay = 6 Then
                    intExTime = Int(.sngExCalory / (.sngWeight * 0.105))
                    lblMuscleTime.Caption = intExTime & " ��"
                Else
                    lblMuscleTime.Caption = "30 ��"
                End If
            Case 15, 22, 29, 36, 43, 50:
                If .intExDay = 6 Then
                    intExTime = Int(.sngExCalory / (.sngWeight * 0.105))
                    lblMuscleTime.Caption = intExTime & " ��"
                Else
                    lblMuscleTime.Caption = "30 ��"
                End If
        End Select
        End With
        '���
        If typExProgram.intExDay = 0 Then
            lblMuscle2.Caption = Is_Null(rValue(1, 0), "")
        Else
            Select Case intProgram
                Case 9, 23, 37
                    lblMuscle2.Caption = typExProgram.intExDay & "��/��"
                Case 11, 18, 25, 32, 39, 46, 30
                    lblMuscle2.Caption = typExProgram.intExDay & "��/��(����ҿ�� ����)"
                Case Else
                    Select Case typExProgram.intExDay
                        Case 3:
                            lblMuscle2.Caption = "3��/��(����ҿ�� ����)"
                        Case 4:
                            lblMuscle2.Caption = "4��/��(����ҿ�� ����)"
                        Case Else
                            lblMuscle2.Caption = "3��/��"
                    End Select
            End Select
        End If
        lblMuscle3.Caption = Is_Null(rValue(2, 0), "")
    End If
    
    Set clsSelect = Nothing
End Sub

'����� � : �ٷ¿, ��Ʈ��Ī �̹����� ���� �����ֱ�
Private Sub DrawExImages(intGrdIndex As Integer, intMaxCol As Integer, intFrom As Integer)
    Dim qrySelect As String, rValue As Variant
    Dim i As Integer

    Set clsSelect = New clsSelect
    With grdExImages(intGrdIndex)
        .Visible = True
        .ColS = intMaxCol
        For i = 0 To intMaxCol - 1
            qrySelect = "SELECT ImageName, Explain FROM tblExerciseNoO2 WHERE xno=" & i + intFrom
            rValue = clsSelect.Query(qrySelect)
            If Not IsNull(rValue) Then
                .Row = 0: .Col = i
                Set .CellPicture = LoadPicture(App.Path & "\Back\Ex\N\" & Trim(rValue(0, 0)) & ".jpg")
                .TextMatrix(1, i) = Trim(rValue(1, 0))
                .Row = 1: .Col = i
                .CellBackColor = FRM_GRAY
            End If
        Next i
        .ColWidth(-1) = 2000
        .ColAlignment(-1) = flexAlignLeftTop
        .Font.Size = 8
    End With
    Set clsSelect = Nothing
End Sub
