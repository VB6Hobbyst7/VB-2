VERSION 5.00
Object = "{4BD5DFC7-B668-44E0-A002-C1347061239D}#1.0#0"; "HSCotrol.ocx"
Begin VB.Form frmHospInfo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '���� ����
   Caption         =   "�����������"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5850
   Icon            =   "frmHospInfo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   5850
   StartUpPosition =   1  '������ ���
   Begin VB.Frame Frame2 
      Caption         =   "Hidden"
      Height          =   1545
      Left            =   5220
      TabIndex        =   39
      Top             =   2550
      Visible         =   0   'False
      Width           =   4305
      Begin VB.TextBox txtColWidth 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BeginProperty Font 
            Name            =   "���� ����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1545
         TabIndex        =   40
         Top             =   390
         Width           =   2565
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         Appearance      =   0  '���
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '����
         Caption         =   "�˻�� ���� "
         BeginProperty Font 
            Name            =   "���� ����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   225
         Index           =   5
         Left            =   420
         TabIndex        =   41
         Top             =   435
         Width           =   1020
      End
   End
   Begin VB.TextBox txtSaveDay 
      Alignment       =   2  '��� ����
      Appearance      =   0  '���
      BeginProperty Font 
         Name            =   "���� ����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   3660
      TabIndex        =   36
      Top             =   5790
      Width           =   705
   End
   Begin VB.CommandButton cmdHospInfo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "���� ����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   150
      Style           =   1  '�׷���
      TabIndex        =   32
      Top             =   6570
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox chkDBCon 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DB����üũ"
      BeginProperty Font 
         Name            =   "���� ����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   345
      Left            =   2460
      TabIndex        =   31
      Top             =   5190
      Width           =   1365
   End
   Begin VB.TextBox txtBarLen 
      Alignment       =   2  '��� ����
      Appearance      =   0  '���
      BeginProperty Font 
         Name            =   "���� ����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2460
      TabIndex        =   29
      Top             =   3990
      Width           =   2565
   End
   Begin VB.ComboBox cboMachs 
      BeginProperty Font 
         Name            =   "Segoe UI Historic"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2460
      TabIndex        =   27
      Top             =   840
      Visible         =   0   'False
      Width           =   2595
   End
   Begin VB.CommandButton cmdLocalDBSet 
      BackColor       =   &H00FFFFFF&
      Caption         =   "���� DB"
      BeginProperty Font 
         Name            =   "���� ����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   150
      Style           =   1  '�׷���
      TabIndex        =   26
      Top             =   5730
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdDBSet 
      BackColor       =   &H00FFFFFF&
      Caption         =   "���� DB"
      BeginProperty Font 
         Name            =   "���� ����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   150
      Style           =   1  '�׷���
      TabIndex        =   25
      Top             =   6150
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '����
      BeginProperty Font 
         Name            =   "���� ����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2460
      TabIndex        =   24
      Top             =   4410
      Width           =   2565
      Begin VB.OptionButton optWorkPos 
         BackColor       =   &H00FFFFFF&
         Caption         =   "�˾�"
         ForeColor       =   &H00404040&
         Height          =   225
         Index           =   1
         Left            =   1230
         TabIndex        =   11
         Top             =   60
         Width           =   1125
      End
      Begin VB.OptionButton optWorkPos 
         BackColor       =   &H00FFFFFF&
         Caption         =   "����"
         ForeColor       =   &H00404040&
         Height          =   225
         Index           =   0
         Left            =   90
         TabIndex        =   10
         Top             =   60
         Value           =   -1  'True
         Width           =   1125
      End
   End
   Begin VB.CheckBox chkLog 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�αױ��"
      BeginProperty Font 
         Name            =   "���� ����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   345
      Left            =   3900
      TabIndex        =   14
      Top             =   5190
      Width           =   1125
   End
   Begin VB.TextBox txtPartNm 
      Appearance      =   0  '���
      BeginProperty Font 
         Name            =   "���� ����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3240
      TabIndex        =   5
      Top             =   2415
      Width           =   1785
   End
   Begin VB.TextBox txtLabNm 
      Appearance      =   0  '���
      BeginProperty Font 
         Name            =   "���� ����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3240
      TabIndex        =   3
      Top             =   2010
      Width           =   1785
   End
   Begin VB.TextBox txtHospNm 
      Alignment       =   2  '��� ����
      Appearance      =   0  '���
      BeginProperty Font 
         Name            =   "���� ����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2460
      TabIndex        =   1
      Top             =   1620
      Width           =   2565
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  '���
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '����
      BeginProperty Font 
         Name            =   "���� ����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   2460
      TabIndex        =   22
      Top             =   4800
      Width           =   2565
      Begin VB.OptionButton optLoginUse 
         BackColor       =   &H00FFFFFF&
         Caption         =   "�̻��"
         ForeColor       =   &H00404040&
         Height          =   225
         Index           =   0
         Left            =   90
         TabIndex        =   12
         Top             =   90
         Value           =   -1  'True
         Width           =   1125
      End
      Begin VB.OptionButton optLoginUse 
         BackColor       =   &H00FFFFFF&
         Caption         =   "���"
         ForeColor       =   &H00404040&
         Height          =   225
         Index           =   1
         Left            =   1230
         TabIndex        =   13
         Top             =   90
         Width           =   1125
      End
   End
   Begin VB.TextBox txtUserNm 
      Alignment       =   2  '��� ����
      Appearance      =   0  '���
      BeginProperty Font 
         Name            =   "���� ����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2460
      TabIndex        =   9
      Top             =   3600
      Width           =   2565
   End
   Begin VB.TextBox txtUserID 
      Alignment       =   2  '��� ����
      Appearance      =   0  '���
      BeginProperty Font 
         Name            =   "���� ����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2460
      TabIndex        =   8
      Top             =   3195
      Width           =   2565
   End
   Begin VB.TextBox txtMachNm 
      Appearance      =   0  '���
      BeginProperty Font 
         Name            =   "���� ����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3240
      TabIndex        =   7
      Top             =   2805
      Width           =   1785
   End
   Begin VB.TextBox txtLabCd 
      Alignment       =   2  '��� ����
      Appearance      =   0  '���
      BeginProperty Font 
         Name            =   "���� ����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2460
      TabIndex        =   2
      Top             =   2010
      Width           =   765
   End
   Begin VB.TextBox txtPartCd 
      Alignment       =   2  '��� ����
      Appearance      =   0  '���
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "���� ����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2460
      TabIndex        =   4
      Top             =   2415
      Width           =   765
   End
   Begin VB.TextBox txtHospCd 
      Alignment       =   2  '��� ����
      Appearance      =   0  '���
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "���� ����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2460
      TabIndex        =   0
      Top             =   1230
      Width           =   2565
   End
   Begin VB.TextBox txtMachCd 
      Alignment       =   2  '��� ����
      Appearance      =   0  '���
      BeginProperty Font 
         Name            =   "���� ����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2460
      TabIndex        =   6
      Top             =   2805
      Width           =   765
   End
   Begin HSCotrol.CButton cmdSave 
      Height          =   495
      Left            =   2160
      TabIndex        =   34
      Top             =   6450
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   873
      BackColor       =   12632256
      Caption         =   " ��������"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmHospInfo.frx":08CA
      MaskColor       =   0
      PicCapAlign     =   2
      BorderStyle     =   1
      BorderColor     =   16777215
      HoverColor      =   -2147483630
   End
   Begin HSCotrol.CButton cmdCancel 
      Height          =   495
      Left            =   3630
      TabIndex        =   35
      Top             =   6450
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   873
      BackColor       =   12632256
      Caption         =   " ��    ��"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmHospInfo.frx":0A24
      MaskColor       =   0
      PicCapAlign     =   2
      BorderStyle     =   1
      BorderColor     =   16777215
      HoverColor      =   -2147483630
   End
   Begin HSCotrol.HSLabel HSLabel1 
      Height          =   345
      Left            =   150
      Top             =   150
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   609
      BackColor       =   16311496
      ForeColor       =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   " �� �����������"
      BevelOut        =   0
      Begin VB.Label lblHosp 
         BackStyle       =   0  '����
         Caption         =   "�󼼼���"
         BeginProperty Font 
            Name            =   "���� ����"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   4680
         TabIndex        =   42
         Top             =   60
         Visible         =   0   'False
         Width           =   765
      End
   End
   Begin VB.Label Label1 
      Appearance      =   0  '���
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      Caption         =   "�� ����"
      BeginProperty Font 
         Name            =   "���� ����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Index           =   11
      Left            =   4440
      TabIndex        =   38
      Top             =   5850
      Width           =   600
   End
   Begin VB.Label Label1 
      Appearance      =   0  '���
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      Caption         =   "�˻���      ��������Ⱓ"
      BeginProperty Font 
         Name            =   "���� ����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   450
      Index           =   9
      Left            =   2460
      TabIndex        =   37
      Top             =   5640
      Width           =   1170
      WordWrap        =   -1  'True
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00BF8B59&
      BorderWidth     =   2
      Height          =   405
      Left            =   120
      Top             =   120
      Width           =   5625
   End
   Begin VB.Label Label1 
      Alignment       =   1  '������ ����
      Appearance      =   0  '���
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      Caption         =   "����ڵ�"
      BeginProperty Font 
         Name            =   "���� ����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Index           =   7
      Left            =   1545
      TabIndex        =   33
      Top             =   1290
      Width           =   720
   End
   Begin VB.Label Label1 
      Alignment       =   1  '������ ����
      Appearance      =   0  '���
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      Caption         =   "���ڵ���� "
      BeginProperty Font 
         Name            =   "���� ����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Index           =   6
      Left            =   1395
      TabIndex        =   30
      Top             =   4035
      Width           =   960
   End
   Begin VB.Label lblMachNm 
      Alignment       =   1  '������ ����
      Appearance      =   0  '���
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "Segoe UI Historic"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Left            =   1710
      TabIndex        =   28
      Top             =   900
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label1 
      Alignment       =   1  '������ ����
      Appearance      =   0  '���
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      Caption         =   "��ũ��ȸ "
      BeginProperty Font 
         Name            =   "���� ����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Index           =   1
      Left            =   1575
      TabIndex        =   23
      Top             =   4470
      Width           =   780
   End
   Begin VB.Label Label1 
      Alignment       =   1  '������ ����
      Appearance      =   0  '���
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      Caption         =   "�α��� "
      BeginProperty Font 
         Name            =   "���� ����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Index           =   4
      Left            =   1755
      TabIndex        =   21
      Top             =   4890
      Width           =   600
   End
   Begin VB.Label Label1 
      Alignment       =   1  '������ ����
      Appearance      =   0  '���
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      Caption         =   "����� �� "
      BeginProperty Font 
         Name            =   "���� ����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Index           =   3
      Left            =   1515
      TabIndex        =   20
      Top             =   3645
      Width           =   840
   End
   Begin VB.Label Label1 
      Alignment       =   1  '������ ����
      Appearance      =   0  '���
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      Caption         =   "����� ID "
      BeginProperty Font 
         Name            =   "���� ����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Index           =   2
      Left            =   1515
      TabIndex        =   19
      Top             =   3240
      Width           =   840
   End
   Begin VB.Label Label1 
      Alignment       =   1  '������ ����
      Appearance      =   0  '���
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      Caption         =   "�˻���Ʈ �ڵ�/�� "
      BeginProperty Font 
         Name            =   "���� ����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Index           =   10
      Left            =   885
      TabIndex        =   18
      Top             =   2475
      Width           =   1455
   End
   Begin VB.Label ����ڸ� 
      Alignment       =   1  '������ ����
      Appearance      =   0  '���
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      Caption         =   "���μ� �ڵ�/�� "
      BeginProperty Font 
         Name            =   "���� ����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Index           =   9
      Left            =   900
      TabIndex        =   17
      Top             =   2055
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  '������ ����
      Appearance      =   0  '���
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      Caption         =   "������ "
      BeginProperty Font 
         Name            =   "���� ����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Index           =   8
      Left            =   1740
      TabIndex        =   16
      Top             =   1680
      Width           =   600
   End
   Begin VB.Label Label1 
      Alignment       =   1  '������ ����
      Appearance      =   0  '���
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      Caption         =   "��� �ڵ�/�� "
      BeginProperty Font 
         Name            =   "���� ����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Index           =   0
      Left            =   1260
      TabIndex        =   15
      Top             =   2850
      Width           =   1095
   End
End
Attribute VB_Name = "frmHospInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboMachs_Click()
    
    lblMachNm.Caption = cboMachs.Text

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDBSet_Click()
    
    If gDBTYPE = "99" Then
        'Call WritePrivateProfileString("HOSP", "HOSPCD", txtHospCd.Text, App.PATH & "\INI\" & gMACH & ".ini")
        Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "HOSPCD", txtHospCd.Text)
    Else
        If InputBox("��й�ȣ �Է�" & Space(5) & "hint:������ogb") = "dev0503" Then
            If gDBTYPE = "1" Then
                frmDB_Oracle.Show vbModal
            ElseIf gDBTYPE = "2" Then
                frmDB_MSSQL.Show vbModal
            ElseIf gDBTYPE = "3" Then
                frmDB_PGSQL.Show vbModal
            Else
                MsgBox App.PATH & "\OKSOFT.ini ���Ͽ���" & vbNewLine & vbNewLine & "DBTYPE�� ���� �����ϼ��� ", vbOKOnly + vbInformation, "DB TYPE ����"
            End If
        End If
    End If
    
End Sub

Private Sub cmdHospInfo_Click()
    
    If InputBox("��й�ȣ �Է�" & Space(5) & "hint:������ogb") = "dev0810" Then
        frmEMRInfo.Show
    End If
    
End Sub

Private Sub cmdLocalDBSet_Click()
    
    frmDB_Local.Show vbModal

End Sub

Private Sub cmdSave_Click()
    Dim strDBType   As String
    
    'Call WritePrivateProfileString("HOSP", "HOSPCD", txtHospCd.Text, App.PATH & "\INI\" & gMACH & ".ini")
    'Call WritePrivateProfileString("HOSP", "HOSPNM", txtHospNm.Text, App.PATH & "\INI\" & gMACH & ".ini")
    'Call WritePrivateProfileString("HOSP", "LABCD", txtLabCd.Text, App.PATH & "\INI\" & gMACH & ".ini")
    'Call WritePrivateProfileString("HOSP", "LABNM", txtLabNm.Text, App.PATH & "\INI\" & gMACH & ".ini")
    'Call WritePrivateProfileString("HOSP", "PARTCD", txtPartCd.Text, App.PATH & "\INI\" & gMACH & ".ini")
    'Call WritePrivateProfileString("HOSP", "PARTNM", txtPartNm.Text, App.PATH & "\INI\" & gMACH & ".ini")
    'Call WritePrivateProfileString("HOSP", "MACHCD", txtMachCd.Text, App.PATH & "\INI\" & gMACH & ".ini")
    'Call WritePrivateProfileString("HOSP", "MACHNM", txtMachNm.Text, App.PATH & "\INI\" & gMACH & ".ini")
    'Call WritePrivateProfileString("HOSP", "USERID", txtUserID.Text, App.PATH & "\INI\" & gMACH & ".ini")
    'Call WritePrivateProfileString("HOSP", "USERNM", txtUserNm.Text, App.PATH & "\INI\" & gMACH & ".ini")
    
    Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "HOSPCD", txtHospCd.Text)
    Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "HOSPNM", txtHospNm.Text)
    Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "LABCD", txtLabCd.Text)
    Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "LABNM", txtLabNm.Text)
    Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "PARTCD", txtPartCd.Text)
    Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "PARTNM", txtPartNm.Text)
    Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "MACHCD", txtMachCd.Text)
    Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "MACHNM", txtMachNm.Text)
    Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "USERID", txtUserID.Text)
    Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "USERNM", txtUserNm.Text)
    
    
    If optLoginUse(0).Value = True Then
        'Call WritePrivateProfileString("HOSP", "LOGINYN", "N", App.PATH & "\INI\" & gMACH & ".ini")
        Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "LOGINYN", "N")
    Else
        'Call WritePrivateProfileString("HOSP", "LOGINYN", "Y", App.PATH & "\INI\" & gMACH & ".ini")
        Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "LOGINYN", "Y")
    End If
    
    If chkLog.Value = "1" Then
        'Call WritePrivateProfileString("HOSP", "LOGWRITE", "1", App.PATH & "\INI\" & gMACH & ".ini")
        Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "LOGWRITE", "1")
    Else
        'Call WritePrivateProfileString("HOSP", "LOGWRITE", "0", App.PATH & "\INI\" & gMACH & ".ini")
        Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "LOGWRITE", "0")
    End If
    
    If optWorkPos(0).Value = True Then
        'Call WritePrivateProfileString("VIEW", "WORKPOS", "M", App.PATH & "\INI\" & gMACH & ".ini")
        Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "WORKPOS", "M")
    Else
        'Call WritePrivateProfileString("VIEW", "WORKPOS", "P", App.PATH & "\INI\" & gMACH & ".ini")
        Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "WORKPOS", "P")
    End If
    
    'Call WritePrivateProfileString("VIEW", "COLWIDTH", txtColWidth.Text, App.PATH & "\INI\" & gMACH & ".ini")
    Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "VIEW", "COLWIDTH", txtColWidth.Text)
    
    If lblMachNm.Caption <> "" Then
        Call WritePrivateProfileString("EXE", "MACH", lblMachNm.Caption, App.PATH & "\OKSOFT.ini")
    End If
    
    If chkDBCon.Value = "1" Then
        'Call WritePrivateProfileString("HOSP", "DBCONCHK", "Y", App.PATH & "\INI\" & gMACH & ".ini")
        Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "DBCONCHK", "Y")
    Else
        'Call WritePrivateProfileString("HOSP", "DBCONCHK", "N", App.PATH & "\INI\" & gMACH & ".ini")
        Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "DBCONCHK", "N")
    End If
    
    If txtBarLen.Text <> "" And IsNumeric(txtBarLen.Text) Then
        'Call WritePrivateProfileString("HOSP", "BARLEN", txtBarLen.Text, App.PATH & "\INI\" & gMACH & ".ini")
        Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "BARLEN", txtBarLen.Text)
    End If
    
    Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "SAVEDAY", txtSaveDay.Text)
    
    
    If gLocalDB.PATH <> "" Then
        Call LetEqpMaster(Trim(txtMachCd.Text))
    End If
    
    SQL = ""
    SQL = SQL & "UPDATE EQPMASTER SET " & vbCrLf
    SQL = SQL & " EQUIPCD = " & STS(txtMachCd.Text)
    
    Call DBExec(AdoCn_Local, SQL)
    
    SQL = ""
    SQL = SQL & "UPDATE AMRMASTER SET " & vbCrLf
    SQL = SQL & " EQUIPCD = " & STS(txtMachCd.Text)
    
    Call DBExec(AdoCn_Local, SQL)
    
    GetSetup
    
    Unload Me

    Call Main
End Sub

Private Sub Form_Load()
    
    Call CtlInitializing
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    
End Sub

Public Sub CtlInitializing()
    Dim i       As Integer
    Dim intIdx  As Integer
    
    txtHospCd.Text = gHOSP.HOSPCD
    txtHospNm.Text = gHOSP.HOSPNM
    txtLabCd.Text = gHOSP.LABCD
    txtLabNm.Text = gHOSP.LABNM
    txtPartCd.Text = gHOSP.PARTCD
    txtPartNm.Text = gHOSP.PARTNM
    txtMachCd.Text = gHOSP.MACHCD
    txtMachNm.Text = gHOSP.MACHNM 'gmach
    txtUserID.Text = gHOSP.USERID
    txtUserNm.Text = gHOSP.USERNM
    txtBarLen.Text = gHOSP.BARLEN
    
    If gHOSP.DBCONCHK = "Y" Then
        chkDBCon.Value = "1"
    Else
        chkDBCon.Value = "0"
    End If
    
    If gHOSP.LOGINYN = "Y" Then
        optLoginUse(1).Value = True
    Else
        optLoginUse(0).Value = True
    End If
    If gHOSP.LOQWRITE = "1" Then
        chkLog.Value = "1"
    Else
        chkLog.Value = "0"
    End If
    
    If gWORKPOS = "P" Then
        optWorkPos(1).Value = True
    Else
        optWorkPos(0).Value = True
    End If
    
    If gCOLWIDTH = "" Then
        txtColWidth.Text = "10"
    Else
        txtColWidth.Text = gCOLWIDTH
    End If
    
    lblMachNm.Caption = ""
    intIdx = 0
    
    cboMachs.Clear
    If IsNumeric(gMACHCOUNT) Then
        For i = 1 To gMACHCOUNT
            cboMachs.AddItem gMACHS(i)
            If gHOSP.MACHNM = gMACHS(i) Then
                intIdx = i
            End If
        Next
        cboMachs.ListIndex = intIdx - 1
    End If
    
    txtSaveDay.Text = gHOSP.SAVEDAY
    
End Sub


Private Sub lblHosp_DblClick()
    Dim strPW   As String
    
    If cboMachs.Visible = False Then
        strPW = InputBox("��й�ȣ �Է�" & Space(5) & "hint:������oyh", "�󼼼���", "")
        If strPW = "dev0503" Then
            cboMachs.Visible = True
            lblMachNm.Visible = True
            cmdLocalDBSet.Visible = True
            cmdDBSet.Visible = True
            cmdHospInfo.Visible = True
            
        End If
    Else
        cboMachs.Visible = False
        lblMachNm.Visible = False
        cmdLocalDBSet.Visible = False
        cmdDBSet.Visible = False
        cmdHospInfo.Visible = False
    End If

End Sub
