VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmEQ����_���˻��ڵ����_�Է� 
   BorderStyle     =   1  '���� ����
   Caption         =   "���˻��ڵ����"
   ClientHeight    =   7275
   ClientLeft      =   6540
   ClientTop       =   2715
   ClientWidth     =   9315
   BeginProperty Font 
      Name            =   "����ü"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEQ����_���˻��ڵ����_�Է�.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   9315
   Begin VB.Frame fra�󼼳��� 
      Caption         =   "[�󼼳���]"
      Height          =   6075
      Left            =   60
      TabIndex        =   30
      Top             =   1140
      Width           =   9195
      Begin VB.ComboBox cboEQORDYN 
         Height          =   300
         Left            =   1680
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   2
         Top             =   600
         Width           =   1515
      End
      Begin FPSpread.vaSpread sprEXCD 
         Height          =   5415
         Left            =   6420
         TabIndex        =   26
         Top             =   540
         Width           =   2655
         _Version        =   393216
         _ExtentX        =   4683
         _ExtentY        =   9551
         _StockProps     =   64
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   3
         SpreadDesigner  =   "frmEQ����_���˻��ڵ����_�Է�.frx":263A
      End
      Begin VB.Frame Frame3 
         Caption         =   "[CutOff] ����� ���ں�ȯ"
         Height          =   2115
         Left            =   120
         TabIndex        =   47
         Top             =   3840
         Width           =   6135
         Begin VB.ComboBox cboEQCUTRTYPE 
            Height          =   300
            Left            =   1020
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   25
            Top             =   1680
            Width           =   4995
         End
         Begin VB.ComboBox cboEQCUTOFFNM 
            Height          =   300
            Left            =   1020
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   24
            Top             =   1320
            Width           =   4995
         End
         Begin VB.ComboBox cboEQCUTMNM 
            Height          =   300
            Left            =   1020
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   23
            Top             =   960
            Width           =   4995
         End
         Begin VB.TextBox txtEQCUTLVAL 
            Height          =   300
            Left            =   4020
            TabIndex        =   21
            Text            =   "txtEQCUTLVAL"
            Top             =   600
            Width           =   855
         End
         Begin VB.ComboBox cboEQCUTLREF 
            Height          =   300
            Left            =   4920
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   22
            Top             =   600
            Width           =   1095
         End
         Begin VB.TextBox txtEQCUTHVAL 
            Height          =   300
            Left            =   1020
            TabIndex        =   19
            Text            =   "txtEQCUTHVAL"
            Top             =   600
            Width           =   855
         End
         Begin VB.ComboBox cboEQCUTHREF 
            Height          =   300
            Left            =   1980
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   20
            Top             =   600
            Width           =   1095
         End
         Begin VB.ComboBox cboEQCUTOFFGB 
            Height          =   300
            Left            =   1020
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   18
            Top             =   240
            Width           =   4995
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  '����
            Caption         =   "ǥ������"
            Height          =   180
            Index           =   19
            Left            =   180
            TabIndex        =   53
            Top             =   1740
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  '����
            Caption         =   "�� �� ��"
            Height          =   180
            Index           =   18
            Left            =   180
            TabIndex        =   52
            Top             =   1380
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  '����
            Caption         =   "�� �� ��"
            Height          =   180
            Index           =   17
            Left            =   180
            TabIndex        =   51
            Top             =   1020
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  '����
            Caption         =   "�� �� ��"
            Height          =   180
            Index           =   16
            Left            =   3240
            TabIndex        =   50
            Top             =   660
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  '����
            Caption         =   "�� �� ��"
            Height          =   180
            Index           =   15
            Left            =   180
            TabIndex        =   49
            Top             =   660
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  '����
            Caption         =   "���뱸��"
            Height          =   180
            Index           =   14
            Left            =   180
            TabIndex        =   48
            Top             =   300
            Width           =   720
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "[Limit] ����� ��ġ����"
         Height          =   675
         Left            =   120
         TabIndex        =   44
         Top             =   3120
         Width           =   6135
         Begin VB.ComboBox cboEQLIMITFLAG2 
            Height          =   300
            Left            =   4020
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   16
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtEQLIMITVALUE2 
            Height          =   300
            Left            =   5160
            TabIndex        =   17
            Text            =   "txtEQLIMITVALUE2"
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txtEQLIMITVALUE1 
            Height          =   300
            Left            =   1980
            TabIndex        =   15
            Text            =   "txtEQLIMITVALUE1"
            Top             =   240
            Width           =   855
         End
         Begin VB.ComboBox cboEQLIMITFLAG1 
            Height          =   300
            Left            =   840
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   14
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  '����
            Caption         =   "���Ѱ�"
            Height          =   180
            Index           =   13
            Left            =   3300
            TabIndex        =   46
            Top             =   300
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  '����
            Caption         =   "���Ѱ�"
            Height          =   180
            Index           =   8
            Left            =   180
            TabIndex        =   45
            Top             =   300
            Width           =   540
         End
      End
      Begin MSMask.MaskEdBox mskEQSEQ 
         Height          =   195
         Left            =   4980
         TabIndex        =   4
         Top             =   1020
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   344
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   "_"
      End
      Begin VB.Frame Frame1 
         Caption         =   "[��������ġ]"
         Height          =   1035
         Left            =   120
         TabIndex        =   36
         Top             =   2040
         Width           =   6135
         Begin VB.ComboBox cboEQRFLREF 
            Height          =   300
            Left            =   1740
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   11
            Top             =   600
            Width           =   1095
         End
         Begin VB.TextBox txtEQRFLVAL 
            Height          =   300
            Left            =   840
            TabIndex        =   10
            Text            =   "txtEQAFLVAL"
            Top             =   600
            Width           =   855
         End
         Begin VB.TextBox txtEQRFHVAL 
            Height          =   300
            Left            =   4020
            TabIndex        =   12
            Text            =   "txtEQAFHVAL"
            Top             =   600
            Width           =   855
         End
         Begin VB.ComboBox cboEQRFHREF 
            Height          =   300
            Left            =   4920
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   13
            Top             =   600
            Width           =   1095
         End
         Begin VB.ComboBox cboEQRMLREF 
            Height          =   300
            Left            =   1740
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   7
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtEQRMLVAL 
            Height          =   300
            Left            =   840
            TabIndex        =   6
            Text            =   "txtEQAMLVAL"
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txtEQRMHVAL 
            Height          =   300
            Left            =   4020
            TabIndex        =   8
            Text            =   "txtEQAMHVAL"
            Top             =   240
            Width           =   855
         End
         Begin VB.ComboBox cboEQRMHREF 
            Height          =   300
            Left            =   4920
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   9
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  '����
            Caption         =   "�� Low"
            Height          =   180
            Index           =   12
            Left            =   180
            TabIndex        =   40
            Top             =   660
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  '����
            Caption         =   "�� High"
            Height          =   180
            Index           =   10
            Left            =   3300
            TabIndex        =   39
            Top             =   660
            Width           =   630
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  '����
            Caption         =   "�� Low"
            Height          =   180
            Index           =   11
            Left            =   180
            TabIndex        =   38
            Top             =   300
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  '����
            Caption         =   "�� High"
            Height          =   180
            Index           =   5
            Left            =   3300
            TabIndex        =   37
            Top             =   300
            Width           =   630
         End
      End
      Begin VB.TextBox txtEQUNIT 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         Height          =   195
         IMEMode         =   8  '����
         Left            =   1680
         TabIndex        =   3
         Text            =   "1234567890"
         Top             =   1020
         Width           =   915
      End
      Begin VB.TextBox txtEQNM 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         Height          =   195
         IMEMode         =   8  '����
         Left            =   1680
         TabIndex        =   1
         Text            =   "12345678901234567890123456789012345678901234567890"
         Top             =   300
         Width           =   4575
      End
      Begin MSMask.MaskEdBox mskEQRSTRANGE 
         Height          =   195
         Left            =   2220
         TabIndex        =   5
         Top             =   1380
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   344
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   "_"
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   120
         X2              =   6240
         Y1              =   900
         Y2              =   900
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "������ۿ���"
         Height          =   195
         Index           =   20
         Left            =   120
         TabIndex        =   55
         Top             =   660
         Width           =   1515
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ó���ڵ�"
         Height          =   195
         Index           =   6
         Left            =   6420
         TabIndex        =   54
         Top             =   300
         Width           =   2655
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '����
         Caption         =   "�� �����Ű���� ������ ��� 0���� ǥ��"
         Height          =   180
         Index           =   7
         Left            =   120
         TabIndex        =   43
         Top             =   1680
         Width           =   3600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '����
         Caption         =   "(0: ��üǥ��, 1 �̻� : ���ڸ�ŭ ǥ��)"
         Height          =   180
         Index           =   2
         Left            =   2700
         TabIndex        =   42
         Top             =   1380
         Width           =   3330
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   4
         X1              =   120
         X2              =   6240
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "�Ҽ������� ǥ�� �ڸ���"
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   41
         Top             =   1380
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ȭ�����ļ���"
         Height          =   195
         Index           =   4
         Left            =   3420
         TabIndex        =   35
         Top             =   1020
         Width           =   1515
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   2
         X1              =   120
         X2              =   6240
         Y1              =   1260
         Y2              =   1260
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "�˻�������"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   34
         Top             =   1020
         Width           =   1515
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "���˻��"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   31
         Top             =   300
         Width           =   1515
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   1
         X1              =   120
         X2              =   6240
         Y1              =   540
         Y2              =   540
      End
   End
   Begin VB.TextBox txtEQCD 
      Appearance      =   0  '���
      BorderStyle     =   0  '����
      Height          =   195
      Left            =   1740
      MaxLength       =   10
      TabIndex        =   0
      Text            =   "1234567890"
      Top             =   780
      Width           =   975
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
      Left            =   8340
      Style           =   1  '�׷���
      TabIndex        =   28
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton cmdSave 
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
      Left            =   7380
      Style           =   1  '�׷���
      TabIndex        =   27
      Top             =   60
      Width           =   915
   End
   Begin MSComctlLib.ProgressBar barStatus 
      Height          =   75
      Left            =   60
      TabIndex        =   32
      Top             =   600
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   132
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label lbl���� 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "���˻��ڵ��Է�"
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
      TabIndex        =   33
      Top             =   60
      Width           =   2880
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   7
      X1              =   120
      X2              =   2700
      Y1              =   1020
      Y2              =   1020
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "���˻��ڵ�"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   29
      Top             =   780
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808000&
      BackStyle       =   1  '�������� ����
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFC0C0&
      FillStyle       =   5  '���� �밢��
      Height          =   495
      Index           =   3
      Left            =   60
      Shape           =   4  '�ձ� �簢��
      Top             =   60
      Width           =   4395
   End
End
Attribute VB_Name = "frmEQ����_���˻��ڵ����_�Է�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function SUB_MM_CANCEL() As Boolean
    barStatus.Max = 100
    barStatus.Value = 100
    
    txtEQCD = ""
    
    Call SUB_MM_KEY_CLEAR
End Function

Public Function MM_DELETE() As Boolean

End Function

Private Sub SUB_MM_INITIAL()
    Me.Height = 7755
    Me.Width = 9435
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    
    '/�ʱ� �ڷ� Setting----------------------------------------------------------------------------------------------------/
    GoSub ADD_ITEM
    '/�ʱ� �ڷ� Setting----------------------------------------------------------------------------------------------------/
    
    Call SUB_MM_CANCEL

    If gstrInputUpdate = "2" Then '/1.Input, 2.Update
        txtEQCD = gstrArgTemp1
        txtEQCD.BackColor = RGB(255, 255, 240)
        txtEQCD.Enabled = False
        Call FUNC_MM_VIEW
    End If
    
Exit Sub

'/----------------------------------------------------------------------------------------------------/

ADD_ITEM:
    '/���˻��ڵ� ������ۿ���(Y.����, N.������)
    cboEQORDYN.AddItem "����" & Space(100) & "Y"
    cboEQORDYN.AddItem "������" & Space(100) & "N"
    
    '/Reference
    cboEQRMHREF.AddItem ""
    cboEQRMHREF.AddItem "�̻�" & Space(100) & "1"
    cboEQRMHREF.AddItem "�ʰ�" & Space(100) & "2"
    cboEQRMHREF.AddItem "����" & Space(100) & "3"
    cboEQRMHREF.AddItem "�̸�" & Space(100) & "4"

    cboEQRMLREF.AddItem ""
    cboEQRMLREF.AddItem "�̻�" & Space(100) & "1"
    cboEQRMLREF.AddItem "�ʰ�" & Space(100) & "2"
    cboEQRMLREF.AddItem "����" & Space(100) & "3"
    cboEQRMLREF.AddItem "�̸�" & Space(100) & "4"

    cboEQRFHREF.AddItem ""
    cboEQRFHREF.AddItem "�̻�" & Space(100) & "1"
    cboEQRFHREF.AddItem "�ʰ�" & Space(100) & "2"
    cboEQRFHREF.AddItem "����" & Space(100) & "3"
    cboEQRFHREF.AddItem "�̸�" & Space(100) & "4"

    cboEQRFLREF.AddItem ""
    cboEQRFLREF.AddItem "�̻�" & Space(100) & "1"
    cboEQRFLREF.AddItem "�ʰ�" & Space(100) & "2"
    cboEQRFLREF.AddItem "����" & Space(100) & "3"
    cboEQRFLREF.AddItem "�̸�" & Space(100) & "4"

    '/LIMIT ���Ѱ� ����(0.������, 1.����, 2.�̸�)
    cboEQLIMITFLAG1.AddItem "������" & Space(100) & "0"
    cboEQLIMITFLAG1.AddItem "����" & Space(100) & "1"
    cboEQLIMITFLAG1.AddItem "�̸�" & Space(100) & "2"

    '/LIMIT ���Ѱ� ����(0.������, 1.�̻�, 2.�ʰ�)
    cboEQLIMITFLAG2.AddItem "������" & Space(100) & "0"
    cboEQLIMITFLAG2.AddItem "�̻�" & Space(100) & "1"
    cboEQLIMITFLAG2.AddItem "�ʰ�" & Space(100) & "2"

    '/CUTOFF ���뱸��(0.�������, 1.���� Positive, 2.���� Positive, 3.�����������(��ġ�� CutOff���� ���ÿ� ���ð��)
    cboEQCUTOFFGB.AddItem "�������" & Space(100) & "0"
    cboEQCUTOFFGB.AddItem "���� Positive" & Space(100) & "1"
    cboEQCUTOFFGB.AddItem "���� Positive" & Space(100) & "2"
    cboEQCUTOFFGB.AddItem "�����������" & Space(100) & "3"
        
    '/CUTOFF ���Ѱ� Equal ���� ����(1.�̻�, 2.�ʰ�)
    cboEQCUTHREF.AddItem ""
    cboEQCUTHREF.AddItem "�̻�" & Space(100) & "1"
    cboEQCUTHREF.AddItem "�ʰ�" & Space(100) & "2"
            
    '/CUTOFF ���Ѱ� Equal ���� ����(1.����, 2.�̸�)
    cboEQCUTLREF.AddItem ""
    cboEQCUTLREF.AddItem "����" & Space(100) & "1"
    cboEQCUTLREF.AddItem "�̸�" & Space(100) & "2"
            
    '/CUTOFF �߰���(CUTOFF ��,���Ѱ��� �ٸ� ��� ���� 1.Grayzone, 2.Weakly positive, 3.Low Titer) �ڵ尪�� ����� �䱸�� ���� �� ����
    cboEQCUTMNM.AddItem ""
    cboEQCUTMNM.AddItem "Grayzone" & Space(100) & "1"
    cboEQCUTMNM.AddItem "Weakly positive" & Space(100) & "2"
    cboEQCUTMNM.AddItem "Low Titer" & Space(100) & "3"
            
    '/CUTOFF ������(1.Negative/Positive, 2.Neg/Pos. 3.Nonreactive/Reactive, 4.NEGATIVE/POSITIVE, 5.NEG/POS) �ڵ尪�� ����� �䱸�� ���� �� ����
    cboEQCUTOFFNM.AddItem ""
    cboEQCUTOFFNM.AddItem "Negative/Positive" & Space(100) & "1"
    cboEQCUTOFFNM.AddItem "Neg/Pos" & Space(100) & "2"
    cboEQCUTOFFNM.AddItem "Nonreactive/Reactive" & Space(100) & "3"
    cboEQCUTOFFNM.AddItem "NEGATIVE/POSITIVE" & Space(100) & "4"
    cboEQCUTOFFNM.AddItem "NEG/POS" & Space(100) & "5"
            
    '/CUTOFF ǥ������(1.Negative/Positive, 2.Negative/Positive(��ġ), 3.Negative/Grayzone(��ġ)/Positive(��ġ), 4.Negative(��ġ)/Grayzone(��ġ)/Positive(��ġ)
    cboEQCUTRTYPE.AddItem ""
    cboEQCUTRTYPE.AddItem "Negative/Positive" & Space(100) & "1"
    cboEQCUTRTYPE.AddItem "Negative/Positive(��ġ)" & Space(100) & "2"
    cboEQCUTRTYPE.AddItem "Negative/Grayzone(��ġ)/Positive(��ġ)" & Space(100) & "3"
    cboEQCUTRTYPE.AddItem "Negative(��ġ)/Grayzone(��ġ)/Positive(��ġ)" & Space(100) & "4"
Return
End Sub

Public Function MM_INPUT() As Boolean

End Function

Private Sub SUB_MM_KEY_CLEAR()
    fra�󼼳���.Enabled = False
    
    txtEQNM = ""
    cboEQORDYN.ListIndex = 0
    txtEQUNIT = ""
    mskEQRSTRANGE = "0"
    mskEQSEQ = ""
    
    txtEQRMHVAL = ""
    txtEQRMLVAL = ""
    txtEQRFHVAL = ""
    txtEQRFLVAL = ""

    cboEQRMHREF.ListIndex = -1      '/Reference ���� High Reference(1.�̻�, 2.�ʰ�, 3.����, 4.�̸�)
    cboEQRMLREF.ListIndex = -1      '/Reference ���� Low Reference(1.�̻�, 2.�ʰ�, 3.����, 4.�̸�)
    cboEQRFHREF.ListIndex = -1      '/Reference ���� High Reference(1.�̻�, 2.�ʰ�, 3.����, 4.�̸�)
    cboEQRFLREF.ListIndex = -1      '/Reference ���� Low Reference(1.�̻�, 2.�ʰ�, 3.����, 4.�̸�)
    
    cboEQLIMITFLAG1.ListIndex = 0   '/LIMIT ���Ѱ� ����(0.������, 1.����, 2.�̸�)
    cboEQLIMITFLAG2.ListIndex = 0   '/LIMIT ���Ѱ� ����(0.������, 1.�̻�, 2.�ʰ�)
    
    cboEQCUTOFFGB.ListIndex = 0     '/CUTOFF ���뱸��
    cboEQCUTHREF.ListIndex = -1     '/CUTOFF ���Ѱ� Equal ���� ����(1.�̻�, 2.�ʰ�)
    cboEQCUTLREF.ListIndex = -1     '/CUTOFF ���Ѱ� Equal ���� ����(1.����, 2.�̸�)
    cboEQCUTMNM.ListIndex = -1      '/CUTOFF �߰���(CUTOFF ��,���Ѱ��� �ٸ� ��� ���� 1.Grayzone, 2.Weakly positive, 3.Low Titer) �ڵ尪�� ����� �䱸�� ���� �� ����
    cboEQCUTOFFNM.ListIndex = -1    '/CUTOFF ������(1.Negative/Positive, 2.Neg/Pos. 3.Nonreactive/Reactive, 4.NEGATIVE/POSITIVE, 5.NEG/POS) �ڵ尪�� ����� �䱸�� ���� �� ����
    cboEQCUTRTYPE.ListIndex = -1    '/CUTOFF ǥ������(1.Negative/Positive, 2.Negative/Positive(��ġ), 3.Negative/Grayzone(��ġ)/Positive(��ġ), 4.Negative(��ġ)/Grayzone(��ġ)/Positive(��ġ)
    
    If sprEXCD.MaxRows > 0 Then sprEXCD.MaxRows = 0: sprEXCD.MaxRows = 1
End Sub

Public Function MM_PRINT() As Boolean

End Function

Public Function FUNC_MM_SAVE() As Boolean
    FUNC_MM_SAVE = False
    
On Error GoTo RTN_ERR

    If ConnDB_LOC = False Then Exit Function
    
    ADC_LOC.BeginTrans
    
    gstrQuy = "SELECT * "
    gstrQuy = gstrQuy & vbCrLf & "  FROM EQ_MST "
    gstrQuy = gstrQuy & vbCrLf & " WHERE EQCD = '" & Trim(txtEQCD) & "' "
    If ReadSQL_LOC(gstrQuy, ADR_LOC) = False Then ADC_LOC.RollbackTrans: Call CloseDB_LOC: Exit Function
    
    If Not ADR_LOC Is Nothing Then
        ADR_LOC.Close: Set ADR_LOC = Nothing
    
        gstrQuy = "UPDATE EQ_MST SET "
        gstrQuy = gstrQuy & vbCrLf & "       EQNM           = '" & Trim(TEXT_LSET(txtEQNM, 50)) & "', "     '/���˻��
        gstrQuy = gstrQuy & vbCrLf & "       EQUNIT         = '" & Trim(TEXT_LSET(txtEQUNIT, 10)) & "', "   '/�������
        gstrQuy = gstrQuy & vbCrLf & "       EQSEQ          =  " & Val(mskEQSEQ) & ",  "                    '/ȭ�����ļ���
        gstrQuy = gstrQuy & vbCrLf & "       EQRSTRANGE     =  " & Val(mskEQRSTRANGE) & ", "                '/���˻��� �Ҽ���ǥ��( 0: ��üǥ��, >1 : ���ڸ�ŭ ǥ��)
        gstrQuy = gstrQuy & vbCrLf & "       EQLIMITFLAG1   = '" & Trim(Right(cboEQLIMITFLAG1, 10)) & "', " '/LIMIT ���Ѱ� ����(0.������, 1.����, 2.�̸�)
        gstrQuy = gstrQuy & vbCrLf & "       EQLIMITVALUE1  = '" & Trim(txtEQLIMITVALUE1) & "', "           '/LIMIT ���Ѱ�
        gstrQuy = gstrQuy & vbCrLf & "       EQLIMITFLAG2   = '" & Trim(Right(cboEQLIMITFLAG2, 10)) & "', " '/LIMIT ���Ѱ� ����(0.������, 1.�̻�, 2.�ʰ�)
        gstrQuy = gstrQuy & vbCrLf & "       EQLIMITVALUE2  = '" & Trim(txtEQLIMITVALUE2) & "', "           '/LIMIT ���Ѱ�
        gstrQuy = gstrQuy & vbCrLf & "       EQCUTOFFGB     = '" & Trim(Right(cboEQCUTOFFGB, 10)) & "', "   '/CUTOFF ���뱸��(0.�������, 1.���� Positive, 2.���� Positive, 3.�����������(��ġ�� CutOff���� ���ÿ� ���ð��)
        gstrQuy = gstrQuy & vbCrLf & "       EQCUTOFFNM     = '" & Trim(Right(cboEQCUTOFFNM, 10)) & "', "   '/CUTOFF ������(1.Negative/Positive, 2.Neg/Pos. 3.Nonreactive/Reactive, 4.NEGATIVE/POSITIVE, 5.NEG/POS) �ڵ尪�� ����� �䱸�� ���� �� ����
        gstrQuy = gstrQuy & vbCrLf & "       EQCUTHVAL      = '" & Trim(txtEQCUTHVAL) & "', "               '/CUTOFF ���Ѱ�
        gstrQuy = gstrQuy & vbCrLf & "       EQCUTHREF      = '" & Trim(Right(cboEQCUTHREF, 10)) & "', "    '/CUTOFF ���Ѱ� Equal ���� ����(1.�̻�, 2.�ʰ�)
        gstrQuy = gstrQuy & vbCrLf & "       EQCUTLVAL      = '" & Trim(txtEQCUTLVAL) & "', "               '/CUTOFF ���Ѱ�
        gstrQuy = gstrQuy & vbCrLf & "       EQCUTLREF      = '" & Trim(Right(cboEQCUTLREF, 10)) & "', "    '/CUTOFF ���Ѱ� Equal ���� ����(1.����, 2.�̸�)
        gstrQuy = gstrQuy & vbCrLf & "       EQCUTMNM       = '" & Trim(Right(cboEQCUTMNM, 10)) & "', "     '/CUTOFF �߰���(CUTOFF ��,���Ѱ��� �ٸ� ��� ���� 1.Grayzone, 2.Weakly positive, 3.Low Titer) �ڵ尪�� ����� �䱸�� ���� �� ����
        gstrQuy = gstrQuy & vbCrLf & "       EQCUTRTYPE     = '" & Trim(Right(cboEQCUTRTYPE, 10)) & "', "   '/CUTOFF ǥ������(1.Negative/Positive, 2.Negative/Positive(��ġ), 3.Negative/Grayzone(��ġ)/Positive(��ġ), 4.Negative(��ġ)/Grayzone(��ġ)/Positive(��ġ)
        gstrQuy = gstrQuy & vbCrLf & "       EQRMHVAL       = '" & Trim(txtEQRMHVAL) & "', "                '/Reference ���� High Value
        gstrQuy = gstrQuy & vbCrLf & "       EQRMHREF       = '" & Trim(Right(cboEQRMHREF, 10)) & "', "     '/Reference ���� High Reference(1.�̻�, 2.�ʰ�, 3.����, 4.�̸�)
        gstrQuy = gstrQuy & vbCrLf & "       EQRMLVAL       = '" & Trim(txtEQRMLVAL) & "', "                '/Reference ���� Low Value
        gstrQuy = gstrQuy & vbCrLf & "       EQRMLREF       = '" & Trim(Right(cboEQRMLREF, 10)) & "', "     '/Reference ���� Low Reference(1.�̻�, 2.�ʰ�, 3.����, 4.�̸�)
        gstrQuy = gstrQuy & vbCrLf & "       EQRFHVAL       = '" & Trim(txtEQRFHVAL) & "', "                '/Reference ���� High Value
        gstrQuy = gstrQuy & vbCrLf & "       EQRFHREF       = '" & Trim(Right(cboEQRFHREF, 10)) & "', "     '/Reference ���� High Reference(1.�̻�, 2.�ʰ�, 3.����, 4.�̸�)
        gstrQuy = gstrQuy & vbCrLf & "       EQRFLVAL       = '" & Trim(txtEQRFLVAL) & "', "                '/Reference ���� Low Value
        gstrQuy = gstrQuy & vbCrLf & "       EQRFLREF       = '" & Trim(Right(cboEQRFLREF, 10)) & "' "      '/Reference ���� Low Reference(1.�̻�, 2.�ʰ�, 3.����, 4.�̸�)
        gstrQuy = gstrQuy & vbCrLf & " WHERE EQCD           = '" & Trim(txtEQCD) & "' "
    Else
        gstrQuy = "INSERT INTO EQ_MST "
        gstrQuy = gstrQuy & vbCrLf & " (EQCD,           EQNM,           EQUNIT,         EQSEQ,          EQRSTRANGE, "
        gstrQuy = gstrQuy & vbCrLf & "  EQLIMITFLAG1,   EQLIMITVALUE1,  EQLIMITFLAG2,   EQLIMITVALUE2,  EQCUTOFFGB, "
        gstrQuy = gstrQuy & vbCrLf & "  EQCUTOFFNM,     EQCUTHVAL,      EQCUTHREF,      EQCUTLVAL,      EQCUTLREF, "
        gstrQuy = gstrQuy & vbCrLf & "  EQCUTMNM,       EQCUTRTYPE,     EQRMHVAL,       EQRMHREF,       EQRMLVAL, "
        gstrQuy = gstrQuy & vbCrLf & "  EQRMLREF,       EQRFHVAL,       EQRFHREF,       EQRFLVAL,       EQRFLREF) "
        gstrQuy = gstrQuy & vbCrLf & " VALUES "
        gstrQuy = gstrQuy & vbCrLf & " ('" & Trim(TEXT_LSET(txtEQCD, 10)) & "', "     '/���˻��ڵ�
        gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(TEXT_LSET(txtEQNM, 50)) & "', "     '/���˻��
        gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(TEXT_LSET(txtEQUNIT, 10)) & "', "   '/�������
        gstrQuy = gstrQuy & vbCrLf & "   " & Val(mskEQSEQ) & ",  "                    '/ȭ�����ļ���
        gstrQuy = gstrQuy & vbCrLf & "   " & Val(mskEQRSTRANGE) & ", "                '/���˻��� �Ҽ���ǥ��( 0: ��üǥ��, >1 : ���ڸ�ŭ ǥ��)
        
        gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(Right(cboEQLIMITFLAG1, 10)) & "', " '/LIMIT ���Ѱ� ����(0.������, 1.����, 2.�̸�)
        gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(txtEQLIMITVALUE1) & "', "           '/LIMIT ���Ѱ�
        gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(Right(cboEQLIMITFLAG2, 10)) & "', " '/LIMIT ���Ѱ� ����(0.������, 1.�̻�, 2.�ʰ�)
        gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(txtEQLIMITVALUE2) & "', "           '/LIMIT ���Ѱ�
        gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(Right(cboEQCUTOFFGB, 10)) & "', "   '/CUTOFF ���뱸��(0.�������, 1.���� Positive, 2.���� Positive, 3.�����������(��ġ�� CutOff���� ���ÿ� ���ð��)
        
        gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(Right(cboEQCUTOFFNM, 10)) & "', "   '/CUTOFF ������(1.Negative/Positive, 2.Neg/Pos. 3.Nonreactive/Reactive, 4.NEGATIVE/POSITIVE, 5.NEG/POS) �ڵ尪�� ����� �䱸�� ���� �� ����
        gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(txtEQCUTHVAL) & "', "               '/CUTOFF ���Ѱ�
        gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(Right(cboEQCUTHREF, 10)) & "', "    '/CUTOFF ���Ѱ� Equal ���� ����(1.�̻�, 2.�ʰ�)
        gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(txtEQCUTLVAL) & "', "               '/CUTOFF ���Ѱ�
        gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(Right(cboEQCUTLREF, 10)) & "', "    '/CUTOFF ���Ѱ� Equal ���� ����(1.����, 2.�̸�)
        
        gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(Right(cboEQCUTMNM, 10)) & "', "     '/CUTOFF �߰���(CUTOFF ��,���Ѱ��� �ٸ� ��� ���� 1.Grayzone, 2.Weakly positive, 3.Low Titer) �ڵ尪�� ����� �䱸�� ���� �� ����
        gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(Right(cboEQCUTRTYPE, 10)) & "', "   '/CUTOFF ǥ������(1.Negative/Positive, 2.Negative/Positive(��ġ), 3.Negative/Grayzone(��ġ)/Positive(��ġ), 4.Negative(��ġ)/Grayzone(��ġ)/Positive(��ġ)
        gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(txtEQRMHVAL) & "', "                '/Reference ���� High Value
        gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(Right(cboEQRMHREF, 10)) & "', "     '/Reference ���� High Reference(1.�̻�, 2.�ʰ�, 3.����, 4.�̸�)
        gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(txtEQRMLVAL) & "', "                '/Reference ���� Low Value
        
        gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(Right(cboEQRMLREF, 10)) & "', "     '/Reference ���� Low Reference(1.�̻�, 2.�ʰ�, 3.����, 4.�̸�)
        gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(txtEQRFHVAL) & "', "                '/Reference ���� High Value
        gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(Right(cboEQRFHREF, 10)) & "', "     '/Reference ���� High Reference(1.�̻�, 2.�ʰ�, 3.����, 4.�̸�)
        gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(txtEQRFLVAL) & "', "                '/Reference ���� Low Value
        gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(Right(cboEQRFLREF, 10)) & "') "     '/Reference ���� Low Reference(1.�̻�, 2.�ʰ�, 3.����, 4.�̸�)
    End If
    If RunSQL_LOC(gstrQuy) = False Then ADC_LOC.RollbackTrans: Call CloseDB_LOC: Exit Function
    
    gstrQuy = "DELETE FROM EX_MST "
    gstrQuy = gstrQuy & vbCrLf & " WHERE EQCD = '" & Trim(txtEQCD) & "' "
    If RunSQL_LOC(gstrQuy) = False Then ADC_LOC.RollbackTrans: Call CloseDB_LOC: Exit Function
    
    For intX = 1 To sprEXCD.DataRowCnt
        If Trim(GET_CELL(sprEXCD, 1, intX)) <> "" Then
            gstrQuy = "SELECT * "
            gstrQuy = gstrQuy & vbCrLf & "  FROM EX_MST "
            gstrQuy = gstrQuy & vbCrLf & " WHERE EQCD = '" & Trim(txtEQCD) & "' "
            gstrQuy = gstrQuy & vbCrLf & "   AND EXCD = '" & Trim(GET_CELL(sprEXCD, 1, intX)) & "' "
            If ReadSQL_LOC(gstrQuy, ADR_LOC) = False Then ADC_LOC.RollbackTrans: Call CloseDB_LOC: Exit Function
            
            If Not ADR_LOC Is Nothing Then
                ADR_LOC.Close: Set ADR_LOC = Nothing
            Else
                gstrQuy = "INSERT INTO EX_MST (EQCD, EXCD, EQORDREADYN, EQRESSENDYN) "
                gstrQuy = gstrQuy & vbCrLf & " VALUES "
                gstrQuy = gstrQuy & vbCrLf & " ('" & Trim(txtEQCD) & "',"                                           '/���˻��ڵ�
                gstrQuy = gstrQuy & vbCrLf & "  '" & Trim(GET_CELL(sprEXCD, 1, intX)) & "', "                       '/ó���ڵ�
                gstrQuy = gstrQuy & vbCrLf & "  '" & IIf(Trim(GET_CELL(sprEXCD, 2, intX)) = "1", "Y", "N") & "', "  '/HISó���б⿩��(Y.���, N.����)
                gstrQuy = gstrQuy & vbCrLf & "  '" & IIf(Trim(GET_CELL(sprEXCD, 3, intX)) = "1", "Y", "N") & "') "   '/���˻��� HIS���ۿ���(Y.���, N.����)
                If RunSQL_LOC(gstrQuy) = False Then ADC_LOC.RollbackTrans: Call CloseDB_LOC: Exit Function
            End If
        End If
    Next intX
    
    ADC_LOC.CommitTrans
    
    Call CloseDB_LOC
    
    FUNC_MM_SAVE = True
Exit Function

'/----------------------------------------------------------------------------------------------------/

RTN_ERR:

End Function

Public Function FUNC_MM_VIEW() As Boolean
    FUNC_MM_VIEW = False
    
On Error GoTo RTN_ERR

    If Trim(txtEQCD) = "" Then Exit Function
    
    Call SUB_MM_KEY_CLEAR
    
    If ConnDB_LOC = False Then Exit Function
    
    gstrQuy = "SELECT * "
    gstrQuy = gstrQuy & vbCrLf & "  FROM EQ_MST "
    gstrQuy = gstrQuy & vbCrLf & " WHERE EQCD = '" & Trim(txtEQCD) & "' "
    If ReadSQL_LOC(gstrQuy, ADR_LOC) = False Then Call CloseDB_LOC: Exit Function
    
    If Not ADR_LOC Is Nothing Then
        If gstrInputUpdate = "2" Then '/1.Input, 2.Update
            fra�󼼳���.Enabled = True
            
            txtEQNM = Trim(ADR_LOC!EQNM & "")               '/���˻��
            Call SET_CBO_DT_R(Trim(ADR_LOC!EQORDYN & ""), cboEQORDYN) '/���˻��ڵ� ������ۿ���(Y.����, N.������)
            txtEQUNIT = Trim(ADR_LOC!EQUNIT & "")           '/�˻�������
            mskEQSEQ = Trim(ADR_LOC!EQSEQ & "")             '/ȭ�����ļ���
            mskEQRSTRANGE = Trim(ADR_LOC!EQRSTRANGE & "")   '/�Ҽ����ڸ���
            
            txtEQRMLVAL = Trim(ADR_LOC!EQRMLVAL & "")                   '/Reference ���� Low Value
            Call SET_CBO_DT_R(Trim(ADR_LOC!EQRMLREF & ""), cboEQRMLREF) '/Reference ���� Low Reference(1.�̻�, 2.�ʰ�, 3.����, 4.�̸�)
            txtEQRMHVAL = Trim(ADR_LOC!EQRMHVAL & "")                   '/Reference ���� High Value
            Call SET_CBO_DT_R(Trim(ADR_LOC!EQRMHREF & ""), cboEQRMHREF) '/Reference ���� High Reference(1.�̻�, 2.�ʰ�, 3.����, 4.�̸�)
            txtEQRFLVAL = Trim(ADR_LOC!EQRFLVAL & "")                   '/Reference ���� Low Value
            Call SET_CBO_DT_R(Trim(ADR_LOC!EQRFLREF & ""), cboEQRFLREF) '/Reference ���� Low Reference(1.�̻�, 2.�ʰ�, 3.����, 4.�̸�)
            txtEQRFHVAL = Trim(ADR_LOC!EQRFHVAL & "")                   '/Reference ���� High Value
            Call SET_CBO_DT_R(Trim(ADR_LOC!EQRFHREF & ""), cboEQRFHREF) '/Reference ���� High Reference(1.�̻�, 2.�ʰ�, 3.����, 4.�̸�)
        
            Call SET_CBO_DT_R(Trim(ADR_LOC!EQLIMITFLAG1 & ""), cboEQLIMITFLAG1) '/LIMIT ���Ѱ� ����(0.������, 1.����, 2.�̸�)
            txtEQLIMITVALUE1 = Trim(ADR_LOC!EQLIMITVALUE1 & "")                 '/LIMIT ���Ѱ�
            Call SET_CBO_DT_R(Trim(ADR_LOC!EQLIMITFLAG2 & ""), cboEQLIMITFLAG2) '/LIMIT ���Ѱ� ����(0.������, 1.�̻�, 2.�ʰ�)
            txtEQLIMITVALUE2 = Trim(ADR_LOC!EQLIMITVALUE2 & "")                 '/LIMIT ���Ѱ�
        
            Call SET_CBO_DT_R(Trim(ADR_LOC!EQCUTOFFGB & ""), cboEQCUTOFFGB) '/CUTOFF ����(0.�������, 1.���� Positive, 2.���� Positive, 3.�����������(��ġ�� CutOff���� ���ÿ� ���ð��)
            txtEQCUTHVAL = Trim(ADR_LOC!EQCUTHVAL & "")                     '/CUTOFF ���Ѱ�
            Call SET_CBO_DT_R(Trim(ADR_LOC!EQCUTHREF & ""), cboEQCUTHREF)   '/CUTOFF ���Ѱ� Equal ���� ����(1.�̻�, 2.�ʰ�)
            txtEQCUTLVAL = Trim(ADR_LOC!EQCUTLVAL & "")                     '/CUTOFF ���Ѱ�
            Call SET_CBO_DT_R(Trim(ADR_LOC!EQCUTLREF & ""), cboEQCUTLREF)   '/CUTOFF ���Ѱ� Equal ���� ����(1.����, 2.�̸�)
            Call SET_CBO_DT_R(Trim(ADR_LOC!EQCUTMNM & ""), cboEQCUTMNM)     '/CUTOFF �߰���(CUTOFF ��,���Ѱ��� �ٸ� ��� ���� 1.Grayzone, 2.Weakly positive, 3.Low Titer) �ڵ尪�� ����� �䱸�� ���� �� ����
            Call SET_CBO_DT_R(Trim(ADR_LOC!EQCUTOFFNM & ""), cboEQCUTOFFNM) '/CUTOFF ������(1.Negative/Positive, 2.Neg/Pos. 3.Nonreactive/Reactive, 4.NEGATIVE/POSITIVE, 5.NEG/POS) �ڵ尪�� ����� �䱸�� ���� �� ����
            Call SET_CBO_DT_R(Trim(ADR_LOC!EQCUTRTYPE & ""), cboEQCUTRTYPE) '/CUTOFF ǥ������(1.Negative/Positive, 2.Negative/Positive(��ġ), 3.Negative/Grayzone(��ġ)/Positive(��ġ), 4.Negative(��ġ)/Grayzone(��ġ)/Positive(��ġ)
        Else
            MsgBox "�����ڷᰡ �ֽ��ϴ�!", vbInformation, "Ȯ��"
        End If
        ADR_LOC.Close: Set ADR_LOC = Nothing
        
        With sprEXCD
            If .MaxRows > 0 Then .MaxRows = 0
            
            gstrQuy = "SELECT * "
            gstrQuy = gstrQuy & vbCrLf & "  FROM EX_MST "
            gstrQuy = gstrQuy & vbCrLf & " WHERE EQCD = '" & Trim(txtEQCD) & "' "
            gstrQuy = gstrQuy & vbCrLf & " ORDER BY EXCD "
            If ReadSQL_LOC(gstrQuy, ADR_LOC) = False Then Call CloseDB_LOC: Exit Function
            
            If Not ADR_LOC Is Nothing Then
                Do Until ADR_LOC.EOF
                    .MaxRows = .MaxRows + 1: .Row = .MaxRows
                
                    .Col = 1: .Text = Trim(ADR_LOC!EXCD & "")
                    .Col = 2: .Text = IIf(Trim(ADR_LOC!EQORDREADYN & "") = "Y", "1", "0")
                    .Col = 3: .Text = IIf(Trim(ADR_LOC!EQRESSENDYN & "") = "Y", "1", "0")
                    
                    ADR_LOC.MoveNext
                Loop
            End If
            
            .MaxRows = .MaxRows + 1
        End With
    Else
        If gstrInputUpdate = "1" Then '/1.Input, 2.Update
            fra�󼼳���.Enabled = True
            
            txtEQNM.SetFocus
        End If
    End If

    Call CloseDB_LOC
    
    FUNC_MM_VIEW = True
Exit Function

'/----------------------------------------------------------------------------------------------------/

RTN_ERR:

End Function

Private Sub cboEQCUTHREF_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cboEQCUTLREF_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cboEQCUTMNM_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cboEQCUTOFFGB_Click()
    '/CUTOFF ���뱸��(0.�������, 1.���� Positive, 2.���� Positive, 3.�����������(��ġ�� CutOff���� ���ÿ� ���ð��)
    If Trim(Right(cboEQCUTOFFGB, 10)) = "0" Then
        txtEQCUTHVAL = ""               '/CUTOFF ���Ѱ�
        cboEQCUTHREF.ListIndex = -1     '/CUTOFF ���Ѱ� Equal ���� ����(1.�̻�, 2.�ʰ�)
        txtEQCUTLVAL = ""               '/CUTOFF ���Ѱ�
        cboEQCUTLREF.ListIndex = -1     '/CUTOFF ���Ѱ� Equal ���� ����(1.����, 2.�̸�)
        cboEQCUTMNM.ListIndex = -1      '/CUTOFF �߰���(CUTOFF ��,���Ѱ��� �ٸ� ��� ���� 1.Grayzone, 2.Weakly positive, 3.Low Titer) �ڵ尪�� ����� �䱸�� ���� �� ����
        cboEQCUTOFFNM.ListIndex = -1    '/CUTOFF ������(1.Negative/Positive, 2.Neg/Pos. 3.Nonreactive/Reactive, 4.NEGATIVE/POSITIVE, 5.NEG/POS) �ڵ尪�� ����� �䱸�� ���� �� ����
        cboEQCUTRTYPE.ListIndex = -1    '/CUTOFF ǥ������(1.Negative/Positive, 2.Negative/Positive(��ġ), 3.Negative/Grayzone(��ġ)/Positive(��ġ), 4.Negative(��ġ)/Grayzone(��ġ)/Positive(��ġ)
    
        txtEQCUTHVAL.Enabled = False
        cboEQCUTHREF.Enabled = False
        txtEQCUTLVAL.Enabled = False
        cboEQCUTLREF.Enabled = False
        cboEQCUTMNM.Enabled = False
        cboEQCUTOFFNM.Enabled = False
        cboEQCUTRTYPE.Enabled = False
    Else
        txtEQCUTHVAL.Enabled = True
        cboEQCUTHREF.Enabled = True
        txtEQCUTLVAL.Enabled = True
        cboEQCUTLREF.Enabled = True
        cboEQCUTMNM.Enabled = True
        cboEQCUTOFFNM.Enabled = True
        cboEQCUTRTYPE.Enabled = True
    End If
End Sub

Private Sub cboEQCUTOFFGB_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cboEQCUTOFFNM_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cboEQCUTRTYPE_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cboEQLIMITFLAG1_Click()
    '/LIMIT ���Ѱ� ����(0.������, 1.����, 2.�̸�)
    If Trim(Right(cboEQLIMITFLAG1, 10)) = "0" Then
        txtEQLIMITVALUE1 = ""
        txtEQLIMITVALUE1.Enabled = False
    Else
        txtEQLIMITVALUE1.Enabled = True
    End If
End Sub

Private Sub cboEQLIMITFLAG1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cboEQLIMITFLAG2_Click()
    '/LIMIT ���Ѱ� ����(0.������, 1.�̻�, 2.�ʰ�)
    If Trim(Right(cboEQLIMITFLAG2, 10)) = "0" Then
        txtEQLIMITVALUE2 = ""
        txtEQLIMITVALUE2.Enabled = False
    Else
        txtEQLIMITVALUE2.Enabled = True
    End If
End Sub

Private Sub cboEQLIMITFLAG2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cboEQORDYN_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cboEQRFHREF_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cboEQRFLREF_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cboEQRMHREF_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cboEQRMLREF_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    If Trim(txtEQCD) = "" Then MsgBox "���˻��ڵ带 (��)�Է��Ͻʽÿ�!", vbInformation, "Ȯ��": txtEQCD.SetFocus: Exit Sub
    
    If IsNumeric(mskEQSEQ) = False Then
        MsgBox "ȭ�����ļ����� (��)�Է��Ͻʽÿ�!" & vbCrLf & _
               "�Է������� ����Ÿ���Դϴ�.", vbInformation, "Ȯ��": mskEQSEQ.SetFocus: Exit Sub
    End If
    
    If IsNumeric(mskEQRSTRANGE) = False Then
        MsgBox "�Ҽ������� ǥ�� �ڸ����� (��)�Է��Ͻʽÿ�!" & vbCrLf & _
               "�Է������� ����Ÿ���Դϴ�.", vbInformation, "Ȯ��": mskEQRSTRANGE.SetFocus: Exit Sub
    End If
    
    '/��������ġ(��Low)
    If Trim(txtEQRMLVAL) = "" Or Trim(cboEQRMLREF) = "" Then '/���� �ϳ��� �����϶�
        If Not (Trim(txtEQRMLVAL) = "" And Trim(cboEQRMLREF) = "") Then '/�Ѵ� ������ �ƴҶ�
            If Trim(txtEQRMLVAL) = "" Then MsgBox "[��������ġ] �� Low ���� (��)�Է��Ͻʽÿ�!", vbInformation, "Ȯ��": txtEQRMLVAL.SetFocus: Exit Sub
            If Trim(cboEQRMLREF) = "" Then MsgBox "[��������ġ] �� Low ������ (��)�Է��Ͻʽÿ�!", vbInformation, "Ȯ��": cboEQRMLREF.SetFocus: Exit Sub
        End If
    End If
    
    '/��������ġ(��High)
    If Trim(txtEQRMHVAL) = "" Or Trim(cboEQRMHREF) = "" Then '/���� �ϳ��� �����϶�
        If Not (Trim(txtEQRMHVAL) = "" And Trim(cboEQRMHREF) = "") Then '/�Ѵ� ������ �ƴҶ�
            If Trim(txtEQRMHVAL) = "" Then MsgBox "[��������ġ] �� High ���� (��)�Է��Ͻʽÿ�!", vbInformation, "Ȯ��": txtEQRMHVAL.SetFocus: Exit Sub
            If Trim(cboEQRMHREF) = "" Then MsgBox "[��������ġ] �� High ������ (��)�Է��Ͻʽÿ�!", vbInformation, "Ȯ��": cboEQRMHREF.SetFocus: Exit Sub
        End If
    End If
    
    '/��������ġ(��Low)
    If Trim(txtEQRFLVAL) = "" Or Trim(cboEQRFLREF) = "" Then '/���� �ϳ��� �����϶�
        If Not (Trim(txtEQRFLVAL) = "" And Trim(cboEQRFLREF) = "") Then '/�Ѵ� ������ �ƴҶ�
            If Trim(txtEQRFLVAL) = "" Then MsgBox "[��������ġ] �� Low ���� (��)�Է��Ͻʽÿ�!", vbInformation, "Ȯ��": txtEQRFLVAL.SetFocus: Exit Sub
            If Trim(cboEQRFLREF) = "" Then MsgBox "[��������ġ] �� Low ������ (��)�Է��Ͻʽÿ�!", vbInformation, "Ȯ��": cboEQRFLREF.SetFocus: Exit Sub
        End If
    End If
    
    '/��������ġ(��High)
    If Trim(txtEQRFHVAL) = "" Or Trim(cboEQRFHREF) = "" Then '/���� �ϳ��� �����϶�
        If Not (Trim(txtEQRFHVAL) = "" And Trim(cboEQRFHREF) = "") Then '/�Ѵ� ������ �ƴҶ�
            If Trim(txtEQRFHVAL) = "" Then MsgBox "[��������ġ] �� High ���� (��)�Է��Ͻʽÿ�!", vbInformation, "Ȯ��": txtEQRFHVAL.SetFocus: Exit Sub
            If Trim(cboEQRFHREF) = "" Then MsgBox "[��������ġ] �� High ������ (��)�Է��Ͻʽÿ�!", vbInformation, "Ȯ��": cboEQRFHREF.SetFocus: Exit Sub
        End If
    End If
    
    '/CUTOFF ���Ѱ� Equal ���� ����(1.�̻�, 2.�ʰ�)
    If Trim(txtEQCUTHVAL) = "" Or Trim(cboEQCUTHREF) = "" Then '/���� �ϳ��� �����϶�
        If Not (Trim(txtEQCUTHVAL) = "" And Trim(cboEQCUTHREF) = "") Then '/�Ѵ� ������ �ƴҶ�
            If Trim(txtEQCUTHVAL) = "" Then MsgBox "CUTOFF ���Ѱ��� (��)�Է��Ͻʽÿ�!", vbInformation, "Ȯ��": txtEQCUTHVAL.SetFocus: Exit Sub
            If Trim(cboEQCUTHREF) = "" Then MsgBox "CUTOFF ���� ������ (��)�Է��Ͻʽÿ�!", vbInformation, "Ȯ��": cboEQCUTHREF.SetFocus: Exit Sub
        End If
    End If
    
    '/CUTOFF ���Ѱ� Equal ���� ����(1.����, 2.�̸�)
    If Trim(txtEQCUTLVAL) = "" Or Trim(cboEQCUTLREF) = "" Then '/���� �ϳ��� �����϶�
        If Not (Trim(txtEQCUTLVAL) = "" And Trim(cboEQCUTLREF) = "") Then '/�Ѵ� ������ �ƴҶ�
            If Trim(txtEQCUTLVAL) = "" Then MsgBox "CUTOFF ���Ѱ��� (��)�Է��Ͻʽÿ�!", vbInformation, "Ȯ��": txtEQCUTLVAL.SetFocus: Exit Sub
            If Trim(cboEQCUTLREF) = "" Then MsgBox "CUTOFF ���� ������ (��)�Է��Ͻʽÿ�!", vbInformation, "Ȯ��": cboEQCUTLREF.SetFocus: Exit Sub
        End If
    End If
    
    If MsgBox("�����ϰڽ��ϱ�?", vbQuestion + vbOKCancel, "��������") = vbCancel Then Exit Sub
    
    If FUNC_MM_SAVE = True Then
        gstrInputUpdateYN = True '/���� �������� Set(��������)
        
        MsgBox "����Ǿ����ϴ�!", vbInformation, "Ȯ��"
        
        Call SUB_MM_CANCEL
        
        If gstrInputUpdate = "1" Then '/1.Input, 2.Update(�ű� �� ���� �� ȭ��ó��)
            txtEQCD.SetFocus
        Else
            Unload Me
        End If
    Else
        gstrInputUpdateYN = False '/���� �������� Set(��������)
        MsgBox "�������!", vbCritical, "Ȯ��"
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: Unload Me
    End Select
End Sub

Private Sub Form_Load()
    Call SUB_MM_INITIAL
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CloseDB_LOC
    Set frmEQ����_���˻��ڵ����_�Է� = Nothing
End Sub

Private Sub sprEXCD_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    If ChangeMade = False Then Exit Sub
    If Col <> 1 Then Exit Sub
    
    If sprEXCD.DataRowCnt = sprEXCD.MaxRows Then
        sprEXCD.MaxRows = sprEXCD.MaxRows + 1
    End If
End Sub

Private Sub txtEQCD_Change()
    Call SUB_MM_KEY_CLEAR
End Sub

Private Sub txtEQCD_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub txtEQCD_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call FUNC_MM_VIEW
End Sub

Private Sub txtEQCUTHVAL_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub txtEQCUTHVAL_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtEQCUTLVAL_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub txtEQCUTLVAL_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtEQLIMITVALUE1_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub txtEQLIMITVALUE1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtEQLIMITVALUE2_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub txtEQLIMITVALUE2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtEQRFHVAL_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub txtEQRFHVAL_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtEQRFLVAL_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub txtEQRFLVAL_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtEQRMHVAL_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub txtEQRMHVAL_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtEQRMLVAL_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub txtEQRMLVAL_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtEQNM_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub txtEQNM_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub mskEQSEQ_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub mskEQSEQ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtEQUNIT_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub txtEQUNIT_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub mskEQRSTRANGE_GotFocus()
    Call TEXTGF(Me.ActiveControl)
End Sub

Private Sub mskEQRSTRANGE_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub
