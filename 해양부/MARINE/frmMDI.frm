VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.MDIForm frmMDI 
   BackColor       =   &H00FFFFFF&
   Caption         =   "�ѱ��ؾ�������ȸ �ڷ������Ȳ ����͸�"
   ClientHeight    =   9810
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   20100
   Icon            =   "frmMDI.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows �⺻��
   WindowState     =   2  '�ִ�ȭ
   Begin VB.PictureBox picMenu 
      Align           =   1  '�� ����
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   20040
      TabIndex        =   0
      Top             =   0
      Width           =   20100
      Begin VB.Timer tmrNow 
         Left            =   13650
         Top             =   90
      End
      Begin VB.CommandButton cmdView 
         BackColor       =   &H00C0FFFF&
         Caption         =   "����������"
         BeginProperty Font 
            Name            =   "���� ���"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   0
         Left            =   150
         Style           =   1  '�׷���
         TabIndex        =   3
         Top             =   60
         Width           =   1575
      End
      Begin VB.CommandButton cmdView 
         BackColor       =   &H00FFFFFF&
         Caption         =   "�ؾ��������"
         BeginProperty Font 
            Name            =   "���� ���"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   1
         Left            =   1740
         Style           =   1  '�׷���
         TabIndex        =   2
         Top             =   60
         Width           =   1575
      End
      Begin VB.CommandButton cmdView 
         BackColor       =   &H00FFFFFF&
         Caption         =   "�����������ڷ�"
         BeginProperty Font 
            Name            =   "���� ���"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   2
         Left            =   3330
         MaskColor       =   &H00FFFFFF&
         Style           =   1  '�׷���
         TabIndex        =   1
         Top             =   60
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker dtpToday 
         Height          =   345
         Left            =   7530
         TabIndex        =   4
         Top             =   120
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "���� ���"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   135987200
         CurrentDate     =   43884
      End
      Begin MSComCtl2.DTPicker dtpTotime 
         Height          =   345
         Left            =   10140
         TabIndex        =   6
         Top             =   120
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "���� ���"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   135987202
         CurrentDate     =   43884
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '����
         Caption         =   "����ð�"
         BeginProperty Font 
            Name            =   "���� ���"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6510
         TabIndex        =   5
         Top             =   180
         Width           =   1005
      End
   End
   Begin VB.Menu mnuMenu01 
      Caption         =   "  ���� "
      Begin VB.Menu mnuHelp01 
         Caption         =   "��������(TeamViewer)"
         Visible         =   0   'False
      End
      Begin VB.Menu menuSep01 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu menuMDBSync 
         Caption         =   "MDB Sync"
         Visible         =   0   'False
      End
      Begin VB.Menu menuSep02 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuExit 
         Caption         =   "����"
      End
   End
   Begin VB.Menu mnuMenu02 
      Caption         =   " ���� "
      Visible         =   0   'False
      Begin VB.Menu menuUser 
         Caption         =   " ����� ���� "
      End
      Begin VB.Menu menuSep001 
         Caption         =   "-"
      End
      Begin VB.Menu menuComp 
         Caption         =   " ���� ���� "
      End
      Begin VB.Menu menuSep002 
         Caption         =   "-"
      End
      Begin VB.Menu menuPack 
         Caption         =   " ���� ���� "
      End
      Begin VB.Menu mnuSep12 
         Caption         =   "-"
      End
      Begin VB.Menu menuRegProd 
         Caption         =   " ��ǰ ������ "
      End
      Begin VB.Menu mnuSep26 
         Caption         =   "-"
      End
      Begin VB.Menu menuMastr 
         Caption         =   " �����ڵ� ���� "
      End
      Begin VB.Menu mnuSep27 
         Caption         =   "-"
      End
      Begin VB.Menu menuRegComm 
         Caption         =   " ���ڵ� ��ż���"
      End
   End
   Begin VB.Menu menuReg 
      Caption         =   " ��� "
      Visible         =   0   'False
      Begin VB.Menu menuRegLabel 
         Caption         =   " �� ��� "
      End
      Begin VB.Menu menuSep201 
         Caption         =   "-"
      End
      Begin VB.Menu menuRegBar 
         Caption         =   " ���ڵ��� "
      End
      Begin VB.Menu menuSep202 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHosp 
         Caption         =   "�� ������� ����"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu menuWork 
      Caption         =   " �۾� "
      Visible         =   0   'False
      Begin VB.Menu menuOrder 
         Caption         =   " �۾����ü� ��� "
      End
   End
   Begin VB.Menu menuPrint 
      Caption         =   " ��� "
      Visible         =   0   'False
      Begin VB.Menu menuReelPrint 
         Caption         =   " Reel ����� "
      End
      Begin VB.Menu menuSep501 
         Caption         =   "-"
      End
      Begin VB.Menu menuPPBoxPrint 
         Caption         =   " PP Box ����� "
      End
      Begin VB.Menu menuSep502 
         Caption         =   "-"
      End
      Begin VB.Menu menuICEBoxPrint 
         Caption         =   " ICE Box ����� "
      End
      Begin VB.Menu menuSep503 
         Caption         =   "-"
      End
      Begin VB.Menu menuRePrint 
         Caption         =   " �� ����� "
      End
      Begin VB.Menu menuSep504 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu menuTestPrint 
         Caption         =   " �׽�Ʈ ��� "
         Visible         =   0   'False
      End
      Begin VB.Menu menuSep506 
         Caption         =   "-"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMenu05 
      Caption         =   " �ɼ� "
      Visible         =   0   'False
      Begin VB.Menu mnuOpt 
         Caption         =   "�� �ɼ� ����"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuBarcodeUse 
         Caption         =   "�� ���ڵ� ���"
         WindowList      =   -1  'True
         Begin VB.Menu mnuBarcode 
            Caption         =   "���ڵ���"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuSeqno 
            Caption         =   "�������"
         End
         Begin VB.Menu mnuRackPos 
            Caption         =   "Rack/Pos"
         End
         Begin VB.Menu mnuCheckBox 
            Caption         =   "üũ��"
         End
      End
      Begin VB.Menu mnuSep13 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveResult 
         Caption         =   "�� ���� ���"
         Begin VB.Menu mnuEqpResult 
            Caption         =   "�����"
         End
         Begin VB.Menu mnuLisResult 
            Caption         =   "LIS���"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuSep14 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "�� ��� ����"
         Begin VB.Menu mnuSaveAuto 
            Caption         =   "�ڵ�"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuSaveManual 
            Caption         =   "����"
         End
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEMR 
         Caption         =   "�� EMR ����"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMenu03 
      Caption         =   " ��Ÿ "
      Visible         =   0   'False
      Begin VB.Menu mnuHelp02 
         Caption         =   "��������(LG Uplus)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHelp03 
         Caption         =   "��������(ez Help)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep21 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCommTest 
         Caption         =   "����׽�Ʈ"
      End
   End
End
Attribute VB_Name = "frmMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdView_Click(Index As Integer)
    Dim i   As Integer

    For i = 0 To 2
        cmdView(i).BackColor = vbWhite
    Next

    cmdView(Index).BackColor = &HC0FFFF
    
    If Index = 0 Then
        Call frmShow(frmJoui)
    ElseIf Index = 1 Then
        Call frmShow(frmBuwi)
    ElseIf Index = 2 Then
        Call frmShow(frmJouiDetail)
    End If
    
End Sub

Private Sub MDIForm_Load()
    
    tmrNow.Interval = 1000
    tmrNow.Enabled = True
    
    dtpToday = Now
    dtpTotime = Now
    
    Call frmShow(frmJoui)
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    
    End

End Sub


Private Sub mnuExit_Click()
    
    End

End Sub

Private Sub tmrNow_Timer()

    dtpTotime = Now
    
End Sub
