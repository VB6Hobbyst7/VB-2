VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "OK SOFT"
   ClientHeight    =   11880
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   21900
   BeginProperty Font 
      Name            =   "����"
      Size            =   9
      Charset         =   129
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11880
   ScaleWidth      =   21900
   WindowState     =   2  '�ִ�ȭ
   Begin VB.PictureBox picHeader 
      Align           =   1  '�� ����
      BackColor       =   &H00A5704B&
      BorderStyle     =   0  '����
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   21900
      TabIndex        =   0
      Top             =   0
      Width           =   21900
      Begin VB.TextBox txtTestNm 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "���� ���"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   10890
         TabIndex        =   56
         Text            =   "1234567890"
         Top             =   90
         Width           =   1335
      End
      Begin MSCommLib.MSComm comEqp 
         Left            =   18000
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
         RThreshold      =   1
         EOFEnable       =   -1  'True
      End
      Begin MSWinsockLib.Winsock wSck 
         Left            =   17520
         Top             =   30
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.CommandButton cmdTestNmSave 
         BackColor       =   &H00FFC0FF&
         Caption         =   "����"
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
         Left            =   12300
         Style           =   1  '�׷���
         TabIndex        =   51
         Top             =   90
         Width           =   615
      End
      Begin VB.CommandButton cmdHide 
         Appearance      =   0  '���
         BackColor       =   &H00A7FAEB&
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   18
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   16710
         Style           =   1  '�׷���
         TabIndex        =   50
         Top             =   60
         Width           =   495
      End
      Begin VB.CommandButton cmdView 
         Appearance      =   0  '���
         BackColor       =   &H00A7FAEB&
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   18
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   16200
         Style           =   1  '�׷���
         TabIndex        =   49
         Top             =   60
         Width           =   495
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00FFFFC0&
         Caption         =   "ȭ������"
         Height          =   405
         Left            =   14670
         Style           =   1  '�׷���
         TabIndex        =   19
         ToolTipText     =   "����ȭ���� ��� ����ϴ�"
         Top             =   60
         Width           =   1455
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00C0E0FF&
         Caption         =   "��������"
         Height          =   405
         Left            =   13230
         Style           =   1  '�׷���
         TabIndex        =   18
         ToolTipText     =   "������ ����� EMR������ �����մϴ�"
         Top             =   60
         Width           =   1395
      End
      Begin VB.TextBox txtTestID 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "���� ���"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   8400
         TabIndex        =   13
         Text            =   "1234567890"
         Top             =   90
         Width           =   1335
      End
      Begin VB.Label lblHospInfo 
         Appearance      =   0  '���
         BackColor       =   &H80000005&
         BackStyle       =   0  '����
         Caption         =   "�������б����� HITACHI 7020"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   150
         TabIndex        =   54
         Top             =   120
         Width           =   4000
      End
      Begin VB.Label lblHospInfo 
         Appearance      =   0  '���
         BackColor       =   &H80000005&
         BackStyle       =   0  '����
         Caption         =   "�������б����� HITACHI 7020[H36] ȫ�浿[12345]"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0FF&
         Height          =   285
         Index           =   0
         Left            =   150
         TabIndex        =   53
         Top             =   90
         Width           =   4000
      End
      Begin VB.Label Label5 
         BackStyle       =   0  '����
         Caption         =   "�˻��ڸ� : "
         BeginProperty Font 
            Name            =   "���� ���"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   9900
         TabIndex        =   52
         Top             =   120
         Width           =   975
      End
      Begin VB.Label lblTestDate 
         BackStyle       =   0  '����
         Caption         =   "1971-03-11"
         BeginProperty Font 
            Name            =   "���� ���"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5730
         TabIndex        =   5
         Top             =   120
         Width           =   1365
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '����
         Caption         =   "�˻���ID : "
         BeginProperty Font 
            Name            =   "���� ���"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7410
         TabIndex        =   3
         Top             =   120
         Width           =   975
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00ACFFEF&
         BackStyle       =   1  '�������� ����
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   0  '����
         Height          =   405
         Left            =   7260
         Top             =   60
         Width           =   5805
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '����
         Caption         =   "�˻����� :"
         BeginProperty Font 
            Name            =   "���� ���"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   4500
         TabIndex        =   1
         Top             =   120
         Width           =   945
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00ACFFEF&
         BackStyle       =   1  '�������� ����
         BorderColor     =   &H00C8FFFF&
         BorderStyle     =   0  '����
         Height          =   405
         Left            =   4320
         Top             =   60
         Width           =   2865
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00ACFFEF&
         BackStyle       =   1  '�������� ����
         BorderColor     =   &H00C8FFFF&
         BorderStyle     =   0  '����
         Height          =   405
         Left            =   90
         Top             =   60
         Width           =   4200
      End
   End
   Begin VB.PictureBox picComm 
      Align           =   2  '�Ʒ� ����
      Height          =   675
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   21840
      TabIndex        =   29
      Top             =   10620
      Visible         =   0   'False
      Width           =   21900
      Begin VB.CommandButton cmdRcvView 
         Caption         =   "V"
         Height          =   525
         Left            =   13410
         TabIndex        =   55
         Top             =   60
         Width           =   495
      End
      Begin VB.CommandButton cmdRcvClear 
         Caption         =   "C"
         Height          =   525
         Left            =   12900
         TabIndex        =   39
         Top             =   60
         Width           =   495
      End
      Begin VB.CommandButton cmdEot 
         Caption         =   "EOT"
         Height          =   405
         Left            =   20880
         TabIndex        =   38
         Top             =   120
         Width           =   585
      End
      Begin VB.CommandButton cmdEtx 
         Caption         =   "ETX"
         Height          =   405
         Left            =   20280
         TabIndex        =   37
         Top             =   120
         Width           =   585
      End
      Begin VB.CommandButton cmdStx 
         Caption         =   "STX"
         Height          =   405
         Left            =   19680
         TabIndex        =   36
         Top             =   120
         Width           =   585
      End
      Begin VB.CommandButton cmdAck 
         Caption         =   "ACK"
         Height          =   405
         Left            =   19080
         TabIndex        =   35
         Top             =   120
         Width           =   585
      End
      Begin VB.CommandButton cmdEnq 
         Caption         =   "ENQ"
         Height          =   405
         Left            =   18480
         TabIndex        =   34
         Top             =   120
         Width           =   585
      End
      Begin VB.TextBox txtSend 
         Height          =   555
         Left            =   13950
         MultiLine       =   -1  'True
         TabIndex        =   33
         Top             =   60
         Width           =   3045
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "Send"
         Height          =   525
         Left            =   17010
         TabIndex        =   32
         Top             =   60
         Width           =   1125
      End
      Begin VB.TextBox txtRcv 
         Height          =   525
         Left            =   60
         MultiLine       =   -1  'True
         TabIndex        =   31
         Top             =   60
         Width           =   11805
      End
      Begin VB.CommandButton cmdRcv 
         Caption         =   "Rcv"
         Height          =   525
         Left            =   11910
         TabIndex        =   30
         Top             =   60
         Width           =   975
      End
   End
   Begin VB.Frame fraPatInfo 
      BackColor       =   &H00FFFFFF&
      Height          =   1155
      Left            =   16440
      TabIndex        =   16
      Top             =   570
      Width           =   6525
      Begin VB.TextBox txtSA 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4350
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   "1004"
         Top             =   660
         Width           =   1935
      End
      Begin VB.TextBox txtPName 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "1004"
         Top             =   660
         Width           =   1935
      End
      Begin VB.TextBox txtPatID 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4350
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "1004"
         Top             =   270
         Width           =   1935
      End
      Begin VB.TextBox txtBarcode 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "1004"
         Top             =   270
         Width           =   1935
      End
      Begin VB.Label Label7 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '����
         Caption         =   "Sex/Age"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   3390
         TabIndex        =   28
         Top             =   720
         Width           =   885
      End
      Begin VB.Label Label6 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '����
         Caption         =   "��      ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   300
         TabIndex        =   26
         Top             =   720
         Width           =   885
      End
      Begin VB.Label Label4 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '����
         Caption         =   "���Ϲ�ȣ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   3390
         TabIndex        =   24
         Top             =   330
         Width           =   885
      End
      Begin VB.Label Label3 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '����
         Caption         =   "��ü��ȣ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   300
         TabIndex        =   22
         Top             =   330
         Width           =   885
      End
   End
   Begin VB.Frame fraHidden 
      Caption         =   "Hidden"
      Height          =   2355
      Left            =   20610
      TabIndex        =   14
      Top             =   8880
      Visible         =   0   'False
      Width           =   6255
      Begin VB.CommandButton cmdWork 
         Caption         =   "��ũ��ȸ"
         Height          =   315
         Left            =   960
         TabIndex        =   17
         Top             =   270
         Width           =   1425
      End
      Begin VB.CommandButton cmdResult 
         Caption         =   "�����ȸ"
         Height          =   315
         Left            =   2610
         TabIndex        =   15
         Top             =   270
         Width           =   1425
      End
      Begin VB.Label lblPatInfo 
         BackStyle       =   0  '����
         Caption         =   "�ڰ˻�"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   765
         Left            =   1170
         TabIndex        =   20
         Top             =   990
         Width           =   3465
      End
      Begin VB.Shape shpPatInfo 
         BorderColor     =   &H00FF0000&
         Height          =   1155
         Left            =   900
         Shape           =   4  '�ձ� �簢��
         Top             =   810
         Visible         =   0   'False
         Width           =   4035
      End
      Begin VB.Shape Shape10 
         BorderColor     =   &H0080FFFF&
         Height          =   405
         Left            =   900
         Top             =   210
         Width           =   1545
      End
      Begin VB.Shape Shape11 
         BorderColor     =   &H0080FFFF&
         Height          =   405
         Left            =   2550
         Top             =   210
         Width           =   1545
      End
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  '�Ʒ� ����
      BackColor       =   &H00404040&
      BorderStyle     =   0  '����
      Height          =   585
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   21900
      TabIndex        =   6
      Top             =   11295
      Width           =   21900
      Begin VB.Timer tmrConn 
         Left            =   18060
         Top             =   90
      End
      Begin VB.Timer tmrSend 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   18960
         Top             =   90
      End
      Begin VB.Timer tmrReceive 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   18540
         Top             =   90
      End
      Begin VB.Timer tmrDBConn 
         Left            =   17610
         Top             =   90
      End
      Begin MSComctlLib.ImageList imlStatus 
         Left            =   19440
         Top             =   -30
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   7
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":0E42
               Key             =   "RUN"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":13DC
               Key             =   "NOT"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1976
               Key             =   "STOP"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1F10
               Key             =   "LST"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":27A2
               Key             =   "ITM"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":28FC
               Key             =   "ERR"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2A56
               Key             =   "NOF"
            EndProperty
         EndProperty
      End
      Begin FPSpread.vaSpread spdComStatus 
         Height          =   330
         Left            =   9840
         TabIndex        =   12
         Top             =   120
         Width           =   5220
         _Version        =   393216
         _ExtentX        =   9208
         _ExtentY        =   582
         _StockProps     =   64
         BackColorStyle  =   1
         ButtonDrawMode  =   4
         DisplayColHeaders=   0   'False
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
         GrayAreaBackColor=   16777215
         GridShowVert    =   0   'False
         MaxCols         =   3
         MaxRows         =   3
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         ScrollBarExtMode=   -1  'True
         SelectBlockOptions=   2
         ShadowColor     =   16777215
         SpreadDesigner  =   "frmMain.frx":2BB0
         UserResize      =   0
         TextTip         =   2
         ScrollBarTrack  =   3
         ShowScrollTips  =   3
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H00FFFFFF&
         Height          =   405
         Left            =   9810
         Top             =   90
         Width           =   5295
      End
      Begin VB.Image imgPort 
         Height          =   240
         Left            =   3900
         Picture         =   "frmMain.frx":3001
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgSend 
         Height          =   240
         Left            =   4845
         Picture         =   "frmMain.frx":358B
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgReceive 
         Height          =   240
         Left            =   5760
         Picture         =   "frmMain.frx":3B15
         Top             =   180
         Width           =   240
      End
      Begin VB.Label lblPort 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "��Ʈ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   180
         Left            =   3360
         TabIndex        =   11
         Top             =   210
         Width           =   360
      End
      Begin VB.Label lblSend 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "�۽�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   195
         Left            =   4335
         TabIndex        =   10
         Top             =   210
         Width           =   420
      End
      Begin VB.Label lblRcv 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   195
         Left            =   5220
         TabIndex        =   9
         Top             =   210
         Width           =   420
      End
      Begin VB.Image imgNet1 
         Height          =   240
         Left            =   390
         Picture         =   "frmMain.frx":409F
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgNet2 
         Height          =   240
         Left            =   390
         Picture         =   "frmMain.frx":41E9
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgNet3 
         Height          =   240
         Left            =   390
         Picture         =   "frmMain.frx":4333
         Top             =   180
         Width           =   240
      End
      Begin VB.Label lblComStatus 
         Alignment       =   1  '������ ����
         BackStyle       =   0  '����
         Caption         =   "Com1 ���Ἲ��"
         BeginProperty Font 
            Name            =   "���� ���"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   285
         Left            =   6210
         TabIndex        =   8
         Top             =   180
         Width           =   3345
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H0080FFFF&
         Height          =   405
         Left            =   210
         Top             =   90
         Width           =   2955
      End
      Begin VB.Label lblDBStatus 
         Alignment       =   2  '��� ����
         BackStyle       =   0  '����
         Caption         =   "�����ͺ��̽� ���Ἲ��"
         BeginProperty Font 
            Name            =   "���� ���"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   285
         Left            =   750
         TabIndex        =   7
         Top             =   180
         Width           =   2295
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00FFFFFF&
         Height          =   405
         Left            =   3180
         Top             =   90
         Width           =   6585
      End
   End
   Begin FPSpread.vaSpread spdResult 
      Height          =   6555
      Left            =   16440
      TabIndex        =   4
      Top             =   1740
      Width           =   6495
      _Version        =   393216
      _ExtentX        =   11456
      _ExtentY        =   11562
      _StockProps     =   64
      BackColorStyle  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "���� ���"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      GridColor       =   15921919
      MaxCols         =   13
      MaxRows         =   20
      ScrollBarExtMode=   -1  'True
      ShadowColor     =   16777215
      SpreadDesigner  =   "frmMain.frx":447D
      ScrollBarTrack  =   3
      ShowScrollTips  =   3
   End
   Begin FPSpread.vaSpread spdOrder 
      Height          =   7935
      Left            =   6000
      TabIndex        =   2
      Top             =   660
      Width           =   15345
      _Version        =   393216
      _ExtentX        =   27067
      _ExtentY        =   13996
      _StockProps     =   64
      ColsFrozen      =   22
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "���� ���"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      GridColor       =   15921919
      GridShowVert    =   0   'False
      MaxCols         =   22
      MaxRows         =   20
      OperationMode   =   2
      ScrollBarExtMode=   -1  'True
      SelectBlockOptions=   0
      ShadowColor     =   16777215
      SpreadDesigner  =   "frmMain.frx":51FF
      ScrollBarTrack  =   3
      ShowScrollTips  =   3
   End
   Begin FPSpread.vaSpread spdWork 
      Height          =   7275
      Left            =   60
      TabIndex        =   48
      Top             =   1320
      Width           =   5895
      _Version        =   393216
      _ExtentX        =   10398
      _ExtentY        =   12832
      _StockProps     =   64
      ColsFrozen      =   22
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "���� ���"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      GridColor       =   15921919
      GridShowVert    =   0   'False
      MaxCols         =   23
      MaxRows         =   20
      OperationMode   =   2
      ScrollBarExtMode=   -1  'True
      SelectBlockOptions=   0
      ShadowColor     =   16777215
      SpreadDesigner  =   "frmMain.frx":725D
      ScrollBarTrack  =   3
      ShowScrollTips  =   3
   End
   Begin VB.Frame fraWorkInfo 
      BackColor       =   &H00FFFFFF&
      Height          =   705
      Left            =   60
      TabIndex        =   40
      Top             =   570
      Width           =   5895
      Begin VB.CommandButton cmdAll 
         BackColor       =   &H00A7FAEB&
         Caption         =   "����"
         Height          =   375
         Left            =   4800
         Style           =   1  '�׷���
         TabIndex        =   43
         Top             =   180
         Width           =   975
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00A7FAEB&
         Caption         =   "��ȸ"
         Height          =   375
         Left            =   3720
         Style           =   1  '�׷���
         TabIndex        =   42
         Top             =   180
         Width           =   1005
      End
      Begin VB.CommandButton cmdMatch 
         BackColor       =   &H00FFFFFF&
         Caption         =   ">>"
         Height          =   375
         Left            =   5340
         Style           =   1  '�׷���
         TabIndex        =   41
         Top             =   420
         Visible         =   0   'False
         Width           =   495
      End
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   375
         Left            =   690
         TabIndex        =   44
         Top             =   210
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   661
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
         Format          =   140574721
         CurrentDate     =   40457
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   375
         Left            =   2280
         TabIndex        =   45
         Top             =   210
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   661
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
         Format          =   140574721
         CurrentDate     =   40457
      End
      Begin VB.Label Label1 
         Appearance      =   0  '���
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '����
         Caption         =   "~"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   2
         Left            =   2100
         TabIndex        =   47
         Top             =   300
         Width           =   150
      End
      Begin VB.Label Label1 
         Appearance      =   0  '���
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '����
         Caption         =   "��ȸ�Ⱓ"
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
         Height          =   360
         Index           =   1
         Left            =   150
         TabIndex        =   46
         Top             =   210
         Width           =   480
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Menu mnuMenu01 
      Caption         =   "  ����  "
      Begin VB.Menu mnuResult 
         Caption         =   "�� ��� ��ȸ"
      End
      Begin VB.Menu mnuSep29 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWork 
         Caption         =   "�� ��ũ ��ȸ"
      End
      Begin VB.Menu mnuSep9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "����"
      End
   End
   Begin VB.Menu mnuMenu02 
      Caption         =   "  ����  "
      Begin VB.Menu mnuComm 
         Caption         =   "�� ��� ����"
      End
      Begin VB.Menu mnuSep11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTest 
         Caption         =   "�� �˻� ����"
      End
      Begin VB.Menu mnuSep12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuView 
         Caption         =   "�� ȭ�� ����"
      End
      Begin VB.Menu mnuSep22 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "�� �ɼ� ����"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep23 
         Caption         =   "-"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMenu05 
      Caption         =   " �ɼ� "
      Begin VB.Menu mnuBarcodeUse 
         Caption         =   "�� ���ڵ� ���"
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
      Begin VB.Menu mnuSep14 
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
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHosp 
         Caption         =   "�� ���� ����"
      End
      Begin VB.Menu mnuSep30 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEMR 
         Caption         =   "�� EMR ����"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMenu03 
      Caption         =   " ��Ÿ "
      Begin VB.Menu mnuHelp01 
         Caption         =   "��������(TeamViewer)"
      End
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
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public sStartTime       As Date
Public sStartDate       As Date

Dim pDel                As Boolean
Dim strOldBarno         As String
Dim gMnuIdx             As Integer

Private Sub cmdEnd_Click()

    If MsgBox("���� ������Դϴ�. �����Ͻðڽ��ϱ�?", vbYesNo + vbCritical, "���α׷� ����") = vbYes Then

        If comEqp.PortOpen = True Then
            comEqp.PortOpen = False
        End If

        If gDBTYPE <> "99" Then
            Call DisConnect_Server

            Call DisConnect_Local
        End If

        Unload Me

        End
    End If
    
End Sub

Private Sub cmdAck_Click()
    
    txtSend.Text = txtSend.Text & ACK

End Sub

Private Sub cmdAll_Click()
    Dim i               As Integer
    Dim intRow          As Integer
    Dim intWRow         As Integer
    Dim intORow         As Integer
    Dim intWCol         As Integer
    Dim intOCol         As Integer
    Dim strBarno        As String
    Dim blnSame         As Boolean
    Dim varItems        As Variant
    Dim intItems        As Integer
    
    With spdWork
        For intWRow = 1 To .MaxRows
            .Row = intWRow
            .Col = colCHECKBOX
            If .Value = "1" Then
                blnSame = False
                strBarno = GetText(spdWork, intWRow, colBARCODE)
                For intORow = 1 To spdOrder.MaxRows
                    spdOrder.Row = intORow
                    spdOrder.Col = colBARCODE
                    If strBarno = GetText(spdOrder, intORow, colBARCODE) Then
                        blnSame = True
                    End If
                Next
                
                If blnSame = False Then
                    spdOrder.MaxRows = spdOrder.MaxRows + 1
                    intRow = spdOrder.MaxRows
                    For i = colCHECKBOX To colSTATE
                        Call SetText(spdOrder, GetText(spdWork, intWRow, i), intRow, i)
                
                        varItems = GetText(spdWork, intWRow, colITEMS)
                        varItems = Split(varItems, "/")
                        For intItems = 0 To UBound(varItems)
                            For intOCol = colSTATE + 1 To spdOrder.MaxCols
                                spdOrder.Row = 0
                                spdOrder.Col = intOCol
                                If varItems(intItems) = Trim(spdOrder.Text) Then
                                    .Row = spdOrder.MaxRows
                                    Call SetText(spdOrder, "��", spdOrder.MaxRows, intOCol)
                                End If
                            Next
                        Next
                    Next
                    
                    spdOrder.RowHeight(-1) = 15
                End If
            End If
        Next
        '.MaxRows = 0
    End With
End Sub

Private Sub cmdClear_Click()

    Call frmClear
    
End Sub

Private Sub cmdEnq_Click()
    
    txtSend.Text = txtSend.Text & ENQ
    
End Sub

Private Sub cmdEot_Click()
    
    txtSend.Text = txtSend.Text & EOT

End Sub

Private Sub cmdEtx_Click()
    
    txtSend.Text = txtSend.Text & ETX

End Sub


Private Sub cmdHide_Click()
    
    spdResult.Visible = False
    fraPatInfo.Visible = False
    
    spdOrder.WIDTH = Me.ScaleWidth - spdResult.WIDTH - 200

End Sub

Private Sub cmdMatch_Click()
    Dim intWRow     As Integer
    Dim intORow     As Integer
    Dim blnSame     As Boolean
    Dim i           As Integer
    Dim intCnt      As Integer
    
    blnSame = False
    intCnt = 0
    
    For intWRow = 1 To spdWork.MaxRows
        If GetText(spdWork, intORow, colCHECKBOX) = "1" Then
            intCnt = intCnt + 1
        End If
    Next
    
    If intCnt > 1 Then
        MsgBox "�� ����� �����ϼ���", vbOKOnly, Me.Caption
        Exit Sub
    End If
       
    For intORow = 1 To spdOrder.MaxRows
        If GetText(spdOrder, intORow, colCHECKBOX) = "1" Then
            blnSame = True
            Exit For
        End If
    Next
    
    If blnSame = False Then
        MsgBox "��� ��ü�� �����ϼ���", vbOKOnly, Me.Caption
        Exit Sub
    End If
    
    
    If blnSame = True Then
        For i = colCHECKBOX To colSTATE
            Call SetText(spdOrder, GetText(spdWork, intWRow, i), intORow, i)
            
'            varItems = GetText(spdWork, intWRow, colITEMS)
'            varItems = Split(varItems, "/")
'            For intItems = 0 To UBound(varItems)
'                For intOCol = colSTATE + 1 To frmMain.spdOrder.MaxCols
'                    .Row = 0
'                    .Col = intOCol
'                    If varItems(intItems) = Trim(.Text) Then
'                        .Row = intRow
'                        Call SetText(spdOrder, "��", intRow, intOCol)
'                    End If
'                Next
'            Next
        Next
        
        '��������
        SQL = ""
        SQL = SQL & "UPDATE PATRESULT SET "
        SQL = SQL & "  BARCODE  = '" & Trim(GetText(spdOrder, intORow, colBARCODE)) & "'" & vbCrLf
        SQL = SQL & " ,PID      = '" & Trim(GetText(spdOrder, intORow, colPID)) & "'" & vbCrLf
        SQL = SQL & " ,CHARTNO  = '" & Trim(GetText(spdOrder, intORow, colCHARTNO)) & "'" & vbCrLf
        SQL = SQL & " ,SPECIMEN = '" & Trim(GetText(spdOrder, intORow, colSPECIMEN)) & "'" & vbCrLf
        SQL = SQL & " ,DEPT     = '" & Trim(GetText(spdOrder, intORow, colDEPT)) & "'" & vbCrLf
        SQL = SQL & " ,INOUT    = '" & Trim(GetText(spdOrder, intORow, colINOUT)) & "'" & vbCrLf
        SQL = SQL & " ,ERYN     = '" & Trim(GetText(spdOrder, intORow, colER)) & "'" & vbCrLf
        SQL = SQL & " ,RETESTYN = '" & Trim(GetText(spdOrder, intORow, colRT)) & "'" & vbCrLf
        SQL = SQL & " ,PNAME    = '" & Trim(GetText(spdOrder, intORow, colPNAME)) & "'" & vbCrLf
        SQL = SQL & " ,PSEX     = '" & Trim(GetText(spdOrder, intORow, colPSEX)) & "'" & vbCrLf
        SQL = SQL & " ,PAGE     = '" & Trim(GetText(spdOrder, intORow, colPAGE)) & "'" & vbCrLf
        SQL = SQL & " ,DISKNO   = '" & Trim(GetText(spdOrder, intORow, colRACKNO)) & "'" & vbCrLf
        SQL = SQL & " ,POSNO    = '" & Trim(GetText(spdOrder, intORow, colPOSNO)) & "'" & vbCrLf
        SQL = SQL & " WHERE EQUIPNO  = '" & gHOSP.MACHCD & "'" & vbCrLf
        SQL = SQL & "   AND EXAMDATE = '" & Trim(GetText(spdOrder, intORow, colEXAMDATE)) & "'" & vbCrLf
        SQL = SQL & "   AND EXAMTIME = '" & Trim(GetText(spdOrder, intORow, colEXAMTIME)) & "'" & vbCrLf
        SQL = SQL & "   AND SAVESEQ  = " & Trim(GetText(spdOrder, intORow, colSAVESEQ)) & vbCrLf
        
        If DBExec(AdoCn_Local, SQL) Then
            '-- ����
        End If
    End If
End Sub

Private Sub cmdRcv_Click()
        
    pBuffer = txtRcv.Text
    
    Select Case UCase(gHOSP.MACHNM)
        Case "HITACHI7180":     Call Phase_Serial_HITACHI7180
        Case "XN1000":          Call Phase_Serial_XN1000
        Case "ISMART30":        Call Phase_Serial_ISMART30
        Case "KLITE":           Call Phase_TCP_KLITE
        Case "C8000":           Call Phase_Serial_C8000
    End Select

    pBuffer = ""
    
End Sub

Private Sub cmdRcvClear_Click()
    
    txtRcv.Text = ""
    
End Sub

Private Sub cmdRcvView_Click()

    frmLogView.Show
    
End Sub

Private Sub cmdResult_Click()

    frmResult.Show vbModal

End Sub

Private Sub cmdSave_Click()
    Dim lRow As Long
    Dim Res  As Integer
    
    If spdOrder.MaxRows = 0 Then
        Exit Sub
    End If
    
    If MsgBox("������ ����� �����Ͻðڽ��ϱ�?", vbYesNo + vbCritical, "�������") = vbYes Then
        With spdOrder
            For lRow = 1 To .DataRowCnt
                .Row = lRow
                .Col = 1
                If .Value = 1 Then
                    Res = SaveTransData(lRow, spdOrder)
                    
                    If Res = -1 Then
                        SetForeColor spdOrder, lRow, lRow, 1, colSTATE, 255, 0, 0
                        SetText spdOrder, "�������", lRow, colSTATE
                    
                              SQL = " UPDATE PATRESULT SET " & vbCrLf
                        SQL = SQL & "     SENDFLAG  = '1' " & vbCrLf
                        SQL = SQL & "   , SENDDATE  = '" & Format(Now, "yyyy-mm-dd") & "' " & vbCrLf
                        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "' " & vbCrLf
                        SQL = SQL & "   AND BARCODE = '" & Trim(GetText(spdOrder, lRow, colBARCODE)) & "' "
                        
                        If DBExec(AdoCn_Local, SQL) Then
                            '-- ����
                        End If
                    
                    Else
                        SetBackColor spdOrder, lRow, lRow, 1, colSTATE, 202, 255, 112
                        SetText spdOrder, "����Ϸ�", lRow, colSTATE
                        
                              SQL = " UPDATE PATRESULT SET " & vbCrLf
                        SQL = SQL & "     SENDFLAG  = '2' " & vbCrLf
                        SQL = SQL & "   , SENDDATE  = '" & Format(Now, "yyyy-mm-dd") & "' " & vbCrLf
                        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "' " & vbCrLf
                        SQL = SQL & "   AND BARCODE = '" & Trim(GetText(spdOrder, lRow, colBARCODE)) & "' "
                        
                        If DBExec(AdoCn_Local, SQL) Then
                            '-- ����
                        End If
                        
                    End If
                    spdOrder.Row = lRow
                    spdOrder.Col = 1
                    spdOrder.Value = 0
                End If
            Next lRow
        End With
    End If
    
End Sub

Private Sub cmdSearch_Click()
    
    Call GetWorkList(Format(dtpFrom.Value, "yyyymmdd"), Format(dtpTo.Value, "yyyymmdd"), spdWork)

End Sub

Private Sub cmdSend_Click()
    
    
    Call SendData(txtSend.Text)

End Sub

Private Sub cmdStx_Click()
    
    txtSend.Text = txtSend.Text & STX

End Sub

Private Sub cmdTestNmSave_Click()
    
    Call WritePrivateProfileString("HOSP", "USERID", txtTestID.Text, App.PATH & "\INI\" & gMACH & ".ini")
    Call WritePrivateProfileString("HOSP", "USERNM", txtTestNm.Text, App.PATH & "\INI\" & gMACH & ".ini")

End Sub

Private Sub cmdView_Click()
    
    spdResult.Visible = True
    fraPatInfo.Visible = True
        
    spdOrder.WIDTH = Me.ScaleWidth - spdWork.WIDTH - spdResult.WIDTH - 200
    spdOrder.HEIGHT = Me.ScaleHeight - picHeader.HEIGHT - picBottom.HEIGHT - 280
    spdResult.LEFT = spdOrder.LEFT + spdOrder.WIDTH + 50
    spdResult.HEIGHT = Me.ScaleHeight - picHeader.HEIGHT - picBottom.HEIGHT - fraPatInfo.HEIGHT - 300

End Sub

Private Sub cmdWork_Click()

    frmWorkList.Show vbModal
    
End Sub

Private Sub comEQP_OnComm()
    Dim EVMsg       As String
    Dim ERMsg       As String
    Dim Ret         As Long
    Dim strDate     As String
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    Select Case comEqp.CommEvent
        Case comEvReceive
            imgReceive.Picture = imlStatus.ListImages("RUN").ExtractIcon
            If tmrReceive.Enabled = False Then
                tmrReceive.Enabled = True
            Else
                tmrReceive.Enabled = False
                tmrReceive.Enabled = True
            End If

            pBuffer = comEqp.Input

            'SetRawData "[Rx]" & pBuffer
            SetRawData "" & pBuffer
            

            Select Case UCase(gHOSP.MACHNM)
                        Case "HITACHI7180":     Call Phase_Serial_HITACHI7180
                        Case "XN1000":          Call Phase_Serial_XN1000
                        Case "ISMART30":        Call Phase_Serial_ISMART30
                        Case "C8000":           Call Phase_Serial_C8000
'                        Case "KLITE":           Call Phase_Serial_KLITE
                        'Case "INDIKO":          Call Phase_Serial_INDIKO
                            
            End Select

        Case comEvSend
            imgSend.Picture = imlStatus.ListImages("RUN").ExtractIcon
            If tmrSend.Enabled = False Then
                tmrSend.Enabled = True
            Else
                tmrSend.Enabled = False
                tmrSend.Enabled = True
            End If

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

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Cancel = 1
    Call cmdEnd_Click
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If MsgBox("���� ������Դϴ�. �����Ͻðڽ��ϱ�?", vbYesNo + vbCritical, "���α׷� ����") = vbYes Then
        
        Close #1

        If comEqp.PortOpen = True Then
            comEqp.PortOpen = False
        End If
    
        Call DisConnect_Server
        
        Call DisConnect_Local
        
        Unload Me
        
        End
    End If
    
End Sub

Private Sub GetOrder_AU680(ByVal pBarNo As String, ByVal pType As String)

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String
    Dim GetOrder    As String

    intRow = -1
    GetOrder = ""

    Call SetCommStatus("Q", pBarNo, frmMain.spdComStatus)
    
    '-- 1. �������� ��ȸ
    With frmMain
        Select Case pType
            '-- ���ڵ� ���
            Case "0"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmMain.spdOrder, i, colBARCODE)) = pBarNo Then
                        intRow = i
                        Exit For
                    End If
                Next i
        
            '-- Seq
            Case "1"
                For i = 1 To .spdOrder.DataRowCnt
                    If Val(Trim(GetText(frmMain.spdOrder, i, colSEQNO))) = Val(mOrder.Seq) Then
                        pBarNo = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarNo
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Rack/Pos
            Case "2"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmMain.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmMain.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
                        pBarNo = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Check Top
            Case "3"
                For i = 1 To .spdOrder.DataRowCnt
                    If GetText(frmMain.spdOrder, i, colCHECKBOX) = "1" Then
                        pBarNo = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarNo
                        intRow = i
                        Exit For
                    End If
                Next i
        End Select
    

        '-- �������忡�� ��ã����..
        If intRow < 0 Then
            intRow = .spdOrder.DataRowCnt + 1
            If .spdOrder.MaxRows < intRow Then
                .spdOrder.MaxRows = intRow
            End If
        End If

        '-- ���������� ȭ��ǥ��
        Call SetText(.spdOrder, mOrder.BarNo, intRow, colBARCODE)
        Call SetText(.spdOrder, mOrder.Seq, intRow, colSEQNO)
        Call SetText(.spdOrder, mOrder.RackNo, intRow, colRACKNO)
        Call SetText(.spdOrder, mOrder.TubePos, intRow, colPOSNO)

        '-- ����������� �����
        .spdResult.MaxRows = 0

        '-- �˻��� ���� ��������
        Call GetSampleInfo(intRow, .spdOrder)

        .spdOrder.RowHeight(-1) = 15

        '-- �������̺��� �˻��׸� �ش��ϴ� �˻�ä�� ã�ƿ��� (intRow = ���� �˻��ߴ� ���ڵ尡 �ٽ� �ö�� ��� ��ġ�� ��ã�´�.)
        strItems = GetEquipExamCode_AU680(gHOSP.MACHCD, pBarNo, intRow)

        '-- �˻�ä�η� ������ �����
        If Trim(strItems) = "" Then
            mOrder.NoOrder = True
            mOrder.Order = ""

            GetOrder = STX & "S " & mOrder.RackNo & mOrder.TubePos & mOrder.Seq & Space(gHOSP.BARLEN - Len(mOrder.BarNo)) & mOrder.BarNo & Space(4) & "E" & ETX

            '-- �������(Order) ǥ��
            Call SetText(frmMain.spdOrder, "��������", intRow, colSTATE)
        Else
            mOrder.NoOrder = False
            mOrder.Order = strItems

            GetOrder = STX & "S " & mOrder.RackNo & mOrder.TubePos & mOrder.Seq & Space(gHOSP.BARLEN - Len(mOrder.BarNo)) & mOrder.BarNo & Space(4) & "E" & strItems & ETX

            '-- �������(Order) ǥ��
            Call SetText(frmMain.spdOrder, "��������", intRow, colSTATE)
        End If

        '-- ���� ����
        Call SendData(GetOrder)
        
        Call SetCommStatus("S", pBarNo, spdComStatus)

        '-- ���� Row
        gRow = intRow

    End With

End Sub

Private Sub GetOrder_HITACHI7180(ByVal pBarNo As String, ByVal pType As String)

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String
    Dim GetOrder    As String

    intRow = -1
    GetOrder = ""

    Call SetCommStatus("Q", pBarNo, frmMain.spdComStatus)
    
    '-- 1. �������� ��ȸ
    With frmMain
        Select Case pType
            '-- ���ڵ� ���
            Case "0"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmMain.spdOrder, i, colBARCODE)) = pBarNo Then
                        intRow = i
                        Exit For
                    End If
                Next i
        
            '-- Seq
            Case "1"
                For i = 1 To .spdOrder.DataRowCnt
                    If Val(Trim(GetText(frmMain.spdOrder, i, colSEQNO))) = Val(mOrder.Seq) Then
                        pBarNo = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarNo
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Rack/Pos
            Case "2"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmMain.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmMain.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
                        pBarNo = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Check Top
            Case "3"
                For i = 1 To .spdOrder.DataRowCnt
                    If GetText(frmMain.spdOrder, i, colCHECKBOX) = "1" Then
                        pBarNo = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarNo
                        intRow = i
                        Exit For
                    End If
                Next i
        End Select
    

        '-- �������忡�� ��ã����..
        If intRow < 0 Then
            intRow = .spdOrder.DataRowCnt + 1
            If .spdOrder.MaxRows < intRow Then
                .spdOrder.MaxRows = intRow
            End If
        End If

        '-- ���������� ȭ��ǥ��
        Call SetText(.spdOrder, mOrder.BarNo, intRow, colBARCODE)
        Call SetText(.spdOrder, mOrder.Seq, intRow, colSEQNO)
        Call SetText(.spdOrder, mOrder.RackNo, intRow, colRACKNO)
        Call SetText(.spdOrder, mOrder.TubePos, intRow, colPOSNO)

        '-- ����������� �����
        .spdResult.MaxRows = 0

        '-- �˻��� ���� ��������
        Call GetSampleInfo(intRow, .spdOrder)

        .spdOrder.RowHeight(-1) = 15

        '-- �������̺��� �˻��׸� �ش��ϴ� �˻�ä�� ã�ƿ��� (intRow = ���� �˻��ߴ� ���ڵ尡 �ٽ� �ö�� ��� ��ġ�� ��ã�´�.)
        strItems = ""
        
        'MsgBox "1"
        
        strItems = GetEquipExamCode_HITACHI7180(gHOSP.MACHCD, pBarNo, intRow)
        
        mOrder.Func = Replace(mOrder.Func, String(13, "#"), LEFT(mOrder.BarNo & Space(13), 13))
        
        '-- �˻�ä�η� ������ �����
        'If Trim(strItems) = "" Then
        If mOrder.SendCnt = 0 Then
            mOrder.NoOrder = True
            mOrder.Order = ""
        
            GetOrder = STX & ";" & mOrder.Func & " 88" & mOrder.Order & "100000" & LEFT(mOrder.PID & Space(30), 30) & ETX  '& vbCrLf
            'GetOrder = STX & ";" & mOrder.Func & " 88" & mOrder.Order & "100000" & Space(30) & ETX  '& vbCrLf
            
            '-- �������(Order) ǥ��
            Call SetText(frmMain.spdOrder, "��������", intRow, colSTATE)
        Else
            mOrder.NoOrder = False
            mOrder.Order = strItems
        
            GetOrder = STX & ";" & mOrder.Func & " 88" & mOrder.Order & "100000" & LEFT(mOrder.PID & Space(30), 30) & ETX  '& vbCrLf
            'GetOrder = STX & ";" & mOrder.Func & " 88" & mOrder.Order & "100000" & Space(30) & ETX  '& vbCrLf
        
            '-- �������(Order) ǥ��
            Call SetText(frmMain.spdOrder, "��������", intRow, colSTATE)
        End If

        Call SendData(GetOrder)
        
        '-- ���� Row
        gRow = intRow

    End With

End Sub

Private Sub GetOrder_INDIKO(ByVal pBarNo As String, ByVal pType As String)

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String
    Dim GetOrder    As String

    intRow = -1
    GetOrder = ""

    Call SetCommStatus("Q", pBarNo, frmMain.spdComStatus)
    
    '-- 1. �������� ��ȸ
    With frmMain
        Select Case pType
            '-- ���ڵ� ���
            Case "0"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmMain.spdOrder, i, colBARCODE)) = pBarNo Then
                        intRow = i
                        Exit For
                    End If
                Next i
        
            '-- Seq
            Case "1"
                For i = 1 To .spdOrder.DataRowCnt
                    If Val(Trim(GetText(frmMain.spdOrder, i, colSEQNO))) = Val(mOrder.Seq) Then
                        pBarNo = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarNo
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Rack/Pos
            Case "2"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmMain.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmMain.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
                        pBarNo = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Check Top
            Case "3"
                For i = 1 To .spdOrder.DataRowCnt
                    If GetText(frmMain.spdOrder, i, colCHECKBOX) = "1" Then
                        pBarNo = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarNo
                        intRow = i
                        Exit For
                    End If
                Next i
        End Select
    

        '-- �������忡�� ��ã����..
        If intRow < 0 Then
            intRow = .spdOrder.DataRowCnt + 1
            If .spdOrder.MaxRows < intRow Then
                .spdOrder.MaxRows = intRow
            End If
        End If

        '-- ���������� ȭ��ǥ��
        Call SetText(.spdOrder, mOrder.BarNo, intRow, colBARCODE)
        Call SetText(.spdOrder, mOrder.Seq, intRow, colSEQNO)
        Call SetText(.spdOrder, mOrder.RackNo, intRow, colRACKNO)
        Call SetText(.spdOrder, mOrder.TubePos, intRow, colPOSNO)

        '-- ����������� �����
        .spdResult.MaxRows = 0

        '-- �˻��� ���� ��������
        Call GetSampleInfo(intRow, .spdOrder)

        .spdOrder.RowHeight(-1) = 15

        '-- �������̺��� �˻��׸� �ش��ϴ� �˻�ä�� ã�ƿ��� (intRow = ���� �˻��ߴ� ���ڵ尡 �ٽ� �ö�� ��� ��ġ�� ��ã�´�.)
        strItems = GetEquipExamCode_INDIKO(gHOSP.MACHCD, pBarNo, intRow)
        
        '-- �˻�ä�η� ������ �����
        If Trim(strItems) = "" Then
            mOrder.NoOrder = True
            mOrder.Order = ""
            
            '-- �������(Order) ǥ��
            Call SetText(frmMain.spdOrder, "��������", intRow, colSTATE)
        Else
            mOrder.NoOrder = False
            mOrder.Order = strItems
        
            '-- �������(Order) ǥ��
            Call SetText(frmMain.spdOrder, "��������", intRow, colSTATE)
        End If


        '-- ���� Row
        gRow = intRow

    End With

End Sub

Private Sub GetOrder_XN1000(ByVal pBarNo As String, ByVal pType As String)

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String
    Dim GetOrder    As String

    intRow = -1
    GetOrder = ""

    Call SetCommStatus("Q", pBarNo, frmMain.spdComStatus)
    
    '-- 1. �������� ��ȸ
    With frmMain
        Select Case pType
            '-- ���ڵ� ���
            Case "0"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmMain.spdOrder, i, colBARCODE)) = pBarNo Then
                        intRow = i
                        Exit For
                    End If
                Next i
        
            '-- Seq
            Case "1"
                For i = 1 To .spdOrder.DataRowCnt
                    If Val(Trim(GetText(frmMain.spdOrder, i, colSEQNO))) = Val(mOrder.Seq) Then
                        pBarNo = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarNo
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Rack/Pos
            Case "2"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmMain.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmMain.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
                        pBarNo = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Check Top
            Case "3"
                For i = 1 To .spdOrder.DataRowCnt
                    If GetText(frmMain.spdOrder, i, colCHECKBOX) = "1" Then
                        pBarNo = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarNo
                        intRow = i
                        Exit For
                    End If
                Next i
        End Select
    

        '-- �������忡�� ��ã����..
        If intRow < 0 Then
            intRow = .spdOrder.DataRowCnt + 1
            If .spdOrder.MaxRows < intRow Then
                .spdOrder.MaxRows = intRow
            End If
        End If

        '-- ���������� ȭ��ǥ��
        Call SetText(.spdOrder, mOrder.BarNo, intRow, colBARCODE)
        Call SetText(.spdOrder, mOrder.Seq, intRow, colSEQNO)
        Call SetText(.spdOrder, mOrder.RackNo, intRow, colRACKNO)
        Call SetText(.spdOrder, mOrder.TubePos, intRow, colPOSNO)

        '-- ����������� �����
        .spdResult.MaxRows = 0

        '-- �˻��� ���� ��������
        Call GetSampleInfo(intRow, .spdOrder)

        .spdOrder.RowHeight(-1) = 15

        '-- �������̺��� �˻��׸� �ش��ϴ� �˻�ä�� ã�ƿ��� (intRow = ���� �˻��ߴ� ���ڵ尡 �ٽ� �ö�� ��� ��ġ�� ��ã�´�.)
        strItems = GetEquipExamCode_XN1000(gHOSP.MACHCD, pBarNo, intRow)
        
        '-- �˻�ä�η� ������ �����
        If Trim(strItems) = "" Then
            mOrder.NoOrder = True
            mOrder.Order = ""
            
            '-- �������(Order) ǥ��
            Call SetText(frmMain.spdOrder, "��������", intRow, colSTATE)
        Else
            mOrder.NoOrder = False
            mOrder.Order = strItems
        
            '-- �������(Order) ǥ��
            Call SetText(frmMain.spdOrder, "��������", intRow, colSTATE)
        End If


        '-- ���� Row
        gRow = intRow

    End With

End Sub

Private Sub GetOrder_C8000(ByVal pBarNo As String, ByVal pType As String)

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String
    Dim GetOrder    As String

    intRow = -1
    GetOrder = ""

    Call SetCommStatus("Q", pBarNo, frmMain.spdComStatus)
    
    '-- 1. �������� ��ȸ
    With frmMain
        Select Case pType
            '-- ���ڵ� ���
            Case "0"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmMain.spdOrder, i, colBARCODE)) = pBarNo Then
                        intRow = i
                        Exit For
                    End If
                Next i
        
            '-- Seq
            Case "1"
                For i = 1 To .spdOrder.DataRowCnt
                    If Val(Trim(GetText(frmMain.spdOrder, i, colSEQNO))) = Val(mOrder.Seq) Then
                        pBarNo = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarNo
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Rack/Pos
            Case "2"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmMain.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmMain.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
                        pBarNo = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Check Top
            Case "3"
                For i = 1 To .spdOrder.DataRowCnt
                    If GetText(frmMain.spdOrder, i, colCHECKBOX) = "1" Then
                        pBarNo = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                        mOrder.BarNo = pBarNo
                        intRow = i
                        Exit For
                    End If
                Next i
        End Select
    

        '-- �������忡�� ��ã����..
        If intRow < 0 Then
            intRow = .spdOrder.DataRowCnt + 1
            If .spdOrder.MaxRows < intRow Then
                .spdOrder.MaxRows = intRow
            End If
        End If

        '-- ���������� ȭ��ǥ��
        Call SetText(.spdOrder, mOrder.BarNo, intRow, colBARCODE)
        Call SetText(.spdOrder, mOrder.Seq, intRow, colSEQNO)
        Call SetText(.spdOrder, mOrder.RackNo, intRow, colRACKNO)
        Call SetText(.spdOrder, mOrder.TubePos, intRow, colPOSNO)

        '-- ����������� �����
        .spdResult.MaxRows = 0

        '-- �˻��� ���� ��������
        Call GetSampleInfo(intRow, .spdOrder)

        .spdOrder.RowHeight(-1) = 15

        '-- �������̺��� �˻��׸� �ش��ϴ� �˻�ä�� ã�ƿ��� (intRow = ���� �˻��ߴ� ���ڵ尡 �ٽ� �ö�� ��� ��ġ�� ��ã�´�.)
        strItems = GetEquipExamCode_C8000(gHOSP.MACHCD, pBarNo, intRow)
        
        '-- �˻�ä�η� ������ �����
        If Trim(strItems) = "" Then
            mOrder.NoOrder = True
            mOrder.Order = ""
            
            '-- �������(Order) ǥ��
            Call SetText(frmMain.spdOrder, "��������", intRow, colSTATE)
        Else
            mOrder.NoOrder = False
            mOrder.Order = strItems
        
            '-- �������(Order) ǥ��
            Call SetText(frmMain.spdOrder, "��������", intRow, colSTATE)
        End If


        '-- ���� Row
        gRow = intRow

    End With

End Sub


Private Sub SendData(ByVal pSendData As Variant)

    '-- ����
    comEqp.Output = pSendData
    
    imgSend.Picture = imlStatus.ListImages("RUN").ExtractIcon
    If tmrSend.Enabled = False Then
        tmrSend.Enabled = True
    Else
        tmrSend.Enabled = False
        tmrSend.Enabled = True
    End If
    
    '-- �αױ��
    Call SetRawData("[Tx]" & pSendData)

    '-- ����ǥ��
    Call SetCommStatus("S", pSendData, spdComStatus)

End Sub

Private Sub SerialRcvData_AU680()
    '��� ���� ����
    Dim strRcvBuf       As String   '������ Data
    Dim strType         As String   '������ Record Type
    Dim strBarno        As String   '������ ���ڵ��ȣ
    Dim strSeq          As String   '������ Sequence
    Dim strRackNo       As String   '������ Rack Or Disk No
    Dim strTubePos      As String   '������ Tube Position
    Dim strIntBase      As String   '������ ������ �˻��
    Dim strMachResult   As String   '������ �����
    Dim strResult       As String   '������ ���(����)
    Dim strIntResult    As String   '������ ���(����)
    Dim strQCResult     As String   '������ ���(QC)
    Dim strFlag         As String   '������ Abnormal Flag
    Dim strComm         As String   '������ Comment
    
    '������ ����
    Dim strCheck        As String   '�˻����üũ
    Dim strSeqno        As String   '�˻����
    Dim strOrderCode    As String   'ó���ڵ�
    Dim strTestName     As String   '�˻��ڵ�
    Dim strTestCode     As String   '�˻��ڵ�
    Dim strTestCodeSub  As String   '�˻��ڵ�SUB
    Dim intResPrecUse   As Integer  '�Ҽ�����ȯ����
    Dim intResPrec      As Integer  '�Ҽ����ڸ���
    Dim strResType      As String   '�Ҽ�����ȯ����
    Dim strLow          As String
    Dim strHigh         As String
    Dim strJudge        As String   '�������
    Dim strPrevRslt     As String   '�������
    
    Dim intRstRow       As String   '����������� ���� Row
    Dim intCnt          As Integer  '��� Frame ����
    Dim intCol          As Integer  '����÷� ����
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    '���� ����
    Dim strCrea         As String
    Dim streGFR         As String
    
On Error GoTo ErrHandle

    ReDim Preserve strRData(UBound(strRecvData))
    
    For i = 1 To UBound(strRecvData)
        strRData(i) = strRecvData(i)
    Next
    
    With frmMain
        For intCnt = 1 To UBound(strRData)
            strRcvBuf = strRData(intCnt)

            Call SetSQLData("RCV", strRcvBuf, "")

            strType = Mid$(strRcvBuf, 1, 2)

            Select Case strType
                Case "R "
                    strRackNo = Mid(strRcvBuf, 3, 4)
                    strTubePos = Mid(strRcvBuf, 7, 2)
                    strSeq = Mid(strRcvBuf, 9, 5)
                    strBarno = Trim(Mid(strRcvBuf, 14, gHOSP.BARLEN))
                    '-- ��������
                    With mOrder
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .Seq = strSeq
                    End With
                    '-- ����ȯ������
                    Call GetOrder_AU680(Trim$(strBarno), gHOSP.RSTTYPE)

                Case "D "    '## Result
                    strRackNo = Mid(strRcvBuf, 3, 4)
                    strTubePos = Mid(strRcvBuf, 7, 2)
                    strSeq = Mid(strRcvBuf, 10, 4)
                    strBarno = Trim$(Mid$(strRcvBuf, 14, gHOSP.BARLEN))
                    
                    '-- �������
                    With mResult
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .Seq = strSeq
                        .RsltDate = Format(Now, "yyyy-mm-dd")
                        .RsltTime = Format(Now, "hh:mm:ss")
                        .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                    End With
                    
                    '-- ���ȯ������
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    If gRow <= 0 Then
                        Exit Sub
                    End If
                    
                    If mResult.BarNo = "" Then
                        Exit Sub
                    End If

                    strTmp = Mid$(strRcvBuf, gHOSP.BARLEN + 19)
                    
                    Do While Len(strTmp) >= 11
                        strIntBase = Mid$(strTmp, 1, 3)
                        strResult = Trim(Mid$(strTmp, 4, 6))
                        strComm = Mid$(strTmp, 10, 1)
                        
                        strSeqno = ""
                        strTestCode = ""
                        strTestName = ""
                        intResPrecUse = -1
                        intResPrec = -1
                        
                        If strIntBase <> "" And strResult <> "" Then
                            blnSame = False
                            '-- �˻縶���� ���� ��������
                            For intTestNmCnt = 1 To UBound(gArrEQPNm)
                                '-- ���ä���� ����...
                                If strIntBase = gArrEQPNm(intTestNmCnt, 3) Then
                                    strCheck = "0"
                                    strSeqno = gArrEQPNm(intTestNmCnt, 1)
                                    strState = ""
                                    '-- ȯ�� ó������ ��������
'                                    If UBound(gPatTest) > 0 Then
'                                        For intOrdCnt = 1 To UBound(gPatTest)
'                                            For intTestCdCnt = 1 To UBound(gArrEQP)
'                                                '-- �˻��ڵ嵵 ���ٸ�...
'                                                If strIntBase = gArrEQP(intTestCdCnt, 3) Then
'                                                    strTestCode = gArrEQP(intTestCdCnt, 2)
'                                                    strTestName = gArrEQP(intTestCdCnt, 5)
'                                                    intResPrecUse = gArrEQP(intTestCdCnt, 7)
'                                                    intResPrec = gArrEQP(intTestCdCnt, 8)
'                                                    '-- ��������ġ�� �⺻���� �Ѵ�
'                                                    strLow = gArrEQP(intTestCdCnt, 9)
'                                                    strHigh = gArrEQP(intTestCdCnt, 10)
'
'                                                    If gPatTest(intOrdCnt) = gArrEQP(intTestCdCnt, 2) Then
'                                                        strCheck = "1"
'
'                                                        strOrderCode = gArrEQP(intTestCdCnt, 16)
'                                                        strTestCodeSub = gArrEQP(intTestCdCnt, 17)
'
'                                                        If mPatient.SEX = "M" Then
'                                                            strLow = gArrEQP(intTestCdCnt, 9)
'                                                            strHigh = gArrEQP(intTestCdCnt, 10)
'                                                        ElseIf mPatient.SEX = "F" Then
'                                                            strLow = gArrEQP(intTestCdCnt, 11)
'                                                            strHigh = gArrEQP(intTestCdCnt, 12)
'                                                        Else
'                                                            strLow = ""
'                                                            strHigh = ""
'                                                        End If
'                                                        strState = "R"
'                                                        blnSame = True
'                                                        Exit For
'                                                    End If
'                                                End If
'                                            Next
'                                        Next
'                                    End If

                                    For intTestCdCnt = 1 To UBound(gArrEQP)
                                        '-- �˻��ڵ嵵 ���ٸ�...
                                        If strIntBase = gArrEQP(intTestCdCnt, 3) Then
                                            strTestCode = gArrEQP(intTestCdCnt, 2)
                                            strTestName = gArrEQP(intTestCdCnt, 5)
                                            intResPrecUse = gArrEQP(intTestCdCnt, 7)
                                            intResPrec = gArrEQP(intTestCdCnt, 8)
                                            '-- ��������ġ�� �⺻���� �Ѵ�
                                            strLow = gArrEQP(intTestCdCnt, 9)
                                            strHigh = gArrEQP(intTestCdCnt, 10)
                                            
                                            If UBound(gPatTest) > 0 Then
                                                For intOrdCnt = 1 To UBound(gPatTest)
                                                    If gPatTest(intOrdCnt) = gArrEQP(intTestCdCnt, 2) Then
                                                        strCheck = "1"
                                                        
                                                        strOrderCode = gArrEQP(intTestCdCnt, 16)
                                                        strTestCodeSub = gArrEQP(intTestCdCnt, 17)
                                                        
                                                        If mPatient.SEX = "M" Then
                                                            strLow = gArrEQP(intTestCdCnt, 9)
                                                            strHigh = gArrEQP(intTestCdCnt, 10)
                                                        ElseIf mPatient.SEX = "F" Then
                                                            strLow = gArrEQP(intTestCdCnt, 11)
                                                            strHigh = gArrEQP(intTestCdCnt, 12)
                                                        Else
                                                            strLow = ""
                                                            strHigh = ""
                                                        End If
                                                        strState = "R"
                                                        blnSame = True
                                                        Exit For
                                                    End If
                                                Next
                                            End If
                                        End If
                                    Next
                                    
                                    If blnSame = True Then
                                        Exit For
                                    End If
                                End If
                            Next

                            '-- ���Row �߰�
                            intRstRow = .spdResult.DataRowCnt + 1
                            If .spdResult.MaxRows < intRstRow Then
                                .spdResult.MaxRows = intRstRow
                            End If

                            '-- �Ҽ��� ó��
                            strMachResult = strResult
                            If intResPrecUse = 1 Then
                                For i = 0 To intResPrec
                                    If i = 0 Then
                                        strResType = "#0"
                                    ElseIf i = 1 Then
                                        strResType = strResType & ".0"
                                    Else
                                        strResType = strResType & "0"
                                    End If
                                Next
                                strResult = Format(strResult, strResType)
                                
                            End If
                            
                            '--- �������
                            strJudge = ""
                            If IsNumeric(strLow) = True And IsNumeric(strHigh) = True Then
                                If CCur(strResult) > CCur(strLow) And CCur(strResult) < CCur(strHigh) Then
                                    strJudge = ""
                                ElseIf CCur(strHigh) <= CCur(strResult) Then
                                    strJudge = "H"
                                ElseIf CCur(strLow) >= CCur(strResult) Then
                                    strJudge = "L"
                                End If
                            End If

                            '-- ������� ǥ��("���")
                            SetText .spdOrder, "���", gRow, colSTATE

                            '-- ����ȭ�� ����� ǥ��
                            For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                If strTestName = gArrEQPNm(intCol - colSTATE, 6) Then
                                    SetText .spdOrder, strResult, gRow, intCol
                                    Exit For
                                End If
                            Next

                            '-- ��� List
                            SetText .spdResult, strCheck, intRstRow, colRCHECKBOX               'üũ
                            SetText .spdResult, strSeqno, intRstRow, colRSEQNO                  '����
                            SetText .spdResult, strOrderCode, intRstRow, colRORDERCD            'ó���ڵ�
                            '-- ó���� �������� �˻��ڵ带 �����Ѵ�.
                            If strState = "R" Then
                                SetText .spdResult, strTestCode, intRstRow, colRTESTCD          '�˻��ڵ�
                                SetText .spdResult, strTestCodeSub, intRstRow, colRSUBCD        '�˻��ڵ�SUB
                            Else
                                SetText .spdResult, "", intRstRow, colRTESTCD                   '�˻��ڵ�
                                SetText .spdResult, "", intRstRow, colRSUBCD                    '�˻��ڵ�SUB
                            End If
                            SetText .spdResult, strTestName, intRstRow, colRTESTNM              '�˻��
                            SetText .spdResult, strIntBase, intRstRow, colRCHANNEL              '���ä��
                            SetText .spdResult, strMachResult, intRstRow, colRMACHRESULT        '�����
                            SetText .spdResult, strResult, intRstRow, colRLISRESULT             'LIS���
                            SetText .spdResult, strJudge, intRstRow, colRJUDGE                  '����
                            SetText .spdResult, strLow & "~" & strHigh, intRstRow, colRREF      '����ġ
                            '-- ������� ��ȸ
                            strPrevRslt = GetPrevResult(mResult.BarNo, strIntBase, strTestCode)
                            SetText .spdResult, strPrevRslt, intRstRow, colRPREVRESULT          '�������
                            '-- H/L ����ǥ��
                            If strJudge = "H" Then
                                .spdResult.Row = intRstRow
                                .spdResult.Col = colRLISRESULT
                                .spdResult.ForeColor = vbRed
                                .spdResult.FontBold = True
                            ElseIf strJudge = "L" Then
                                .spdResult.Row = intRstRow
                                .spdResult.Col = colRLISRESULT
                                .spdResult.ForeColor = vbBlue
                                .spdResult.FontBold = True
                            Else
                                .spdResult.Row = intRstRow
                                .spdResult.Col = colRLISRESULT
                                .spdResult.ForeColor = vbBlack
                                .spdResult.FontBold = False
                            End If
                            
                            '-- ���� ����
                            Call SetLocalDB(gRow, intRstRow, "1", "")

                            '-- ���Count
                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                SetText .spdOrder, "1", gRow, colRCNT
                            Else
                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                            End If
                        End If
                        strTmp = Mid$(strTmp, 12)
                    Loop

                    .spdResult.RowHeight(-1) = 15

                    '## DB�� �������
                    If gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow, spdOrder)

                        If Res = -1 Then
                            '-- ���� ����
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "�������", gRow, colSTATE
                        Else
                            '-- ���� ����
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "����Ϸ�", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX

                                  SQL = "Update PATRESULT Set " & vbCrLf
                            SQL = SQL & " sendflag = '2' " & vbCrLf
                            SQL = SQL & " Where equipno = '" & gHOSP.HOSPCD & "' " & vbCrLf
                            SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                            SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                            If DBExec(AdoCn_Local, SQL) Then
                                '-- ����
                            End If
                        End If
                        strState = ""
                    End If
            End Select
        Next
    End With

Exit Sub

ErrHandle:
    strErrMsg = ""
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "SerialRcvData_AU680" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub SerialRcvData_TEST1()
    Dim RS_L            As ADODB.Recordset
    
    '��� ���� ����
    Dim strRcvBuf       As String   '������ Data
    Dim strType         As String   '������ Record Type
    Dim strBarno        As String   '������ ���ڵ��ȣ
    Dim strSeq          As String   '������ Sequence
    Dim strRackNo       As String   '������ Rack Or Disk No
    Dim strTubePos      As String   '������ Tube Position
    Dim strIntBase      As String   '������ ������ �˻��
    Dim strMachResult   As String   '������ �����
    Dim strResult       As String   '������ ���(����)
    Dim strIntResult    As String   '������ ���(����)
    Dim strQCResult     As String   '������ ���(QC)
    Dim strFlag         As String   '������ Abnormal Flag
    Dim strComm         As String   '������ Comment
    
    '������ ����
    Dim strCheck        As String   '�˻����üũ
    Dim strSeqno        As String   '�˻����
    Dim strOrderCode    As String   'ó���ڵ�
    Dim strTestName     As String   '�˻��ڵ�
    Dim strTestCode     As String   '�˻��ڵ�
    Dim strTestCodeSub  As String   '�˻��ڵ�SUB
    Dim intResPrecUse   As Integer  '�Ҽ�����ȯ����
    Dim intResPrec      As Integer  '�Ҽ����ڸ���
    Dim strResType      As String   '�Ҽ�����ȯ����
    Dim strLow          As String
    Dim strHigh         As String
    Dim strJudge        As String   '�������
    Dim strPrevRslt     As String   '�������
    
    Dim intRstRow       As String   '����������� ���� Row
    Dim intCnt          As Integer  '��� Frame ����
    Dim intCol          As Integer  '����÷� ����
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    
On Error GoTo ErrHandle

    ReDim Preserve strRData(UBound(strRecvData))
    
    For i = 1 To UBound(strRecvData)
        strRData(i) = strRecvData(i)
    Next
    
    With frmMain
        For intCnt = 1 To UBound(strRData)
            strRcvBuf = strRData(intCnt)

            Call SetSQLData("RCV", strRcvBuf, "")
            
            strBarno = Trim(Mid(strRcvBuf, 4, 15))
            strRackNo = Trim(Mid(strRcvBuf, 19, 2))
            strTubePos = Trim(Mid(strRcvBuf, 21, 2))
            
            strIntBase = "ESR"
            strResult = Trim$(Mid$(strRcvBuf, 25, 4))
            If IsNumeric(strResult) Then
                strResult = Val(strResult)
            End If
            
            If Trim(strBarno) <> "" Then
                '-- �������
                With mResult
                    .BarNo = strBarno
                    .RackNo = strRackNo
                    .TubePos = strTubePos
                    .RsltDate = Format(Now, "yyyy-mm-dd")
                    .RsltTime = Format(Now, "hh:mm:ss")
                    .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                End With
                
                '-- ���ȯ������
                Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                
                If gRow <= 0 Then
                    Exit Sub
                End If
                
                If mResult.BarNo = "" Then
                    Exit Sub
                End If

                strSeqno = ""
                strTestCode = ""
                strTestName = ""
                intResPrecUse = -1
                intResPrec = -1
                
                If strIntBase <> "" And strResult <> "" Then
                    SQL = ""
                    SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFMLOW,REFMHIGH,REFFLOW,REFFHIGH,RESPRECUSE,RESPREC" & vbCrLf
                    SQL = SQL & "  FROM EQPMASTER" & vbCr
                    SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                    SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                    If gPatOrdCd <> "" Then
                        SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                    End If
                    
                    Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                    If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                        strSeqno = Trim(RS_L.Fields("SEQNO"))
                        strTestCode = Trim(RS_L.Fields("TESTCODE"))
                        strTestName = Trim(RS_L.Fields("TESTNAME"))
                        '-- ����ġ
                        If mPatient.SEX = "M" Then
                            strLow = Trim(RS_L.Fields("REFMLOW"))
                            strHigh = Trim(RS_L.Fields("REFMHIGH"))
                        ElseIf mPatient.SEX = "F" Then
                            strLow = Trim(RS_L.Fields("REFFLOW"))
                            strHigh = Trim(RS_L.Fields("REFFHIGH"))
                        Else
                            '-- ��������ġ�� �⺻���� �Ѵ�
                            strLow = Trim(RS_L.Fields("REFMLOW"))
                            strHigh = Trim(RS_L.Fields("REFMHIGH"))
                        End If
                        intResPrecUse = Trim(RS_L.Fields("RESPRECUSE"))
                        intResPrec = Trim(RS_L.Fields("RESPREC"))

                        '-- ���Row �߰�
                        intRstRow = .spdResult.DataRowCnt + 1
                        If .spdResult.MaxRows < intRstRow Then
                            .spdResult.MaxRows = intRstRow
                        End If

                        '-- �Ҽ��� ó��
                        strMachResult = strResult
                        If intResPrecUse = 1 Then
                            For i = 0 To intResPrec
                                If i = 0 Then
                                    strResType = "#0"
                                ElseIf i = 1 Then
                                    strResType = strResType & ".0"
                                Else
                                    strResType = strResType & "0"
                                End If
                            Next
                            strResult = Format(strResult, strResType)
                        End If
                    
                        '--- �������
                        strJudge = ""
                        If IsNumeric(strLow) = True And IsNumeric(strHigh) = True Then
                            If CCur(strResult) > CCur(strLow) And CCur(strResult) < CCur(strHigh) Then
                                strJudge = ""
                            ElseIf CCur(strHigh) <= CCur(strResult) Then
                                strJudge = "H"
                            ElseIf CCur(strLow) >= CCur(strResult) Then
                                strJudge = "L"
                            End If
                        End If
    
                        '-- ������� ǥ��("���")
                        SetText .spdOrder, "���", gRow, colSTATE

                        '-- ����ȭ�� ����� ǥ��
                        For intCol = colSTATE + 1 To .spdOrder.MaxCols
                            If strTestName = gArrEQPNm(intCol - colSTATE, 6) Then
                                SetText .spdOrder, strResult, gRow, intCol
                                
                                strOrderCode = gArrEQP(intCol - colSTATE, 16)
                                
                                Exit For
                            End If
                        Next

                        '-- ��� List
                        SetText .spdResult, strCheck, intRstRow, colRCHECKBOX               'üũ
                        SetText .spdResult, strSeqno, intRstRow, colRSEQNO                  '����
                        SetText .spdResult, strOrderCode, intRstRow, colRORDERCD            'ó���ڵ�
                        SetText .spdResult, strTestCode, intRstRow, colRTESTCD          '�˻��ڵ�
                        SetText .spdResult, strTestCodeSub, intRstRow, colRSUBCD        '�˻��ڵ�SUB
                        SetText .spdResult, strTestName, intRstRow, colRTESTNM              '�˻��
                        SetText .spdResult, strIntBase, intRstRow, colRCHANNEL              '���ä��
                        SetText .spdResult, strMachResult, intRstRow, colRMACHRESULT        '�����
                        SetText .spdResult, strResult, intRstRow, colRLISRESULT             'LIS���
                        SetText .spdResult, strJudge, intRstRow, colRJUDGE                  '����
                        SetText .spdResult, strLow & "~" & strHigh, intRstRow, colRREF      '����ġ
                        
                        '-- ������� ��ȸ
                        strPrevRslt = GetPrevResult(mResult.BarNo, strIntBase, strTestCode)
                        SetText .spdResult, strPrevRslt, intRstRow, colRPREVRESULT          '�������
                        
                        '-- H/L ����ǥ��
                        If strJudge = "H" Then
                            .spdResult.Row = intRstRow
                            .spdResult.Col = colRLISRESULT
                            .spdResult.ForeColor = vbRed
                            .spdResult.FontBold = True
                        ElseIf strJudge = "L" Then
                            .spdResult.Row = intRstRow
                            .spdResult.Col = colRLISRESULT
                            .spdResult.ForeColor = vbBlue
                            .spdResult.FontBold = True
                        Else
                            .spdResult.Row = intRstRow
                            .spdResult.Col = colRLISRESULT
                            .spdResult.ForeColor = vbBlack
                            .spdResult.FontBold = False
                        End If
                        
                        '-- ���� ����
                        Call SetLocalDB(gRow, intRstRow, "1", "")
    
                        '-- ���Count
                        If GetText(.spdOrder, gRow, colRCNT) = "" Then
                            SetText .spdOrder, "1", gRow, colRCNT
                        Else
                            SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                        End If
                        strState = "R"
                        
                    End If

                    .spdResult.RowHeight(-1) = 15
    
                    '## DB�� �������
                    If gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow, spdOrder)
    
                        If Res = -1 Then
                            '-- ���� ����
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "�������", gRow, colSTATE
                        Else
                            '-- ���� ����
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "����Ϸ�", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX
    
                                  SQL = "Update PATRESULT Set " & vbCrLf
                            SQL = SQL & " sendflag = '2' " & vbCrLf
                            SQL = SQL & " Where equipno = '" & gHOSP.HOSPCD & "' " & vbCrLf
                            SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                            SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
    
                            If DBExec(AdoCn_Local, SQL) Then
                                '-- ����
                            End If
                        End If
                        strState = ""
                    End If
                End If
            End If
        Next
    End With

Exit Sub

ErrHandle:
    strErrMsg = ""
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "SerialRcvData_TEST1" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub SerialRcvData_INDIKO()
    Dim RS_L            As ADODB.Recordset
    
    '��� ���� ����
    Dim strRcvBuf       As String   '������ Data
    Dim strType         As String   '������ Record Type
    Dim strBarno        As String   '������ ���ڵ��ȣ
    Dim strSeq          As String   '������ Sequence
    Dim strRackNo       As String   '������ Rack Or Disk No
    Dim strTubePos      As String   '������ Tube Position
    Dim strIntBase      As String   '������ ������ �˻��
    Dim strMachResult   As String   '������ �����
    Dim strResult       As String   '������ ���(����)
    Dim strIntResult    As String   '������ ���(����)
    Dim strQCResult     As String   '������ ���(QC)
    Dim strFlag         As String   '������ Abnormal Flag
    Dim strComm         As String   '������ Comment
    
    '������ ����
    Dim strCheck        As String   '�˻����üũ
    Dim strSeqno        As String   '�˻����
    Dim strOrderCode    As String   'ó���ڵ�
    Dim strTestName     As String   '�˻��ڵ�
    Dim strTestCode     As String   '�˻��ڵ�
    Dim strTestCodeSub  As String   '�˻��ڵ�SUB
    Dim intResPrecUse   As Integer  '�Ҽ�����ȯ����
    Dim intResPrec      As Integer  '�Ҽ����ڸ���
    Dim strResType      As String   '�Ҽ�����ȯ����
    Dim strLow          As String
    Dim strHigh         As String
    Dim strJudge        As String   '�������
    Dim strPrevRslt     As String   '�������
    
    Dim intRstRow       As String   '����������� ���� Row
    Dim intCnt          As Integer  '��� Frame ����
    Dim intCol          As Integer  '����÷� ����
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    
On Error GoTo ErrHandle

    ReDim Preserve strRData(UBound(strRecvData))
    
    For i = 1 To UBound(strRecvData)
        strRData(i) = strRecvData(i)
    Next
    
    With frmMain
        For intCnt = 1 To UBound(strRData)
            strRcvBuf = strRData(intCnt)

            Call SetSQLData("RCV", strRcvBuf, "")
            
            strType = Mid$(strRcvBuf, 2, 1)
            If strType = "|" Then
                strType = Mid$(strRcvBuf, 1, 1)
            End If
            
            Select Case strType
                Case "H"    '## Header
                Case "P"    '## Patient
                Case "Q"    '## Request Information
                    strBarno = Trim$(mGetP(mGetP(strRcvBuf, 3, "|"), 2, "^"))
                    
                    With mOrder
                        .NoOrder = False
                        .BarNo = strBarno
                    End With
                    
                    Call GetOrder_INDIKO(strBarno, gHOSP.RSTTYPE)
                    
                    strState = "Q"
                
                Case "O"
                    strBarno = Trim$(mGetP(mGetP(strRcvBuf, 3, "|"), 1, "^"))
                    
                    If Trim(strBarno) <> "" And strOldBarno <> strBarno Then
                        '-- �������
                        With mResult
                            .BarNo = strBarno
                            .RsltDate = Format(Now, "yyyy-mm-dd")
                            .RsltTime = Format(Now, "hh:mm:ss")
                            .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                        End With
                        
                        '-- ���ȯ������
                        Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                        
                        If gRow <= 0 Then
                            Exit Sub
                        End If
                        
                        If mResult.BarNo = "" Then
                            Exit Sub
                        End If
                        
                        strOldBarno = strBarno
                    End If
                    
                Case "R"
                    '6R|1|^^^ALP^0.0|71|U/l|0^0^|N||F\R||||20170119170306|Analyzer 1

                    strSeqno = ""
                    strTestCode = ""
                    strTestName = ""
                    intResPrecUse = -1
                    intResPrec = -1
                    
                    strIntBase = Trim$(mGetP(mGetP(strRcvBuf, 3, "|"), 4, "^"))
                    '## ������� ����
                    strResult = Trim$(mGetP(strRcvBuf, 4, "|"))
                    If InStr(strResult, "^") > 0 Then
                        '## ������� ����
                        strResult = mGetP(strResult, 2, "^")
                    End If


                    If strIntBase <> "" And strResult <> "" Then
                        SQL = ""
                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFMLOW,REFMHIGH,REFFLOW,REFFHIGH,RESPRECUSE,RESPREC" & vbCrLf
                        SQL = SQL & "  FROM EQPMASTER" & vbCr
                        SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                        If gPatOrdCd <> "" Then
                            SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                        End If
                        
                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                            strSeqno = Trim(RS_L.Fields("SEQNO"))
                            strTestCode = Trim(RS_L.Fields("TESTCODE")) & ""
                            strTestName = Trim(RS_L.Fields("TESTNAME")) & ""
                            '-- ����ġ
                            If mPatient.SEX = "M" Then
                                strLow = Trim(RS_L.Fields("REFMLOW")) & ""
                                strHigh = Trim(RS_L.Fields("REFMHIGH")) & ""
                            ElseIf mPatient.SEX = "F" Then
                                strLow = Trim(RS_L.Fields("REFFLOW")) & ""
                                strHigh = Trim(RS_L.Fields("REFFHIGH")) & ""
                            Else
                                '-- ��������ġ�� �⺻���� �Ѵ�
                                strLow = Trim(RS_L.Fields("REFMLOW")) & ""
                                strHigh = Trim(RS_L.Fields("REFMHIGH")) & ""
                            End If
                            intResPrecUse = Trim(RS_L.Fields("RESPRECUSE")) & ""
                            intResPrec = Trim(RS_L.Fields("RESPREC")) & ""
    
                            '-- ���Row �߰�
                            intRstRow = .spdResult.DataRowCnt + 1
                            If .spdResult.MaxRows < intRstRow Then
                                .spdResult.MaxRows = intRstRow
                            End If
    
                            '-- �Ҽ��� ó��
                            strMachResult = strResult
                            If intResPrecUse = 1 Then
                                For i = 0 To intResPrec
                                    If i = 0 Then
                                        strResType = "#0"
                                    ElseIf i = 1 Then
                                        strResType = strResType & ".0"
                                    Else
                                        strResType = strResType & "0"
                                    End If
                                Next
                                strResult = Format(strResult, strResType)
                            End If
                        
                            '--- �������
                            strJudge = ""
                            If IsNumeric(strLow) = True And IsNumeric(strHigh) = True Then
                                If CCur(strResult) > CCur(strLow) And CCur(strResult) < CCur(strHigh) Then
                                    strJudge = ""
                                ElseIf CCur(strHigh) <= CCur(strResult) Then
                                    strJudge = "H"
                                ElseIf CCur(strLow) >= CCur(strResult) Then
                                    strJudge = "L"
                                End If
                            End If
        
                            '-- ������� ǥ��("���")
                            SetText .spdOrder, "���", gRow, colSTATE
    
                            '-- ����ȭ�� ����� ǥ��
                            For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                If strTestName = gArrEQPNm(intCol - colSTATE, 6) Then
                                    SetText .spdOrder, strResult, gRow, intCol
                                    
                                    strOrderCode = gArrEQP(intCol - colSTATE, 16)
                                    
                                    Exit For
                                End If
                            Next
    
                            '-- ��� List
                            SetText .spdResult, strCheck, intRstRow, colRCHECKBOX               'üũ
                            SetText .spdResult, strSeqno, intRstRow, colRSEQNO                  '����
                            SetText .spdResult, strOrderCode, intRstRow, colRORDERCD            'ó���ڵ�
                            SetText .spdResult, strTestCode, intRstRow, colRTESTCD          '�˻��ڵ�
                            SetText .spdResult, strTestCodeSub, intRstRow, colRSUBCD        '�˻��ڵ�SUB
                            SetText .spdResult, strTestName, intRstRow, colRTESTNM              '�˻��
                            SetText .spdResult, strIntBase, intRstRow, colRCHANNEL              '���ä��
                            SetText .spdResult, strMachResult, intRstRow, colRMACHRESULT        '�����
                            SetText .spdResult, strResult, intRstRow, colRLISRESULT             'LIS���
                            SetText .spdResult, strJudge, intRstRow, colRJUDGE                  '����
                            SetText .spdResult, strLow & "~" & strHigh, intRstRow, colRREF      '����ġ
                            
                            '-- ������� ��ȸ
                            strPrevRslt = GetPrevResult(mResult.BarNo, strIntBase, strTestCode)
                            SetText .spdResult, strPrevRslt, intRstRow, colRPREVRESULT          '�������
                            
                            '-- H/L ����ǥ��
                            If strJudge = "H" Then
                                .spdResult.Row = intRstRow
                                .spdResult.Col = colRLISRESULT
                                .spdResult.ForeColor = vbRed
                                .spdResult.FontBold = True
                            ElseIf strJudge = "L" Then
                                .spdResult.Row = intRstRow
                                .spdResult.Col = colRLISRESULT
                                .spdResult.ForeColor = vbBlue
                                .spdResult.FontBold = True
                            Else
                                .spdResult.Row = intRstRow
                                .spdResult.Col = colRLISRESULT
                                .spdResult.ForeColor = vbBlack
                                .spdResult.FontBold = False
                            End If
                            
                            '-- ���� ����
                            Call SetLocalDB(gRow, intRstRow, "1", "")
        
                            '-- ���Count
                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                SetText .spdOrder, "1", gRow, colRCNT
                            Else
                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                            End If
                            strState = "R"
                            
                        End If
    
                        .spdResult.RowHeight(-1) = 15
        
                    End If
                    
                Case "L"
            
                    '## DB�� �������
                    If gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow, spdOrder)
    
                        If Res = -1 Then
                            '-- ���� ����
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "�������", gRow, colSTATE
                        Else
                            '-- ���� ����
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "����Ϸ�", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX
    
                                  SQL = "Update PATRESULT Set " & vbCrLf
                            SQL = SQL & " sendflag = '2' " & vbCrLf
                            SQL = SQL & " Where equipno = '" & gHOSP.HOSPCD & "' " & vbCrLf
                            SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                            SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
    
                            If DBExec(AdoCn_Local, SQL) Then
                                '-- ����
                            End If
                        End If
                        strState = ""
                    End If
                
            End Select
        Next
    End With

Exit Sub

ErrHandle:
    strErrMsg = ""
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "SerialRcvData_INDIKO" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub SerialRcvData_XN1000()
    Dim RS_L            As ADODB.Recordset
    
    '��� ���� ����
    Dim strRcvBuf       As String   '������ Data
    Dim strType         As String   '������ Record Type
    Dim strBarno        As String   '������ ���ڵ��ȣ
    Dim strSeq          As String   '������ Sequence
    Dim strRackNo       As String   '������ Rack Or Disk No
    Dim strTubePos      As String   '������ Tube Position
    Dim strIntBase      As String   '������ ������ �˻��
    Dim strMachResult   As String   '������ �����
    Dim strResult       As String   '������ ���(����)
    Dim strIntResult    As String   '������ ���(����)
    Dim strQCResult     As String   '������ ���(QC)
    Dim strFlag         As String   '������ Abnormal Flag
    Dim strComm         As String   '������ Comment
    
    '������ ����
    Dim strCheck        As String   '�˻����üũ
    Dim strSeqno        As String   '�˻����
    Dim strOrderCode    As String   'ó���ڵ�
    Dim strTestName     As String   '�˻��ڵ�
    Dim strTestCode     As String   '�˻��ڵ�
    Dim strTestCodeSub  As String   '�˻��ڵ�SUB
    Dim intResPrecUse   As Integer  '�Ҽ�����ȯ����
    Dim intResPrec      As Integer  '�Ҽ����ڸ���
    Dim strResType      As String   '�Ҽ�����ȯ����
    Dim strLow          As String
    Dim strHigh         As String
    Dim strJudge        As String   '�������
    Dim strPrevRslt     As String   '�������
    
    Dim intRstRow       As String   '����������� ���� Row
    Dim intCnt          As Integer  '��� Frame ����
    Dim intCol          As Integer  '����÷� ����
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    Dim strTemp1        As String
    Dim strTemp2        As String
    
    Dim strWBC          As String
    Dim strNeut         As String
    
    
On Error GoTo ErrHandle

    ReDim Preserve strRData(UBound(strRecvData))
    
    For i = 1 To UBound(strRecvData)
        strRData(i) = strRecvData(i)
    Next
    
    With frmMain
        For intCnt = 1 To UBound(strRData)
            strRcvBuf = strRData(intCnt)

            Call SetSQLData("RCV", strRcvBuf, "")
            
            strType = Mid$(strRcvBuf, 2, 1)
            If strType = "|" Then
                strType = Mid$(strRcvBuf, 1, 1)
            End If
            
            Select Case strType
                Case "H"    '## Header
                    strWBC = ""
                    strNeut = ""
                
                Case "P"    '## Patient
                Case "Q"    '## Request Information
                    strTemp1 = mGetP(strRcvBuf, 3, "|")
                    
                    strRackNo = mGetP(strTemp1, 1, "^")
                    strTubePos = mGetP(strTemp1, 2, "^")
                    strBarno = Trim$(mGetP(strTemp1, 3, "^"))
                    
                    With mOrder
                        .NoOrder = False
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                    End With
                    
                    Call GetOrder_XN1000(strBarno, gHOSP.RSTTYPE)
                    
                    strState = "Q"
                
                Case "O"
                    strTemp1 = mGetP(strRcvBuf, 4, "|")
                    strBarno = Trim(mGetP(strTemp1, 3, "^"))
                    strRackNo = mGetP(strTemp1, 1, "^")
                    strTubePos = mGetP(strTemp1, 2, "^")
                

                    With mResult
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .RsltDate = Format(Now, "yyyy-mm-dd")
                        .RsltTime = Format(Now, "hh:mm:ss")
                        .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                    End With


                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                Case "R"
                    strTemp1 = mGetP(strRcvBuf, 3, "|")
                    strIntBase = mGetP(strTemp1, 5, "^")
                    strTemp2 = mGetP(strRcvBuf, 4, "|")
                    strFlag = mGetP(strRcvBuf, 7, "|")
                    
                    If InStr(strTemp2, "^") > 0 Then
                        '## ������� ����
                        strResult = mGetP(strTemp2, 2, "^")
                    Else
                        '## ������� ����
                        strResult = strTemp2
                    End If
                    
'                    If strIntBase = "WBC" And IsNumeric(strResult) Then
'                        strWBC = strResult * 1000
'                    End If
'
'                    If strIntBase = "NEUT%" And IsNumeric(strResult) Then
'                        strNeut = strResult / 100
'                    End If
                        
RST:
                    
                    If strIntBase <> "" And strResult <> "" Then
                        SQL = ""
                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFMLOW,REFMHIGH,REFFLOW,REFFHIGH,RESPRECUSE,RESPREC" & vbCrLf
                        SQL = SQL & "  FROM EQPMASTER" & vbCr
                        SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                        If gPatOrdCd <> "" Then
                            SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                        End If
                        
                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                            strSeqno = Trim(RS_L.Fields("SEQNO"))
                            strTestCode = Trim(RS_L.Fields("TESTCODE")) & ""
                            strTestName = Trim(RS_L.Fields("TESTNAME")) & ""
                            '-- ����ġ
                            If mPatient.SEX = "M" Then
                                strLow = Trim(RS_L.Fields("REFMLOW")) & ""
                                strHigh = Trim(RS_L.Fields("REFMHIGH")) & ""
                            ElseIf mPatient.SEX = "F" Then
                                strLow = Trim(RS_L.Fields("REFFLOW")) & ""
                                strHigh = Trim(RS_L.Fields("REFFHIGH")) & ""
                            Else
                                '-- ��������ġ�� �⺻���� �Ѵ�
                                strLow = Trim(RS_L.Fields("REFMLOW")) & ""
                                strHigh = Trim(RS_L.Fields("REFMHIGH")) & ""
                            End If
                            intResPrecUse = Trim(RS_L.Fields("RESPRECUSE")) & ""
                            intResPrec = Trim(RS_L.Fields("RESPREC")) & ""
    
                            '-- ���Row �߰�
                            intRstRow = .spdResult.DataRowCnt + 1
                            If .spdResult.MaxRows < intRstRow Then
                                .spdResult.MaxRows = intRstRow
                            End If
    
                            '-- �Ҽ��� ó��
                            strMachResult = strResult
                            If intResPrecUse = 1 Then
                                For i = 0 To intResPrec
                                    If i = 0 Then
                                        strResType = "#0"
                                    ElseIf i = 1 Then
                                        strResType = strResType & ".0"
                                    Else
                                        strResType = strResType & "0"
                                    End If
                                Next
                                strResult = Format(strResult, strResType)
                            End If
                        
                            If strIntBase = "WBC" And IsNumeric(strResult) Then
                                strWBC = strResult * 1000
                            End If
                            
                            If strIntBase = "NEUT%" And IsNumeric(strResult) Then
                                strNeut = strResult / 100
                            End If
                        
                            '--- �������
                            strJudge = ""
                            If IsNumeric(strLow) = True And IsNumeric(strHigh) = True Then
                                If CCur(strResult) > CCur(strLow) And CCur(strResult) < CCur(strHigh) Then
                                    strJudge = ""
                                ElseIf CCur(strHigh) <= CCur(strResult) Then
                                    strJudge = "H"
                                ElseIf CCur(strLow) >= CCur(strResult) Then
                                    strJudge = "L"
                                End If
                            End If
        
                            '-- ������� ǥ��("���")
                            SetText .spdOrder, "���", gRow, colSTATE
    
                            '-- ����ȭ�� ����� ǥ��
                            For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                If strTestName = gArrEQPNm(intCol - colSTATE, 6) Then
                                    SetText .spdOrder, strResult, gRow, intCol
                                    
                                    strTestCodeSub = gArrEQP(intCol - colSTATE, 16)
                                    
                                    Exit For
                                End If
                            Next
    
                            '-- ��� List
                            SetText .spdResult, strCheck, intRstRow, colRCHECKBOX               'üũ
                            SetText .spdResult, strSeqno, intRstRow, colRSEQNO                  '����
                            SetText .spdResult, strOrderCode, intRstRow, colRORDERCD            'ó���ڵ�
                            SetText .spdResult, strTestCode, intRstRow, colRTESTCD          '�˻��ڵ�
                            SetText .spdResult, strTestCodeSub, intRstRow, colRSUBCD        '�˻��ڵ�SUB
                            SetText .spdResult, strTestName, intRstRow, colRTESTNM              '�˻��
                            SetText .spdResult, strIntBase, intRstRow, colRCHANNEL              '���ä��
                            SetText .spdResult, strMachResult, intRstRow, colRMACHRESULT        '�����
                            SetText .spdResult, strResult, intRstRow, colRLISRESULT             'LIS���
                            SetText .spdResult, strJudge, intRstRow, colRJUDGE                  '����
                            SetText .spdResult, strLow & "~" & strHigh, intRstRow, colRREF      '����ġ
                            
                            '-- ������� ��ȸ
                            strPrevRslt = GetPrevResult(mResult.BarNo, strIntBase, strTestCode)
                            SetText .spdResult, strPrevRslt, intRstRow, colRPREVRESULT          '�������
                            
                            '-- H/L ����ǥ��
                            If strJudge = "H" Then
                                .spdResult.Row = intRstRow
                                .spdResult.Col = colRLISRESULT
                                .spdResult.ForeColor = vbRed
                                .spdResult.FontBold = True
                            ElseIf strJudge = "L" Then
                                .spdResult.Row = intRstRow
                                .spdResult.Col = colRLISRESULT
                                .spdResult.ForeColor = vbBlue
                                .spdResult.FontBold = True
                            Else
                                .spdResult.Row = intRstRow
                                .spdResult.Col = colRLISRESULT
                                .spdResult.ForeColor = vbBlack
                                .spdResult.FontBold = False
                            End If
                            
                            '-- ���� ����
                            Call SetLocalDB(gRow, intRstRow, "1", "")
        
                            '-- ���Count
                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                SetText .spdOrder, "1", gRow, colRCNT
                            Else
                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                            End If
                            strState = "R"
                            
                        End If
    
                        .spdResult.RowHeight(-1) = 15
        
                    End If
                    
                    If strWBC <> "" And strNeut <> "" Then
                        ''ANC = (wbc * 1000 * neut%) / 100
                        strIntBase = "ANC"
                        strResult = (strWBC * strNeut)
                        strResult = Format(strResult, "##0")
                        strWBC = ""
                        strNeut = ""
                        GoTo RST
                    End If
                    
                    
                Case "L"
            
                    '## DB�� �������
                    If gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow, spdOrder)
    
                        If Res = -1 Then
                            '-- ���� ����
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "�������", gRow, colSTATE
                        Else
                            '-- ���� ����
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "����Ϸ�", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX
    
                                  SQL = "Update PATRESULT Set " & vbCrLf
                            SQL = SQL & " sendflag = '2' " & vbCrLf
                            SQL = SQL & " Where equipno = '" & gHOSP.HOSPCD & "' " & vbCrLf
                            SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                            SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
    
                            If DBExec(AdoCn_Local, SQL) Then
                                '-- ����
                            End If
                        End If
                        strState = ""
                    End If
                
            End Select
        Next
    End With

Exit Sub

ErrHandle:
    strErrMsg = ""
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "SerialRcvData_INDIKO" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub SerialRcvData_C8000()
    Dim RS_L            As ADODB.Recordset
    
    '��� ���� ����
    Dim strRcvBuf       As String   '������ Data
    Dim strType         As String   '������ Record Type
    Dim strBarno        As String   '������ ���ڵ��ȣ
    Dim strSeq          As String   '������ Sequence
    Dim strRackNo       As String   '������ Rack Or Disk No
    Dim strTubePos      As String   '������ Tube Position
    Dim strIntBase      As String   '������ ������ �˻��
    Dim strMachResult   As String   '������ �����
    Dim strResult       As String   '������ ���(����)
    Dim strIntResult    As String   '������ ���(����)
    Dim strQCResult     As String   '������ ���(QC)
    Dim strFlag         As String   '������ Abnormal Flag
    Dim strComm         As String   '������ Comment
    
    '������ ����
    Dim strCheck        As String   '�˻����üũ
    Dim strSeqno        As String   '�˻����
    Dim strOrderCode    As String   'ó���ڵ�
    Dim strTestName     As String   '�˻��ڵ�
    Dim strTestCode     As String   '�˻��ڵ�
    Dim strTestCodeSub  As String   '�˻��ڵ�SUB
    Dim intResPrecUse   As Integer  '�Ҽ�����ȯ����
    Dim intResPrec      As Integer  '�Ҽ����ڸ���
    Dim strResType      As String   '�Ҽ�����ȯ����
    Dim strLow          As String
    Dim strHigh         As String
    Dim strJudge        As String   '�������
    Dim strPrevRslt     As String   '�������
    
    Dim intRstRow       As String   '����������� ���� Row
    Dim intCnt          As Integer  '��� Frame ����
    Dim intCol          As Integer  '����÷� ����
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    Dim strTemp1        As String
    Dim strTemp2        As String
    
    Dim strWBC          As String
    Dim strNeut         As String
    
    
On Error GoTo ErrHandle

    ReDim Preserve strRData(UBound(strRecvData))
    
    For i = 1 To UBound(strRecvData)
        strRData(i) = strRecvData(i)
    Next
    
    With frmMain
        For intCnt = 1 To UBound(strRData)
            strRcvBuf = strRData(intCnt)

            Call SetSQLData("RCV", strRcvBuf, "")
            
            strType = Mid$(strRcvBuf, 2, 1)
            If strType = "|" Then
                strType = Mid$(strRcvBuf, 1, 1)
            End If
            
            Select Case strType
                Case "H"    '## Header
                Case "P"    '## Patient
                Case "Q"    '## Request Information
                    strTemp1 = mGetP(strRcvBuf, 3, "|")
                    
                    strBarno = Trim$(mGetP(strTemp1, 3, "^"))
                    strSeq = Trim$(mGetP(strTemp1, 4, "^"))
                    strRackNo = Trim$(mGetP(strTemp1, 5, "^"))
                    strTubePos = Trim$(mGetP(strTemp1, 6, "^"))
                    
                    With mOrder
                        .NoOrder = False
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .Seq = strSeq
                    End With
                    
                    Call GetOrder_C8000(strBarno, gHOSP.RSTTYPE)
                    
                    strState = "Q"
                
                Case "O"
                    strBarno = Trim(mGetP(strRcvBuf, 3, "^"))
                    
                    strTemp1 = mGetP(strRcvBuf, 4, "|")
                    strBarno = mGetP(strTemp1, 1, "^")
                    strSeq = mGetP(strTemp1, 2, "^")
                    strTubePos = Trim(mGetP(strTemp1, 3, "^"))

                    With mResult
                        .BarNo = strBarno
                        .Seq = strSeq
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .RsltDate = Format(Now, "yyyy-mm-dd")
                        .RsltTime = Format(Now, "hh:mm:ss")
                        .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                    End With

                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                Case "R"
                    strTemp1 = mGetP(mGetP(strRcvBuf, 3, "|"), 4, "^")
                    strTemp1 = mGetP(strTemp1, 1, "/")
                    strIntBase = strTemp1
                    
                    strTemp2 = mGetP(strRcvBuf, 4, "|")
                    strFlag = mGetP(strRcvBuf, 7, "|")
                    
                    If InStr(strTemp2, "^") > 0 Then
                        '## ������� ����
                        strResult = mGetP(strTemp2, 1, "^")
                        Select Case strResult
                            Case "-1":  strResult = "Negative"
                            Case "0":   strResult = "Grayzone"
                            Case "1":   strResult = "Posiive"
                        End Select
                        '## ������� ����
                        strIntResult = mGetP(strTemp2, 2, "^")
                        
                        '--Negative(0.23)
                        strResult = strResult & "(" & strIntResult & ")"
                    Else
                        '## ������� ����
                        strIntResult = strTemp2
                        'HBsAb
                        If strIntBase = "739" Then
                            If IsNumeric(strIntResult) Then
                                If CCur(strIntResult) < 10 Then
                                    strResult = "Negative(" & strIntResult & ")"
                                Else
                                    strResult = "Positive(" & strIntResult & ")"
                                End If
                            Else
                                strResult = strIntResult
                            End If
                        Else
                            Select Case strFlag
                                Case "LL"
                                    strResult = "<" & strIntResult
                                Case "HH"
                                    strResult = ">" & strIntResult
                                Case Else
                                    strResult = strIntResult
                            End Select
                        End If
                    End If
                    
                    If strIntBase <> "" And strResult <> "" Then
                        SQL = ""
                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFMLOW,REFMHIGH,REFFLOW,REFFHIGH,RESPRECUSE,RESPREC" & vbCrLf
                        SQL = SQL & "  FROM EQPMASTER" & vbCr
                        SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                        If gPatOrdCd <> "" Then
                            SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                        End If
                        
                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                            strSeqno = Trim(RS_L.Fields("SEQNO"))
                            strTestCode = Trim(RS_L.Fields("TESTCODE")) & ""
                            strTestName = Trim(RS_L.Fields("TESTNAME")) & ""
                            '-- ����ġ
                            If mPatient.SEX = "M" Then
                                strLow = Trim(RS_L.Fields("REFMLOW")) & ""
                                strHigh = Trim(RS_L.Fields("REFMHIGH")) & ""
                            ElseIf mPatient.SEX = "F" Then
                                strLow = Trim(RS_L.Fields("REFFLOW")) & ""
                                strHigh = Trim(RS_L.Fields("REFFHIGH")) & ""
                            Else
                                '-- ��������ġ�� �⺻���� �Ѵ�
                                strLow = Trim(RS_L.Fields("REFMLOW")) & ""
                                strHigh = Trim(RS_L.Fields("REFMHIGH")) & ""
                            End If
                            intResPrecUse = Trim(RS_L.Fields("RESPRECUSE")) & ""
                            intResPrec = Trim(RS_L.Fields("RESPREC")) & ""
    
                            '-- ���Row �߰�
                            intRstRow = .spdResult.DataRowCnt + 1
                            If .spdResult.MaxRows < intRstRow Then
                                .spdResult.MaxRows = intRstRow
                            End If
    
                            '-- �Ҽ��� ó��
                            strMachResult = strResult
                            If intResPrecUse = 1 Then
                                For i = 0 To intResPrec
                                    If i = 0 Then
                                        strResType = "#0"
                                    ElseIf i = 1 Then
                                        strResType = strResType & ".0"
                                    Else
                                        strResType = strResType & "0"
                                    End If
                                Next
                                strResult = Format(strResult, strResType)
                            End If
                        
                            '--- �������
                            strJudge = ""
                            If IsNumeric(strLow) = True And IsNumeric(strHigh) = True Then
                                If CCur(strResult) > CCur(strLow) And CCur(strResult) < CCur(strHigh) Then
                                    strJudge = ""
                                ElseIf CCur(strHigh) <= CCur(strResult) Then
                                    strJudge = "H"
                                ElseIf CCur(strLow) >= CCur(strResult) Then
                                    strJudge = "L"
                                End If
                            End If
        
                            '-- ������� ǥ��("���")
                            SetText .spdOrder, "���", gRow, colSTATE
    
                            '-- ����ȭ�� ����� ǥ��
                            For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                If strTestName = gArrEQPNm(intCol - colSTATE, 6) Then
                                    SetText .spdOrder, strResult, gRow, intCol
                                    
                                    strTestCodeSub = gArrEQP(intCol - colSTATE, 16)
                                    
                                    Exit For
                                End If
                            Next
    
                            '-- ��� List
                            SetText .spdResult, strCheck, intRstRow, colRCHECKBOX               'üũ
                            SetText .spdResult, strSeqno, intRstRow, colRSEQNO                  '����
                            SetText .spdResult, strOrderCode, intRstRow, colRORDERCD            'ó���ڵ�
                            SetText .spdResult, strTestCode, intRstRow, colRTESTCD          '�˻��ڵ�
                            SetText .spdResult, strTestCodeSub, intRstRow, colRSUBCD        '�˻��ڵ�SUB
                            SetText .spdResult, strTestName, intRstRow, colRTESTNM              '�˻��
                            SetText .spdResult, strIntBase, intRstRow, colRCHANNEL              '���ä��
                            SetText .spdResult, strMachResult, intRstRow, colRMACHRESULT        '�����
                            SetText .spdResult, strResult, intRstRow, colRLISRESULT             'LIS���
                            SetText .spdResult, strJudge, intRstRow, colRJUDGE                  '����
                            SetText .spdResult, strLow & "~" & strHigh, intRstRow, colRREF      '����ġ
                            
                            '-- ������� ��ȸ
                            strPrevRslt = GetPrevResult(mResult.BarNo, strIntBase, strTestCode)
                            SetText .spdResult, strPrevRslt, intRstRow, colRPREVRESULT          '�������
                            
                            '-- H/L ����ǥ��
                            If strJudge = "H" Then
                                .spdResult.Row = intRstRow
                                .spdResult.Col = colRLISRESULT
                                .spdResult.ForeColor = vbRed
                                .spdResult.FontBold = True
                            ElseIf strJudge = "L" Then
                                .spdResult.Row = intRstRow
                                .spdResult.Col = colRLISRESULT
                                .spdResult.ForeColor = vbBlue
                                .spdResult.FontBold = True
                            Else
                                .spdResult.Row = intRstRow
                                .spdResult.Col = colRLISRESULT
                                .spdResult.ForeColor = vbBlack
                                .spdResult.FontBold = False
                            End If
                            
                            '-- ���� ����
                            Call SetLocalDB(gRow, intRstRow, "1", "")
        
                            '-- ���Count
                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                SetText .spdOrder, "1", gRow, colRCNT
                            Else
                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                            End If
                            strState = "R"
                            
                        End If
    
                        .spdResult.RowHeight(-1) = 15
        
                    End If
                    
                Case "L"
            
                    '## DB�� �������
                    If gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow, spdOrder)
    
                        If Res = -1 Then
                            '-- ���� ����
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "�������", gRow, colSTATE
                        Else
                            '-- ���� ����
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "����Ϸ�", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX
    
                                  SQL = "Update PATRESULT Set " & vbCrLf
                            SQL = SQL & " sendflag = '2' " & vbCrLf
                            SQL = SQL & " Where equipno = '" & gHOSP.HOSPCD & "' " & vbCrLf
                            SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                            SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
    
                            If DBExec(AdoCn_Local, SQL) Then
                                '-- ����
                            End If
                        End If
                        strState = ""
                    End If
                
            End Select
        Next
    End With

Exit Sub

ErrHandle:
    strErrMsg = ""
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "SerialRcvData_INDIKO" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub


Private Sub SerialRcvData_ISMART30()
    Dim RS_L            As ADODB.Recordset
    
    '��� ���� ����
    Dim strRcvBuf       As String   '������ Data
    Dim strType         As String   '������ Record Type
    Dim strBarno        As String   '������ ���ڵ��ȣ
    Dim strSeq          As String   '������ Sequence
    Dim strRackNo       As String   '������ Rack Or Disk No
    Dim strTubePos      As String   '������ Tube Position
    Dim strIntBase      As String   '������ ������ �˻��
    Dim strMachResult   As String   '������ �����
    Dim strResult       As String   '������ ���(����)
    Dim strIntResult    As String   '������ ���(����)
    Dim strQCResult     As String   '������ ���(QC)
    Dim strFlag         As String   '������ Abnormal Flag
    Dim strComm         As String   '������ Comment
    
    '������ ����
    Dim strCheck        As String   '�˻����üũ
    Dim strSeqno        As String   '�˻����
    Dim strOrderCode    As String   'ó���ڵ�
    Dim strTestName     As String   '�˻��ڵ�
    Dim strTestCode     As String   '�˻��ڵ�
    Dim strTestCodeSub  As String   '�˻��ڵ�SUB
    Dim intResPrecUse   As Integer  '�Ҽ�����ȯ����
    Dim intResPrec      As Integer  '�Ҽ����ڸ���
    Dim strResType      As String   '�Ҽ�����ȯ����
    Dim strLow          As String
    Dim strHigh         As String
    Dim strJudge        As String   '�������
    Dim strPrevRslt     As String   '�������
    
    Dim intRstRow       As String   '����������� ���� Row
    Dim intCnt          As Integer  '��� Frame ����
    Dim intCol          As Integer  '����÷� ����
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    Dim strTemp1        As String
    Dim strTemp2        As String
    
On Error GoTo ErrHandle

    ReDim Preserve strRData(UBound(strRecvData))
    
    For i = 1 To UBound(strRecvData)
        strRData(i) = strRecvData(i)
    Next
    
    With frmMain
        For intCnt = 1 To UBound(strRData)
            strRcvBuf = strRData(intCnt)

            Call SetSQLData("RCV", strRcvBuf, "")
            
            strType = Mid$(strRcvBuf, 2, 1)
            If strType = "|" Then
                strType = Mid$(strRcvBuf, 1, 1)
            End If
            
            Select Case strType
                Case "H"    '## Header
                Case "P"    '## Patient
                Case "O"
                    strBarno = Trim(mGetP(mGetP(strRcvBuf, 4, "|"), 3, "^"))
                
                    With mResult
                        .BarNo = strBarno
                        .RsltDate = Format(Now, "yyyy-mm-dd")
                        .RsltTime = Format(Now, "hh:mm:ss")
                        .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                    End With

                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                Case "R"
                    strTemp1 = mGetP(strRcvBuf, 3, "|")
                    strIntBase = mGetP(strTemp1, 4, "^")
                    
                    strTemp2 = mGetP(strRcvBuf, 4, "|")
                    'strFlag = mGetP(strRcvBuf, 7, "|")
                    
                    If InStr(strTemp2, "^") > 0 Then
                        '## ������� ����
                        strResult = mGetP(strTemp2, 2, "^")
                    Else
                        '## ������� ����
                        strResult = strTemp2
                    End If

                    If strIntBase <> "" And strResult <> "" Then
                        SQL = ""
                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFMLOW,REFMHIGH,REFFLOW,REFFHIGH,RESPRECUSE,RESPREC" & vbCrLf
                        SQL = SQL & "  FROM EQPMASTER" & vbCr
                        SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                        If gPatOrdCd <> "" Then
                            SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                        End If
                        
                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                            strSeqno = Trim(RS_L.Fields("SEQNO"))
                            strTestCode = Trim(RS_L.Fields("TESTCODE")) & ""
                            strTestName = Trim(RS_L.Fields("TESTNAME")) & ""
                            '-- ����ġ
                            If mPatient.SEX = "M" Then
                                strLow = Trim(RS_L.Fields("REFMLOW")) & ""
                                strHigh = Trim(RS_L.Fields("REFMHIGH")) & ""
                            ElseIf mPatient.SEX = "F" Then
                                strLow = Trim(RS_L.Fields("REFFLOW")) & ""
                                strHigh = Trim(RS_L.Fields("REFFHIGH")) & ""
                            Else
                                '-- ��������ġ�� �⺻���� �Ѵ�
                                strLow = Trim(RS_L.Fields("REFMLOW")) & ""
                                strHigh = Trim(RS_L.Fields("REFMHIGH")) & ""
                            End If
                            intResPrecUse = Trim(RS_L.Fields("RESPRECUSE")) & ""
                            intResPrec = Trim(RS_L.Fields("RESPREC")) & ""
    
                            '-- ���Row �߰�
                            intRstRow = .spdResult.DataRowCnt + 1
                            If .spdResult.MaxRows < intRstRow Then
                                .spdResult.MaxRows = intRstRow
                            End If
    
                            '-- �Ҽ��� ó��
                            strMachResult = strResult
                            If intResPrecUse = 1 Then
                                For i = 0 To intResPrec
                                    If i = 0 Then
                                        strResType = "#0"
                                    ElseIf i = 1 Then
                                        strResType = strResType & ".0"
                                    Else
                                        strResType = strResType & "0"
                                    End If
                                Next
                                strResult = Format(strResult, strResType)
                            End If
                        
                            '--- �������
                            strJudge = ""
                            If IsNumeric(strLow) = True And IsNumeric(strHigh) = True Then
                                If CCur(strResult) > CCur(strLow) And CCur(strResult) < CCur(strHigh) Then
                                    strJudge = ""
                                ElseIf CCur(strHigh) <= CCur(strResult) Then
                                    strJudge = "H"
                                ElseIf CCur(strLow) >= CCur(strResult) Then
                                    strJudge = "L"
                                End If
                            End If
        
                            '-- ������� ǥ��("���")
                            SetText .spdOrder, "���", gRow, colSTATE
    
                            '-- ����ȭ�� ����� ǥ��
                            For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                If strTestName = gArrEQPNm(intCol - colSTATE, 6) Then
                                    SetText .spdOrder, strResult, gRow, intCol
                                    
                                    strTestCodeSub = gArrEQP(intCol - colSTATE, 16)
                                    
                                    Exit For
                                End If
                            Next
    
                            '-- ��� List
                            SetText .spdResult, strCheck, intRstRow, colRCHECKBOX               'üũ
                            SetText .spdResult, strSeqno, intRstRow, colRSEQNO                  '����
                            SetText .spdResult, strOrderCode, intRstRow, colRORDERCD            'ó���ڵ�
                            SetText .spdResult, strTestCode, intRstRow, colRTESTCD              '�˻��ڵ�
                            SetText .spdResult, strTestCodeSub, intRstRow, colRSUBCD            '�˻��ڵ�SUB
                            SetText .spdResult, strTestName, intRstRow, colRTESTNM              '�˻��
                            SetText .spdResult, strIntBase, intRstRow, colRCHANNEL              '���ä��
                            SetText .spdResult, strMachResult, intRstRow, colRMACHRESULT        '�����
                            SetText .spdResult, strResult, intRstRow, colRLISRESULT             'LIS���
                            SetText .spdResult, strJudge, intRstRow, colRJUDGE                  '����
                            SetText .spdResult, strLow & "~" & strHigh, intRstRow, colRREF      '����ġ
                            
                            '-- ������� ��ȸ
                            strPrevRslt = GetPrevResult(mResult.BarNo, strIntBase, strTestCode)
                            SetText .spdResult, strPrevRslt, intRstRow, colRPREVRESULT          '�������
                            
                            '-- H/L ����ǥ��
                            If strJudge = "H" Then
                                .spdResult.Row = intRstRow
                                .spdResult.Col = colRLISRESULT
                                .spdResult.ForeColor = vbRed
                                .spdResult.FontBold = True
                            ElseIf strJudge = "L" Then
                                .spdResult.Row = intRstRow
                                .spdResult.Col = colRLISRESULT
                                .spdResult.ForeColor = vbBlue
                                .spdResult.FontBold = True
                            Else
                                .spdResult.Row = intRstRow
                                .spdResult.Col = colRLISRESULT
                                .spdResult.ForeColor = vbBlack
                                .spdResult.FontBold = False
                            End If
                            
                            '-- ���� ����
                            Call SetLocalDB(gRow, intRstRow, "1", "")
        
                            '-- ���Count
                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                SetText .spdOrder, "1", gRow, colRCNT
                            Else
                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                            End If
                            strState = "R"
                            
                        End If
    
                        .spdResult.RowHeight(-1) = 15
        
                    End If
                    
                Case "L"
            
                    '## DB�� �������
                    If gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow, spdOrder)
    
                        If Res = -1 Then
                            '-- ���� ����
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "�������", gRow, colSTATE
                        Else
                            '-- ���� ����
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "����Ϸ�", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX
    
                                  SQL = "Update PATRESULT Set " & vbCrLf
                            SQL = SQL & " sendflag = '2' " & vbCrLf
                            SQL = SQL & " Where equipno = '" & gHOSP.HOSPCD & "' " & vbCrLf
                            SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                            SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
    
                            If DBExec(AdoCn_Local, SQL) Then
                                '-- ����
                            End If
                        End If
                        strState = ""
                    End If
                
            End Select
        Next
    End With

Exit Sub

ErrHandle:
    strErrMsg = ""
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "SerialRcvData_INDIKO" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub TCPRcvData_KLITE()
    Dim RS_L            As ADODB.Recordset
    
    '��� ���� ����
    Dim strRcvBuf       As String   '������ Data
    Dim strType         As String   '������ Record Type
    Dim strBarno        As String   '������ ���ڵ��ȣ
    Dim strSeq          As String   '������ Sequence
    Dim strRackNo       As String   '������ Rack Or Disk No
    Dim strTubePos      As String   '������ Tube Position
    Dim strIntBase      As String   '������ ������ �˻��
    Dim strMachResult   As String   '������ �����
    Dim strResult       As String   '������ ���(����)
    Dim strIntResult    As String   '������ ���(����)
    Dim strQCResult     As String   '������ ���(QC)
    Dim strFlag         As String   '������ Abnormal Flag
    Dim strComm         As String   '������ Comment
    
    '������ ����
    Dim strCheck        As String   '�˻����üũ
    Dim strSeqno        As String   '�˻����
    Dim strOrderCode    As String   'ó���ڵ�
    Dim strTestName     As String   '�˻��ڵ�
    Dim strTestCode     As String   '�˻��ڵ�
    Dim strTestCodeSub  As String   '�˻��ڵ�SUB
    Dim intResPrecUse   As Integer  '�Ҽ�����ȯ����
    Dim intResPrec      As Integer  '�Ҽ����ڸ���
    Dim strResType      As String   '�Ҽ�����ȯ����
    Dim strLow          As String
    Dim strHigh         As String
    Dim strJudge        As String   '�������
    Dim strPrevRslt     As String   '�������
    
    Dim intRstRow       As String   '����������� ���� Row
    Dim intCnt          As Integer  '��� Frame ����
    Dim intCol          As Integer  '����÷� ����
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    '���� ����
    Dim strCrea         As String
    Dim streGFR         As String
    Dim strFunction     As String
    Dim strFunc         As String
    
    Dim strResultA      As String
    Dim strResultB      As String
    Dim strResultA_NTE  As String
    Dim strResultB_NTE  As String
    Dim strNGSP         As String
    
    Dim strHeader       As String
    Dim strHeaderType   As String
    
    
    Dim strSend         As String
    
On Error GoTo ErrHandle

    ReDim Preserve strRData(UBound(strRecvData))
    
    For i = 1 To UBound(strRecvData)
        strRData(i) = strRecvData(i)
    Next
    
    With frmMain
        For intCnt = 1 To UBound(strRData)
            strRcvBuf = strRData(intCnt)

            strType = mGetP(strRcvBuf, 1, "|")

            Select Case strType
                Case "MSH"
                    'MSH|^~\&|Medicong|KLITE-8-1012484|||20120530104611||ORU^R01|TR03-025|P|2.4||||||ASCII<CR>
                    'MSH|^~\&|Medicong|KLITE-8-1012484|||20190611090403||ORU^R01|TR14-009|P|2.4||||||ASCII
                    strHeader = mGetP(strRcvBuf, 10, "|")
                    strHeaderType = mGetP(strRcvBuf, 18, "|")
                    
                Case "PID"
                    'PID|03-025||12345678||UnKnowName||<CR>
                    strBarno = Trim(mGetP(strRcvBuf, 4, "|"))
                    If Trim(strBarno) <> Trim(strOldBarno) Then
                        strOldBarno = strBarno
                        '-- �������
                        With mResult
                            .BarNo = strBarno
                            .RsltDate = Format(Now, "yyyy-mm-dd")
                            .RsltTime = Format(Now, "hh:mm:ss")
                            .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                        End With
                    End If
                    
                    '-- ���ȯ������
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    If gRow <= 0 Then
                        Exit Sub
                    End If

                Case "OBR"
                    'OBR||12345678^R||Medicong^KLITE-8^LN||201205301046<CR>
                    
                    '-- �������̽� ����
                    strSend = ""
                    strSend = strSend & SB
                    strSend = strSend & "MSH|^~$&|||||||ACK^R01|1|P|2.4||||0||" & strHeaderType & "|||" & vbCr '"MSH|^~\&|Virtual SDB HL7Server^FB6590F3-E233-41A5-BB5F-CB17F5015295^GUID|Instr RnD DeptSDBIOSENSOR|||20180117093204+0900||ACK^R01^ACK|0B140FC8-ABE7-4955-BFCF-7882A9A25FC6|P|2.6" & vbCr
                    strSend = strSend & "MSA|AA|" & strHeader & "|message accepted|||0|" & vbCr
                    strSend = strSend & EB & vbCr

                    'If wSck.State = sckOpen Then
                        SetRawData "[Tx]" & strSend
                        wSck.SendData strSend
                    'End If
                Case "OBX"
                    'OBX|1|NM|Blood^K^LN|K|20.10|mmol/L^R^R|||||F<CR>
                    'OBX|2|NM|Blood^Na^LN|Na|20.11|mmol/L^R^R|||||F<CR>
                    'OBX|3|NM|Blood^Cl^LN|Cl|20.12|mmol/L^R^R|||||F<CR>

                    strIntBase = mGetP(mGetP(strRcvBuf, 4, "|"), 2, "^")
                    strResult = mGetP(strRcvBuf, 6, "|")
                    
                    '-- �˻縶���� ���� ��������
                    If strIntBase <> "" And strResult <> "" Then
                        SQL = ""
                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFMLOW,REFMHIGH,REFFLOW,REFFHIGH,RESPRECUSE,RESPREC" & vbCrLf
                        SQL = SQL & "  FROM EQPMASTER" & vbCr
                        SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                        If gPatOrdCd <> "" Then
                            SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                        End If
                        
                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                            strSeqno = Trim(RS_L.Fields("SEQNO"))
                            strTestCode = Trim(RS_L.Fields("TESTCODE")) & ""
                            strTestName = Trim(RS_L.Fields("TESTNAME")) & ""
                            '-- ����ġ
                            If mPatient.SEX = "M" Then
                                strLow = Trim(RS_L.Fields("REFMLOW")) & ""
                                strHigh = Trim(RS_L.Fields("REFMHIGH")) & ""
                            ElseIf mPatient.SEX = "F" Then
                                strLow = Trim(RS_L.Fields("REFFLOW")) & ""
                                strHigh = Trim(RS_L.Fields("REFFHIGH")) & ""
                            Else
                                '-- ��������ġ�� �⺻���� �Ѵ�
                                strLow = Trim(RS_L.Fields("REFMLOW")) & ""
                                strHigh = Trim(RS_L.Fields("REFMHIGH")) & ""
                            End If
                            intResPrecUse = Trim(RS_L.Fields("RESPRECUSE")) & ""
                            intResPrec = Trim(RS_L.Fields("RESPREC")) & ""
    
                            '-- ���Row �߰�
                            intRstRow = .spdResult.DataRowCnt + 1
                            If .spdResult.MaxRows < intRstRow Then
                                .spdResult.MaxRows = intRstRow
                            End If
    
                            '-- �Ҽ��� ó��
                            strMachResult = strResult
                            If intResPrecUse = 1 Then
                                For i = 0 To intResPrec
                                    If i = 0 Then
                                        strResType = "#0"
                                    ElseIf i = 1 Then
                                        strResType = strResType & ".0"
                                    Else
                                        strResType = strResType & "0"
                                    End If
                                Next
                                strResult = Format(strResult, strResType)
                            End If
                        
                            '--- �������
                            strJudge = ""
                            If IsNumeric(strLow) = True And IsNumeric(strHigh) = True Then
                                If CCur(strResult) > CCur(strLow) And CCur(strResult) < CCur(strHigh) Then
                                    strJudge = ""
                                ElseIf CCur(strHigh) <= CCur(strResult) Then
                                    strJudge = "H"
                                ElseIf CCur(strLow) >= CCur(strResult) Then
                                    strJudge = "L"
                                End If
                            End If
        
                            '-- ������� ǥ��("���")
                            SetText .spdOrder, "���", gRow, colSTATE
    
                            '-- ����ȭ�� ����� ǥ��
                            For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                If strTestName = gArrEQPNm(intCol - colSTATE, 6) Then
                                    SetText .spdOrder, strResult, gRow, intCol
                                    
                                    strTestCodeSub = gArrEQP(intCol - colSTATE, 16)
                                    
                                    Exit For
                                End If
                            Next
    
                            '-- ��� List
                            SetText .spdResult, strCheck, intRstRow, colRCHECKBOX               'üũ
                            SetText .spdResult, strSeqno, intRstRow, colRSEQNO                  '����
                            SetText .spdResult, strOrderCode, intRstRow, colRORDERCD            'ó���ڵ�
                            SetText .spdResult, strTestCode, intRstRow, colRTESTCD          '�˻��ڵ�
                            SetText .spdResult, strTestCodeSub, intRstRow, colRSUBCD        '�˻��ڵ�SUB
                            SetText .spdResult, strTestName, intRstRow, colRTESTNM              '�˻��
                            SetText .spdResult, strIntBase, intRstRow, colRCHANNEL              '���ä��
                            SetText .spdResult, strMachResult, intRstRow, colRMACHRESULT        '�����
                            SetText .spdResult, strResult, intRstRow, colRLISRESULT             'LIS���
                            SetText .spdResult, strJudge, intRstRow, colRJUDGE                  '����
                            SetText .spdResult, strLow & "~" & strHigh, intRstRow, colRREF      '����ġ
                            
                            '-- ������� ��ȸ
                            strPrevRslt = GetPrevResult(mResult.BarNo, strIntBase, strTestCode)
                            SetText .spdResult, strPrevRslt, intRstRow, colRPREVRESULT          '�������
                            
                            '-- H/L ����ǥ��
                            If strJudge = "H" Then
                                .spdResult.Row = intRstRow
                                .spdResult.Col = colRLISRESULT
                                .spdResult.ForeColor = vbRed
                                .spdResult.FontBold = True
                            ElseIf strJudge = "L" Then
                                .spdResult.Row = intRstRow
                                .spdResult.Col = colRLISRESULT
                                .spdResult.ForeColor = vbBlue
                                .spdResult.FontBold = True
                            Else
                                .spdResult.Row = intRstRow
                                .spdResult.Col = colRLISRESULT
                                .spdResult.ForeColor = vbBlack
                                .spdResult.FontBold = False
                            End If
                            
                            '-- ���� ����
                            Call SetLocalDB(gRow, intRstRow, "1", "")
        
                            '-- ���Count
                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                SetText .spdOrder, "1", gRow, colRCNT
                            Else
                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                            End If
                            strState = "R"
                            
                        End If
    
                        .spdResult.RowHeight(-1) = 15
        
                    End If

                    .spdResult.RowHeight(-1) = 15

            End Select
        Next
    
        '## DB�� �������
        If gHOSP.SAVEAUTO = "Y" And strState = "R" Then
            Res = SaveTransData(gRow, spdOrder)

            If Res = -1 Then
                '-- ���� ����
                SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                SetText .spdOrder, "�������", gRow, colSTATE
            Else
                '-- ���� ����
                SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                SetText .spdOrder, "����Ϸ�", gRow, colSTATE
                SetText .spdOrder, "0", gRow, colCHECKBOX

                      SQL = "Update PATRESULT Set " & vbCrLf
                SQL = SQL & " sendflag = '2' " & vbCrLf
                SQL = SQL & " Where equipno = '" & gHOSP.HOSPCD & "' " & vbCrLf
                SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                If DBExec(AdoCn_Local, SQL) Then
                    '-- ����
                End If
            End If
            strState = ""
        End If
    End With

Exit Sub

ErrHandle:
    strErrMsg = ""
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "TCPRcvData_F200" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub


Private Sub TCPRcvData_F200()
    Dim RS_L            As ADODB.Recordset
    
    '��� ���� ����
    Dim strRcvBuf       As String   '������ Data
    Dim strType         As String   '������ Record Type
    Dim strBarno        As String   '������ ���ڵ��ȣ
    Dim strSeq          As String   '������ Sequence
    Dim strRackNo       As String   '������ Rack Or Disk No
    Dim strTubePos      As String   '������ Tube Position
    Dim strIntBase      As String   '������ ������ �˻��
    Dim strMachResult   As String   '������ �����
    Dim strResult       As String   '������ ���(����)
    Dim strIntResult    As String   '������ ���(����)
    Dim strQCResult     As String   '������ ���(QC)
    Dim strFlag         As String   '������ Abnormal Flag
    Dim strComm         As String   '������ Comment
    
    '������ ����
    Dim strCheck        As String   '�˻����üũ
    Dim strSeqno        As String   '�˻����
    Dim strOrderCode    As String   'ó���ڵ�
    Dim strTestName     As String   '�˻��ڵ�
    Dim strTestCode     As String   '�˻��ڵ�
    Dim strTestCodeSub  As String   '�˻��ڵ�SUB
    Dim intResPrecUse   As Integer  '�Ҽ�����ȯ����
    Dim intResPrec      As Integer  '�Ҽ����ڸ���
    Dim strResType      As String   '�Ҽ�����ȯ����
    Dim strLow          As String
    Dim strHigh         As String
    Dim strJudge        As String   '�������
    Dim strPrevRslt     As String   '�������
    
    Dim intRstRow       As String   '����������� ���� Row
    Dim intCnt          As Integer  '��� Frame ����
    Dim intCol          As Integer  '����÷� ����
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    '���� ����
    Dim strCrea         As String
    Dim streGFR         As String
    Dim strFunction     As String
    Dim strFunc         As String
    
    Dim strResultA      As String
    Dim strResultB      As String
    Dim strResultA_NTE  As String
    Dim strResultB_NTE  As String
    Dim strNGSP         As String
    
    
    Dim strSend         As String
    
On Error GoTo ErrHandle

    ReDim Preserve strRData(UBound(strRecvData))
    
    For i = 1 To UBound(strRecvData)
        strRData(i) = strRecvData(i)
    Next
    
    With frmMain
        For intCnt = 1 To UBound(strRData)
            strRcvBuf = strRData(intCnt)

            strType = mGetP(strRcvBuf, 1, "|")

            Select Case strType
                Case "MSH"
                    'MSH|^~\&|Medicong|KLITE-8-1012484|||20120530104611||ORU^R01|TR03-025|P|2.4||||||ASCII<CR>
                Case "PID"
                    'PID|03-025||12345678||UnKnowName||<CR>
                    strBarno = Trim(mGetP(strRcvBuf, 4, "|"))
                    If Trim(strBarno) <> Trim(strOldBarno) Then
                        strOldBarno = strBarno
                        '-- �������
                        With mResult
                            .BarNo = strBarno
                            .RsltDate = Format(Now, "yyyy-mm-dd")
                            .RsltTime = Format(Now, "hh:mm:ss")
                            .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                        End With
                    End If
                    
                    '-- ���ȯ������
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    If gRow <= 0 Then
                        Exit Sub
                    End If

                Case "OBR"
                    'OBR||12345678^R||Medicong^KLITE-8^LN||201205301046<CR>
                    
                    '-- �������̽� ����
                    strSend = ""
                    strSend = strSend & SB
                    strSend = strSend & "MSH|^~$&|||||||ACK^R01|1|P|2.4||||0||ASCII|||" & vbCr '"MSH|^~\&|Virtual SDB HL7Server^FB6590F3-E233-41A5-BB5F-CB17F5015295^GUID|Instr RnD DeptSDBIOSENSOR|||20180117093204+0900||ACK^R01^ACK|0B140FC8-ABE7-4955-BFCF-7882A9A25FC6|P|2.6" & vbCr
                    strSend = strSend & "MSA|AA|TR03-025|message accepted|||0|" & vbCr
                    strSend = strSend & EB & vbCr

                    If wSck.State = sckOpen Then
                        SetRawData "[Tx]" & strSend
                        wSck.SendData strSend
                    End If
                Case "OBX"
                    'OBX|1|NM|Blood^K^LN|K|20.10|mmol/L^R^R|||||F<CR>
                    'OBX|2|NM|Blood^Na^LN|Na|20.11|mmol/L^R^R|||||F<CR>
                    'OBX|3|NM|Blood^Cl^LN|Cl|20.12|mmol/L^R^R|||||F<CR>

                    strIntBase = mGetP(mGetP(strRcvBuf, 4, "|"), 2, "^")
                    strResult = mGetP(strRcvBuf, 6, "|")
                    
                    '-- �˻縶���� ���� ��������
                    If strIntBase <> "" And strResult <> "" Then
                        SQL = ""
                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFMLOW,REFMHIGH,REFFLOW,REFFHIGH,RESPRECUSE,RESPREC" & vbCrLf
                        SQL = SQL & "  FROM EQPMASTER" & vbCr
                        SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                        If gPatOrdCd <> "" Then
                            SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                        End If
                        
                        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                            strSeqno = Trim(RS_L.Fields("SEQNO"))
                            strTestCode = Trim(RS_L.Fields("TESTCODE")) & ""
                            strTestName = Trim(RS_L.Fields("TESTNAME")) & ""
                            '-- ����ġ
                            If mPatient.SEX = "M" Then
                                strLow = Trim(RS_L.Fields("REFMLOW")) & ""
                                strHigh = Trim(RS_L.Fields("REFMHIGH")) & ""
                            ElseIf mPatient.SEX = "F" Then
                                strLow = Trim(RS_L.Fields("REFFLOW")) & ""
                                strHigh = Trim(RS_L.Fields("REFFHIGH")) & ""
                            Else
                                '-- ��������ġ�� �⺻���� �Ѵ�
                                strLow = Trim(RS_L.Fields("REFMLOW")) & ""
                                strHigh = Trim(RS_L.Fields("REFMHIGH")) & ""
                            End If
                            intResPrecUse = Trim(RS_L.Fields("RESPRECUSE")) & ""
                            intResPrec = Trim(RS_L.Fields("RESPREC")) & ""
    
                            '-- ���Row �߰�
                            intRstRow = .spdResult.DataRowCnt + 1
                            If .spdResult.MaxRows < intRstRow Then
                                .spdResult.MaxRows = intRstRow
                            End If
    
                            '-- �Ҽ��� ó��
                            strMachResult = strResult
                            If intResPrecUse = 1 Then
                                For i = 0 To intResPrec
                                    If i = 0 Then
                                        strResType = "#0"
                                    ElseIf i = 1 Then
                                        strResType = strResType & ".0"
                                    Else
                                        strResType = strResType & "0"
                                    End If
                                Next
                                strResult = Format(strResult, strResType)
                            End If
                        
                            '--- �������
                            strJudge = ""
                            If IsNumeric(strLow) = True And IsNumeric(strHigh) = True Then
                                If CCur(strResult) > CCur(strLow) And CCur(strResult) < CCur(strHigh) Then
                                    strJudge = ""
                                ElseIf CCur(strHigh) <= CCur(strResult) Then
                                    strJudge = "H"
                                ElseIf CCur(strLow) >= CCur(strResult) Then
                                    strJudge = "L"
                                End If
                            End If
        
                            '-- ������� ǥ��("���")
                            SetText .spdOrder, "���", gRow, colSTATE
    
                            '-- ����ȭ�� ����� ǥ��
                            For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                If strTestName = gArrEQPNm(intCol - colSTATE, 6) Then
                                    SetText .spdOrder, strResult, gRow, intCol
                                    
                                    strOrderCode = gArrEQP(intCol - colSTATE, 16)
                                    
                                    Exit For
                                End If
                            Next
    
                            '-- ��� List
                            SetText .spdResult, strCheck, intRstRow, colRCHECKBOX               'üũ
                            SetText .spdResult, strSeqno, intRstRow, colRSEQNO                  '����
                            SetText .spdResult, strOrderCode, intRstRow, colRORDERCD            'ó���ڵ�
                            SetText .spdResult, strTestCode, intRstRow, colRTESTCD          '�˻��ڵ�
                            SetText .spdResult, strTestCodeSub, intRstRow, colRSUBCD        '�˻��ڵ�SUB
                            SetText .spdResult, strTestName, intRstRow, colRTESTNM              '�˻��
                            SetText .spdResult, strIntBase, intRstRow, colRCHANNEL              '���ä��
                            SetText .spdResult, strMachResult, intRstRow, colRMACHRESULT        '�����
                            SetText .spdResult, strResult, intRstRow, colRLISRESULT             'LIS���
                            SetText .spdResult, strJudge, intRstRow, colRJUDGE                  '����
                            SetText .spdResult, strLow & "~" & strHigh, intRstRow, colRREF      '����ġ
                            
                            '-- ������� ��ȸ
                            strPrevRslt = GetPrevResult(mResult.BarNo, strIntBase, strTestCode)
                            SetText .spdResult, strPrevRslt, intRstRow, colRPREVRESULT          '�������
                            
                            '-- H/L ����ǥ��
                            If strJudge = "H" Then
                                .spdResult.Row = intRstRow
                                .spdResult.Col = colRLISRESULT
                                .spdResult.ForeColor = vbRed
                                .spdResult.FontBold = True
                            ElseIf strJudge = "L" Then
                                .spdResult.Row = intRstRow
                                .spdResult.Col = colRLISRESULT
                                .spdResult.ForeColor = vbBlue
                                .spdResult.FontBold = True
                            Else
                                .spdResult.Row = intRstRow
                                .spdResult.Col = colRLISRESULT
                                .spdResult.ForeColor = vbBlack
                                .spdResult.FontBold = False
                            End If
                            
                            '-- ���� ����
                            Call SetLocalDB(gRow, intRstRow, "1", "")
        
                            '-- ���Count
                            If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                SetText .spdOrder, "1", gRow, colRCNT
                            Else
                                SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                            End If
                            strState = "R"
                            
                        End If
    
                        .spdResult.RowHeight(-1) = 15
        
                    End If

                    .spdResult.RowHeight(-1) = 15

            End Select
        Next
    
        '## DB�� �������
        If gHOSP.SAVEAUTO = "Y" And strState = "R" Then
            Res = SaveTransData(gRow, spdOrder)

            If Res = -1 Then
                '-- ���� ����
                SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                SetText .spdOrder, "�������", gRow, colSTATE
            Else
                '-- ���� ����
                SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                SetText .spdOrder, "����Ϸ�", gRow, colSTATE
                SetText .spdOrder, "0", gRow, colCHECKBOX

                      SQL = "Update PATRESULT Set " & vbCrLf
                SQL = SQL & " sendflag = '2' " & vbCrLf
                SQL = SQL & " Where equipno = '" & gHOSP.HOSPCD & "' " & vbCrLf
                SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                If DBExec(AdoCn_Local, SQL) Then
                    '-- ����
                End If
            End If
            strState = ""
        End If
    End With

Exit Sub

ErrHandle:
    strErrMsg = ""
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "TCPRcvData_F200" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub SerialRcvData_HITACHI7180()
    Dim RS_L            As ADODB.Recordset
    
    '��� ���� ����
    Dim strRcvBuf       As String   '������ Data
    Dim strType         As String   '������ Record Type
    Dim strBarno        As String   '������ ���ڵ��ȣ
    Dim strSeq          As String   '������ Sequence
    Dim strRackNo       As String   '������ Rack Or Disk No
    Dim strTubePos      As String   '������ Tube Position
    Dim strIntBase      As String   '������ ������ �˻��
    Dim strMachResult   As String   '������ �����
    Dim strResult       As String   '������ ���(����)
    Dim strIntResult    As String   '������ ���(����)
    Dim strQCResult     As String   '������ ���(QC)
    Dim strFlag         As String   '������ Abnormal Flag
    Dim strComm         As String   '������ Comment
    
    '������ ����
    Dim strCheck        As String   '�˻����üũ
    Dim strSeqno        As String   '�˻����
    Dim strOrderCode    As String   'ó���ڵ�
    Dim strTestName     As String   '�˻��ڵ�
    Dim strTestCode     As String   '�˻��ڵ�
    Dim strTestCodeSub  As String   '�˻��ڵ�SUB
    Dim intResPrecUse   As Integer  '�Ҽ�����ȯ����
    Dim intResPrec      As Integer  '�Ҽ����ڸ���
    Dim strResType      As String   '�Ҽ�����ȯ����
    Dim strLow          As String
    Dim strHigh         As String
    Dim strJudge        As String   '�������
    Dim strPrevRslt     As String   '�������
    
    Dim intRstRow       As String   '����������� ���� Row
    Dim intCnt          As Integer  '��� Frame ����
    Dim intCol          As Integer  '����÷� ����
    Dim Res             As Integer

    Dim strTmp          As String
    Dim strQCTemp       As String
    Dim strRData()      As String

    Dim i               As Integer
    Dim ii              As Integer
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    
    '���� ����
    Dim strCrea         As String
    Dim streGFR         As String
    Dim strFunction     As String
    Dim strFunc         As String
    Dim sFunc           As String
    
On Error GoTo RST

    ReDim Preserve strRData(UBound(strRecvData))
    
    For i = 1 To UBound(strRecvData)
        strRData(i) = strRecvData(i)
    Next
    
    With frmMain
        For intCnt = 1 To UBound(strRData)
            strRcvBuf = strRData(intCnt)

            'Call SetSQLData("RCV", strRcvBuf, "")

            strType = Mid$(strRcvBuf, 1, 1)

            Select Case strType
                Case ">", "?", "@"      'ANY ����
                    
                    '-- ���� ����
                    Call SendData(SndMore)
                
                Case "1", "2", "3", "4", "5", "6", "7", "8", "9" '##Result
                    
                    '-- ���� ����
                    Call SendData(SndMore)
                    
                Case ";"    '## TS inquiry
'                    'gHOSP.BARLEN = 13
'                    ';A1     0  11   1019000157    00610191202
'                    strBarno = Trim(Mid(strRcvBuf, 14, 13))
'                    strFunction = Mid$(strRcvBuf, 4, 38)
'                    strFunc = Mid(strRcvBuf, 2, 12) & String(13, "#") & Mid(strRcvBuf, 27, 15)
'                    strSeq = Mid(strRcvBuf, 4, 5)
'                    strRackNo = Mid(strRcvBuf, 9, 1)
'                    strTubePos = Mid(strRcvBuf, 10, 3)
'
'                    '-- ��������
'                    With mOrder
'                        .Function = strFunction
'                        .Func = strFunc
'                        .BarNo = strBarno
'                        .RackNo = strRackNo
'                        .TubePos = strTubePos
'                        .Seq = strSeq
'                    End With
'
'                    '-- ����ȯ������
'                    Call GetOrder_HITACHI7180(Trim$(strBarno), gHOSP.RSTTYPE)

                    strBarno = Trim(Mid(strRcvBuf, 14, 13))
                    sFunc = Mid(strRcvBuf, 2, 12) & String(13, "#") & Mid(strRcvBuf, 27, 15)
                    sFunc = Mid(strRcvBuf, 2, 40)
                    With mOrder
                        .BarNo = strBarno
                        '.Func = Mid$(strRcvBuf, 2, 2)
                        .Func = sFunc
                        .Function = Mid$(strRcvBuf, 4, 38)
                        .Seq = Mid(strRcvBuf, 4, 5)
                        .RackNo = Mid$(strRcvBuf, 9, 1)
                        .TubePos = Mid$(strRcvBuf, 10, 3)
                        'tmpSeqNo = Mid(RcvBuffer, 4, 5)
                    End With
                    
                    Call GetOrder_HITACHI7180(Trim$(strBarno), gHOSP.RSTTYPE)


                Case ":"    '## End
                    '## Control, Calibration �����ʹ� ������
                    strFunc = Mid$(strRcvBuf, 2, 1)
                    'If UCase(strFunc) = "H" Or UCase(strFunc) = "G" Or UCase(strFunc) = "F" Then
                    If UCase(strFunc) = "H" Or UCase(strFunc) = "G" Then
                        '-- ���� ����
                        Call SendData(SndMore)
                        strState = ""
                        Exit Sub
                    End If
                    
                    If UCase(strFunc) = "F" Then
                        '-- ���� ����
                        Call SendData(SndMore)
                        strState = ""
                        Exit Sub
                    End If
                    
                    strSeq = Mid(strRcvBuf, 4, 5)
                    strRackNo = Mid(strRcvBuf, 9, 1)
                    strTubePos = Mid(strRcvBuf, 10, 3)
                    strBarno = Trim$(Mid$(strRcvBuf, 14, gHOSP.BARLEN)) '13
                    
                    '-- �������
                    With mResult
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .Seq = strSeq
                        .RsltDate = Format(Now, "yyyy-mm-dd")
                        .RsltTime = Format(Now, "hh:mm:ss")
                        .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                    End With
                    
                    '-- ���ȯ������
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    If gRow <= 0 Then
                        Call SendData(SndMore)
                        Exit Sub
                    End If

'                    If mResult.BarNo = "" Then
'                        Call SendData(SndMore)
'                        Exit Sub
'                    End If

                    For ii = 51 To Len(strRcvBuf) Step 10
                        strIntBase = Trim(Mid(strRcvBuf, ii, 3))
                        strResult = Trim(Mid(strRcvBuf, ii + 3, 6))
                        
                        'CRP : 16
                        '2019.08.27 ���� : CRP�� �״�� �����
'                        If strIntBase = "16" Then
'                            If IsNumeric(strResult) Then
'                                If CCur(strResult) < 0.1 Then
'                                    strResult = "<0.1"
'                                End If
'                            End If
'                        End If
                    
                        strSeqno = ""
                        strTestCode = ""
                        strTestName = ""
                        intResPrecUse = -1
                        intResPrec = -1
                            
                        '-- �˻縶���� ���� ��������
                        If strIntBase <> "" And strResult <> "" Then
                            SQL = ""
                            SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFMLOW,REFMHIGH,REFFLOW,REFFHIGH,RESPRECUSE,RESPREC" & vbCrLf
                            SQL = SQL & "  FROM EQPMASTER                                                                   " & vbCrLf
                            SQL = SQL & " WHERE EQUIPCD     = '" & gHOSP.MACHCD & "'                                        " & vbCrLf
                            SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "'                                          " & vbCrLf
                            If gPatOrdCd <> "" Then
                                SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                            End If
                            
                            Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                            If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                strSeqno = Trim(RS_L.Fields("SEQNO"))
                                strTestCode = Trim(RS_L.Fields("TESTCODE")) & ""
                                strTestName = Trim(RS_L.Fields("TESTNAME")) & ""
                                '-- ����ġ
                                If mPatient.SEX = "M" Then
                                    strLow = Trim(RS_L.Fields("REFMLOW")) & ""
                                    strHigh = Trim(RS_L.Fields("REFMHIGH")) & ""
                                ElseIf mPatient.SEX = "F" Then
                                    strLow = Trim(RS_L.Fields("REFFLOW")) & ""
                                    strHigh = Trim(RS_L.Fields("REFFHIGH")) & ""
                                Else
                                    '-- ��������ġ�� �⺻���� �Ѵ�
                                    strLow = Trim(RS_L.Fields("REFMLOW")) & ""
                                    strHigh = Trim(RS_L.Fields("REFMHIGH")) & ""
                                End If
                                
                                '-- �Ҽ�����ȯ ��뿩�ο� ��ȯ�ڸ���
                                intResPrecUse = Trim(RS_L.Fields("RESPRECUSE")) & ""
                                intResPrec = Trim(RS_L.Fields("RESPREC")) & ""
                                
                                '-- �Ҽ��� ó��
                                strMachResult = strResult
                                If intResPrecUse = 1 Then
                                    For i = 0 To intResPrec
                                        If i = 0 Then
                                            strResType = "#0"
                                        ElseIf i = 1 Then
                                            strResType = strResType & ".0"
                                        Else
                                            strResType = strResType & "0"
                                        End If
                                    Next
                                    strResult = Format(strResult, strResType)
                                End If
                            
                                '--- �������
                                strJudge = ""
                                If IsNumeric(strLow) = True And IsNumeric(strHigh) = True Then
                                    If CCur(strResult) > CCur(strLow) And CCur(strResult) < CCur(strHigh) Then
                                        strJudge = ""
                                    ElseIf CCur(strHigh) <= CCur(strResult) Then
                                        strJudge = "H"
                                    ElseIf CCur(strLow) >= CCur(strResult) Then
                                        strJudge = "L"
                                    End If
                                End If
            
                                '-- ���Row �߰�
                                intRstRow = .spdResult.DataRowCnt + 1
                                If .spdResult.MaxRows < intRstRow Then
                                    .spdResult.MaxRows = intRstRow
                                End If
            
                                '-- ������� ǥ��("���")
                                SetText .spdOrder, "���", gRow, colSTATE
        
                                '-- ����ȭ�� ����� ǥ��
                                For intCol = colSTATE + 1 To .spdOrder.MaxCols
                                    If strTestName = gArrEQPNm(intCol - colSTATE, 6) Then
                                        SetText .spdOrder, strResult, gRow, intCol
                                        
                                        strTestCodeSub = gArrEQP(intCol - colSTATE, 16)
                                        
                                        Exit For
                                    End If
                                Next
        
                                '-- ��� List
                                SetText .spdResult, strCheck, intRstRow, colRCHECKBOX               'üũ
                                SetText .spdResult, strSeqno, intRstRow, colRSEQNO                  '����
                                SetText .spdResult, strOrderCode, intRstRow, colRORDERCD            'ó���ڵ�
                                SetText .spdResult, strTestCode, intRstRow, colRTESTCD          '�˻��ڵ�
                                SetText .spdResult, strTestCodeSub, intRstRow, colRSUBCD        '�˻��ڵ�SUB
                                SetText .spdResult, strTestName, intRstRow, colRTESTNM              '�˻��
                                SetText .spdResult, strIntBase, intRstRow, colRCHANNEL              '���ä��
                                SetText .spdResult, strMachResult, intRstRow, colRMACHRESULT        '�����
                                SetText .spdResult, strResult, intRstRow, colRLISRESULT             'LIS���
                                SetText .spdResult, strJudge, intRstRow, colRJUDGE                  '����
                                SetText .spdResult, strLow & "~" & strHigh, intRstRow, colRREF      '����ġ
                                
                                '-- ������� ��ȸ
                                strPrevRslt = GetPrevResult(mResult.BarNo, strIntBase, strTestCode)
                                SetText .spdResult, strPrevRslt, intRstRow, colRPREVRESULT          '�������
                                
                                '-- H/L ����ǥ��
                                If strJudge = "H" Then
                                    .spdResult.Row = intRstRow
                                    .spdResult.Col = colRLISRESULT
                                    .spdResult.ForeColor = vbRed
                                    .spdResult.FontBold = True
                                ElseIf strJudge = "L" Then
                                    .spdResult.Row = intRstRow
                                    .spdResult.Col = colRLISRESULT
                                    .spdResult.ForeColor = vbBlue
                                    .spdResult.FontBold = True
                                Else
                                    .spdResult.Row = intRstRow
                                    .spdResult.Col = colRLISRESULT
                                    .spdResult.ForeColor = vbBlack
                                    .spdResult.FontBold = False
                                End If
                                
                                '-- ���Count
                                If GetText(.spdOrder, gRow, colRCNT) = "" Then
                                    SetText .spdOrder, "1", gRow, colRCNT
                                Else
                                    SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                                End If
                                
                                '-- ���� ����
                                Call SetLocalDB(gRow, intRstRow, "1", "")
                                
                                strState = "R"
                                
                            End If
                        End If
                    Next
                    
                    Call SendData(SndMore)
                    
                    .spdResult.RowHeight(-1) = 15

                    '## DB�� �������
                    If gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow, spdOrder)

                        If Res = -1 Then
                            '-- ���� ����
                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "�������", gRow, colSTATE
                        Else
                            '-- ���� ����
                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "����Ϸ�", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX

                                  SQL = "Update PATRESULT Set " & vbCrLf
                            SQL = SQL & " sendflag = '2' " & vbCrLf
                            SQL = SQL & " Where equipno = '" & gHOSP.HOSPCD & "' " & vbCrLf
                            SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                            SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf

                            If DBExec(AdoCn_Local, SQL) Then
                                '-- ����
                            End If
                        End If
                        strState = ""
                    End If
            End Select
        Next
    End With

Exit Sub

RST:
    strErrMsg = ""
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "SerialRcvData_H7180" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub


Private Sub Phase_Serial_AU680()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case intPhase
            Case 1
                Select Case BufChar
                    Case STX
                        Erase strRecvData
                        intBufCnt = 1
                        ReDim Preserve strRecvData(intBufCnt)
                        intPhase = 2
                End Select
            Case 2
                Select Case BufChar
                    Case STX
                        Erase strRecvData
                        intBufCnt = 1
                        ReDim Preserve strRecvData(intBufCnt)
                    Case ETB
                    Case ETX
                        intPhase = 1
                        lblTestDate.Caption = Format(Now, "yyyy-mm-dd")
                        Call SerialRcvData_AU680
                    Case Else
                        If intBufCnt > 0 Then
                            strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                        End If
                End Select
        End Select
    Next i

End Sub



'-----------------------------------------------------------------------------'
'   ��� : �������� ����
'-----------------------------------------------------------------------------'
Private Sub SendOrder_XN1000()
    Dim strOutput   As String     '�۽��� ������

    Select Case intSndPhase
        Case 1  '## Header
            strOutput = intFrameNo & "H|\^&||||||||||P|1" & vbCr & ETX
            intSndPhase = 2
            intFrameNo = intFrameNo + 1

        Case 2  '## Patient
            strOutput = intFrameNo & "P|1" & vbCr & ETX
            intSndPhase = 3
            intFrameNo = intFrameNo + 1
        
        Case 3  '## Order
            If mOrder.NoOrder = True Then
                strOutput = intFrameNo & "O|1|" & mOrder.RackNo & "^" & mOrder.TubePos & "^" & Right(Space(15) & mOrder.BarNo, 15) & "^B||" & mOrder.Order & "|||||||N||||||||||||||Q" & vbCr & ETX
                intSndPhase = 4
            Else
                '## ���� ������
                If mOrder.IsSending = False Then
                    strOutput = "O|1|" & mOrder.RackNo & "^" & mOrder.TubePos & "^" & Right(Space(15) & mOrder.BarNo, 15) & "^B||" & mOrder.Order & "|||||||N||||||||||||||Q"
                    
                    If Len(strOutput) > 230 Then
                        mOrder.IsSending = True
                        mOrder.Order = Mid$(strOutput, 231)
                        strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
                        intSndPhase = 3
                    Else
                        strOutput = intFrameNo & strOutput & vbCr & ETX
                        intSndPhase = 4
                    End If
                '## ���� ���ڿ��� ������
                Else
                    strOutput = mOrder.Order
                    If Len(strOutput) > 230 Then
                        mOrder.Order = Mid$(strOutput, 231)
                        strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
                        intSndPhase = 3
                    Else
                        mOrder.IsSending = False
                        strOutput = intFrameNo & strOutput & vbCr & ETX
                        intSndPhase = 4
                    End If
                End If
            End If
            intFrameNo = intFrameNo + 1

        Case 4  '## Termianator
            strOutput = intFrameNo & "L|1|N" & vbCr & ETX
            intSndPhase = 5
            intFrameNo = intFrameNo + 1

        Case 5  '## EOT
            strState = ""
            frmMain.comEqp.Output = EOT
            SetRawData "[Tx]" & EOT
            intFrameNo = 1

            Exit Sub
    End Select

    If intFrameNo = 8 Then
        intFrameNo = 0
    End If

    strOutput = STX & strOutput & GetChkSum(strOutput) & vbCrLf
    
    Call SendData(strOutput)

End Sub

'-----------------------------------------------------------------------------'
'   ��� : �������� ����
'-----------------------------------------------------------------------------'
Private Sub SendOrder_C8000()
    Dim strOutput   As String     '�۽��� ������

    Select Case intSndPhase
        Case 1  '## Header
            strOutput = intFrameNo & "H|\^&||||||||||P|1" & vbCr & ETX
            intSndPhase = 2
            intFrameNo = intFrameNo + 1

        Case 2  '## Patient
            strOutput = intFrameNo & "P|1" & vbCr & ETX
            intSndPhase = 3
            intFrameNo = intFrameNo + 1
        
        Case 3  '## Order
            If mOrder.NoOrder = True Then
                strOutput = intFrameNo & "O|1|" & mOrder.RackNo & "^" & mOrder.TubePos & "^" & Right(Space(15) & mOrder.BarNo, 15) & "^B||" & mOrder.Order & "|||||||N||||||||||||||Q" & vbCr & ETX
                intSndPhase = 4
            Else
                '## ���� ������
                If mOrder.IsSending = False Then
                    strOutput = "O|1|" & mOrder.RackNo & "^" & mOrder.TubePos & "^" & Right(Space(15) & mOrder.BarNo, 15) & "^B||" & mOrder.Order & "|||||||N||||||||||||||Q"
                    
                    If Len(strOutput) > 230 Then
                        mOrder.IsSending = True
                        mOrder.Order = Mid$(strOutput, 231)
                        strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
                        intSndPhase = 3
                    Else
                        strOutput = intFrameNo & strOutput & vbCr & ETX
                        intSndPhase = 4
                    End If
                '## ���� ���ڿ��� ������
                Else
                    strOutput = mOrder.Order
                    If Len(strOutput) > 230 Then
                        mOrder.Order = Mid$(strOutput, 231)
                        strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
                        intSndPhase = 3
                    Else
                        mOrder.IsSending = False
                        strOutput = intFrameNo & strOutput & vbCr & ETX
                        intSndPhase = 4
                    End If
                End If
            End If
            intFrameNo = intFrameNo + 1

        Case 4  '## Termianator
            strOutput = intFrameNo & "L|1|N" & vbCr & ETX
            intSndPhase = 5
            intFrameNo = intFrameNo + 1

        Case 5  '## EOT
            strState = ""
            frmMain.comEqp.Output = EOT
            SetRawData "[Tx]" & EOT
            intFrameNo = 1

            Exit Sub
    End Select

    If intFrameNo = 8 Then
        intFrameNo = 0
    End If

    strOutput = STX & strOutput & GetChkSum(strOutput) & vbCrLf
    
    Call SendData(strOutput)

End Sub


Private Sub Phase_Serial_INDIKO()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)

        Select Case intPhase
            Case 1      '## Estabilshment Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        intBufCnt = 1
                        ReDim Preserve strRecvData(intBufCnt)
                        intPhase = 2

                        Call SendData(ACK)
                        
                    Case ACK
                        If strState = "Q" Then
                            Call SendOrder_XN1000
                        Else
                            Call SendData(ACK)
                        End If
                End Select
            
            Case 2      '## Transfer Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        Call SendData(ACK)

                    Case STX
                        If intBufCnt = 0 Then
                            intBufCnt = 1
                            Erase strRecvData
                            ReDim Preserve strRecvData(intBufCnt)
                        Else
                            intBufCnt = intBufCnt + 1
                            ReDim Preserve strRecvData(intBufCnt)
                        End If
                        
                    Case ETB
                        blnIsETB = True
                        intPhase = 3
                    
                    Case ETX
                        intBufCnt = intBufCnt + 1
                        ReDim Preserve strRecvData(intBufCnt)
                        intPhase = 3
                    
                    Case vbCr
                    Case vbLf
                    
                    Case EOT
                        intPhase = 1
                    
                    Case Else
                        If blnIsETB = False Then
                            strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                        Else
                            blnIsETB = False
                        End If
                End Select
            
            Case 3      '## Transfer Phase
                Select Case BufChar
                    Case vbCr
                    Case vbLf
                        intPhase = 4
                        Call SendData(ACK)
                End Select
            Case 4      '## Termination Phase
                Select Case BufChar
                    Case STX
                        intPhase = 2
                    Case EOT
                        lblTestDate.Caption = Format(Now, "yyyy-mm-dd")
                        
                        Call SerialRcvData_INDIKO
                        
                        If strState = "Q" Then
                            intSndPhase = 1
                            intFrameNo = 1
                            Call SendData(ENQ)
                        End If
                        
                        intPhase = 1
                End Select
        End Select
    Next i

End Sub

Private Sub Phase_Serial_XN1000()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)

        Select Case intPhase
            Case 1      '## Estabilshment Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        intBufCnt = 1
                        ReDim Preserve strRecvData(intBufCnt)
                        intPhase = 2

                        Call SendData(ACK)
                        
                    Case ACK
                        If strState = "Q" Then
                            Call SendOrder_XN1000
                        Else
                            Call SendData(ACK)
                        End If
                End Select
            
            Case 2      '## Transfer Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        Call SendData(ACK)

                    Case STX
                        If intBufCnt = 0 Then
                            intBufCnt = 1
                            Erase strRecvData
                            ReDim Preserve strRecvData(intBufCnt)
                        Else
                            intBufCnt = intBufCnt + 1
                            ReDim Preserve strRecvData(intBufCnt)
                        End If
                        
                    Case ETB
                        blnIsETB = True
                        intPhase = 3
                    
                    Case ETX
                        intBufCnt = intBufCnt + 1
                        ReDim Preserve strRecvData(intBufCnt)
                        intPhase = 3
                    
                    Case vbCr, vbLf
                    
                    Case EOT
                        intPhase = 1
                    
                    Case Else
                        If blnIsETB = False Then
                            strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                        Else
                            blnIsETB = False
                        End If
                End Select
            
            Case 3      '## Transfer Phase
                Select Case BufChar
                    Case vbCr
                    Case vbLf
                        intPhase = 4
                        Call SendData(ACK)
                End Select
            Case 4      '## Termination Phase
                Select Case BufChar
                    Case STX
                        intPhase = 2
                    Case EOT
                        lblTestDate.Caption = Format(Now, "yyyy-mm-dd")
                        
                        Call SerialRcvData_XN1000
                        
                        If strState = "Q" Then
                            intSndPhase = 1
                            intFrameNo = 1
                            Call SendData(ENQ)
                        End If
                        
                        intPhase = 1
                End Select
        End Select
    Next i

End Sub

Private Sub Phase_Serial_C8000()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)

        Select Case intPhase
            Case 1      '## Estabilshment Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        intBufCnt = 1
                        ReDim Preserve strRecvData(intBufCnt)
                        intPhase = 2

                        Call SendData(ACK)
                        
                    Case ACK
                        If strState = "Q" Then
                            Call SendOrder_C8000
                        Else
                            Call SendData(ACK)
                        End If
                End Select
            
            Case 2      '## Transfer Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        Call SendData(ACK)

                    Case STX
                        If intBufCnt = 0 Then
                            intBufCnt = 1
                            Erase strRecvData
                            ReDim Preserve strRecvData(intBufCnt)
                        Else
                            intBufCnt = intBufCnt + 1
                            ReDim Preserve strRecvData(intBufCnt)
                        End If
                        
                    Case ETB
                        blnIsETB = True
                        intPhase = 3
                    
                    Case ETX
                        intPhase = 3
                    
                    Case vbCr
                        intBufCnt = intBufCnt + 1
                        ReDim Preserve strRecvData(intBufCnt)
                    
                    Case vbLf
                    
                    Case EOT
                        intPhase = 1
                    
                    Case Else
                        If blnIsETB = False Then
                            strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                        Else
                            blnIsETB = False
                        End If
                End Select
            
            Case 3      '## Transfer Phase
                Select Case BufChar
                    Case vbCr
                    Case vbLf
                        intPhase = IIf(blnIsETB = False, 4, 2)
                        'intPhase = 4
                        Call SendData(ACK)
                End Select
            Case 4      '## Termination Phase
                Select Case BufChar
                    Case STX
                        intPhase = 2
                    Case EOT
                        lblTestDate.Caption = Format(Now, "yyyy-mm-dd")
                        intPhase = 1
                        
                        Call SerialRcvData_C8000
                        
                        If strState = "Q" Then
                            intSndPhase = 1
                            intFrameNo = 1
                            Call SendData(ENQ)
                        End If
                        
                        
                End Select
        End Select
    Next i

End Sub

Private Sub Phase_Serial_ISMART30()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)

        Select Case intPhase
            Case 1      '## Estabilshment Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        intBufCnt = 1
                        ReDim Preserve strRecvData(intBufCnt)
                        intPhase = 2

                        Call SendData(ACK)
                        
                    Case ACK
'                        If strState = "Q" Then
'                            Call SendOrder_XN1000
'                        Else
'                            Call SendData(ACK)
'                        End If
                End Select
            
            Case 2      '## Transfer Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        Call SendData(ACK)

                    Case STX
                        If intBufCnt = 0 Then
                            intBufCnt = 1
                            Erase strRecvData
                            ReDim Preserve strRecvData(intBufCnt)
                        Else
                            intBufCnt = intBufCnt + 1
                            ReDim Preserve strRecvData(intBufCnt)
                        End If
                        
                    Case ETB
                        blnIsETB = True
                        intPhase = 3
                    
                    Case ETX
                        intBufCnt = intBufCnt + 1
                        ReDim Preserve strRecvData(intBufCnt)
                        intPhase = 3
                    
                    Case vbCr, vbLf
                    
                    Case EOT
                        intPhase = 1
                    
                    Case Else
                        If blnIsETB = False Then
                            strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                        Else
                            blnIsETB = False
                        End If
                End Select
            
            Case 3      '## Transfer Phase
                Select Case BufChar
                    Case vbCr
                    Case vbLf
                        intPhase = 4
                        Call SendData(ACK)
                End Select
            Case 4      '## Termination Phase
                Select Case BufChar
                    Case STX
                        intPhase = 2
                    Case EOT
                        lblTestDate.Caption = Format(Now, "yyyy-mm-dd")
                        
                        Call SerialRcvData_ISMART30
                        
                        If strState = "Q" Then
                            intSndPhase = 1
                            intFrameNo = 1
                            Call SendData(ENQ)
                        End If
                        
                        intPhase = 1
                End Select
        End Select
    Next i

End Sub


Private Sub Phase_Serial_TEST1()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case intPhase
            Case 1
                Select Case BufChar
                    Case STX
                        Erase strRecvData
                        intBufCnt = 1
                        ReDim Preserve strRecvData(intBufCnt)
                        intPhase = 2
                End Select
            Case 2
                Select Case BufChar
                    Case ETX
                        intPhase = 1
                        lblTestDate.Caption = Format(Now, "yyyy-mm-dd")
                        Call SerialRcvData_TEST1
                    Case Else
                        If intBufCnt > 0 Then
                            strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                        End If
                End Select
        End Select
    Next i

End Sub

Private Sub Phase_Serial_HITACHI7180()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case intPhase
            Case 1
                Select Case BufChar
                    Case STX
                        Erase strRecvData
                        intBufCnt = 1
                        ReDim Preserve strRecvData(intBufCnt)
                        intPhase = 2
                End Select
            Case 2
                Select Case BufChar
                    Case STX
                        Erase strRecvData
                        intBufCnt = 1
                        ReDim Preserve strRecvData(intBufCnt)
                    Case ETX
                        intPhase = 1
                        lblTestDate.Caption = Format(Now, "yyyy-mm-dd")
                        
                        Call SerialRcvData_HITACHI7180
                        
                    Case vbCr
                    Case vbLf
                    Case Else
                        If intBufCnt > 0 Then
                            strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                        End If
                End Select
        End Select
    Next i

End Sub

Private Sub Phase_TCP_KLITE()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case intPhase
            Case 1
                Select Case BufChar
                    Case SB
                        Erase strRecvData
                        intBufCnt = 1
                        ReDim Preserve strRecvData(intBufCnt)
                        intPhase = 2
                End Select
            Case 2
                Select Case BufChar
                    Case SB
                        Erase strRecvData
                        intBufCnt = 1
                        ReDim Preserve strRecvData(intBufCnt)
                    Case EB
                        intPhase = 1
                        lblTestDate.Caption = Format(Now, "yyyy-mm-dd")
                        Call TCPRcvData_KLITE
                    Case vbCr
                        intBufCnt = intBufCnt + 1
                        ReDim Preserve strRecvData(intBufCnt)
                    Case vbLf
                    Case Else
                        If intBufCnt > 0 Then
                            strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                        End If
                End Select
        End Select
    Next i


End Sub



Private Sub frmClear()
    
    shpPatInfo.Visible = False
    lblPatInfo.Caption = ""
    
    spdOrder.MaxRows = 0
    spdResult.MaxRows = 0
    spdWork.MaxRows = 0
    
    txtBarcode.Text = ""
    txtPatID.Text = ""
    txtPName.Text = ""
    txtSA.Text = ""

    dtpFrom.Value = Now
    dtpTo.Value = Now
        
End Sub


Private Sub Form_Load()
    Dim strTmp      As String
    Dim strSaveDt   As String
    Dim intCnt      As Integer
    
On Error GoTo ErrHandle
    
'    Me.Width = 20940
'    Me.Height = 12585


    'Me.Caption = gHOSP.MACHNM
    'Me.Caption = gHOSP.MACHNM & Space$(5) & "�¢¢¢¢�     [���� �������̽�]     �¢¢¢¢�"

    If Mid(gForm.MAXYN, 1, 1) = "Y" Then
        Me.WindowState = 2
    Else
        Me.WindowState = 0
        Me.TOP = gForm.TOP
        Me.LEFT = gForm.LEFT
        Me.WIDTH = gForm.WIDTH
        Me.HEIGHT = gForm.HEIGHT
    End If
    
    Me.Caption = gHOSP.HOSPNM & Space$(5) & gHOSP.MACHNM
    lblHospInfo(0).Caption = "�� " & gHOSP.PARTNM & " - " & gHOSP.MACHNM & " �������̽�"
    lblHospInfo(1).Caption = "�� " & gHOSP.PARTNM & " - " & gHOSP.MACHNM & " �������̽�"

    If gWORKPOS = "M" Then
        spdWork.Visible = True
        fraWorkInfo.Visible = True
        
        cmdView.Visible = True
        cmdHide.Visible = True
        
    Else
        spdWork.Visible = False
        fraWorkInfo.Visible = False
    
        cmdView.Visible = False
        cmdHide.Visible = False
    
    End If


    Call CtlInitializing

    Call frmClear
    
    '-- Menu Set
    Call SetMenu

    '-- �÷����̱⼳��
    Call SetColumnView(spdOrder)

    '-- �˻��ڵ�
    Call GetTestList

    Call GetTestListName

    '-- �˻�� ���̱�
    Call SetExamCode(spdOrder)

    '-- ��ſ���
    Call OpenCommunication

    pDel = False

    spdComStatus.MaxRows = 0
    spdComStatus.Font.Bold = True
    
    lblTestDate.Caption = Format(Now, "yyyy-mm-dd")
    txtTestID.Text = gHOSP.USERID
    txtTestNm.Text = gHOSP.USERNM
    lblPatInfo.Caption = ""

    imgNet1.ZOrder 0
    tmrDBConn.Interval = 500
    tmrDBConn.Enabled = True
    
    
    '-- ������� ����
    strTmp = Format$(DateAdd("d", -Val(gHOSP.SAVEDAY), Format$(Now, "YYYY-MM-DD")), "YYYY-MM-DD")

    SQL = "Select count(*) From PATRESULT Where EXAMDATE <= '" & strTmp & "'"
    Set AdoRs_Local = New ADODB.Recordset
    
    AdoRs_Local.CursorLocation = adUseClient
    AdoRs_Local.Open SQL, AdoCn_Local
    If AdoRs_Local.RecordCount > 0 Then AdoRs_Local.MoveFirst
    If Not AdoRs_Local.EOF Then intCnt = AdoRs_Local(0) & ""
    AdoRs_Local.Close:    Set AdoRs_Local = Nothing
    
    If intCnt > 0 Then
        If MsgBox(gHOSP.SAVEDAY + "���� ����Ÿ�� �����Ͻðڽ��ϱ�?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
            strSaveDt = Format$(DateAdd("d", -Val(gHOSP.SAVEDAY), Format(Now, "YYYY-MM-DD")), "YYYY-MM-DD")
            
            SQL = "DELETE From PATRESULT Where EXAMDATE <= '" & strTmp & "'"
            AdoCn_Local.Execute SQL
        End If
    End If
    
    'tmrConn.Interval = 60000
    'tmrConn.Enabled = True
    
    
    Exit Sub

ErrHandle:

    If Err.Number = "8002" Then
        If (MsgBox("��Ʈ ��ȣ�� �߸��Ǿ����ϴ�." & vbNewLine & vbNewLine & "   ��� �����Ͻðڽ��ϱ�?", vbYesNo + vbCritical, Me.Caption)) = vbYes Then
            lblComStatus.Caption = "COM" & comEqp.CommPort & " �������"
            
            imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
            imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
            imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
            
            Resume Next
        Else
            End
        End If
    Else
                
        strErrMsg = ""
        strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "Form_Load" & vbNewLine & vbNewLine
        strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
        strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
        frmErrMsg.txtErr = vbNewLine & strErrMsg
        frmErrMsg.Show
    
    End If

End Sub
 
'
Public Sub OpenCommunication()

    If gComm.COMTYPE = "1" Then

        comEqp.CommPort = gComm.COMPORT
        comEqp.RTSEnable = gComm.RTSEnable
        comEqp.DTREnable = gComm.DTREnable
        comEqp.Settings = gComm.SPEED & "," & gComm.Parity & "," & gComm.DATABIT & "," & gComm.STOPBIT

        If comEqp.PortOpen = False Then
            comEqp.PortOpen = True
        End If

        If comEqp.PortOpen Then
            lblComStatus.Caption = "COM" & comEqp.CommPort & " ���Ἲ��"
            
            imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
            imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
            imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
        Else
            lblComStatus.Caption = "COM" & comEqp.CommPort & " �������"
            
            imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
            imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
            imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
        End If
    ElseIf gComm.COMTYPE = "2" Then
        'lblComStatus.Left = imgPort.Left + 500
        'lblComStatus.Width = 6000
        If gComm.TCPTYPE = "SERVER" Then
            wSck.LocalPort = CInt(gComm.TCPPORT)
            wSck.Listen
                
            
            lblComStatus.Caption = "TCP " & gComm.TCPPORT & " ������.."

            imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
            imgSend.Visible = False
            imgReceive.Visible = False
            lblSend.Visible = False
            lblRcv.Visible = False

        Else
            wSck.Close
            wSck.Connect gComm.TCPIP, CInt(gComm.TCPPORT)
            
            lblComStatus.Caption = "TCP " & gComm.TCPIP & ":" & gComm.TCPPORT & " ������..."

            imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
            imgSend.Visible = False
            imgReceive.Visible = False
            lblSend.Visible = False
            lblRcv.Visible = False
        End If
    ElseIf gComm.COMTYPE = "" Then

    End If

End Sub


Private Sub Form_Resize()

    On Error Resume Next

    If Me.ScaleHeight = 0 Then Exit Sub

    'Me.TOP = 0

    If gWORKPOS = "M" Then
        spdWork.HEIGHT = Me.ScaleHeight - picHeader.HEIGHT - picBottom.HEIGHT - fraWorkInfo.HEIGHT - 300
        
        spdOrder.WIDTH = Me.ScaleWidth - spdWork.WIDTH - spdResult.WIDTH - 200
        spdOrder.HEIGHT = Me.ScaleHeight - picHeader.HEIGHT - picBottom.HEIGHT - 350
        spdResult.LEFT = spdOrder.LEFT + spdOrder.WIDTH + 50
        spdResult.HEIGHT = Me.ScaleHeight - picHeader.HEIGHT - picBottom.HEIGHT - fraPatInfo.HEIGHT - 300
    
        fraPatInfo.LEFT = spdOrder.LEFT + spdOrder.WIDTH + 50
        fraPatInfo.HEIGHT = Me.ScaleHeight - picHeader.HEIGHT - picBottom.HEIGHT - spdResult.HEIGHT - 300
    Else
        spdOrder.LEFT = 100
        spdOrder.WIDTH = Me.ScaleWidth - spdResult.WIDTH - 200
        spdOrder.HEIGHT = Me.ScaleHeight - picHeader.HEIGHT - picBottom.HEIGHT - 350
        
        spdResult.LEFT = spdOrder.LEFT + spdOrder.WIDTH + 50
        spdResult.HEIGHT = Me.ScaleHeight - picHeader.HEIGHT - picBottom.HEIGHT - fraPatInfo.HEIGHT - 300
        
        fraPatInfo.LEFT = spdOrder.LEFT + spdOrder.WIDTH + 50
        fraPatInfo.HEIGHT = Me.ScaleHeight - picHeader.HEIGHT - picBottom.HEIGHT - spdResult.HEIGHT - 300

    End If
    
    If Me.WindowState = 2 Then
        'gForm.MAXYN = True
        Call WritePrivateProfileString("FORM", "MAXYN", "Y", App.PATH & "\INI\" & gMACH & ".ini")
    Else
        'gForm.MAXYN = False
        gForm.TOP = Me.TOP
        gForm.LEFT = Me.LEFT
        gForm.WIDTH = Me.WIDTH
        gForm.HEIGHT = Me.HEIGHT
        
        Call WritePrivateProfileString("FORM", "MAXYN", "N", App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("FORM", "TOP", gForm.TOP, App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("FORM", "LEFT", gForm.LEFT, App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("FORM", "WIDTH", gForm.WIDTH, App.PATH & "\INI\" & gMACH & ".ini")
        Call WritePrivateProfileString("FORM", "HEIGHT", gForm.HEIGHT, App.PATH & "\INI\" & gMACH & ".ini")

    End If
    
End Sub

'�������̽� ȯ�ڼ��ý� ������ �˻��׸�/��������ֱ�
Private Function GetPatTRestResult(ByVal asRow As Integer) As Integer
    Dim strBarno    As String
    Dim intSeq      As String
    Dim strExamDate As String
    Dim intRow   As Integer

On Error GoTo ErrHandle

    GetPatTRestResult = -1
    intRow = 0

    intSeq = GetText(spdOrder, asRow, colSAVESEQ)
    strExamDate = GetText(spdOrder, asRow, colEXAMDATE)
    strBarno = GetText(spdOrder, asRow, colBARCODE)
    
    If intSeq = "" Then
        Exit Function
    End If

    SQL = ""
    SQL = SQL & "SELECT DISTINCT SEQNO, EXAMNAME, EXAMCODE, RESULT, PREVRESULT, REFJUDGE" & vbCr
    SQL = SQL & "  FROM PATRESULT " & vbCr
    SQL = SQL & " WHERE SAVESEQ = " & intSeq & vbCr
    SQL = SQL & "   AND EXAMDATE = '" & strExamDate & "'" & vbCr
    SQL = SQL & "   AND BARCODE = '" & strBarno & "'" & vbCr
    SQL = SQL & " ORDER BY SEQNO "

    '-- Record Count ������
    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        With frmMain.spdResult
            .MaxRows = 0
            .MaxRows = AdoRs_Local.RecordCount
            Do Until AdoRs_Local.EOF
                intRow = intRow + 1
                If AdoRs_Local.Fields("EXAMCODE").Value & "" = "" Then
                    Call SetText(frmMain.spdResult, "0", intRow, colCHECKBOX)
                Else
                    Call SetText(frmMain.spdResult, "1", intRow, colCHECKBOX)
                End If
                Call SetText(frmMain.spdResult, AdoRs_Local.Fields("EXAMCODE").Value & "", intRow, colRTESTCD)
                Call SetText(frmMain.spdResult, AdoRs_Local.Fields("EXAMNAME").Value & "", intRow, colRTESTNM)
                Call SetText(frmMain.spdResult, AdoRs_Local.Fields("RESULT").Value & "", intRow, colRLISRESULT)
                If AdoRs_Local.Fields("REFJUDGE").Value & "" = "H" Then
                    .ForeColor = vbRed
                    .FontBold = True
                ElseIf AdoRs_Local.Fields("REFJUDGE").Value & "" = "L" Then
                    .ForeColor = vbBlue
                    .FontBold = True
                Else
                    .ForeColor = vbBlack
                    .FontBold = False
                End If
                Call SetText(frmMain.spdResult, AdoRs_Local.Fields("PREVRESULT").Value & "", intRow, colRPREVRESULT)
                AdoRs_Local.MoveNext
            Loop
            .RowHeight(-1) = 15
        End With
        GetPatTRestResult = 1
    End If

    AdoRs_Local.Close

Exit Function

ErrHandle:
    GetPatTRestResult = -1

    strErrMsg = ""
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "GetPatTRestResult" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Function

Private Sub mnuBarcode_Click()
    
    mnuBarcode.Checked = True
    mnuSeqno.Checked = False
    mnuRackPos.Checked = False
    mnuCheckBox.Checked = False

    Call WritePrivateProfileString("HOSP", "BARUSE", "Y", App.PATH & "\INI\" & gMACH & ".ini")
    Call WritePrivateProfileString("HOSP", "RSTTYPE", "0", App.PATH & "\INI\" & gMACH & ".ini")

End Sub

Private Sub mnuCheckBox_Click()
    
    mnuBarcode.Checked = False
    mnuSeqno.Checked = False
    mnuRackPos.Checked = False
    mnuCheckBox.Checked = True

    Call WritePrivateProfileString("HOSP", "BARUSE", "N", App.PATH & "\INI\" & gMACH & ".ini")
    Call WritePrivateProfileString("HOSP", "RSTTYPE", "3", App.PATH & "\INI\" & gMACH & ".ini")

End Sub

Private Sub mnuComm_Click()
    
    frmConfig.Show

End Sub

Private Sub mnuComTest_Click()

End Sub

Private Sub mnuCommTest_Click()

    If picComm.Visible = True Then
        picComm.Visible = False
    Else
        picComm.Visible = True
    End If
    
End Sub

Private Sub mnuEqpResult_Click()

    mnuEqpResult.Checked = True
    mnuLisResult.Checked = False

    Call WritePrivateProfileString("HOSP", "SAVELIS", "N", App.PATH & "\INI\" & gMACH & ".ini")

End Sub

Private Sub mnuHelp01_Click()

    Call WinExec(App.PATH & "\TeamViewerQS.exe", 1)
    
End Sub

Private Sub mnuHosp_Click()

    frmHospInfo.Show vbModal

End Sub

Private Sub mnuLisResult_Click()

    mnuEqpResult.Checked = False
    mnuLisResult.Checked = True

    Call WritePrivateProfileString("HOSP", "SAVELIS", "Y", App.PATH & "\INI\" & gMACH & ".ini")

End Sub

Private Sub mnuOpt_Click()
    
    frmTestOptSet.Show vbModal
    
End Sub

Private Sub mnuRackPos_Click()
    
    mnuBarcode.Checked = False
    mnuSeqno.Checked = False
    mnuRackPos.Checked = True
    mnuCheckBox.Checked = False

    Call WritePrivateProfileString("HOSP", "BARUSE", "N", App.PATH & "\INI\" & gMACH & ".ini")
    Call WritePrivateProfileString("HOSP", "RSTTYPE", "2", App.PATH & "\INI\" & gMACH & ".ini")

End Sub

Private Sub mnuResult_Click()
    
    frmResult.Show vbModal
    
End Sub

Private Sub mnuSaveAuto_Click()

    mnuSaveAuto.Checked = True
    mnuSaveManual.Checked = False

    Call WritePrivateProfileString("HOSP", "SAVEAUTO", "Y", App.PATH & "\INI\" & gMACH & ".ini")

End Sub

Private Sub mnuSaveManual_Click()

    mnuSaveAuto.Checked = False
    mnuSaveManual.Checked = True

    Call WritePrivateProfileString("HOSP", "SAVEAUTO", "N", App.PATH & "\INI\" & gMACH & ".ini")

End Sub

Private Sub mnuSeqno_Click()
    
    mnuBarcode.Checked = False
    mnuSeqno.Checked = True
    mnuRackPos.Checked = False
    mnuCheckBox.Checked = False

    Call WritePrivateProfileString("HOSP", "BARUSE", "N", App.PATH & "\INI\" & gMACH & ".ini")
    Call WritePrivateProfileString("HOSP", "RSTTYPE", "1", App.PATH & "\INI\" & gMACH & ".ini")
    
End Sub

Private Sub mnuTest_Click()
    
    frmTestSet.Show vbModal
    
End Sub

Private Sub mnuView_Click()
    frmScreenSet.Show vbModal
End Sub

Private Sub mnuWork_Click()
    
    frmWorkList.Show vbModal

End Sub

Public Sub spdOrder_Click(ByVal Col As Long, ByVal Row As Long)
    Dim intCol      As Integer
    Dim i           As Integer
    Dim strPatInfo  As String
    
    '-- ����
'    If Row = 0 Then
'        '-- ���� �߰�
'        Exit Sub
'    End If
    
    
    If Row = 0 And Col = colCHECKBOX Then
        If GetText(spdOrder, 1, colCHECKBOX) = "1" Then
            For i = 1 To spdOrder.DataRowCnt
                Call SetText(spdOrder, "0", i, colCHECKBOX)
            Next
        Else
            For i = 1 To spdOrder.DataRowCnt
                Call SetText(spdOrder, "1", i, colCHECKBOX)
            Next
        End If
        Exit Sub
    End If
    
    If Row > 0 And Col = colCHECKBOX Then
        If GetText(spdOrder, Row, colCHECKBOX) = "1" Then
            Call SetText(spdOrder, "0", Row, colCHECKBOX)
        Else
            Call SetText(spdOrder, "1", Row, colCHECKBOX)
        End If
        Exit Sub
    End If
    
    If Row = 0 Then
        Exit Sub
    End If
    
    '-- ȯ������ǥ��
    shpPatInfo.Visible = True
    
    strPatInfo = ""
    strPatInfo = strPatInfo & "����    ��: " & GetText(spdOrder, Row, colPNAME)
    If GetText(spdOrder, Row, colPSEX) <> "" Then
        strPatInfo = strPatInfo & " [" & GetText(spdOrder, Row, colPSEX) & "/" & GetText(spdOrder, Row, colPAGE) & "] " & vbCrLf
    Else
        strPatInfo = strPatInfo & vbCrLf
    End If
    strPatInfo = strPatInfo & "�°�ü��ȣ: " & GetText(spdOrder, Row, colBARCODE) & vbCrLf
    strPatInfo = strPatInfo & "��ȯ�ڹ�ȣ: " & GetText(spdOrder, Row, colPID) & vbCrLf
    
    lblPatInfo.Caption = strPatInfo
    
    txtBarcode.Text = GetText(spdOrder, Row, colBARCODE)
    txtPatID.Text = GetText(spdOrder, Row, colPID)
    txtPName.Text = GetText(spdOrder, Row, colPNAME)
    txtSA.Text = GetText(spdOrder, Row, colPSEX) & "/" & GetText(spdOrder, Row, colPAGE)
    
    '-- ���ǥ��
    If GetPatTRestResult(Row) = -1 Then
        '������� ������� �˻�� �����ֱ�
        spdResult.MaxRows = 0
        With spdOrder
            For intCol = colSTATE + 1 To .MaxCols
                If GetText(spdOrder, Row, intCol) <> "" Then    '��
                    spdResult.MaxRows = spdResult.MaxRows + 1
                    Call SetText(spdResult, GetText(spdOrder, 0, intCol), spdResult.MaxRows, colRTESTNM)
                    spdResult.RowHeight(-1) = 15
                End If
            Next
        End With
    End If

End Sub

Private Sub spdOrder_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sRow        As Long
    Dim sCol        As Long
    Dim strNewBarNo As String
    Dim intRow      As Integer
    Dim strSeq      As String
    
    
    sRow = spdOrder.ActiveRow
    sCol = spdOrder.ActiveCol
    
    If sRow = 0 Then
        Exit Sub
    End If
    
    strNewBarNo = GetText(spdOrder, sRow, sCol)
    
    If KeyCode = vbKeyReturn Then
        If colBARCODE = sCol Then
            If GetSampleInfo(sRow, spdOrder) = -1 Then
                MsgBox "�Է��� ���ڵ忡�� ȯ�������� ã�� ���߽��ϴ�." & vbNewLine & " ���ڵ� ��ȣ�� Ȯ���ϼ���", vbOKOnly + vbCritical, Me.Caption
            Else
                '��������
                SQL = ""
                SQL = SQL & "UPDATE PATRESULT SET "
                SQL = SQL & "  BARCODE  = '" & Trim(GetText(spdOrder, sRow, colBARCODE)) & "'" & vbCrLf
                SQL = SQL & " ,PID      = '" & Trim(GetText(spdOrder, sRow, colPID)) & "'" & vbCrLf
                SQL = SQL & " ,CHARTNO  = '" & Trim(GetText(spdOrder, sRow, colCHARTNO)) & "'" & vbCrLf
                SQL = SQL & " ,SPECIMEN = '" & Trim(GetText(spdOrder, sRow, colSPECIMEN)) & "'" & vbCrLf
                SQL = SQL & " ,DEPT     = '" & Trim(GetText(spdOrder, sRow, colDEPT)) & "'" & vbCrLf
                SQL = SQL & " ,INOUT    = '" & Trim(GetText(spdOrder, sRow, colINOUT)) & "'" & vbCrLf
                SQL = SQL & " ,ERYN     = '" & Trim(GetText(spdOrder, sRow, colER)) & "'" & vbCrLf
                SQL = SQL & " ,RETESTYN = '" & Trim(GetText(spdOrder, sRow, colRT)) & "'" & vbCrLf
                SQL = SQL & " ,PNAME    = '" & Trim(GetText(spdOrder, sRow, colPNAME)) & "'" & vbCrLf
                SQL = SQL & " ,PSEX     = '" & Trim(GetText(spdOrder, sRow, colPSEX)) & "'" & vbCrLf
                SQL = SQL & " ,PAGE     = '" & Trim(GetText(spdOrder, sRow, colPAGE)) & "'" & vbCrLf
                SQL = SQL & " ,DISKNO   = '" & Trim(GetText(spdOrder, sRow, colRACKNO)) & "'" & vbCrLf
                SQL = SQL & " ,POSNO    = '" & Trim(GetText(spdOrder, sRow, colPOSNO)) & "'" & vbCrLf
                SQL = SQL & " WHERE EQUIPNO  = '" & gHOSP.MACHCD & "'" & vbCrLf
                SQL = SQL & "   AND EXAMDATE = '" & Trim(GetText(spdOrder, sRow, colEXAMDATE)) & "'" & vbCrLf
                SQL = SQL & "   AND EXAMTIME = '" & Trim(GetText(spdOrder, sRow, colEXAMTIME)) & "'" & vbCrLf
                SQL = SQL & "   AND SAVESEQ  = " & Trim(GetText(spdOrder, sRow, colSAVESEQ)) & vbCrLf
                
                If DBExec(AdoCn_Local, SQL) Then
                    '-- ����
                End If
            End If
        ElseIf sCol = colSEQNO Then
            With spdOrder
                strSeq = GetText(spdOrder, .ActiveRow, .ActiveCol)
                If Not IsNumeric(strSeq) Then
                    MsgBox "���ڸ� �Է��� �����մϴ�"
                    Exit Sub
                End If
                For intRow = .ActiveRow + 1 To .MaxRows
                    Call SetText(spdOrder, strSeq + 1, intRow, colSEQNO)
                    strSeq = strSeq + 1
                Next
            End With
        End If
    ElseIf KeyCode = vbKeyDelete Then
        If strNewBarNo = "" Then
        
        End If
        
        If MsgBox(strNewBarNo & " �� ����ðڽ��ϱ�?", vbInformation + vbYesNo, "�˸�") = vbNo Then
            Exit Sub
        End If
        
        DeleteRow spdOrder, sRow, sRow
        spdOrder.MaxRows = spdOrder.MaxRows - 1
        spdResult.MaxRows = 0
    ElseIf KeyCode = vbKeyDown Then
        DoEvents
        If sRow = spdOrder.MaxRows Then
            Exit Sub
        End If
        Call spdOrder_Click(colPNAME, sRow + 1)
        DoEvents
    ElseIf KeyCode = vbKeyUp Then
        DoEvents
        If sRow = 1 Then
            Exit Sub
        End If
        Call spdOrder_Click(colPNAME, sRow - 1)
        DoEvents
        
    End If
End Sub


Private Sub spdWork_Click(ByVal Col As Long, ByVal Row As Long)
    Dim intCol      As Integer
    Dim i           As Integer
    Dim strPatInfo  As String
    
    '-- ����
    If Row = 0 Then
        '-- ���� �߰�
        Call SetSpreadSort(spdWork, 0)
        Exit Sub
    End If
    
    
    If Row = 0 And Col = colCHECKBOX Then
        If GetText(spdWork, 1, colCHECKBOX) = "1" Then
            For i = 1 To spdWork.DataRowCnt
                Call SetText(spdWork, "0", i, colCHECKBOX)
            Next
        Else
            For i = 1 To spdWork.DataRowCnt
                Call SetText(spdWork, "1", i, colCHECKBOX)
            Next
        End If
        Exit Sub
    End If
    
    If Row > 0 And Col = colCHECKBOX Then
        If GetText(spdWork, Row, colCHECKBOX) = "1" Then
            Call SetText(spdWork, "0", Row, colCHECKBOX)
        Else
            Call SetText(spdWork, "1", Row, colCHECKBOX)
        End If
        Exit Sub
    End If
End Sub

Private Sub spdWork_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim i               As Integer
    Dim intRow          As Integer
    Dim intWRow         As Integer
    Dim intORow         As Integer
    Dim intWCol         As Integer
    Dim intOCol         As Integer
    Dim strBarno        As String
    Dim blnSame         As Boolean
    Dim varItems        As Variant
    Dim intItems        As Integer
    
    Dim strBarno_Work   As String
    
    If Row = 0 Then Exit Sub
    If Col <> colBARCODE Then
        Exit Sub
    End If
    
    intWRow = Row
    spdWork.Row = Row
    spdWork.Col = colBARCODE
    strBarno_Work = Trim(spdWork.Text)
    
    With spdOrder
        blnSame = False
        For intORow = 1 To .MaxRows
            .Row = intORow
            .Col = colBARCODE
            If strBarno_Work = Trim(.Text) Then
                blnSame = True
                Exit For
            End If
        Next
        
        If blnSame = False Then
            .MaxRows = .MaxRows + 1
            intRow = .MaxRows
            
            For i = colCHECKBOX To colSTATE
                Call SetText(spdOrder, GetText(spdWork, intWRow, i), intRow, i)
                
                varItems = GetText(spdWork, intWRow, colITEMS)
                varItems = Split(varItems, "/")
                For intItems = 0 To UBound(varItems)
                    For intOCol = colSTATE + 1 To frmMain.spdOrder.MaxCols
                        .Row = 0
                        .Col = intOCol
                        If varItems(intItems) = Trim(.Text) Then
                            .Row = intRow
                            Call SetText(spdOrder, "��", intRow, intOCol)
                        End If
                    Next
                Next
            Next
            
            Call DeleteRow(spdWork, intWRow, intWRow)
            spdWork.MaxRows = spdWork.MaxRows - 1
            .RowHeight(-1) = 15
        End If
    
    End With
End Sub

Private Sub tmrConn_Timer()
    Dim sqlRet          As Long
    Dim RS          As ADODB.Recordset
    
On Error GoTo ErrHandle
    If DbConnect_SQL = True Then
        AdoCn.CursorLocation = adUseClient
        Set RS = AdoCn.Execute("Select sysdate from DUAL", sqlRet)
        RS.Close
        
        Call SetCommStatus("R", Format(Now, "yyyy-mm-dd"), frmMain.spdComStatus)
    End If
Exit Sub

ErrHandle:
    strErrMsg = ""
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "tmrConn_Timer" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    lblDBStatus.Caption = "�����ͺ��̽� �������"
'    frmErrMsg.Show
    
End Sub

Private Sub tmrDBConn_Timer()

    DoEvents

    If imgNet2.Visible = True Then
        imgNet2.Visible = False
        imgNet3.Visible = True
        imgNet3.ZOrder
    Else
        imgNet3.Visible = False
        imgNet2.Visible = True
        imgNet2.ZOrder
    End If
    
End Sub

Private Sub tmrReceive_Timer()
    
    imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
    tmrReceive.Enabled = False

End Sub

Private Sub tmrSend_Timer()
    
    imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
    tmrSend.Enabled = False

End Sub

Private Sub wSCK_Close()
        
    If gComm.TCPTYPE = "SERVER" Then
        wSck.Close
        wSck.LocalPort = CInt(gComm.TCPPORT)
        wSck.Listen

        lblComStatus.Caption = "TCP " & gComm.TCPPORT & " ��Ʈ ���Ἲ��"
    Else
        wSck.Close
        wSck.Connect gComm.TCPIP, CInt(gComm.TCPPORT)

        lblComStatus.Caption = "TCP " & gComm.TCPIP & ":" & gComm.TCPPORT & " ��Ʈ ���Ἲ��"
    End If

End Sub

Private Sub wSCK_ConnectionRequest(ByVal requestID As Long)
            
    If wSck.State <> sckClosed Then
        wSck.Close

        wSck.Accept requestID
        imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
        If gComm.TCPTYPE = "SERVER" Then
            lblComStatus.Caption = "TCP " & gComm.TCPPORT & " ��Ʈ ���Ἲ��"
        Else
            lblComStatus.Caption = "TCP " & gComm.TCPIP & ":" & gComm.TCPPORT & " ��Ʈ ���Ἲ��"
        End If
    End If
            
End Sub

Private Sub wSCK_DataArrival(ByVal bytesTotal As Long)
    Dim strText     As String
    Dim varBuffers  As Variant
    
    wSck.GetData strText
    SetRawData "[Rx]" & strText
    
    '-- �����Ͻ� ������ ��!!
'    strText = Replace(strText, vbLf, "")
    pBuffer = strText
    
    If Len(pBuffer) > 0 Then
    
        Select Case UCase(gHOSP.MACHNM)
            Case "KLITE":                    Call Phase_TCP_KLITE
        
        End Select
    End If

End Sub

