VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "OK SOFT"
   ClientHeight    =   11880
   ClientLeft      =   60
   ClientTop       =   -1530
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
   StartUpPosition =   1  '������ ���
   WindowState     =   2  '�ִ�ȭ
   Begin VB.Frame fraWorkInfo 
      BackColor       =   &H00FFFFFF&
      Height          =   705
      Left            =   60
      TabIndex        =   44
      Top             =   660
      Width           =   5895
      Begin VB.CommandButton cmdMatch 
         BackColor       =   &H00FFFFFF&
         Caption         =   ">>"
         Height          =   375
         Left            =   5340
         Style           =   1  '�׷���
         TabIndex        =   52
         Top             =   420
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��ȸ"
         Height          =   375
         Left            =   4020
         Style           =   1  '�׷���
         TabIndex        =   46
         Top             =   180
         Width           =   855
      End
      Begin VB.CommandButton cmdAll 
         BackColor       =   &H00FFFFFF&
         Caption         =   "All"
         Height          =   375
         Left            =   4950
         Style           =   1  '�׷���
         TabIndex        =   45
         Top             =   180
         Width           =   825
      End
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   315
         Left            =   1020
         TabIndex        =   47
         Top             =   240
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
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
         Format          =   138674177
         CurrentDate     =   40457
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   315
         Left            =   2610
         TabIndex        =   48
         Top             =   240
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
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
         Format          =   138674177
         CurrentDate     =   40457
      End
      Begin VB.Label Label1 
         Appearance      =   0  '���
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '����
         Caption         =   "��ȸ�Ⱓ"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   150
         TabIndex        =   50
         Top             =   330
         Width           =   780
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
         Left            =   2430
         TabIndex        =   49
         Top             =   330
         Width           =   150
      End
   End
   Begin VB.PictureBox picComm 
      Align           =   2  '�Ʒ� ����
      Height          =   675
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   21840
      TabIndex        =   32
      Top             =   10620
      Visible         =   0   'False
      Width           =   21900
      Begin VB.CommandButton cmdRcvClear 
         Caption         =   "C"
         Height          =   495
         Left            =   12930
         TabIndex        =   42
         Top             =   60
         Width           =   495
      End
      Begin VB.CommandButton cmdEot 
         Caption         =   "EOT"
         Height          =   405
         Left            =   20880
         TabIndex        =   41
         Top             =   120
         Width           =   585
      End
      Begin VB.CommandButton cmdEtx 
         Caption         =   "ETX"
         Height          =   405
         Left            =   20280
         TabIndex        =   40
         Top             =   120
         Width           =   585
      End
      Begin VB.CommandButton cmdStx 
         Caption         =   "STX"
         Height          =   405
         Left            =   19680
         TabIndex        =   39
         Top             =   120
         Width           =   585
      End
      Begin VB.CommandButton cmdAck 
         Caption         =   "ACK"
         Height          =   405
         Left            =   19080
         TabIndex        =   38
         Top             =   120
         Width           =   585
      End
      Begin VB.CommandButton cmdEnq 
         Caption         =   "ENQ"
         Height          =   405
         Left            =   18480
         TabIndex        =   37
         Top             =   120
         Width           =   585
      End
      Begin VB.TextBox txtSend 
         Height          =   555
         Left            =   13560
         MultiLine       =   -1  'True
         TabIndex        =   36
         Top             =   30
         Width           =   3435
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "Send"
         Height          =   525
         Left            =   17010
         TabIndex        =   35
         Top             =   60
         Width           =   1125
      End
      Begin VB.TextBox txtRcv 
         Height          =   525
         Left            =   60
         MultiLine       =   -1  'True
         TabIndex        =   34
         Top             =   60
         Width           =   11805
      End
      Begin VB.CommandButton cmdRcv 
         Caption         =   "Rcv"
         Height          =   525
         Left            =   11940
         TabIndex        =   33
         Top             =   60
         Width           =   975
      End
   End
   Begin VB.Frame fraPatInfo 
      BackColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   16440
      TabIndex        =   20
      Top             =   660
      Width           =   6525
      Begin VB.TextBox txtSA 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4350
         Locked          =   -1  'True
         TabIndex        =   30
         Text            =   "1004"
         Top             =   750
         Width           =   1935
      End
      Begin VB.TextBox txtPName 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   28
         Text            =   "1004"
         Top             =   750
         Width           =   1935
      End
      Begin VB.TextBox txtPatID 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4350
         Locked          =   -1  'True
         TabIndex        =   26
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
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "1004"
         Top             =   270
         Width           =   1935
      End
      Begin VB.Label Label7 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '����
         Caption         =   "S / A"
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
         TabIndex        =   31
         Top             =   810
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
         TabIndex        =   29
         Top             =   810
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
         TabIndex        =   27
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
         TabIndex        =   25
         Top             =   330
         Width           =   885
      End
   End
   Begin VB.Frame fraHidden 
      Caption         =   "Hidden"
      Height          =   2355
      Left            =   11910
      TabIndex        =   18
      Top             =   7620
      Visible         =   0   'False
      Width           =   6255
      Begin VB.TextBox Text1 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   540
         TabIndex        =   58
         Text            =   "0"
         Top             =   330
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.TextBox txtLastSeq 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   930
         TabIndex        =   53
         Text            =   "0"
         Top             =   330
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdResult 
         Caption         =   "�����ȸ"
         Height          =   315
         Left            =   2610
         TabIndex        =   19
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
         TabIndex        =   23
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
      TabIndex        =   7
      Top             =   11295
      Width           =   21900
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
         Left            =   1770
         Top             =   -30
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
               Picture         =   "frmMain.frx":113A
               Key             =   "RUN"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":16D4
               Key             =   "NOT"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1C6E
               Key             =   "STOP"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2208
               Key             =   "LST"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2A9A
               Key             =   "ITM"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2BF4
               Key             =   "ERR"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2D4E
               Key             =   "NOF"
            EndProperty
         EndProperty
      End
      Begin FPSpread.vaSpread spdComStatus 
         Height          =   330
         Left            =   8010
         TabIndex        =   13
         Top             =   120
         Width           =   3570
         _Version        =   393216
         _ExtentX        =   6297
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
         SpreadDesigner  =   "frmMain.frx":2EA8
         UserResize      =   0
         TextTip         =   2
         ScrollBarTrack  =   3
         ShowScrollTips  =   3
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H00FFFFFF&
         Height          =   405
         Left            =   7980
         Top             =   90
         Width           =   3645
      End
      Begin VB.Image imgPort 
         Height          =   240
         Left            =   3900
         Picture         =   "frmMain.frx":32F9
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgSend 
         Height          =   240
         Left            =   4845
         Picture         =   "frmMain.frx":3883
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgReceive 
         Height          =   240
         Left            =   5760
         Picture         =   "frmMain.frx":3E0D
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
         Top             =   210
         Width           =   420
      End
      Begin VB.Image imgNet1 
         Height          =   240
         Left            =   390
         Picture         =   "frmMain.frx":4397
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgNet2 
         Height          =   240
         Left            =   390
         Picture         =   "frmMain.frx":44E1
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgNet3 
         Height          =   240
         Left            =   390
         Picture         =   "frmMain.frx":462B
         Top             =   180
         Width           =   240
      End
      Begin VB.Label lblComStatus 
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
         TabIndex        =   9
         Top             =   180
         Width           =   1695
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
         TabIndex        =   8
         Top             =   180
         Width           =   2295
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00FFFFFF&
         Height          =   405
         Left            =   3210
         Top             =   90
         Width           =   4785
      End
   End
   Begin FPSpread.vaSpread spdResult 
      Height          =   6555
      Left            =   16440
      TabIndex        =   4
      Top             =   2010
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
      SpreadDesigner  =   "frmMain.frx":4775
      ScrollBarTrack  =   3
      ShowScrollTips  =   3
   End
   Begin FPSpread.vaSpread spdOrder 
      Height          =   7935
      Left            =   5970
      TabIndex        =   2
      Top             =   630
      Width           =   15255
      _Version        =   393216
      _ExtentX        =   26908
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
      SpreadDesigner  =   "frmMain.frx":54F7
      ScrollBarTrack  =   3
      ShowScrollTips  =   3
   End
   Begin VB.PictureBox picHeader 
      Align           =   1  '�� ����
      BackColor       =   &H00800000&
      BorderStyle     =   0  '����
      Height          =   585
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   21900
      TabIndex        =   0
      Top             =   0
      Width           =   21900
      Begin VB.Frame fraVISION 
         Appearance      =   0  '���
         BackColor       =   &H00800000&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   14880
         TabIndex        =   54
         Top             =   0
         Width           =   4965
         Begin VB.CommandButton Command1 
            Appearance      =   0  '���
            Caption         =   "�ޱ�"
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
            Left            =   3900
            TabIndex        =   61
            Top             =   150
            Width           =   645
         End
         Begin VB.TextBox Text2 
            Alignment       =   2  '��� ����
            Height          =   285
            Left            =   3090
            TabIndex        =   60
            Text            =   "1"
            Top             =   150
            Width           =   345
         End
         Begin VB.TextBox Text3 
            Alignment       =   2  '��� ����
            Height          =   285
            Left            =   3450
            TabIndex        =   59
            Text            =   "1"
            Top             =   150
            Width           =   345
         End
         Begin VB.TextBox txtRCnt 
            Alignment       =   2  '��� ����
            Appearance      =   0  '���
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1170
            TabIndex        =   56
            Text            =   "1"
            Top             =   180
            Width           =   525
         End
         Begin VB.CommandButton cmdGetRslt 
            Appearance      =   0  '���
            Caption         =   "����ޱ�"
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
            Left            =   1830
            TabIndex        =   55
            Top             =   150
            Width           =   1155
         End
         Begin VB.Label Label8 
            BackStyle       =   0  '����
            Caption         =   "�������"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   240
            TabIndex        =   57
            Top             =   210
            Width           =   915
         End
      End
      Begin VB.CommandButton cmdWork 
         Caption         =   "��ũ��ȸ"
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
         Left            =   13290
         TabIndex        =   43
         Top             =   150
         Width           =   1425
      End
      Begin MSWinsockLib.Winsock wSCK 
         Left            =   20820
         Top             =   150
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "ȭ������"
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
         Left            =   10110
         TabIndex        =   22
         ToolTipText     =   "����ȭ���� ��� ����ϴ�"
         Top             =   150
         Width           =   1425
      End
      Begin VB.CommandButton cmdSave 
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
         Height          =   315
         Left            =   11700
         TabIndex        =   21
         ToolTipText     =   "������ ����� EMR������ �����մϴ�"
         Top             =   150
         Width           =   1425
      End
      Begin VB.CommandButton cmdTestNmSave 
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
         Left            =   9390
         TabIndex        =   17
         Top             =   150
         Width           =   555
      End
      Begin VB.TextBox txtTestNm 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   7680
         TabIndex        =   16
         Text            =   "1004"
         Top             =   180
         Width           =   1635
      End
      Begin VB.TextBox txtTestID 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4230
         TabIndex        =   15
         Text            =   "1004"
         Top             =   180
         Width           =   1635
      End
      Begin VB.CommandButton cmdTestIDSave 
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
         Left            =   5940
         TabIndex        =   14
         Top             =   150
         Width           =   555
      End
      Begin MSCommLib.MSComm comEqp 
         Left            =   20100
         Top             =   90
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
         RThreshold      =   1
         EOFEnable       =   -1  'True
      End
      Begin VB.Shape Shape10 
         BorderColor     =   &H0080FFFF&
         Height          =   405
         Left            =   13230
         Top             =   90
         Width           =   1545
      End
      Begin VB.Shape Shape8 
         BorderColor     =   &H0080FFFF&
         Height          =   405
         Left            =   10050
         Top             =   90
         Width           =   1545
      End
      Begin VB.Shape Shape9 
         BorderColor     =   &H0080FFFF&
         Height          =   405
         Left            =   11640
         Top             =   90
         Width           =   1545
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
         ForeColor       =   &H00C0FFC0&
         Height          =   285
         Left            =   1560
         TabIndex        =   6
         Top             =   180
         Width           =   1365
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00C0FFFF&
         BorderWidth     =   2
         Height          =   405
         Left            =   210
         Top             =   90
         Width           =   2865
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FFFFFF&
         Height          =   405
         Left            =   6570
         Top             =   90
         Width           =   3405
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   6690
         TabIndex        =   5
         Top             =   180
         Width           =   975
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3240
         TabIndex        =   3
         Top             =   180
         Width           =   975
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         Height          =   405
         Left            =   3120
         Top             =   90
         Width           =   3405
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   330
         TabIndex        =   1
         Top             =   180
         Width           =   945
      End
   End
   Begin FPSpread.vaSpread spdWork 
      Height          =   7275
      Left            =   60
      TabIndex        =   51
      Top             =   1410
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
      SpreadDesigner  =   "frmMain.frx":7555
      ScrollBarTrack  =   3
      ShowScrollTips  =   3
   End
   Begin VB.Menu mnuMenu01 
      Caption         =   " ���� "
      Begin VB.Menu mnuHosp 
         Caption         =   "�� ���� ����"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEMR 
         Caption         =   "�� EMR ����"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep9 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuExit 
         Caption         =   "����"
      End
   End
   Begin VB.Menu mnuMenu04 
      Caption         =   " ��ȸ "
      Begin VB.Menu mnuResult 
         Caption         =   "�� ��� ��ȸ"
      End
      Begin VB.Menu mnuSep29 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWork 
         Caption         =   "�� ��ũ ��ȸ"
      End
   End
   Begin VB.Menu mnuMenu02 
      Caption         =   " ���� "
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
      Caption         =   " �˻�ɼ� "
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
   End
   Begin VB.Menu mnuMenu03 
      Caption         =   " �������� "
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

Private Sub cmdGetRslt_Click()
    Dim strSendData As String
    Dim strFirstSeq  As String
    Dim strLastSeq  As String
    Dim db_tmp As String * 100

On Error GoTo Err

    strFirstSeq = txtLastSeq.Text
    strFirstSeq = (strFirstSeq - 1) - (txtRCnt.Text - 1)

    strLastSeq = strFirstSeq + (txtRCnt.Text - 1)

    strSendData = Text1 & vbTab & "GET" & vbTab & strFirstSeq & vbTab & strLastSeq

    wSCK.SendData strSendData & vbLf
    SetRawData "[Tx]" & strSendData & vbLf

    

        
Exit Sub

Err:
    strErrMsg = ""
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "cmdGetRslt_Click" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
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
                
        Case "VISION":          Call Phase_TCP_VISION
        Case "BC5800":          Call Phase_TCP_BC5800
        Case "UROMETER720":     Call Phase_Serial_UROMETER720
        Case "HITACHI7060":     Call Phase_Serial_HITACHI7060
        Case "STAGO":           Call Phase_Serial_STAGO
        Case "ISMART300":       Call Phase_Serial_ISMART300
                    
    End Select

    pBuffer = ""
    
End Sub

Private Sub cmdRcvClear_Click()
    
    txtRcv.Text = ""
    
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
                If .Value = 1 And Trim(GetText(spdOrder, lRow, colSTATE)) <> "" Then
                    
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
    
    Call GetWorkList(Format(dtpFrom.Value, "yyyy-mm-dd"), Format(dtpTo.Value, "yyyy-mm-dd"), spdWork)

End Sub

Private Sub cmdSend_Click()
    
    
    Call SendData(txtSend.Text)

End Sub

Private Sub cmdStx_Click()
    
    txtSend.Text = txtSend.Text & STX

End Sub

Private Sub cmdTestIDSave_Click()
    
    Call WritePrivateProfileString("HOSP", "USERID", txtTestID.Text, App.PATH & "\INI\" & gMACH & ".ini")

End Sub

Private Sub cmdTestNmSave_Click()
    
    Call WritePrivateProfileString("HOSP", "USERNM", txtTestNm.Text, App.PATH & "\INI\" & gMACH & ".ini")

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
                Case "HITACHI7060":     Call Phase_Serial_HITACHI7060
                Case "ISMART300":       Call Phase_Serial_ISMART300
                Case "STAGO":           Call Phase_Serial_STAGO
                Case "UROMETER720":     Call Phase_Serial_UROMETER720
                            
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

Private Sub Command1_Click()
    
    Dim strSendData As String


    strSendData = Text1 & vbTab & "GET" & vbTab & Text2 & vbTab & Text3

    wSCK.SendData strSendData & vbLf
    SetRawData "[Tx]" & strSendData & vbLf

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
        strItems = GetEquipExamCode_HITACHI7180(gHOSP.MACHCD, pBarNo, intRow)
        mOrder.Order = strItems
        
        If gHOSP.BARUSE = "N" Then
            mOrder.Function = Replace(mOrder.Function, String(gHOSP.BARLEN, "#"), Left(mOrder.BarNo & Space(gHOSP.BARLEN), gHOSP.BARLEN))
        End If
        
        '-- �˻�ä�η� ������ �����
        If mOrder.SendCnt = 0 Then
            mOrder.NoOrder = True

            GetOrder = STX & ";" & mOrder.Function & " 88" & mOrder.Order & "100000" & Left(mOrder.PID & Space(30), 30) & ETX

            '-- �������(Order) ǥ��
            Call SetText(frmMain.spdOrder, "��������", intRow, colSTATE)
        Else
            mOrder.NoOrder = False

            GetOrder = STX & ";" & mOrder.Function & " 88" & mOrder.Order & "100000" & Left(mOrder.PID & Space(30), 30) & ETX

            '-- �������(Order) ǥ��
            Call SetText(frmMain.spdOrder, "��������", intRow, colSTATE)
        End If

        Call SendData(GetOrder)
        
        Call SetCommStatus("S", pBarNo, spdComStatus)

        '-- ���� Row
        gRow = intRow

    End With

End Sub

Private Sub GetOrder_HITACHI7060(ByVal pBarNo As String, ByVal pType As String)

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
        strItems = GetEquipExamCode_HITACHI7060(gHOSP.MACHCD, pBarNo, intRow)
        mOrder.Order = strItems
        
        'If gHOSP.BARUSE = "N" Then
        '    mOrder.Function = Replace(mOrder.Function, String(gHOSP.BARLEN, "#"), Left(mOrder.BarNo & Space(gHOSP.BARLEN), gHOSP.BARLEN))
        'End If
        
        '-- �˻�ä�η� ������ �����
        If mOrder.SendCnt = 0 Then
            mOrder.NoOrder = True

            GetOrder = STX & ";" & Mid(mOrder.Function, 1, 3) & Left(mOrder.Seq & Space(3), 3) & "0" & Left(mOrder.TubePos & Space(2), 2) & Left(mOrder.BarNo & Space(13), 13) & Mid(mOrder.Function, 23, 15) & "48" & mOrder.Order & ETX

            '-- �������(Order) ǥ��
            Call SetText(frmMain.spdOrder, "��������", intRow, colSTATE)
        Else
            mOrder.NoOrder = False
            
            GetOrder = STX & ";" & Mid(mOrder.Function, 1, 3) & Left(mOrder.Seq & Space(3), 3) & "0" & Left(mOrder.TubePos & Space(2), 2) & Left(mOrder.BarNo & Space(13), 13) & Mid(mOrder.Function, 23, 15) & "48" & mOrder.Order & ETX

            '-- �������(Order) ǥ��
            Call SetText(frmMain.spdOrder, "��������", intRow, colSTATE)
        End If

        Call SendData(GetOrder)
        
        Call SetCommStatus("S", pBarNo, spdComStatus)

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

End Sub



Private Sub SerialRcvData_HITACHI7060()
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
                    'gHOSP.BARLEN = 13
                    strRackNo = Mid(strRcvBuf, 7, 1)
                    strTubePos = Mid(strRcvBuf, 8, 2)
                    strSeq = Mid(strRcvBuf, 10, 3)
                    strBarno = strSeq
                    strFunction = strRcvBuf
                    '-- ��������
                    With mOrder
                        .Function = strFunction
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .Seq = strSeq
                    End With
                    '-- ����ȯ������
                    Call GetOrder_HITACHI7060(Trim$(strBarno), gHOSP.RSTTYPE)

                Case ":"    '## End
                    
                    ':s   10 1                3 0513191028 212      V13   1.0 
                    ':s   20 2                3 050819135514 1      V 2      V 3      V 5      V 6      V 7      V 8      V 9      V10      V11      V13      V14      V15      V16      V
                    
                    ':S   10 1                3 053119113016 1   6.8  2   4.0  3   0.3  6    13  7    20  8    14  9   156 10   190 11   120 12   4.9 13   0.7 14    15 15   324 16    33 18  0.00 23     0 25

                    '## Control, Calibration �����ʹ� ������
                    strFunc = Mid$(strRcvBuf, 2, 1)
                    If UCase(strFunc) = "H" Or UCase(strFunc) = "G" Or UCase(strFunc) = "F" Then
                        '-- ���� ����
                        Call SendData(SndMore)
                        strState = ""
                        Exit Sub
                    End If
                    
                    strSeq = Mid(strRcvBuf, 4, 3)
                    strRackNo = Mid(strRcvBuf, 7, 1)
                    strTubePos = Mid(strRcvBuf, 8, 2)
                    'strBarno = Trim$(Mid$(strRcvBuf, 14, gHOSP.BARLEN)) '13
                    strBarno = strSeq
                    
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

                    If mResult.BarNo = "" Then
                        Call SendData(SndMore)
                        Exit Sub
                    End If

                    strTmp = Mid$(strRcvBuf, 40)
                    
                    strState = ""
                    
                    Do While Len(strTmp) >= 9
                        strIntBase = Trim(Mid$(strTmp, 1, 2))
                        strResult = Trim(Mid$(strTmp, 3, 6))
                        strFlag = Trim(Mid$(strTmp, 9, 1))
                        
                        strSeqno = ""
                        strTestCode = ""
                        strTestName = ""
                        intResPrecUse = -1
                        intResPrec = -1
                            
                        '-- �˻縶���� ���� ��������
                        If strIntBase <> "" And strResult <> "" Then
                            blnSame = False
                            '-- �˻縶���� ���� ��������
                            For intTestNmCnt = 1 To UBound(gArrEQPNm)
                                '-- ���ä���� ����...
                                If strIntBase = gArrEQPNm(intTestNmCnt, 3) Then
                                    strCheck = "0"
                                    strSeqno = gArrEQPNm(intTestNmCnt, 1)
                                    If strState <> "R" Then
                                        strState = ""
                                    End If
                                    
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
                                            
                                            '-- ȯ�� ó������ ��������
                                            If UBound(gPatTest) > 0 Then
                                                For intOrdCnt = 1 To UBound(gPatTest)
                                                    If gPatTest(intOrdCnt) = gArrEQP(intTestCdCnt, 2) Then
                                                        'strTestCode = gArrEQP(intTestCdCnt, 2)
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
                                                        
                                                        strCheck = "1"
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
                            
                            If strCheck <> "" Then
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
                                        strTestCodeSub = gArrEQP(intCol - colSTATE, 17)
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
                            End If
                        End If
                        strTmp = Mid$(strTmp, 10)
                    Loop

                    .spdResult.RowHeight(-1) = 15

                    '## DB�� �������
                    If gHOSP.SAVEAUTO = "Y" Then
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
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "SerialRcvData_HITACHI7060" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub



Private Sub SerialRcvData_UROMETER720()
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
    Dim Pos As Integer

On Error GoTo RST

    With frmMain
        RcvBuffer = Replace(RcvBuffer, vbLf, "")
        strRecvData = Split(RcvBuffer, vbCr)
        
        For intCnt = 0 To UBound(strRecvData)
            strRcvBuf = strRecvData(intCnt)
'~NOV/28/2016 10:10:17
'Name:             SEX:
'Ward:             AGE:
'SD (10)      (0001)
'BLD -neg
'BIL -neg
'URO    norm    0.1mg/dl
'KET -neg
'PRO -neg
'NIT -neg
'GLU -neg
'pH 5.5
'S.G 1.005
'LEU -neg
'(SN=60158394)
'ID(             )
'OP(00000000) LOT(000000) ~
'~FinDb~

            Call SetSQLData("RCV", strRcvBuf, "A")
            
            Select Case intCnt
                Case 3
                    strSeq = Mid(strRcvBuf, 10)
                    strSeq = Replace(strSeq, ")", "")
                    strSeq = Replace(strSeq, "(", "")
                    strSeq = Val(Trim(strSeq))
                    strBarno = strSeq
                    
                    '-- �������
                    With mResult
                        .BarNo = strBarno
                        '.RackNo = strRackNo
                        '.TubePos = strTubePos
                        .Seq = strSeq
                        .RsltDate = Format(Now, "yyyy-mm-dd")
                        .RsltTime = Format(Now, "hh:mm:ss")
                        .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                    End With
                    
                    '-- ���ȯ������
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    strState = ""
            
                Case 4 To 13
                    strIntBase = Mid(strRcvBuf, 1, 4)
                    strIntBase = Trim(strIntBase)
                    
                    'strResult = Mid(strRcvBuf, 7, 5) '-- ����
                    strResult = Mid(strRcvBuf, 8, 4) '-- ����
                    strResult = Trim(strResult)
            
                    If strIntBase = "pH" Or strIntBase = "p.H" Or strIntBase = "S.G" Then
                        strResult = Trim(Mid(strRcvBuf, 4))  '-- ����
                        strResult = Replace(strResult, "mg/dl", "")
                        strResult = Replace(strResult, "RBC/ul", "")
                        strResult = Replace(strResult, "WBC/ul", "")
                        
                        strResult = Replace(strResult, "<", "")
                        strResult = Replace(strResult, ">", "")
                        strResult = Replace(strResult, "=", "")
                    End If
                    
                    '-- URO
                    If strResult = "norm" Then
                        strResult = "-"
                    End If
    '
    '                '-- NIT
                    If strResult = "pos" Then
                        strResult = "+"
                    End If
            
                    Select Case Trim(strResult)
                        Case "-":       strResult = "negative"
                        Case "+":       strResult = "1positive"
                        Case "++":      strResult = "2positive"
                        Case "+++":     strResult = "3positive"
                        Case "++++":    strResult = "4positive"
                        Case "+/-":     strResult = "trace"
                    End Select
                    
                    strCheck = ""
                    '-- �˻縶���� ���� ��������
                    If strIntBase <> "" And strResult <> "" Then
                        blnSame = False
                        '-- �˻縶���� ���� ��������
                        For intTestNmCnt = 1 To UBound(gArrEQPNm)
                            '-- ���ä���� ����...
                            If strIntBase = gArrEQPNm(intTestNmCnt, 3) Then
                                strCheck = "0"
                                strSeqno = gArrEQPNm(intTestNmCnt, 1)
                                If strState <> "R" Then
                                    strState = ""
                                End If
                                
                                
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
                                        
                                        '-- ȯ�� ó������ ��������
                                        If UBound(gPatTest) > 0 Then
                                            For intOrdCnt = 1 To UBound(gPatTest)
                                                If gPatTest(intOrdCnt) = gArrEQP(intTestCdCnt, 2) Then
                                                    
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
                                                    
                                                    strCheck = "1"
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
    
                        If strCheck <> "" Then
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
                                    strTestCodeSub = gArrEQP(intCol - colSTATE, 17)
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
                        End If
        
                        .spdResult.RowHeight(-1) = 15
        
                        '## DB�� �������
                        If gHOSP.SAVEAUTO = "Y" Then
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
            End Select
        Next
    End With

Exit Sub

RST:
    strErrMsg = ""
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "SerialRcvData_UROMETER720" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub


Private Sub SerialRcvData_STAGO()
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
    Dim strTemp1        As String
    
On Error GoTo RST

    ReDim Preserve strRData(UBound(strRecvData))
    
    For i = 1 To UBound(strRecvData)
        strRData(i) = strRecvData(i)
    Next
    
    With frmMain
        For intCnt = 1 To UBound(strRData)
            strRcvBuf = strRData(intCnt)

            strType = Mid$(strRcvBuf, 2, 1)
            If strType = "|" Then
                strType = Mid$(strRcvBuf, 1, 1)
            End If

            Select Case strType
                Case "H"
                Case "Q"
'                    strTemp1 = mGetP(strRcvBuf, 3, "|")
'                    strBarno = Trim$(mGetP(strTemp1, 2, "^"))
'
'                    With mOrder
'                        .NoOrder = False
'                        .BarNo = strBarno
'                    End With
'
'                    Call GetOrder_STAGO(strBarno, gHOSP.RSTTYPE)
'
'                    strState = "Q"
                Case "P"
                Case "O"
                    'gHOSP.BARLEN = 13
                    strTemp1 = mGetP(strRcvBuf, 3, "|")
                    strBarno = mGetP(strTemp1, 1, "^")
                    strSeq = mGetP(strTemp1, 2, "^")
                    strTubePos = mGetP(strTemp1, 3, "^")
                    
                    '-- �������
                    With mResult
                        .BarNo = strBarno
                        .TubePos = strTubePos
                        .Seq = strSeq
                        .RsltDate = Format(Now, "yyyy-mm-dd")
                        .RsltTime = Format(Now, "hh:mm:ss")
                        .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                    End With
                    '-- ���ȯ������
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    strState = ""
                Case "R"
                
                    strTemp1 = mGetP(strRcvBuf, 3, "|")
                    strIntBase = mGetP(strTemp1, 4, "^")
                    strFlag = mGetP(strRcvBuf, 9, "|")
                    strIntResult = mGetP(strRcvBuf, 4, "|")
                    
                    Select Case strFlag
                        Case "F"    '## ����
                            'strIntBase = strIntBase & "N"
                            strResult = strIntResult
                        Case "I"    '## ����
                            'strIntBase = strIntBase & "C"
                            Select Case Mid$(strIntResult, 1, 1)
                                Case "N":   strResult = "Negative"
                                Case "G":   strResult = "GRAYZONE"
                                Case "R":   strResult = "Positive"
                                Case "P":   strResult = "Positive"
                            End Select
                    End Select


                    '-- �˻縶���� ���� ��������
                    If strIntBase <> "" And strResult <> "" Then
                        blnSame = False
                        '-- �˻縶���� ���� ��������
                        For intTestNmCnt = 1 To UBound(gArrEQPNm)
                            '-- ���ä���� ����...
                            If strIntBase = gArrEQPNm(intTestNmCnt, 3) Then
                                strCheck = "0"
                                strSeqno = gArrEQPNm(intTestNmCnt, 1)
                                If strState <> "R" Then
                                    strState = ""
                                End If
                                
                                
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
                                        
                                        '-- ȯ�� ó������ ��������
                                        If UBound(gPatTest) > 0 Then
                                            For intOrdCnt = 1 To UBound(gPatTest)
                                                If gPatTest(intOrdCnt) = gArrEQP(intTestCdCnt, 2) Then
                                                    
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
                                                    
                                                    strCheck = "1"
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
    
                        If strCheck <> "" Then
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
                        End If
                    End If

                    .spdResult.RowHeight(-1) = 15
                
                Case "L"
                    '## DB�� �������
                    If gHOSP.SAVEAUTO = "Y" Then
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
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "SerialRcvData_STAGO" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub SerialRcvData_ISMART300()
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
    Dim strTemp1        As String
    
On Error GoTo RST

    ReDim Preserve strRData(UBound(strRecvData))
    
    For i = 1 To UBound(strRecvData)
        strRData(i) = strRecvData(i)
    Next
    
'1H|\^|||i-Smart 300^G10152^-^1.1.2.2||||||||1394-97|20190515100120D2
'2P|1||17112800041||^|||||||||||||||||||||||||||||B5
'3O|1||170912-1-6-S4||||||||||||Arterial|||||||||||||||1B
'4C|1|I||G2B
'5R|1|^^^pH^M|7.509||7.350^7.450^Ref. Range|^H^||R||||20171128164905|BB
'6R|2|^^^pCO2^M|34.4|mmHg|32.0^48.0^Ref. Range|^N^||R|||||51
'7R|3|^^^pO2^M|128|mmHg|83^108^Ref. Range|^H^||R|||||53
'0R|4|^^^cNa+^M|122|mmol/L|136^145^Ref. Range|^L^||R|||||6E
'1R|5|^^^cK+^M|4.8|mmol/L|3.4^4.5^Ref. Range|^H^||R|||||05
'2R|6|^^^cCa2+^M|1.20|mmol/L|1.15^1.33^Ref. Range|^N^||R|||||1D
'3R|7|^^^cCl-^M|106|mmol/L|98^107^Ref. Range|^N^||R|||||4F
    
    With frmMain
        For intCnt = 1 To UBound(strRData)
            strRcvBuf = strRData(intCnt)

            strType = Mid$(strRcvBuf, 2, 1)
            If strType = "|" Then
                strType = Mid$(strRcvBuf, 1, 1)
            End If

            Select Case strType
                Case "H"
                Case "Q"
                Case "P"
                    '2P|1||test||^|||||||||||||||||||||||||||||4C
                    strBarno = mGetP(strRcvBuf, 4, "|")
                Case "O"
                    '3O|1||190507-1-60-S20||||||||||||Arterial|||||||||||||||7B
                    'gHOSP.BARLEN = 13
                    strSeq = mGetP(strRcvBuf, 3, "|")
                    
                    '-- �������
                    With mResult
                        .BarNo = strBarno
                        .TubePos = strTubePos
                        .Seq = strSeq
                        .RsltDate = Format(Now, "yyyy-mm-dd")
                        .RsltTime = Format(Now, "hh:mm:ss")
                        .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                    End With
                    '-- ���ȯ������
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    strState = ""
                    
                Case "R"
                    '0R|4|^^^cNa+^M|122|mmol/L|136^145^Ref. Range|^L^||R|||||6E
                    strCheck = ""

                    '0R|4|^^^cNa+^M|137|mmol/L||^-^||R|||||0D
                    
                    strTemp1 = mGetP(strRcvBuf, 3, "|")
                    strIntBase = mGetP(strTemp1, 4, "^")
                    strIntResult = mGetP(strRcvBuf, 4, "|")
                    strFlag = mGetP(strRcvBuf, 9, "|")
                    
'                    Select Case strFlag
'                        Case "F"    '## ����
'                            'strIntBase = strIntBase & "N"
'                            strResult = strIntResult
'                        Case "I"    '## ����
'                            'strIntBase = strIntBase & "C"
'                            Select Case Mid$(strIntResult, 1, 1)
'                                Case "N":   strResult = "Negative"
'                                Case "G":   strResult = "GRAYZONE"
'                                Case "R":   strResult = "Positive"
'                                Case "P":   strResult = "Positive"
'                            End Select
'                    End Select
                    
                    strResult = strIntResult
        
                    '-- �˻縶���� ���� ��������
                    If strIntBase <> "" And strResult <> "" Then
                        blnSame = False
                        '-- �˻縶���� ���� ��������
                        For intTestNmCnt = 1 To UBound(gArrEQPNm)
                            '-- ���ä���� ����...
                            If strIntBase = gArrEQPNm(intTestNmCnt, 3) Then
                                strCheck = "0"
                                strSeqno = gArrEQPNm(intTestNmCnt, 1)
                                
                                If strState <> "R" Then
                                    strState = ""
                                End If
                                
                                
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
                                        
                                        '-- ȯ�� ó������ ��������
                                        If UBound(gPatTest) > 0 Then
                                            For intOrdCnt = 1 To UBound(gPatTest)
                                                If gPatTest(intOrdCnt) = gArrEQP(intTestCdCnt, 2) Then
                                                    
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
                                                    
                                                    strCheck = "1"
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
    
                        If strCheck <> "" Then
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
                                    strTestCodeSub = gArrEQP(intCol - colSTATE, 17)
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
                        End If
                    End If

                    .spdResult.RowHeight(-1) = 15
                
                Case "L"
                    '## DB�� �������
                    If gHOSP.SAVEAUTO = "Y" Then
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
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "SerialRcvData_ISMART300" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub

Private Sub TCPRcvData_BC5800()
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

    
    RcvBuffer = Replace(RcvBuffer, vbLf, "")
    
    strRecvData = Split(RcvBuffer, vbCr)
    
    ReDim Preserve strRData(UBound(strRecvData))
    
    strRData = strRecvData
'    For i = 1 To UBound(strRecvData)
'        strRData(i) = strRecvData(i)
'    Next
    
    strOldBarno = ""
    
  '  MsgBox UBound(strRData)
    
    With frmMain
        For intCnt = 1 To UBound(strRData)
            
            If intCnt >= 50 Then
                Exit For
            End If
            
            strRcvBuf = strRData(intCnt)

            Call SetSQLData("RCV", intCnt & ":" & strRcvBuf, "A")

            strType = mGetP(strRcvBuf, 1, "|")

            Select Case strType
                Case "MSH"
                    'MSH|^~\&|BC5800|MINDRAY|||20190625143335||ORU^R01|1|P|2.3.1||||||UNICODE
                
                Case "PID"
                    'PID|1||^^^^MR
                    
                    strBarno = Trim(mGetP(strRcvBuf, 2, "|"))

                Case "OBR"
                    'OBR|1||1|00001^Automated Count^99MRC|||20190625143333|||||||||||||||||HM||||||||1

                    '-- �������̽� ����
'                    strSend = ""
'                    strSend = strSend & SB & "MSH|^~\&|MINDRAY|BC5800|||" & Format(Now, "yyyymmddhhmmss") & "||ACK^R01|" & strBarno & "|P|2.3.1||||0||UNICODE|||" & vbCr
'                    strSend = strSend & "MSA|AA|1|Message accepted|||0|" & vbCr
'                    strSend = strSend & EB & vbCr
'
'                    SetRawData "[Tx]" & strSend
'                    wSCK.SendData strSend
                
                    strBarno = Trim(mGetP(strRcvBuf, 4, "|"))
                    
                    If strBarno = "" Then
                        Exit Sub
                    End If
                
                    If strOldBarno <> strBarno Then
                        '-- �������
                        With mResult
                            .BarNo = strBarno
                            .RsltDate = Format(Now, "yyyy-mm-dd")
                            .RsltTime = Format(Now, "hh:mm:ss")
                            .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                        End With
                        
                        '-- ���ȯ������
                        Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    End If
                    
                    strOldBarno = strBarno
                    
                    If gRow <= 0 Then
                        Exit Sub
                    End If
                    
                    strState = ""
                    
                Case "OBX"
                    'OBX|5|NM|6690-2^WBC^LN||5.68|10*3/uL|4.00-10.00||||F

                    strIntBase = Trim(mGetP(mGetP(strRcvBuf, 4, "|"), 2, "^"))
                    strResult = Trim(mGetP(strRcvBuf, 6, "|"))
                    strCheck = ""
                    
                    'Call SetSQLData("RCV", strIntBase & vbCrLf, "A")
                    
                    '-- �˻縶���� ���� ��������
                    If strIntBase <> "" And strResult <> "" Then
                        blnSame = False
                        '-- �˻縶���� ���� ��������
                        For intTestNmCnt = 1 To UBound(gArrEQPNm)
                            '-- ���ä���� ����...
                            If strIntBase = gArrEQPNm(intTestNmCnt, 3) Then
                                strCheck = "0"
                                strSeqno = gArrEQPNm(intTestNmCnt, 1)
                                
                                If strState <> "R" Then
                                    strState = ""
                                End If
                                
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
                                        
                                        '-- ȯ�� ó������ ��������
                                        If UBound(gPatTest) > 0 Then
                                            For intOrdCnt = 1 To UBound(gPatTest)
                                                If gPatTest(intOrdCnt) = gArrEQP(intTestCdCnt, 2) Then
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
                                                    
                                                    strCheck = "1"
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

                        If strCheck <> "" Then
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
                                    strTestCodeSub = gArrEQP(intCol - colSTATE, 17)
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
                        End If
                    End If

                    .spdResult.RowHeight(-1) = 15

            End Select
        Next
    
        '## DB�� �������
        If gHOSP.SAVEAUTO = "Y" Then
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
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "_TCPRcvData_BC5800" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub


Private Sub TCPRcvData_VISION()
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

'    ReDim Preserve strRData(UBound(strRecvData))
'
'    For i = 1 To UBound(strRecvData)
'        strRData(i) = strRecvData(i)
'    Next
'
'    MsgBox strBuffer
    
    strOldBarno = ""
    
    strRData = Split(strBuffer, vbLf)
    
    With frmMain
        For intCnt = 0 To UBound(strRData)
            strRcvBuf = strRData(intCnt)
            strIntBase = "ESR"
            strSeqno = mGetP(strRcvBuf, 1, vbTab)
            strBarno = mGetP(strRcvBuf, 7, vbTab)
            
            
            '-- 18�� ���
            strResult = mGetP(strRcvBuf, 10, vbTab)
            
            If strBarno = "" Then
                strBarno = strSeqno
            End If
        
'MsgBox strBarno
'MsgBox strIntBase
'MsgBox strResult

            If strOldBarno <> strBarno Then
                '-- �������
                With mResult
                    .BarNo = strBarno
                    .RsltDate = Format(Now, "yyyy-mm-dd")
                    .RsltTime = Format(Now, "hh:mm:ss")
                    .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
                End With
                
                '-- ���ȯ������
                Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
            
            End If
            
            strOldBarno = strBarno
            
            If gRow <= 0 Then
                Exit Sub
            End If
            
            strState = ""
            
            '-- �˻縶���� ���� ��������
            If strIntBase <> "" And strResult <> "" Then
                blnSame = False
                '-- �˻縶���� ���� ��������
                For intTestNmCnt = 1 To UBound(gArrEQPNm)
                    '-- ���ä���� ����...
                    If strIntBase = gArrEQPNm(intTestNmCnt, 3) Then
                        strCheck = "0"
                        strSeqno = gArrEQPNm(intTestNmCnt, 1)
                        
                        If strState <> "R" Then
                            strState = ""
                        End If
                        
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
                                
                                '-- ȯ�� ó������ ��������
                                If UBound(gPatTest) > 0 Then
                                    For intOrdCnt = 1 To UBound(gPatTest)
                                        If gPatTest(intOrdCnt) = gArrEQP(intTestCdCnt, 2) Then
                                            
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
                                            
                                            strCheck = "1"
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

                If strCheck <> "" Then
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
                            strTestCodeSub = gArrEQP(intCol - colSTATE, 17)
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
                End If
            End If

            .spdResult.RowHeight(-1) = 15
        
            '## DB�� �������
            If gHOSP.SAVEAUTO = "Y" Then
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
        Next
    End With

Exit Sub

ErrHandle:
    strErrMsg = ""
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "TCPRcvData_VISION" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show

End Sub


Private Sub Phase_Serial_HITACHI7060()
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
                        Call SerialRcvData_HITACHI7060
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

Private Sub Phase_Serial_UROMETER720()
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
                    Case "~"
                        RcvBuffer = ""
                        intPhase = 2
                    Case Else
                        RcvBuffer = RcvBuffer & BufChar
                End Select
            Case 2
            
                Select Case BufChar
                    Case "~"
                        Call SerialRcvData_UROMETER720
                        RcvBuffer = ""
                        intPhase = 1
                    Case Else
                        RcvBuffer = RcvBuffer & BufChar
                End Select
        End Select
    Next i

End Sub

Private Sub Phase_Serial_STAGO()
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
                        intPhase = 2
                        frmMain.comEqp.Output = ACK
                        SetRawData "[Tx]" & ACK
                    Case ACK
                        If strState = "Q" Then
'                            Call SendOrder_STAGO
                        End If
                End Select
            Case 2      '## Transfer Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        frmMain.comEqp.Output = ACK
                        SetRawData "[Tx]" & ACK
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
                        frmMain.comEqp.Output = ACK
                        SetRawData "[Tx]" & ACK
                End Select
            Case 4      '## Termination Phase
                Select Case BufChar
                    Case STX
                        intPhase = 2
                    Case EOT
                        lblTestDate.Caption = Format(Now, "yyyy-mm-dd")
                        Call SerialRcvData_STAGO
                        If strState = "Q" Then
                            intSndPhase = 1
                            intFrameNo = 1
                            frmMain.comEqp.Output = ENQ
                            SetRawData "[Tx]" & ENQ
                        End If
                        intPhase = 1
                End Select
        End Select
    Next i
            
End Sub

Private Sub Phase_Serial_ISMART300()
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
                        intPhase = 2
                        frmMain.comEqp.Output = ACK
                        SetRawData "[Tx]" & ACK
                    Case ACK
                        If strState = "Q" Then
'                            Call SendOrder_STAGO
                        End If
                End Select
            Case 2      '## Transfer Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        frmMain.comEqp.Output = ACK
                        SetRawData "[Tx]" & ACK
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
                        frmMain.comEqp.Output = ACK
                        SetRawData "[Tx]" & ACK
                End Select
            Case 4      '## Termination Phase
                Select Case BufChar
                    Case STX
                        intPhase = 2
                    Case EOT
                        lblTestDate.Caption = Format(Now, "yyyy-mm-dd")
                        Call SerialRcvData_ISMART300
                        If strState = "Q" Then
                            intSndPhase = 1
                            intFrameNo = 1
                            frmMain.comEqp.Output = ENQ
                            SetRawData "[Tx]" & ENQ
                        End If
                        intPhase = 1
                End Select
        End Select
    Next i
            
End Sub



Private Sub Phase_TCP_VISION()
    Dim strText As String
    Dim strTmp As String
    Dim i           As Long
    
    Dim strLastSeq  As String
    Dim strRcvSign  As String
    Dim strSendAck  As String
    Dim strRcvCnt   As String
    
    Dim strNS       As String
    Dim strNE       As String
    Dim intNS       As Integer
    Dim intNE       As Integer
    
    Dim strSendData  As String
    Dim varBuffers   As Variant
    
    varBuffers = Split(pBuffer, vbLf)
    strBuffer = ""
    
    For i = 0 To UBound(varBuffers)
        strBuffer = varBuffers(i)
        If strBuffer = "" Then
            Exit For
        End If
        strLastSeq = mGetP(strBuffer, 1, vbTab)
        strRcvSign = mGetP(strBuffer, 2, vbTab)
        
        strSendAck = strLastSeq & vbTab & "ACK"
        
        Select Case UCase(strRcvSign)
            Case "RESULT"
                    Call TCPRcvData_VISION
                    strBuffer = ""
                    
            Case "CONNECT"
                    
                    wSCK.SendData strSendAck & vbLf
                    SetRawData "[Tx]" & strSendAck & vbLf
                    
            Case "RESULTS"
                    
                    strRcvCnt = CInt(mGetP(strBuffer, 3, vbTab))
                    
                    strNS = strRcvCnt
                    strNE = mGetP(strBuffer, 4, vbTab)
                    
                    strNS = strNS - strNE
                    strNE = strNS + strNE
                    
                    strSendData = strLastSeq & vbTab & "GET" & vbTab & strNS & vbTab & strNE
                    
                    wSCK.SendData strSendData & vbLf
                    SetRawData "[Tx]" & strSendData & vbLf
                    
                    Call WritePrivateProfileString("config", "LASTSEQ", strRcvCnt, App.PATH & "\Interface.ini")
                    txtLastSeq.Text = strRcvCnt
                    
'                    blnResults = False
                    
        End Select
    Next
    

End Sub

Private Sub Phase_TCP_BC5800()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case BufChar
            Case STX
                'If blnIsSB = True Then
                '    lblTestDate.Caption = Format(Now, "yyyy-mm-dd")
                '    Call TCPRcvData_BC5800
                'End If
                'blnIsSB = False
'                intBufCnt = 1
'                Erase strRecvData
'                ReDim Preserve strRecvData(intBufCnt)
            Case SB '
                'intBufCnt = 1
                'Erase strRecvData
                'ReDim Preserve strRecvData(intBufCnt)
                RcvBuffer = ""
                blnIsSB = True
            Case EB

                'If blnIsSB = True Then
                    lblTestDate.Caption = Format(Now, "yyyy-mm-dd")
                    Call TCPRcvData_BC5800
                    RcvBuffer = ""
                'End If
                'blnIsSB = False

'                dtpToday.Value = Now
'                Call TCPRcvData_BC5800
'                Erase strRecvData
'                intBufCnt = 0
'                RcvBuffer = ""
'            Case vbCr
'                intBufCnt = intBufCnt + 1
'                ReDim Preserve strRecvData(intBufCnt)
'            Case vbLf
            
            
            Case Else
                RcvBuffer = RcvBuffer & BufChar
                'strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
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
    
End Sub

Private Sub Form_Load()
    Dim strTmp      As String
    Dim strSaveDt   As String
    Dim intCnt      As Integer
    
On Error GoTo ErrHandle
    
    Me.Width = 20940
    Me.Height = 12585

    'Me.Caption = Space(5) & gHOSP.HOSPNM & Space(5) & gHOSP.MACHNM & " �������̽�"
    Me.Caption = gHOSP.MACHNM & Space$(5) & "�¢¢¢¢�     [���� �������̽�]     �¢¢¢¢�"

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

    dtpFrom.Value = Now
    dtpTo.Value = Now
    
    imgNet1.ZOrder 0
    tmrDBConn.Interval = 500
    tmrDBConn.Enabled = True
    
    If gHOSP.MACHNM = "VISION" Then
        fraVISION.Visible = True
    Else
        fraVISION.Visible = False
    End If
    
    
    intPhase = 1

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
        
        If gComm.TCPTYPE = "SERVER" Then
            wSCK.LocalPort = CInt(gComm.TCPPORT)
            wSCK.Listen
            
            lblComStatus.Left = lblSend.Left
            lblComStatus.Width = lblComStatus.Width * 2
            
            lblComStatus.Caption = "TCP " & gComm.TCPPORT & " ������..."
    
            imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
            imgSend.Visible = False
            imgReceive.Visible = False
            lblSend.Visible = False
            lblRcv.Visible = False
    
        Else
            wSCK.Close
            wSCK.Connect gComm.TCPIP, CInt(gComm.TCPPORT)
    
            lblComStatus.Left = lblSend.Left
            lblComStatus.Width = lblComStatus.Width * 2
            
            lblComStatus.Caption = "TCP " & gComm.TCPIP & ":" & gComm.TCPPORT & " ������..."
    
            imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
            imgSend.Visible = False
            imgReceive.Visible = False
            lblSend.Visible = False
            lblRcv.Visible = False
        End If
    End If
    
End Sub


Private Sub Form_Resize()

    On Error Resume Next

    If Me.ScaleHeight = 0 Then Exit Sub

    Me.Top = 0

    'fraWorkInfo.Left
    'fraWorkInfo.Height = Me.ScaleHeight - picHeader.Height - picBottom.Height - spdWork.Top - 300
    spdWork.Height = Me.ScaleHeight - picHeader.Height - picBottom.Height - fraWorkInfo.Height - 300
    
    spdOrder.Top = spdOrder.Top + 40
    spdOrder.Width = Me.ScaleWidth - spdWork.Width - spdResult.Width - 200
    spdOrder.Height = Me.ScaleHeight - picHeader.Height - picBottom.Height - 300
    
    spdResult.Left = spdOrder.Left + spdOrder.Width + 50
    spdResult.Height = Me.ScaleHeight - picHeader.Height - picBottom.Height - fraPatInfo.Height - 300

    fraPatInfo.Left = spdOrder.Left + spdOrder.Width + 50
    fraPatInfo.Height = Me.ScaleHeight - picHeader.Height - picBottom.Height - spdResult.Height - 300

    
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
        wSCK.Close
        wSCK.LocalPort = CInt(gComm.TCPPORT)
        wSCK.Listen

        lblComStatus.Caption = "TCP " & gComm.TCPPORT & " ��Ʈ ���Ἲ��"
    Else
        wSCK.Close
        wSCK.Connect gComm.TCPIP, CInt(gComm.TCPPORT)

        lblComStatus.Caption = "TCP " & gComm.TCPIP & ":" & gComm.TCPPORT & " ��Ʈ ���Ἲ��"
    End If

End Sub

Private Sub wSCK_ConnectionRequest(ByVal requestID As Long)
            
    If wSCK.State <> sckClosed Then
        wSCK.Close

        wSCK.Accept requestID
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
    
    wSCK.GetData strText
    SetRawData "[Rx]" & strText
    
    pBuffer = strText
    
    If Len(pBuffer) > 0 Then
        Select Case UCase(gHOSP.MACHNM)
            Case "BC5800":                  Call Phase_TCP_BC5800
            Case "VISION":                  Call Phase_TCP_VISION
        
        End Select
    End If

End Sub
