VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
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
      Height          =   1425
      Left            =   60
      TabIndex        =   44
      Top             =   660
      Width           =   5895
      Begin VB.TextBox txtStopNum 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2760
         TabIndex        =   52
         Text            =   "009999"
         Top             =   630
         Width           =   1365
      End
      Begin VB.TextBox txtStartNum 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1170
         TabIndex        =   51
         Text            =   "000001"
         Top             =   630
         Width           =   1365
      End
      Begin VB.CommandButton cmdAll 
         BackColor       =   &H00FFFFFF&
         Caption         =   ">>"
         Height          =   705
         Left            =   5250
         Style           =   1  '�׷���
         TabIndex        =   50
         Top             =   210
         Width           =   495
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��ũ��ȸ"
         Height          =   705
         Left            =   4200
         Style           =   1  '�׷���
         TabIndex        =   49
         Top             =   210
         Width           =   975
      End
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   315
         Left            =   1170
         TabIndex        =   45
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
         Format          =   137166849
         CurrentDate     =   40457
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   315
         Left            =   2760
         TabIndex        =   46
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
         Format          =   137166849
         CurrentDate     =   40457
      End
      Begin VB.Label lblFileNm 
         Appearance      =   0  '���
         BackColor       =   &H80000005&
         BackStyle       =   0  '����
         Caption         =   "Label8"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   1170
         TabIndex        =   59
         Top             =   1020
         Width           =   4485
      End
      Begin VB.Label Label1 
         Appearance      =   0  '���
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '����
         Caption         =   "�����̸� :"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   5
         Left            =   150
         TabIndex        =   58
         Top             =   1020
         Width           =   930
      End
      Begin VB.Label Label1 
         Appearance      =   0  '���
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '����
         Caption         =   "������ȣ :"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   4
         Left            =   150
         TabIndex        =   57
         Top             =   660
         Width           =   930
      End
      Begin VB.Label Label1 
         Appearance      =   0  '���
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '����
         Caption         =   "~"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   3
         Left            =   2580
         TabIndex        =   53
         Top             =   690
         Width           =   150
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
         Left            =   2580
         TabIndex        =   48
         Top             =   330
         Width           =   150
      End
      Begin VB.Label Label1 
         Appearance      =   0  '���
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '����
         Caption         =   "�������� :"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   150
         TabIndex        =   47
         Top             =   330
         Width           =   930
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
         Left            =   11910
         TabIndex        =   33
         Top             =   60
         Width           =   975
      End
   End
   Begin VB.Frame fraPatInfo 
      BackColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   16440
      TabIndex        =   19
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
      Left            =   19350
      TabIndex        =   17
      Top             =   9150
      Visible         =   0   'False
      Width           =   6255
      Begin VB.CommandButton cmdWork 
         Caption         =   "��ũ��ȸ"
         Height          =   315
         Left            =   960
         TabIndex        =   20
         Top             =   270
         Width           =   1425
      End
      Begin VB.CommandButton cmdResultSearch 
         Caption         =   "�����ȸ"
         Height          =   315
         Left            =   2610
         TabIndex        =   18
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
         Left            =   8010
         TabIndex        =   12
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
         SpreadDesigner  =   "frmMain.frx":2BB0
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
         Picture         =   "frmMain.frx":2F8C
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgSend 
         Height          =   240
         Left            =   4845
         Picture         =   "frmMain.frx":3516
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgReceive 
         Height          =   240
         Left            =   5760
         Picture         =   "frmMain.frx":3AA0
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
         Picture         =   "frmMain.frx":402A
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgNet2 
         Height          =   240
         Left            =   390
         Picture         =   "frmMain.frx":4174
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgNet3 
         Height          =   240
         Left            =   390
         Picture         =   "frmMain.frx":42BE
         Top             =   180
         Width           =   240
      End
      Begin VB.Label lblComStatus 
         Alignment       =   2  '��� ����
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
         TabIndex        =   7
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
      TabIndex        =   3
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
      SpreadDesigner  =   "frmMain.frx":4408
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
      Begin VB.FileListBox FileAllo 
         Height          =   270
         Left            =   18570
         Pattern         =   "*.patlist"
         TabIndex        =   56
         Top             =   150
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Timer tmrAllo 
         Left            =   17850
         Top             =   60
      End
      Begin VB.CommandButton cmdOrder 
         Caption         =   "��������"
         Enabled         =   0   'False
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
         Left            =   10260
         TabIndex        =   55
         ToolTipText     =   "����ȭ���� ��� ����ϴ�"
         Top             =   150
         Width           =   1425
      End
      Begin VB.CommandButton cmdResult 
         Caption         =   "����ޱ�"
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
         Left            =   11790
         TabIndex        =   54
         ToolTipText     =   "������ ����� EMR������ �����մϴ�"
         Top             =   150
         Width           =   1425
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
         Left            =   13320
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
         Left            =   14850
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
         Top             =   150
         Width           =   555
      End
      Begin MSCommLib.MSComm comEqp 
         Left            =   20970
         Top             =   30
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
         RThreshold      =   1
         EOFEnable       =   -1  'True
      End
      Begin MSComDlg.CommonDialog AlloFile 
         Left            =   16620
         Top             =   60
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Shape Shape7 
         BorderColor     =   &H0080FFFF&
         Height          =   405
         Left            =   10200
         Top             =   90
         Width           =   1545
      End
      Begin VB.Shape Shape12 
         BorderColor     =   &H0080FFFF&
         Height          =   405
         Left            =   11730
         Top             =   90
         Width           =   1545
      End
      Begin VB.Shape Shape8 
         BorderColor     =   &H0080FFFF&
         Height          =   405
         Left            =   13260
         Top             =   90
         Width           =   1545
      End
      Begin VB.Shape Shape9 
         BorderColor     =   &H0080FFFF&
         Height          =   405
         Left            =   14790
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
         TabIndex        =   5
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
         TabIndex        =   4
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
         TabIndex        =   2
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
      TabIndex        =   43
      Top             =   2130
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
      SpreadDesigner  =   "frmMain.frx":5026
      ScrollBarTrack  =   3
      ShowScrollTips  =   3
   End
   Begin FPSpread.vaSpread spdOrder 
      Height          =   7935
      Left            =   5970
      TabIndex        =   60
      Top             =   720
      Width           =   16335
      _Version        =   393216
      _ExtentX        =   28813
      _ExtentY        =   13996
      _StockProps     =   64
      ColsFrozen      =   22
      EditEnterAction =   2
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
      SpreadDesigner  =   "frmMain.frx":71FA
      ScrollBarTrack  =   3
      ShowScrollTips  =   3
   End
   Begin VB.Menu mnuMenu01 
      Caption         =   "  ����  "
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
      Caption         =   "  ��ȸ  "
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

Private Sub cmdOrder_Click()
'    Dim objAccInfo As clsIISAccInfo     '�������� Ŭ����
'    Dim objResult   As clsIISResult     '������� Ŭ����
'    Dim objQCResult As clsIISQCResult   'QC������� Ŭ����
'    Dim mLogOn      As clsIISLogOn
    Dim strAlloFile As String
    Dim lngFIleNum  As Long
    Dim strInFo     As String
    Dim strOldInFo  As String
    'Dim iCnt        As Integer
    Dim intRow      As Integer
    Dim varTmp      As Variant
    Dim OrderPath   As String
    Dim i           As Integer
    Dim strOrgBarno    As String
    Dim strBarno    As String
    Dim strPtID     As String
    Dim strDeptNm   As String
    Dim strPtNm     As String
    Dim strINFD     As String

'    Dim strBarNo    As String
    Dim varBuffer   As Variant
    Dim intCnt As Integer
    Dim strBuf      As String
    Dim RS          As ADODB.Recordset
    Dim Rs1 As New Recordset
    Dim strKey As String
    Dim strTemp As String
    Dim J As Integer
    
    varBuffer = Split(pBuffer, vbCr)
    
    AlloFile.Filename = lblFileNm.Caption
    AlloFile.CancelError = True
    
    J = 0
    
    '-----------------------------------------------------------------------------
    'NAME: AdvanSure
    'TYPE: Patient List Files
    '-----------------------------------------------------------------------------
    'ID   NAME    PANEL   A   B   AGE     GENDER      ADDR    CONTACT     BIRTH   RRN     CLIENT      INSPECTOR   HOSPITAL    EXAMDATE
    'C1600707952 ����  2   1   1                           ��������            2016-06-10
    'C1600707952 ����  1   1   1                           ��������            2016-06-10
    'C1600707962 �׽�Ʈ  1   1   1                           ��������            2016-06-10
    'C1600707972 ���ֱ�  2   1   1                           ��������            2016-06-10
    'C1600707972 ���ֱ�  1   1   1                           ��������            2016-06-10
    
    For intCnt = 0 To UBound(varBuffer) - 1
        strBuf = varBuffer(intCnt)
        If Mid(strBuf, 1, 2) = "--" Or Mid(strBuf, 1, 2) = "NA" Or Mid(strBuf, 1, 2) = "TY" Or Mid(strBuf, 1, 2) = "ID" Then
    
        Else
            J = J + 1
            
            strOrgBarno = Mid(strBuf, 1, gHOSP.BARLEN)
            
            If Mid(strBuf, gHOSP.BARLEN + 2, 10) <> "" Then
                lblFileNm.Caption = ""
                Exit Sub
            End If
'            If Len(strBuf) > 13 Then
'                Close #lngFIleNum
'                pBuffer = ""
'                Exit Sub
'            End If
            
            If J = 1 Then
                If Len(Dir(AlloFile.Filename)) Then
                     Close #lngFIleNum
                     Kill AlloFile.Filename
                End If
                lngFIleNum = FreeFile
                        
                Open AlloFile.Filename For Append As #lngFIleNum
            
                '-- �������ϸ� ���� : PatList 2012-01-16.patlist
                'OrderPath = GetAlloConfig("OrderPath")
            
                Print #lngFIleNum, "-----------------------------------------------------------------------------"
                Print #lngFIleNum, "NAME: AdvanSure"
                Print #lngFIleNum, "TYPE: Patient List Files"
                Print #lngFIleNum, "-----------------------------------------------------------------------------"
                Print #lngFIleNum, "ID " & vbTab & " NAME " & vbTab & " PANEL " & vbTab & " A " & vbTab & " B " & vbTab & " AGE " & vbTab & " GENDER " & vbTab & " ADDR " & vbTab & _
                                   " CONTACT " & vbTab & " BIRTH " & vbTab & " RRN " & vbTab & " CLIENT " & vbTab & " INSPECTOR " & vbTab & " HOSPITAL " & vbTab & " EXAMDATE"
            End If
                        
            '-- ��������
            With mOrder
                .OrgBarNo = strOrgBarno
                
                SQL = "SELECT fn_ack_get_bcno_normal('" & strOrgBarno & "') as BCD FROM DUAL"
                Set RS = AdoCn.Execute(SQL, , 1)
                If Not RS.EOF = True And Not RS.BOF = True Then
                    strBarno = Trim(RS.Fields("BCD")) & ""
                End If
                RS.Close
                
                .BarNo = strBarno
                
                Call GetOrder_ALLOSTATION(strBarno, gHOSP.RSTTYPE)
                
            End With
        End If
    Next
    
    'MsgBox "1"
    
    For intRow = 1 To spdOrder.MaxRows
        spdOrder.Row = intRow
        spdOrder.Col = colCHECKBOX
        If spdOrder.Value = "1" And GetText(spdOrder, intRow, colSTATE) = "ORDER" Then
            strPtID = ""
            strDeptNm = ""
            strPtNm = ""
            strBarno = ""
            strINFD = ""
            
            strPtID = GetText(spdOrder, intRow, colPID)
            'strDeptNm = GetText(spdOrder, intRow, colSTATE)
            strPtNm = GetText(spdOrder, intRow, colPNAME)
            'strBarno = GetText(spdOrder, intRow, colBARCODE)
            strBarno = GetText(spdOrder, intRow, colRACKNO)
            strINFD = GetText(spdOrder, intRow, colINOUT)
            
            If strBarno = "" Then
                Exit For
            End If
                
            '-- ȣ���
            If strINFD = "IN" Then
                strInFo = "1"
                Print #lngFIleNum, strBarno & vbTab & strPtNm & vbTab & strInFo & vbTab & "1" & vbTab & "1" & vbTab & "" & vbTab & "" & vbTab & vbTab & vbTab & vbTab & "" & vbTab & strDeptNm & vbTab & "" & vbTab & vbTab & Format(Now, "yyyy-mm-dd")
            '-- ����
            ElseIf strINFD = "FD" Then
                strInFo = "2"
                Print #lngFIleNum, strBarno & vbTab & strPtNm & vbTab & strInFo & vbTab & "1" & vbTab & "1" & vbTab & "" & vbTab & "" & vbTab & vbTab & vbTab & vbTab & "" & vbTab & strDeptNm & vbTab & "" & vbTab & vbTab & Format(Now, "yyyy-mm-dd")
            '-- ������
            'ElseIf varTmp(i) = "AT" Then
            '    strInFo = "3"
            '    Print #lngFIleNum, strBarNo & vbTab & strPtNm & vbTab & strInFo & vbTab & "1" & vbTab & "0" & vbTab & "" & vbTab & "" & vbTab & vbTab & vbTab & vbTab & "" & vbTab & strDeptNm & vbTab & "" & vbTab & vbTab & Format(Now, "yyyy-mm-dd")
            End If
            spdOrder.SetText 1, intRow, "SEND"
        End If
    Next
        
    'MsgBox "���� ���� ���� �Ϸ�", vbOKOnly + vbInformation, Me.Caption
        
    Close #lngFIleNum
    
    pBuffer = ""
    
End Sub

Private Sub cmdRcv_Click()
        
    pBuffer = txtRcv.Text
    
    Select Case UCase(gHOSP.MACHNM)
                
        Case "AU680":           Call Phase_Serial_AU680
        Case "HITACHI7180":     Call Phase_Serial_HITACHI7180
                    
    End Select

    pBuffer = ""
    
End Sub

Private Sub cmdRcvClear_Click()
    
    txtRcv.Text = ""
    
End Sub

Private Sub cmdResult_Click()

    tmrAllo.Enabled = False
    
    If Not Set_DbConnect_Jet Then
        MsgBox "AlloScan �����ͺ��̽� ��θ� Ȯ���ϼ���", vbOKOnly + vbCritical, Me.Caption
'        End
    Else
        Call Get_SearchListIN
        Call Get_SearchListFD
    End If
    
    tmrAllo.Enabled = True
    
End Sub

'   Access DB Connect
Public Function Set_DbConnect_Jet() As Boolean
    Dim DB_Name         As String
    Dim UserName        As String
    Dim Password        As String
    Dim blnWinNTAuth    As Boolean
    Dim strSrcfile      As String
    Dim strDestFile     As String
    
    Set AdoCn_Allo = New ADODB.Connection

On Error GoTo ConnectError

    DB_Name = "C:\Program Files (x86)\LG Life Sciences\AdvanSure AlloStationSmart\DB\DBUnit.mdb" 'GetAlloConfig("MDBPath")      'C:\Program Files\LG Life Sciences\AdvanSure AlloStationSmart\DB
    UserName = "admin"
    Password = "reader_admin"
    
    
    If (DB_Name = "") Or (UserName = "") Then
        Set_DbConnect_Jet = False
        Set AdoCn_Allo = Nothing
        Exit Function
    End If
        
    With AdoCn_Allo
        .ConnectionTimeout = 25
        .CursorLocation = adUseClient
        .Provider = "Microsoft.Jet.OLEDB.4.0"
        .Properties("Mode").Value = adModeReadWrite
        .Properties("Persist Security Info").Value = False
        .Properties("Data Source").Value = DB_Name
        .Properties("User ID").Value = UserName
        .Properties("Jet OLEDB:Database Password").Value = Password
        .Properties("Jet OLEDB:Compact Without Replica Repair").Value = True
        .Open
    End With

    Set_DbConnect_Jet = True
    
 Exit Function

ConnectError:
    '   ����ó��
    MsgBox "   Error No. : " & Err.Number & vbCrLf & _
           " Description : " & Err.Description & vbCrLf & _
           "      Source : " & Err.Source & vbCrLf & vbCrLf _
           , vbCritical, " DB Open Error"

    If AdoCn_Allo.State <> adStateOpen Then
        Set_DbConnect_Jet = False
        Set AdoCn_Allo = Nothing
    End If
    
End Function

Private Sub Get_SearchListIN()
    Dim intRow      As Long
    Dim strAge      As String
    Dim strTransDt  As String
    Dim strHMsg     As String
    Dim strDMsg     As String
    
    Dim vWorkNo      As Variant  'Spread�� WorkNo
    Dim vBarNo       As Variant  'Spread�� ���ڵ��ȣ
    Dim strRcvBuf    As String   '������ Data
    Dim strType      As String   '������ Record Type
    Dim strOrgBarno  As String   '������ BarNO
    Dim strBarno     As String   '������ BarNO
    Dim strWorkNo    As String   '������ WorkNo
    Dim strIntResult As String   '������ �˻���
    Dim strIntBase   As String   '������ ������ �˻��
    Dim strResult    As String   'LIS �˻���
    Dim strMachResult   As String   '������ �����
    Dim strClass     As String   'Class �������
    Dim strClassA    As String   'Class �������
    Dim strMaxClass  As String   'Class �������
    
    Dim strTemp      As String
    Dim i            As Long
    Dim blnSameBar   As Boolean
    Dim intCnt       As Integer
    Dim strTIgE      As String   'Total IgE
    Dim strSIgE      As String   'Ŭ������ ���� ���� ��
    Dim strRemark    As String
    Dim SIgE         As String
    
    Dim Y, Y1, Y2, Y3, X1, X2
    

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
    Dim intCol          As Integer  '����÷� ����
    
    
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    Dim Res             As Integer
    Dim RS          As ADODB.Recordset

    On Error GoTo DBErr
    
    Set AdoRs_Allo = Get_ResultList_IN
    
    If Not AdoRs_Allo.BOF Then
        intRow = 1
        strTransDt = ""
        mResult.BarNo = ""
        
        Do Until AdoRs_Allo.EOF
            strOrgBarno = AdoRs_Allo.Fields("PATIENTID").Value & ""
                        
            '-- �������
            mOrder.OrgBarNo = strOrgBarno
            With mResult
                .Kind = "IN"
                .BarNo = strOrgBarno
                .RsltDate = Format(Now, "yyyy-mm-dd")
                .RsltTime = Format(Now, "hh:mm:ss")
                .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
            End With
            
            Call SetPatInfo(strOrgBarno, gHOSP.RSTTYPE)
            
            strIntBase = ""
            
            For intCnt = 2 To 32
                If intCnt = 2 Then
                    strIntBase = "IN"
                End If
                
                strIntResult = ""
                strResult = ""
                strTemp = AdoRs_Allo.Fields("BANDVAL" & intCnt) & ""
                '> 100|Valid|PC|158.24
                '1.31|2|d70|1.31
                
                If strTemp <> "" Then
                    strIntBase = mResult.Kind & mGetP(strTemp, 3, "|")
                    strIntResult = mGetP(strTemp, 1, "|")
                End If
                
                If strIntBase = "INtIgE" Or strIntBase = "FDtIgE" Then
                    If IsNumeric(strIntResult) Then
                        If strIntResult > 100 Then
                            strIntResult = ">100"
                        Else ' strIntResult <= 100 Then
                            strIntResult = "=<100"
                        End If
                    Else
                        If Trim(strIntResult) = "> 100" Then
                            strIntResult = ">100"
                        ElseIf Trim(strIntResult) = "�� 100" Then
                            strIntResult = "��100"
                        Else
                            strIntResult = "��100"
                        End If
                    End If
                Else
                    strIntResult = mGetP(strTemp, 2, "|")
                    If strIntResult = "0" Then
                        strIntResult = "Nondetect"
                    Else
                        strIntResult = "Class " & strIntResult
                    End If
                End If
                strResult = strIntResult
                
                If strIntBase <> "" And strResult <> "" Then
                    blnSame = False
                    '-- �˻縶���� ���� ��������
                    For intTestNmCnt = 1 To UBound(gArrEQPNm)
                        '-- ���ä���� ����...
                        If strIntBase = gArrEQPNm(intTestNmCnt, 3) Then
                            strCheck = "0"
                            strSeqno = gArrEQPNm(intTestNmCnt, 1)
                            strState = ""
                            For intTestCdCnt = 1 To UBound(gArrEQP)
                                '-- �˻��ڵ嵵 ���ٸ�...
                                If strIntBase = gArrEQP(intTestCdCnt, 3) Then
                                    strTestCode = gArrEQP(intTestCdCnt, 2)
                                    strTestName = gArrEQP(intTestCdCnt, 5)
                                    intResPrecUse = gArrEQP(intTestCdCnt, 7)
                                    intResPrec = gArrEQP(intTestCdCnt, 8)
                                    If UBound(gPatTest) > 0 Then
                                        For intOrdCnt = 1 To UBound(gPatTest)
                                            If gPatTest(intOrdCnt) = gArrEQP(intTestCdCnt, 2) Then
                                                strCheck = "1"
                                                
                                                strOrderCode = gArrEQP(intTestCdCnt, 16)
                                                strTestCodeSub = gArrEQP(intTestCdCnt, 17)
                                                
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
                    intRstRow = spdResult.DataRowCnt + 1
                    If spdResult.MaxRows < intRstRow Then
                        spdResult.MaxRows = intRstRow
                    End If

                    '-- ������� ǥ��("���")
                    SetText spdOrder, "���", gRow, colSTATE

                    '-- ����ȭ�� ����� ǥ��
                    For intCol = colSTATE + 1 To spdOrder.MaxCols
                        If strTestName = gArrEQPNm(intCol - colSTATE, 7) Then
                            SetText spdOrder, strResult, gRow, intCol
                            Exit For
                        End If
                    Next

                    '-- ��� List
                    SetText spdResult, strCheck, intRstRow, colRCHECKBOX               'üũ
                    SetText spdResult, strSeqno, intRstRow, colRSEQNO                  '����
                    SetText spdResult, strOrderCode, intRstRow, colRORDERCD            'ó���ڵ�
                    SetText spdResult, strTestCode, intRstRow, colRTESTCD          '�˻��ڵ�
                    SetText spdResult, strTestCodeSub, intRstRow, colRSUBCD        '�˻��ڵ�SUB
                    SetText spdResult, strTestName, intRstRow, colRTESTNM              '�˻��
                    SetText spdResult, strIntBase, intRstRow, colRCHANNEL              '���ä��
                    SetText spdResult, strMachResult, intRstRow, colRMACHRESULT        '�����
                    SetText spdResult, strResult, intRstRow, colRLISRESULT             'LIS���
                    SetText spdResult, strJudge, intRstRow, colRJUDGE                  '����
                    SetText spdResult, strLow & "~" & strHigh, intRstRow, colRREF      '����ġ

                    '-- ���� ����
                    Call SetLocalDB(gRow, intRstRow, "1", "")

                End If
            Next
            
'            spdResult.RowHeight(-1) = 15
'
'            '## DB�� �������
'            If gHOSP.SAVEAUTO = "Y" And strState = "R" Then
'                Res = SaveTransData(gRow, spdOrder)
'
'                If Res = -1 Then
'                    '-- ���� ����
'                    SetForeColor spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
'                    SetText spdOrder, "�������", gRow, colSTATE
'                Else
'                    '-- ���� ����
'                    SetBackColor spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
'                    SetText spdOrder, "����Ϸ�", gRow, colSTATE
'                    SetText spdOrder, "0", gRow, colCHECKBOX
'
'                          SQL = "Update PATRESULT Set " & vbCrLf
'                    SQL = SQL & " sendflag = '2' " & vbCrLf
'                    SQL = SQL & " Where equipno = '" & gHOSP.HOSPCD & "' " & vbCrLf
'                    SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
'                    SQL = SQL & "   And barcode = '" & Trim(GetText(spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
'                    SQL = SQL & "   And saveseq = " & Trim(GetText(spdOrder, gRow, colSAVESEQ)) & vbCrLf
'
'                    If DBExec(AdoCn_Local, SQL) Then
'                        '-- ����
'                    End If
'                End If
'                strState = ""
'            End If
'
'            'intRow = intRow + 1
'
            
            Set AdoRs_Allo2 = Get_ResultList_ST(strOrgBarno)
            
            If Not AdoRs_Allo2.BOF Then
                'intRow = 1
                strTransDt = ""
                mResult.BarNo = ""
                
                Do Until AdoRs_Allo2.EOF
                    
                    strIntBase = ""
                    
                    For intCnt = 2 To 32
                        If intCnt = 2 Then
                            strIntBase = "IN"
                        End If
                        
                        strIntResult = ""
                        strResult = ""
                        strTemp = AdoRs_Allo2.Fields("BANDVAL" & intCnt) & ""
                        '> 100|Valid|PC|158.24
                        '1.31|2|d70|1.31
                        
                        If strTemp <> "" Then
                            strIntBase = mResult.Kind & mGetP(strTemp, 3, "|")
                            strIntResult = mGetP(strTemp, 1, "|")
                        End If
                        
                        If strIntBase = "INtIgE" Or strIntBase = "FDtIgE" Then
                            If IsNumeric(strIntResult) Then
                                If strIntResult > 100 Then
                                    strIntResult = ">100"
                                Else ' strIntResult <= 100 Then
                                    strIntResult = "=<100"
                                End If
                            Else
                                If Trim(strIntResult) = "> 100" Then
                                    strIntResult = ">100"
                                ElseIf Trim(strIntResult) = "�� 100" Then
                                    strIntResult = "��100"
                                Else
                                    strIntResult = "��100"
                                End If
                            End If
                        Else
                            strIntResult = mGetP(strTemp, 2, "|")
                            If strIntResult = "0" Then
                                strIntResult = "Nondetect"
                            Else
                                strIntResult = "Class " & strIntResult
                            End If
                        End If
                        strResult = strIntResult
                        
                        If strIntBase <> "" And strResult <> "" Then
                            blnSame = False
                            '-- �˻縶���� ���� ��������
                            For intTestNmCnt = 1 To UBound(gArrEQPNm)
                                '-- ���ä���� ����...
                                If strIntBase = gArrEQPNm(intTestNmCnt, 3) Then
                                    strCheck = "0"
                                    strSeqno = gArrEQPNm(intTestNmCnt, 1)
                                    strState = ""
                                    For intTestCdCnt = 1 To UBound(gArrEQP)
                                        '-- �˻��ڵ嵵 ���ٸ�...
                                        If strIntBase = gArrEQP(intTestCdCnt, 3) Then
                                            strTestCode = gArrEQP(intTestCdCnt, 2)
                                            strTestName = gArrEQP(intTestCdCnt, 5)
                                            intResPrecUse = gArrEQP(intTestCdCnt, 7)
                                            intResPrec = gArrEQP(intTestCdCnt, 8)
                                            If UBound(gPatTest) > 0 Then
                                                For intOrdCnt = 1 To UBound(gPatTest)
                                                    If gPatTest(intOrdCnt) = gArrEQP(intTestCdCnt, 2) Then
                                                        strCheck = "1"
                                                        
                                                        strOrderCode = gArrEQP(intTestCdCnt, 16)
                                                        strTestCodeSub = gArrEQP(intTestCdCnt, 17)
                                                        
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
                            intRstRow = spdResult.DataRowCnt + 1
                            If spdResult.MaxRows < intRstRow Then
                                spdResult.MaxRows = intRstRow
                            End If
        
                            '-- ������� ǥ��("���")
                            SetText spdOrder, "���", gRow, colSTATE
        
                            '-- ����ȭ�� ����� ǥ��
                            For intCol = colSTATE + 1 To spdOrder.MaxCols
                                If strTestName = gArrEQPNm(intCol - colSTATE, 7) Then
                                    SetText spdOrder, strResult, gRow, intCol
                                    Exit For
                                End If
                            Next
        
                            '-- ��� List
                            SetText spdResult, strCheck, intRstRow, colRCHECKBOX               'üũ
                            SetText spdResult, strSeqno, intRstRow, colRSEQNO                  '����
                            SetText spdResult, strOrderCode, intRstRow, colRORDERCD            'ó���ڵ�
                            SetText spdResult, strTestCode, intRstRow, colRTESTCD          '�˻��ڵ�
                            SetText spdResult, strTestCodeSub, intRstRow, colRSUBCD        '�˻��ڵ�SUB
                            SetText spdResult, strTestName, intRstRow, colRTESTNM              '�˻��
                            SetText spdResult, strIntBase, intRstRow, colRCHANNEL              '���ä��
                            SetText spdResult, strMachResult, intRstRow, colRMACHRESULT        '�����
                            SetText spdResult, strResult, intRstRow, colRLISRESULT             'LIS���
                            SetText spdResult, strJudge, intRstRow, colRJUDGE                  '����
                            SetText spdResult, strLow & "~" & strHigh, intRstRow, colRREF      '����ġ
        
                            '-- ���� ����
                            Call SetLocalDB(gRow, intRstRow, "1", "")
        
                        End If
                    Next
                    
                    spdResult.RowHeight(-1) = 15
        
                    '## DB�� �������
                    If gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow, spdOrder)
        
                        If Res = -1 Then
                            '-- ���� ����
                            SetForeColor spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText spdOrder, "�������", gRow, colSTATE
                        Else
                            '-- ���� ����
                            SetBackColor spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText spdOrder, "����Ϸ�", gRow, colSTATE
                            SetText spdOrder, "0", gRow, colCHECKBOX
        
                                  SQL = "Update PATRESULT Set " & vbCrLf
                            SQL = SQL & " sendflag = '2' " & vbCrLf
                            SQL = SQL & " Where equipno = '" & gHOSP.HOSPCD & "' " & vbCrLf
                            SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And barcode = '" & Trim(GetText(spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                            SQL = SQL & "   And saveseq = " & Trim(GetText(spdOrder, gRow, colSAVESEQ)) & vbCrLf
        
                            If DBExec(AdoCn_Local, SQL) Then
                                '-- ����
                            End If
                        End If
                        strState = ""
                    End If
                    
                    intRow = intRow + 1
                    AdoRs_Allo2.MoveNext
                Loop
            End If
            
            Set AdoRs_Allo2 = Nothing
            
            AdoRs_Allo.MoveNext
        Loop
    End If
    
    Set AdoRs_Allo = Nothing

Exit Sub

DBErr:
    Screen.MousePointer = 0
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "Get_SearchListIN" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Sub

'''Private Sub Get_SearchListFD()
'''    Dim intRow      As Long
'''    Dim strAge      As String
'''    Dim strTransDt  As String
'''    Dim strHMsg     As String
'''    Dim strDMsg     As String
'''
'''    Dim vWorkNo      As Variant  'Spread�� WorkNo
'''    Dim vBarNo       As Variant  'Spread�� ���ڵ��ȣ
'''    Dim strRcvBuf    As String   '������ Data
'''    Dim strType      As String   '������ Record Type
'''    Dim strBarno     As String   '������ BarNO
'''    Dim strWorkNo    As String   '������ WorkNo
'''    Dim strIntResult As String   '������ �˻���
'''    Dim strIntBase   As String   '������ ������ �˻��
'''    Dim strResult    As String   'LIS �˻���
'''    Dim strMachResult   As String   '������ �����
'''    Dim strClass     As String   'Class �������
'''    Dim strClassA    As String   'Class �������
'''    Dim strMaxClass  As String   'Class �������
'''
'''    Dim strTemp      As String
'''    Dim i            As Long
'''    Dim blnSameBar   As Boolean
'''    Dim intCnt       As Integer
'''    Dim strTIgE      As String   'Total IgE
'''    Dim strSIgE      As String   'Ŭ������ ���� ���� ��
'''    Dim strRemark    As String
'''    Dim SIgE         As String
'''
'''    Dim Y, Y1, Y2, Y3, X1, X2
'''
'''
'''    '������ ����
'''    Dim strCheck        As String   '�˻����üũ
'''    Dim strSeqno        As String   '�˻����
'''    Dim strOrderCode    As String   'ó���ڵ�
'''    Dim strTestName     As String   '�˻��ڵ�
'''    Dim strTestCode     As String   '�˻��ڵ�
'''    Dim strTestCodeSub  As String   '�˻��ڵ�SUB
'''    Dim intResPrecUse   As Integer  '�Ҽ�����ȯ����
'''    Dim intResPrec      As Integer  '�Ҽ����ڸ���
'''    Dim strResType      As String   '�Ҽ�����ȯ����
'''    Dim strLow          As String
'''    Dim strHigh         As String
'''    Dim strJudge        As String   '�������
'''    Dim strPrevRslt     As String   '�������
'''
'''    Dim intRstRow       As String   '����������� ���� Row
'''    Dim intCol          As Integer  '����÷� ����
'''
'''
'''    Dim intTestNmCnt    As Integer
'''    Dim intTestCdCnt    As Integer
'''    Dim intOrdCnt       As Integer
'''    Dim blnSame         As Boolean
'''    Dim Res             As Integer
'''    Dim strOrgBarno  As String   '������ BarNO
'''    Dim Rs          As ADODB.Recordset
'''
'''    On Error GoTo DBErr
'''
'''    Set AdoRs_Allo = Get_ResultList_FD
'''
'''    If Not AdoRs_Allo.BOF Then
'''        intRow = 1
'''        strTransDt = ""
'''        mResult.BarNo = ""
'''
'''        Do Until AdoRs_Allo.EOF
'''            'If mResult.BarNo = "" And strBarno <> AdoRs_allo.Fields("PATIENTID").Value Then
'''                strOrgBarno = AdoRs_Allo.Fields("PATIENTID").Value
'''
'''                '-- �������
'''                With mResult
'''                    .Kind = "FD"
'''                    .BarNo = strOrgBarno
'''                    .RsltDate = Format(Now, "yyyy-mm-dd")
'''                    .RsltTime = Format(Now, "hh:mm:ss")
'''                    .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
'''                End With
'''
'''''                           If Len(strOrgBarno) = 11 Then
''''                    SQL = "SELECT fn_ack_get_bcno_normal(" & strOrgBarno & ") as BCD FROM DUAL"
''''                    Set Rs = AdoCn.Execute(SQL, , 1)
''''                    If Not Rs.EOF = True And Not Rs.BOF = True Then
''''                        strBarno = Trim(Rs.Fields("BCD")) & ""
''''                        mResult.BarNo = strBarno
''''                    End If
''''                    Rs.Close
''''                End If
'''
'''                Call SetPatInfo(strOrgBarno, gHOSP.RSTTYPE)
'''            'End If
'''
'''            strSIgE = "0"
'''            strMaxClass = "0"
'''            strClassA = "0"
'''            strIntBase = ""
'''
'''            For intCnt = 2 To 32
'''                If intCnt = 2 Then
'''                    strIntBase = "FD"
'''                End If
'''
'''                strIntResult = ""
'''                strTemp = AdoRs_Allo.Fields("BANDVAL" & intCnt) & ""
'''                If strTemp <> "" Then
'''                    strIntBase = mResult.Kind & mGetP(strTemp, 3, "|")
'''                    strIntResult = mGetP(strTemp, 1, "|")
'''                End If
'''
'''
'''
'''
'''                If strIntBase = "INtIgE" Or strIntBase = "FDtIgE" Then
'''                    If IsNumeric(strIntResult) Then
'''                        If strIntResult > 100 Then
'''                            strIntResult = ">100"
'''                        Else ' strIntResult <= 100 Then
'''                            strIntResult = "=<100"
'''                        End If
'''                    Else
'''                        If Trim(strIntResult) = "> 100" Then
'''                            strIntResult = ">100"
'''                        ElseIf Trim(strIntResult) = "�� 100" Then
'''                            strIntResult = "��100"
'''                        Else
'''                            strIntResult = "��100"
'''                        End If
'''                    End If
'''                Else
''''                    If strIntResult > strSIgE Then
''''                        strSIgE = strIntResult
''''                    End If
'''                    strIntResult = mGetP(strTemp, 2, "|")
'''                    If strIntResult = "0" Then
'''                        strIntResult = "Nondetect"
'''                    Else
'''                        strIntResult = "Class " & strIntResult
'''                    End If
'''
''''                    If strIntResult > strSIgE Then
''''                        strSIgE = strIntResult
''''                    End If
''''
''''                    If IsNumeric(strIntResult) Then
''''                        If strIntResult < 0.35 Then
''''                            'strClass = "����"
''''                            strClass = "0"
''''                            strClassA = "0"
''''                            strIntResult = "0.00"
''''                            strIntResult = "Nondetect"
''''                        ElseIf strIntResult >= 0.35 And strIntResult < 0.7 Then
''''                            'strClass = "����"
''''                            strClass = "1"
''''                            strClassA = "1"
''''                            strIntResult = Format(strIntResult, "#0.#0")
''''                            strIntResult = "Class 1"
''''                        ElseIf strIntResult >= 0.7 And strIntResult < 3.5 Then
''''                            'strClass = "�߰�"
''''                            strClass = "2"
''''                            strClassA = "2"
''''                            strIntResult = Format(strIntResult, "#0.#0")
''''                            strIntResult = "Class 2"
''''                        ElseIf strIntResult >= 3.5 And strIntResult < 17.5 Then
''''                            'strClass = "�߰�/����"
''''                            strClass = "3"
''''                            strClassA = "3"
''''                            strIntResult = Format(strIntResult, "#0.#0")
''''                            strIntResult = "Class 3"
''''                        ElseIf strIntResult >= 17.5 And strIntResult < 50 Then
''''                            'strClass = "����"
''''                            strClass = "4"
''''                            strClassA = "4"
''''                            strIntResult = Format(strIntResult, "#0.#0")
''''                            strIntResult = "Class 4"
''''                        ElseIf strIntResult >= 50 And strIntResult < 100 Then
''''                            'strClass = "�ſ� ����"
''''                            strClass = "5"
''''                            strClassA = "5"
''''                            strIntResult = Format(strIntResult, "#0.#0")
''''                            strIntResult = "Class 5"
''''                        ElseIf strIntResult >= 100 Then
''''                            'strClass = "���� ����"
''''                            strClass = "6"
''''                            strClassA = "6"
''''                            strIntResult = ">=100"
''''                            strIntResult = "Class 6"
''''                        End If
''''                    Else
''''                         strClass = mGetP(strTemp, 2, "|")
''''                         strClassA = mGetP(strTemp, 2, "|")
''''                    End If
''''
''''                    If strClassA > strMaxClass Then
''''                        strMaxClass = strClassA
''''                    End If
'''
'''                End If
''''                If strIntResult = "0." Then strIntResult = "0.00"
'''
'''                'strIntResult = strIntResult '& " :" & strClass
'''                strResult = strIntResult
'''
'''                '-- ������� Class���� ����
'''                'strResult = strClass & " Class" & "(" & strResult & ")"
'''
'''                If strIntBase <> "" And strResult <> "" Then
'''                    blnSame = False
'''                    '-- �˻縶���� ���� ��������
'''                    For intTestNmCnt = 1 To UBound(gArrEQPNm)
'''                        '-- ���ä���� ����...
'''                        If strIntBase = gArrEQPNm(intTestNmCnt, 3) Then
'''                            strCheck = "0"
'''                            strSeqno = gArrEQPNm(intTestNmCnt, 1)
'''                            strState = ""
'''                            For intTestCdCnt = 1 To UBound(gArrEQP)
'''                                '-- �˻��ڵ嵵 ���ٸ�...
'''                                If strIntBase = gArrEQP(intTestCdCnt, 3) Then
'''                                    strTestCode = gArrEQP(intTestCdCnt, 2)
'''                                    strTestName = gArrEQP(intTestCdCnt, 5)
'''                                    intResPrecUse = gArrEQP(intTestCdCnt, 7)
'''                                    intResPrec = gArrEQP(intTestCdCnt, 8)
'''                                    If UBound(gPatTest) > 0 Then
'''                                        For intOrdCnt = 1 To UBound(gPatTest)
'''                                            If gPatTest(intOrdCnt) = gArrEQP(intTestCdCnt, 2) Then
'''                                                strCheck = "1"
'''
'''                                                strOrderCode = gArrEQP(intTestCdCnt, 16)
'''                                                strTestCodeSub = gArrEQP(intTestCdCnt, 17)
'''
'''                                                strState = "R"
'''                                                blnSame = True
'''                                                Exit For
'''                                            End If
'''                                        Next
'''                                    End If
'''                                End If
'''                            Next
'''
'''                            If blnSame = True Then
'''                                Exit For
'''                            End If
'''                        End If
'''                    Next
'''
'''                    '-- ���Row �߰�
'''                    intRstRow = spdResult.DataRowCnt + 1
'''                    If spdResult.MaxRows < intRstRow Then
'''                        spdResult.MaxRows = intRstRow
'''                    End If
'''
'''                    '-- ������� ǥ��("���")
'''                    SetText spdOrder, "���", gRow, colSTATE
'''
'''                    '-- ����ȭ�� ����� ǥ��
'''                    For intCol = colSTATE + 1 To spdOrder.MaxCols
'''                        If strTestName = gArrEQPNm(intCol - colSTATE, 6) Then
'''                            SetText spdOrder, strResult, gRow, intCol
'''                            Exit For
'''                        End If
'''                    Next
'''
'''                    '-- ��� List
'''                    SetText spdResult, strCheck, intRstRow, colRCHECKBOX               'üũ
'''                    SetText spdResult, strSeqno, intRstRow, colRSEQNO                  '����
'''                    SetText spdResult, strOrderCode, intRstRow, colRORDERCD            'ó���ڵ�
'''                    SetText spdResult, strTestCode, intRstRow, colRTESTCD          '�˻��ڵ�
'''                    SetText spdResult, strTestCodeSub, intRstRow, colRSUBCD        '�˻��ڵ�SUB
'''                    SetText spdResult, strTestName, intRstRow, colRTESTNM              '�˻��
'''                    SetText spdResult, strIntBase, intRstRow, colRCHANNEL              '���ä��
'''                    SetText spdResult, strMachResult, intRstRow, colRMACHRESULT        '�����
'''                    SetText spdResult, strResult, intRstRow, colRLISRESULT             'LIS���
'''                    SetText spdResult, strJudge, intRstRow, colRJUDGE                  '����
'''                    SetText spdResult, strLow & "~" & strHigh, intRstRow, colRREF      '����ġ
'''
'''                    '-- ���� ����
'''                    Call SetLocalDB(gRow, intRstRow, "1", "")
'''
'''                End If
'''            Next
'''
'''            spdResult.RowHeight(-1) = 15
'''
'''            '## DB�� �������
'''            If gHOSP.SAVEAUTO = "Y" And strState = "R" Then
'''                Res = SaveTransData(gRow, spdOrder)
'''
'''                If Res = -1 Then
'''                    '-- ���� ����
'''                    SetForeColor spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
'''                    SetText spdOrder, "�������", gRow, colSTATE
'''                Else
'''                    '-- ���� ����
'''                    SetBackColor spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
'''                    SetText spdOrder, "����Ϸ�", gRow, colSTATE
'''                    SetText spdOrder, "0", gRow, colCHECKBOX
'''
'''                          SQL = "Update PATRESULT Set " & vbCrLf
'''                    SQL = SQL & " sendflag = '2' " & vbCrLf
'''                    SQL = SQL & " Where equipno = '" & gHOSP.HOSPCD & "' " & vbCrLf
'''                    SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
'''                    SQL = SQL & "   And barcode = '" & Trim(GetText(spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
'''                    SQL = SQL & "   And saveseq = " & Trim(GetText(spdOrder, gRow, colSAVESEQ)) & vbCrLf
'''
'''                    If DBExec(AdoCn_Local, SQL) Then
'''                        '-- ����
'''                    End If
'''                End If
'''                strState = ""
'''            End If
'''
'''            intRow = intRow + 1
'''            AdoRs_Allo.MoveNext
'''        Loop
'''    End If
'''
'''    Set AdoRs_Allo = Nothing
'''
'''Exit Sub
'''
'''DBErr:
'''    Screen.MousePointer = 0
'''
'''    strErrMsg = ""
'''    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "Get_SearchListFD" & vbNewLine & vbNewLine
'''    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
'''    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
'''    frmErrMsg.txtErr = vbNewLine & strErrMsg
'''    frmErrMsg.Show
'''
'''End Sub


Private Sub Get_SearchListFD()
    Dim intRow      As Long
    Dim strAge      As String
    Dim strTransDt  As String
    Dim strHMsg     As String
    Dim strDMsg     As String
    
    Dim vWorkNo      As Variant  'Spread�� WorkNo
    Dim vBarNo       As Variant  'Spread�� ���ڵ��ȣ
    Dim strRcvBuf    As String   '������ Data
    Dim strType      As String   '������ Record Type
    Dim strOrgBarno  As String   '������ BarNO
    Dim strBarno     As String   '������ BarNO
    Dim strWorkNo    As String   '������ WorkNo
    Dim strIntResult As String   '������ �˻���
    Dim strIntBase   As String   '������ ������ �˻��
    Dim strResult    As String   'LIS �˻���
    Dim strMachResult   As String   '������ �����
    Dim strClass     As String   'Class �������
    Dim strClassA    As String   'Class �������
    Dim strMaxClass  As String   'Class �������
    
    Dim strTemp      As String
    Dim i            As Long
    Dim blnSameBar   As Boolean
    Dim intCnt       As Integer
    Dim strTIgE      As String   'Total IgE
    Dim strSIgE      As String   'Ŭ������ ���� ���� ��
    Dim strRemark    As String
    Dim SIgE         As String
    
    Dim Y, Y1, Y2, Y3, X1, X2
    

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
    Dim intCol          As Integer  '����÷� ����
    
    
    Dim intTestNmCnt    As Integer
    Dim intTestCdCnt    As Integer
    Dim intOrdCnt       As Integer
    Dim blnSame         As Boolean
    Dim Res             As Integer
    Dim RS          As ADODB.Recordset

    On Error GoTo DBErr
    
    Set AdoRs_Allo = Get_ResultList_FD
    
    If Not AdoRs_Allo.BOF Then
        intRow = 1
        strTransDt = ""
        mResult.BarNo = ""
        
        Do Until AdoRs_Allo.EOF
            strOrgBarno = AdoRs_Allo.Fields("PATIENTID").Value & ""
                        
            '-- �������
            mOrder.OrgBarNo = strOrgBarno
            With mResult
                .Kind = "FD"
                .BarNo = strOrgBarno
                .RsltDate = Format(Now, "yyyy-mm-dd")
                .RsltTime = Format(Now, "hh:mm:ss")
                .RsltSeq = getMaxTestNum(Format(Now, "yyyy-mm-dd"))
            End With
            
            Call SetPatInfo(strOrgBarno, gHOSP.RSTTYPE)
            
            strIntBase = ""
            
            For intCnt = 2 To 32
                If intCnt = 2 Then
                    strIntBase = "FD"
                End If
                
                strIntResult = ""
                strResult = ""
                strTemp = AdoRs_Allo.Fields("BANDVAL" & intCnt) & ""
                '> 100|Valid|PC|158.24
                '1.31|2|d70|1.31
                
                If strTemp <> "" Then
                    strIntBase = mResult.Kind & mGetP(strTemp, 3, "|")
                    strIntResult = mGetP(strTemp, 1, "|")
                End If
                
                If strIntBase = "INtIgE" Or strIntBase = "FDtIgE" Then
                    If IsNumeric(strIntResult) Then
                        If strIntResult > 100 Then
                            strIntResult = ">100"
                        Else ' strIntResult <= 100 Then
                            strIntResult = "=<100"
                        End If
                    Else
                        If Trim(strIntResult) = "> 100" Then
                            strIntResult = ">100"
                        ElseIf Trim(strIntResult) = "�� 100" Then
                            strIntResult = "��100"
                        Else
                            strIntResult = "��100"
                        End If
                    End If
                Else
                    strIntResult = mGetP(strTemp, 2, "|")
                    If strIntResult = "0" Then
                        strIntResult = "Nondetect"
                    Else
                        strIntResult = "Class " & strIntResult
                    End If
                End If
                strResult = strIntResult
                
                If strIntBase <> "" And strResult <> "" Then
                    blnSame = False
                    '-- �˻縶���� ���� ��������
                    For intTestNmCnt = 1 To UBound(gArrEQPNm)
                        '-- ���ä���� ����...
                        If strIntBase = gArrEQPNm(intTestNmCnt, 3) Then
                            strCheck = "0"
                            strSeqno = gArrEQPNm(intTestNmCnt, 1)
                            strState = ""
                            For intTestCdCnt = 1 To UBound(gArrEQP)
                                '-- �˻��ڵ嵵 ���ٸ�...
                                If strIntBase = gArrEQP(intTestCdCnt, 3) Then
                                    strTestCode = gArrEQP(intTestCdCnt, 2)
                                    strTestName = gArrEQP(intTestCdCnt, 5)
                                    intResPrecUse = gArrEQP(intTestCdCnt, 7)
                                    intResPrec = gArrEQP(intTestCdCnt, 8)
                                    If UBound(gPatTest) > 0 Then
                                        For intOrdCnt = 1 To UBound(gPatTest)
                                            If gPatTest(intOrdCnt) = gArrEQP(intTestCdCnt, 2) Then
                                                strCheck = "1"
                                                
                                                strOrderCode = gArrEQP(intTestCdCnt, 16)
                                                strTestCodeSub = gArrEQP(intTestCdCnt, 17)
                                                
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
                    intRstRow = spdResult.DataRowCnt + 1
                    If spdResult.MaxRows < intRstRow Then
                        spdResult.MaxRows = intRstRow
                    End If

                    '-- ������� ǥ��("���")
                    SetText spdOrder, "���", gRow, colSTATE

                    '-- ����ȭ�� ����� ǥ��
                    For intCol = colSTATE + 1 To spdOrder.MaxCols
                        If strTestName = gArrEQPNm(intCol - colSTATE, 7) Then
                            SetText spdOrder, strResult, gRow, intCol
                            Exit For
                        End If
                    Next

                    '-- ��� List
                    SetText spdResult, strCheck, intRstRow, colRCHECKBOX               'üũ
                    SetText spdResult, strSeqno, intRstRow, colRSEQNO                  '����
                    SetText spdResult, strOrderCode, intRstRow, colRORDERCD            'ó���ڵ�
                    SetText spdResult, strTestCode, intRstRow, colRTESTCD          '�˻��ڵ�
                    SetText spdResult, strTestCodeSub, intRstRow, colRSUBCD        '�˻��ڵ�SUB
                    SetText spdResult, strTestName, intRstRow, colRTESTNM              '�˻��
                    SetText spdResult, strIntBase, intRstRow, colRCHANNEL              '���ä��
                    SetText spdResult, strMachResult, intRstRow, colRMACHRESULT        '�����
                    SetText spdResult, strResult, intRstRow, colRLISRESULT             'LIS���
                    SetText spdResult, strJudge, intRstRow, colRJUDGE                  '����
                    SetText spdResult, strLow & "~" & strHigh, intRstRow, colRREF      '����ġ

                    '-- ���� ����
                    Call SetLocalDB(gRow, intRstRow, "1", "")

                End If
            Next
            
'            spdResult.RowHeight(-1) = 15
'
'            '## DB�� �������
'            If gHOSP.SAVEAUTO = "Y" And strState = "R" Then
'                Res = SaveTransData(gRow, spdOrder)
'
'                If Res = -1 Then
'                    '-- ���� ����
'                    SetForeColor spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
'                    SetText spdOrder, "�������", gRow, colSTATE
'                Else
'                    '-- ���� ����
'                    SetBackColor spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
'                    SetText spdOrder, "����Ϸ�", gRow, colSTATE
'                    SetText spdOrder, "0", gRow, colCHECKBOX
'
'                          SQL = "Update PATRESULT Set " & vbCrLf
'                    SQL = SQL & " sendflag = '2' " & vbCrLf
'                    SQL = SQL & " Where equipno = '" & gHOSP.HOSPCD & "' " & vbCrLf
'                    SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
'                    SQL = SQL & "   And barcode = '" & Trim(GetText(spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
'                    SQL = SQL & "   And saveseq = " & Trim(GetText(spdOrder, gRow, colSAVESEQ)) & vbCrLf
'
'                    If DBExec(AdoCn_Local, SQL) Then
'                        '-- ����
'                    End If
'                End If
'                strState = ""
'            End If
'
'            'intRow = intRow + 1
'
            
            Set AdoRs_Allo2 = Get_ResultList_ST(strOrgBarno)
            
            If Not AdoRs_Allo2.BOF Then
                'intRow = 1
                strTransDt = ""
                mResult.BarNo = ""
                
                Do Until AdoRs_Allo2.EOF
                    
                    strIntBase = ""
                    
                    For intCnt = 2 To 32
                        If intCnt = 2 Then
                            strIntBase = "FD"
                        End If
                        
                        strIntResult = ""
                        strResult = ""
                        strTemp = AdoRs_Allo2.Fields("BANDVAL" & intCnt) & ""
                        '> 100|Valid|PC|158.24
                        '1.31|2|d70|1.31
                        
                        If strTemp <> "" Then
                            strIntBase = mResult.Kind & mGetP(strTemp, 3, "|")
                            strIntResult = mGetP(strTemp, 1, "|")
                        End If
                        
                        If strIntBase = "INtIgE" Or strIntBase = "FDtIgE" Then
                            If IsNumeric(strIntResult) Then
                                If strIntResult > 100 Then
                                    strIntResult = ">100"
                                Else ' strIntResult <= 100 Then
                                    strIntResult = "=<100"
                                End If
                            Else
                                If Trim(strIntResult) = "> 100" Then
                                    strIntResult = ">100"
                                ElseIf Trim(strIntResult) = "�� 100" Then
                                    strIntResult = "��100"
                                Else
                                    strIntResult = "��100"
                                End If
                            End If
                        Else
                            strIntResult = mGetP(strTemp, 2, "|")
                            If strIntResult = "0" Then
                                strIntResult = "Nondetect"
                            Else
                                strIntResult = "Class " & strIntResult
                            End If
                        End If
                        strResult = strIntResult
                        
                        If strIntBase <> "" And strResult <> "" Then
                            blnSame = False
                            '-- �˻縶���� ���� ��������
                            For intTestNmCnt = 1 To UBound(gArrEQPNm)
                                '-- ���ä���� ����...
                                If strIntBase = gArrEQPNm(intTestNmCnt, 3) Then
                                    strCheck = "0"
                                    strSeqno = gArrEQPNm(intTestNmCnt, 1)
                                    strState = ""
                                    For intTestCdCnt = 1 To UBound(gArrEQP)
                                        '-- �˻��ڵ嵵 ���ٸ�...
                                        If strIntBase = gArrEQP(intTestCdCnt, 3) Then
                                            strTestCode = gArrEQP(intTestCdCnt, 2)
                                            strTestName = gArrEQP(intTestCdCnt, 5)
                                            intResPrecUse = gArrEQP(intTestCdCnt, 7)
                                            intResPrec = gArrEQP(intTestCdCnt, 8)
                                            If UBound(gPatTest) > 0 Then
                                                For intOrdCnt = 1 To UBound(gPatTest)
                                                    If gPatTest(intOrdCnt) = gArrEQP(intTestCdCnt, 2) Then
                                                        strCheck = "1"
                                                        
                                                        strOrderCode = gArrEQP(intTestCdCnt, 16)
                                                        strTestCodeSub = gArrEQP(intTestCdCnt, 17)
                                                        
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
                            intRstRow = spdResult.DataRowCnt + 1
                            If spdResult.MaxRows < intRstRow Then
                                spdResult.MaxRows = intRstRow
                            End If
        
                            '-- ������� ǥ��("���")
                            SetText spdOrder, "���", gRow, colSTATE
        
                            '-- ����ȭ�� ����� ǥ��
                            For intCol = colSTATE + 1 To spdOrder.MaxCols
                                If strTestName = gArrEQPNm(intCol - colSTATE, 7) Then
                                    SetText spdOrder, strResult, gRow, intCol
                                    Exit For
                                End If
                            Next
        
                            '-- ��� List
                            SetText spdResult, strCheck, intRstRow, colRCHECKBOX               'üũ
                            SetText spdResult, strSeqno, intRstRow, colRSEQNO                  '����
                            SetText spdResult, strOrderCode, intRstRow, colRORDERCD            'ó���ڵ�
                            SetText spdResult, strTestCode, intRstRow, colRTESTCD          '�˻��ڵ�
                            SetText spdResult, strTestCodeSub, intRstRow, colRSUBCD        '�˻��ڵ�SUB
                            SetText spdResult, strTestName, intRstRow, colRTESTNM              '�˻��
                            SetText spdResult, strIntBase, intRstRow, colRCHANNEL              '���ä��
                            SetText spdResult, strMachResult, intRstRow, colRMACHRESULT        '�����
                            SetText spdResult, strResult, intRstRow, colRLISRESULT             'LIS���
                            SetText spdResult, strJudge, intRstRow, colRJUDGE                  '����
                            SetText spdResult, strLow & "~" & strHigh, intRstRow, colRREF      '����ġ
        
                            '-- ���� ����
                            Call SetLocalDB(gRow, intRstRow, "1", "")
        
                        End If
                    Next
                    
                    spdResult.RowHeight(-1) = 15
        
                    '## DB�� �������
                    If gHOSP.SAVEAUTO = "Y" And strState = "R" Then
                        Res = SaveTransData(gRow, spdOrder)
        
                        If Res = -1 Then
                            '-- ���� ����
                            SetForeColor spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText spdOrder, "�������", gRow, colSTATE
                        Else
                            '-- ���� ����
                            SetBackColor spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText spdOrder, "����Ϸ�", gRow, colSTATE
                            SetText spdOrder, "0", gRow, colCHECKBOX
        
                                  SQL = "Update PATRESULT Set " & vbCrLf
                            SQL = SQL & " sendflag = '2' " & vbCrLf
                            SQL = SQL & " Where equipno = '" & gHOSP.HOSPCD & "' " & vbCrLf
                            SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And barcode = '" & Trim(GetText(spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                            SQL = SQL & "   And saveseq = " & Trim(GetText(spdOrder, gRow, colSAVESEQ)) & vbCrLf
        
                            If DBExec(AdoCn_Local, SQL) Then
                                '-- ����
                            End If
                        End If
                        strState = ""
                    End If
                    
                    intRow = intRow + 1
                    AdoRs_Allo2.MoveNext
                Loop
            End If
            
            Set AdoRs_Allo2 = Nothing
            
            AdoRs_Allo.MoveNext
        Loop
    End If
    
    Set AdoRs_Allo = Nothing

Exit Sub

DBErr:
    Screen.MousePointer = 0
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "Get_SearchListFD" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Sub

'   Record Set Open
Public Function Get_Recordset(ByVal AdoCn As ADODB.Connection, ByVal strSql As String, _
                             ByVal AdoRs As ADODB.Recordset, _
                             Optional Call_Name As String, _
                             Optional Cursor_Location As ADODB.CursorLocationEnum = adUseClient, _
                             Optional Cursor_Type As ADODB.CursorTypeEnum = adOpenStatic, _
                             Optional Lock_Type As ADODB.LockTypeEnum = adLockPessimistic) As Boolean

On Error GoTo DBOpenRsError
    
    With AdoRs
        .CursorLocation = Cursor_Location
        .Source = strSql
        .ActiveConnection = AdoCn
        .CursorType = Cursor_Type
        .LockType = Lock_Type
        .Open
    End With
    
    Get_Recordset = True

Exit Function

DBOpenRsError:
    Set AdoRs = Nothing
    Get_Recordset = False

End Function

'   Result List Recordset
Public Function Get_ResultList_IN() As ADODB.Recordset
    Dim strSql      As String

On Error GoTo ErrorTrap

             strSql = "SELECT * "
    strSql = strSql & "  FROM UNITS "
    strSql = strSql & " WHERE EXAMDATE >= '" & Format(dtpFrom.Value, "yyyy-mm-dd") & "' "
    strSql = strSql & "   AND EXAMDATE <= '" & Format(dtpTo.Value, "yyyy-mm-dd") & "' "
'    strSql = strSql & "   AND PANEL IN ('Inhalant','Standard') "
    strSql = strSql & "   AND PANEL IN ('Inhalant') "
    strSql = strSql & " ORDER BY EXAMDATE,PATIENTID, PANEL "
    
    
    'MsgBox strSql
    
    'SetRawData "[IN] " & strSql
    
    Set AdoRs_Allo = New ADODB.Recordset
    If Get_Recordset(AdoCn_Allo, strSql, AdoRs_Allo, "") Then
        Set Get_ResultList_IN = AdoRs_Allo
    Else
        Set Get_ResultList_IN = Nothing
    End If
    
    Set AdoRs_Allo = Nothing

Exit Function

ErrorTrap:
    Set AdoRs_Allo = Nothing

End Function

Public Function Get_ResultList_ST(ByVal pPID As String) As ADODB.Recordset
    Dim strSql      As String

On Error GoTo ErrorTrap

             strSql = "SELECT * "
    strSql = strSql & "  FROM UNITS "
    strSql = strSql & " WHERE PATIENTID =  '" & pPID & "' "
    strSql = strSql & "   AND PANEL IN ('Standard') "
    strSql = strSql & " ORDER BY EXAMDATE,PATIENTID, PANEL "
    
    Set AdoRs_Allo2 = New ADODB.Recordset
    If Get_Recordset(AdoCn_Allo, strSql, AdoRs_Allo2, "") Then
        Set Get_ResultList_ST = AdoRs_Allo2
    Else
        Set Get_ResultList_ST = Nothing
    End If
    
    Set AdoRs_Allo2 = Nothing

Exit Function

ErrorTrap:
    Set AdoRs_Allo2 = Nothing

End Function

'   Result List Recordset
Public Function Get_ResultList_FD() As ADODB.Recordset
    Dim strSql      As String

On Error GoTo ErrorTrap

             strSql = "SELECT * "
    strSql = strSql & "  FROM UNITS "
    strSql = strSql & " WHERE EXAMDATE >= '" & Format(dtpFrom.Value, "yyyy-mm-dd") & "' "
    strSql = strSql & "   AND EXAMDATE <= '" & Format(dtpTo.Value, "yyyy-mm-dd") & "' "
'    strSql = strSql & "   AND PANEL IN ('Food','Standard') "
    strSql = strSql & "   AND PANEL IN ('Food') "
    strSql = strSql & " ORDER BY EXAMDATE,PATIENTID, PANEL "
    
    'SetRawData "[FD] " & strSql
    
    
    Set AdoRs_Allo = New ADODB.Recordset
    If Get_Recordset(AdoCn_Allo, strSql, AdoRs_Allo, "") Then
        Set Get_ResultList_FD = AdoRs_Allo
    Else
        Set Get_ResultList_FD = Nothing
    End If
    
    Set AdoRs_Allo = Nothing

Exit Function

ErrorTrap:
    Set AdoRs_Allo = Nothing

End Function



Private Sub cmdResultSearch_Click()

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
                        Case "AU680":           Call Phase_Serial_AU680
                        Case "HITACHI7180":     Call Phase_Serial_HITACHI7180
                            
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

Private Sub GetOrder_ALLOSTATION(ByVal pBarNo As String, ByVal pType As String)

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
        Call SetText(.spdOrder, mOrder.OrgBarNo, intRow, colRACKNO)
        Call SetText(.spdOrder, mOrder.TubePos, intRow, colPOSNO)

        '-- ����������� �����
        .spdResult.MaxRows = 0

        '-- �˻��� ���� ��������
        Call GetSampleInfo(intRow, .spdOrder)

        .spdOrder.RowHeight(-1) = 15

        Call SetText(frmMain.spdOrder, "ORDER", intRow, colSTATE)

'        '-- �������̺��� �˻��׸� �ش��ϴ� �˻�ä�� ã�ƿ��� (intRow = ���� �˻��ߴ� ���ڵ尡 �ٽ� �ö�� ��� ��ġ�� ��ã�´�.)
'        strItems = GetEquipExamCode_AU680(gHOSP.MACHCD, pBarNo, intRow)
'
'        '-- �˻�ä�η� ������ �����
'        If Trim(strItems) = "" Then
'            mOrder.NoOrder = True
'            mOrder.Order = ""
'
'            GetOrder = STX & "S " & mOrder.RackNo & mOrder.TubePos & mOrder.Seq & Space(gHOSP.BARLEN - Len(mOrder.BarNo)) & mOrder.BarNo & Space(4) & "E" & ETX
'
'            '-- �������(Order) ǥ��
'            Call SetText(frmMain.spdOrder, "��������", intRow, colSTATE)
'        Else
'            mOrder.NoOrder = False
'            mOrder.Order = strItems
'
'            GetOrder = STX & "S " & mOrder.RackNo & mOrder.TubePos & mOrder.Seq & Space(gHOSP.BARLEN - Len(mOrder.BarNo)) & mOrder.BarNo & Space(4) & "E" & strItems & ETX
'
'            '-- �������(Order) ǥ��
'            Call SetText(frmMain.spdOrder, "��������", intRow, colSTATE)
'        End If
'
'        '-- ���� ����
'        Call SendData(GetOrder)
'
'        Call SetCommStatus("S", pBarNo, spdComStatus)

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

Private Sub SerialRcvData_HITACHI7180()
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
                    If gHOSP.BARUSE = "Y" Then
                        strFunction = Mid(strRcvBuf, 2, 40)
                    Else
                        strFunction = Mid(strRcvBuf, 2, 12) & String(gHOSP.BARLEN, "#") & Mid(strRcvBuf, 27, 15)
                    End If
                    strBarno = Trim(Mid(strRcvBuf, 14, gHOSP.BARLEN))
                    strSeq = Mid(strRcvBuf, 4, 5)
                    strRackNo = Mid(strRcvBuf, 9, 1)
                    strTubePos = Mid(strRcvBuf, 10, 3)
                    '-- ��������
                    With mOrder
                        .Function = strFunction
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .Seq = strSeq
                    End With
                    '-- ����ȯ������
                    Call GetOrder_HITACHI7180(Trim$(strBarno), gHOSP.RSTTYPE)

                Case ":"    '## End
                    '## Control, Calibration �����ʹ� ������
                    strFunc = Mid$(strRcvBuf, 2, 1)
                    If UCase(strFunc) = "H" Or UCase(strFunc) = "G" Or UCase(strFunc) = "F" Then
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

                    If mResult.BarNo = "" Then
                        Call SendData(SndMore)
                        Exit Sub
                    End If

                    strTmp = Mid$(strRcvBuf, 51)
                    
                    Do While Len(strTmp) >= 10
                        strIntBase = Trim(Mid$(strTmp, 1, 3))
                        strResult = Trim(Mid$(strTmp, 4, 6))
                        strComm = Mid$(strTmp, 9, 1)
                        
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
                                    strState = ""
'                                    '-- ȯ�� ó������ ��������
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
                            Else
                                SetText .spdResult, "", intRstRow, colRTESTCD                   '�˻��ڵ�
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
                        strTmp = Mid$(strTmp, 11)
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

RST:
    strErrMsg = ""
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "SerialRcvData_AU680" & vbNewLine & vbNewLine
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

Private Sub frmClear()
    
    shpPatInfo.Visible = False
    lblPatInfo.Caption = ""
    
    spdWork.MaxRows = 0
    spdOrder.MaxRows = 0
    spdResult.MaxRows = 0
    
    txtBarcode.Text = ""
    txtPatID.Text = ""
    txtPName.Text = ""
    txtSA.Text = ""
    
    lblFileNm.Caption = ""
    
End Sub

Private Sub Form_Load()
    Dim strTmp      As String
    Dim strSaveDt   As String
    Dim intCnt      As Integer
    
    'MsgBox "00"
    
On Error GoTo ErrHandle
    
    'MsgBox "0"
    
    Me.Width = 20940
    Me.Height = 12585

    'Me.Caption = gHOSP.MACHNM
    Me.Caption = gHOSP.MACHNM & Space$(5) & "�¢¢¢¢�     [���� �������̽�]     �¢¢¢¢�"

    'MsgBox "1"
    
    Call CtlInitializing

    'MsgBox "2"
    Call frmClear
    
    'MsgBox "3"
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

'    MsgBox "3"

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
    
    
    tmrAllo.Interval = CLng(30000) '30000
'    tmrAllo.Interval = CLng(2000) '30000
    tmrAllo.Enabled = True
    
    
    Call WinExec("C:\TEMP\HostInterface.exe", 0)
    
    FileAllo.PATH = "C:\TEMP"
        
    DoEvents
    
    
    '-- ������� ����
'    strTmp = Format$(DateAdd("d", -Val(gHOSP.SAVEDAY), Format$(Now, "YYYY-MM-DD")), "YYYY-MM-DD")
'
'    SQL = "Select count(*) From PATRESULT Where EXAMDATE <= '" & strTmp & "'"
'    Set AdoRs_Local = New ADODB.Recordset
'
'    AdoRs_Local.CursorLocation = adUseClient
'    AdoRs_Local.Open SQL, AdoCn_Local
'    If AdoRs_Local.RecordCount > 0 Then AdoRs_Local.MoveFirst
'    If Not AdoRs_Local.EOF Then intCnt = AdoRs_Local(0) & ""
'    AdoRs_Local.Close:    Set AdoRs_Local = Nothing
'
'    If intCnt > 0 Then
'        If MsgBox(gHOSP.SAVEDAY + "���� ����Ÿ�� �����Ͻðڽ��ϱ�?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
'            strSaveDt = Format$(DateAdd("d", -Val(gHOSP.SAVEDAY), Format(Now, "YYYY-MM-DD")), "YYYY-MM-DD")
'
'            SQL = "DELETE From PATRESULT Where EXAMDATE <= '" & strTmp & "'"
'            AdoCn_Local.Execute SQL
'        End If
'    End If
'
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
        strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "_Form_Load" & vbNewLine & vbNewLine
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
'    ElseIf gComm.COMTYPE = "2" Then
'        If gComm.TCPTYPE = "1" Then
'            wSck.LocalPort = CInt(gComm.TCPPORT)
'            wSck.Listen
'
'            lblComStatus.Caption = "TCP " & gComm.TCPPORT & " ����..."
'
'            imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
'            imgSend.Visible = False
'            imgReceive.Visible = False
'            lblSend.Visible = False
'            lblRcv.Visible = False
'
'        Else
'            wSck.Close
'            wSck.Connect gComm.TCPIP, CInt(gComm.TCPPORT)
'
'            lblComStatus.Caption = "TCP " & gComm.TCPIP & ":" & gComm.TCPPORT & " ����..."
'
'            imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
'            imgSend.Visible = False
'            imgReceive.Visible = False
'            lblSend.Visible = False
'            lblRcv.Visible = False
'        End If
'    ElseIf gComm.COMTYPE = "" Then

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


Private Sub tmrAllo_Timer()
    Dim intIdx      As Integer
    Dim strSrcfile  As String
    Dim strDestFile As String
    Dim strtmpBuf   As String
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long
    Dim intCnt      As Integer
    Dim strBuf      As String
    Dim TextLine

    FileAllo.Refresh
    
    DoEvents
    
    For intIdx = 0 To FileAllo.ListCount - 1
        FileAllo.ListIndex = intIdx
        
        'FileAllo.FileName = ALLO_201606081241.patlist
        '���� ��¥�Ÿ� �����´�.
        If Mid(FileAllo.Filename, 6, 8) = Format(Now, "yyyymmdd") Then
            
            If Right(FileAllo.PATH, 1) = "\" Then
                strSrcfile = FileAllo.PATH & FileAllo.Filename     ' ���� ���� �̸��� �����մϴ�.
            Else
                strSrcfile = FileAllo.PATH & "\" & FileAllo.Filename    ' ���� ���� �̸��� �����մϴ�.
            End If
            
            '���� �����̸��� ��� ó������ �ʴ´�.
            If lblFileNm.Caption = strSrcfile Then
                Exit Sub
            End If
            
            Open strSrcfile For Input As #3
    
            pBuffer = ""
            strBuf = ""
            
            Do While Not EOF(3)
                Line Input #3, TextLine ' ������ ������ ���� �о���Դϴ�.
                strBuf = strBuf & TextLine & vbCr
            Loop
    
            Close #3
            
            '��� ���� �̸��� ����
            strDestFile = App.PATH & "\Log\" & FileAllo.Filename
            
            '������ ��� ����
            FileCopy strSrcfile, strDestFile
            
            lblFileNm.Caption = strSrcfile
            FileAllo.Refresh
            
            lngBufLen = Len(strBuf)
            
            pBuffer = strBuf
            
            If Len(pBuffer) > 300 Then
                Call cmdOrder_Click
                pBuffer = ""
            End If
        End If
    Next
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
