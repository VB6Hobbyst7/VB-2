VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{4BD5DFC7-B668-44E0-A002-C1347061239D}#1.0#0"; "HSCotrol.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmComm 
   Caption         =   "Interface"
   ClientHeight    =   7095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11985
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7095
   ScaleWidth      =   11985
   WindowState     =   2  '�ִ�ȭ
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '���
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   8310
      ScaleHeight     =   345
      ScaleWidth      =   3555
      TabIndex        =   57
      Top             =   570
      Width           =   3585
      Begin VB.CheckBox chkMan 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Manual"
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
         Left            =   120
         TabIndex        =   60
         Top             =   60
         Width           =   1065
      End
      Begin VB.ComboBox cboLevel 
         Height          =   300
         Left            =   2370
         TabIndex        =   59
         Text            =   "Combo1"
         Top             =   30
         Width           =   1065
      End
      Begin VB.CheckBox chkQC 
         BackColor       =   &H00FFFFC0&
         Caption         =   "QC"
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1440
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   60
         Width           =   720
      End
   End
   Begin VB.Timer tmrReceive 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3420
      Top             =   6630
   End
   Begin VB.Timer tmrSend 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3840
      Top             =   6660
   End
   Begin MSComctlLib.ImageList imlList 
      Left            =   2250
      Top             =   6510
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_2.frx":0000
            Key             =   "ITM"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_2.frx":059A
            Key             =   "ERR"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_2.frx":0B34
            Key             =   "NOF"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_2.frx":10CE
            Key             =   "LST"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_2.frx":1668
            Key             =   "LSE"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_2.frx":1C02
            Key             =   "LSN"
         EndProperty
      EndProperty
   End
   Begin MSCommLib.MSComm comEQP 
      Left            =   2850
      Top             =   6510
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      RThreshold      =   1
      SThreshold      =   1
   End
   Begin MSComctlLib.ImageList imlStatus 
      Left            =   4350
      Top             =   6510
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
            Picture         =   "frmComm_2.frx":219C
            Key             =   "RUN"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_2.frx":2736
            Key             =   "NOT"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_2.frx":2CD0
            Key             =   "STOP"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_2.frx":326A
            Key             =   "LST"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_2.frx":3AFC
            Key             =   "ITM"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_2.frx":3C56
            Key             =   "ERR"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComm_2.frx":3DB0
            Key             =   "NOF"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraCmdBar 
      BeginProperty Font 
         Name            =   "����"
         Size            =   1.5
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   580
      Left            =   30
      TabIndex        =   1
      Top             =   6495
      Width           =   11940
      Begin HSCotrol.CButton cmdOrder 
         Height          =   360
         Left            =   5160
         TabIndex        =   52
         Top             =   135
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   635
         Caption         =   "��������"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         BorderStyle     =   1
         BorderColor     =   -2147483632
      End
      Begin HSCotrol.CButton cmdAction 
         Height          =   360
         Index           =   0
         Left            =   6375
         TabIndex        =   2
         Top             =   135
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   635
         Caption         =   "Run"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         BorderStyle     =   1
         BorderColor     =   -2147483632
      End
      Begin HSCotrol.CButton cmdAction 
         Height          =   360
         Index           =   1
         Left            =   7740
         TabIndex        =   3
         Top             =   135
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   635
         Caption         =   "Stop"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         BorderStyle     =   1
         BorderColor     =   -2147483632
      End
      Begin HSCotrol.CButton cmdAction 
         Height          =   360
         Index           =   2
         Left            =   9120
         TabIndex        =   4
         Top             =   135
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   635
         Caption         =   "Clear"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         BorderStyle     =   1
         BorderColor     =   -2147483632
      End
      Begin HSCotrol.CButton cmdAction 
         Height          =   360
         Index           =   3
         Left            =   10485
         TabIndex        =   5
         Top             =   135
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   635
         Caption         =   "Close"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         BorderStyle     =   1
         BorderColor     =   -2147483632
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         Caption         =   "�۾���� ��.."
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   180
         Left            =   960
         TabIndex        =   10
         Top             =   225
         Width           =   1200
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   " ���� :"
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
         Left            =   210
         TabIndex        =   9
         Top             =   225
         Width           =   615
      End
   End
   Begin HSCotrol.CaptionBar CaptionBar1 
      Align           =   1  '�� ����
      Height          =   555
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11985
      _ExtentX        =   21140
      _ExtentY        =   979
      Border          =   1
      CaptionBackColor=   16777215
      Picture         =   "frmComm_2.frx":3F0A
      Caption         =   " Communication"
      SubCaption      =   "�˻� ���� ����Ͽ� ����� ���� �մϴ�."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty SubCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "Receive : "
         Height          =   180
         Left            =   10140
         TabIndex        =   8
         Top             =   285
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "Send : "
         Height          =   180
         Left            =   9105
         TabIndex        =   7
         Top             =   285
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "Port : "
         Height          =   180
         Left            =   8040
         TabIndex        =   6
         Top             =   285
         Width           =   510
      End
      Begin VB.Image imgReceive 
         Height          =   240
         Left            =   11010
         Picture         =   "frmComm_2.frx":518C
         Top             =   255
         Width           =   240
      End
      Begin VB.Image imgSend 
         Height          =   240
         Left            =   9780
         Picture         =   "frmComm_2.frx":5716
         Top             =   255
         Width           =   240
      End
      Begin VB.Image imgPort 
         Height          =   240
         Left            =   8640
         Picture         =   "frmComm_2.frx":5CA0
         Top             =   255
         Width           =   240
      End
   End
   Begin TabDlg.SSTab tabWork 
      Height          =   5850
      Left            =   45
      TabIndex        =   11
      Top             =   630
      Width           =   11865
      _ExtentX        =   20929
      _ExtentY        =   10319
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   " WorkList"
      TabPicture(0)   =   "frmComm_2.frx":622A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "pnlCom2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "pnlCom"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "spdWorkList"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdSel(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdWorkList"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "spdResult1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdWordQuery"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdAppend(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "mskOrdDate"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cboRstgbn(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtBarCode"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdSel(1)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Command1"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "optBar"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "optSeq"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cmdRackNo"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cmdStartNo"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "chkAuto"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtStaCom"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).ControlCount=   20
      TabCaption(1)   =   " ���� ���"
      TabPicture(1)   =   "frmComm_2.frx":6246
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label4"
      Tab(1).Control(1)=   "spdResult2"
      Tab(1).Control(2)=   "cmdSel(2)"
      Tab(1).Control(3)=   "cmdSel(3)"
      Tab(1).Control(4)=   "cmdRstQuery"
      Tab(1).Control(5)=   "cmdAppend(1)"
      Tab(1).Control(6)=   "mskRstDate"
      Tab(1).Control(7)=   "cboRstgbn(1)"
      Tab(1).Control(8)=   "lvwCuData"
      Tab(1).ControlCount=   9
      Begin VB.TextBox txtStaCom 
         Height          =   4215
         Left            =   4350
         TabIndex        =   56
         Top             =   1260
         Visible         =   0   'False
         Width           =   7095
      End
      Begin VB.CheckBox chkAuto 
         Caption         =   "Auto(����)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6795
         TabIndex        =   55
         Top             =   540
         Value           =   1  'Ȯ��
         Width           =   1410
      End
      Begin HSCotrol.CButton cmdStartNo 
         Height          =   300
         Left            =   7080
         TabIndex        =   54
         Top             =   495
         Visible         =   0   'False
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         Caption         =   "���۹�ȣ����"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         BorderStyle     =   1
         BorderColor     =   -2147483632
      End
      Begin HSCotrol.CButton cmdRackNo 
         Height          =   300
         Left            =   8310
         TabIndex        =   53
         Top             =   495
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         Caption         =   "Rack��ȣ����"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         BorderStyle     =   1
         BorderColor     =   -2147483632
      End
      Begin VB.OptionButton optSeq 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Seq"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8790
         TabIndex        =   51
         Top             =   0
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.OptionButton optBar 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Bar"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9720
         TabIndex        =   50
         Top             =   0
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSComctlLib.ListView lvwCuData 
         Height          =   4920
         Left            =   -68160
         TabIndex        =   47
         Top             =   810
         Visible         =   0   'False
         Width           =   4725
         _ExtentX        =   8334
         _ExtentY        =   8678
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   180
         Left            =   2745
         TabIndex        =   46
         Top             =   315
         Visible         =   0   'False
         Width           =   420
      End
      Begin Threed.SSCommand cmdSel 
         Height          =   360
         Index           =   1
         Left            =   360
         TabIndex        =   27
         Top             =   900
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   644
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "frmComm_2.frx":6262
      End
      Begin VB.ComboBox cboRstgbn 
         Height          =   300
         Index           =   1
         ItemData        =   "frmComm_2.frx":66E4
         Left            =   -72570
         List            =   "frmComm_2.frx":66F1
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   16
         Top             =   495
         Visible         =   0   'False
         Width           =   2085
      End
      Begin VB.TextBox txtBarCode 
         Height          =   300
         Left            =   2415
         MaxLength       =   12
         TabIndex        =   13
         Top             =   495
         Width           =   1500
      End
      Begin VB.ComboBox cboRstgbn 
         Height          =   300
         Index           =   0
         ItemData        =   "frmComm_2.frx":671B
         Left            =   3930
         List            =   "frmComm_2.frx":6728
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   12
         Top             =   495
         Visible         =   0   'False
         Width           =   1500
      End
      Begin MSMask.MaskEdBox mskRstDate 
         Height          =   300
         Left            =   -73695
         TabIndex        =   17
         Top             =   495
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   10
         Mask            =   "####-##-##"
         PromptChar      =   "_"
      End
      Begin HSCotrol.CButton cmdAppend 
         Height          =   300
         Index           =   1
         Left            =   -64380
         TabIndex        =   18
         Top             =   495
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   529
         Caption         =   "�������"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         BorderStyle     =   1
         BorderColor     =   -2147483632
      End
      Begin HSCotrol.CButton cmdRstQuery 
         Height          =   300
         Left            =   -65460
         TabIndex        =   19
         Top             =   495
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   529
         Caption         =   "�� ȸ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         BorderStyle     =   1
         BorderColor     =   -2147483632
      End
      Begin MSMask.MaskEdBox mskOrdDate 
         Height          =   300
         Left            =   1305
         TabIndex        =   20
         Top             =   495
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   10
         Mask            =   "####-##-##"
         PromptChar      =   "_"
      End
      Begin HSCotrol.CButton cmdAppend 
         Height          =   300
         Index           =   0
         Left            =   10620
         TabIndex        =   21
         Top             =   495
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   529
         Caption         =   "�������"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         BorderStyle     =   1
         BorderColor     =   -2147483632
      End
      Begin HSCotrol.CButton cmdWordQuery 
         Height          =   300
         Left            =   9555
         TabIndex        =   22
         Top             =   495
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   529
         Caption         =   "�� ȸ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         BorderStyle     =   1
         BorderColor     =   -2147483632
      End
      Begin FPSpread.vaSpread spdResult1 
         Height          =   4815
         Left            =   2610
         TabIndex        =   25
         Top             =   900
         Width           =   9060
         _Version        =   196608
         _ExtentX        =   15981
         _ExtentY        =   8493
         _StockProps     =   64
         BackColorStyle  =   1
         ColHeaderDisplay=   0
         ColsFrozen      =   4
         DisplayRowHeaders=   0   'False
         EditEnterAction =   2
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
         MaxCols         =   14
         MaxRows         =   14
         ScrollBarMaxAlign=   0   'False
         SelectBlockOptions=   0
         SpreadDesigner  =   "frmComm_2.frx":6752
      End
      Begin HSCotrol.CButton cmdWorkList 
         Height          =   300
         Left            =   90
         TabIndex        =   26
         Top             =   5445
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   529
         Caption         =   "WorkList �ۼ�"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         BorderStyle     =   1
         BorderColor     =   -2147483632
      End
      Begin Threed.SSCommand cmdSel 
         Height          =   360
         Index           =   0
         Left            =   90
         TabIndex        =   28
         Top             =   900
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   644
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "frmComm_2.frx":6E1E
      End
      Begin FPSpread.vaSpread spdWorkList 
         Height          =   4515
         Left            =   90
         TabIndex        =   15
         Top             =   900
         Width           =   2490
         _Version        =   196608
         _ExtentX        =   4392
         _ExtentY        =   7964
         _StockProps     =   64
         BackColorStyle  =   1
         ColHeaderDisplay=   0
         ColsFrozen      =   1
         EditEnterAction =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   4
         MaxRows         =   14
         ScrollBarMaxAlign=   0   'False
         SpreadDesigner  =   "frmComm_2.frx":728C
         UserResize      =   0
      End
      Begin HSCotrol.UserPanel pnlCom 
         Height          =   5355
         Left            =   45
         TabIndex        =   29
         Top             =   495
         Visible         =   0   'False
         Width           =   11760
         _ExtentX        =   20743
         _ExtentY        =   9446
         CloseEnabled    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox txtCom 
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4395
            Left            =   45
            MultiLine       =   -1  'True
            ScrollBars      =   2  '����
            TabIndex        =   35
            Top             =   270
            Width           =   11595
         End
         Begin VB.Frame Frame1 
            Height          =   645
            Left            =   45
            TabIndex        =   30
            Top             =   4650
            Width           =   11610
            Begin HSCotrol.CButton cmdCOMSave 
               Height          =   360
               Left            =   10515
               TabIndex        =   31
               TabStop         =   0   'False
               Top             =   180
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   635
               Caption         =   "File Save"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaskColor       =   0
               BorderStyle     =   1
               BorderColor     =   -2147483632
            End
            Begin HSCotrol.CButton cmdCOMOutput 
               Height          =   360
               Left            =   1155
               TabIndex        =   32
               Top             =   180
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   635
               Caption         =   "Send"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaskColor       =   0
               BorderStyle     =   1
               BorderColor     =   8421504
            End
            Begin HSCotrol.CButton cmdCOMClear 
               Height          =   360
               Left            =   9450
               TabIndex        =   33
               TabStop         =   0   'False
               Top             =   180
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   635
               Caption         =   "Clear"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaskColor       =   0
               BorderStyle     =   1
               BorderColor     =   -2147483632
            End
            Begin HSCotrol.CButton cmdCOMInput 
               Height          =   360
               Left            =   90
               TabIndex        =   34
               Top             =   180
               Width           =   1000
               _ExtentX        =   1773
               _ExtentY        =   635
               Caption         =   "Receive"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaskColor       =   0
               BorderStyle     =   1
               BorderColor     =   8421504
            End
         End
      End
      Begin Threed.SSCommand cmdSel 
         Height          =   360
         Index           =   3
         Left            =   -74640
         TabIndex        =   48
         Top             =   900
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   644
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "frmComm_2.frx":774A
      End
      Begin Threed.SSCommand cmdSel 
         Height          =   360
         Index           =   2
         Left            =   -74910
         TabIndex        =   49
         Top             =   900
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   644
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "frmComm_2.frx":7BCC
      End
      Begin FPSpread.vaSpread spdResult2 
         Height          =   4830
         Left            =   -74910
         TabIndex        =   14
         Top             =   900
         Width           =   11670
         _Version        =   196608
         _ExtentX        =   20585
         _ExtentY        =   8520
         _StockProps     =   64
         BackColorStyle  =   1
         ColHeaderDisplay=   0
         ColsFrozen      =   4
         DisplayRowHeaders=   0   'False
         EditEnterAction =   2
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
         MaxCols         =   15
         MaxRows         =   14
         ScrollBarMaxAlign=   0   'False
         SelectBlockOptions=   0
         SpreadDesigner  =   "frmComm_2.frx":803A
      End
      Begin HSCotrol.UserPanel pnlCom2 
         Height          =   3975
         Left            =   5880
         TabIndex        =   36
         Top             =   1815
         Visible         =   0   'False
         Width           =   5940
         _ExtentX        =   10478
         _ExtentY        =   7011
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox txtCOM2 
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4395
            Left            =   90
            MultiLine       =   -1  'True
            ScrollBars      =   2  '����
            TabIndex        =   45
            Top             =   270
            Width           =   5730
         End
         Begin VB.Frame Frame2 
            Height          =   645
            Left            =   90
            TabIndex        =   37
            Top             =   4635
            Width           =   5760
            Begin MSComDlg.CommonDialog cdlFile 
               Left            =   5265
               Top             =   60
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin HSCotrol.CButton cmdChksum 
               Height          =   360
               Left            =   2205
               TabIndex        =   38
               Top             =   180
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   635
               Caption         =   "SUM"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaskColor       =   0
               BorderStyle     =   1
               BorderColor     =   8421504
            End
            Begin HSCotrol.CButton cmdCOMOutput2 
               Height          =   360
               Left            =   1155
               TabIndex        =   39
               Top             =   180
               Width           =   1000
               _ExtentX        =   1773
               _ExtentY        =   635
               Caption         =   "Send"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaskColor       =   0
               BorderStyle     =   1
               BorderColor     =   8421504
            End
            Begin HSCotrol.CButton cmdCOMClear2 
               Height          =   360
               Left            =   3600
               TabIndex        =   40
               TabStop         =   0   'False
               Top             =   180
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   635
               Caption         =   "Clear"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaskColor       =   0
               BorderStyle     =   1
               BorderColor     =   -2147483632
            End
            Begin HSCotrol.CButton cmdCOMInput2 
               Height          =   360
               Left            =   90
               TabIndex        =   41
               Top             =   180
               Width           =   1000
               _ExtentX        =   1773
               _ExtentY        =   635
               Caption         =   "Receive"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaskColor       =   0
               BorderStyle     =   1
               BorderColor     =   8421504
            End
            Begin HSCotrol.CButton cmdCOMLoad 
               Height          =   360
               Left            =   4635
               TabIndex        =   42
               TabStop         =   0   'False
               Top             =   180
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   635
               Caption         =   "File Load"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaskColor       =   0
               BorderStyle     =   1
               BorderColor     =   -2147483632
            End
            Begin HSCotrol.CButton cmdACK 
               Height          =   360
               Left            =   3105
               TabIndex        =   43
               Top             =   180
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   635
               Caption         =   "ACK"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaskColor       =   0
               BorderStyle     =   1
               BorderColor     =   8421504
            End
            Begin HSCotrol.CButton cmdENQ 
               Height          =   360
               Left            =   2655
               TabIndex        =   44
               Top             =   180
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   635
               Caption         =   "ENQ"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaskColor       =   0
               BorderStyle     =   1
               BorderColor     =   8421504
            End
         End
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "�˻����� :"
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
         Left            =   -74910
         TabIndex        =   24
         Top             =   570
         Width           =   1125
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "��ü������ :"
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
         Left            =   90
         TabIndex        =   23
         Top             =   570
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frmComm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Const COL_KEY       As String = "K"
Private Const COL_EQP_NUM   As String = "EQP_ID"

Private Const KEY_SEQ       As String = "KEY_SEQ"   ' "����"
Private Const KEY_PTID      As String = "KEY_PTID"  ' "��Ϲ�ȣ"
Private Const KEY_PTNM      As String = "KEY_PTNM"  ' "��  ��"
Private Const KEY_SPCNO     As String = "KEY_SPCNO" ' "��ü��ȣ"
Private Const KEY_EQPNO     As String = "KEY_EQPNO" ' "��ü��ȣ"
Private Const KEY_STAT      As String = "KEY_STAT"  ' "�� ��"
Private Const KEY_TEST      As String = "KEY_TEST"  ' "�˻��׸�"

Private Const TEST_NM_EQP   As String = "EQP_NM"    '��� �ڵ�
Private Const TEST_CD_LIS   As String = "LIS_CD"    '�˻�� �ڵ�
Private Const TEST_NM_LIS   As String = "LIS_NM"    '�˻�� �̸�
Private Const TEST_VALUES   As String = "VALUES"    '���

Public WithEvents Result As clsMsg_Result
Attribute Result.VB_VarHelpID = -1
Public WithEvents Order  As clsMsg_Query
Attribute Order.VB_VarHelpID = -1
Public Result1 As clsResult
Attribute Result1.VB_VarHelpID = -1

Private mAdoRs      As ADODB.Recordset
Private CallForm    As String
Private IS_SET      As Boolean

Private f_strBuffer     As String
Private f_strJOB_FLAG   As String
Private f_strOrdList    As String
Private f_intSampleNo   As Integer

Private f_blnWorkList   As Boolean
Private f_lngWork_Row   As Long

Private MSG_STX     As String
Private MSG_ETX     As String
Private MSG_ENQ     As String
Private MSG_EOT     As String
Private MSG_ACK     As String
Private MSG_NAK     As String
Private MSG_CR      As String
Private MSG_LF      As String
Private MSG_CRLF    As String

Dim fRcvString As String
Dim fChannel() As String
Dim fElecsys2010(100) As String
Dim fElecsys2010Cfg(100) As Integer
Dim Flag_HQL As String
Dim Patiant_Recevid As Boolean

Const STX As String = ""
Const ETX As String = ""
Const ENQ As String = ""
Const ACK As String = ""
Const NAK As String = ""
Const EOT As String = ""

Dim PatientID As String    'Q Message Pattern Check
Dim PatientSeq As String
Dim PatientDisk As String
Dim PatientRack As String
Dim PatientPos As String
Dim SendCount As Integer

Dim OrderCnt As Integer
Dim vRow As Integer

Private Type ITEMCODE
    MachCode As String
    MachCnt As Integer
    MachTst(1 To 99) As String
End Type

Private Channel() As ITEMCODE

Private Function f_funAdd_Server(ByVal strBarno As String, ByVal strTestcd As String, _
                                 ByVal strTestval As String, ByRef strOrdLst() As String) As Boolean
                                 
    Dim strErrMsg       As String
    Dim strSampleno()   As String
    Dim strOrdcd()      As String, strRstval()  As String
    Dim strTmp1()       As String, strTmp2()    As String, strTmp   As String
    Dim intPos          As Integer, intIdx      As Integer
    Dim blnFlag         As Boolean
    
    blnFlag = False
    f_funAdd_Server = False
    
    strTmp = strTestcd: intPos = InStr(strTmp, ",")
    Do While intPos > 0
        blnFlag = False
        For intIdx = 0 To UBound(strOrdLst) - 1
            If strOrdLst(intIdx) = Mid$(strTmp, 1, intPos - 1) Then
                blnFlag = True
                strTmp = Mid$(strTmp, 1, intPos - 1)
                Exit Do
            End If
        Next
        
        strTmp = Mid$(strTmp, intPos + 1)
        intPos = InStr(strTmp, ",")
    Loop
    
    If Not blnFlag Then
        For intIdx = 0 To UBound(strOrdLst) - 1
            If strOrdLst(intIdx) = strTmp Then blnFlag = True: Exit For
        Next
    End If
    
    If blnFlag Then
        ReDim Preserve strSampleno(1 To 1) As String
        ReDim Preserve strOrdcd(1 To 1) As String
        ReDim Preserve strRstval(1 To 1) As String
        ReDim Preserve strTmp1(1 To 1) As String
        ReDim Preserve strTmp2(1 To 1) As String
        
        strSampleno(1) = strBarno
        strOrdcd(1) = strTmp
        strRstval(1) = strTestval
        
        Call sl_online_result_ul_4&(strErrMsg, strSampleno, strOrdcd, strRstval, strTmp1, strTmp2, Chr(0))
        If strErrMsg = "" Then
            f_funAdd_Server = True
        Else
            Call ErrMsgProc("", strErrMsg)
        End If
'    Else
'        Call ErrMsgProc("", "��ü��ȣ [" + strBarno + "]�� �������� ���߽��ϴ�.")
    End If
                                
End Function


Private Function f_funAdd_QcServer(ByVal strBarno As String, ByVal strTestcd As String, _
                                 ByVal strTestval As String, ByRef strOrdLst() As String) As Boolean
                                 
    Dim strErrMsg       As String
    Dim strSampleno()   As String
    Dim strOrdcd()      As String, strRstval()  As String
    Dim strTmp1()       As String, strTmp2()    As String, strTmp   As String
    Dim intPos          As Integer, intIdx      As Integer
    Dim blnFlag         As Boolean
    
    blnFlag = False
    f_funAdd_Server = False
    
    strTmp = strTestcd: intPos = InStr(strTmp, ",")
    Do While intPos > 0
        blnFlag = False
        For intIdx = 0 To UBound(strOrdLst) - 1
            If strOrdLst(intIdx) = Mid$(strTmp, 1, intPos - 1) Then
                blnFlag = True
                strTmp = Mid$(strTmp, 1, intPos - 1)
                Exit Do
            End If
        Next
        
        strTmp = Mid$(strTmp, intPos + 1)
        intPos = InStr(strTmp, ",")
    Loop
    
    If Not blnFlag Then
        For intIdx = 0 To UBound(strOrdLst) - 1
            If strOrdLst(intIdx) = strTmp Then blnFlag = True: Exit For
        Next
    End If
    
    If blnFlag Then
        ReDim Preserve strSampleno(1 To 1) As String
        ReDim Preserve strOrdcd(1 To 1) As String
        ReDim Preserve strRstval(1 To 1) As String
        ReDim Preserve strTmp1(1 To 1) As String
        ReDim Preserve strTmp2(1 To 1) As String
        
        strSampleno(1) = strBarno
        strOrdcd(1) = strTmp
        strRstval(1) = strTestval
        
        Call sl_online_pc_98&(strErrMsg, strSampleno, strOrdcd, strRstval, strTmp1, strTmp2, Chr(0))
        If strErrMsg = "" Then
            f_funAdd_Server = True
        Else
            Call ErrMsgProc("", strErrMsg)
        End If
'    Else
'        Call ErrMsgProc("", "��ü��ȣ [" + strBarno + "]�� �������� ���߽��ϴ�.")
    End If
                                
End Function

Private Function f_funGet_ConvertResult(ByVal strRstval As String) As String

    Dim intPos  As Integer
    Dim strTmp1 As String, strTmp2  As String
    
    intPos = InStr(strRstval, "E")
    If intPos > 0 Then
        strTmp1 = Mid$(strRstval, 1, intPos - 1)
        strTmp2 = Mid$(strRstval, intPos + 1)
        
        If Mid$(strTmp2, 1, 1) = "-" Then
            f_funGet_ConvertResult = Round(Val(strTmp1) * (0.1 ^ Val(Mid$(strTmp2, 2))), 2)
        Else
            f_funGet_ConvertResult = Round(Val(strTmp1) * (10 ^ Val(Mid$(strTmp2, 2))), 2)
        End If
    Else
        f_funGet_ConvertResult = strRstval
    End If
    
End Function

Private Function f_funGet_SpreadRow(ByVal objSpd As vaSpread, ByVal intCol As Integer, _
                                    ByVal strPara As String) As Integer

    Dim varTmp  As Variant
    Dim intRow  As Integer
    
    f_funGet_SpreadRow = 0
    
    With objSpd
        For intRow = 1 To .maxrows
            .GetText intCol, intRow, varTmp
            If Trim$(varTmp) = strPara Then
                f_funGet_SpreadRow = intRow
                Exit For
            End If
        Next
    End With
    
End Function

Private Sub f_subGet_JobList(ByVal strKeyno As String, ByRef strOrder As String, _
                             ByRef intOrdCnt As Integer, ByRef strSpec As String, _
                             ByRef strPcFlag As String)

    Dim adoRS1  As New ADODB.Recordset
    Dim adoRS2  As New ADODB.Recordset
    Dim sqlDoc  As String
    
    strOrder = "":  strPcFlag = "  ":   strSpec = "SE": intOrdCnt = 0
    sqlDoc = "select ORD_CODE, CHART_NO From L3A01" & _
             " where SAMPLE_DATE = '" & Mid$(strKeyno, 1, 8) & "'" & _
             "   and SAMPLE_SEQ  = " & Format(Mid$(strKeyno, 9, 3), "##0") & "" & _
             "   and PART        = '" & Mid$(strKeyno, 12, 2) & "'"
    adoRS1.CursorLocation = adUseClient
    adoRS1.Open sqlDoc, AdoCn_SQL
    If adoRS1.RecordCount > 0 Then adoRS1.MoveFirst
    
    sqlDoc = "select TESTCD_EQP, TESTCD, REMARK, AUTOVERIFY from INTERFACE002 where (EQP_CD = " & STS(INS_CODE) & ") AND (TESTCD <> '')"
    adoRS2.CursorLocation = adUseClient
    adoRS2.Open sqlDoc, AdoCn_Jet
    If adoRS2.RecordCount > 0 Then adoRS2.MoveFirst
    Do While Not adoRS2.EOF
        If adoRS1.RecordCount > 0 Then adoRS1.MoveFirst
        adoRS1.Find "ORD_CODE = " & STS(Trim(adoRS2("TESTCD") & ""))
        If Not adoRS1.EOF Then
            Select Case Trim(adoRS2(2) & "")
                Case "128": strSpec = "PL"
                Case Else:  strSpec = "SE"
            End Select
            
            If Trim(adoRS2("TESTCD_EQP") & "") = "XXX" Then
                strOrder = strOrder + "06A ," + Trim$(adoRS2("AUTOVERIFY") & "") + ",": strPcFlag = "PC"
            Else
                strOrder = strOrder + Trim(adoRS2("TESTCD_EQP") & "") + " ," + Trim$(adoRS2("AUTOVERIFY") & "") + ","
            End If
            intOrdCnt = intOrdCnt + 1
        End If
        adoRS2.MoveNext
    Loop
    adoRS2.Close:   Set adoRS2 = Nothing
    adoRS1.Close:   Set adoRS1 = Nothing
    
    If strOrder <> "" Then strOrder = Mid$(strOrder, 1, Len(strOrder) - 1)
    
End Sub

Private Sub f_subGet_WorkList(ByRef strOrder As String, ByRef intOrdCnt As Integer, _
                              ByRef strSpec As String, ByRef strPcFlag As String, _
                              ByVal intRow As Integer)

    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String

    Dim varTmp  As Variant
    Dim intCol  As Integer
    
    Dim itemX   As ListItems
    
    Set itemX = lvwCuData.ListItems
    
    strOrder = "":  strPcFlag = "  ": strSpec = "SE":   intOrdCnt = 0
    With spdWorkList
        For intCol = 5 To .MaxCols
            .Row = intRow:  .Col = intCol
            If .BackColor = &HC6FEFF Then
                Select Case itemX.Item(intCol - 4).SubItems(11)
                    Case "128": strSpec = "PL"
                    Case Else:  strSpec = "SE"
                End Select
                .GetText intCol, 0, varTmp
                
                If itemX.Item(intCol - 4).tag = "XXX" Then
                    strOrder = strOrder + "06A ," + itemX.Item(intCol - 4).SubItems(10) + ",": strPcFlag = "PC"
                Else
                    strOrder = strOrder + itemX.Item(intCol - 4).tag + " ," & itemX.Item(intCol - 4).SubItems(10) + ","
                End If
                intOrdCnt = intOrdCnt + 1
            End If
        Next
    End With
    
    If strOrder <> "" Then strOrder = Mid$(strOrder, 1, Len(strOrder) - 1)
   
End Sub

Private Sub f_subSet_ComCharacter()

    MSG_STX = Chr(COM_STX)
    MSG_ETX = Chr(COM_ETX)
    MSG_ENQ = Chr(COM_ENQ)
    MSG_EOT = Chr(COM_EOT)
    MSG_ACK = Chr(COM_ACK)
    MSG_NAK = Chr(COM_NACK)
    MSG_CR = Chr(COM_CR)
    MSG_LF = Chr(COM_LF)
    MSG_CRLF = Chr(COM_CR) & Chr(COM_LF)
    
End Sub

Private Sub f_subSet_ItemHeader()
    
    '�˻��ڵ� ���̺�
    With lvwCuData
        .View = lvwReport
        Set .ColumnHeaderIcons = imlList
        Set .SmallIcons = imlList
        .FullRowSelect = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .HideColumnHeaders = True
        With .ColumnHeaders
            .Clear
'            Call .Add(, TEST_NM_EQP, "ID", (lvwCuData.Width - 310) * 0.2)
'            Call .Add(, TEST_CD_LIS, "�˻��ڵ�", (lvwCuData.Width - 310) * 0.2)
'            Call .Add(, TEST_NM_LIS, "�� �� ��", (lvwCuData.Width - 310) * 0.2)
'            Call .Add(, TEST_VALUES, "�˻���", (lvwCuData.Width - 310) * 0.2)
'            Call .Add(, "DELTA", "DELTA", (lvwCuData.Width - 310) * 0.2)
'            Call .Add(, "DELTAGBN", "DELTAGBN", (lvwCuData.Width - 310) * 0.2)
'            Call .Add(, "PANICL", "PANIC(L)", (lvwCuData.Width - 310) * 0.2)
'            Call .Add(, "PANICH", "PANIC(H)", (lvwCuData.Width - 310) * 0.2)
'            Call .Add(, "REFL", "����ġ(L)", (lvwCuData.Width - 310) * 0.2)
'            Call .Add(, "REFH", "����ġ(H)", (lvwCuData.Width - 310) * 0.2)
'            Call .Add(, "AUTOVERIFY", "���", (lvwCuData.Width - 310) * 0.1)
'            Call .Add(, "REMARK", "��ü�ڵ�", (lvwCuData.Width - 310) * 0.1)
            Call .Add(, TEST_NM_EQP, "ID", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, TEST_CD_LIS, "�˻��ڵ�", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, TEST_NM_LIS, "�� �� ��", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, TEST_VALUES, "�˻���", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "DELTA", "DELTA", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "DELTAGBN", "DELTAGBN", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "PANICL", "PANIC(L)", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "PANICH", "PANIC(H)", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "REFL", "����ġ(L)", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "REFH", "����ġ(H)", (lvwCuData.Width - 310) * 0.2)
            Call .Add(, "AUTOVERIFY", "���", (lvwCuData.Width - 310) * 0.1)
            Call .Add(, "REMARK", "��ü�ڵ�", (lvwCuData.Width - 310) * 0.1)
            Call .Add(, "ETC", "���", (lvwCuData.Width - 310) * 0.1)
        End With
        .HideColumnHeaders = False
    End With
    
   
End Sub

Private Sub f_subSet_ItemComplete(lvw As Listview)

    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String
    
    Dim itemH           As ColumnHeader
    Dim objHeadeItem    As clsCommon
    
    Dim intCol  As Integer
    
    lvw.ColumnHeaders.Clear
    Call lvw.ColumnHeaders.Add(, "EQP_ID", "��ü ��ȣ")
    
    intCol = 4
    sqlDoc = "select RTRIM(LTRIM(TESTCD_EQP)) AS TESTCD_EQP, TESTNM_EQP, OUT_SEQ, TESTCD, TESTNM, AUTOVERIFY, REMARK," & _
             "       REFL, REFH, DELTA, DELTAGBN, PANICL, PANICH" & _
             "  from INTERFACE002" & _
             " where (EQP_CD = '" & INS_CODE & "') AND ((TESTCD <> '') AND (TESTCD IS NOT NULL))" & _
             " order by TESTCD"
             
    adoRS.CursorLocation = adUseClient
    adoRS.Open sqlDoc, AdoCn_Jet
    If adoRS.RecordCount > 0 Then adoRS.MoveFirst
    Do While Not adoRS.EOF
        With lvw
            .Enabled = True
            Set itemH = .ColumnHeaders.Add
            With itemH
                '�÷� ���Ű�� ���˻� �ڵ��
                .Key = COL_KEY & Trim(adoRS.Fields("TESTCD_EQP") & "")
                '�÷����� �˻� �׸� �̸�
                .Text = Trim(adoRS.Fields("TESTNM") & "")
                '�ױ״� �˻� �ڵ��
                .tag = Trim(adoRS.Fields("TESTCD") & "")
                .Width = 700
                .Alignment = lvwColumnCenter
            End With
            Set itemH = Nothing
        End With
        
        With spdWorkList
            intCol = intCol + 1
            If intCol > .MaxCols Then .MaxCols = .MaxCols + 1:  .ColWidth(.MaxCols) = 6.5
            
            .SetText intCol, 0, adoRS.Fields("TESTNM")
        End With
        adoRS.MoveNext
    Loop
    adoRS.Close:    Set adoRS = Nothing
    
End Sub

Private Sub f_subSet_ItemList()
    Dim itemX   As ListItem
    Dim itemA   As ListItem
    
    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String
    
    Dim intCol  As Integer, intPos  As Integer
    Dim strTmp  As String, strTest  As String
    Dim intCnt  As Integer
    
On Error GoTo ErrRoutine
    CallForm = "frmInterface - Private Sub f_subSet_ItemList()"
    
    lvwCuData.ListItems.Clear:  f_strOrdList = ""
    
    intCol = 7
    With spdWorkList
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .maxrows = 14
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
    End With
    
    With spdResult1
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .maxrows = 15
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
    End With
    
    With spdResult2
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .maxrows = 15
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
    End With
    
    sqlDoc = "select RTRIM(LTRIM(TESTCD_EQP)) as TEST_EQP, TESTNM_EQP, OUT_SEQ, TESTCD, TESTNM, AUTOVERIFY, REMARK," & _
             "       REFL, REFH, DELTA, DELTAGBN, PANICL, PANICH" & _
             "  from INTERFACE002" & _
             " where (EQP_CD = " & STS(INS_CODE) & ") AND ((TESTCD <> '') AND (TESTCD IS NOT NULL))"
             
    sqlDoc = sqlDoc + " order by OUT_SEQ, TESTCD "
    
    adoRS.CursorLocation = adUseClient
    adoRS.Open sqlDoc, AdoCn_Jet
    If adoRS.RecordCount > 0 Then adoRS.MoveFirst: ReDim fChannel(adoRS.RecordCount)
    Do While Not adoRS.EOF
        Set itemX = lvwCuData.ListItems.Add(, , Trim(adoRS.Fields("TEST_EQP") & ""), , "LST")
            
            intCnt = intCnt + 1
            ReDim Preserve Channel(1 To intCnt) As ITEMCODE
            
            Channel(intCnt).MachCode = Trim$(adoRS.Fields("TEST_EQP"))
            Channel(intCnt).MachCnt = 0
            
            strTmp = Trim$(adoRS.Fields("TESTCD"))
            intPos = InStr(strTmp, ",")
            Do While intPos > 0
                strTest = strTest + "'" + Mid$(strTmp, 1, intPos - 1) + "',"
                
                Channel(intCnt).MachCnt = Channel(intCnt).MachCnt + 1
                Channel(intCnt).MachTst(Channel(intCnt).MachCnt) = Mid$(strTmp, 1, intPos - 1)
                strTmp = Mid$(strTmp, intPos + 1)
                
                intPos = InStr(strTmp, ",")
            Loop
            strTest = strTest + "'" + strTmp + "',"
            Channel(intCnt).MachCnt = Channel(intCnt).MachCnt + 1
            Channel(intCnt).MachTst(Channel(intCnt).MachCnt) = strTmp
            
            itemX.SubItems(1) = Trim(adoRS.Fields("TESTCD") & "")
            itemX.SubItems(2) = Trim(adoRS.Fields("TESTNM") & "")
            itemX.SubItems(3) = ""
            itemX.SubItems(4) = Trim(adoRS.Fields("DELTA") & "")
            itemX.SubItems(5) = Trim(adoRS.Fields("DELTAGBN") & "")
            itemX.SubItems(6) = Trim(adoRS.Fields("PANICL") & "")
            itemX.SubItems(7) = Trim(adoRS.Fields("PANICH") & "")
            itemX.SubItems(8) = Trim(adoRS.Fields("REFL") & "")
            itemX.SubItems(9) = Trim(adoRS.Fields("REFH") & "")
            itemX.SubItems(10) = Trim(adoRS.Fields("AUTOVERIFY") & "")
            itemX.SubItems(11) = Trim(adoRS.Fields("REMARK") & "")
            itemX.SubItems(12) = strTest
            itemX.tag = Trim(adoRS.Fields("TEST_EQP") & "")
            'itemX.Text = Trim(adoRS.Fields("TESTCD") & "")
            itemX.Text = Trim(adoRS.Fields("TESTCD") & "")
        Set itemX = Nothing
        
        With spdWorkList
            If intCol - 2 > .MaxCols Then .MaxCols = .MaxCols + 1
            
            .SetText intCol - 2, 0, Trim$(adoRS("TESTNM") & "")
'            .Col= intcol-2: .ColHidden = true
        End With
        
        With spdResult1
            If intCol > .MaxCols Then .MaxCols = .MaxCols + 1
            .SetText intCol, 0, Trim$(adoRS("TESTNM") & "")
        End With
        
        With spdResult2
            If intCol > .MaxCols Then .MaxCols = .MaxCols + 1
            .SetText intCol, 0, Trim$(adoRS("TESTNM") & "")
        End With
        
        fChannel(intCol - 6) = adoRS.Fields("TEST_EQP")
        
        intCol = intCol + 1
        If InStr(Trim$(adoRS.Fields("TESTCD")), ",") > 0 Then
            f_strOrdList = f_strOrdList + ", '" & Mid(Trim$(adoRS.Fields("TESTCD")), 1, InStr(Trim$(adoRS.Fields("TESTCD")), ",") - 1) & "'"
            f_strOrdList = f_strOrdList + ", '" & Mid(Trim$(adoRS.Fields("TESTCD")), InStr(Trim$(adoRS.Fields("TESTCD")), ",") + 1) & "'"
        Else
            f_strOrdList = f_strOrdList + ", '" & Trim$(adoRS.Fields("TESTCD")) & "'"
        End If
        
        adoRS.MoveNext
    Loop
    Set adoRS = Nothing
    
    With spdResult2
        If intCol > .MaxCols Then .MaxCols = .MaxCols + 1
        .SetText intCol, 0, ""
        .Col = intCol:  .ColHidden = True
    End With
    
    f_strOrdList = Mid$(f_strOrdList, 3)
    
Exit Sub
ErrRoutine:
    Set adoRS = Nothing
    Call ErrMsgProc(CallForm)
        
End Sub


Private Sub CButton2_Click()

End Sub



Private Sub cmdACK_Click()
'
'    Call COM_OUTPUT(charCOM_Convert(COM_ACK))
Call COM_OUTPUT(Chr(1))
End Sub

Private Sub cmdAction_Click(Index As Integer)
    
    Select Case Index
        Case 0
            Call cmdRun
        Case 1
            Call cmdStop
        Case 2
            Call cmdClear
        Case 3 'cmd close
            Call cmdExit
        Case Else
    End Select

End Sub

Private Sub cmdClear()
    
    f_strJOB_FLAG = "1"
    f_intSampleNo = 0
    
    With spdWorkList
        .maxrows = 14
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 12
    End With
    
    With spdResult1
        .maxrows = 15
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .BlockMode = True
        .Action = ActionClearText
        .BackColor = vbWhite
        .BlockMode = False
        .RowHeight(-1) = 12
        '.MaxCols = 4
    End With

    With spdResult2
        .maxrows = 15
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 12
        '.MaxCols = 4
    End With

End Sub

Private Sub cmdExit()
    
    Unload Me

End Sub

Private Sub cmdRun()
    
    Dim itemX As ListItem
    
On Error GoTo ErrRoutine
    CallForm = "frmInterface - Private Sub cmdRun()"
    
    If Not comEQP.PortOpen Then comEQP.PortOpen = True
    If comEQP.PortOpen Then
        Call ShowMessage("���� �Ǿ����ϴ�.")
        imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
        imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
        imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
        lblStatus = "�۾���.."
    Else
        Call ShowMessage("���� ���� �ʾҽ��ϴ�.")
        imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
        imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
        imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
        lblStatus = "�۾� �����.."
    End If
        
Exit Sub
ErrRoutine:
    Call ErrMsgProc(CallForm)
End Sub

Private Sub cmdStop()
On Error GoTo ErrRoutine
    CallForm = "frmInterface - Private Sub cmdRun()"
    
    If comEQP.PortOpen Then comEQP.PortOpen = False
    If comEQP.PortOpen Then
        Call ShowMessage("���� ���� �ʾҽ��ϴ�.")
        imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
        imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
        imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
        lblStatus = "�۾���.."
    Else
        imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
        imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
        imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
        lblStatus = "�۾� �����.."
    End If
Exit Sub
ErrRoutine:
    Call ErrMsgProc(CallForm)
End Sub

Private Sub cmdAppend_Click(Index As Integer)
    
    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String
        
    Dim varTmp  As Variant, strErrMsg   As String
    Dim strSampleno()   As String, strBarno     As String, strTime      As String
    Dim strOrdcd()      As String, strRstval()  As String, intCnt       As Integer
    Dim strTmp1()       As String, strTmp2()    As String
    Dim intPos          As String, strTestcd    As String, strTestRst   As String
    
    Dim strOrdLst()     As String, strPid()    As String, strPnm() As String
    
    Dim intRow  As Integer, intCol  As Integer, intIdx  As Integer, blnFlag As Boolean
    Dim itemX   As ListItem
    Dim objSpd  As vaSpread
    
    CallForm = "frmComm - Private Sub cmdAppend_Click()"
    
On Error GoTo ErrorRoutine
    
    Me.MousePointer = 11
    
    If Index = 0 Then
        Set objSpd = spdResult1
    Else
        Set objSpd = spdResult2
    End If
    
    With objSpd
        For intRow = 1 To .maxrows
            .GetText 2, intRow, varTmp:         strBarno = Trim$(varTmp)
            .GetText .MaxCols, intRow, varTmp:  strTime = Trim$(varTmp)
            .GetText 1, intRow, varTmp
            
            If strBarno = "" Then Exit For
            If chkQC.Value = 0 Then
                Call sl_spcid_tstcd_select&(strBarno, strOrdLst, strPid, strPnm)
            Else
                Call sl_spcid_tstcd_select_qc&(INS_CODE, strBarno, strOrdLst)
            End If
            
            intCnt = 0: Erase strOrdcd: Erase strRstval
            If Trim$(varTmp) = "1" Then
                For intCol = 7 To .MaxCols
                    .GetText intCol, intRow, varTmp
                    If Trim$(varTmp) <> "" Then
                        .GetText intCol, 0, varTmp
                        Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
                        If Not itemX Is Nothing Then
                            .GetText intCol, intRow, varTmp
                            strTestcd = itemX.ListSubItems(1)
                            intPos = InStr(strTestcd, ",")
                            Do While intPos > 0
                                
                                blnFlag = False
                                For intIdx = 0 To UBound(strOrdLst)
                                    If strOrdLst(intIdx) = Mid$(strTestcd, 1, intPos - 1) Then blnFlag = True:  Exit For
                                Next
                                
                                If blnFlag Then
                                    intCnt = intCnt + 1
                                    ReDim Preserve strSampleno(1 To intCnt) As String
                                    ReDim Preserve strOrdcd(1 To intCnt) As String
                                    ReDim Preserve strRstval(1 To intCnt) As String
                                    ReDim Preserve strTmp1(1 To intCnt) As String
                                    ReDim Preserve strTmp2(1 To intCnt) As String
                                    
                                    strSampleno(intCnt) = strBarno
                                    strOrdcd(intCnt) = Mid$(strTestcd, 1, intPos - 1)
                                    strRstval(intCnt) = Trim$(varTmp)
                                End If
                                
                                strTestcd = Mid$(strTestcd, intPos + 1)
                                intPos = InStr(strTestcd, ",")
                            Loop
                            
                            blnFlag = False
                            For intIdx = 0 To UBound(strOrdLst)
                                If strOrdLst(intIdx) = strTestcd Then blnFlag = True: Exit For
                            Next
                            
                            If blnFlag Then
                                intCnt = intCnt + 1
                                ReDim Preserve strSampleno(1 To intCnt) As String
                                ReDim Preserve strOrdcd(1 To intCnt) As String
                                ReDim Preserve strRstval(1 To intCnt) As String
                                ReDim Preserve strTmp1(1 To intCnt) As String
                                ReDim Preserve strTmp2(1 To intCnt) As String
                                
                                strSampleno(intCnt) = strBarno
                                strOrdcd(intCnt) = strTestcd
                                strRstval(intCnt) = Trim$(varTmp)
                            End If
                        End If
                        Set itemX = Nothing
                    End If
                Next
                
                If intCnt > 0 Then
                    If chkQC.Value = 0 Then
                        Call sl_online_result_ul_4&(strErrMsg, strSampleno, strOrdcd, strRstval, strTmp1, strTmp2, Chr(0))
                    Else
                        Call sl_online_pc_98&(strErrMsg, strSampleno, strOrdcd, strRstval, strTmp1, strTmp2, Chr(0))
                    End If
                    If strErrMsg = "" Then
                        If Index = 1 Then
                            sqlDoc = "Update INTERFACE003 set SERVERGBN = 'Y'" & _
                                     " where SPCNO   = '" & strBarno & "'" & _
                                     "   and TRANSDT = '" & mskRstDate.Text & "'" & _
                                     "   and TRANSTM = '" & strTime & "'"
                            AdoCn_Jet.Execute sqlDoc
                        End If
                    Else
                        MsgBox strErrMsg, vbInformation, Me.Caption
                    End If
                Else
                    MsgBox "��ü��ȣ [" + strBarno + "]�� �������� ���߽��ϴ�.", vbInformation, Me.Caption
                End If
            End If
        Next
    End With
    Me.MousePointer = 0
    MsgBox "�۾��� �Ϸ�Ǿ����ϴ�.", vbInformation, Me.Caption
    
    Exit Sub
ErrorRoutine:
    Set itemX = Nothing
    
    Me.MousePointer = 0
    Call ErrMsgProc(CallForm)
    
End Sub

Private Sub cmdENQ_Click()
    
    Call COM_OUTPUT(charCOM_Convert(COM_ENQ))

End Sub

Private Sub cmdOrder_Click()
Dim ii As Integer
                           
    OrderCnt = 0
    
    With spdResult1
        For ii = 1 To .maxrows
            .Col = 1: .Row = ii
            If .Value = 1 Then
                .Col = 2
                If Len(Trim(.Text)) > 0 And .BackColor = vbWhite Then
                    comEQP.Output = ENQ
                    Debug.Print "[HOST] " & ENQ
                    txtStaCom.Text = txtStaCom.Text + "[HOST] " & ENQ
                    lblStatus = "����������.."
                    .BackColor = vbCyan
                    OrderCnt = ii
                    .Col = 3
                    .BackColor = vbCyan
                    .Col = 4
                    .BackColor = vbCyan
                    
                    Exit For
                End If
            End If
        Next
    End With
    
End Sub

Private Sub cmdRackNo_Click()

Dim sNo As String, sCnt As Integer, sAdd As Integer
Dim fNum1 As Integer, fNum2 As Integer
Dim intRow1 As Integer

AgainInput:
    fNum1 = 1: fNum2 = 0
    sNo = InputBox("���� ��ȣ�� �Է��ϼ��� !")
    If Len(sNo) > 0 And spdResult1.maxrows > 0 Then
        If Not IsNumeric(sNo) Then
            MsgBox "���ڸ� �Է��ϼ���.!", vbCritical
            GoTo AgainInput
        End If
        
        With spdResult1
            sAdd = 0
            For sCnt = .ActiveRow To .maxrows
                intRow1 = intRow1 + 1
                .Row = sCnt
                .Col = 1
                If .Value >= 1 Then
                    'If .ActiveCol = 3 Then
                        .Col = 3 '.ActiveCol
                        If intRow1 = (5 * fNum1) + 1 Then fNum1 = fNum1 + 1: fNum2 = 0
                        fNum2 = fNum2 + 1
                        .Text = Format(Trim((fNum1 + Val(sNo)) - 1), "0000")
                        .Col = 4 '.ActiveCol + 1
                        .Text = fNum2
                    'End If
                End If
            Next sCnt
        End With
'        spdResult1.StartingRowNumber = Val(sNo)
    End If
End Sub

Private Sub cmdRstQuery_Click()

    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String, intRet   As Integer
    
    Dim strSpcno    As String
    Dim intRow      As Integer, intCol  As Integer
    
    Dim strOrdcd()  As String, strPid() As String, strPnm() As String
    Dim itemX       As ListItem

    intRow = 0
    With spdResult2
        .maxrows = 14
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .maxrows
        .BlockMode = True
        .Action = ActionClearText
        .BackColor = vbWhite
        .BlockMode = False
    End With
    
    sqlDoc = "select SPCNO, TESTCD, EQUIPCD, TRANSTM, RSTVAL, REFVAL" & _
             "  from INTERFACE003" & _
             " where TRANSDT = '" & mskRstDate.Text & "'"
    If cboRstgbn(1).ListIndex = 0 Then
        sqlDoc = sqlDoc & "   and SERVERGBN = ''"
    ElseIf cboRstgbn(1).ListIndex = 1 Then
        sqlDoc = sqlDoc & "   and SERVERGBN = 'Y'"
    End If
    sqlDoc = sqlDoc & " order by SPCNO, TRANSTM"
    
    adoRS.CursorLocation = adUseClient
    adoRS.Open sqlDoc, AdoCn_Jet
    If adoRS.RecordCount > 0 Then adoRS.MoveFirst
    Do While Not adoRS.EOF
        With spdResult2
            If strSpcno <> Trim$(adoRS(0) & "") + Trim$(adoRS(3) & "") Then
                intRow = intRow + 1
                If intRow > .maxrows Then .maxrows = .maxrows + 1:  .RowHeight(.maxrows) = 13
                
                intRet = sl_spcid_tstcd_select&(Trim$(adoRS(0) & ""), strOrdcd, strPid, strPnm)

                .SetText 1, intRow, "1"
                .SetText 2, intRow, Trim$(adoRS(0) & "")
                .SetText 5, intRow, strPnm(0)
                .SetText 6, intRow, strPid(0)
                .SetText .MaxCols, intRow, Trim$(adoRS(3) & "")
            End If
            strSpcno = Trim$(adoRS(0) & "") + Trim$(adoRS(3) & "")
            Set itemX = lvwCuData.FindItem(Trim$(adoRS(1) & ""), lvwSubItem, , lvwWhole)
            If Not itemX Is Nothing Then
                intCol = itemX.Index + 6
                .SetText intCol, intRow, Trim$(adoRS(4)) & ""
                .Col = intCol:  .Row = intRow:  .ForeColor = IIf(Trim$(adoRS(5) & "") <> "", vbRed, vbBlack)
            End If
            
        End With
        adoRS.MoveNext
    Loop
    adoRS.Close:    Set adoRS = Nothing
    
End Sub

Private Sub cmdSel_Click(Index As Integer)

    Dim varTmp  As Variant
    Dim intRow  As Integer
    
    If Index = 2 Or Index = 3 Then
        With spdResult2
            For intRow = 1 To .maxrows
                .GetText 2, intRow, varTmp
                If Trim$(varTmp) <> "" Then .SetText 1, intRow, IIf(Index = 2, "1", "")
            Next
        End With
    Else
        With spdWorkList
            For intRow = 1 To .maxrows
                .GetText 2, intRow, varTmp
                If Trim$(varTmp) <> "" Then .SetText 1, intRow, IIf(Index = 0, "1", "")
            Next
        End With
    End If
    
End Sub

Private Sub cmdStartNo_Click()
Dim sNo As String, sCnt As Integer, sAdd As Integer

AgainInput:
    sNo = InputBox("���� ��ȣ�� �Է��ϼ��� !")
    If Len(sNo) > 0 And spdResult1.maxrows > 0 Then
        If Not IsNumeric(sNo) Then
            MsgBox "���ڸ� �Է��ϼ���.!", vbCritical
            GoTo AgainInput
        End If
        
        With spdResult1
            sAdd = 0
            For sCnt = .ActiveRow To .maxrows
                .Row = sCnt
                .Col = 0:       .Text = Trim(sAdd + Val(sNo))
                sAdd = sAdd + 1
            Next sCnt
        
            .StartingRowNumber = Val(sNo)
        End With
    End If
    
End Sub

Private Sub cmdWordQuery_Click()

    On Error GoTo ErrRoutine
    CallForm = "frmInterface - Privete sub cmdWorkQuery_Click()"
    
    Dim strKeyno    As String
    Dim strOrdcd()  As String, strPid() As String, strPnm() As String, strBarno()   As String
    Dim strTestcd() As String, strTPid()    As String, strTPnm() As String
    Dim strEqpCd    As String
    Dim intRow  As String, intIdx  As Integer, intCol   As Integer
    Dim itemX   As ListItem
    Dim ii As Integer
    Dim strLevel()  As String
    
    If chkQC.Value = 1 Then
        If Trim(cboLevel.Text) <> "" Then
            intIdx = sl_spcid_tstcd_select_qc1&(INS_CODE, cboLevel.Text, strBarno, strOrdcd, strLevel)
        Else
            lblStatus.Caption = "QC Level�� �����ϼ���!"
            cboLevel.SetFocus
            Exit Sub
        End If
    Else
        intIdx = sl_tstcd_spcid_select&(mskOrdDate.Text, "(" + f_strOrdList + ")", strBarno, strPid, strPnm, strOrdcd)
    End If
    
    intRow = 0
    For intIdx = 0 To UBound(strOrdcd) - 1
        With spdWorkList
            If strKeyno <> strBarno(intIdx) Then
                intRow = intRow + 1
                If intRow > .maxrows Then .maxrows = .maxrows + 1:  .RowHeight(.maxrows) = 13
                
                If chkQC.Value = 1 Then
                    .SetText 2, intRow, strBarno(intIdx)
                    .SetText 3, intRow, strOrdcd(intIdx)
                    .SetText 4, intRow, strLevel(intIdx)
                Else
                    .SetText 2, intRow, strBarno(intIdx)
                    .SetText 3, intRow, strPnm(intIdx)
                    .SetText 4, intRow, strPid(intIdx)
                End If
            End If
            If chkQC.Value = 1 Then
                intCol = sl_spcid_tstcd_select_qc&(INS_CODE, strBarno(intIdx), strTestcd)
            Else
                intCol = sl_spcid_tstcd_select(strBarno(intIdx), strTestcd, strTPid, strTPnm)
            End If
                
            strEqpCd = f_funget_code(strTestcd(ii))
            Set itemX = lvwCuData.FindItem(strEqpCd, lvwTag, , lvwWhole)
            If Not itemX Is Nothing Then .SetText 6 + itemX.Index, intRow, "V"
            Set itemX = Nothing
            strKeyno = strBarno(intIdx)
        End With
    Next
    Exit Sub
    
ErrRoutine:

    Call ErrMsgProc(CallForm)

End Sub

Private Sub cmdWorkList_Click()

    Dim varTmp  As Variant
    Dim intRow1 As Integer, intRow2 As Integer
    Dim intIdx  As Integer
    Dim fNum1   As Integer, fNum2  As Integer
    Dim Rev     As Long
    Dim Test_Cd() As String, strPid()   As String, strPnm() As String
    Dim ii As Integer
    Dim itemX As ListItem
'    Dim adoRS As ADODB.Recordset
'    Dim sqlDoc As String
    Dim bgetWork As Boolean
'    Dim iCol As Integer, iRow As Integer
    Dim strEqpCd As String
    
    bgetWork = False
    fNum1 = 1: fNum2 = 0
    With spdWorkList
        For intRow1 = 1 To .maxrows
            .GetText 1, intRow1, varTmp
            If Trim$(varTmp) = "1" Then
                .GetText 2, intRow1, varTmp
                intRow2 = f_funGet_SpreadRow(spdResult1, 2, Trim$(varTmp))
                If intRow2 < 1 Then
                    intRow2 = f_funGet_SpreadRow(spdResult1, 2, "")
                    If intRow2 < 1 Then
                        spdResult1.maxrows = spdResult1.maxrows + 1
                        spdResult1.RowHeight(spdResult1.maxrows) = 12
                        intRow2 = spdResult1.maxrows
                    End If
                    
                    If chkMan.Value = 0 Then
                        Rev = sl_spcid_tstcd_select(Trim(varTmp), Test_Cd, strPid, strPnm)
                        For ii = 0 To Rev - 1
                            strEqpCd = f_funget_code(Test_Cd(ii))
                            Set itemX = lvwCuData.FindItem(strEqpCd, lvwTag, , lvwWhole)
                            If Not itemX Is Nothing Then
                                bgetWork = True
                                spdResult1.Row = intRow2
                                spdResult1.Col = itemX.Index + 6
                                spdResult1.BackColor = &HC6FEFF
                                DoEvents
                            End If
                        Next
                    End If
                    
                    If chkMan.Value = 0 Then
                        If bgetWork = True Then
                            spdResult1.SetText 2, intRow2, Trim(varTmp)
                            If intRow1 > 1 Then
                                spdResult1.Row = intRow1 - 1
                                spdResult1.Col = 4
                                If spdResult1.Text = "5" Then
                                    fNum1 = fNum1 + 1
                                    fNum2 = 1
                                Else
                                    fNum2 = Val(spdResult1.Text) + 1
                                    spdResult1.Col = 3
                                    fNum1 = spdResult1.Text
                                End If
                            Else
                                spdResult1.SetText 2, intRow2, Trim(varTmp)
                                If intRow1 = (5 * fNum1) + 1 Then fNum1 = fNum1 + 1: fNum2 = 0
                                fNum2 = fNum2 + 1
                            End If
                            
                            If chkQC.Value = 0 Then
                                spdResult1.SetText 1, intRow2, "1"
                                spdResult1.SetText 3, intRow2, Format(Val(fNum1), "0000")
                                spdResult1.SetText 4, intRow2, Trim(fNum2)
                                spdResult1.SetText 5, intRow2, strPnm(0)
                                spdResult1.SetText 6, intRow2, strPid(0)
                            Else
                                spdResult1.SetText 1, intRow2, "1"
                                spdResult1.SetText 2, intRow2, Trim(fNum2)
                                spdResult1.SetText 3, intRow2, "QC�˻�"
                                spdResult1.SetText 4, intRow2, "Level : " & strPid(0)
                            End If
                        Else
                            spdResult1.maxrows = spdResult1.maxrows - 1
                        End If
                        bgetWork = False
                    Else
                        spdResult1.SetText 2, intRow2, Trim(varTmp)
                        If intRow1 > 1 Then
                            spdResult1.Row = intRow1 - 1
                            spdResult1.Col = 4
                            If spdResult1.Text = "5" Then
                                fNum1 = fNum1 + 1
                                fNum2 = 1
                            Else
                                fNum2 = Val(spdResult1.Text) + 1
                                spdResult1.Col = 3
                                fNum1 = spdResult1.Text
                            End If
                        Else
                            spdResult1.SetText 2, intRow2, Trim(varTmp)
                            If intRow1 = (5 * fNum1) + 1 Then fNum1 = fNum1 + 1: fNum2 = 0
                            fNum2 = fNum2 + 1
                        End If
                        
                        If chkQC.Value = 0 Then
                            spdResult1.SetText 1, intRow2, "1"
                            spdResult1.SetText 3, intRow2, Format(Val(fNum1), "0000")
                            spdResult1.SetText 4, intRow2, Trim(fNum2)
                            spdResult1.SetText 5, intRow2, "" 'strPnm(0)
                            spdResult1.SetText 6, intRow2, "" 'strPid(0)
                        Else
                            spdResult1.SetText 1, intRow2, "1"
                            spdResult1.SetText 2, intRow2, Trim(fNum2)
                            spdResult1.SetText 3, intRow2, "QC�˻�"
                            spdResult1.SetText 4, intRow2, "Level : " & strPid(0)
                        End If
                    End If
                End If
                spdResult1.maxrows = intRow2
                
                .Row = intRow1: .Col = 1
                .Value = 0
            End If
        Next
    End With
    
    '-- �ӽ�
    'comEQP.Output = STX + "M     " + ETX
    
    f_strJOB_FLAG = "1"
    f_intSampleNo = 0
    
End Sub


Private Sub comEQP_OnComm()
    
    Dim strEVMsg    As String
    Dim strERMsg    As String
    Dim Arr()       As Byte
    Dim strdata     As String
    Dim brStr As String
    Dim sStxCheck As Integer, sEtxCheck As Integer, sCrcheck As Integer
    Dim com_sTemp As String
    Dim ii As Integer, jj As Integer
    Dim MHead  As String, Pinfo As String
    Dim PatientID As String
    
    Dim Orderoutput As String
    Dim OutPutData  As String
    Dim Rev As Long
    Dim Test_Cd() As String, strPid()    As String, strPnm() As String
    Dim sRow As Integer
    Dim oPatNo As String
    Dim oRackNo As String
    Dim oPosNo As String
    
    Dim adoRS As ADODB.Recordset
    Dim sqlDoc As String
    Dim itemX As ListItem
    Dim strEqpCd1 As String
    
    Dim varTmp  As Variant  '-- yej
    Dim intCol  As Integer  '-- yej
    
    Select Case comEQP.CommEvent
        Case comEvReceive
            imgReceive.Picture = imlStatus.ListImages("RUN").ExtractIcon
            If tmrReceive.Enabled = False Then
                tmrReceive.Enabled = True
            Else
                tmrReceive.Enabled = False
                tmrReceive.Enabled = True
            End If
            brStr = ""
            brStr = comEQP.Input
            
            txtStaCom.Text = txtStaCom.Text + brStr
            
            If left$(brStr, 1) = ENQ Then
                Debug.Print "[ELEC] " & ENQ
                comEQP.Output = ACK
                Debug.Print "[HOST] " & ACK
                fRcvString = ""
                Exit Sub
            End If
            
            For ii = 1 To Len(brStr)
                fRcvString = fRcvString + Mid(brStr, ii, 1)
            Next ii
            
            spdResult1.Col = 1: spdResult1.Row = OrderCnt
            If Len(Trim(spdResult1.Text)) > 0 Then
                '-- Ordering
                If left$(brStr, 1) = ACK Then
                    Debug.Print "[ELEC] " & ACK
                    Select Case SendCount
                        Case 0      'Message Header
                            MHead = "1H|\^&|||ASTM-Host"
                            comEQP.Output = STX & MHead & vbCr & ETX & MakeCS(MHead) & vbCr & vbLf
                            Debug.Print "[HOST] " & STX & MHead & vbCr & ETX & MakeCS(MHead) & vbCr & vbLf
                            SendCount = SendCount + 1
                            MHead = ""
                        Case 1      'patient information
                            Pinfo = "2P|1"
                            comEQP.Output = STX & Pinfo & vbCr & ETX & MakeCS(Pinfo) & vbCr & vbLf
                            Debug.Print "[HOST] " & STX & Pinfo & vbCr & ETX & MakeCS(Pinfo) & vbCr & vbLf
                            SendCount = SendCount + 1
                            Pinfo = ""
                        Case 2      'Test Order
                            SendCount = SendCount + 1
                            spdResult1.Row = OrderCnt
                            spdResult1.Col = 2
                            Rev = sl_spcid_tstcd_select(Trim(spdResult1.Text), Test_Cd, strPid, strPnm)
                            oPatNo = spdResult1.Text
                            
                            spdResult1.Col = 3: oRackNo = spdResult1.Text
                            spdResult1.Col = 4: oPosNo = spdResult1.Text
                            
                            If Rev = 0 And chkMan.Value = 0 Then
                                lblStatus.Caption = oPatNo & "�� �ش� �˻��׸��� �����ϴ�.!"
                                comEQP.Output = EOT
                                Debug.Print "[HOST] " & EOT
                                OrderCnt = 0
                                SendCount = 0
                                Exit Sub
                            ElseIf chkMan.Value = 0 Then
                                Orderoutput = ""
                                For jj = 0 To Rev - 1
                                    If Len(Test_Cd(jj)) > 0 Then
                                        strEqpCd1 = f_funget_code(Test_Cd(jj))
                                        Set itemX = lvwCuData.FindItem(strEqpCd1, lvwTag, , lvwWhole)
                                        If Not itemX Is Nothing Then
                                            spdResult1.Col = itemX.Index + 6
                                            If spdResult1.BackColor = &HC6FEFF Then
                                                Orderoutput = Orderoutput + "^^^" + strEqpCd1 + "^0\"
                                            End If
                                            'DoEvents
                                        End If
                                    End If
                                Next jj

                                Orderoutput = Mid(Orderoutput, 1, Len(Orderoutput) - 1)
                            End If
                            
                            'If chkMan.Value = 1 And Orderoutput = "" Then
                            If chkMan.Value = 1 Then
                                Orderoutput = ""
                                For intCol = 7 To spdResult1.MaxCols
                                    spdResult1.GetText intCol, 0, varTmp
                                    If Trim$(varTmp) = "" Then Exit For
                                    Set itemX = lvwCuData.FindItem(Trim$(varTmp), lvwSubItem, , lvwWhole)
                                    If Not itemX Is Nothing Then
                                        spdResult1.Col = intCol:    spdResult1.Row = OrderCnt
                                        If spdResult1.BackColor = &HC6FEFF Then
                                            Orderoutput = Orderoutput + "^^^" + itemX.tag + "^0\"
                                        End If
                                    End If
                                    Set itemX = Nothing
                                Next
                            End If
                            
                            Orderoutput = "3O|1|" + oPatNo + "|^" + oRackNo + "^" + oPosNo + "^^SAMPLE^NORMAL|" + Orderoutput + "|R||||||N||||||||||||||O"
                            OutPutData = STX & Orderoutput & vbCr & ETX & MakeCS(Orderoutput) & vbCr & vbLf
                            comEQP.Output = OutPutData
                            'MsgBox OutPutData
                            Debug.Print "[HOST] " & OutPutData
                        Case 3      'Message Terminator
                            SendCount = SendCount + 1
                            comEQP.Output = STX & "4L|1|N" & vbCr & ETX & "07" & vbCr & vbLf
                            Debug.Print "[HOST] " & STX & "4L|1|N" & vbCr & ETX & "07" & vbCr & vbLf
                        Case Else
                            comEQP.Output = EOT
                            Debug.Print "[HOST] " & EOT
                            '-- ���������� �����ִ��� üũ
                            With spdResult1
                                For ii = OrderCnt To .maxrows
                                    .Col = 1: .Row = ii
                                    If .Value = 1 Then
                                        .Col = 2
                                        If .BackColor = vbWhite Then
                                            .BackColor = vbCyan
                                            .Col = 3
                                            .BackColor = vbCyan
                                            .Col = 4
                                            .BackColor = vbCyan
                                            OrderCnt = OrderCnt + 1
                                            .Col = 2
                                            If Len(Trim(.Text)) > 0 Then
                                                '-- osw edit
                                                .Row = ii '+ 1
                                                If Len(Trim(.Text)) > 0 Then
                                                    comEQP.Output = ENQ
                                                    Debug.Print "[HOST] " & ENQ
                                                End If
                                                SendCount = 0
                                                Exit For
                                            End If
                                        End If
                                    End If
                                    Me.Enabled = True
                                Next
                            End With
                            SendCount = 0
    '                        Flag_HQL = ""
                    End Select
'                    fRcvString = ""
                End If
            End If
                        
            sStxCheck = InStr(fRcvString, STX)
            sEtxCheck = InStr(fRcvString, ETX)
            sCrcheck = InStr(fRcvString, vbCr)
            
            If sStxCheck <> 0 And sEtxCheck <> 0 And sCrcheck <> 0 Then
                Call psDataDefine(fRcvString, fChannel(), spdResult1)
                Debug.Print fRcvString
                fRcvString = ""
            End If
            
        Case comEvSend
        
            imgSend.Picture = imlStatus.ListImages("RUN").ExtractIcon
            If tmrSend.Enabled = False Then
                tmrSend.Enabled = True
            Else
                tmrSend.Enabled = False
                tmrSend.Enabled = True
            End If
        Case comEvCTS
            strEVMsg = " CTS(Clear to Send) ���� ����"
        Case comEvDSR
            strEVMsg = " DSR(Data Set Read) ���� ����"
        Case comEvCD
            strEVMsg = " CD(Carrier Detecr) ���� ����"
        Case comEvRing
            strEVMsg = " ��ȭ ���� �︮�� ��"
        Case comEvEOF
            strEVMsg = " EOF(End Of File) ����"

        ' ���� �޽���
        Case comBreak
            strERMsg = " �ߴ� ��ȣ ����"
        Case comCDTO
            strERMsg = " �ݼ��� ���� �ð� �ʰ�"
        Case comCTSTO
            strERMsg = " CTS(Clear to Send) �ð� �ʰ�"
        Case comDCB
            strERMsg = " ��Ʈ�� ���� ��ġ ���� ���(DCB) �˻� �� ����ġ ���� ����"
        Case comDSRTO
            strERMsg = " DSR(Data Set Read) �ð� �ʰ�"
        Case comFrame
            strERMsg = " �����̹� ����"
        Case comOverrun
            strERMsg = " �и�Ƽ ����"
        Case comRxOver
            strERMsg = " ���� ���� �ʰ�"
        Case comRxParity
            strERMsg = " �и�Ƽ ����"
        Case comTxFull
            strERMsg = " ���� ���ۿ� ������ ����"
        Case Else
            strERMsg = " �� �� ���� ���� �Ǵ� �̺�Ʈ"
    End Select
    If Len(strERMsg) > 0 Then Call ShowMessage(strERMsg)
End Sub

Public Sub OrderingTheDataElecsys(ByVal brCom As Object, ByVal brbarcd As String, ByVal brSpread As Object, ByRef brChannel() As String) ', ByRef brItemdeci() As String)
Dim Orderoutput As String
Dim OutPutData As String
Dim Testcd As String, sOrderLst As String
Dim ii As Integer, Loop_count As Integer, pDoCount As Integer
Dim rv As Long
Dim tst_cd() As String

    For Loop_count = 1 To 100: fElecsys2010(Loop_count) = "": Next Loop_count

    pDoCount = 0
    '-- ��¥�� ��ü��ȣ�� ������ ��ȸ�Ѵ�.
'    rv = sl_spcid_tstcd_select(PatientID, tst_cd)
'    If rv = 0 Then
'        MsgBox "�ش��ü ����"
'    Else
'        ReDim tst_cd(rv)
'        For II = 1 To rv
'            sOrderLst = sOrderLst + "," + tst_cd(II)
'        Next
'    End If
'    sOrderLst = sOrderLst + ","
'
'    Do While InStr(sOrderLst, ",") > 0
'        pDoCount = pDoCount + 1
'        fElecsys2010(pDoCount) = Text_Redefine(sOrderLst, ",")
'        sOrderLst = Mid$(sOrderLst, InStr(sOrderLst, ",") + 1)
'        If pDoCount > 99 Then
'            sOrderLst = ""
'            Exit Do
'        End If
'    Loop
    
'    For II = 1 To pDoCount
'        If Val(fElecsys2010(II)) > 0 Then
'            OutPutData = OutPutData + "^^^" & fElecsys2010(II) & "^" & "0"
'            If II = pDoCount Then
'                Exit For
'            Else
'                OutPutData = OutPutData + "\"
'            End If
'        End If
'    Next II
    '-- �ӽ� �˻��׸�
    OutPutData = "^^^250^0\^^^10^0"
    
'    Orderoutput = "3O" & "|1|" & PatientID & "|" & PatientSeq & "|" & OutPutData & "|R|" & Format(Now, "YYYYMMDDHHMMSS") & "|||||N||||||||||||||Q"
'    OutPutData = STX & Orderoutput & vbCr & ETX & MakeCS(Orderoutput) & vbCr & vbLf
    
    '-- �ӽ�
    Orderoutput = "3O" & "|1|" & "0001" & "|" & "011" & "|" & OutPutData & "|R|" & Format(Now, "YYYYMMDDHHMMSS") & "|||||N||||||||||||||Q"
    OutPutData = STX & Orderoutput & vbCr & ETX & MakeCS(Orderoutput) & vbCr & vbLf
    
    comEQP.Output = OutPutData
    Debug.Print "[HOST] " & OutPutData
    OutPutData = ""
    PatientID = ""
    PatientSeq = ""
    PatientDisk = ""
    PatientPos = ""
    
End Sub

Private Function MakeCS(Source As String) As String
    Dim x      As Long
    Dim ChkCS  As String
    Dim SumCS  As String
    Dim AddCS  As Long
    
    For x = 1 To Len(Source)
        AddCS = AddCS + Asc(Mid(Source, x, 1))
    Next x
    AddCS = AddCS + Asc(Chr(13)) + Asc(ETX)
    AddCS = AddCS Mod &H100
    SumCS = Hex(AddCS)
    If Len(SumCS) = 1 Then
        ChkCS = "0" & SumCS
    Else
        ChkCS = Mid(SumCS, Len(SumCS) - 1, 1)
        ChkCS = ChkCS & Right(SumCS, 1)
    End If
    MakeCS = ChkCS
End Function

Private Sub ComReceive(ByRef RecData() As Byte)
    
    Dim strRec  As String, strBuff  As String
    
    Dim strTmp  As String, strDta() As String
    Dim intIdx  As Integer, intCnt  As Integer, intPos  As Integer
    Dim sStxCheck As Integer, sEtxCheck As Integer
    Dim com_sTemp As String
    Dim fOpt        As String

    Static OrgMsg As String
    
'    strRec = StrConv(RecData, vbUnicode)
    strRec = RecData
    Print #1, strRec;
    'Debug.Print RecData
    
    Call COM_INPUT(strRec)
    
    For intIdx = 1 To Len(strRec)
        strBuff = Mid$(strRec, intIdx, 1)
        Select Case f_strJOB_FLAG
            Case "1"    '-- ���
                        Select Case Asc(strBuff)
                            Case 2  '-- STX
                                    f_strJOB_FLAG = "2"
                        End Select
                        f_strBuffer = f_strBuffer + strBuff
            Case "2"    '--  �ޱ�
                        Select Case Asc(strBuff)
                            Case 2      '-- STX
                                        f_strJOB_FLAG = "2"
                            Case 3      '-- ETX
                                        f_strBuffer = f_strBuffer + strBuff
                                        Debug.Print "dsfdsfdsf" + f_strBuffer
                                        'strTmp = f_strBuffer
                                        sStxCheck = InStr(f_strBuffer, Chr(2))
                                        sEtxCheck = InStr(f_strBuffer, Chr(3))
                                        If sStxCheck <> 0 And sEtxCheck <> 0 Then
                                            com_sTemp = Mid$(f_strBuffer, sStxCheck, sEtxCheck)
                                            f_strBuffer = Mid$(f_strBuffer, sEtxCheck + 1)
                                            If optSeq.Value = True Then
                                                fOpt = "0"
                                            Else
                                                fOpt = "1"
                                            End If
                                            'Call psDataDefine(com_sTemp, fChannel(), spdResult1, fOpt) ', brSpread, brChannel(), brItemdeci(), brOpt
                                        End If
                                        
                                        
                                        Call COM_OUTPUT(Chr(6))
                                        f_strBuffer = ""
                                        f_strJOB_FLAG = "1"
                            
                            Case Else
                                        f_strBuffer = f_strBuffer + strBuff
                        End Select
        End Select
     Next
End Sub

Private Sub psDataDefine(ByVal brbarcd As String, ByRef brChannel() As String, ByVal brSpread As Object)

Dim sTemp       As String       ' On Com���κ��� �Ѱܹ��� Receive Data
Dim Channel_No  As String       ' ������ ����
Dim Patiant_No  As String       ' ȯ�ڹ�ȣ
Dim pGrid_Point As Integer      ' �ش� �˻��� Point
Dim Max_Arary_Cnt As Integer    ' �˻� �׸��
'-------------------------------' �ӽ� ������.....
Dim sDeCnt      As Integer
Dim pDoCount    As Integer
Dim Loop_count  As Integer
Dim sRtn As Integer, sChannel As String, sRstText As String, sRstValue As Single, sUnit As String
Dim sPatiant_No As String
Dim itemX As ListItem
Dim strRstval(1 To 19) As String, strRefVal(1 To 19)  As String
Dim FunStr As String
Dim sqlDoc  As String
Dim intCol As Integer
Dim Test_Cd() As String, strPid()    As String, strPnm() As String
Dim Rev As Long
Dim ii As Integer
Dim tmpTstCd As String

Dim strDate As String, strTime  As String, sqlRet   As Integer

    On Error GoTo errDefine
    sRstText = brbarcd
    '------------------------------<<< fElecsys2010() �迭 Clear �Ѵ�.         >>>----------
    For Loop_count = 1 To 100: fElecsys2010(Loop_count) = "": Next Loop_count
    '------------------------------<<< fElecsys2010() �迭�� �����Ͽ� �ִ´�.  >>>----------
        
    pDoCount = 0
'    sRstText = Mid(sRstText, STX)
    sRstText = Mid(sRstText, InStr(fRcvString, STX))
    Do While InStr(sRstText, "|") > 0
        pDoCount = pDoCount + 1
        fElecsys2010(pDoCount) = Text_Redefine(sRstText, "|")
        sRstText = Mid$(sRstText, InStr(sRstText, "|") + 1)   ' �����ڰ� "|" �̴�....
        If pDoCount > 99 Then
            sRstText = ""
            Exit Do
        End If
    Loop
   
    sRstText = ""
    If Mid$(fElecsys2010(1), 3, 1) = "H" Then          ' "H" Head Message Display
        comEQP.Output = ACK
    ElseIf Mid$(fElecsys2010(1), 3, 1) = "P" Then      ' "P" Patiant Information Data Process
        comEQP.Output = ACK
    ElseIf Mid$(fElecsys2010(1), 3, 1) = "C" Then
        comEQP.Output = ACK
    ElseIf Mid$(fElecsys2010(1), 3, 1) = "Q" Then      ' "Q" Patiant Information Data Process
        comEQP.Output = ACK
    ElseIf Mid$(fElecsys2010(1), 3, 1) = "O" Then      ' "O" Order Data Process
        comEQP.Output = ACK
        PatientID = fElecsys2010(3)
        pDoCount = 0
        Do While InStr(fElecsys2010(4), "^") > 0
            pDoCount = pDoCount + 1
            Select Case pDoCount
                Case 1:    PatientSeq = Text_Redefine(fElecsys2010(4), "^")
                Case 2:    PatientRack = Text_Redefine(fElecsys2010(4), "^")
                Case 3:    PatientPos = Text_Redefine(fElecsys2010(4), "^")
                Case Else: Exit Do
            End Select
            fElecsys2010(4) = Mid$(fElecsys2010(4), InStr(fElecsys2010(4), "^") + 1)   ' �����ڰ� "^" �̴�....
        Loop

        Patiant_Recevid = False        ' ȯ�ڹ�ȣ Flag
        sPatiant_No = fElecsys2010(3)  ' ȯ�ڹ�ȣ
        '-------------------------------------------<<< �ش�˻����� �ش�ȯ�ڸ� �O�´�.       >>>----------
        With brSpread
            For pDoCount = 1 To .maxrows
                .Row = pDoCount: .Col = 2
                If Trim$(.Text) = sPatiant_No Then
                    vRow = pDoCount
                    Patiant_Recevid = True
                    Exit For
                End If
            Next pDoCount
        End With

    ElseIf Mid$(fElecsys2010(1), 3, 1) = "R" Then      ' "R" Result Data Process
        comEQP.Output = ACK
        Debug.Print "[HOST] " & ACK
        If Patiant_Recevid = True Then
            fElecsys2010(3) = Replace(fElecsys2010(3), "^", "")    ' channel
            fElecsys2010(3) = fElecsys2010(3) / 10
            Channel_No = Val(fElecsys2010(3))                                   ' channel
            '  fElecsys2010(4)                                                  ' result
            '-------------------------------------------<<< �ش�˻����� �O�´�.       >>>----------
            Max_Arary_Cnt = brSpread.MaxCols - 6   ' �տ������� 5������ ȯ�� ���� �̱⶧����.... -6�� �Ѵ�.
                                                   ' �ش� �迭��  brItem(),brChannel() �̴�.
            With brSpread
                '----------------------------------------------<<<<<<<<<,  ���ΰ˻��׸��� �O�´�.  >>>>>>>----------
                For pDoCount = 1 To Max_Arary_Cnt
                    .Row = vRow
                    .Col = pDoCount + 6
                    If Channel_No > 0 And Channel_No = Val("" & brChannel(pDoCount)) Then          ' �˻����� ������...
                        If Trim(fElecsys2010(4)) <> "" Then
                            .Text = Trim(fElecsys2010(4))
                        Else
                            .Text = ""
                        End If
                        
                        If .Text <> "" Then
                            Set itemX = lvwCuData.FindItem(Channel_No, lvwTag, , lvwWhole)
                            If Not itemX Is Nothing Then '-- itemX : �˻��ڵ�
                                If itemX.ListSubItems(8) <> "" And itemX.ListSubItems(9) <> "" Then
                                    If Val(.Text) < itemX.ListSubItems(8) Then
                                        strRefVal(pDoCount) = "L"
                                    ElseIf Val(.Text) > itemX.ListSubItems(9) Then
                                        strRefVal(pDoCount) = "H"
                                    End If
                                End If
                                
                                If chkQC.Value = 1 Then
                                    Rev = sl_spcid_tstcd_select_qc1&(INS_CODE, cboLevel.Text, strTPid, Test_Cd, strLevel)
                                Else
                                    Rev = sl_spcid_tstcd_select&(Trim(PatientID), Test_Cd, strPid, strPnm)
                                End If
                                
                                For ii = 1 To Rev
                                    If InStr(itemX.ListSubItems(1), Trim(Test_Cd(ii))) > 0 Then
                                          tmpTstCd = "" & Trim(Test_Cd(ii))
                                          Exit For
                                    End If
                                Next
                                    
                                strDate = Format$(Now, "YYYYMMDD"): strTime = Format$(Now, "MMSS")
                                sqlDoc = "insert into INTERFACE003(" & _
                                         "            SPCNO, TESTCD, EQPNUM, TRANSDT, TRANSTM, RSTVAL, REFVAL, EQUIPCD, SERVERGBN)" & _
                                         "    values( '" & PatientID & "', '" & itemX.ListSubItems(1) & "', '" & Channel_No & "'," & _
                                         "            '" & strDate & "', '" & strTime & "'," & _
                                         "            '" & Trim(fElecsys2010(4)) & "', '" & strRefVal(pDoCount) & "'," & _
                                         "            '" & INS_CODE & "', '')"
                                AdoCn_Jet.Execute sqlDoc
                                
                                intCol = itemX.Index + 6
                                .Col = intCol
                                .ForeColor = IIf(strRefVal(pDoCount) <> "", vbRed, vbBlack)
                                    
                                '-- ����������
                                If Rev > 0 And chkAuto.Value = vbChecked And chkQC.Value = 0 Then
                                    If f_funAdd_Server(PatientID, itemX.ListSubItems(1), Trim(fElecsys2010(4)), Test_Cd) Then
                                        .Col = -1:  .BackColor = &HFFF8F0
                                    
                                        sqlDoc = "Update INTERFACE003 set SERVERGBN = 'Y'" & _
                                                 " where SPCNO   = '" & PatientID & "'" & _
                                                 "   and EQPNUM  = '" & Channel_No & "'" & _
                                                 "   and TRANSDT = '" & strDate & "'" & _
                                                 "   and TRANSTM = '" & strTime & "'"
                                        AdoCn_Jet.Execute sqlDoc, sqlRet
                                    End If
                                Else
                                    If f_funAdd_QcServer(PatientID, itemX.ListSubItems(1), Trim(fElecsys2010(4)), Test_Cd) Then
                                        .Col = -1:  .BackColor = &HFFF8F0
                                    
                                        sqlDoc = "Update INTERFACE003 set SERVERGBN = 'Y'" & _
                                                 " where SPCNO   = '" & PatientID & "'" & _
                                                 "   and EQPNUM  = '" & Channel_No & "'" & _
                                                 "   and TRANSDT = '" & strDate & "'" & _
                                                 "   and TRANSTM = '" & strTime & "'"
                                        AdoCn_Jet.Execute sqlDoc, sqlRet
                                    End If
                                End If
                                
                            End If
                            Set itemX = Nothing
                        End If
                    End If

                Next pDoCount
            End With
        End If
    ElseIf Mid$(fElecsys2010(1), 3, 1) = "L" Then      ' "L" Data Last
        comEQP.Output = ACK
        Debug.Print "[HOST] " & ACK
        Flag_HQL = Flag_HQL & Mid(brbarcd, 3, 1)
        Patiant_Recevid = False                        ' ȯ�� ��ȣ  Flag
    End If
                        
    Exit Sub
errDefine:

End Sub

Public Function Text_Redefine(FSend_Str As String, FCheck_Char As String) As String
    If InStr(FSend_Str, FCheck_Char) > 0 Then
        Text_Redefine = left$(FSend_Str, InStr(FSend_Str, FCheck_Char) - 1)
    Else
        Text_Redefine = FSend_Str
    End If
    
End Function

Public Function SeqSearch(ByVal brSpread As Object, ByVal brSeq As Long, ByVal brCol As Integer) As Long
Dim sCnt As Long

    SeqSearch = 0
    If brSpread.maxrows <= 0 Then
        Exit Function
    End If
    
    With brSpread
        For sCnt = 1 To .maxrows
            .Row = sCnt
            .Col = brCol
            If Val(.Text) = brSeq Then
                SeqSearch = sCnt
                .Action = ActionActiveCell
                .Refresh
                Exit For
            End If
        Next sCnt
    End With

End Function

Public Function f_funGet_CheckSum(ByVal strPara As String) As String

    Dim intIdx      As Integer
    Dim intChkSum   As Integer
    
    intChkSum = 0
    For intIdx = 1 To Len(strPara)
        intChkSum = intChkSum + (0 Xor Asc(Mid$(strPara, intIdx, 1)))
    Next
    
    f_funGet_CheckSum = Chr(intChkSum) '-Format$(Hex(intChkSum), "00")
        
End Function

Private Sub Command1_Click()

    Dim intIdx          As Integer
    Dim strdata(1 To 5)         As Byte
    
'                 strData(1) = "D1U   XE-2100^A10930000000073000             830107231544000000090302                0000000000000000000000000000000000000000000000000000000000000000000000000000000000XE-2100^99337319^A1093"
'    strData(1) = strData(1) & "D2U   XE-2100^A10930000000073000             8300461004120013800400009710033500345001950         011900424001250010700304000000000000000000000000000000000210           00000000000000000XE-2100^99337319^A1093"
'D1U   XE-2100^A10930000000072000             820107231543000000090202                0000000000000000000000000000000000000000000000000000000000000000000000000000000000XE-2100^99337319^A1093
'D2U   XE-2100^A10930000000072000             8200510004290013700412009600031900333002500         013000452001150010100264000000000000000000000000000000000250           00000000000000000XE-2100^99337319^A1093
'D1U   XE-2100^A10930000000071000             810107231543000000090102                0000000000000000000000000000000000000000000000000000000000000000000000000000000000XE-2100^99337319^A1093
'D2U   XE-2100^A10930000000071000             8100422004160012300378009090029600325001690         012800423001450011800401000000000000000000000000000000000200           00000000000000000XE-2100^99337319^A1093
'D1U   XE-2100^A10930000000070000             800107231543000000081002            0000000000000000000000000000000000000000000000000000000000000000000000000000000000XE-2100^99337319^A1093
'D2U   XE-2100^A10930000000070000             8000492004340013700403009290031600340001740         01250042500108000970022000000000000000000000000000000000017000000000000000000XE-2100^99337319^A1093
'D1U   XE-2100^A10930000000073000             830107231544000000090302                0000000000000000000000000000000000000000000000000000000000000000000000000000000000XE-2100^99337319^A1093
'D2U   XE-2100^A10930000000073000             8300461004120013800400009710033500345001950         011900424001250010700304000000000000000000000000000000000210           00000000000000000XE-2100^99337319^A1093
'D1U   XE-2100^A10930000000072000             820107231543000000090202                0000000000000000000000000000000000000000000000000000000000000000000000000000000000XE-2100^99337319^A1093
'D2U   XE-2100^A10930000000072000             8200510004290013700412009600031900333002500         013000452001150010100264000000000000000000000000000000000250           00000000000000000XE-2100^99337319^A1093
'D1U   XE-2100^A10930000000071000             810107231543000000090102                0000000000000000000000000000000000000000000000000000000000000000000000000000000000XE-2100^99337319^A1093
'D2U   XE-2100^A10930000000071000             8100422004160012300378009090029600325001690     012800423001450011800401000000000000000000000000000000000200           00000000000000000XE-2100^99337319^A1093
'D1U   XE-2100^A10930000000070000             800107231543000000081002            0000000000000000000000000000000000000000000000000000000000000000000000000000000000XE-2100^99337319^A1093
'D2U   XE-2100^A10930000000070000             8000492004340013700403009290031600340001740         01250042500108000970022000000000000000000000000000000000017000000000000000000XE-2100^99337319^A1093
    
    Call comEQP_OnComm
    
End Sub

Private Sub Form_Activate()
    
    If IS_SET = False Then Unload Me

End Sub

Private Sub Form_Load()
    
'    Me.Show
    imgPort.Picture = imlStatus.ListImages("NOT").ExtractIcon
    imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
    imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
    
    CaptionBar1.Caption = INS_NAME & " Communication"
    
    Call cmdClear               ' �ʱ�ȭ
    Call f_subSet_ItemHeader    ' ����Ʈ�ش�
    Call f_subSet_ItemList      ' �˻��׸�
    
    Call f_subSet_ComCharacter  ' ��Ź���
    Call f_subGet_Setting       ' ��ż���
    
    Call cmdRun           ' ����
    
    mskRstDate.Text = Format$(Now, "YYYYMMDD")
    mskOrdDate.Text = Format$(Now, "YYYYMMDD")
    Open App.Path + "\" + "dump_job.log" For Append As #1

    f_strJOB_FLAG = "1":    f_intSampleNo = 0
    cboRstgbn(0).ListIndex = 0: cboRstgbn(1).ListIndex = 2
    SendCount = 0
    
    tabWork.Tab = 0
    
    cboLevel.Clear
    cboLevel.AddItem "H"
    cboLevel.AddItem "M"
    cboLevel.AddItem "L"

End Sub

Private Sub f_subGet_Setting()
    
    Dim objComSetting As clsCommon
    Dim Baudratio As String
    Dim Paritybit As String
    Dim Databit As String
    Dim Stopbit As String
    
    On Error GoTo ErrRoutine
    CallForm = "frmInterface - Private Sub f_subGet_Setting()"
    Set objComSetting = New clsCommon
    
    With objComSetting
        .SetAdoCn AdoCn_Jet
        Set mAdoRs = .Get_EqpProperty(INS_CODE)
    End With
    Set objComSetting = Nothing
    
    If mAdoRs Is Nothing Then
        IS_SET = False
        MsgBox INS_CODE & " �� ���� ��� ��� ������ �����ϴ�. ��� ������ �ٽ� �õ� �Ͻʽÿ�.", vbExclamation
        Exit Sub
    Else
        If mAdoRs.EOF Then
            IS_SET = False
            MsgBox INS_CODE & " �� ���� ��� ��� ������ �����ϴ�. ��� ������ �ٽ� �õ� �Ͻʽÿ�.", vbExclamation
            Set mAdoRs = Nothing
            Exit Sub
        Else
            IS_SET = True
            Baudratio = Trim(mAdoRs.Fields("COM_SPEED") & "")
            Paritybit = Trim(mAdoRs.Fields("COM_PARITYBIT") & "")
            Databit = Trim(mAdoRs.Fields("COM_DATABIT") & "")
            Stopbit = Trim(mAdoRs.Fields("COM_STOPBIT") & "")
            
            With comEQP
                .CommPort = Trim(mAdoRs.Fields("COM_PORT") & "")
                .Handshaking = Trim(mAdoRs.Fields("COM_HANDSHAK") & "")
'                .InputMode = Trim(mAdoRs.Fields("COM_INPUTMOD") & "")
'                .DTREnable = Trim(mAdoRs.Fields("COM_DTR") & "")
'                .EOFEnable = Trim(mAdoRs.Fields("COM_EOF") & "")
'                .NullDiscard = Trim(mAdoRs.Fields("COM_NULDIS") & "")
'                .RTSEnable = Trim(mAdoRs.Fields("COM_RTS") & "")
'                .InBufferSize = Trim(mAdoRs.Fields("COM_IBS") & "")
'                .InputLen = Trim(mAdoRs.Fields("COM_INLEN") & "")
'                .OutBufferSize = Trim(mAdoRs.Fields("COM_OBS") & "")
'                .ParityReplace = Trim(mAdoRs.Fields("COM_PTR") & "")
                .RThreshold = 1 'Trim(mAdoRs.Fields("COM_RTH") & "")
                .SThreshold = 1 'Trim(mAdoRs.Fields("COM_STH") & "")
                .Settings = Baudratio & "," & Paritybit & "," & Databit & "," & Stopbit
            End With
            Call Del_OldData
        End If
    End If
    
    Set mAdoRs = Nothing
Exit Sub

ErrRoutine:
    Set objComSetting = Nothing
    Set mAdoRs = Nothing
    Call ErrMsgProc(CallForm)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Call cmdStop
    Set Result = Nothing
    
    Close #1

End Sub

Private Sub imgPort_DblClick()
    
    If lvwCuData.Visible Then
        lvwCuData.Visible = False
    Else
        lvwCuData.Visible = True
        lvwCuData.ZOrder 0
    End If
    
End Sub

Private Sub imgReceive_DblClick()

    If pnlCom2.Visible = True Then
        pnlCom2.Visible = False
    Else
        pnlCom2.Visible = True
        pnlCom2.ZOrder 0
    End If
    
End Sub

Private Sub imgSend_DblClick()
    
    If pnlCom.Visible = True Then
        pnlCom.Visible = False
    Else
        pnlCom.Visible = True
        pnlCom.ZOrder 0
    End If

End Sub

Private Sub Label9_DblClick()

    If COM_MODE = "1" Then
        COM_MODE = "0"
        ShowMessage "�������̽� ������ ȭ�鿡 ������� �ʽ��ϴ�."
        txtStaCom.Visible = False
    Else
        COM_MODE = "1"
        ShowMessage "�������̽� ������ ȭ�鿡 ����մϴ�."
        txtStaCom.Visible = True
    End If
End Sub

Private Sub mskOrdDate_GotFocus()

    With mskOrdDate
        .SelStart = 8
        .SelLength = Len(.Text)
    End With
    
End Sub


Private Sub mskOrdDate_KeyPress(KeyAscii As Integer)

    If Not KeyAscii = vbKeyBack Then mskOrdDate.SelLength = 1
    
End Sub


Private Sub mskRstDate_GotFocus()

    With mskRstDate
        .SelStart = 0
        .SelLength = Len(.Text) + 2
    End With '
    
End Sub


Private Sub mskRstDate_KeyPress(KeyAscii As Integer)

    If Not KeyAscii = vbKeyBack Then mskRstDate.SelLength = 1
    
End Sub


Private Sub Order_Ready(ByVal ACK As String)

    Static msgIndex As Long
    
    Select Case ACK
        Case Chr(COM_ENQ)
            msgIndex = 1
        Case Chr(COM_ACK)
            msgIndex = msgIndex + 1
        Case Chr(COM_NACK)
            msgIndex = msgIndex
        Case Chr(COM_EOT)
            msgIndex = 7
            Set Order = Nothing
        Case Else
        
    End Select
    
    Select Case msgIndex
        Case 1
            Call COM_OUTPUT(Order.MSG_ENQ)
        Case 2
            Call COM_OUTPUT(Order.MSG_HEADER)
        Case 3
            Call COM_OUTPUT(Order.MSG_PATIENT)
        Case 4
            Call COM_OUTPUT(Order.MSG_ORDER)
        Case 5
            Call COM_OUTPUT(Order.MSG_TERMINATION)
        Case 6
            Call COM_OUTPUT(Order.MSG_EOT)
        Case Else
    End Select
    
End Sub

Private Sub Result_MsgBegin(ByVal SID As String)
    
'    Dim itemX As ListItem
'
'    Set itemX = lvwComplete.FindItem(Trim(SID), lvwTag, , lvwWhole)
'    If itemX Is Nothing Then
'        Set itemX = lvwComplete.ListItems.Add(, , Trim(SID))
'        If Not itemX Is Nothing Then
'            With itemX
'                .Key = COL_KEY & Trim(SID)
'                .tag = Trim(SID)
'                .SmallIcon = "LST"
'            End With
'        End If
'    End If
'
End Sub

Private Sub Result_MsgSplit(ByVal Result As clsResult)

'On Error GoTo ErrorRoutine
'
'    Dim sqlDoc  As String, sqlRet   As Integer
'
'    Dim strTime As String
'    Dim itemX   As ListItem
'    Dim itemH   As ListItem
'    Dim itemS   As ListSubItem
'
'    CallForm = "frmComm - Private Sub Result_MsgSplit()"
'
'    '��ġ ���̺��� �˻��ڵ带 ������
'    Set itemX = lvwCuData.FindItem(Trim(Result.Rst_Test), lvwTag, , lvwWhole)
'    If Not itemX Is Nothing Then
'        If Mid$(Result.Rst_Sid, 10, 2) = "PC" And Trim(Result.Rst_Test) = "06A" Then
'            Result.Rst_Sid = Mid$(Result.Rst_Sid, 1, 9)
'            Result.Rst_Test = "XXX"
'            Result.Rst_Tag = ""
'        Else
'            Result.Rst_Sid = Mid$(Result.Rst_Sid, 1, 9)
'            Result.Rst_Tag = Trim(itemX.SubItems(1))
'        End If
'
'        sqlDoc = "Update INTERFACE003 set RSTVAL = '" & Result.Rst_Values & "', REFVAL = '" & Result.Rst_Eid & "'" & _
'                 " where SPCNO  = '" & Result.Rst_Sid & "'" & _
'                 "   and TESTCD = '" & Result.Rst_Test & "'" & _
'                 "   and TRANSDT = '" & Format$(Now, "YYYYMMDD") & "'" & _
'                 "   and TRANSTM = '" & Format$(Now, "MMSS") & "'"
'        AdoCn_Jet.Execute sqlDoc, sqlRet
'        If sqlRet = 0 Then
'            sqlDoc = "insert into INTERFACE003(" & _
'                     "            SPCNO, TESTCD, EQPNUM, TRANSDT, TRANSTM, RSTVAL, REFVAL, EQUIPCD)" & _
'                     "    values( '" & Result.Rst_Sid & "', '" & Result.Rst_Test & "'," & _
'                     "            '" & Result.Rst_Eid & "', '" & Format$(Now, "YYYYMMDD") & "'," & _
'                     "            '" & Format$(Now, "MMSS") & "', '" & Result.Rst_Values & "'," & _
'                     "            '" & Result.Rst_Eid & "', '" & INS_CODE & "')"
'            AdoCn_Jet.Execute sqlDoc
'        End If
'
'        '��� ǥ��
'        Set itemH = lvwComplete.FindItem(Result.Rst_Sid, lvwText, , lvwWhole)
'        If itemH Is Nothing Then
'            Set itemH = lvwComplete.ListItems.Add()
'            With itemH
'                .Key = COL_KEY & Result.Rst_Sid '������ Ű�� ��ü��ȣ
'                .Text = Result.Rst_Sid          '������ �� ��ü��ȣ
'                .tag = Result.Rst_Type          '�ױ׿� ��� Ÿ��
'                .SmallIcon = "LSE"
'            End With
'        End If
'        '����� ���
'        itemH.SubItems(lvwComplete.ColumnHeaders(COL_KEY & Result.Rst_Test).SubItemIndex) = Result.Rst_Values
'
'        '--- ����
'        itemH.ListSubItems(lvwComplete.ColumnHeaders(COL_KEY & Result.Rst_Test).SubItemIndex).ForeColor = vbBlack
'        If Val(itemX.SubItems(7)) < Val(Result.Rst_Values) Or Val(itemX.SubItems(8)) > Val(Result.Rst_Values) Then
'            itemH.ListSubItems(lvwComplete.ColumnHeaders(COL_KEY & Result.Rst_Test).SubItemIndex).ForeColor = vbRed
'        End If
'
'        Set itemS = itemH.ListSubItems(lvwComplete.ColumnHeaders(COL_KEY & Result.Rst_Test).SubItemIndex)
'
'        itemS.tag = Result.Rst_Error '��������� �ױ׿� ���� �޽���
'        Set itemS = Nothing
'        Set itemX = Nothing
'        Set itemX = Nothing
'    End If
'    '�˻��ڵ尡 ���°��� ��� ���� ����
'    Exit Sub
'ErrorRoutine:
'
'    Set itemS = Nothing
'    Set itemX = Nothing
'    Set itemX = Nothing
'
'    Call ErrMsgProc(CallForm)
'    Err.Clear
    
End Sub



Private Sub spdResult1_Click(ByVal Col As Long, ByVal Row As Long)
    
    If Col > 6 Then
        spdResult1.Col = Col
        spdResult1.Row = Row
        If spdResult1.BackColor = &HC6FEFF Then
            spdResult1.BackColor = vbWhite
        Else
            spdResult1.BackColor = &HC6FEFF
        End If
    End If

End Sub

Private Sub spdResult1_DblClick(ByVal Col As Long, ByVal Row As Long)
    If Col > 4 Then
        spdResult1.Col = Col
        spdResult1.Row = Row
        spdResult1.CellType = CellTypeEdit
'        spdResult1.EditMode = True
    End If
End Sub

Private Sub spdWorkList_Click(ByVal Col As Long, ByVal Row As Long)

    If Col < 3 Then Exit Sub
    
    Dim varTmp  As Variant
    
    With spdWorkList
        If Col = 1 Then
            .GetText 2, Row, varTmp
            If Trim$(varTmp) = "" Then Exit Sub
            
            .SetText 1, Row, IIf(Trim$(varTmp) = "1", "", "1")
        ElseIf Col > 6 Then
            .GetText Col, 0, varTmp
            If Trim$(varTmp) = "" Then Exit Sub
            
            .Row = Row: .Col = Col
            If .BackColor = vbWhite Then
                .BackColor = &HC6FEFF
            Else
                .BackColor = vbWhite
            End If
        End If
    End With
    
End Sub


Private Sub tmrReceive_Timer()
    
    imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
    tmrReceive.Enabled = False

End Sub

Private Sub tmrSend_Timer()
    
    imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
    tmrSend.Enabled = False

End Sub

Private Sub Form_Resize()
    Dim i As Integer
    If ScaleHeight < 650 Then Exit Sub
    If ScaleWidth < 60 Then Exit Sub
    fraCmdBar.Move ScaleLeft + 30, ScaleHeight - fraCmdBar.Height - 30, ScaleWidth - 60
    For i = cmdAction.LBound To cmdAction.UBound
        Call cmdAction(i).Move(fraCmdBar.Width - ((1300 * (cmdAction.Count - i)) + (70 * (cmdAction.UBound - i)) + 100), _
                               (fraCmdBar.Height - 360) / 2, 1300, 360)
    Next
End Sub

Private Sub txtBarCode_Change()

    If txtBarCode.SelStart = txtBarCode.MaxLength Then SendKeys "{TAB}"
    
End Sub

Private Sub txtBarCode_GotFocus()

    With txtBarCode
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub


Private Sub txtBarCode_KeyPress(KeyAscii As Integer)
    
    Dim tst_no() As String, strPid()    As String, strPnm() As String
    Dim TMP() As String
    Dim rv As Long
    Dim samChk As Boolean
    Dim ii As Integer
    Dim bgetWork As Boolean
    Dim itemX As ListItem
    
    Dim strEqpCd   As String
    
On Error GoTo ErrRoutine
    CallForm = "frmInterface - Private Sub txtBarCode_KeyPress()"
        
    samChk = False
    If KeyAscii = vbKeyReturn Then
        If chkQC.Value = 1 Then
            rv = sl_spcid_tstcd_select_qc&(Trim(txtBarCode.Text), INS_CODE, tst_no)
        Else
            rv = sl_spcid_tstcd_select&(Trim(txtBarCode.Text), tst_no, strPid, strPnm)
        End If
'        rv = sl_spcid_tstcd_select(Trim(txtBarCode.Text), tst_no, strPid, strPnm)
        If (rv = 0) Then
            MsgBox "������ ��ü�Դϴ�.!", vbCritical
        Else
            If psDataExists Then
                MsgBox "�̹� ��ϵ� ��ü�Դϴ�.!", vbCritical
                txtBarCode.Text = ""
                Exit Sub
            End If
            
            bgetWork = False
            For ii = 0 To rv - 1
                strEqpCd = f_funget_code(tst_no(ii))
                Set itemX = lvwCuData.FindItem(strEqpCd, lvwTag, , lvwWhole)
                If Not itemX Is Nothing Then
                    bgetWork = True
                    Exit For
                End If
            Next
                    
             With spdWorkList
                If bgetWork = True Then
                    .Col = 2
                    For ii = 1 To .maxrows
                        .Row = ii
                        If Trim(.Text) = "" Then
                            .Text = txtBarCode.Text
                            .SetText 3, ii, strPnm(0)
                            .SetText 4, ii, strPid(0)
                            txtBarCode.Text = ""
                            .Col = 1
                            .Value = 1
                            samChk = True
                            Exit For
                        End If
                    Next
                    If samChk = False Then
                        .Col = 1
                        .Row = .maxrows + 1
                        .Value = 1
                        .Col = 2
                        .maxrows = .maxrows + 1
                        .Row = .maxrows
                        .Action = ActionActiveCell
                        .Text = txtBarCode.Text
                        .SetText 3, .maxrows, strPnm(0)
                        .SetText 4, .maxrows, strPid(0)
                        
                        .Col = 1
                        .Value = 1
                        
                        txtBarCode.Text = ""
                    End If
                Else
                   MsgBox "�ش�˻��׸��� �������� �ʴ� ��ü�Դϴ�.", vbOKOnly + vbInformation, Me.Caption
                End If
             End With
        End If
    End If

Exit Sub
ErrRoutine:
    Call ErrMsgProc(CallForm)
    
End Sub

Private Function f_funget_code(ByVal Testcd As String)

    Dim intIdx1 As Integer, intIdx2 As Integer
    
    f_funget_code = ""
    For intIdx1 = 1 To UBound(Channel)
        For intIdx2 = 1 To Channel(intIdx1).MachCnt
            If Channel(intIdx1).MachTst(intIdx2) = Testcd Then
                f_funget_code = Channel(intIdx1).MachCode
                Exit Function
            End If
        Next
    Next
    
End Function


Private Function psDataExists() As Boolean
Dim sCnt As Long
    
    psDataExists = False
    With spdWorkList
        For sCnt = 1 To .maxrows
            .Row = sCnt:    .Col = 2
            If Trim(.Text) = Mid(txtBarCode.Text, 1, 11) Then
                psDataExists = True
                Exit For
            End If
        Next sCnt
    End With

End Function

Private Sub txtBarCode_LostFocus()
'
'    Dim intRow      As Integer
'    Dim strOrdcd(1 To 100) As String
'
'    Call sl_spcid_tstcd_select&(txtBarCode.Text, strOrdcd)
'    If strOrdcd(1) = "" Then
'        MsgBox "�ش� �˻��׸��� �������� ���� ��ü�Դϴ�.", vbInformation, Me.Caption
'        Exit Sub
'    End If
'
'    intRow = f_funGet_SpreadRow(spdWorkList, 2, txtBarCode.Text)
'    If intRow < 1 Then
'        intRow = f_funGet_SpreadRow(spdWorkList, 2, "")
'        If intRow < 1 Then
'            spdWorkList.maxrows = spdWorkList.maxrows + 1
'            spdWorkList.RowHeight(spdWorkList.maxrows) = 13
'            intRow = spdWorkList.maxrows
'        End If
'        spdWorkList.SetText 2, intRow, txtBarCode.Text
'    End If
'    spdWorkList.SetText 1, intRow, "1"
'
End Sub


' ��Ż��� Ȯ�� �����̺�Ʈ
' ------------------------------------------------------------------------
Private Sub txtCom_Change()
    txtCom.SelStart = Len(txtCom.Text)
End Sub

Private Sub cmdCOMLoad_Click()
    Dim i               As Long
    Dim lngFIleNum      As Long
    Dim strTemp         As String
    Dim strTemp2        As String
    Dim bteBuffer()     As Byte
    
On Error GoTo ErrorRoutine
    
    With cdlFile
        .CancelError = True
        .FileName = App.Path & "\comm.txt"
        .ShowOpen
        lngFIleNum = FreeFile
        
        Open .FileName For Binary Access Read As #lngFIleNum
        
        txtCOM2.Text = ""
        ReDim bteBuffer(LOF(lngFIleNum))
        Get #lngFIleNum, , bteBuffer

        strTemp = StrConv(bteBuffer, vbUnicode)
        txtCOM2.Text = strTemp
                
        Close #lngFIleNum
    End With
    Exit Sub
    
ErrorRoutine:
    Close #lngFIleNum
        
End Sub

Private Sub cmdCOMSave_Click()
    Dim lngFIleNum      As Long
    
On Error GoTo ErrorRoutine

    With cdlFile
        .CancelError = True
        .FileName = App.Path & "\comm.txt"
        .ShowSave
        lngFIleNum = FreeFile
        
        Open .FileName For Append As #lngFIleNum
        Print #lngFIleNum, _
              Format(Date, "YYYY�� MM�� DD��") & "  "; Time & vbNewLine & _
              "������������������������������������������������������������" & vbNewLine & _
              txtCom.Text & _
              "������������������������������������������������������������" & vbNewLine
    Close #lngFIleNum
    End With
    Exit Sub
    
ErrorRoutine:
    Close #lngFIleNum

End Sub

Private Sub cmdCOMOutput_Click()
    'Call COM_OUTPUT(StrConv(charCOM_Convert(txtCom.SelText), vbFromUnicode))
    Call COM_OUTPUT(charCOM_Convert(txtCom.SelText))
End Sub

Private Sub cmdCOMClear_Click()
    mlngRecLen = 0
    txtCom.Text = ""
End Sub

Private Sub cmdCOMClear2_Click()
    txtCOM2.Text = ""
End Sub

Private Sub cmdCOMInput_Click()
    Dim bytTemp() As Byte
    
    bytTemp = StrConv(charCOM_Convert(txtCom.SelText), vbFromUnicode)

    Call ComReceive(bytTemp)
End Sub

Private Sub cmdCOMInput2_Click()
    
    Dim bytTemp() As Byte
    
    If txtCOM2.SelLength = 0 Then
        bytTemp = StrConv(charCOM_Convert(txtCOM2.Text), vbFromUnicode)
    Else
        bytTemp = StrConv(charCOM_Convert(txtCOM2.SelText), vbFromUnicode)
    End If

    Call ComReceive(bytTemp)

End Sub

Private Sub cmdCOMOutput2_Click()
    
    If txtCOM2.SelLength = 0 Then
        Call COM_OUTPUT(charCOM_Convert(txtCOM2.Text))
    Else
        Call COM_OUTPUT(charCOM_Convert(txtCOM2.SelText))
    End If
    
End Sub
' ------------------------------------------------------------------------
' ��Ż��� Ȯ�� �����̺�Ʈ


