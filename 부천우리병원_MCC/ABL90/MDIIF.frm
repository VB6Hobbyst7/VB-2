VERSION 5.00
Object = "{4BD5DFC7-B668-44E0-A002-C1347061239D}#1.0#0"; "HSCotrol.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIIF 
   BackColor       =   &H00FFFFFF&
   Caption         =   "SANSOFT Interface"
   ClientHeight    =   9315
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   20115
   Icon            =   "MDIIF.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.PictureBox picHeader 
      Align           =   1  '�� ����
      BackColor       =   &H00F8E4D8&
      BorderStyle     =   0  '����
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   20115
      TabIndex        =   2
      Top             =   0
      Width           =   20115
      Begin HSCotrol.CButton cmdPrint 
         Height          =   495
         Left            =   20730
         TabIndex        =   29
         ToolTipText     =   "������ ��ü����� OCS/EMR�� �����մϴ�."
         Top             =   30
         Visible         =   0   'False
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   873
         BackColor       =   16777215
         Caption         =   "ȭ�����"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MDIIF.frx":25CA
         MaskColor       =   0
         PicCapAlign     =   2
         BorderStyle     =   1
         BorderColor     =   4210752
         HoverColor      =   16711680
         HoverPicture    =   "MDIIF.frx":341C
      End
      Begin VB.Frame fraJWINFO 
         BackColor       =   &H00F8E4D8&
         Height          =   495
         Left            =   22410
         TabIndex        =   25
         Top             =   30
         Visible         =   0   'False
         Width           =   2595
         Begin VB.OptionButton optSch_JW 
            Appearance      =   0  '���
            BackColor       =   &H00F8E4D8&
            Caption         =   "�ܷ�"
            ForeColor       =   &H00808080&
            Height          =   225
            Index           =   2
            Left            =   1740
            TabIndex        =   28
            Top             =   180
            Width           =   735
         End
         Begin VB.OptionButton optSch_JW 
            Appearance      =   0  '���
            BackColor       =   &H00F8E4D8&
            Caption         =   "�Կ�"
            ForeColor       =   &H00808080&
            Height          =   225
            Index           =   1
            Left            =   930
            TabIndex        =   27
            Top             =   180
            Width           =   735
         End
         Begin VB.OptionButton optSch_JW 
            Appearance      =   0  '���
            BackColor       =   &H00F8E4D8&
            Caption         =   "��ü"
            ForeColor       =   &H00808080&
            Height          =   225
            Index           =   0
            Left            =   150
            TabIndex        =   26
            Top             =   180
            Value           =   -1  'True
            Width           =   735
         End
      End
      Begin VB.Timer Timer1 
         Left            =   1530
         Top             =   -180
      End
      Begin VB.Frame fraStatus 
         Appearance      =   0  '���
         BackColor       =   &H00F8E4D8&
         BorderStyle     =   0  '����
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   14190
         TabIndex        =   18
         Top             =   30
         Visible         =   0   'False
         Width           =   8655
         Begin HSCotrol.CButton cmdXML 
            Height          =   405
            Left            =   4380
            TabIndex        =   30
            Top             =   90
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   714
            BackColor       =   15698777
            Caption         =   "XML ����"
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
            MaskColor       =   0
            BorderStyle     =   1
            BorderColor     =   16777215
            HoverColor      =   65535
         End
         Begin VB.Label lblIFStatus 
            Appearance      =   0  '���
            BackColor       =   &H00F8E4D8&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   225
            Left            =   4380
            TabIndex        =   23
            Top             =   90
            Width           =   3075
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblComStatus 
            Appearance      =   0  '���
            BackColor       =   &H00F8E4D8&
            Caption         =   "Com1 ���Ἲ��"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   225
            Left            =   2070
            TabIndex        =   22
            Top             =   30
            Width           =   2115
         End
         Begin VB.Image imgSend 
            Height          =   240
            Left            =   2595
            Picture         =   "MDIIF.frx":3CF6
            Top             =   300
            Width           =   240
         End
         Begin VB.Image imgReceive 
            Height          =   240
            Left            =   3915
            Picture         =   "MDIIF.frx":4280
            Top             =   300
            Width           =   240
         End
         Begin VB.Label lblSend 
            AutoSize        =   -1  'True
            BackStyle       =   0  '����
            Caption         =   "�޴½�ȣ"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   225
            Left            =   1785
            TabIndex        =   21
            Top             =   330
            Width           =   720
         End
         Begin VB.Label lblRcv 
            AutoSize        =   -1  'True
            BackStyle       =   0  '����
            Caption         =   "�����½�ȣ"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   225
            Left            =   2940
            TabIndex        =   20
            Top             =   300
            Width           =   900
         End
         Begin VB.Image imgPort 
            Height          =   240
            Left            =   1755
            Picture         =   "MDIIF.frx":480A
            Top             =   30
            Width           =   240
         End
         Begin VB.Image imgNet1 
            Height          =   240
            Left            =   30
            Picture         =   "MDIIF.frx":4D94
            Top             =   210
            Width           =   240
         End
         Begin VB.Image imgNet2 
            Height          =   240
            Left            =   30
            Picture         =   "MDIIF.frx":4EDE
            Top             =   210
            Width           =   240
         End
         Begin VB.Image imgNet3 
            Height          =   240
            Left            =   30
            Picture         =   "MDIIF.frx":5028
            Top             =   210
            Width           =   240
         End
         Begin VB.Label lblDBStatus 
            BackStyle       =   0  '����
            Caption         =   "�����ͺ��̽� ���Ἲ��"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   465
            Left            =   390
            TabIndex        =   19
            Top             =   90
            Width           =   1185
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  '���
         BackColor       =   &H00F8E4D8&
         BorderStyle     =   0  '����
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   9165
         TabIndex        =   4
         Top             =   -60
         Width           =   3645
         Begin HSCotrol.CButton cmdTestNmSave 
            Height          =   495
            Left            =   2100
            TabIndex        =   14
            ToolTipText     =   "�˻���ID/���� �����մϴ�."
            Top             =   120
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   873
            BackColor       =   16777215
            Caption         =   "����ں���"
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����ü"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Picture         =   "MDIIF.frx":5172
            MaskColor       =   0
            PicCapAlign     =   2
            BorderStyle     =   1
            BorderColor     =   32768
            HoverColor      =   16711680
         End
         Begin VB.TextBox txtTestID 
            Alignment       =   2  '��� ����
            Appearance      =   0  '���
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   870
            TabIndex        =   11
            Top             =   120
            Visible         =   0   'False
            Width           =   1185
         End
         Begin VB.TextBox txtTestNm 
            Alignment       =   2  '��� ����
            Appearance      =   0  '���
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   870
            TabIndex        =   10
            Top             =   360
            Visible         =   0   'False
            Width           =   1185
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '����
            Caption         =   "�˻��ڸ� :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Left            =   0
            TabIndex        =   15
            Top             =   360
            Width           =   825
         End
         Begin VB.Label lblTestID 
            BackStyle       =   0  '����
            Caption         =   "�˻���ID"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   960
            TabIndex        =   13
            ToolTipText     =   "�˻���ID�� �����Ϸ��� ����Ŭ�� �ϼ���"
            Top             =   150
            Width           =   975
         End
         Begin VB.Label lblTestNm 
            BackStyle       =   0  '����
            Caption         =   "�˻��ڸ�"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Left            =   960
            TabIndex        =   12
            ToolTipText     =   "�˻��ڸ��� �����Ϸ��� ����Ŭ�� �ϼ���"
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '����
            Caption         =   "�˻���ID :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   0
            TabIndex        =   9
            Top             =   150
            Width           =   825
         End
      End
      Begin VB.CheckBox chkLock 
         BackColor       =   &H00F8E4D8&
         Caption         =   "�޴�����"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   435
         Left            =   60
         TabIndex        =   3
         Top             =   120
         Width           =   705
      End
      Begin HSCotrol.CButton cmdClear 
         Height          =   495
         Left            =   4620
         TabIndex        =   6
         ToolTipText     =   "�������̽� ȭ���� ����ϴ�."
         Top             =   60
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   873
         BackColor       =   16777215
         Caption         =   "ȭ������"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MDIIF.frx":5C97
         MaskColor       =   0
         PicCapAlign     =   2
         BorderStyle     =   1
         BorderColor     =   4210752
         HoverColor      =   16711680
         HoverPicture    =   "MDIIF.frx":6971
      End
      Begin HSCotrol.CButton cmdSave 
         Height          =   495
         Left            =   7530
         TabIndex        =   7
         ToolTipText     =   "������ ��ü����� OCS/EMR�� �����մϴ�."
         Top             =   60
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   873
         BackColor       =   16777215
         Caption         =   "��������"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MDIIF.frx":79F8
         MaskColor       =   0
         PicCapAlign     =   2
         BorderStyle     =   1
         BorderColor     =   4210752
         HoverColor      =   16711680
         HoverPicture    =   "MDIIF.frx":82D2
      End
      Begin HSCotrol.CButton cmdView 
         Height          =   495
         Left            =   6060
         TabIndex        =   8
         ToolTipText     =   "������ ��ü�� �󼼰���� ���Դϴ�."
         Top             =   60
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   873
         BackColor       =   16777215
         Caption         =   "�󼼰��"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MDIIF.frx":8BAC
         MaskColor       =   0
         PicCapAlign     =   2
         BorderStyle     =   1
         BorderColor     =   4210752
         HoverColor      =   16711680
         HoverPicture    =   "MDIIF.frx":9833
      End
      Begin HSCotrol.CButton cmdOrderSend 
         Height          =   495
         Left            =   12840
         TabIndex        =   31
         Top             =   60
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   873
         BackColor       =   32768
         Caption         =   "��������"
         ForeColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MDIIF.frx":A885
         MaskColor       =   0
         PicCapAlign     =   2
         BorderStyle     =   1
         BorderColor     =   16777215
         HoverColor      =   0
      End
      Begin VB.Label lblMenuInfo 
         BackStyle       =   0  '����
         Caption         =   "UROMETER120"
         BeginProperty Font 
            Name            =   "Segoe UI Historic"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   1440
         TabIndex        =   24
         Top             =   180
         Width           =   1845
      End
      Begin VB.Label Label9 
         BackStyle       =   0  '����
         Caption         =   "�˻����� : "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3480
         TabIndex        =   17
         Top             =   90
         Width           =   975
      End
      Begin VB.Label lblTestDate 
         BackStyle       =   0  '����
         Caption         =   "1971-03-11"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   3480
         TabIndex        =   16
         Top             =   330
         UseMnemonic     =   0   'False
         Width           =   975
      End
      Begin VB.Image Image2 
         Height          =   420
         Left            =   810
         Picture         =   "MDIIF.frx":AAE1
         Top             =   90
         Width           =   2580
      End
   End
   Begin VB.PictureBox picNode 
      Align           =   3  '���� ����
      BackColor       =   &H00FFFFFF&
      Height          =   8700
      Left            =   0
      ScaleHeight     =   8640
      ScaleWidth      =   2940
      TabIndex        =   0
      Top             =   615
      Width           =   3000
      Begin HSCotrol.CButton cmdNode 
         Height          =   9855
         Left            =   2625
         TabIndex        =   5
         Top             =   120
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   17383
         BackColor       =   16777215
         Caption         =   "��"
         ForeColor       =   12553049
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         BorderStyle     =   1
         BorderColor     =   16777215
         HoverColor      =   -2147483630
      End
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   14445
         Left            =   60
         TabIndex        =   1
         Top             =   0
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   25479
         _Version        =   393217
         LineStyle       =   1
         Style           =   7
         ImageList       =   "imlSubList(1)"
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imlSubList 
      Index           =   11
      Left            =   4680
      Top             =   690
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIIF.frx":C2F0
            Key             =   "LIS1101"
            Object.Tag             =   "Menu"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIIF.frx":D342
            Key             =   "LIS1102"
            Object.Tag             =   "SubMenus"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIIF.frx":E394
            Key             =   "LIS1104"
            Object.Tag             =   "SubMenus"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIIF.frx":F3E6
            Key             =   "LIS1103"
            Object.Tag             =   "SubMenu"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   " ���� "
      Begin VB.Menu mnuExit 
         Caption         =   "����"
      End
   End
   Begin VB.Menu mnuMenu00 
      Caption         =   "  �������̽� "
      Visible         =   0   'False
      Begin VB.Menu mnuHoriba 
         Caption         =   " HORIBA "
      End
   End
   Begin VB.Menu mnuMenu01 
      Caption         =   "  ��ȸ���� "
      Visible         =   0   'False
      Begin VB.Menu mnuResult 
         Caption         =   " ��� ��ȸ"
      End
      Begin VB.Menu mnuSep29 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWork 
         Caption         =   " ��ũ ��ȸ"
      End
      Begin VB.Menu mnuSep9 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuMenu02 
      Caption         =   " �������� "
      Visible         =   0   'False
      Begin VB.Menu mnuComm 
         Caption         =   " ��� ����"
      End
      Begin VB.Menu mnuSep11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTest 
         Caption         =   " �˻� ����"
      End
      Begin VB.Menu mnuSep12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuView 
         Caption         =   " ȭ�� ����"
      End
      Begin VB.Menu mnuSep22 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpt 
         Caption         =   " �ɼ� ����"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep23 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHosp 
         Caption         =   " ������� ����"
      End
      Begin VB.Menu mnuSep25 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEMRInfo 
         Caption         =   " �������� ����"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMenu06 
      Caption         =   " ��� "
      Begin VB.Menu mnuWorkSave 
         Caption         =   " ��ũ����Ʈ ���� "
      End
      Begin VB.Menu mnuSep27 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWorkOpen 
         Caption         =   " ��ũ����Ʈ ����"
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
      WindowList      =   -1  'True
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
      Begin VB.Menu mnComStatus 
         Caption         =   "��Ż��º���"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep26 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCommTest 
         Caption         =   "����׽�Ʈ"
      End
      Begin VB.Menu mnuSep28 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "SANIF ����"
      End
   End
End
Attribute VB_Name = "MDIIF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    

Private Sub chkLock_Click()
    Dim strMenuLock As String
    
    If chkLock.Value = "1" Then
        strMenuLock = "1"
    Else
        strMenuLock = "0"
    End If
    
    'Call WritePrivateProfileString("HOSP", "MENULOCK", strMenuLock, App.PATH & "\INI\" & gMACH & ".ini")
    Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "MENULOCK", strMenuLock)

End Sub

Private Sub cmdClear_Click()

    If frmInterface.WindowState = 2 Then
        Call frmInterface.frmClear
    End If
    
End Sub

Private Sub cmdNode_Click()
    
'    Call FrmMove

        With MDIIF
            If .cmdNode.Caption = "��" Then
                .cmdNode.Caption = "��"
                .TreeView1.Visible = True
                .picNode.WIDTH = 3000 '3930
                .cmdNode.LEFT = (.picNode.WIDTH - .cmdNode.WIDTH) - 30
            Else
                .cmdNode.Caption = "��"
                .TreeView1.Visible = False
                .picNode.WIDTH = 400 '300
                .cmdNode.LEFT = (.picNode.WIDTH - .cmdNode.WIDTH) - 30
            End If
        End With

End Sub

Private Sub cmdNode_MouseIn()
    
    Call FrmMove

End Sub

Private Sub cmdOrderSend_Click()

    If frmInterface.spdOrder.MaxRows < 1 Then
        MsgBox "�˻����ڰ� �����ϴ�.", vbOKOnly + vbCritical, Me.Caption
    Else
        intPhase = 3
        intSndPhase = 1
        strState = "Q"
        Call frmInterface.SendData(ENQ)
    End If
    
End Sub

Private Sub cmdPrint_Click()
    
    With frmInterface
        .spdOrder.PrintOrientation = PrintOrientationLandscape       '�������
        .spdOrder.Action = 13
    End With
    
End Sub

Private Sub cmdSave_Click()
    Dim intRow      As Integer
    Dim intRes      As Integer
    Dim strRCnt     As String
    Dim intRCnt     As Integer
    
    If frmInterface.WindowState <> 2 Then
        Exit Sub
    End If
    
    If frmInterface.spdOrder.MaxRows = 0 Then
        Exit Sub
    End If
    
    If MsgBox("������ ����� �����Ͻðڽ��ϱ�?", vbYesNo + vbCritical, "�������") = vbYes Then
        With frmInterface.spdOrder
            For intRow = 1 To .DataRowCnt
                strRCnt = GetText(frmInterface.spdOrder, intRow, colRCNT)
                If Not IsNumeric(strRCnt) Then
                    intRCnt = 0
                Else
                    intRCnt = strRCnt
                End If
                .Row = intRow
                .Col = colCHECKBOX
                If .Value = 1 And GetText(frmInterface.spdOrder, intRow, colSTATE) <> "" And intRCnt > 0 Then
                    intRes = SaveTransData(intRow, frmInterface.spdOrder)
                    Call SetUpdateStatus(frmInterface.spdOrder, intRow, intRes)
                End If
            Next
        End With
    End If
End Sub

Private Sub cmdView_Click()

    With frmInterface
        If .WindowState = 2 Then
            If gWORKPOS = "M" Then
                If .spdResult.Visible = False Then
                    .spdResult.Visible = True
                    Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "VIEW", "DETAILVIEW", "Y")
                    
                    .spdOrder.HEIGHT = Me.ScaleHeight - .picHeader.HEIGHT - 100
                    .spdOrder.WIDTH = Me.ScaleWidth - .spdWork.WIDTH - .spdResult.WIDTH - 200
                    
                    .spdResult.LEFT = .spdOrder.LEFT + .spdOrder.WIDTH + 50
                    .spdResult.HEIGHT = .spdOrder.HEIGHT
                    .spdResult.TOP = .spdOrder.TOP
                Else
                    .spdResult.Visible = False
                    .spdOrder.WIDTH = Me.ScaleWidth - .spdWork.WIDTH - 200
                    
                    Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "VIEW", "DETAILVIEW", "N")
                End If
            Else
                If .spdResult.Visible = False Then
                    .spdResult.Visible = True
                    Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "VIEW", "DETAILVIEW", "Y")
                    
                    .spdOrder.HEIGHT = Me.ScaleHeight - .picHeader.HEIGHT - 100
                    .spdOrder.WIDTH = Me.ScaleWidth - .spdResult.WIDTH - 200
                    
                    .spdResult.LEFT = .spdOrder.LEFT + .spdOrder.WIDTH + 50
                    .spdResult.HEIGHT = .spdOrder.HEIGHT
                    .spdResult.TOP = .spdOrder.TOP
                Else
                    .spdResult.Visible = False
                    .spdOrder.WIDTH = Me.ScaleWidth - 200
                    
                    Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "VIEW", "DETAILVIEW", "N")
                End If
            End If
        End If
    End With
End Sub


Private Sub lblMenuInfo_Click()

    frmInterface.ZOrder 0
    
End Sub

Private Sub lblMenuInfo_DblClick()

    If fraStatus.Visible = False Then
        fraStatus.Visible = True
    Else
        fraStatus.Visible = False
    End If
    
End Sub

Private Sub MDIForm_Load()
    
    'MDI�� ũ��
    If Mid(gForm.MAXYN, 1, 1) = "Y" Then
        Me.WindowState = 2
    Else
        Me.TOP = IIf(gForm.TOP < 0, 0, gForm.TOP)
        Me.LEFT = IIf(gForm.LEFT < 0, 0, gForm.LEFT)
        Me.WIDTH = IIf(gForm.WIDTH < 0, 10000, gForm.WIDTH)
        Me.HEIGHT = IIf(gForm.HEIGHT < 0, 10000, gForm.HEIGHT)
    End If
    
    cmdNode.HEIGHT = TreeView1.HEIGHT
    Me.Caption = "SANSOFT �������̽�"
    lblMenuInfo.Caption = gHOSP.MACHNM '"�������̽�"
    MDIIF.lblTestDate.Caption = Format(Now, "yyyy-mm-dd")
    lblTestID.Caption = gHOSP.USERID
    lblTestNm.Caption = gHOSP.USERNM
    
    Call SetTreeNode
    Call FrmMove
    Call frmShow(frmInterface)
    chkLock.Value = gHOSP.MENULOCK

    If gEMR = "JWINFO" Then
        fraJWINFO.Visible = True
    Else
        fraJWINFO.Visible = False
    End If
    
    If InStr(gHOSP.MACHNM, "BATCH") > 0 Then
        cmdOrderSend.Visible = True
    Else
        cmdOrderSend.Visible = False
    End If
    
    fraStatus.Visible = True


End Sub

'-----------------------------------------------------------------------------'
'   ��� : �̰� ����...
'-----------------------------------------------------------------------------'
Public Sub FrmMove()
    
    If chkLock.Value = "0" Then
        With MDIIF
            If .cmdNode.Caption = "��" Then
                .cmdNode.Caption = "��"
                .TreeView1.Visible = True
                .picNode.WIDTH = 3000 '3930
                .cmdNode.LEFT = (.picNode.WIDTH - .cmdNode.WIDTH) - 30
            Else
                .cmdNode.Caption = "��"
                .TreeView1.Visible = False
                .picNode.WIDTH = 400 '300
                .cmdNode.LEFT = (.picNode.WIDTH - .cmdNode.WIDTH) - 30
            End If
        End With
    End If
End Sub

Private Sub SetTreeNode()
    Dim nodX As Node

    picNode.Visible = True
    
    With TreeView1
        .Refresh
        .Visible = False
        .LabelEdit = lvwManual
        
        .ImageList = imlSubList(11)
        .HideSelection = False
        .Nodes.Clear
        
        Set nodX = .Nodes.Add(, tvwTreeLines, "LIS000", "�������̽�", "LIS1101")
        .Nodes("LIS000").Expanded = True
        Set nodX = .Nodes.Add(, tvwTreeLines, "LIS001", "��ȸ����", "LIS1101")
        .Nodes("LIS001").Expanded = True
        Set nodX = .Nodes.Add(, tvwTreeLines, "LIS002", "��������", "LIS1101")
        .Nodes("LIS002").Expanded = True
        .LineStyle = tvwTreeLines
        .Indentation = 300
        
        Set nodX = Nothing
        .Visible = True
        
    End With

End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Cancel = 1
    Call mnuExit_Click

End Sub

Private Sub MDIForm_Resize()
    
    If Me.WindowState = 2 Then
        Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "FORM", "MAXYN", "Y")
    Else
        gForm.TOP = Me.TOP
        gForm.LEFT = Me.LEFT
        gForm.WIDTH = Me.WIDTH
        gForm.HEIGHT = Me.HEIGHT
        
        Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "FORM", "MAXYN", "N")
        Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "FORM", "TOP", gForm.TOP)
        Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "FORM", "LEFT", gForm.LEFT)
        Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "FORM", "WIDTH", gForm.WIDTH)
        Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "FORM", "HEIGHT", gForm.HEIGHT)
    End If
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)

    If MsgBox("�����Ͻðڽ��ϱ�?", vbYesNo + vbCritical, "���α׷� ����") = vbYes Then
        If frmInterface.comEqp.PortOpen = True Then
            frmInterface.comEqp.PortOpen = False
        End If
        
        Close #1
        
        If gDBTYPE <> "99" Then
            Call DisConnect_Server

            Call DisConnect_Local
        End If

        End
    End If
    
End Sub

Private Sub mnComStatus_Click()
    
    If fraStatus.Visible = True Then
        fraStatus.Visible = False
    Else
        fraStatus.Visible = True
    End If
    
End Sub

Private Sub mnuAbout_Click()

    Call ShowForm(frmAbout, "�����Ʈ SANIF ����")

End Sub

Private Sub mnuHoriba_Click()
    
    Call ShowForm(frmInterface, "�������̽�")

End Sub

Private Sub mnuWorkOpen_Click()
    Dim strPath  As String
    Dim TextLine
    Dim strBuffer
    Dim strCount    As String
    
    If frmInterface.WindowState <> 2 Then
        Exit Sub
    End If
    
    If frmInterface.spdOrder.MaxRows > 0 Then
        If MsgBox("���� ȭ���� ����� ��ũ����Ʈ�� �ҷ����ڽ��ϱ�?", vbYesNo + vbInformation, "��ũ����Ʈ �ҷ�����") = vbNo Then
            Exit Sub
        End If
    End If
    
    With frmInterface.CommonDialog1
        .CancelError = True
        On Error GoTo ErrHandler
        .Flags = cdlOFNHideReadOnly
        .InitDir = App.PATH & "\WorkList"
        .Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*|"
        .FilterIndex = 1
        .Filename = ""
        .ShowOpen
        strPath = .Filename
    End With
    
    Open strPath For Input As #1
    Do While Not EOF(1)
        Line Input #1, TextLine
        strBuffer = strBuffer & TextLine & vbCr & vbLf
    Loop
    Close #1
 
    strCount = mGetP(mGetP(mGetP(strPath, 2, "WL_"), 3, "_"), 1, ".")
    
    frmInterface.spdOrder.MaxRows = strCount
    
    With frmInterface.spdOrder
        .Row = 1:       .Row2 = .MaxRows
        .Col = 1:       .Col2 = .MaxCols
        .BlockMode = True
        .Action = ActionClearText
        .Clip = strBuffer
        .ClipboardPaste
        .BlockMode = False
        
        .RowHeight(-1) = gROWHEIGHT
    End With
    
Exit Sub
ErrHandler:
                        
End Sub

Private Sub mnuWorkSave_Click()
    Dim strBuffer As String
    
    If frmInterface.WindowState <> 2 Then
        Exit Sub
    End If
    
    With frmInterface
        If .spdOrder.MaxRows < 1 Then
            MsgBox "������ ��ũ����Ʈ�� �����ϴ�.", vbOKOnly + vbCritical, "��ũ ����Ʈ"
            Exit Sub
        End If
        
        Call .spdOrder.SetSelection(1, 1, .spdOrder.MaxCols, .spdOrder.MaxRows)
        'Ŭ������ ī��
        .spdOrder.ClipboardCopy
        
        strBuffer = Clipboard.GetText()
        
        Call SetWorkData(strBuffer, .spdOrder.MaxRows)
        
        MsgBox "��ũ����Ʈ�� ���� �Ǿ����ϴ�.", vbOKOnly + vbInformation, "��ũ ����Ʈ"
        
    End With
    
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)

    Call TreeFromLoad(Node)
    
End Sub

Private Sub TreeFromLoad(ByVal Button As MSComctlLib.Node, Optional ByVal intIdx As Integer)
    
    If Button.Children <> 0 Then
        Exit Sub
    End If
    
    With TreeView1
        Select Case Button.Key
            '�������̽� ===========================================================================================================
            Case "LIS000":
                            TreeView1.Nodes.Add "LIS000", tvwChild, "LIS00001", gHOSP.MACHNM, "LIS1103"
                            'TreeView1.Nodes.Add "LIS000", tvwChild, "LIS00002", "XP300", "LIS1103"
                            
                            Case "LIS00001":        Call ShowForm(frmInterface, frmInterface.Caption)
                            'Case "LIS00002":        Call ShowForm(frmInterface2, frmInterface2.Caption)
                            
            '��ȸ���� ===========================================================================================================
            Case "LIS001":
                            TreeView1.Nodes.Add "LIS001", tvwChild, "LIS00101", "��� ��ȸ", "LIS1103"
                            TreeView1.Nodes.Add "LIS001", tvwChild, "LIS00102", "��ũ ��ȸ", "LIS1103"
                            TreeView1.Nodes.Add "LIS001", tvwChild, "LIS00103", "�˻� ���", "LIS1103"

                            Case "LIS00101":        Call ShowForm(frmResult, frmResult.Caption)
                            Case "LIS00102":        Call ShowForm(frmWorkList, frmWorkList.Caption)
                            Case "LIS00103":        Call ShowForm(frmStatistics, frmStatistics.Caption)
            '�������� =======================================================================================================
            Case "LIS002":
                            TreeView1.Nodes.Add "LIS002", tvwChild, "LIS00201", "�˻缳��", "LIS1103"
                            TreeView1.Nodes.Add "LIS002", tvwChild, "LIS00202", "��ż���", "LIS1103"
                            TreeView1.Nodes.Add "LIS002", tvwChild, "LIS00203", "ȭ�鼳��", "LIS1103"
                            TreeView1.Nodes.Add "LIS002", tvwChild, "LIS00204", "�����������", "LIS1103"
                            'TreeView1.Nodes.Add "LIS002", tvwChild, "LIS00205", "�ɼǼ���", "LIS1103"

                            Case "LIS00201":        Call ShowForm(frmTestSet, frmTestSet.Caption)
                            Case "LIS00202":        Call ShowForm(frmConfig, frmConfig.Caption)
                            Case "LIS00203":        Call ShowForm(frmScreenSet, frmScreenSet.Caption)
                            Case "LIS00204":        Call ShowForm(frmHospInfo, frmHospInfo.Caption)
                            'Case "LIS00205":        Call ShowForm(frmTestOptSet, frmTestOptSet.Caption)
            
            
            
        End Select
    End With
    
End Sub

Private Sub cmdTestNmSave_Click()
    
    If txtTestID.Text <> "" Then
        lblTestID.Caption = txtTestID.Text
        'Call WritePrivateProfileString("HOSP", "USERID", lblTestID.Caption, App.PATH & "\INI\" & gMACH & ".ini")
        Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "USERID", txtTestID.Text)
        
        txtTestID.Visible = False
        lblTestID.Visible = True
    End If
    
    If txtTestNm.Text <> "" Then
        lblTestNm.Caption = txtTestNm.Text
        'Call WritePrivateProfileString("HOSP", "USERNM", lblTestNm.Caption, App.PATH & "\INI\" & gMACH & ".ini")
        Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "USERNM", txtTestNm.Text)
        txtTestNm.Visible = False
        lblTestNm.Visible = True
    End If
    
End Sub


Private Sub lblTestID_DblClick()
    If txtTestID.Visible = False Then
        txtTestID.Text = lblTestID.Caption
        lblTestID.Visible = False
        txtTestID.Visible = True
    Else
        txtTestID.Text = ""
        lblTestID.Visible = True
        txtTestID.Visible = False
    End If
End Sub


Private Sub lblTestNm_DblClick()
    If txtTestNm.Visible = False Then
        txtTestNm.Text = lblTestNm.Caption
        lblTestNm.Visible = False
        txtTestNm.Visible = True
    Else
        txtTestNm.Text = ""
        lblTestNm.Visible = True
        txtTestNm.Visible = False
    End If
End Sub


Private Sub mnuBarcode_Click()
    
    mnuBarcode.Checked = True
    mnuSeqno.Checked = False
    mnuRackPos.Checked = False
    mnuCheckBox.Checked = False

    'Call WritePrivateProfileString("HOSP", "BARUSE", "Y", App.PATH & "\INI\" & gMACH & ".ini")
    'Call WritePrivateProfileString("HOSP", "RSTTYPE", "0", App.PATH & "\INI\" & gMACH & ".ini")

    Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "BARUSE", "Y")
    Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "RSTTYPE", "0")


End Sub

Private Sub mnuCheckBox_Click()
    
    mnuBarcode.Checked = False
    mnuSeqno.Checked = False
    mnuRackPos.Checked = False
    mnuCheckBox.Checked = True

    'Call WritePrivateProfileString("HOSP", "BARUSE", "N", App.PATH & "\INI\" & gMACH & ".ini")
    'Call WritePrivateProfileString("HOSP", "RSTTYPE", "3", App.PATH & "\INI\" & gMACH & ".ini")

    Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "BARUSE", "N")
    Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "RSTTYPE", "3")

End Sub

Private Sub mnuComm_Click()
    
    frmConfig.Show

End Sub

Private Sub mnuComTest_Click()

End Sub

Private Sub mnuCommTest_Click()

    If frmInterface.picComm.Visible = True Then
        frmInterface.picComm.Visible = False
    Else
        frmInterface.picComm.Visible = True
        frmInterface.picComm.ZOrder 0
    End If
    
End Sub

Private Sub mnuEMRInfo_Click()
    
    If InputBox("��й�ȣ �Է�" & Space(5) & "hint:������oyh") = "dev0503" Then
        frmEMRInfo.Show
    End If
    
End Sub

Private Sub mnuEqpResult_Click()

    mnuEqpResult.Checked = True
    mnuLisResult.Checked = False

    'Call WritePrivateProfileString("HOSP", "SAVELIS", "N", App.PATH & "\INI\" & gMACH & ".ini")
    Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "SAVELIS", "N")

End Sub

Private Sub mnuExit_Click()
    
    If MsgBox("�����Ͻðڽ��ϱ�?", vbYesNo + vbCritical, "���α׷� ����") = vbYes Then

        If KillProcess("EPOC") = False Then
            Call Shell("taskkill.exe /im EPOC.exe", 0)
        End If
        
        If frmInterface.comEqp.PortOpen = True Then
            frmInterface.comEqp.PortOpen = False
        End If
        
        Close #1
        
        If gDBTYPE <> "99" Then
            Call DisConnect_Server

            Call DisConnect_Local
        End If

        End
    End If
    
End Sub

Private Sub mnuHelp01_Click()

    Call WinExec(App.PATH & "\TeamViewerQS.exe", 1)
    
End Sub

Private Sub mnuHosp_Click()

    frmHospInfo.Show 'vbModal

End Sub

Private Sub mnuLisResult_Click()

    mnuEqpResult.Checked = False
    mnuLisResult.Checked = True

    'Call WritePrivateProfileString("HOSP", "SAVELIS", "Y", App.PATH & "\INI\" & gMACH & ".ini")
    Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "SAVELIS", "Y")

End Sub

Private Sub mnuOpt_Click()
    
    frmTestOptSet.Show 'vbModal
    
End Sub

Private Sub mnuRackPos_Click()
    
    mnuBarcode.Checked = False
    mnuSeqno.Checked = False
    mnuRackPos.Checked = True
    mnuCheckBox.Checked = False

    'Call WritePrivateProfileString("HOSP", "BARUSE", "N", App.PATH & "\INI\" & gMACH & ".ini")
    'Call WritePrivateProfileString("HOSP", "RSTTYPE", "2", App.PATH & "\INI\" & gMACH & ".ini")

    Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "BARUSE", "N")
    Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "RSTTYPE", "2")

End Sub

Private Sub mnuResult_Click()
    
    frmResult.Show 'vbModal
    
End Sub

Private Sub mnuSaveAuto_Click()

    mnuSaveAuto.Checked = True
    mnuSaveManual.Checked = False

    'Call WritePrivateProfileString("HOSP", "SAVEAUTO", "Y", App.PATH & "\INI\" & gMACH & ".ini")
    Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "SAVEAUTO", "Y")

End Sub

Private Sub mnuSaveManual_Click()

    mnuSaveAuto.Checked = False
    mnuSaveManual.Checked = True

    'Call WritePrivateProfileString("HOSP", "SAVEAUTO", "N", App.PATH & "\INI\" & gMACH & ".ini")
    Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "SAVEAUTO", "N")

End Sub

Private Sub mnuSeqno_Click()
    
    mnuBarcode.Checked = False
    mnuSeqno.Checked = True
    mnuRackPos.Checked = False
    mnuCheckBox.Checked = False

    'Call WritePrivateProfileString("HOSP", "BARUSE", "N", App.PATH & "\INI\" & gMACH & ".ini")
    'Call WritePrivateProfileString("HOSP", "RSTTYPE", "1", App.PATH & "\INI\" & gMACH & ".ini")
    
    Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "BARUSE", "N")
    Call SaveString(HKEY_CURRENT_USER, REG_MACH & "\" & "HOSP", "RSTTYPE", "1")
    
End Sub

Private Sub mnuTest_Click()
    
    frmTestSet.Show 'vbModal
    
End Sub

Private Sub mnuView_Click()
    frmScreenSet.Show 'vbModal
End Sub

Private Sub mnuWork_Click()
    
    frmWorkList.Show 'vbModal

End Sub


