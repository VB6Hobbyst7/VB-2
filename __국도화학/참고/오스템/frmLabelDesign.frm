VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{2200CD23-1176-101D-85F5-0020AF1EF604}#1.7#0"; "barcod32.ocx"
Object = "{FBFE812A-0A7F-4497-A6DA-BD90928ACE7F}#1.1#0"; "HexUniControls30.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Begin VB.Form frmLabelDesign 
   Caption         =   "Label Designer"
   ClientHeight    =   11370
   ClientLeft      =   -60
   ClientTop       =   -5100
   ClientWidth     =   19140
   BeginProperty Font 
      Name            =   "����"
      Size            =   11.25
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLabelDesign.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   758
   ScaleMode       =   3  '�ȼ�
   ScaleWidth      =   1276
   StartUpPosition =   1  '������ ���
   Begin HexUniControls.ctlUniTextBoxXP txtSpdEdit 
      Height          =   345
      Left            =   1350
      TabIndex        =   167
      Top             =   8220
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   609
      BorderColor     =   16711680
      BackColor       =   12648447
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmLabelDesign.frx":17D2A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   0
      MultiLine       =   0   'False
      Alignment       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      TrapTab         =   0   'False
      EnableContextMenu=   -1  'True
      RaiseChangeEvent=   -1  'True
      Tip             =   "frmLabelDesign.frx":17D4A
      NoHideSel       =   0   'False
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmLabelDesign.frx":17D6A
   End
   Begin VB.HScrollBar hscPaint 
      Height          =   285
      LargeChange     =   10
      Left            =   60
      SmallChange     =   10
      TabIndex        =   149
      Top             =   7860
      Width           =   10215
   End
   Begin VB.Frame Frame1 
      Height          =   7125
      Left            =   90
      TabIndex        =   147
      Top             =   660
      Width           =   10125
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '���
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   7155
         Left            =   0
         ScaleHeight     =   475
         ScaleMode       =   3  '�ȼ�
         ScaleWidth      =   675
         TabIndex        =   148
         Top             =   0
         Width           =   10155
      End
   End
   Begin MSComctlLib.Toolbar tlbMain 
      Align           =   1  '�� ����
      Height          =   555
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   19140
      _ExtentX        =   33761
      _ExtentY        =   979
      ButtonWidth     =   609
      ButtonHeight    =   926
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      Begin ActiveResizeCtl.ActiveResize ActiveResize1 
         Left            =   8790
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         Resolution      =   99
         ResizeFonts     =   0   'False
         ScreenHeight    =   1050
         ScreenWidth     =   1680
         ScreenHeightDT  =   1050
         ScreenWidthDT   =   1680
         AutoResizeControls=   0   'False
         AutoResizeOnLoad=   0   'False
         FormHeightDT    =   12180
         FormWidthDT     =   19260
         FormScaleHeightDT=   758
         FormScaleWidthDT=   1276
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   15480
         Picture         =   "frmLabelDesign.frx":17D86
         ScaleHeight     =   495
         ScaleWidth      =   3765
         TabIndex        =   138
         Top             =   30
         Width           =   3765
      End
   End
   Begin VB.VScrollBar vscPaint 
      Height          =   7185
      LargeChange     =   10
      Left            =   10290
      SmallChange     =   10
      TabIndex        =   146
      Top             =   660
      Width           =   285
   End
   Begin VB.Frame Frame3 
      Caption         =   "Hidden Value"
      Height          =   2505
      Left            =   5070
      TabIndex        =   10
      Top             =   4950
      Visible         =   0   'False
      Width           =   4485
      Begin HexUniControls.ctlUniLabel ctlUniLabel1 
         Height          =   405
         Left            =   3090
         Top             =   240
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   714
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmLabelDesign.frx":1CB38
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmLabelDesign.frx":1CB70
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmLabelDesign.frx":1CB90
      End
      Begin VB.PictureBox Picture2 
         Height          =   585
         Left            =   2160
         ScaleHeight     =   525
         ScaleWidth      =   495
         TabIndex        =   142
         Top             =   750
         Width           =   555
      End
      Begin VB.Timer tmrMove 
         Left            =   2760
         Top             =   870
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   360
         TabIndex        =   14
         Top             =   1560
         Width           =   3945
      End
      Begin BarcodLib.Barcod Barcod1 
         Height          =   315
         Left            =   360
         TabIndex        =   141
         Tag             =   "GF07J030A195"
         Top             =   300
         Width           =   2805
         _Version        =   65543
         _ExtentX        =   4948
         _ExtentY        =   556
         _StockProps     =   75
         Caption         =   "gf07j030a195"
         BackColor       =   16777215
         BarWidth        =   0
         Direction       =   0
         Style           =   7
         UPCNotches      =   0
         Alignment       =   0
         Extension       =   ""
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   3210
         Top             =   780
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.ImageList imlToolbar 
         Left            =   3750
         Top             =   690
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   7
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLabelDesign.frx":1CBAC
               Key             =   "Make"
               Object.Tag             =   "Job ���ϸ����"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLabelDesign.frx":227CE
               Key             =   "Save"
               Object.Tag             =   "LOF ��������"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLabelDesign.frx":22C20
               Key             =   "New"
               Object.Tag             =   "���θ����"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLabelDesign.frx":22F3A
               Key             =   "Open"
               Object.Tag             =   "����"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLabelDesign.frx":23814
               Key             =   "Exit"
               Object.Tag             =   "������"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLabelDesign.frx":23B2E
               Key             =   "Edit"
               Object.Tag             =   "����"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLabelDesign.frx":24DB0
               Key             =   "View"
               Object.Tag             =   "�̸�����"
            EndProperty
         EndProperty
      End
      Begin VB.Image Didim_DImg 
         Height          =   600
         Left            =   1230
         Top             =   750
         Width           =   765
      End
      Begin VB.Image Didim_SImg 
         Height          =   660
         Left            =   360
         Top             =   720
         Width           =   795
      End
   End
   Begin VB.Frame Frame11 
      Height          =   1185
      Left            =   10590
      TabIndex        =   125
      Top             =   7410
      Width           =   8505
      Begin VB.PictureBox picPrint 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         ForeColor       =   &H80000008&
         Height          =   525
         Left            =   6180
         Picture         =   "frmLabelDesign.frx":250CA
         ScaleHeight     =   525
         ScaleWidth      =   555
         TabIndex        =   136
         Top             =   390
         Width           =   555
      End
      Begin VB.TextBox txtPaperWSize 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         Height          =   345
         Left            =   4050
         MaxLength       =   5
         TabIndex        =   132
         Text            =   "3.5"
         Top             =   1320
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.TextBox txtPaperHSize 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         Height          =   345
         Left            =   2040
         MaxLength       =   5
         TabIndex        =   131
         Text            =   "7.5"
         Top             =   1320
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.OptionButton optHW 
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   6870
         TabIndex        =   130
         Top             =   1350
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.OptionButton optHW 
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   5970
         TabIndex        =   129
         Top             =   1350
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.ComboBox cmbPrinter 
         Height          =   345
         Left            =   300
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   128
         Top             =   690
         Width           =   5295
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "    �� ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   6000
         TabIndex        =   127
         Top             =   270
         Width           =   1965
      End
      Begin VB.CheckBox chkCorrect 
         Caption         =   "����������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4140
         TabIndex        =   126
         Top             =   300
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.Label lblPrint 
         Caption         =   "�����ͼ��� :"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   300
         TabIndex        =   137
         Top             =   330
         Width           =   1515
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '����
         Caption         =   "cm"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5370
         TabIndex        =   135
         Top             =   1410
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '����
         Caption         =   "cm  X"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3360
         TabIndex        =   134
         Top             =   1410
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '����
         Caption         =   "��������(����X����)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   90
         TabIndex        =   133
         Top             =   1410
         Visible         =   0   'False
         Width           =   1875
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   16080
      TabIndex        =   7
      Top             =   540
      Width           =   2985
      Begin VB.TextBox txtYmm 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         ForeColor       =   &H00400040&
         Height          =   345
         Left            =   1890
         MaxLength       =   5
         TabIndex        =   140
         Top             =   720
         Width           =   915
      End
      Begin VB.TextBox txtXmm 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         ForeColor       =   &H00400040&
         Height          =   345
         Left            =   1890
         MaxLength       =   5
         TabIndex        =   139
         Top             =   270
         Width           =   915
      End
      Begin VB.CheckBox Check2 
         Caption         =   "�̼�����"
         Height          =   345
         Left            =   270
         TabIndex        =   13
         Top             =   2790
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "�̵�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   330
         TabIndex        =   2
         Top             =   2220
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.TextBox txtYpos 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         ForeColor       =   &H00400040&
         Height          =   345
         Left            =   960
         MaxLength       =   5
         TabIndex        =   1
         Top             =   720
         Width           =   915
      End
      Begin VB.TextBox txtXpos 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         ForeColor       =   &H00400040&
         Height          =   345
         Left            =   960
         MaxLength       =   5
         TabIndex        =   0
         Top             =   270
         Width           =   915
      End
      Begin VB.Label Label10 
         Caption         =   "Y ��ǥ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   210
         TabIndex        =   9
         Top             =   780
         Width           =   675
      End
      Begin VB.Label Label9 
         Caption         =   "X ��ǥ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   8
         Top             =   330
         Width           =   735
      End
   End
   Begin VB.Frame Frame10 
      Height          =   1215
      Left            =   10590
      TabIndex        =   107
      Top             =   540
      Width           =   5475
      Begin VB.ComboBox cboType 
         Height          =   345
         Left            =   1020
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   111
         Top             =   270
         Width           =   2715
      End
      Begin VB.CheckBox chkPrint 
         Caption         =   "��¾���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3780
         TabIndex        =   110
         Top             =   270
         Width           =   1365
      End
      Begin VB.TextBox txtTitle 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1050
         MaxLength       =   20
         TabIndex        =   109
         Text            =   "LINE"
         Top             =   750
         Width           =   2625
      End
      Begin VB.TextBox txtTag 
         Appearance      =   0  '���
         Height          =   345
         Left            =   3720
         TabIndex        =   108
         Top             =   720
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.Shape Shape8 
         BorderColor     =   &H0080C0FF&
         BorderWidth     =   4
         FillColor       =   &H000080FF&
         Height          =   375
         Left            =   1020
         Shape           =   4  '�ձ� �簢��
         Top             =   720
         Width           =   2715
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  '������ ����
         Caption         =   "�׸�� "
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   120
         TabIndex        =   113
         Top             =   780
         Width           =   825
      End
      Begin VB.Label Label4 
         Alignment       =   1  '������ ����
         Caption         =   "���� "
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   120
         TabIndex        =   112
         Top             =   330
         Width           =   825
      End
   End
   Begin VB.Frame Frame9 
      Height          =   3615
      Left            =   16860
      TabIndex        =   98
      Top             =   1830
      Width           =   2205
      Begin VB.PictureBox picDelobj 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   270
         Picture         =   "frmLabelDesign.frx":253D4
         ScaleHeight     =   435
         ScaleWidth      =   495
         TabIndex        =   104
         Top             =   2040
         Width           =   495
      End
      Begin VB.PictureBox picSet 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         ForeColor       =   &H80000008&
         Height          =   525
         Left            =   270
         Picture         =   "frmLabelDesign.frx":25C9E
         ScaleHeight     =   525
         ScaleWidth      =   465
         TabIndex        =   103
         Top             =   1140
         Width           =   465
      End
      Begin VB.PictureBox picMake 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         ForeColor       =   &H80000008&
         Height          =   525
         Left            =   270
         Picture         =   "frmLabelDesign.frx":26568
         ScaleHeight     =   525
         ScaleWidth      =   495
         TabIndex        =   102
         Top             =   330
         Width           =   495
      End
      Begin VB.CommandButton cmdMake 
         Caption         =   "    �����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   101
         Top             =   240
         Width           =   1965
      End
      Begin VB.CommandButton cmdSet 
         Caption         =   "     �׸�����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   100
         Top             =   1050
         Width           =   1965
      End
      Begin VB.CommandButton cmdDelobj 
         Caption         =   "     �׸����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   99
         Top             =   1890
         Width           =   1965
      End
      Begin VB.PictureBox picUndo 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   270
         Picture         =   "frmLabelDesign.frx":26E32
         ScaleHeight     =   435
         ScaleWidth      =   495
         TabIndex        =   105
         Top             =   2850
         Width           =   495
      End
      Begin VB.CommandButton cmdUndo 
         Caption         =   "     �������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   106
         Top             =   2700
         Width           =   1965
      End
   End
   Begin VB.Frame Frame6 
      Height          =   3615
      Left            =   10590
      TabIndex        =   16
      Top             =   1830
      Width           =   6255
      Begin TabDlg.SSTab sstType 
         Height          =   3315
         Left            =   30
         TabIndex        =   17
         Top             =   210
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   5847
         _Version        =   393216
         Tabs            =   7
         Tab             =   4
         TabsPerRow      =   7
         TabHeight       =   520
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "S_Text"
         TabPicture(0)   =   "frmLabelDesign.frx":27274
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Label8(6)"
         Tab(0).Control(1)=   "Label6(0)"
         Tab(0).Control(2)=   "Label7(0)"
         Tab(0).Control(3)=   "Label8(0)"
         Tab(0).Control(4)=   "txtContent(0)"
         Tab(0).Control(5)=   "txtFontSize(0)"
         Tab(0).Control(6)=   "txtFontName(0)"
         Tab(0).Control(7)=   "chkFontBold(0)"
         Tab(0).Control(8)=   "chkFontUnder(0)"
         Tab(0).Control(9)=   "chkFontItalic(0)"
         Tab(0).Control(10)=   "txtContent9(0)"
         Tab(0).Control(11)=   "chkTStatic"
         Tab(0).Control(12)=   "cmdFont(0)"
         Tab(0).Control(13)=   "picFont(0)"
         Tab(0).Control(14)=   "Frame7"
         Tab(0).ControlCount=   15
         TabCaption(1)   =   "D_Text"
         TabPicture(1)   =   "frmLabelDesign.frx":27290
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label8(1)"
         Tab(1).Control(1)=   "Label7(1)"
         Tab(1).Control(2)=   "Label8(11)"
         Tab(1).Control(3)=   "Label6(1)"
         Tab(1).Control(4)=   "txtContent(1)"
         Tab(1).Control(5)=   "cmdFont(1)"
         Tab(1).Control(6)=   "txtContent9(1)"
         Tab(1).Control(7)=   "chkFontItalic(1)"
         Tab(1).Control(8)=   "chkFontUnder(1)"
         Tab(1).Control(9)=   "chkFontBold(1)"
         Tab(1).Control(10)=   "txtFontName(1)"
         Tab(1).Control(11)=   "txtFontSize(1)"
         Tab(1).Control(12)=   "picFont(1)"
         Tab(1).Control(13)=   "Frame8"
         Tab(1).ControlCount=   14
         TabCaption(2)   =   "S_Image"
         TabPicture(2)   =   "frmLabelDesign.frx":272AC
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Label6(2)"
         Tab(2).Control(1)=   "Label7(2)"
         Tab(2).Control(2)=   "Label8(2)"
         Tab(2).Control(3)=   "Label8(7)"
         Tab(2).Control(4)=   "Label8(8)"
         Tab(2).Control(5)=   "cmdImage(0)"
         Tab(2).Control(6)=   "txtImageWSize(2)"
         Tab(2).Control(7)=   "txtImageHSize(2)"
         Tab(2).Control(8)=   "cmdImageDevSet(0)"
         Tab(2).Control(9)=   "txtImageDevide(0)"
         Tab(2).Control(10)=   "chkIStatic"
         Tab(2).Control(11)=   "txtImageHSize(0)"
         Tab(2).Control(12)=   "txtImageName(0)"
         Tab(2).Control(13)=   "txtImageWSize(0)"
         Tab(2).Control(14)=   "picImage(0)"
         Tab(2).ControlCount=   15
         TabCaption(3)   =   "D_Image"
         TabPicture(3)   =   "frmLabelDesign.frx":272C8
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "picImage(1)"
         Tab(3).Control(1)=   "cmdImage(1)"
         Tab(3).Control(2)=   "txtImageWSize(1)"
         Tab(3).Control(3)=   "txtImageName(1)"
         Tab(3).Control(4)=   "txtImageHSize(1)"
         Tab(3).Control(5)=   "txtImageDevide(1)"
         Tab(3).Control(6)=   "cmdImageDevSet(1)"
         Tab(3).Control(7)=   "txtImageHSize(3)"
         Tab(3).Control(8)=   "txtImageWSize(3)"
         Tab(3).Control(9)=   "Label6(3)"
         Tab(3).Control(10)=   "Label7(3)"
         Tab(3).Control(11)=   "Label8(3)"
         Tab(3).Control(12)=   "Label8(9)"
         Tab(3).Control(13)=   "Label8(10)"
         Tab(3).ControlCount=   14
         TabCaption(4)   =   "Barcode"
         TabPicture(4)   =   "frmLabelDesign.frx":272E4
         Tab(4).ControlEnabled=   -1  'True
         Tab(4).Control(0)=   "Label6(7)"
         Tab(4).Control(0).Enabled=   0   'False
         Tab(4).Control(1)=   "Label7(7)"
         Tab(4).Control(1).Enabled=   0   'False
         Tab(4).Control(2)=   "Label8(12)"
         Tab(4).Control(2).Enabled=   0   'False
         Tab(4).Control(3)=   "Label8(13)"
         Tab(4).Control(3).Enabled=   0   'False
         Tab(4).Control(4)=   "Label8(4)"
         Tab(4).Control(4).Enabled=   0   'False
         Tab(4).Control(5)=   "chkBarRotate"
         Tab(4).Control(5).Enabled=   0   'False
         Tab(4).Control(6)=   "txtBarData"
         Tab(4).Control(6).Enabled=   0   'False
         Tab(4).Control(7)=   "txtBarDevide"
         Tab(4).Control(7).Enabled=   0   'False
         Tab(4).Control(8)=   "txtBarWSize"
         Tab(4).Control(8).Enabled=   0   'False
         Tab(4).Control(9)=   "txtBarHSize"
         Tab(4).Control(9).Enabled=   0   'False
         Tab(4).Control(10)=   "cboBarType"
         Tab(4).Control(10).Enabled=   0   'False
         Tab(4).ControlCount=   11
         TabCaption(5)   =   "Line"
         TabPicture(5)   =   "frmLabelDesign.frx":27300
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "chkLineRotate"
         Tab(5).Control(1)=   "txtLineHSize"
         Tab(5).Control(2)=   "txtLineWSize"
         Tab(5).Control(3)=   "Label7(4)"
         Tab(5).Control(4)=   "Label8(5)"
         Tab(5).ControlCount=   5
         TabCaption(6)   =   "RFID"
         TabPicture(6)   =   "frmLabelDesign.frx":2731C
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "chkFontItalic(2)"
         Tab(6).Control(1)=   "chkFontUnder(2)"
         Tab(6).Control(2)=   "chkFontBold(2)"
         Tab(6).Control(3)=   "txtFontName(2)"
         Tab(6).Control(4)=   "txtFontSize(2)"
         Tab(6).Control(5)=   "picFont(2)"
         Tab(6).Control(6)=   "Frame12"
         Tab(6).Control(7)=   "txtContent(2)"
         Tab(6).Control(8)=   "cmdFont(2)"
         Tab(6).Control(9)=   "Label8(15)"
         Tab(6).Control(10)=   "Label7(5)"
         Tab(6).Control(11)=   "Label8(14)"
         Tab(6).Control(12)=   "Label6(4)"
         Tab(6).ControlCount=   13
         Begin VB.CheckBox chkFontItalic 
            Caption         =   "���� "
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   -70740
            TabIndex        =   160
            Top             =   1110
            Width           =   1065
         End
         Begin VB.CheckBox chkFontUnder 
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   -71700
            TabIndex        =   159
            Top             =   1110
            Width           =   825
         End
         Begin VB.CheckBox chkFontBold 
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   -72690
            TabIndex        =   158
            Top             =   1110
            Width           =   825
         End
         Begin VB.TextBox txtFontName 
            Appearance      =   0  '���
            Height          =   345
            Index           =   2
            Left            =   -73800
            MaxLength       =   20
            TabIndex        =   157
            Top             =   600
            Width           =   1995
         End
         Begin VB.TextBox txtFontSize 
            Alignment       =   2  '��� ����
            Appearance      =   0  '���
            Height          =   345
            Index           =   2
            Left            =   -73800
            MaxLength       =   3
            TabIndex        =   156
            Top             =   1050
            Width           =   1005
         End
         Begin VB.PictureBox picFont 
            Appearance      =   0  '���
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  '����
            ForeColor       =   &H80000008&
            Height          =   480
            Index           =   2
            Left            =   -71790
            Picture         =   "frmLabelDesign.frx":27338
            ScaleHeight     =   480
            ScaleWidth      =   480
            TabIndex        =   155
            Top             =   2520
            Width           =   480
         End
         Begin VB.Frame Frame12 
            Height          =   465
            Left            =   -73800
            TabIndex        =   150
            Top             =   1860
            Visible         =   0   'False
            Width           =   4095
            Begin VB.OptionButton optITRotate 
               Caption         =   "0��"
               Height          =   255
               Index           =   7
               Left            =   180
               TabIndex        =   154
               Top             =   150
               Value           =   -1  'True
               Width           =   705
            End
            Begin VB.OptionButton optITRotate 
               Caption         =   "90��"
               Height          =   255
               Index           =   6
               Left            =   930
               TabIndex        =   153
               Top             =   150
               Width           =   735
            End
            Begin VB.OptionButton optITRotate 
               Caption         =   "180��"
               Height          =   255
               Index           =   5
               Left            =   1740
               TabIndex        =   152
               Top             =   150
               Width           =   825
            End
            Begin VB.OptionButton optITRotate 
               Caption         =   "270��"
               Height          =   255
               Index           =   4
               Left            =   2640
               TabIndex        =   151
               Top             =   150
               Width           =   825
            End
         End
         Begin VB.Frame Frame8 
            Height          =   465
            Left            =   -73800
            TabIndex        =   86
            Top             =   1860
            Visible         =   0   'False
            Width           =   4095
            Begin VB.OptionButton optITRotate 
               Caption         =   "270��"
               Height          =   255
               Index           =   3
               Left            =   2640
               TabIndex        =   90
               Top             =   150
               Width           =   825
            End
            Begin VB.OptionButton optITRotate 
               Caption         =   "180��"
               Height          =   255
               Index           =   2
               Left            =   1740
               TabIndex        =   89
               Top             =   150
               Width           =   825
            End
            Begin VB.OptionButton optITRotate 
               Caption         =   "90��"
               Height          =   255
               Index           =   1
               Left            =   930
               TabIndex        =   88
               Top             =   150
               Width           =   735
            End
            Begin VB.OptionButton optITRotate 
               Caption         =   "0��"
               Height          =   255
               Index           =   0
               Left            =   180
               TabIndex        =   87
               Top             =   150
               Value           =   -1  'True
               Width           =   705
            End
         End
         Begin VB.Frame Frame7 
            Height          =   465
            Left            =   -73800
            TabIndex        =   81
            Top             =   1860
            Visible         =   0   'False
            Width           =   4095
            Begin VB.OptionButton optSTRotate 
               Caption         =   "270��"
               Height          =   255
               Index           =   3
               Left            =   2640
               TabIndex        =   85
               Top             =   150
               Width           =   825
            End
            Begin VB.OptionButton optSTRotate 
               Caption         =   "180��"
               Height          =   255
               Index           =   2
               Left            =   1740
               TabIndex        =   84
               Top             =   150
               Width           =   825
            End
            Begin VB.OptionButton optSTRotate 
               Caption         =   "90��"
               Height          =   255
               Index           =   1
               Left            =   930
               TabIndex        =   83
               Top             =   150
               Width           =   735
            End
            Begin VB.OptionButton optSTRotate 
               Caption         =   "0��"
               Height          =   255
               Index           =   0
               Left            =   180
               TabIndex        =   82
               Top             =   150
               Value           =   -1  'True
               Width           =   705
            End
         End
         Begin VB.PictureBox picFont 
            Appearance      =   0  '���
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  '����
            ForeColor       =   &H80000008&
            Height          =   480
            Index           =   1
            Left            =   -71790
            Picture         =   "frmLabelDesign.frx":28ABA
            ScaleHeight     =   480
            ScaleWidth      =   480
            TabIndex        =   75
            Top             =   2520
            Width           =   480
         End
         Begin VB.PictureBox picFont 
            Appearance      =   0  '���
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  '����
            ForeColor       =   &H80000008&
            Height          =   480
            Index           =   0
            Left            =   -71790
            Picture         =   "frmLabelDesign.frx":2A23C
            ScaleHeight     =   480
            ScaleWidth      =   480
            TabIndex        =   74
            Top             =   2520
            Width           =   480
         End
         Begin VB.PictureBox picImage 
            Appearance      =   0  '���
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  '����
            ForeColor       =   &H80000008&
            Height          =   480
            Index           =   0
            Left            =   -71790
            Picture         =   "frmLabelDesign.frx":2B9BE
            ScaleHeight     =   480
            ScaleWidth      =   480
            TabIndex        =   72
            Top             =   2520
            Width           =   480
         End
         Begin VB.PictureBox picImage 
            Appearance      =   0  '���
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  '����
            ForeColor       =   &H80000008&
            Height          =   480
            Index           =   1
            Left            =   -71790
            Picture         =   "frmLabelDesign.frx":2D140
            ScaleHeight     =   480
            ScaleWidth      =   480
            TabIndex        =   66
            Top             =   2520
            Width           =   480
         End
         Begin VB.CheckBox chkLineRotate 
            Caption         =   "ȸ��"
            Height          =   345
            Left            =   -73590
            TabIndex        =   56
            Top             =   1890
            Width           =   1275
         End
         Begin VB.TextBox txtLineHSize 
            Alignment       =   2  '��� ����
            Appearance      =   0  '���
            Height          =   345
            Left            =   -73590
            MaxLength       =   1
            TabIndex        =   55
            Top             =   930
            Width           =   2505
         End
         Begin VB.TextBox txtLineWSize 
            Alignment       =   2  '��� ����
            Appearance      =   0  '���
            Height          =   345
            Left            =   -73590
            MaxLength       =   5
            TabIndex        =   54
            Top             =   1410
            Width           =   2505
         End
         Begin VB.ComboBox cboBarType 
            Height          =   345
            Left            =   1290
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   53
            Top             =   600
            Width           =   3225
         End
         Begin VB.TextBox txtBarHSize 
            Alignment       =   2  '��� ����
            Appearance      =   0  '���
            Height          =   345
            Left            =   3150
            MaxLength       =   5
            TabIndex        =   51
            Top             =   1500
            Width           =   1365
         End
         Begin VB.TextBox txtBarWSize 
            Alignment       =   2  '��� ����
            Appearance      =   0  '���
            Height          =   345
            Left            =   1290
            MaxLength       =   5
            TabIndex        =   50
            Top             =   1500
            Width           =   1245
         End
         Begin VB.TextBox txtBarDevide 
            Appearance      =   0  '���
            Height          =   345
            Left            =   1290
            MaxLength       =   1
            TabIndex        =   49
            Top             =   1050
            Visible         =   0   'False
            Width           =   3225
         End
         Begin VB.TextBox txtBarData 
            Appearance      =   0  '���
            Height          =   345
            Left            =   1290
            MaxLength       =   20
            TabIndex        =   52
            Top             =   1980
            Width           =   3225
         End
         Begin VB.CheckBox chkBarRotate 
            Caption         =   "ȸ��"
            Height          =   345
            Left            =   1290
            TabIndex        =   48
            Top             =   2430
            Width           =   1665
         End
         Begin VB.CommandButton cmdImage 
            Caption         =   "      �̹��� ã��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   675
            Index           =   1
            Left            =   -71910
            Picture         =   "frmLabelDesign.frx":2E8C2
            TabIndex        =   47
            Top             =   2430
            Width           =   2205
         End
         Begin VB.TextBox txtImageWSize 
            Alignment       =   2  '��� ����
            Appearance      =   0  '���
            Height          =   345
            Index           =   1
            Left            =   -73800
            MaxLength       =   5
            TabIndex        =   46
            Top             =   1050
            Width           =   1605
         End
         Begin VB.TextBox txtImageName 
            Appearance      =   0  '���
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   1
            Left            =   -73800
            TabIndex        =   45
            Top             =   600
            Width           =   4035
         End
         Begin VB.TextBox txtImageHSize 
            Alignment       =   2  '��� ����
            Appearance      =   0  '���
            Height          =   345
            Index           =   1
            Left            =   -73800
            MaxLength       =   5
            TabIndex        =   44
            Top             =   1500
            Width           =   1605
         End
         Begin VB.TextBox txtImageWSize 
            Alignment       =   2  '��� ����
            Appearance      =   0  '���
            Height          =   345
            Index           =   0
            Left            =   -73800
            MaxLength       =   5
            TabIndex        =   43
            Top             =   1050
            Width           =   1605
         End
         Begin VB.TextBox txtImageName 
            Appearance      =   0  '���
            Height          =   345
            Index           =   0
            Left            =   -73800
            TabIndex        =   42
            Top             =   600
            Width           =   4035
         End
         Begin VB.TextBox txtImageHSize 
            Alignment       =   2  '��� ����
            Appearance      =   0  '���
            Height          =   345
            Index           =   0
            Left            =   -73800
            MaxLength       =   5
            TabIndex        =   41
            Top             =   1500
            Width           =   1605
         End
         Begin VB.CheckBox chkIStatic 
            Caption         =   "������ ����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   -73800
            TabIndex        =   40
            Top             =   2400
            Width           =   1665
         End
         Begin VB.TextBox txtFontSize 
            Alignment       =   2  '��� ����
            Appearance      =   0  '���
            Height          =   345
            Index           =   1
            Left            =   -73800
            MaxLength       =   3
            TabIndex        =   39
            Top             =   1050
            Width           =   1005
         End
         Begin VB.TextBox txtFontName 
            Appearance      =   0  '���
            Height          =   345
            Index           =   1
            Left            =   -73800
            MaxLength       =   20
            TabIndex        =   38
            Top             =   600
            Width           =   1995
         End
         Begin VB.CheckBox chkFontBold 
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   -72690
            TabIndex        =   37
            Top             =   1110
            Width           =   825
         End
         Begin VB.CheckBox chkFontUnder 
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   -71700
            TabIndex        =   36
            Top             =   1110
            Width           =   825
         End
         Begin VB.CheckBox chkFontItalic 
            Caption         =   "���� "
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   -70740
            TabIndex        =   35
            Top             =   1110
            Width           =   1065
         End
         Begin VB.TextBox txtContent9 
            Appearance      =   0  '���
            Height          =   345
            Index           =   1
            Left            =   -74820
            TabIndex        =   34
            Top             =   2730
            Visible         =   0   'False
            Width           =   4065
         End
         Begin VB.CommandButton cmdFont 
            Caption         =   "     ��Ʈ ����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Index           =   0
            Left            =   -71910
            TabIndex        =   33
            Top             =   2460
            Width           =   2205
         End
         Begin VB.CheckBox chkTStatic 
            Caption         =   "������ ����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   -73800
            TabIndex        =   32
            Top             =   2400
            Width           =   1665
         End
         Begin VB.TextBox txtContent9 
            Appearance      =   0  '���
            Height          =   345
            Index           =   0
            Left            =   -74850
            TabIndex        =   31
            Top             =   2880
            Visible         =   0   'False
            Width           =   4065
         End
         Begin VB.CheckBox chkFontItalic 
            Caption         =   "���� "
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   -70740
            TabIndex        =   30
            Top             =   1140
            Width           =   1065
         End
         Begin VB.CheckBox chkFontUnder 
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   -71700
            TabIndex        =   29
            Top             =   1110
            Width           =   825
         End
         Begin VB.CheckBox chkFontBold 
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   -72690
            TabIndex        =   28
            Top             =   1110
            Width           =   825
         End
         Begin VB.TextBox txtFontName 
            Appearance      =   0  '���
            Height          =   345
            Index           =   0
            Left            =   -73800
            MaxLength       =   20
            TabIndex        =   27
            Top             =   600
            Width           =   1995
         End
         Begin VB.TextBox txtFontSize 
            Alignment       =   2  '��� ����
            Appearance      =   0  '���
            Height          =   345
            Index           =   0
            Left            =   -73800
            MaxLength       =   3
            TabIndex        =   26
            Top             =   1050
            Width           =   1005
         End
         Begin VB.TextBox txtImageDevide 
            Alignment       =   2  '��� ����
            Appearance      =   0  '���
            Height          =   345
            Index           =   0
            Left            =   -73800
            MaxLength       =   5
            TabIndex        =   25
            Top             =   1950
            Width           =   585
         End
         Begin VB.CommandButton cmdImageDevSet 
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
            Height          =   375
            Index           =   0
            Left            =   -72870
            TabIndex        =   24
            Top             =   1950
            Width           =   705
         End
         Begin VB.TextBox txtImageDevide 
            Alignment       =   2  '��� ����
            Appearance      =   0  '���
            Height          =   345
            Index           =   1
            Left            =   -73800
            MaxLength       =   5
            TabIndex        =   23
            Top             =   1950
            Width           =   585
         End
         Begin VB.CommandButton cmdImageDevSet 
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
            Height          =   375
            Index           =   1
            Left            =   -72870
            TabIndex        =   22
            Top             =   1950
            Width           =   705
         End
         Begin VB.TextBox txtImageHSize 
            Appearance      =   0  '���
            BackColor       =   &H80000000&
            Enabled         =   0   'False
            Height          =   345
            Index           =   2
            Left            =   -72180
            MaxLength       =   5
            TabIndex        =   21
            Top             =   1500
            Width           =   1605
         End
         Begin VB.TextBox txtImageWSize 
            Appearance      =   0  '���
            BackColor       =   &H80000000&
            Enabled         =   0   'False
            Height          =   345
            Index           =   2
            Left            =   -72180
            MaxLength       =   5
            TabIndex        =   20
            Top             =   1050
            Width           =   1605
         End
         Begin VB.TextBox txtImageHSize 
            Appearance      =   0  '���
            BackColor       =   &H80000000&
            Enabled         =   0   'False
            Height          =   345
            Index           =   3
            Left            =   -72180
            MaxLength       =   5
            TabIndex        =   19
            Top             =   1500
            Width           =   1605
         End
         Begin VB.TextBox txtImageWSize 
            Appearance      =   0  '���
            BackColor       =   &H80000000&
            Enabled         =   0   'False
            Height          =   345
            Index           =   3
            Left            =   -72180
            MaxLength       =   5
            TabIndex        =   18
            Top             =   1050
            Width           =   1605
         End
         Begin VB.CommandButton cmdImage 
            Caption         =   "      �̹��� ã��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   675
            Index           =   0
            Left            =   -71910
            Picture         =   "frmLabelDesign.frx":2EF84
            TabIndex        =   73
            Top             =   2430
            Width           =   2205
         End
         Begin VB.CommandButton cmdFont 
            Caption         =   "     ��Ʈ ����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Index           =   1
            Left            =   -71910
            TabIndex        =   76
            Top             =   2460
            Width           =   2205
         End
         Begin HexUniControls.ctlUniTextBoxXP txtContent 
            Height          =   375
            Index           =   0
            Left            =   -73800
            TabIndex        =   143
            Top             =   1470
            Width           =   4065
            _ExtentX        =   7170
            _ExtentY        =   661
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmLabelDesign.frx":2F646
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   0
            MultiLine       =   0   'False
            Alignment       =   0
            ScrollBars      =   0
            PasswordChar    =   ""
            TrapTab         =   0   'False
            EnableContextMenu=   -1  'True
            RaiseChangeEvent=   -1  'True
            Tip             =   "frmLabelDesign.frx":2F666
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmLabelDesign.frx":2F686
         End
         Begin HexUniControls.ctlUniTextBoxXP txtContent 
            Height          =   405
            Index           =   1
            Left            =   -73800
            TabIndex        =   144
            Top             =   1440
            Width           =   4065
            _ExtentX        =   7170
            _ExtentY        =   714
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmLabelDesign.frx":2F6A2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   0
            MultiLine       =   0   'False
            Alignment       =   0
            ScrollBars      =   0
            PasswordChar    =   ""
            TrapTab         =   0   'False
            EnableContextMenu=   -1  'True
            RaiseChangeEvent=   -1  'True
            Tip             =   "frmLabelDesign.frx":2F6C2
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmLabelDesign.frx":2F6E2
         End
         Begin HexUniControls.ctlUniTextBoxXP txtContent 
            Height          =   405
            Index           =   2
            Left            =   -73800
            TabIndex        =   162
            Top             =   1440
            Width           =   4065
            _ExtentX        =   7170
            _ExtentY        =   714
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmLabelDesign.frx":2F6FE
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   0
            MultiLine       =   0   'False
            Alignment       =   0
            ScrollBars      =   0
            PasswordChar    =   ""
            TrapTab         =   0   'False
            EnableContextMenu=   -1  'True
            RaiseChangeEvent=   -1  'True
            Tip             =   "frmLabelDesign.frx":2F71E
            NoHideSel       =   0   'False
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmLabelDesign.frx":2F73E
         End
         Begin VB.CommandButton cmdFont 
            Caption         =   "     ��Ʈ ����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Index           =   2
            Left            =   -71910
            TabIndex        =   161
            Top             =   2460
            Width           =   2205
         End
         Begin VB.Label Label8 
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   15
            Left            =   -74340
            TabIndex        =   166
            Top             =   1590
            Width           =   435
         End
         Begin VB.Label Label7 
            Caption         =   "��Ʈũ�� "
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   5
            Left            =   -74730
            TabIndex        =   165
            Top             =   1140
            Width           =   825
         End
         Begin VB.Label Label8 
            Caption         =   "ȸ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   14
            Left            =   -74340
            TabIndex        =   164
            Top             =   2040
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.Label Label6 
            Caption         =   "��Ʈ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   -74520
            TabIndex        =   163
            Top             =   690
            Width           =   675
         End
         Begin VB.Label Label6 
            Caption         =   "��Ʈ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   -74520
            TabIndex        =   79
            Top             =   690
            Width           =   675
         End
         Begin VB.Label Label7 
            Caption         =   "������"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   4
            Left            =   -74400
            TabIndex        =   97
            Top             =   990
            Width           =   765
         End
         Begin VB.Label Label8 
            Caption         =   "������"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   5
            Left            =   -74400
            TabIndex        =   96
            Top             =   1440
            Width           =   585
         End
         Begin VB.Label Label8 
            Caption         =   "���ڵ尪"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   330
            TabIndex        =   95
            Top             =   2070
            Width           =   825
         End
         Begin VB.Label Label8 
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   13
            Left            =   2670
            TabIndex        =   94
            Top             =   1590
            Width           =   405
         End
         Begin VB.Label Label8 
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   12
            Left            =   720
            TabIndex        =   93
            Top             =   1560
            Width           =   405
         End
         Begin VB.Label Label7 
            Caption         =   "���ݺ��� "
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   7
            Left            =   300
            TabIndex        =   92
            Top             =   1110
            Visible         =   0   'False
            Width           =   825
         End
         Begin VB.Label Label6 
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   7
            Left            =   690
            TabIndex        =   91
            Top             =   660
            Width           =   405
         End
         Begin VB.Label Label8 
            Caption         =   "ȸ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   11
            Left            =   -74340
            TabIndex        =   80
            Top             =   2040
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.Label Label7 
            Caption         =   "��Ʈũ�� "
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   -74730
            TabIndex        =   78
            Top             =   1140
            Width           =   825
         End
         Begin VB.Label Label8 
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   -74340
            TabIndex        =   77
            Top             =   1590
            Width           =   435
         End
         Begin VB.Label Label8 
            Alignment       =   1  '������ ����
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   8
            Left            =   -73140
            TabIndex        =   71
            Top             =   2040
            Width           =   165
         End
         Begin VB.Label Label8 
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   7
            Left            =   -74280
            TabIndex        =   70
            Top             =   2040
            Width           =   405
         End
         Begin VB.Label Label8 
            Caption         =   "���α���"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   2
            Left            =   -74670
            TabIndex        =   69
            Top             =   1560
            Width           =   795
         End
         Begin VB.Label Label7 
            Caption         =   "���α��� "
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   2
            Left            =   -74700
            TabIndex        =   68
            Top             =   1110
            Width           =   825
         End
         Begin VB.Label Label6 
            Caption         =   "�̹�����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   2
            Left            =   -74700
            TabIndex        =   67
            Top             =   660
            Width           =   825
         End
         Begin VB.Label Label6 
            Caption         =   "�̹�����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   3
            Left            =   -74700
            TabIndex        =   65
            Top             =   660
            Width           =   825
         End
         Begin VB.Label Label7 
            Caption         =   "���α��� "
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   3
            Left            =   -74700
            TabIndex        =   64
            Top             =   1110
            Width           =   825
         End
         Begin VB.Label Label8 
            Caption         =   "���α���"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   3
            Left            =   -74670
            TabIndex        =   63
            Top             =   1560
            Width           =   795
         End
         Begin VB.Label Label8 
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   -74340
            TabIndex        =   62
            Top             =   1590
            Width           =   435
         End
         Begin VB.Label Label7 
            Caption         =   "��Ʈũ�� "
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   -74730
            TabIndex        =   61
            Top             =   1140
            Width           =   825
         End
         Begin VB.Label Label6 
            Caption         =   "��Ʈ�� "
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   -74520
            TabIndex        =   60
            Top             =   690
            Width           =   615
         End
         Begin VB.Label Label8 
            Caption         =   "ȸ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   6
            Left            =   -74340
            TabIndex        =   59
            Top             =   2040
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.Label Label8 
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   9
            Left            =   -74280
            TabIndex        =   58
            Top             =   2040
            Width           =   405
         End
         Begin VB.Label Label8 
            Alignment       =   1  '������ ����
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   10
            Left            =   -73140
            TabIndex        =   57
            Top             =   2040
            Width           =   165
         End
      End
   End
   Begin VB.Frame Frame4 
      Height          =   2025
      Left            =   10590
      TabIndex        =   12
      Top             =   5400
      Width           =   8505
      Begin VB.Frame Frame5 
         Height          =   1305
         Left            =   4470
         TabIndex        =   117
         Top             =   420
         Width           =   3585
         Begin VB.OptionButton optDevide 
            Caption         =   "2��"
            Height          =   315
            Index           =   1
            Left            =   5760
            TabIndex        =   123
            Tag             =   "2"
            Top             =   180
            Visible         =   0   'False
            Width           =   885
         End
         Begin VB.OptionButton optDevide 
            Caption         =   "1��"
            Height          =   315
            Index           =   0
            Left            =   3660
            TabIndex        =   122
            Tag             =   "1"
            Top             =   180
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.OptionButton optDevide 
            Caption         =   "1.4��"
            Height          =   315
            Index           =   2
            Left            =   4620
            TabIndex        =   121
            Tag             =   "1.4"
            Top             =   180
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CommandButton cmdDevide 
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   26.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Index           =   1
            Left            =   2730
            TabIndex        =   120
            Top             =   480
            Width           =   555
         End
         Begin VB.CommandButton cmdDevide 
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   26.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Index           =   0
            Left            =   1140
            TabIndex        =   119
            Top             =   480
            Width           =   555
         End
         Begin VB.TextBox txtDevide 
            Alignment       =   2  '��� ����
            Appearance      =   0  '���
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "����"
               Size            =   21.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   555
            Left            =   1740
            TabIndex        =   118
            Top             =   450
            Width           =   915
         End
         Begin VB.Shape Shape7 
            BorderColor     =   &H00C0FFFF&
            BorderWidth     =   4
            FillColor       =   &H000080FF&
            Height          =   585
            Left            =   2670
            Shape           =   4  '�ձ� �簢��
            Top             =   450
            Width           =   675
         End
         Begin VB.Shape Shape6 
            BorderColor     =   &H00C0FFFF&
            BorderWidth     =   4
            FillColor       =   &H000080FF&
            Height          =   585
            Left            =   1110
            Shape           =   4  '�ձ� �簢��
            Top             =   450
            Width           =   615
         End
         Begin VB.Label Label6 
            Alignment       =   1  '������ ����
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   24
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   6
            Left            =   90
            TabIndex        =   124
            Top             =   510
            Width           =   975
         End
      End
      Begin VB.CheckBox chkDetail 
         Caption         =   "�̼�����"
         Height          =   345
         Left            =   2640
         TabIndex        =   116
         Top             =   960
         Width           =   1275
      End
      Begin VB.CheckBox chkContinue 
         Caption         =   "�����̵�"
         Height          =   345
         Left            =   2640
         TabIndex        =   115
         Top             =   1410
         Width           =   1275
      End
      Begin VB.CheckBox chkChoice 
         Caption         =   "�����̵�"
         Height          =   345
         Left            =   2640
         TabIndex        =   114
         Top             =   540
         Width           =   1275
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   26.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   1
         Left            =   1740
         TabIndex        =   4
         Top             =   810
         Width           =   585
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   26.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   3
         Left            =   1110
         TabIndex        =   6
         Top             =   1350
         Width           =   585
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   26.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   2
         Left            =   1110
         TabIndex        =   5
         Top             =   240
         Width           =   585
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   26.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   0
         Left            =   480
         TabIndex        =   3
         Top             =   810
         Width           =   585
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H00C0FFFF&
         BorderWidth     =   4
         FillColor       =   &H000080FF&
         Height          =   585
         Left            =   1080
         Shape           =   4  '�ձ� �簢��
         Top             =   1350
         Width           =   675
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00C0FFFF&
         BorderWidth     =   4
         FillColor       =   &H000080FF&
         Height          =   585
         Left            =   1710
         Shape           =   4  '�ձ� �簢��
         Top             =   780
         Width           =   675
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00C0FFFF&
         BorderWidth     =   4
         FillColor       =   &H000080FF&
         Height          =   585
         Left            =   450
         Shape           =   4  '�ձ� �簢��
         Top             =   780
         Width           =   675
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00C0FFFF&
         BorderWidth     =   4
         FillColor       =   &H000080FF&
         Height          =   585
         Left            =   1080
         Shape           =   4  '�ձ� �簢��
         Top             =   210
         Width           =   675
      End
   End
   Begin HexUniControls.ctlUniGrid spdList 
      Height          =   2715
      Left            =   30
      TabIndex        =   145
      Top             =   8610
      Width           =   19065
      _ExtentX        =   33629
      _ExtentY        =   4789
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      SelBackColor    =   -2147483635
      SelForeColor    =   -2147483634
      HeaderBackColor =   -2147483633
      HeaderForeColor =   -2147483630
      Tip             =   "frmLabelDesign.frx":2F75A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MinRowHeight    =   0
      HeaderStyle     =   -1
      SelectorStyle   =   -1
      ShowFocus       =   -1  'True
      AllowMultiLineText=   0   'False
      Enabled         =   -1  'True
   End
   Begin FPSpread.vaSpread spdList7 
      Height          =   3795
      Left            =   7650
      TabIndex        =   11
      Top             =   10560
      Visible         =   0   'False
      Width           =   19005
      _Version        =   196608
      _ExtentX        =   33523
      _ExtentY        =   6694
      _StockProps     =   64
      AutoCalc        =   0   'False
      AutoClipboard   =   0   'False
      BackColorStyle  =   1
      ColHeaderDisplay=   0
      ColsFrozen      =   3
      DisplayRowHeaders=   0   'False
      EditEnterAction =   5
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����ü"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FormulaSync     =   0   'False
      GridShowHoriz   =   0   'False
      GridSolid       =   0   'False
      MaxCols         =   29
      MaxRows         =   5
      MoveActiveOnFocus=   0   'False
      OperationMode   =   2
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   14735309
      SpreadDesigner  =   "frmLabelDesign.frx":2F77A
      ScrollBarTrack  =   3
   End
   Begin VB.Label Label8 
      Caption         =   "������ �Է�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   16
      Left            =   120
      TabIndex        =   168
      Top             =   8280
      Width           =   1125
   End
   Begin VB.Menu MnuFile 
      Caption         =   "������(&F)"
      Begin VB.Menu mnuNew 
         Caption         =   "�ű�"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "����"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "����"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuClose 
         Caption         =   "����"
      End
   End
   Begin VB.Menu mnuExe 
      Caption         =   "�ѽ���(&R)"
      Begin VB.Menu mnuView 
         Caption         =   "�̸�����"
      End
      Begin VB.Menu mnuMake 
         Caption         =   "�۾����ϻ���"
      End
   End
   Begin VB.Menu mnuSet 
      Caption         =   "�Ѽ���(&U)"
   End
End
Attribute VB_Name = "frmLabelDesign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================================
'  ���α׷� : ������ ��Ǯ��Ʈ ���� �� [���ڵ巹�̾ƿ� �ҷ�����/����/����,���� ��Ʈ�� ����/�̺�Ʈ ó��]
'  �� �� �� : frmLabelDesign.frm
'  �� �� �� : 2011.09.21
'  �� �� �� : ������
'  Ȩ������ : http://www.didiminfoinfo.co.kr
'  ��    �� :
'  �����̷� :
'===============================================================================
Option Explicit


Private m_ColCommandButton              As Collection               ' ���� ���� ��Ʈ�� ������ ���� �÷���
Private WithEvents ClsEventMonitor      As ClassEventMonitor        ' �̺�Ʈ ������ ���� Ŭ����
Attribute ClsEventMonitor.VB_VarHelpID = -1

Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long


'==== API ���� ���� ���� =================================================
Const FW_NORMAL = 400
Const DEFAULT_CHARSET = 1
Const OUT_DEFAULT_PRECIS = 0
Const CLIP_DEFAULT_PRECIS = 0
Const DEFAULT_QUALITY = 0
Const DEFAULT_PITCH = 0
Const FF_ROMAN = 16
Const CF_PRINTERFONTS = &H2
Const CF_SCREENFONTS = &H1
Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Const CF_EFFECTS = &H100&
Const CF_FORCEFONTEXIST = &H10000
Const CF_INITTOLOGFONTSTRUCT = &H40&
Const CF_LIMITSIZE = &H2000&
Const REGULAR_FONTTYPE = &H400
Const LF_FACESIZE = 32
Const CCHDEVICENAME = 32
Const CCHFORMNAME = 32
Const GMEM_MOVEABLE = &H2
Const GMEM_ZEROINIT = &H40
Const DM_DUPLEX = &H1000&
Const DM_ORIENTATION = &H1&
Const PD_PRINTSETUP = &H40
Const PD_DISABLEPRINTTOFILE = &H80000

Private Const ANSI_CHARSET = 0
Private Const VARIABLE_PITCH = 2
Private Const FF_DONTCARE = 0
Private Const FW_BOLD = 700
Private Const LOGPIXELSY = 90


Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type PAGESETUPDLG
    lStructSize As Long
    hwndOwner As Long
    hDevMode As Long
    hDevNames As Long
    flags As Long
    ptPaperSize As POINTAPI
    rtMinMargin As RECT
    rtMargin As RECT
    hInstance As Long
    lCustData As Long
    lpfnPageSetupHook As Long
    lpfnPagePaintHook As Long
    lpPageSetupTemplateName As String
    hPageSetupTemplate As Long
End Type

Private Type CHOOSECOLOR
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName As String * 31
End Type

Private Type CHOOSEFONT
        lStructSize As Long
        hwndOwner As Long          '  caller's window handle
        hdc As Long                '  printer DC/IC or NULL
        lpLogFont As Long          '  ptr. to a LOGFONT struct
        iPointSize As Long         '  10 * size in points of selected font
        flags As Long              '  enum. type flags
        rgbColors As Long          '  returned text color
        lCustData As Long          '  data passed to hook fn.
        lpfnHook As Long           '  ptr. to hook function
        lpTemplateName As String     '  custom template name
        hInstance As Long          '  instance handle of.EXE that
                                       '    contains cust. dlg. template
        lpszStyle As String          '  return the style field here
                                       '  must be LF_FACESIZE or bigger
        nFontType As Integer          '  same value reported to the EnumFonts
                                       '    call back with the extra FONTTYPE_
                                       '    bits added
        MISSING_ALIGNMENT As Integer
        nSizeMin As Long           '  minimum pt size allowed &
        nSizeMax As Long           '  max pt size allowed if
                                       '    CF_LIMITSIZE is used
End Type

Private Type PRINTDLG_TYPE
    lStructSize As Long
    hwndOwner As Long
    hDevMode As Long
    hDevNames As Long
    hdc As Long
    flags As Long
    nFromPage As Integer
    nToPage As Integer
    nMinPage As Integer
    nMaxPage As Integer
    nCopies As Integer
    hInstance As Long
    lCustData As Long
    lpfnPrintHook As Long
    lpfnSetupHook As Long
    lpPrintTemplateName As String
    lpSetupTemplateName As String
    hPrintTemplate As Long
    hSetupTemplate As Long
End Type

Private Type DEVNAMES_TYPE
    wDriverOffset As Integer
    wDeviceOffset As Integer
    wOutputOffset As Integer
    wDefault As Integer
    extra As String * 100
End Type

Private Type DEVMODE_TYPE
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

Private Declare Function CHOOSECOLOR Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function PrintDialog Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PRINTDLG_TYPE) As Long
Private Declare Function PAGESETUPDLG Lib "comdlg32.dll" Alias "PageSetupDlgA" (pPagesetupdlg As PAGESETUPDLG) As Long
Private Declare Function CHOOSEFONT Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As CHOOSEFONT) As Long
'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long

Dim OFName As OPENFILENAME
Dim CustomColors() As Byte
'==== API ���� ���� ���� =================================================


Public Enum eGlobalImgIndex
    egii_x = 1
    egii_help
    egii_big
    egii_vbico
    
    eim_copy
    eim_paste
    eim_cut
    eim_open
    eim_save
    eim_undo
    eim_redo
    eim_new
End Enum


Dim gblPicVval As Long
Dim gblPicHval As Long

Dim gColWidth(30) As Integer

'Public Function DrawRotatedText(lhDC As Long, FontInfo As StdFont, iRot As Integer, sText As String, lX As Long, lY As Long) As Boolean
'
''On Error GoTo DrawRotatedText_E
'
''Parameters:
''   lhDC     - The device context to draw the text on
''   FontInfo - A font structure with the font to use
''   iRot     - Rotation in tenths of degrees (900 equals 90 degrees)
''   sText    - The text to draw
''   lX       - X coordinate of starting point (in pixels)
''   lY       - Y coordinate of starting point (in pixels)
''
''Return value:
''   returns true if successful, false otherwise
''
''Last modified: June 9, 1999
''Special thanks to: Sebastian Strand
'
'Dim hlFont As Long, hlOld As Long
'Dim uLogFont As LOGFONT, b As Byte
'Dim abChars() As Byte
'
''Fill logfont structure with proper font data
'With uLogFont
'
'.lfCharSet = ANSI_CHARSET
'.lfClipPrecision = CLIP_DEFAULT_PRECIS
'.lfEscapement = iRot
'
''We can't assign directly to fixed length array
''so we have to use a temp array and copy the chars
''one by one
'abChars = StrConv(FontInfo.Name, vbFromUnicode)
'For b = 0 To IIf(UBound(abChars) > UBound(.lfFaceName), UBound(.lfFaceName), UBound(abChars))
'.lfFaceName(b) = abChars(b)
'Next b
'
'.lfHeight = FontInfo.Size / 72 * GetDeviceCaps(lhDC, LOGPIXELSY)
'.lfWidth = 0 'When zero windows calculates proper width based on the height setting
'.lfItalic = Abs(FontInfo.Italic)
'.lfOrientation = .lfEscapement
'.lfOutPrecision = OUT_DEFAULT_PRECIS
'.lfPitchAndFamily = VARIABLE_PITCH Or FF_DONTCARE
'.lfQuality = DEFAULT_QUALITY
'.lfStrikeOut = Abs(FontInfo.Strikethrough)
'.lfUnderline = Abs(FontInfo.Underline)
'.lfWeight = IIf(FontInfo.Bold, FW_BOLD, FW_NORMAL)
'End With
'
''Create font
'hlFont = CreateFontIndirect(uLogFont)
'If hlFont = 0 Then Exit Function
'
''Select created font into dc to use it
'hlOld = SelectObject(lhDC, hlFont)
'
''Draw text and return result
'DrawRotatedText = (TextOut(lhDC, lX, lY, sText, Len(sText)) <> 0)
'
''Select old font back
'hlOld = SelectObject(lhDC, hlOld)
'
'DrawRotatedText_X:
'Exit Function
'
'DrawRotatedText_E:
'Resume DrawRotatedText_X
'
'End Function

Private Sub ActiveResize1_BeforeResize(Cancel As Boolean)

End Sub

Private Sub ActiveResize1_ResizeComplete()

    Call objMove(1)
'    Dim varBuffer() As Variant
'    Dim varBuf      As Variant
'    Dim utf8() As Byte
'    Dim ucs2 As Variant
'    Dim chars As Long
'    Dim varTmp As Variant
'    Dim i As Integer
'    Dim LineCount As Long
    Dim intRow As Integer
    Dim intCol      As Integer
    Dim strBuf()    As String
    
'    If gOpenFileNm <> "" Then
'        ' �÷��� �ʱ�ȭ
'        Set m_ColCommandButton = Nothing
'        Set m_ColCommandButton = New Collection
'
'
'        With spdList
'            sstType.Visible = False
'            For intRow = 1 To .Rows
'                '.Row = intRow
'                '.Col = 1
'                Erase strBuf
'                'If Trim(.Text) <> "" Then
'
'                If Trim(.CellValue(intRow, 1)) <> "" Then
'                    ReDim Preserve strBuf(.Cols) As String
'                    For intCol = 1 To .Cols
'                        '.Col = intCol
'                        'strBuf(intCol - 1) = Trim(.Text)
'                        strBuf(intCol - 1) = .CellValue(intRow, intCol)
'                    Next
'                    Call MakeLayout(strBuf)
'                    Erase strBuf
'                End If
'            Next
'            sstType.Visible = True
'        End With
'
'        Call PaintLine
'
'        vscPaint.Max = 100 * gDevide
'        hscPaint.Max = 100 * gDevide
'
'    End If
    
End Sub

Private Sub cboType_Click()
    
    sstType.Tab = cboType.ListIndex
    txtTitle.Enabled = True
    
    Select Case cboType.ListIndex
        Case 0
            txtTitle.Text = "S_TEXT" & gblCtrlIdx
        Case 1
            txtTitle.Text = "D_TEXT" & gblCtrlIdx
        Case 2
            txtTitle.Text = "S_Image" & gblCtrlIdx
        Case 3
            txtTitle.Text = "D_Image" & gblCtrlIdx
        Case 4
            txtTitle.Text = "BARCODE" & gblCtrlIdx
        Case 5
            txtTitle.Text = "LINE" & gblCtrlIdx
            txtLineHSize.Text = "1"
        Case 6
            txtTitle.Text = "RFID"
'            txtTitle.Enabled = False
    End Select
    
    txtXpos.Text = 1
    txtYpos.Text = 10
    
End Sub






'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' ���� ���� ��Ʈ�ѿ����� �̺�Ʈ ó��
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Sub ClsEventMonitor_EventRaised(EventObject As ClassEventObject, ByVal StrEventName As String)

    Dim StrEvent        As String
    Dim obj             As Object
    Dim val1            As Variant
    
    On Error Resume Next

    ' ���� �̺�Ʈ�� �߻��� Object
    Set obj = EventObject.EventObject

    StrEvent = ""
    StrEvent = StrEvent & Format(Now, "HH:MM:SS") & " "
    StrEvent = StrEvent & obj.Name & " - " & StrEventName & "("
    
    ' �Ķ���� ����
    For Each val1 In EventObject.Params
        StrEvent = StrEvent & CStr(val1) & ", "
    Next

    If Right(StrEvent, 2) = ", " Then
        StrEvent = Left(StrEvent, Len(StrEvent) - 2)
    End If

    StrEvent = StrEvent & "" & ")"
    
    ' �̺�Ʈ �α�
    List1.AddItem StrEvent, 0

End Sub

Private Sub cmdDelobj_Click()
    Dim intRow          As Integer
    Dim strObjType      As Variant
    Dim strObjName      As Variant
    Dim strObjRotate    As Variant
    
    If txtTag.Text = "" Then
        Exit Sub
    End If
    
    If InStr(txtTag.Text, "LineH_") > 0 Then
        Exit Sub
    End If
    
    Me.Controls(txtTag.Text).Visible = False
    
    With spdList
        For intRow = 1 To .Rows
            .Row = intRow
            
            strObjType = .CellValue(intRow, 2)
            strObjName = .CellValue(intRow, 29)
            
            If strObjType = sstType.Tab And strObjName = Trim(txtTag.Text) Then
                
                Call .RemoveRow(intRow)
                Exit For
            End If
        Next
    End With

End Sub

Private Sub cmdDevide_Click(Index As Integer)
    Dim intRow      As Integer
    Dim intCol      As Integer
    Dim strBuf()    As String
    
    Me.MousePointer = 11
    
    intMode = 2
    
    If Index = 0 Then
        If txtDevide.Text = "1" Then
            txtDevide.Text = "1"
            gDevide = "1"
        Else
            txtDevide.Text = txtDevide.Text - 1
            gDevide = txtDevide.Text - 0.01
        End If
    Else
        If txtDevide.Text = "5" Then
            txtDevide.Text = "5"
            gDevide = "5.01"
        
        Else
            txtDevide.Text = txtDevide.Text + 1
            gDevide = txtDevide.Text + 0.01
        End If
    End If
    
'    gDevide = txtDevide.Text
    
    ' �÷��� �ʱ�ȭ
    Set m_ColCommandButton = Nothing
    Set m_ColCommandButton = New Collection
    
    
    With spdList
        sstType.Visible = False
        For intRow = 1 To .Rows
            '.Row = intRow
            '.Col = 1
            Erase strBuf
            'If Trim(.Text) <> "" Then
                
            If Trim(.CellValue(intRow, 1)) <> "" Then
                ReDim Preserve strBuf(.Cols) As String
                For intCol = 1 To .Cols
                    '.Col = intCol
                    'strBuf(intCol - 1) = Trim(.Text)
                    strBuf(intCol - 1) = .CellValue(intRow, intCol)
                Next
                Call MakeLayout(strBuf)
                Erase strBuf
            End If
        Next
        sstType.Visible = True
    End With
    
    Call PaintLine
    
'    vscPaint.Max = 50 * gDevide
'    hscPaint.Max = 50 * gDevide
    vscPaint.Max = vscPaint.Max * gDevide
    hscPaint.Max = hscPaint.Max * gDevide

    Me.MousePointer = 0
    
End Sub

'-- ��Ʈ ����
Private Sub cmdFont_Click(Index As Integer)
 
    'Cancel�� True�� �����մϴ�.
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    
    'Flags �Ӽ��� �����մϴ�.
    CommonDialog1.flags = cdlCFEffects Or cdlCFBoth
    
    '��Ʈ �Ӽ��� �����մϴ�.[Default]
    If txtFontName(Index).Text = "" Then
        CommonDialog1.FontName = "����"
    Else
        CommonDialog1.FontName = txtFontName(Index).Text
    End If
    
    If txtFontSize(Index).Text = "" Then
        CommonDialog1.FontSize = 9
    Else
        CommonDialog1.FontSize = txtFontSize(Index).Text
    End If
    
    '[�۲�] ��ȭ ���ڸ� ǥ���մϴ�.
    CommonDialog1.ShowFont
    txtFontName(Index).Text = CommonDialog1.FontName
    txtFontSize(Index).Text = CommonDialog1.FontSize
    chkFontBold(Index).Value = IIf(CommonDialog1.FontBold = True, 1, 0)
    chkFontItalic(Index).Value = IIf(CommonDialog1.FontItalic = True, 1, 0)
    chkFontUnder(Index).Value = IIf(CommonDialog1.FontUnderline = True, 1, 0)

    Exit Sub

ErrHandler:
  '" ����ڰ� [���] ���߸� �������ϴ�.
  Exit Sub
  
End Sub

'-- �̹��� ��� ����
Private Sub cmdImage_Click(Index As Integer)
'
'    Dim sFile As String
'    sFile = ShowOpen("JPG����(*.jpg)|*.jpg", App.Path & "\" & gImage)
'    If sFile <> "" Then
'        txtImageName(Index).Text = sFile
'        If Index = 0 Then
'            Didim_SImg.Picture = LoadPicture(txtImageName(Index).Text)
'            txtImageWSize(Index).Text = Round(Didim_SImg.Width / gScaleCal, 0)
'            txtImageHSize(Index).Text = Round(Didim_SImg.Height / gScaleCal, 0)
'
'            txtImageWSize(Index + 2).Text = txtImageWSize(Index).Text
'            txtImageHSize(Index + 2).Text = txtImageHSize(Index).Text
'
'            txtImageDevide(Index).SetFocus
'        Else
'            Didim_DImg.Picture = LoadPicture(txtImageName(Index).Text)
'            txtImageWSize(Index).Text = Round(Didim_DImg.Width / gScaleCal, 0)
'            txtImageHSize(Index).Text = Round(Didim_DImg.Height / gScaleCal, 0)
'
'            txtImageWSize(Index + 2).Text = txtImageWSize(Index).Text
'            txtImageHSize(Index + 2).Text = txtImageHSize(Index).Text
'
'            txtImageDevide(Index).SetFocus
'        End If
'    Else
''        MsgBox "You pressed cancel"
'    End If




'
'
Dim x
    'Cancel�� True�� �����մϴ�.
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler

    'Flags �Ӽ��� �����մϴ�.
    CommonDialog1.flags = cdlCFEffects Or cdlCFBoth

    '��� �Ӽ��� �����մϴ�.
    CommonDialog1.InitDir = App.Path & "\" & gImage

    CommonDialog1.Filter = "JPG����(*.jpg)|*.jpg"

    '[����] ��ȭ ���ڸ� ǥ���մϴ�.
    CommonDialog1.ShowOpen
    txtImageName(Index).Text = CommonDialog1.FileName

    If Index = 0 Then
        Didim_SImg.Picture = LoadPicture(txtImageName(Index).Text)
        txtImageWSize(Index).Text = Round(Didim_SImg.Width / gScaleCal, 0)
        txtImageHSize(Index).Text = Round(Didim_SImg.Height / gScaleCal, 0)
    Else
        Didim_DImg.Picture = LoadPicture(txtImageName(Index).Text)
        txtImageWSize(Index).Text = Round(Didim_DImg.Width / gScaleCal, 0)
        txtImageHSize(Index).Text = Round(Didim_DImg.Height / gScaleCal, 0)
    End If

    Exit Sub

ErrHandler:
  '" ����ڰ� [���] ���߸� �������ϴ�.
  Exit Sub

End Sub

Private Sub MakeSpdSaveList(obj As Object, idx As Integer)
    
    With spdList
        .Rows = .Rows + 1
        '.Action = ActionActiveCell
        Select Case idx
        Case 0, 1, 6
            .CellValue(.Rows, 1) = .Rows - 1                                        '��������
            .CellValue(.Rows, 2) = idx                                              '�׸񱸺�
            
            If idx = 6 Then idx = 2
            
            .CellValue(.Rows, 3) = txtTitle.Text                                    '�׸��
            .CellValue(.Rows, 4) = txtXpos.Text                                     'X1��ǥ
            .CellValue(.Rows, 5) = 0                                                'X2��ǥ
            .CellValue(.Rows, 6) = txtYpos.Text                                     'Y1��ǥ
            .CellValue(.Rows, 7) = 0                                                'Y2��ǥ
            .CellValue(.Rows, 8) = txtFontName(idx).Text                            '��Ʈ��
            .CellValue(.Rows, 9) = txtFontSize(idx).Text                            '��Ʈũ��
            .CellValue(.Rows, 10) = IIf(chkFontBold(idx).Value = "0", "0", "1")     '��Ʈ����
            .CellValue(.Rows, 11) = IIf(chkFontUnder(idx).Value = "0", "0", "1")    '��Ʈ����
            .CellValue(.Rows, 12) = IIf(chkFontItalic(idx).Value = "0", "0", "1")   '��Ʈ����
            .CellValue(.Rows, 13) = "0"                                             '��Ʈȸ��
            .CellValue(.Rows, 14) = "0"                                             '���ڵ�����
            .CellValue(.Rows, 15) = "0"                                             '���ڵ���
            .CellValue(.Rows, 16) = "0"                                             '���ڵ�ȸ��
            .CellValue(.Rows, 17) = ""                                              '�̹������
            .CellValue(.Rows, 18) = "0"                                             '����ȸ��
            .CellValue(.Rows, 19) = "0"                                             '���εβ�
            .CellValue(.Rows, 20) = "0"                                             '������
            .CellValue(.Rows, 21) = IIf(chkPrint.Value = "1", "0", "1")             '��¿���
            .CellValue(.Rows, 22) = txtContent(idx).Text                            '��°�
            .CellValue(.Rows, 23) = gScaleCal                                             'X��ǥ ������
            .CellValue(.Rows, 24) = gScaleCal                                             'Y��ǥ ������
            .CellValue(.Rows, 25) = txtPaperHSize.Text                              '��������
            .CellValue(.Rows, 26) = txtPaperWSize.Text                              '������
            .CellValue(.Rows, 27) = IIf(chkFontItalic(idx).Value = "0", "0", "1")   '�����ǰ���
            .CellValue(.Rows, 28) = "0"                                             '��������
            .CellValue(.Rows, 29) = gblCtrlNm                                       'Tag

        Case 2
            .CellValue(.Rows, 1) = .Rows - 1                                     '��������
            .CellValue(.Rows, 2) = idx                                              '�׸񱸺�
            .CellValue(.Rows, 3) = txtTitle.Text                                    '�׸��
            .CellValue(.Rows, 4) = txtXpos.Text                                     'X1��ǥ
            .CellValue(.Rows, 5) = txtImageWSize(0).Text                            'X2��ǥ
            .CellValue(.Rows, 6) = txtYpos.Text                                     'Y1��ǥ
            .CellValue(.Rows, 7) = txtImageHSize(0).Text                            'Y2��ǥ
            .CellValue(.Rows, 8) = ""                            '��Ʈ��
            .CellValue(.Rows, 9) = "0"                            '��Ʈũ��
            .CellValue(.Rows, 10) = "0"     '��Ʈ����
            .CellValue(.Rows, 11) = "0"    '��Ʈ����
            .CellValue(.Rows, 12) = "0"    '��Ʈ����
            .CellValue(.Rows, 13) = "0"                                             '��Ʈȸ��
            .CellValue(.Rows, 14) = "0"                                             '���ڵ�����
            .CellValue(.Rows, 15) = "0"                                             '���ڵ���
            .CellValue(.Rows, 16) = "0"                                             '���ڵ�ȸ��
            .CellValue(.Rows, 17) = txtImageName(0).Text                                               '�̹������
            .CellValue(.Rows, 18) = "0"                                             '����ȸ��
            .CellValue(.Rows, 19) = "0"                                             '���εβ�
            .CellValue(.Rows, 20) = "0"                                             '������
            .CellValue(.Rows, 21) = IIf(chkPrint.Value = "1", "0", "1")             '��¿���
            .CellValue(.Rows, 22) = ""                            '��°�
            .CellValue(.Rows, 23) = gScaleCal                                             'X��ǥ ������
            .CellValue(.Rows, 24) = gScaleCal                                             'Y��ǥ ������
            .CellValue(.Rows, 25) = txtPaperHSize.Text                              '��������
            .CellValue(.Rows, 26) = txtPaperWSize.Text                              '������
            .CellValue(.Rows, 27) = IIf(chkIStatic.Value = "0", "0", "1")           '�����ǰ���
            .CellValue(.Rows, 28) = "0"                                             '��������
            .CellValue(.Rows, 29) = gblCtrlNm                                       'Tag
        

        Case 3
            .CellValue(.Rows, 1) = .Rows - 1                                     '��������
            .CellValue(.Rows, 2) = idx                                              '�׸񱸺�
            .CellValue(.Rows, 3) = txtTitle.Text                                    '�׸��
            .CellValue(.Rows, 4) = txtXpos.Text                                     'X1��ǥ
            .CellValue(.Rows, 5) = txtImageWSize(1).Text                            'X2��ǥ
            .CellValue(.Rows, 6) = txtYpos.Text                                     'Y1��ǥ
            .CellValue(.Rows, 7) = txtImageHSize(1).Text                            'Y2��ǥ
            .CellValue(.Rows, 8) = ""                            '��Ʈ��
            .CellValue(.Rows, 9) = "0"                            '��Ʈũ��
            .CellValue(.Rows, 10) = "0"     '��Ʈ����
            .CellValue(.Rows, 11) = "0"    '��Ʈ����
            .CellValue(.Rows, 12) = "0"    '��Ʈ����
            .CellValue(.Rows, 13) = "0"                                             '��Ʈȸ��
            .CellValue(.Rows, 14) = "0"                                             '���ڵ�����
            .CellValue(.Rows, 15) = "0"                                             '���ڵ���
            .CellValue(.Rows, 16) = "0"                                             '���ڵ�ȸ��
            .CellValue(.Rows, 17) = txtImageName(1).Text                                               '�̹������
            .CellValue(.Rows, 18) = "0"                                             '����ȸ��
            .CellValue(.Rows, 19) = "0"                                             '���εβ�
            .CellValue(.Rows, 20) = "0"                                             '������
            .CellValue(.Rows, 21) = IIf(chkPrint.Value = "1", "0", "1")             '��¿���
            .CellValue(.Rows, 22) = ""                            '��°�
            .CellValue(.Rows, 23) = gScaleCal                                             'X��ǥ ������
            .CellValue(.Rows, 24) = gScaleCal                                             'Y��ǥ ������
            .CellValue(.Rows, 25) = txtPaperHSize.Text                              '��������
            .CellValue(.Rows, 26) = txtPaperWSize.Text                              '������
            .CellValue(.Rows, 27) = IIf(chkIStatic.Value = "0", "0", "1")           '�����ǰ���
            .CellValue(.Rows, 28) = "0"                                             '��������
            .CellValue(.Rows, 29) = gblCtrlNm                                       'Tag
       
        Case 4
            .CellValue(.Rows, 1) = .Rows - 1                                     '��������
            .CellValue(.Rows, 2) = idx                                              '�׸񱸺�
            .CellValue(.Rows, 3) = txtTitle.Text                                    '�׸��
            .CellValue(.Rows, 4) = txtXpos.Text                                     'X1��ǥ
            .CellValue(.Rows, 5) = txtBarWSize.Text                                 'X2��ǥ
            .CellValue(.Rows, 6) = txtYpos.Text                                     'Y1��ǥ
            .CellValue(.Rows, 7) = txtBarHSize.Text                                 'Y2��ǥ
            .CellValue(.Rows, 8) = ""                                               '��Ʈ��
            .CellValue(.Rows, 9) = "0"                                              '��Ʈũ��
            .CellValue(.Rows, 10) = "0"                                             '��Ʈ����
            .CellValue(.Rows, 11) = "0"                                             '��Ʈ����
            .CellValue(.Rows, 12) = "0"                                             '��Ʈ����
            .CellValue(.Rows, 13) = "0"                                             '��Ʈȸ��
            .CellValue(.Rows, 14) = cboBarType.ListIndex                            '���ڵ�����
            .CellValue(.Rows, 15) = "0" 'txtBarDevide.Text                           '���ڵ���
            .CellValue(.Rows, 16) = IIf(chkBarRotate.Value = "0", 0, 2)             '���ڵ�ȸ��
            .CellValue(.Rows, 17) = ""                                              '�̹������
            .CellValue(.Rows, 18) = "0"                                             '����ȸ��
            .CellValue(.Rows, 19) = "0"                                             '���εβ�
            .CellValue(.Rows, 20) = "0"                                             '������
            .CellValue(.Rows, 21) = IIf(chkPrint.Value = "1", "0", "1")             '��¿���
            .CellValue(.Rows, 22) = Trim(txtBarData.Text)                           '��°�
            .CellValue(.Rows, 23) = gScaleCal                                       'X��ǥ ������
            .CellValue(.Rows, 24) = gScaleCal                                       'Y��ǥ ������
            .CellValue(.Rows, 25) = txtPaperHSize.Text                              '��������
            .CellValue(.Rows, 26) = txtPaperWSize.Text                              '������
            .CellValue(.Rows, 27) = IIf(chkIStatic.Value = "0", "0", "1")           '�����ǰ���
            .CellValue(.Rows, 28) = "0"                                             '��������
            .CellValue(.Rows, 29) = gblCtrlNm                                       'Tag
        
        Case 5
            .CellValue(.Rows, 1) = .Rows - 1                                     '��������
            .CellValue(.Rows, 2) = idx                                              '�׸񱸺�
            .CellValue(.Rows, 3) = txtTitle.Text                                    '�׸��
            If chkLineRotate.Value = "0" Then
                .CellValue(.Rows, 4) = txtXpos.Text                                 'X1��ǥ
                .CellValue(.Rows, 5) = txtLineWSize.Text                            'X2��ǥ
                .CellValue(.Rows, 6) = txtYpos.Text                                 'Y1��ǥ
                .CellValue(.Rows, 7) = txtYpos.Text                                 'Y2��ǥ
            Else
                .CellValue(.Rows, 4) = txtXpos.Text                                 'X1��ǥ
                .CellValue(.Rows, 5) = txtXpos.Text                                 'X2��ǥ
                .CellValue(.Rows, 6) = txtYpos.Text                                 'Y1��ǥ
                .CellValue(.Rows, 7) = txtLineWSize.Text                            'Y2��ǥ
            End If
            .CellValue(.Rows, 8) = ""                                               '��Ʈ��
            .CellValue(.Rows, 9) = "1"                                              '��Ʈũ��
            .CellValue(.Rows, 10) = "0"                                             '��Ʈ����
            .CellValue(.Rows, 11) = "0"                                             '��Ʈ����
            .CellValue(.Rows, 12) = "0"                                             '��Ʈ����
            .CellValue(.Rows, 13) = "0"                                             '��Ʈȸ��
            .CellValue(.Rows, 14) = "0"                                             '���ڵ�����
            .CellValue(.Rows, 15) = "0"                                             '���ڵ���
            .CellValue(.Rows, 16) = "0"                                             '���ڵ�ȸ��
            .CellValue(.Rows, 17) = ""                                              '�̹������
            .CellValue(.Rows, 18) = IIf(chkLineRotate.Value = "0", "0", "1")        '����ȸ��
            .CellValue(.Rows, 19) = txtLineHSize.Text                               '���εβ�
            .CellValue(.Rows, 20) = txtLineWSize.Text                               '������
            .CellValue(.Rows, 21) = IIf(chkPrint.Value = "1", "0", "1")             '��¿���
            .CellValue(.Rows, 22) = ""                                              '��°�
            .CellValue(.Rows, 23) = gScaleCal                                       'X��ǥ ������
            .CellValue(.Rows, 24) = gScaleCal                                       'Y��ǥ ������
            .CellValue(.Rows, 25) = txtPaperHSize.Text                              '��������
            .CellValue(.Rows, 26) = txtPaperWSize.Text                              '������
            .CellValue(.Rows, 27) = IIf(chkIStatic.Value = "0", "0", "1")           '�����ǰ���
            .CellValue(.Rows, 28) = "0"                                             '��������
            .CellValue(.Rows, 29) = gblCtrlNm                                       'Tag
        
        End Select
        
'        .ColWidth(-1) = 5
    End With
    
End Sub

' ������Ʈ�� ������Ų��.
Private Function objMake() As String
    Dim obj                 As Object
    Dim ClsEventObject      As ClassEventObject
    Dim intTab As Integer
    
    Set ClsEventObject = New ClassEventObject
    
    intTab = sstType.Tab
    objMake = "0"
    
    Select Case intTab
    Case 0  'Static Label
        Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectSLabel, txtTag.Text)
        If Not obj Is Nothing Then
            obj.Tag = txtTitle.Text
            obj.BackColor = vbWhite
            obj.Font = txtFontName(sstType.Tab).Text
            obj.Font.Charset = 163
            obj.Font.Size = Round(txtFontSize(sstType.Tab).Text * gDevide, 0)
            obj.Font.Bold = IIf(chkFontBold(sstType.Tab).Value = 1, True, False)
            obj.Font.Italic = IIf(chkFontItalic(sstType.Tab).Value = 1, True, False)
            obj.Font.Underline = IIf(chkFontUnder(sstType.Tab).Value = 1, True, False)
            obj.Top = Round(txtYpos.Text * gDevide, 0)
            obj.Left = Round(txtXpos.Text * gDevide, 0)
            obj.Caption = txtContent(sstType.Tab).Text
            obj.DataMember = chkTStatic.Value                       '-- �����ǰ���
            obj.DataField = IIf(chkPrint.Value = "1", "0", "1")     '-- ��¾���
            obj.AutoSize = True
            'obj.Width = Len(txtContent(sstType.Tab).Text) * 6
            'obj.Height = Round(txtFontSize(sstType.Tab).Text * gDevide, 0) * 1.6
            'obj.AutoSize = True
            obj.MousePointer = 5
            
        Else
            MsgBox "������ �׸���� ����� �� �����ϴ�.", vbInformation, Me.Caption
            Set ClsEventObject = Nothing
            Exit Function
'            Set ClsEventObject = Nothing
'            If MsgBox("������ �׸���� ����� �� �����ϴ�." & vbNewLine & "�����Ͻðڽ��ϱ�?", vbYesNo + vbCritical, Me.Caption) = vbYes Then
'                objMake = txtTag.Text & "_EDIT"
'                Exit Function
'            End If
        End If
    Case 1  'Dynamic Label
        Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectDLabel, txtTag.Text)
        If Not obj Is Nothing Then
            obj.Tag = txtTitle.Text
            obj.AutoSize = True
            obj.BackColor = vbWhite
            obj.Font = txtFontName(sstType.Tab).Text
            obj.Font.Size = Round(txtFontSize(sstType.Tab).Text * gDevide, 0)
            obj.Font.Bold = IIf(chkFontBold(sstType.Tab).Value = 1, True, False)
            obj.Font.Italic = IIf(chkFontItalic(sstType.Tab).Value = 1, True, False)
            obj.Font.Underline = IIf(chkFontUnder(sstType.Tab).Value = 1, True, False)
            obj.Top = Round(txtYpos.Text * gDevide, 0)
            obj.Left = Round(txtXpos.Text * gDevide, 0)
            obj.Caption = txtContent(sstType.Tab).Text
            obj.DataMember = IIf(chkPrint.Value = "1", "0", "1")   '-- ��¾���
            obj.MousePointer = 5
        Else
            MsgBox "������ �׸���� ����� �� �����ϴ�.", vbInformation, Me.Caption
            Set ClsEventObject = Nothing
            Exit Function
'            Set ClsEventObject = Nothing
'            If MsgBox(txtTag.Text & " �׸���� ����� �� �����ϴ�." & vbNewLine & "�׸���� �����Ͻðڽ��ϱ�?", vbYesNo + vbInformation, Me.Caption) = vbYes Then
'                objMake = txtTag.Text & "_EDIT"
'                Exit Function
'            End If
        End If
    Case 2 'Static Image
        Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectSImage, txtTag.Text)
        If Not obj Is Nothing Then
            If Dir(txtImageName(0).Text) = "" Then
                obj.Picture = LoadPicture(App.Path & "\" & gImage & "noimage.bmp")
            Else
                obj.Picture = LoadPicture(txtImageName(0).Text)
            End If
            obj.Tag = txtTitle.Text
            obj.DataMember = txtImageName(0).Text   '-- �̹������
            obj.Stretch = True
            obj.Width = Round(txtImageWSize(0).Text * gDevide, 0)
            obj.Height = Round(txtImageHSize(0).Text * gDevide, 0)
            obj.Top = Round(txtYpos.Text * gDevide, 0)
            obj.Left = Round(txtXpos.Text * gDevide, 0)
            obj.ToolTipText = chkIStatic.Value      '-- �����ǰ���
            obj.DataField = IIf(chkPrint.Value = "1", "0", "1")   '-- ��¾���
            obj.MousePointer = 5
        Else
            MsgBox "������ �׸���� ����� �� �����ϴ�.", vbInformation, Me.Caption
            Set ClsEventObject = Nothing
            Exit Function
'            Set ClsEventObject = Nothing
'            If MsgBox(txtTag.Text & " �׸���� ����� �� �����ϴ�." & vbNewLine & "�׸���� �����Ͻðڽ��ϱ�?", vbYesNo + vbInformation, Me.Caption) = vbYes Then
'                objMake = txtTag.Text & "_EDIT"
'                Exit Function
'            End If
        
        End If
    Case 3 'Dynamic Image
        Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectDImage, txtTag.Text)
        If Not obj Is Nothing Then
            If Dir(txtImageName(1).Text) = "" Then
                obj.Picture = LoadPicture(App.Path & "\" & gImage & "noimage.bmp")
            Else
                obj.Picture = LoadPicture(txtImageName(1).Text)
            End If
            obj.Tag = txtTitle.Text
            obj.DataMember = txtImageName(1).Text       '-- �̹������
            obj.Width = Round(txtImageWSize(1).Text * gDevide, 0)
            obj.Height = Round(txtImageHSize(1).Text * gDevide, 0)
            obj.Top = Round(txtYpos.Text * gDevide, 0)
            obj.Left = Round(txtXpos.Text * gDevide, 0)
            obj.DataField = IIf(chkPrint.Value = "1", "0", "1")   '-- ��¾���
            obj.MousePointer = 5
            obj.Stretch = True
        Else
            MsgBox "������ �׸���� ����� �� �����ϴ�.", vbInformation, Me.Caption
            Set ClsEventObject = Nothing
            Exit Function
'            Set ClsEventObject = Nothing
'            If MsgBox(txtTag.Text & " �׸���� ����� �� �����ϴ�." & vbNewLine & "�׸���� �����Ͻðڽ��ϱ�?", vbYesNo + vbInformation, Me.Caption) = vbYes Then
'                objMake = txtTag.Text & "_EDIT"
'                Exit Function
'            End If
        
        End If

    Case 4 'Barcode
        Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectBarcode, txtTag.Text)
        If Not obj Is Nothing Then
            obj.Tag = txtTitle.Text
            obj.Caption = txtBarData.Text
            obj.Style = cboBarType.ListIndex
            obj.Alignment = bcALeft
            obj.BarWidth = 0
            obj.Width = Round(txtBarWSize.Text * gDevide, 0)
            obj.Height = Round(txtBarHSize.Text * gDevide, 0)
            obj.Top = Round(txtYpos.Text * gDevide, 0)
            obj.Left = Round(txtXpos.Text * gDevide, 0)
            obj.Direction = IIf(chkBarRotate.Value = "0", 0, 2)
            obj.Visible = False
            'obj.Visible = True
        
            Set obj.Container = Picture1
            m_ColCommandButton.Add ClsEventObject
            Set ClsEventObject = Nothing
            
            Me.Controls(txtTag.Text).Top = Round(txtYpos.Text * gDevide, 0)
            Me.Controls(txtTag.Text).Height = Round(txtBarHSize.Text * gDevide, 0)
            
            '== ���ڵ带 �̹��� ���·� �ø��� ===================================================================
'            If intMode = 0 Then '==== Mode Set [0:�ε�,1:����,2:�̵�,3:����]
'                If strBarImgName = "" Then
'                    'strBarImgName = txtTag.Text & "_IMG1"
'                    strBarImgName = txtTag.Text & "_IMG"
'                Else
'                    strBarImgName = Mid(strBarImgName, 1, Len(strBarImgName) - 1) & Right(strBarImgName, 1) + 1
'                End If
'            Else
'                If strBarImgName = "" Then
'                    strBarImgName = gblCtrlNm & "_IMG"
'                End If
'            End If
            
            gblCtrlNm = gblCtrlNm & "_IMG"
            
            Set ClsEventObject = New ClassEventObject
            Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectBImage, gblCtrlNm)
            If Not obj Is Nothing Then
                obj.Tag = txtTitle.Text
                obj.Stretch = True
                If chkBarRotate.Value = "0" Then
                    obj.Picture = LoadPicture(App.Path & "\" & gImage & "\barcode.bmp")
                    obj.DataMember = App.Path & "\" & gImage & "\barcode.bmp"   '-- �̹��� ���
                    obj.Width = Round(txtBarWSize.Text * gDevide, 0)
                    obj.Height = Round(txtBarHSize.Text * gDevide, 0)
                Else
                    obj.Picture = LoadPicture(App.Path & "\" & gImage & "\barcode90.bmp")
                    obj.DataMember = App.Path & "\" & gImage & "\barcode90.bmp"   '-- �̹��� ���
                    obj.Width = Round(txtBarHSize.Text * gDevide, 0)
                    obj.Height = Round(txtBarWSize.Text * gDevide, 0)
                End If
                obj.Top = Round(txtYpos.Text * gDevide, 0)
                obj.Left = Round(txtXpos.Text * gDevide, 0)
                obj.ToolTipText = cboBarType.ListIndex                      '-- ���ڵ� Ÿ��
                obj.DataField = IIf(chkPrint.Value = "1", "0", "1")         '-- ��¾���
                obj.MousePointer = 5
            Else
                MsgBox "������ �׸���� ����� �� �����ϴ�.[���ڵ� ���� ����]", vbInformation, Me.Caption
                Set ClsEventObject = Nothing
                Exit Function
            End If
            '== ���ڵ带 �̹��� ���·� �ø��� ===================================================================
        Else
            MsgBox "������ �׸���� ����� �� �����ϴ�.", vbInformation, Me.Caption
            Set ClsEventObject = Nothing
            Exit Function
        End If
    Case 5  'Line
        Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectLImage, txtTag.Text)
        If Not obj Is Nothing Then
            obj.Tag = txtTitle.Text
            If chkLineRotate.Value = 0 Then
                obj.Picture = LoadPicture(App.Path & "\" & gImage & "wline.jpg")
                obj.Stretch = True
                obj.Width = Round(txtLineWSize * gDevide, 0)
                obj.Height = Round(txtLineHSize * gDevide, 0)
                obj.Top = Round(txtYpos.Text * gDevide, 0)
                obj.Left = Round(txtXpos.Text * gDevide, 0)
                obj.ToolTipText = IIf(chkPrint.Value = "1", "0", "1")   '-- ��¾���
                obj.DataMember = "0"                                    '-- Rotate
                obj.MousePointer = 5
            Else
                obj.Picture = LoadPicture(App.Path & "\" & gImage & "hline.jpg")
                obj.Stretch = True
                obj.Width = Round(txtLineHSize * gDevide, 0)
                obj.Height = Round(txtLineWSize * gDevide, 0)
                obj.Top = Round(txtYpos.Text * gDevide, 0)
                obj.Left = Round(txtXpos.Text * gDevide, 0)
                obj.ToolTipText = IIf(chkPrint.Value = "1", "0", "1")   '-- ��¾���
                obj.DataMember = "1"                                    '-- Rotate
                obj.MousePointer = 5
            End If
        Else
            MsgBox "������ �׸���� ����� �� �����ϴ�.", vbInformation, Me.Caption
            Set ClsEventObject = Nothing
            Exit Function
'            Set ClsEventObject = Nothing
'            If MsgBox(txtTag.Text & " �׸���� ����� �� �����ϴ�." & vbNewLine & "�׸���� �����Ͻðڽ��ϱ�?", vbYesNo + vbInformation, Me.Caption) = vbYes Then
'                objMake = txtTag.Text & "_EDIT"
'                Exit Function
'            End If
        End If
    Case 6  'Dynamic Label - RFID
        Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectRLabel, txtTag.Text)
        If Not obj Is Nothing Then
            obj.Tag = txtTitle.Text
            obj.AutoSize = True
            obj.BackColor = vbWhite
            obj.Font = txtFontName(2).Text
            obj.Font.Size = Round(txtFontSize(2).Text * gDevide, 0)
            obj.Font.Bold = IIf(chkFontBold(2).Value = 1, True, False)
            obj.Font.Italic = IIf(chkFontItalic(2).Value = 1, True, False)
            obj.Font.Underline = IIf(chkFontUnder(2).Value = 1, True, False)
            obj.Top = Round(txtYpos.Text * gDevide, 0)
            obj.Left = Round(txtXpos.Text * gDevide, 0)
            obj.Caption = txtContent(2).Text
            obj.DataMember = IIf(chkPrint.Value = "1", "0", "1")   '-- ��¾���
            obj.MousePointer = 5
        Else
            MsgBox "������ �׸���� ����� �� �����ϴ�.", vbInformation, Me.Caption
            Set ClsEventObject = Nothing
            Exit Function
        End If
    
    End Select
        
    Picture1.ScaleMode = vbPixels
    obj.Visible = True
    Set obj.Container = Picture1
    m_ColCommandButton.Add ClsEventObject
    Set ClsEventObject = Nothing
    
    Select Case intTab
    Case 0
        Me.Controls(txtTag.Text).Top = Round(txtYpos.Text * gDevide, 0)
        Me.Controls(txtTag.Text).Font.Size = Round(Me.Controls(txtTag.Text).Font.Size, 0)
    Case 1
        Me.Controls(txtTag.Text).Top = Round(txtYpos.Text * gDevide, 0)

    Case 2
        Me.Controls(txtTag.Text).Top = Round(txtYpos.Text * gDevide, 0)   '39.99912 ==> 34.06667
        Me.Controls(txtTag.Text).Height = Round(txtImageHSize(0).Text * gDevide, 0)    '61.33333
    
    Case 3
        Me.Controls(txtTag.Text).Top = Round(txtYpos.Text * gDevide, 0)   '39.99912 ==> 34.06667
        Me.Controls(txtTag.Text).Height = Round(txtImageHSize(1).Text * gDevide, 0)    '61.33333
    Case 4
        '-- ���ڵ�
        If chkBarRotate.Value = "0" Then
            obj.Height = Round(txtBarHSize.Text * gDevide, 0)
        Else
            obj.Height = Round(txtBarWSize.Text * gDevide, 0)
        End If
        obj.Top = Round(txtYpos.Text * gDevide, 0)
    Case 5
        If chkLineRotate.Value = 0 Then
            Me.Controls(txtTag.Text).Width = Round(txtLineWSize * gDevide, 0)
            Me.Controls(txtTag.Text).Top = Round(txtYpos.Text * gDevide, 0)
        Else
            Me.Controls(txtTag.Text).Width = Round(txtLineHSize * gDevide, 0)
            Me.Controls(txtTag.Text).Height = Round(txtLineWSize * gDevide, 0)
            Me.Controls(txtTag.Text).Top = Round(txtYpos.Text * gDevide, 0)
        End If
    Case 6
        Me.Controls(txtTag.Text).Top = Round(txtYpos.Text * gDevide, 0)

    
    End Select
    
End Function

Private Sub MakeBarImage(ByVal BarObj As Object)
    
    Picture2.Height = BarObj.Height
    Picture2.Width = BarObj.Width
    Barcod1.PrinterScaleMode = vbTwips 'Form1.ScaleMode
    Barcod1.PrinterWidth = BarObj.Width
    Barcod1.PrinterHeight = BarObj.Height
    Barcod1.PrinterTop = 0
    Barcod1.PrinterLeft = 0
    Barcod1.PrinterHDC = Picture2.hdc
    Picture2.Refresh
    Clipboard.Clear
    Clipboard.SetData Picture2.Image

'    SavePicture Picture2.Image, "C:\TEST.BMP"
    SavePicture Picture2.Image, "C:\TEST.BMP"

End Sub

Private Function findSameCtrlNm(strIdx As String, strTitle As String) As Boolean
    Dim i As Integer
    Dim strCtrlIdx  As String
    Dim strCtrlNm   As String
    
    findSameCtrlNm = False
    With spdList
        For i = 1 To .Rows
            '.Row = i
            '.Col = 2: strCtrlIdx = Trim(.Text)
            '.Col = 3: strCtrlNm = Trim(.Text)
            strCtrlIdx = .CellValue(i, 2)
            strCtrlNm = .CellValue(i, 3)
            
            If strIdx = strCtrlIdx And strTitle = strCtrlNm Then
                findSameCtrlNm = True
                Exit For
            End If
        Next
    End With
    
End Function

Private Sub objNewMake()
    Dim obj                 As Object
    Dim i                   As Integer
    Dim ClsEventObject      As ClassEventObject
    
    '-- ��ȿ�� �˻� [�׸��]
    If Trim(txtTitle.Text) = "" Then
        MsgBox "�׸���� �Է��ϼ���.", vbInformation, Me.Caption
        txtTitle.SetFocus
        Exit Sub
    End If
    '-- ��ȿ�� �˻� [X ��ǥ��]
    If Trim(txtXpos.Text) = "" Then
        MsgBox "X��ǥ�� �Է��ϼ���.", vbInformation, Me.Caption
        txtXpos.SetFocus
        Exit Sub
    End If
    '-- ��ȿ�� �˻� [X ��ǥ]
    If Not IsNumeric(Trim(txtXpos.Text)) Then
        MsgBox "X��ǥ�� ���ڸ� �Է��� �����մϴ�.", vbOKOnly + vbInformation, Me.Caption
        txtXpos.SetFocus
        Exit Sub
    End If
    '-- ��ȿ�� �˻� [Y ��ǥ��]
    If Trim(txtYpos.Text) = "" Then
        MsgBox "Y��ǥ�� �Է��ϼ���.", vbInformation, Me.Caption
        txtYpos.SetFocus
        Exit Sub
    End If
    '-- ��ȿ�� �˻� [Y ��ǥ]
    If Not IsNumeric(Trim(txtYpos.Text)) Then
        MsgBox "Y��ǥ�� ���ڸ� �Է��� �����մϴ�.", vbOKOnly + vbInformation, Me.Caption
        txtYpos.SetFocus
        Exit Sub
    End If
            
    Select Case sstType.Tab
        Case 0 '## Static Label ##
            '-- ��ȿ�� �˻� [��Ʈ��]
            If Trim(txtFontName(0).Text) = "" Or Trim(txtFontSize(0).Text) = "" Then
                MsgBox "Font�� �����ϼ���.", vbInformation, Me.Caption
                Call cmdFont_Click(0)
                Exit Sub
            End If
            '-- ��ȿ�� �˻� [��Ʈ������]
            If Not IsNumeric(Trim(txtFontSize(0).Text)) Then
                MsgBox "���ڸ� �Է��� �����մϴ�.", vbOKOnly + vbInformation, Me.Caption
                txtFontSize(0).SetFocus
                Exit Sub
            End If
            '-- ��ȿ�� �˻� [�ؽ�Ʈ]
            If Trim(txtContent(0).Text) = "" Then
                MsgBox "Text�� �Է��ϼ���.", vbInformation, Me.Caption
                txtContent(0).SetFocus
                Exit Sub
            End If
            
            '-- ���ϸ�Ī üũ
            If findSameCtrlNm(sstType.Tab, txtTitle.Text) Then
                MsgBox "������ �׸���� ����� �� �����ϴ�.", vbInformation, Me.Caption
                Exit Sub
            End If
            
            '-- Static Label ��ü�����
            If gblCtrlIdx = 0 And gblCtrlNm = "" Then
                gblCtrlIdx = 0
                gblCtrlNm = "Control_" & gblCtrlIdx
            Else
                gblCtrlIdx = gblCtrlIdx + 1
                gblCtrlNm = "Control_" & gblCtrlIdx
            End If
            
            Set ClsEventObject = New ClassEventObject
            'Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectSLabel, txtTitle.Text)
            Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectSLabel, gblCtrlNm)
            If Not obj Is Nothing Then
                obj.Tag = txtTitle.Text
                obj.AutoSize = True
                obj.BackColor = vbWhite
                obj.Font = txtFontName(sstType.Tab).Text
                obj.Font.Size = txtFontSize(sstType.Tab).Text * gDevide
                obj.Font.Bold = IIf(chkFontBold(sstType.Tab).Value = 1, True, False)
                obj.Font.Italic = IIf(chkFontItalic(sstType.Tab).Value = 1, True, False)
                obj.Font.Underline = IIf(chkFontUnder(sstType.Tab).Value = 1, True, False)
                obj.Top = txtYpos.Text * gDevide
                obj.Left = txtXpos.Text * gDevide
                obj.Caption = txtContent(sstType.Tab).Text
                obj.DataMember = chkTStatic.Value              '-- �����ǰ���
                obj.DataField = IIf(chkPrint.Value = "1", "0", "1")     '-- ��¾���
                obj.MousePointer = 5
                
                'obj======�׸��°�
                'X , Y====��ǥ
                'Txt======����
                'TxtGag===������ ����
                'H========������ ����(1�� ���� ����)
                'W========������ �ʺ�(1�� ���� ����)
                'LineSpace ====�ٰ���(1�� ���� ����)
                
'                Call RotateControl(obj, 90)
                
'                If optSTRotate(0).Value = True Then
'                    Call FontStuff(Picture1, obj.Top, obj.Left, obj.Caption, 0, 1, 1, 1)
'
'                ElseIf optSTRotate(1).Value = True Then
'                    Call FontStuff(Picture1, obj.Top, obj.Left, obj.Caption, 90, 1, 1, 1)
'                ElseIf optSTRotate(2).Value = True Then
'                    Call FontStuff(Picture1, obj.Top, obj.Left, obj.Caption, 180, 1, 1, 1)
'                Else
'                    Call FontStuff(Picture1, obj.Top, obj.Left, obj.Caption, 270, 1, 1, 1)
'                End If
        

                Call MakeSpdSaveList(obj, sstType.Tab)
            Else
                If gblCtrlIdx > 0 Then gblCtrlIdx = gblCtrlIdx - 1
                MsgBox "������ �׸���� ����� �� �����ϴ�.", vbInformation, Me.Caption
                Set ClsEventObject = Nothing
                Exit Sub
            End If
        
        Case 1  '## Dynamic Label ##
            '-- ��ȿ�� �˻� [��Ʈ��]
            If Trim(txtFontName(1).Text) = "" Or Trim(txtFontSize(1).Text) = "" Then
                MsgBox "Font�� �����ϼ���.", vbInformation, Me.Caption
                Call cmdFont_Click(1)
                Exit Sub
            End If
            '-- ��ȿ�� �˻� [��Ʈ������]
            If Not IsNumeric(Trim(txtFontSize(1).Text)) Then
                MsgBox "���ڸ� �Է��� �����մϴ�.", vbOKOnly + vbInformation, Me.Caption
                txtFontSize(1).SetFocus
                Exit Sub
            End If
            '-- ��ȿ�� �˻� [�ؽ�Ʈ]
            If Trim(txtContent(1).Text) = "" Then
                MsgBox "Text�� �Է��ϼ���.", vbInformation, Me.Caption
                txtContent(1).SetFocus
                Exit Sub
            End If
            
            '-- ���ϸ�Ī üũ
            If findSameCtrlNm(sstType.Tab, txtTitle.Text) Then
                MsgBox "������ �׸���� ����� �� �����ϴ�.", vbInformation, Me.Caption
                Exit Sub
            End If
            
            '-- Dynamic Label ��ü�����
            If gblCtrlIdx = 0 And gblCtrlNm = "" Then
                gblCtrlIdx = 0
                gblCtrlNm = "Control_" & gblCtrlIdx
            Else
                gblCtrlIdx = gblCtrlIdx + 1
                gblCtrlNm = "Control_" & gblCtrlIdx
            End If
            
            Set ClsEventObject = New ClassEventObject
            'Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectDLabel, txtTitle.Text)
            Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectDLabel, gblCtrlNm)
            If Not obj Is Nothing Then
                obj.Tag = txtTitle.Text
                obj.AutoSize = True
                obj.BackColor = vbWhite
                obj.Font = txtFontName(sstType.Tab).Text
                obj.Font.Size = txtFontSize(sstType.Tab).Text * gDevide
                obj.Font.Bold = IIf(chkFontBold(sstType.Tab).Value = 1, True, False)
                obj.Font.Italic = IIf(chkFontItalic(sstType.Tab).Value = 1, True, False)
                obj.Font.Underline = IIf(chkFontUnder(sstType.Tab).Value = 1, True, False)
                obj.Top = txtYpos.Text * gDevide
                obj.Left = txtXpos.Text * gDevide
                obj.Caption = txtContent(sstType.Tab).Text
                obj.DataMember = IIf(chkPrint.Value = "1", "0", "1")   '-- ��¾���
                obj.MousePointer = 5
                Call MakeSpdSaveList(obj, sstType.Tab)
            Else
                If gblCtrlIdx > 0 Then gblCtrlIdx = gblCtrlIdx - 1
                MsgBox "������ �׸���� ����� �� �����ϴ�.", vbInformation, Me.Caption
                Set ClsEventObject = Nothing
                Exit Sub
            End If
        
        Case 2 '## Static Image ##
            '-- ��ȿ�� �˻� [�̹�����]
            If Trim(txtImageName(0).Text) = "" Then
                MsgBox "�̹����� �����ϼ���.", vbInformation, Me.Caption
                Call cmdImage_Click(0)
                Exit Sub
            End If
            '-- ��ȿ�� �˻� [����Size]
            If Trim(txtImageWSize(0).Text) = "" Then
                MsgBox "����Size�� �Է��ϼ���.", vbInformation, Me.Caption
                txtImageWSize(0).SetFocus
                Exit Sub
            End If
            '-- ��ȿ�� �˻� [����Size]
            If Not IsNumeric(Trim(txtImageWSize(0).Text)) Then
                MsgBox "���ڸ� �Է��� �����մϴ�.", vbInformation, Me.Caption
                txtImageWSize(0).SetFocus
                Exit Sub
            End If
            '-- ��ȿ�� �˻� [����Size]
            If Trim(txtImageHSize(0).Text) = "" Then
                MsgBox "����Size�� �Է��ϼ���.", vbInformation, Me.Caption
                txtImageHSize(0).SetFocus
                Exit Sub
            End If
            '-- ��ȿ�� �˻� [����Size]
            If Not IsNumeric(Trim(txtImageHSize(0).Text)) Then
                MsgBox "���ڸ� �Է��� �����մϴ�.", vbInformation, Me.Caption
                txtImageHSize(0).SetFocus
                Exit Sub
            End If
            
            '-- ���ϸ�Ī üũ
            If findSameCtrlNm(sstType.Tab, txtTitle.Text) Then
                MsgBox "������ �׸���� ����� �� �����ϴ�.", vbInformation, Me.Caption
                Exit Sub
            End If
            
            '-- Static Image ��ü�����
            If gblCtrlIdx = 0 And gblCtrlNm = "" Then
                gblCtrlIdx = 0
                gblCtrlNm = "Control_" & gblCtrlIdx
            Else
                gblCtrlIdx = gblCtrlIdx + 1
                gblCtrlNm = "Control_" & gblCtrlIdx
            End If
            
            Set ClsEventObject = New ClassEventObject
            'Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectSImage, txtTitle.Text)
            Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectSImage, gblCtrlNm)
            If Not obj Is Nothing Then
                If Dir(txtImageName(0).Text) = "" Then
                    obj.Picture = LoadPicture(App.Path & "\image\noimage.bmp")
                Else
                    obj.Picture = LoadPicture(txtImageName(0).Text)
                End If
                obj.Tag = txtTitle.Text
                obj.DataMember = txtImageName(0).Text           '-- �̹��� ���
                obj.Stretch = True
                obj.Width = txtImageWSize(0).Text * gDevide
                obj.Height = txtImageHSize(0).Text * gDevide
                obj.Top = txtYpos.Text * gDevide
                obj.Left = txtXpos.Text * gDevide
                obj.MousePointer = 5
                obj.ToolTipText = chkIStatic.Value              '-- �����ǰ���
                obj.DataField = IIf(chkPrint.Value = "1", "0", "1")   '-- ��¾���
                Call MakeSpdSaveList(obj, sstType.Tab)
            Else
                If gblCtrlIdx > 0 Then gblCtrlIdx = gblCtrlIdx - 1
                MsgBox "������ �׸���� ����� �� �����ϴ�.", vbInformation, Me.Caption
                Set ClsEventObject = Nothing
                Exit Sub
            End If
            
        Case 3 '## Dynamic Image ##
            '-- ��ȿ�� �˻� [�̹�����]
            If Trim(txtImageName(1).Text) = "" Then
                MsgBox "�̹����� �����ϼ���.", vbInformation, Me.Caption
                Call cmdImage_Click(1)
                Exit Sub
            End If
            '-- ��ȿ�� �˻� [����Size]
            If Trim(txtImageWSize(1).Text) = "" Then
                MsgBox "����Size�� �Է��ϼ���.", vbInformation, Me.Caption
                txtImageWSize(1).SetFocus
                Exit Sub
            End If
            '-- ��ȿ�� �˻� [����Size]
            If Not IsNumeric(Trim(txtImageWSize(1).Text)) Then
                MsgBox "���ڸ� �Է��� �����մϴ�.", vbInformation, Me.Caption
                txtImageWSize(1).SetFocus
                Exit Sub
            End If
            '-- ��ȿ�� �˻� [����Size]
            If Trim(txtImageHSize(1).Text) = "" Then
                MsgBox "����Size�� �Է��ϼ���.", vbInformation, Me.Caption
                txtImageHSize(1).SetFocus
                Exit Sub
            End If
            '-- ��ȿ�� �˻� [����Size]
            If Not IsNumeric(Trim(txtImageHSize(1).Text)) Then
                MsgBox "���ڸ� �Է��� �����մϴ�.", vbInformation, Me.Caption
                txtImageHSize(1).SetFocus
                Exit Sub
            End If
            
            '-- ���ϸ�Ī üũ
            If findSameCtrlNm(sstType.Tab, txtTitle.Text) Then
                MsgBox "������ �׸���� ����� �� �����ϴ�.", vbInformation, Me.Caption
                Exit Sub
            End If
            
            '-- Dynamic Image ��ü�����
            If gblCtrlIdx = 0 And gblCtrlNm = "" Then
                gblCtrlIdx = 0
                gblCtrlNm = "Control_" & gblCtrlIdx
            Else
                gblCtrlIdx = gblCtrlIdx + 1
                gblCtrlNm = "Control_" & gblCtrlIdx
            End If
            
            Set ClsEventObject = New ClassEventObject
            'Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectDImage, txtTitle.Text)
            Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectDImage, gblCtrlNm)
            If Not obj Is Nothing Then
                If Dir(txtImageName(1).Text) = "" Then
                    obj.Picture = LoadPicture(App.Path & "\image\noimage.bmp")
                Else
                    obj.Picture = LoadPicture(txtImageName(1).Text)
                End If
                obj.Tag = txtTitle.Text
                obj.DataMember = txtImageName(1).Text           '-- �̹��� ���
                obj.Stretch = True
                obj.Width = txtImageWSize(1).Text * gDevide
                obj.Height = txtImageHSize(1).Text * gDevide
                obj.Top = txtYpos.Text * gDevide
                obj.Left = txtXpos.Text * gDevide
                obj.DataField = IIf(chkPrint.Value = "1", "0", "1")   '-- ��¾���
                obj.MousePointer = 5
                Call MakeSpdSaveList(obj, sstType.Tab)
            Else
                If gblCtrlIdx > 0 Then gblCtrlIdx = gblCtrlIdx - 1
                MsgBox "������ �׸���� ����� �� �����ϴ�.", vbInformation, Me.Caption
                Set ClsEventObject = Nothing
                Exit Sub
            End If
    
        Case 4  '## Barcode ##
            '-- ��ȿ�� �˻� [����Size]
            If Trim(txtBarWSize.Text) = "" Then
                MsgBox "����Size�� �Է��ϼ���.", vbInformation, Me.Caption
                txtBarWSize.SetFocus
                Exit Sub
            End If
            '-- ��ȿ�� �˻� [����Size]
            If Not IsNumeric(Trim(txtBarWSize.Text)) Then
                MsgBox "���ڸ� �Է��� �����մϴ�.", vbInformation, Me.Caption
                txtBarWSize.SetFocus
                Exit Sub
            End If
            '-- ��ȿ�� �˻� [����Size]
            If Trim(txtBarHSize.Text) = "" Then
                MsgBox "����Size�� �Է��ϼ���.", vbInformation, Me.Caption
                txtBarHSize.SetFocus
                Exit Sub
            End If
            '-- ��ȿ�� �˻� [����Size]
            If Not IsNumeric(Trim(txtBarHSize.Text)) Then
                MsgBox "���ڸ� �Է��� �����մϴ�.", vbInformation, Me.Caption
                txtBarHSize.SetFocus
                Exit Sub
            End If
            '-- ��ȿ�� �˻� [����Size]
            If Trim(txtBarData.Text) = "" Then
                MsgBox "Data�� �Է��ϼ���.", vbInformation, Me.Caption
                txtBarData.SetFocus
                Exit Sub
            End If
            
            '-- ���ϸ�Ī üũ
            If findSameCtrlNm(sstType.Tab, txtTitle.Text) Then
                MsgBox "������ �׸���� ����� �� �����ϴ�.", vbInformation, Me.Caption
                Exit Sub
            End If
            
            '-- Barcode ��ü�����
            If gblCtrlIdx = 0 And gblCtrlNm = "" Then
                gblCtrlIdx = 0
                gblCtrlNm = "Control_" & gblCtrlIdx
            Else
                gblCtrlIdx = gblCtrlIdx + 1
                gblCtrlNm = "Control_" & gblCtrlIdx
            End If
            
            Set ClsEventObject = New ClassEventObject
            'Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectBarcode, txtTitle.Text)
            Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectBarcode, gblCtrlNm)
            If Not obj Is Nothing Then
                obj.Tag = txtTitle.Text
                obj.Caption = txtBarData.Text
                obj.Style = cboBarType.ListIndex
                obj.Alignment = bcALeft
                obj.BarWidth = 0
                obj.Width = txtBarWSize.Text * gDevide
                obj.Height = txtBarHSize.Text * gDevide
                obj.Top = txtYpos.Text * gDevide
                obj.Left = txtXpos.Text * gDevide
                obj.Direction = IIf(chkBarRotate.Value = "0", 0, 2)
                'obj.DataField = IIf(chkPrint.Value = "1", "0", "1")         '-- ��¾���
                obj.Visible = False
                                
                Set obj.Container = Picture1
                m_ColCommandButton.Add ClsEventObject
                Set ClsEventObject = Nothing
                
'                If strBarImgName = "" Then
'                    'strBarImgName = txtTitle.Text & "_IMG"
'                    strBarImgName = gblCtrlNm & "_IMG"
'                Else
'                    strBarImgName = Mid(strBarImgName, 1, Len(strBarImgName) - 1) & Right(strBarImgName, 1) + 1
'                End If

                '-- ���ϸ�Ī üũ
                If findSameCtrlNm(sstType.Tab, txtTitle.Text) Then
                    MsgBox "������ �׸���� ����� �� �����ϴ�.", vbInformation, Me.Caption
                    Exit Sub
                End If

                gblCtrlNm = gblCtrlNm & "_IMG"
                Call MakeSpdSaveList(obj, sstType.Tab)
                                
                '== ���ڵ带 �̹��� ���·� �ø��� ===================================================================
                'gblCtrlNm = gblCtrlNm & "_IMG"
                
                Set ClsEventObject = New ClassEventObject
                'Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectBImage, strBarImgName)
                Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectBImage, gblCtrlNm)
                If Not obj Is Nothing Then
                    obj.Tag = txtTitle.Text
                    If chkBarRotate.Value = "0" Then
                        obj.Picture = LoadPicture(App.Path & "\" & gImage & "\barcode.bmp")
                        obj.DataMember = App.Path & "\" & gImage & "\barcode.bmp"
                        obj.Width = txtBarWSize.Text * gDevide
                        obj.Height = txtBarHSize.Text * gDevide
                    Else
                        obj.Picture = LoadPicture(App.Path & "\" & gImage & "\barcode90.bmp")
                        obj.DataMember = App.Path & "\" & gImage & "\barcode90.bmp"
                        obj.Width = txtBarHSize.Text * gDevide
                        obj.Height = txtBarWSize.Text * gDevide
                    End If
                    obj.Stretch = True
                    obj.Top = txtYpos.Text * gDevide
                    obj.Left = txtXpos.Text * gDevide
                    obj.ToolTipText = cboBarType.ListIndex
                    obj.DataField = IIf(chkPrint.Value = "1", "0", "1")         '-- ��¾���
                    obj.MousePointer = 5
                Else
                    If gblCtrlIdx > 0 Then gblCtrlIdx = gblCtrlIdx - 1
                    MsgBox "������ �׸���� ����� �� �����ϴ�.[���ڵ� ���� ����]", vbInformation, Me.Caption
                    Set ClsEventObject = Nothing
                    Exit Sub
                End If
                '== ���ڵ带 �̹��� ���·� �ø��� ===================================================================
            Else
                If gblCtrlIdx > 0 Then gblCtrlIdx = gblCtrlIdx - 1
                MsgBox "������ �׸���� ����� �� �����ϴ�.", vbInformation, Me.Caption
                Set ClsEventObject = Nothing
                Exit Sub
            End If
    
        Case 5  '## Line ##
            '-- ��ȿ�� �˻� [������]
            If Trim(txtLineHSize.Text) = "" Then
                MsgBox "�����⸦ �Է��ϼ���.", vbInformation, Me.Caption
                txtLineHSize.SetFocus
                Exit Sub
            End If
            '-- ��ȿ�� �˻� [������]
            If Not IsNumeric(Trim(txtLineHSize.Text)) Then
                MsgBox "���ڸ� �Է��� �����մϴ�.", vbInformation, Me.Caption
                txtLineHSize.SetFocus
                Exit Sub
            End If
            '-- ��ȿ�� �˻� [������]
            If Trim(txtLineWSize.Text) = "" Then
                MsgBox "�����̸� �Է��ϼ���.", vbInformation, Me.Caption
                txtLineWSize.SetFocus
                Exit Sub
            End If
            '-- ��ȿ�� �˻� [������]
            If Not IsNumeric(Trim(txtLineWSize.Text)) Then
                MsgBox "���ڸ� �Է��� �����մϴ�.", vbInformation, Me.Caption
                txtLineWSize.SetFocus
                Exit Sub
            End If
            
            '-- ���ϸ�Ī üũ
            If findSameCtrlNm(sstType.Tab, txtTitle.Text) Then
                MsgBox "������ �׸���� ����� �� �����ϴ�.", vbInformation, Me.Caption
                Exit Sub
            End If
            
            '-- Line ��ü�����
            If gblCtrlIdx = 0 And gblCtrlNm = "" Then
                gblCtrlIdx = 0
                gblCtrlNm = "Control_" & gblCtrlIdx
            Else
                gblCtrlIdx = gblCtrlIdx + 1
                gblCtrlNm = "Control_" & gblCtrlIdx
            End If
            
            Set ClsEventObject = New ClassEventObject
            Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectLImage, gblCtrlNm)
            If Not obj Is Nothing Then
                obj.Tag = txtTitle.Text
                If chkLineRotate.Value = 0 Then
                    obj.Picture = LoadPicture(App.Path & "\" & gImage & "wline.jpg")
                    obj.Stretch = True
                    obj.Width = txtLineWSize * gScaleCal
                    obj.Height = txtLineHSize * gScaleCal
                    obj.Top = txtYpos.Text * gScaleCal
                    obj.Left = txtXpos.Text * gScaleCal
                    obj.DataMember = "0"
                Else
                    obj.Picture = LoadPicture(App.Path & "\" & gImage & "hline.jpg")
                    obj.Stretch = True
                    obj.Width = txtLineHSize * gScaleCal
                    obj.Height = txtLineWSize * gScaleCal
                    obj.Top = txtYpos.Text * gScaleCal
                    obj.Left = txtXpos.Text * gScaleCal
                    obj.DataMember = "1"
                End If
                obj.MousePointer = 5
                Call MakeSpdSaveList(obj, sstType.Tab)
            Else
                If gblCtrlIdx > 0 Then gblCtrlIdx = gblCtrlIdx - 1
                MsgBox "������ �׸���� ����� �� �����ϴ�.", vbInformation, Me.Caption
                Set ClsEventObject = Nothing
                Exit Sub
            End If
        Case 6  '## Dynamic Label - RFID ##
            '-- ��ȿ�� �˻� [��Ʈ��]
            If Trim(txtFontName(2).Text) = "" Or Trim(txtFontSize(2).Text) = "" Then
                MsgBox "Font�� �����ϼ���.", vbInformation, Me.Caption
                Call cmdFont_Click(2)
                Exit Sub
            End If
            '-- ��ȿ�� �˻� [��Ʈ������]
            If Not IsNumeric(Trim(txtFontSize(2).Text)) Then
                MsgBox "���ڸ� �Է��� �����մϴ�.", vbOKOnly + vbInformation, Me.Caption
                txtFontSize(2).SetFocus
                Exit Sub
            End If
            '-- ��ȿ�� �˻� [�ؽ�Ʈ]
            If Trim(txtContent(2).Text) = "" Then
                MsgBox "Text�� �Է��ϼ���.", vbInformation, Me.Caption
                txtContent(2).SetFocus
                Exit Sub
            End If
            
            '-- ���ϸ�Ī üũ
            If findSameCtrlNm(sstType.Tab, txtTitle.Text) Then
                MsgBox "������ �׸���� ����� �� �����ϴ�.", vbInformation, Me.Caption
                Exit Sub
            End If
            
            '-- Dynamic Label ��ü�����
            If gblCtrlIdx = 0 And gblCtrlNm = "" Then
                gblCtrlIdx = 0
                gblCtrlNm = "Control_" & gblCtrlIdx
            Else
                gblCtrlIdx = gblCtrlIdx + 1
                gblCtrlNm = "Control_" & gblCtrlIdx
            End If
            
            Set ClsEventObject = New ClassEventObject
            'Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectDLabel, txtTitle.Text)
            Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectRLabel, gblCtrlNm)
            If Not obj Is Nothing Then
                obj.Tag = txtTitle.Text
                obj.AutoSize = True
                obj.BackColor = vbWhite
                obj.Font = txtFontName(2).Text
                obj.Font.Size = txtFontSize(2).Text * gDevide
                obj.Font.Bold = IIf(chkFontBold(2).Value = 1, True, False)
                obj.Font.Italic = IIf(chkFontItalic(2).Value = 1, True, False)
                obj.Font.Underline = IIf(chkFontUnder(2).Value = 1, True, False)
                obj.Top = txtYpos.Text * gDevide
                obj.Left = txtXpos.Text * gDevide
                obj.Caption = txtContent(2).Text
                obj.DataMember = IIf(chkPrint.Value = "1", "0", "1")   '-- ��¾���
                obj.MousePointer = 5
                Call MakeSpdSaveList(obj, sstType.Tab)
            Else
                If gblCtrlIdx > 0 Then gblCtrlIdx = gblCtrlIdx - 1
                MsgBox "������ �׸���� ����� �� �����ϴ�.", vbInformation, Me.Caption
                Set ClsEventObject = Nothing
                Exit Sub
            End If
    
    End Select
        
    
'    Dim lnghNewFont As Long
'    Dim lnghOriginalFonrt As Long
'    Dim lngHeight As Long
'    Dim lngWidth As Long
'    Dim intAngle As Integer
    
    
    obj.Visible = True
    Set obj.Container = Picture1
    
    m_ColCommandButton.Add ClsEventObject
    
    Set ClsEventObject = Nothing
    
'    intAngle = 90
'    With Picture1
'        .ScaleMode = vbPixels
'        .AutoRedraw = True
'        lngHeight = .TextHeight(obj)
'        lngWidth = 0
'
'        With .Font
'            lnghNewFont = CreateFont(lngHeight, lngWidth, intAngle * 10, intAngle * 10, .Weight, .Italic, .Underline, .Strikethrough, .Charset, 0, 0, 0, 0, .Name)
'        End With
'        lnghOriginalFonrt = SelectObject(.hdc, lnghNewFont)
'        .CurrentX = obj.Left
'        .CurrentY = obj.Top
'        Picture1.Print obj
'
'        lnghNewFont = SelectObject(.hdc, lnghOriginalFonrt)
'        .AutoRedraw = False
'    End With
'    DeleteObject lnghNewFont
'    'obj.Visible = False
        
    
End Sub

Private Sub objSet()
    Dim strNm As String

    If InStr(txtTag.Text, "LineH_") > 0 Then
        Exit Sub
    End If

    Select Case sstType.Tab
    Case 0  'Static Label
            Me.Controls(txtTag.Text).Tag = txtTitle.Text
            Me.Controls(txtTag.Text).AutoSize = True
            Me.Controls(txtTag.Text).BackColor = vbWhite
            Me.Controls(txtTag.Text).Font = txtFontName(sstType.Tab).Text
            Me.Controls(txtTag.Text).Font.Size = Round(txtFontSize(sstType.Tab).Text * gDevide, 0)
            Me.Controls(txtTag.Text).Font.Bold = IIf(chkFontBold(sstType.Tab).Value = 1, True, False)
            Me.Controls(txtTag.Text).Font.Italic = IIf(chkFontItalic(sstType.Tab).Value = 1, True, False)
            Me.Controls(txtTag.Text).Font.Underline = IIf(chkFontUnder(sstType.Tab).Value = 1, True, False)
            Me.Controls(txtTag.Text).Top = Round(txtYpos.Text * gDevide, 0)
            Me.Controls(txtTag.Text).Left = Round(txtXpos.Text * gDevide, 0)
            Me.Controls(txtTag.Text).Caption = txtContent(sstType.Tab).Text
            Me.Controls(txtTag.Text).DataMember = chkTStatic.Value
            Me.Controls(txtTag.Text).DataField = IIf(chkPrint.Value = "1", "0", "1")    '-- ��¾���
            
    Case 1  'Dynamic Label
            Me.Controls(txtTag.Text).Tag = txtTitle.Text
            Me.Controls(txtTag.Text).AutoSize = True
            Me.Controls(txtTag.Text).BackColor = vbWhite
            Me.Controls(txtTag.Text).Font = txtFontName(sstType.Tab).Text
            Me.Controls(txtTag.Text).Font.Size = Round(txtFontSize(sstType.Tab).Text * gDevide, 0)
            Me.Controls(txtTag.Text).Font.Bold = IIf(chkFontBold(sstType.Tab).Value = 1, True, False)
            Me.Controls(txtTag.Text).Font.Italic = IIf(chkFontItalic(sstType.Tab).Value = 1, True, False)
            Me.Controls(txtTag.Text).Font.Underline = IIf(chkFontUnder(sstType.Tab).Value = 1, True, False)
            Me.Controls(txtTag.Text).Top = Round(txtYpos.Text * gDevide, 0)
            Me.Controls(txtTag.Text).Left = Round(txtXpos.Text * gDevide, 0)
            Me.Controls(txtTag.Text).Caption = txtContent(sstType.Tab).Text
            Me.Controls(txtTag.Text).DataMember = IIf(chkPrint.Value = "1", "0", "1")          '-- ��¾���
    
    Case 2 'Static Image
            Me.Controls(txtTag.Text).Tag = txtTitle.Text
            Me.Controls(txtTag.Text).Width = Round(txtImageWSize(0).Text * gDevide, 0)
            Me.Controls(txtTag.Text).Height = Round(txtImageHSize(0).Text * gDevide, 0)
            Me.Controls(txtTag.Text).Top = Round(txtYpos.Text * gDevide, 0)
            Me.Controls(txtTag.Text).Left = Round(txtXpos.Text * gDevide, 0)
            If Dir(txtImageName(0).Text) = "" Then
                Me.Controls(txtTag.Text).Picture = LoadPicture(App.Path & "\" & gImage & "noimage.bmp")
            Else
                Me.Controls(txtTag.Text).Picture = LoadPicture(txtImageName(0).Text)
            End If
            
            Me.Controls(txtTag.Text).DataMember = txtImageName(0).Text   '-- �̹������
            
            Me.Controls(txtTag.Text).DataField = IIf(chkPrint.Value = "1", "0", "1")    '-- ��¾���
            
    Case 3 'Dynamic Image
            Me.Controls(txtTag.Text).Tag = txtTitle.Text
            Me.Controls(txtTag.Text).Width = Round(txtImageWSize(1).Text * gDevide, 0)
            Me.Controls(txtTag.Text).Height = Round(txtImageHSize(1).Text * gDevide, 0)
            Me.Controls(txtTag.Text).Top = Round(txtYpos.Text * gDevide, 0)
            Me.Controls(txtTag.Text).Left = Round(txtXpos.Text * gDevide, 0)
        
            If Dir(txtImageName(1).Text) = "" Then
                Me.Controls(txtTag.Text).Picture = LoadPicture(App.Path & "\" & gImage & "noimage.bmp")
            Else
                Me.Controls(txtTag.Text).Picture = LoadPicture(txtImageName(1).Text)
            End If

            Me.Controls(txtTag.Text).DataMember = txtImageName(1).Text   '-- �̹������
            Me.Controls(txtTag.Text).DataField = IIf(chkPrint.Value = "1", "0", "1")    '-- ��¾���
        
    Case 4  'Barcode Label
            '-- ���ڵ� �̹��� ����
            strNm = txtTag.Text
            Me.Controls(txtTag.Text).Tag = txtTitle.Text
            Me.Controls(strNm).Top = Round(txtYpos.Text * gDevide, 0)
            Me.Controls(strNm).Left = Round(txtXpos.Text * gDevide, 0)
            If chkBarRotate.Value = "0" Then
                Me.Controls(strNm).Width = Round(txtBarWSize.Text * gDevide, 0)
                Me.Controls(strNm).Height = Round(txtBarHSize.Text * gDevide, 0)
                Me.Controls(strNm).Picture = LoadPicture(App.Path & "\" & gImage & "barcode.bmp")
            Else
                Me.Controls(strNm).Height = Round(txtBarWSize.Text * gDevide, 0)
                Me.Controls(strNm).Width = Round(txtBarHSize.Text * gDevide, 0)
                Me.Controls(strNm).Picture = LoadPicture(App.Path & "\" & gImage & "barcode90.bmp")
            End If
            Me.Controls(strNm).ToolTipText = cboBarType.ListIndex           '-- ���ڵ� Ÿ��
            Me.Controls(strNm).DataField = IIf(chkPrint.Value = "1", "0", "1")    '-- ��¾���
            
            '-- ���ڵ� ����
            strNm = Mid(Trim(txtTag.Text), 1, InStr(Trim(txtTag.Text), "_IMG") - 1)
            Me.Controls(strNm).Tag = txtTitle.Text
            Me.Controls(strNm).Caption = txtBarData.Text
            Me.Controls(strNm).Style = cboBarType.ListIndex
            Me.Controls(strNm).Alignment = bcALeft
            Me.Controls(strNm).Top = Round(txtYpos.Text * gDevide, 0)
            Me.Controls(strNm).Left = Round(txtXpos.Text * gDevide, 0)
            If chkBarRotate.Value = "0" Then
                Me.Controls(strNm).Width = Round(txtBarWSize.Text * gDevide, 0)
                Me.Controls(strNm).Height = Round(txtBarHSize.Text * gDevide, 0)
            Else
                Me.Controls(strNm).Width = Round(txtBarHSize.Text * gDevide, 0)
                Me.Controls(strNm).Height = Round(txtBarWSize.Text * gDevide, 0)
            End If
            Me.Controls(strNm).Direction = IIf(chkBarRotate.Value = "0", 0, 2)
            
            
    Case 5  'Line Image
            Me.Controls(txtTag.Text).Tag = txtTitle.Text
            If chkLineRotate.Value = 0 Then
                Me.Controls(txtTag.Text).Width = Round(txtLineWSize * gDevide, 0)
                Me.Controls(txtTag.Text).Height = Round(txtLineHSize * gDevide, 0)
                Me.Controls(txtTag.Text).Top = Round(txtYpos.Text * gDevide, 0)
                Me.Controls(txtTag.Text).Left = Round(txtXpos.Text * gDevide, 0)
            Else
                Me.Controls(txtTag.Text).Width = Round(txtLineHSize * gDevide, 0)
                Me.Controls(txtTag.Text).Height = Round(txtLineWSize * gDevide, 0)
                Me.Controls(txtTag.Text).Top = Round(txtYpos.Text * gDevide, 0)
                Me.Controls(txtTag.Text).Left = Round(txtXpos.Text * gDevide, 0)
            End If
            Me.Controls(txtTag.Text).ToolTipText = IIf(chkPrint.Value = "1", "0", "1")   '-- ��¾���
            
    Case 6  'Dynamic Label - RFID
            Me.Controls(txtTag.Text).Tag = txtTitle.Text
            Me.Controls(txtTag.Text).AutoSize = True
            Me.Controls(txtTag.Text).BackColor = vbWhite
            Me.Controls(txtTag.Text).Font = txtFontName(2).Text
            Me.Controls(txtTag.Text).Font.Size = Round(txtFontSize(2).Text * gDevide, 0)
            Me.Controls(txtTag.Text).Font.Bold = IIf(chkFontBold(2).Value = 1, True, False)
            Me.Controls(txtTag.Text).Font.Italic = IIf(chkFontItalic(2).Value = 1, True, False)
            Me.Controls(txtTag.Text).Font.Underline = IIf(chkFontUnder(2).Value = 1, True, False)
            Me.Controls(txtTag.Text).Top = Round(txtYpos.Text * gDevide, 0)
            Me.Controls(txtTag.Text).Left = Round(txtXpos.Text * gDevide, 0)
            Me.Controls(txtTag.Text).Caption = txtContent(2).Text
            Me.Controls(txtTag.Text).DataMember = IIf(chkPrint.Value = "1", "0", "1")          '-- ��¾���
            
    End Select
    
'    Dim sText As String
'    sText = "Living on the edge..."
    
'    Call DrawRotatedText(picPrint.hdc, Me.Font, 900, sText, 0, Me.ScaleY(Me.TextWidth(sText), Me.ScaleMode, vbPixels))
    
    Call SetLayout(sstType.Tab)
        
End Sub



Private Sub cmdImageDevSet_Click(Index As Integer)
    
    If txtImageWSize(Index + 2).Text = "" Or txtImageHSize(Index + 2).Text = "" Then
        Exit Sub
    End If
    
    If Trim(txtImageDevide(Index).Text) = "" Or Not IsNumeric(txtImageDevide(Index).Text) Then
        MsgBox "�̹��� ������ Ȯ���ϼ���", vbOKOnly + vbInformation, Me.Caption
        txtImageDevide(Index).SetFocus
        Exit Sub
    End If
    
    If Trim(txtImageWSize(Index).Text) = "" Or Trim(txtImageHSize(Index).Text) = "" Or Not IsNumeric(txtImageWSize(Index).Text) And IsNumeric(txtImageHSize(Index).Text) Then
        MsgBox "�̹��� ����� Ȯ���ϼ���", vbOKOnly + vbInformation, Me.Caption
        Exit Sub
    Else
        txtImageWSize(Index).Text = Round(txtImageWSize(Index + 2).Text * (txtImageDevide(Index).Text / 100), 0)
        txtImageHSize(Index).Text = Round(txtImageHSize(Index + 2).Text * (txtImageDevide(Index).Text / 100), 0)
    End If
        
End Sub

' ���� ��Ʈ�� ����
Private Sub cmdMake_Click()
    
    '-- Mode Set [����]
    intMode = 3
    
    Call objNewMake
    
    Call PaintLine
            
End Sub


Private Sub objMove(Index)
    Dim intRow          As Integer
    Dim strObjType      As Variant
    Dim strObjName      As Variant
    Dim strObjRotate    As Variant
    
    With spdList
        Select Case Index
        Case 0      'left   - x1 ��ǥ
            For intRow = 1 To .Rows
                strObjType = .CellValue(intRow, 2)
                strObjName = .CellValue(intRow, 29)
                
                '-- �����̵�
                If chkChoice.Value = "1" Then
                    If strObjName = Trim(txtTag.Text) Then
                        If strObjType = 5 Then
                            If chkDetail.Value = 1 Then
                                .CellValue(intRow, 5) = .CellValue(intRow, 5) - 1
                                .CellValue(intRow, 4) = .CellValue(intRow, 4) - 1
                            Else
                                .CellValue(intRow, 5) = .CellValue(intRow, 5) - 5
                                .CellValue(intRow, 4) = .CellValue(intRow, 4) - 5
                            End If
                        Else
                            If chkDetail.Value = 1 Then
                                .CellValue(intRow, 4) = .CellValue(intRow, 4) - 1
                            Else
                                .CellValue(intRow, 4) = .CellValue(intRow, 4) - 5
                            End If
                        End If
                        '-- ����ȸ��[strObjRotate]�� "1" �̸� ��/�� �����̴�
                        '-- XI,X2�� ���� ������ �־�� �Ѵ�.
                        'Call .GetText(18, intRow, strObjRotate)
                        Me.Controls(strObjName).Left = .CellValue(intRow, 4) * gDevide
                    
                    End If
                Else
                    If strObjType = 5 Then
                        If chkDetail.Value = 1 Then
                            .CellValue(intRow, 5) = .CellValue(intRow, 5) - 1
                            .CellValue(intRow, 4) = .CellValue(intRow, 4) - 1
                        Else
                            .CellValue(intRow, 5) = .CellValue(intRow, 5) - 5
                            .CellValue(intRow, 4) = .CellValue(intRow, 4) - 5
                        End If
                    Else
                        If chkDetail.Value = 1 Then
                            .CellValue(intRow, 4) = .CellValue(intRow, 4) - 1
                        Else
                            .CellValue(intRow, 4) = .CellValue(intRow, 4) - 5
                        End If
                    End If
                    '-- ����ȸ��[strObjRotate]�� "1" �̸� ��/�� �����̴�
                    '-- XI,X2�� ���� ������ �־�� �Ѵ�.
                    'Call .GetText(18, intRow, strObjRotate)
                    Me.Controls(strObjName).Left = .CellValue(intRow, 4) * gDevide
                End If
            Next
        Case 1      'right  + x1 ��ǥ
            For intRow = 1 To .Rows
                strObjType = .CellValue(intRow, 2)
                strObjName = .CellValue(intRow, 29)
                
                '-- �����̵�
                If chkChoice.Value = "1" Then
                    If strObjName = Trim(txtTag.Text) Then
                        If strObjType = 5 Then
                            If chkDetail.Value = 1 Then
                                .CellValue(intRow, 5) = .CellValue(intRow, 5) + 1
                                .CellValue(intRow, 4) = .CellValue(intRow, 4) + 1
                            Else
                                .CellValue(intRow, 5) = .CellValue(intRow, 5) + 5
                                .CellValue(intRow, 4) = .CellValue(intRow, 4) + 5
                            End If
                        Else
                            If chkDetail.Value = 1 Then
                                .CellValue(intRow, 4) = .CellValue(intRow, 4) + 1
                            Else
                                .CellValue(intRow, 4) = .CellValue(intRow, 4) + 5
                            End If
                        End If
                        '-- ����ȸ��[strObjRotate]�� "1" �̸� ��/�� �����̴�
                        '-- XI,X2�� ���� ������ �־�� �Ѵ�.
                        'Call .GetText(18, intRow, strObjRotate)
                        Me.Controls(strObjName).Left = .CellValue(intRow, 4) * gDevide
                    
                    End If
                Else
                    If strObjType = 5 Then
                        If chkDetail.Value = 1 Then
                            .CellValue(intRow, 5) = .CellValue(intRow, 5) + 1
                            .CellValue(intRow, 4) = .CellValue(intRow, 4) + 1
                        Else
                            .CellValue(intRow, 5) = .CellValue(intRow, 5) + 5
                            .CellValue(intRow, 4) = .CellValue(intRow, 4) + 5
                        End If
                    Else
                        If chkDetail.Value = 1 Then
                            .CellValue(intRow, 4) = .CellValue(intRow, 4) + 1
                        Else
                            .CellValue(intRow, 4) = .CellValue(intRow, 4) + 5
                        End If
                    End If
                    '-- ����ȸ��[strObjRotate]�� "1" �̸� ��/�� �����̴�
                    '-- XI,X2�� ���� ������ �־�� �Ѵ�.
                    'Call .GetText(18, intRow, strObjRotate)
                    Me.Controls(strObjName).Left = .CellValue(intRow, 4) * gDevide
                End If
            Next

        Case 2      'top    - y1 ��ǥ
            For intRow = 1 To .Rows
                strObjType = .CellValue(intRow, 2)
                strObjName = .CellValue(intRow, 29)
                
                '-- �����̵�
                If chkChoice.Value = "1" Then
                    If strObjName = Trim(txtTag.Text) Then
                        If strObjType = 5 Then
                            If chkDetail.Value = 1 Then
                                .CellValue(intRow, 7) = .CellValue(intRow, 7) - 1
                                .CellValue(intRow, 6) = .CellValue(intRow, 6) - 1
                            Else
                                .CellValue(intRow, 7) = .CellValue(intRow, 7) - 5
                                .CellValue(intRow, 6) = .CellValue(intRow, 6) - 5
                            End If
                        Else
                            If chkDetail.Value = 1 Then
                                .CellValue(intRow, 6) = .CellValue(intRow, 6) - 1
                            Else
                                .CellValue(intRow, 6) = .CellValue(intRow, 6) - 5
                            End If
                        End If
                        '-- ����ȸ��[strObjRotate]�� "1" �̸� ��/�� �����̴�
                        '-- XI,X2�� ���� ������ �־�� �Ѵ�.
                        'Call .GetText(18, intRow, strObjRotate)
                        Me.Controls(strObjName).Top = .CellValue(intRow, 6) * gDevide
                    
                    End If
                Else
                    If strObjType = 5 Then
                        If chkDetail.Value = 1 Then
                            .CellValue(intRow, 7) = .CellValue(intRow, 7) - 1
                            .CellValue(intRow, 6) = .CellValue(intRow, 6) - 1
                        Else
                            .CellValue(intRow, 7) = .CellValue(intRow, 7) - 5
                            .CellValue(intRow, 6) = .CellValue(intRow, 6) - 5
                        End If
                    Else
                        If chkDetail.Value = 1 Then
                            .CellValue(intRow, 6) = .CellValue(intRow, 6) - 1
                        Else
                            .CellValue(intRow, 6) = .CellValue(intRow, 6) - 5
                        End If
                    End If
                    '-- ����ȸ��[strObjRotate]�� "1" �̸� ��/�� �����̴�
                    '-- XI,X2�� ���� ������ �־�� �Ѵ�.
                    'Call .GetText(18, intRow, strObjRotate)
                    Me.Controls(strObjName).Top = .CellValue(intRow, 6) * gDevide
                End If
            Next
        Case 3      'bottom + y1 ��ǥ
            For intRow = 1 To .Rows
                strObjType = .CellValue(intRow, 2)
                strObjName = .CellValue(intRow, 29)
                
                '-- �����̵�
                If chkChoice.Value = "1" Then
                    If strObjName = Trim(txtTag.Text) Then
                        If strObjType = 5 Then
                            If chkDetail.Value = 1 Then
                                .CellValue(intRow, 7) = .CellValue(intRow, 7) + 1
                                .CellValue(intRow, 6) = .CellValue(intRow, 6) + 1
                            Else
                                .CellValue(intRow, 7) = .CellValue(intRow, 7) + 5
                                .CellValue(intRow, 6) = .CellValue(intRow, 6) + 5
                            End If
                        Else
                            If chkDetail.Value = 1 Then
                                .CellValue(intRow, 6) = .CellValue(intRow, 6) + 1
                            Else
                                .CellValue(intRow, 6) = .CellValue(intRow, 6) + 5
                            End If
                        End If
                        '-- ����ȸ��[strObjRotate]�� "1" �̸� ��/�� �����̴�
                        '-- XI,X2�� ���� ������ �־�� �Ѵ�.
                        'Call .GetText(18, intRow, strObjRotate)
                        Me.Controls(strObjName).Top = .CellValue(intRow, 6) * gDevide
                    
                    End If
                Else
                    If strObjType = 5 Then
                        If chkDetail.Value = 1 Then
                            .CellValue(intRow, 7) = .CellValue(intRow, 7) + 1
                            .CellValue(intRow, 6) = .CellValue(intRow, 6) + 1
                        Else
                            .CellValue(intRow, 7) = .CellValue(intRow, 7) + 5
                            .CellValue(intRow, 6) = .CellValue(intRow, 6) + 5
                        End If
                    Else
                        If chkDetail.Value = 1 Then
                            .CellValue(intRow, 6) = .CellValue(intRow, 6) + 1
                        Else
                            .CellValue(intRow, 6) = .CellValue(intRow, 6) + 5
                        End If
                    End If
                    '-- ����ȸ��[strObjRotate]�� "1" �̸� ��/�� �����̴�
                    '-- XI,X2�� ���� ������ �־�� �Ѵ�.
                    'Call .GetText(18, intRow, strObjRotate)
                    Me.Controls(strObjName).Top = .CellValue(intRow, 6) * gDevide
                End If
            Next
        Case 4
            '-- X1,Y1 ��ǥ����
            For intRow = 1 To .Rows
                strObjType = .CellValue(intRow, 2)
                strObjName = .CellValue(intRow, 29)
                
                If strObjType = sstType.Tab And strObjName = Trim(txtTag.Text) Then
                    .CellValue(intRow, 4) = Trim(txtXpos.Text)
                    Me.Controls(strObjName).Left = Trim(txtXpos.Text) * gDevide
                    .CellValue(intRow, 6) = Trim(txtYpos.Text)
                    Me.Controls(strObjName).Top = Trim(txtYpos.Text) * gDevide
                    Exit For
                End If
            Next
        End Select
    End With

End Sub

Private Sub cmdMove_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    '-- Mode Set [�̵�]
    intMode = 2
    
    Me.MousePointer = 11
    
    Call objMove(Index)
    
    If Index < 4 Then
        intMoveIdx = Index
        
        If chkContinue.Value = 1 Then
            tmrMove.Interval = 100
            tmrMove.Enabled = True
            DoEvents
        Else
            tmrMove.Enabled = False
        End If
    End If
    
    Me.MousePointer = 0
    
End Sub

Private Sub cmdMove_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    tmrMove.Enabled = False

End Sub

Private Sub cmdPrint_Click()
    Dim prtSelectPrinter As Printer
    Dim boolPrinter_Select_Fales As Boolean
    Dim Buffer As String
    Dim aryPrinter() As String
    Dim strBuffer As String
 
'Printer ��ü�� �̿��� �μ⹰�� �ۼ��Ͻ� ������ ������ ������ ����Ͽ� �ֽʽÿ�.
'
'PaperSize �� Printer Driver�� ���� �ٸ����� �⺻������ A4 ������ �����Ǿ� �ֽ��ϴ�.
'������ ũ�⸦ ����� ������ ũ��� �����ϱ� ���Ͽ� ���� 256 ���� ������ �� ������
'������ ũ�⸦ �����ϴ� ���� ���ǹ��մϴ�. �Դٰ� 256���� ������ �� ������ �����ϴ�
'����̹��鵵 ���� �ֱ� �����Դϴ�.
'������ ũ�⸦ ������ �ʿ�� ������ �μ⹰�� ũ�⸸ �Ű澲�ø� �ǰڽ��ϴ�.
'
'�Բ��� ������� �ڵ带 ���ڸ� ���� 190, ���� 134 mm �� ������ �°� ����� �ϽǷ���
'�ϴ� �� �����ϴ�.
'�̷� ��� ������ ũ��� 190 * 134 ���� ������ �ʴٸ� � �����԰����� �����ص� ����
'�����ϴ�. �̷� ��쿡�� �׳� A4 �� �����ϼŵ� �˴ϴ�.
'Printer�� Width�Ӽ��� Height�Ӽ��� Twip ������ �Ǿ� ������ ���� �μⰡ���� �μ⹰��
'�׵θ�(�Ѱ�, Boundary)�������� �����Ͻô� �� ���� �� �մϴ�.

'�μ��� �� ���� �߿��� ���� ScaleMode, Scale, ScaleWidth, ScaleHeight �Դϴ�.
'
'mm ������������ �������� ����Ͻð��� �Ѵٸ� ScaleMode ��  6 ���� �����Ͻø� �˴ϴ�.
'������ ���� ������ A4��, ScaleMode�� 6 ���� ������ �Ŀ�
'Printer.Line (0, 0)-(210, 297), , B
'���� ���� �������� ��� ������ �ϴ��� �׵θ��� ������ �ʰ��Ͽ� ����� ���� �ʽ��ϴ�.
'�ֳ��ϸ� ������ ũ��� 210 * 297 ������ �����͸��� �μⰡ�ɿ����̶�� �� �����մϴ�.
'��ũ���� ��쿡�� ���������� ���� ���� ������ ���� �μⰡ�ɿ����� �۽��ϴ�.
'�׷��� ScaleMode �� 6���� ���� �� ScaleWidth �� ScaleHeight�� ���� ���� 210 �Ǵ� 297 ����
'���� ������ �Ǿ� �ִٴ� ���� �� �� �ֽ��ϴ�.
'�̷� �κе��� ����Ͽ� �μ⹰�� �ۼ��� ���ñ� �ٶ��ϴ�.
'�׷� ����~~�ϼ���.

 

''    '============== �̹��� ��� ��� ==========================================================
''    Picture1.AutoRedraw = True
''    SendMessage Picture1.hwnd, WM_PAINT, Picture1.hDC, 0
''    'SendMessage Picture1.hwnd, WM_PRINT, Picture1.hDC, PRF_CHILDREN Or PRF_CLIENT Or PRF_OWNED
''    Printer.PaintPicture Picture1.Image, 0, 0, Picture1.Width, Picture1.Height
''    Printer.EndDoc
''    SavePicture Picture1.Image, "C:\TEST.BMP"
    
''    '============== �̹��� ��� ��� ==========================================================
    
'Exit Sub

    Dim intRow As Integer
    Dim intCol As Integer
    Dim intCnt As Integer
    Dim strX1, strX2, strY1, strY2 As String
    Dim strFont As String
    Dim strFontSize As String
    Dim strFontBold As String
    Dim strFontUnder As String
    Dim strFontItalic As String
    Dim strdata As String
    Dim strTitle As String
    Dim strPrtYN    As String
    Dim intPixeltoTwip As Long
    Dim intPixeltoTwipX As Long
    Dim intPixeltoTwipY As Long
    Dim varTmp As Variant
    
    If chkCorrect.Value = "1" Then
'        Call spdList.GetText(23, 1, varTmp): intPixeltoTwip = IIf(varTmp <> "", varTmp, 15)
'        Call spdList.GetText(23, 1, varTmp): intPixeltoTwipX = IIf(varTmp <> "", varTmp, 15)
'        Call spdList.GetText(24, 1, varTmp): intPixeltoTwipX = IIf(varTmp <> "", varTmp, 15)
    
        intPixeltoTwip = gBojung '14.405
        intPixeltoTwipX = gBojung '14.405
        intPixeltoTwipY = gBojung '14.405
    Else
        intPixeltoTwip = 15
        intPixeltoTwipX = 15
        intPixeltoTwipY = 15
    End If
    
    '-- ���õ� �����ͷ� ���
    For Each prtSelectPrinter In Printers
        If UCase(Trim(prtSelectPrinter.DeviceName)) = UCase(Trim(cmbPrinter.Text)) Then
            Set Printer = prtSelectPrinter
            boolPrinter_Select_Fales = True
            Exit For
        End If
    Next
    
    With spdList
        Printer.ScaleMode = vbTwips
        Picture1.AutoRedraw = True
        '-- �ڽ� �׸���
        
        For intRow = 1 To .Rows
            '.Row = intRow
            '.Col = 2
            
            Select Case Trim(.CellValue(intRow, 2))
                Case "0"
                    Printer.ScaleMode = vbTwips
                    strX1 = Trim(.CellValue(intRow, 4)) * intPixeltoTwip
                    strX2 = Trim(.CellValue(intRow, 5)) * intPixeltoTwip
                    strY1 = Trim(.CellValue(intRow, 6)) * intPixeltoTwip
                    strY2 = Trim(.CellValue(intRow, 7)) * intPixeltoTwip
                    strFont = Trim(.CellValue(intRow, 8))
                    strFontSize = Trim(.CellValue(intRow, 9))
                    strFontBold = Trim(.CellValue(intRow, 10))
                    strFontItalic = Trim(.CellValue(intRow, 11))
                    strFontUnder = Trim(.CellValue(intRow, 12))
                    'strdata = Trim(.CellValue(intRow, 22))
                        
                    Printer.Font.Name = strFont
                    Printer.Font.Size = strFontSize
                    Printer.Font.Bold = IIf(strFontBold = "1", True, False)
                    Printer.Font.Italic = IIf(strFontItalic = "1", True, False)
                    Printer.Font.Underline = IIf(strFontUnder = "1", True, False)
                    Printer.CurrentX = strX1
                    Printer.CurrentY = strY1
                    'Printer.Print strData
    
                    'Call TextOutW(Picture1.hdc, strX1, strY1, StrPtr(strData), Len(strData))
'                    Call TextOutW(Printer.hdc, strX1, strY1, StrPtr(strData), Len(strData))
                    Call TextOutW(Printer.hdc, strX1 / 2.4, strY1 / 2.4, StrPtr(Trim(.CellValue(intRow, 22))), Len(Trim(.CellValue(intRow, 22))))
                    
                Case "1"
                    Printer.ScaleMode = vbTwips 'vbPixels 'vbTwips
                    strX1 = Trim(.CellValue(intRow, 4)) * intPixeltoTwip
                    strX2 = Trim(.CellValue(intRow, 5)) * intPixeltoTwip
                    strY1 = Trim(.CellValue(intRow, 6)) * intPixeltoTwip
                    strY2 = Trim(.CellValue(intRow, 7)) * intPixeltoTwip
                    strFont = Trim(.CellValue(intRow, 8))
                    strFontSize = Trim(.CellValue(intRow, 9))
                    strFontBold = Trim(.CellValue(intRow, 10))
                    strFontItalic = Trim(.CellValue(intRow, 11))
                    strFontUnder = Trim(.CellValue(intRow, 12))
                    strdata = Trim(.CellValue(intRow, 22))

'                    Printer.Font.Name = strFont
                    Printer.Font.Size = strFontSize
                    Printer.Font.Bold = IIf(strFontBold = "1", True, False)
                    Printer.Font.Italic = IIf(strFontItalic = "1", True, False)
                    Printer.Font.Underline = IIf(strFontUnder = "1", True, False)

                    Printer.CurrentX = strX1
                    Printer.CurrentY = strY1
                    'Printer.Print strData
                    
                    'Call TextOut(Printer.hdc, strX1, strY1, StrPtr(strData), Len(strData))
                    Call TextOutW(Printer.hdc, strX1 / 2.4, strY1 / 2.4, StrPtr(Trim(.CellValue(intRow, 22))), Len(Trim(.CellValue(intRow, 22))))
'Exit For
''
                Case "2"
                    Printer.ScaleMode = vbTwips
                    strTitle = Trim(.CellValue(intRow, 29))
                    strX1 = Trim(.CellValue(intRow, 4)) * intPixeltoTwip
                    strX2 = Trim(.CellValue(intRow, 5)) * intPixeltoTwip
                    strY1 = Trim(.CellValue(intRow, 6)) * intPixeltoTwip
                    strY2 = Trim(.CellValue(intRow, 7)) * intPixeltoTwip

                    Printer.PaintPicture Me.Controls(strTitle), strX1, strY1, strX2, strY2

                Case "3"
                    Printer.ScaleMode = vbTwips
                    strTitle = Trim(.CellValue(intRow, 29))
                    strX1 = Trim(.CellValue(intRow, 4)) * intPixeltoTwip
                    strX2 = Trim(.CellValue(intRow, 5)) * intPixeltoTwip
                    strY1 = Trim(.CellValue(intRow, 6)) * intPixeltoTwip
                    strY2 = Trim(.CellValue(intRow, 7)) * intPixeltoTwip

                    Printer.PaintPicture Me.Controls(strTitle), strX1, strY1, strX2, strY2

                Case "4"
                    strTitle = Trim(.CellValue(intRow, 29))
                    strTitle = Mid(Trim(strTitle), 1, InStr(Trim(strTitle), "_IMG") - 1)
                    strX1 = Trim(.CellValue(intRow, 4)) * intPixeltoTwip
                    strX2 = Trim(.CellValue(intRow, 5)) * intPixeltoTwip
                    strY1 = Trim(.CellValue(intRow, 6)) * intPixeltoTwip
                    strY2 = Trim(.CellValue(intRow, 7)) * intPixeltoTwip

                    Dim x, y, W, H

                    Printer.ScaleMode = vbTwips
                    Printer.PSet (0, 0), vbWhite

                    x = Printer.ScaleX(strX1, vbTwips) ' X-position = 25 mm from left border
                    y = Printer.ScaleY(strY1, vbTwips)  ' Y-position = 25 mm from top border
                    W = Printer.ScaleX(strX2, vbTwips)  ' Width = 100 mm
                    H = Printer.ScaleY(strY2, vbTwips)  ' Height = 40 mm

                    '-- ���ڵ� ȸ��
                    Me.Controls(strTitle).Direction = IIf(Trim(.CellValue(intRow, 16)) = "0", 0, 0)
                    
                    If Trim(.CellValue(intRow, 16)) = "0" Then
                        Me.Controls(strTitle).PrinterWidth = W
                        Me.Controls(strTitle).PrinterHeight = H
                    Else
                        Me.Controls(strTitle).PrinterWidth = H
                        Me.Controls(strTitle).PrinterHeight = W
                    End If
                    Me.Controls(strTitle).PrinterScaleMode = vbTwips   '3:�ȼ�,1:Ʈ��,6:�и�����
                    Me.Controls(strTitle).Alignment = bcACenter
                    Me.Controls(strTitle).PrinterLeft = x
                    Me.Controls(strTitle).PrinterTop = y
                    Me.Controls(strTitle).PrinterHDC = Printer.hdc

                Case "5"
                    '-- ��¿���
                    strPrtYN = Trim(.CellValue(intRow, 21))
                    Printer.ScaleMode = vbTwips

                    'If strPrtYN = "1" Then
                        Printer.PSet (0, 0), vbWhite
                        strX1 = Trim(.CellValue(intRow, 4)) * intPixeltoTwip '* 13.3
                        strX2 = Trim(.CellValue(intRow, 5)) * intPixeltoTwip '* 13.3
                        strY1 = Trim(.CellValue(intRow, 6)) * intPixeltoTwip '* 13.3
                        strY2 = Trim(.CellValue(intRow, 7)) * intPixeltoTwip '* 13.3
                        '������
                        Printer.DrawWidth = 1
                        Printer.Line (strX1, strY1)-(strX2, strY2)
                    'End If
            End Select
        Next
    End With
    

    Printer.EndDoc
    
    'SavePicture Picture1.Image, "C:\TEST.BMP"
    
End Sub

Public Sub cmdSet_Click()

    '-- Mode Set [���밡��]
    If intMode = 1 Then
        Call objSet
    End If
    
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' ���� ��ư ����
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'''Private Sub Command1_Click()
'''
'''    Dim obj                 As Object
'''    Dim i                   As Integer
'''    Dim ClsEventObject      As ClassEventObject
'''
'''    ' ���α׷� ���� TextBox ����
'''    Text1.Visible = False
'''
'''    List1.Clear
'''
'''    ' �÷��� �ʱ�ȭ
''''    Set m_ColCommandButton = Nothing
''''    Set m_ColCommandButton = New Collection
'''
'''    ' ���� ��Ʈ�� ����
'''    For i = 1 To Val(Combo1.Text)
'''        Set ClsEventObject = New ClassEventObject
'''
'''        If Option1.Value = True Then
'''            ' CommandButton
'''            Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectCommandButton, "DynamicCmd" & CStr(i))
'''            obj.Width = 3600
'''            obj.Height = 525
'''            obj.Top = 300 + (i - 1) * (525 + 30)
'''            obj.Left = 450
'''            obj.Caption = "Button" & CStr(i)
'''        ElseIf Option2.Value = True Then
'''            ' TextBox
'''            Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectTextBox, "DynamicTxt" & CStr(i))
'''            obj.Width = 3600
'''            obj.Height = 525
'''            obj.Top = 300 + (i - 1) * (525 + 30)
'''            obj.Left = 450
'''            obj.Text = "Text" & CStr(i)
'''        ElseIf Option3.Value = True Then
'''            ' Label
'''            Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectLabel, "DynamicLbl" & CStr(i))
'''            obj.Width = 3600
'''            obj.Height = 525
'''            obj.Top = 300 + (i - 1) * (525 + 30)
'''            obj.Left = 450
'''            obj.Caption = "Label" & CStr(i)
'''        ElseIf Option4.Value = True Then
'''            ' Image
'''            Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectImage, "DynamicImg" & CStr(i))
'''            obj.Width = 3600
'''            obj.Height = 525
'''            obj.Top = 300 + (i - 1) * (525 + 30)
'''            obj.Left = 450
'''            obj.Picture = LoadPicture(App.Path & "\ugc.jpg")
'''
'''        ElseIf Option5.Value = True Then
'''            ' line
'''            Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectLine, "DynamicLine" & CStr(i))
'''            '-- ���μ�
'''            obj.X1 = 100 * i
'''            obj.X2 = 100 * i
'''            obj.Y1 = 2070
'''            obj.Y2 = 4560
'''            '-- ���μ�
'''            obj.X1 = 2850
'''            obj.X2 = 7080
'''            obj.Y1 = 100 * i
'''            obj.Y2 = 100 * i
'''
'''        Else
'''            ' barcode
'''            Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectBarcode, "DynamicBar" & CStr(i))
'''            obj.Alignment = bcACenter
'''            obj.Caption = "88006611"
'''            obj.Style = msSCode128B
'''            obj.Width = 3600
'''            obj.Height = 525
'''            obj.Top = 300 + (i - 1) * (525 + 30)
'''            obj.Left = 450
'''
''''            Barcod1.Alignment = bcACenter
'''            'Barcod1.Style = msSCode128B ' msS2of5
'''
'''        End If
'''
'''        obj.Visible = True
'''        'Set obj.Container = Frame2
'''        Set obj.Container = Picture1
'''
'''        m_ColCommandButton.Add ClsEventObject
'''
'''        Set ClsEventObject = Nothing
'''    Next
'''
'''End Sub


Private Sub MDIForm_Tool()
    
On Error GoTo ErrorRouten
    
    With tlbMain
        .AllowCustomize = False
        Set .ImageList = imlToolbar
        '.TextAlignment = tbrTextAlignBottom '= tbrTextAlignRight
        .TextAlignment = tbrTextAlignRight
        .BorderStyle = ccNone
        .Appearance = cc3D
        .Style = tbrFlat
        Call .Buttons.Add(, TLBKEY_NEW, "�ű�", tbrDefault, "New")
        Call .Buttons.Add(, TLBKEY_OPEN, "����", tbrDefault, "Open")
        Call .Buttons.Add(, TLBKEY_SAVE, "����", tbrDefault, "Save")
        
        Call .Buttons.Add(, "", "", tbrSeparator)
        
        Call .Buttons.Add(, TLBKEY_MAKE, "JOB", tbrDefault, "Make")
        Call .Buttons.Add(, TLBKEY_VIEW, "����", tbrDefault, "View")
        Call .Buttons.Add(, "", "", tbrSeparator)
        Call .Buttons.Add(, TLBKEY_EDIT, "����", tbrDefault, "Edit")
        Call .Buttons.Add(, TLBKEY_EXIT, "����", tbrDefault, "Exit")
        Call .Buttons.Add(, "", "", tbrSeparator)
        
        
        .Refresh
    End With

Exit Sub

ErrorRouten:
'    Call ErrMsgProc(CallForm)

End Sub


'Private Sub Command2_Click()
'    Dim i As Integer
'    Dim sTmp As String
'    Text1.Text = "����(��)��"
'
'    Picture1.Cls
'    For i = 1 To Len(Text1.Text)
'        If Mid(Text1.Text, i, 1) = "(" Then
'            sTmp = Mid(Text1.Text, i, 3)
'            i = i + 2
'        Else
'            sTmp = Mid(Text1.Text, i, 1)
'        End If
'        Picture1.CurrentX = (Picture1.ScaleWidth - Picture1.TextWidth(sTmp)) / 2
'        Picture1.Print sTmp
'    Next i
'
'
'
'End Sub


Private Sub cmdUndo_Click()
    Dim Moveobj As Variant
    Dim x, y As Long
    
    If IsEmpty(LMousePos.obj) Then
        Exit Sub
    End If
    
    Moveobj = LMousePos.obj
    x = LMousePos.fromx
    y = LMousePos.fromy

    Me.Controls(Moveobj).Left = x
    Me.Controls(Moveobj).Top = y

End Sub



''Private Sub vscPaint_Change()
''  Dim lngPicPaintTop As Long
''
''  On Error GoTo ErrorHandler
''
''    If Abs(gblPicVval) < CLng(vscPaint.Value) * 10 Then
''        lngPicPaintTop = -(CLng(vscPaint.Value) * 10)
''        Picture1.Top = Picture1.Top + lngPicPaintTop
''        Picture1.Height = Picture1.Height + Abs(lngPicPaintTop)
''    Else
''        lngPicPaintTop = CLng(vscPaint.Value) * 10
''        If lngPicPaintTop = 0 Then
''            Picture1.Top = 0
''            Picture1.Height = Picture1.Height + Abs(lngPicPaintTop)
''        Else
''            Picture1.Top = Picture1.Top + lngPicPaintTop
''            Picture1.Height = Picture1.Height + Abs(lngPicPaintTop)
''        End If
''    End If
''
''    gblPicVval = lngPicPaintTop
''
''
''    Exit Sub
''
''ErrorHandler:
'''  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
''
''End Sub

Private Sub hscPaint_Change()
  Dim lngPicPaintLeft As Long
  
  On Error GoTo ErrorHandler
  
    If Abs(gblPicHval) < CLng(hscPaint.Value) * 10 Then
        lngPicPaintLeft = -(CLng(hscPaint.Value) * 10)
        Picture1.Left = Picture1.Left + lngPicPaintLeft
        Picture1.Width = Picture1.Width + Abs(lngPicPaintLeft)
    Else
        lngPicPaintLeft = CLng(hscPaint.Value) * 10
        If lngPicPaintLeft = 0 Then
            Picture1.Left = 0
            Picture1.Width = Picture1.Width + Abs(lngPicPaintLeft)
        Else
            Picture1.Left = Picture1.Left + lngPicPaintLeft
            Picture1.Width = Picture1.Width + Abs(lngPicPaintLeft)
        End If
    End If
  
    gblPicHval = lngPicPaintLeft
      
  
    Exit Sub

ErrorHandler:
'  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description

End Sub

'Private Sub hscPaint_Scroll()
'  Dim lngPicPaintLeft As Long
'
'  On Error GoTo ErrorHandler
'
'    If Abs(gblPicHval) < CLng(hscPaint.Value) * 10 Then
'        lngPicPaintLeft = -(CLng(hscPaint.Value) * 10)
'        Picture1.Left = Picture1.Left + lngPicPaintLeft
'        Picture1.Width = Picture1.Width + Abs(lngPicPaintLeft)
'    Else
'        lngPicPaintLeft = CLng(vscPaint.Value) * 10
'        If lngPicPaintLeft = 0 Then
'            Picture1.Left = 0
'            Picture1.Width = Picture1.Width + Abs(lngPicPaintLeft)
'        Else
'            Picture1.Left = Picture1.Left + lngPicPaintLeft
'            Picture1.Width = Picture1.Width + Abs(lngPicPaintLeft)
'        End If
'    End If
'
'    gblPicHval = lngPicPaintLeft
'
'
'    Exit Sub
'
'ErrorHandler:
''  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
'
'End Sub

Private Sub lblPrint_DblClick()
    
    If chkCorrect.Visible = True Then
        chkCorrect.Visible = False
    Else
        chkCorrect.Visible = True
    End If
    
End Sub

'Private Sub Command3_Click()
'
'    Call RotateControl(Me.Controls("Control_1"), 90)
'
'End Sub

'Private Sub Form_Activate()
'    MDIActiveX.WindowState = ccMaximize
'End Sub
'
'Private Sub Form_Deactivate()
'    MDIActiveX.WindowState = ccMinimize
'End Sub

Private Sub lblTitle_DblClick()
    
    If txtTag.Visible = True Then
        txtTag.Visible = False
    Else
        txtTag.Visible = True
    End If
    
End Sub

Private Sub mnuClose_Click()
        
    If MsgBox("�����Ͻðڽ��ϱ�?", vbYesNo + vbCritical, Me.Caption) = vbYes Then
        Unload Me
    End If
    
End Sub

Private Sub mnuMake_Click()
    
    If MsgBox("�۾������� �����Ͻðڽ��ϱ�?", vbYesNo + vbInformation, Me.Caption) = vbYes Then
        Call MakeJOB
    End If
    
End Sub


' ù��° ��� : UTF-16�� ��Ÿ���� Byte Order Mark(BOM) �� ���� ���,
'
Public Function UTF8FromUTF16(ByRef abytUTF16() As Byte) As Byte()
     
    Dim lngByteNum As Long
    Dim abytUTF8() As Byte
    Dim lngCharCount As Long
     
    On Error GoTo ConversionErr
     
    lngCharCount = (UBound(abytUTF16) + 1) \ 2
    ' UTF-16 LE ��Ʈ���� ������ ���� ���Խ���, ��ȯ�� �ʿ��� ����Ʈ ���� ���մϴ�.
    lngByteNum = WideCharToMultiByteArray(CP_UTF8, 0, abytUTF16(0), lngCharCount, 0, 0, 0, 0)
                     
    If lngByteNum > 0 Then
        ' ��ȯ�� �ڵ带 ��ȯ���� �޸𸮸� Ȯ���� �� �Լ��� ȣ���մϴ�.
        ReDim abytUTF8(lngByteNum - 1)
        lngByteNum = WideCharToMultiByteArray(CP_UTF8, 0, abytUTF16(0), lngCharCount, _
                                         abytUTF8(0), lngByteNum, 0, 0)
        UTF8FromUTF16 = abytUTF8
    End If
    Exit Function
     
ConversionErr:
    MsgBox " Conversion failed "
    
End Function


' �ι�° ��� : BOM �� ������ ��, UTF-8 ������� ��ȯ�� ��,
'                    UTF-8 ����� ��Ÿ���� Signature �� �߰��Ͽ� ��ȯ
'
Public Function UTF8FromUTF16withMark(ByRef abytUTF16() As Byte) As Byte()
    Dim abytTemp() As Byte
    Dim abytUTF8() As Byte
    Dim lngByteNum As Long
    Dim lngCharCount As Long
    Dim lngUpper As Long
     
    On Error GoTo ConversionErr
                   
    abytTemp = abytUTF16
    lngUpper = UBound(abytTemp)
    If lngUpper > 1 Then
        ' UTF-16 LE �� ����Ʈ����ǥ���� ���� ��� �̸� �ϴ� �����մϴ�.
        ' &HFEFF �����ε�, LE������ ��ġ�Ǿ� ����ǹǷ�, &HFF �� ���� ��ġ��.
        If abytTemp(0) = &HFF And abytTemp(1) = &HFE Then
            Call CopyMemory(abytTemp(0), abytTemp(2), lngUpper - 1)
            ReDim Preserve abytTemp(lngUpper - 2)
            lngUpper = lngUpper - 2
        End If
    End If
    lngCharCount = (lngUpper + 1) \ 2

   ' ���� ��ȯ�� �ʿ��� �޸��� ũ�⸦ ���մϴ�.
    lngByteNum = WideCharToMultiByteArray(CP_UTF8, 0, abytTemp(0), lngCharCount, 0, 0, 0, 0)
                     
    If lngByteNum > 0 Then
        ReDim abytUTF8(lngByteNum - 1)
        lngByteNum = WideCharToMultiByteArray(CP_UTF8, 0, abytTemp(0), lngCharCount, _
                                         abytUTF8(0), lngByteNum, 0, 0)
        lngUpper = UBound(abytUTF8)
        ' ��ȯ�Ǿ� �ִ� UTF-8 ����Ʈ �迭 ���ο� UTF-8 ǥ���� �ֱ� ����
        ' ������ ����Ʈ �迭�� �ڷ� �о��, �迭 �պκп� ǥ���� �߰��մϴ�.
        ReDim Preserve abytUTF8(lngUpper + 3)
        Call CopyMemory(abytUTF8(3), abytUTF8(0), lngUpper + 1)
        abytUTF8(0) = &HEF
        abytUTF8(1) = &HBB
        abytUTF8(2) = &HBF
         
        UTF8FromUTF16withMark = abytUTF8
    End If
    Exit Function
     
ConversionErr:
    MsgBox " Conversion failed "
    
End Function

Private Sub MakeLOF()
    Dim intRow As Integer
    Dim intCol As Integer
    Dim strdata As Variant
    Dim varTmp
    Dim abytUTF16() As Byte
    Dim abytUTF8() As Byte
    Dim sFile As String
    Dim FileName As String
    Dim strMapBar As String
    Dim blnMap      As Boolean
    
    'Cancel�� True�� �����մϴ�.
    'CommonDialog1.CancelError = True
    
    On Error GoTo ErrHandler
    
    'Flags �Ӽ��� �����մϴ�.
    'CommonDialog1.flags = cdlCFEffects Or cdlCFBoth
    
    '[�۲�] ��ȭ ���ڸ� ǥ���մϴ�.
    'CommonDialog1.ShowSave

    FileName = ShowSave("LayoutFile(*.lof)|*.lof", App.Path & "\" & gLayOut)
    If FileName <> "" Then
        '-- filename ���� Ư�����ڰ� �ٴ´�.
        FileName = Mid(FileName, 1, Len(FileName) - 1)
        If Not LCase(Right(FileName, 4)) = ".lof" Then
            FileName = FileName & ".lof"
        End If
        
        Open FileName For Binary As #1
        With spdList
            strdata = ""
            blnMap = False
            For intRow = 1 To .Rows
                For intCol = 1 To .Cols - 1 '-- ������ Control����
                    '.GetText intCol, intRow, varTmp: strData = strData & varTmp & "^"
                    varTmp = .CellValue(intRow, intCol)
                    
                    If intCol = 2 And varTmp = "4" Then
                        blnMap = True
                    End If
                    
                    If blnMap = True And intCol = 14 Then
                        strMapBar = BarIdxReMapper(varTmp)
                        strdata = strdata & varTmp & "^"
                        blnMap = False
                    Else
                        strdata = strdata & varTmp & "^"
                    End If
                Next
                strdata = strdata & vbCr
            Next
            
        End With
    
        abytUTF16 = strdata
        'abytUTF16 = "�����ڵ� ���ڵ� ��ȯ �׽�Ʈ : UTF-16 LE �� UTF-8 ������� ��ȯ�ϱ�"
        abytUTF8 = UTF8FromUTF16withMark(abytUTF16)
         
        'Open "C:\_UTF8TestFile.TXT" For Binary As #1
        Put #1, , abytUTF8
        Close #1
        'MsgBox " ��ȯ �Ϸ�. " & vbCrLf & " ���ͳ� �ͽ��÷η��� _UTF8TestFile.TXT ������ Ȯ���� �� �ֽ��ϴ�. "
    
    
        Close #1
    Else
    
    End If
    
    Exit Sub
    
ErrHandler:

End Sub

Private Sub MakeJOB()
    Dim intRow As Integer
    Dim intCol As Integer
    Dim strdata As Variant
    Dim varTmp
        
    On Error GoTo ErrHandler
    
    Open App.Path & "\" & gWork & "Job.txt" For Output As #1
        
    Print #1, "[JobPK]" & Chr(13) + Chr(10);
    Print #1, Me.Caption & ";" & Format(Now, "yyyy-mm-dd") & ";A;A;A;1;V" & Chr(13) + Chr(10);
    
    With spdList
        Print #1, "[S_Text]" & Chr(13) + Chr(10);
'        strData = ""
'        For intRow = 1 To .MaxRows
'            .GetText 2, intRow, varTmp
'            If varTmp = "0" Then
'                .GetText 3, intRow, varTmp
'                strData = strData & varTmp & ";"
'                .GetText 22, intRow, varTmp
'                strData = strData & varTmp
'                Print #1, strData & Chr(13) + Chr(10);
'                strData = ""
'            End If
'        Next
        
        '[D_Text]
        Print #1, "[D_Text]" & Chr(13) + Chr(10);
        strdata = ""
        For intRow = 1 To .Rows
            varTmp = .CellValue(intRow, 2)
            If varTmp = "1" Then
                strdata = strdata & .CellValue(intRow, 3) & ";"
                strdata = strdata & .CellValue(intRow, 22)
                Print #1, strdata & Chr(13) + Chr(10);
                strdata = ""
            End If
        Next
        
        '-- RFID
        strdata = ""
        For intRow = 1 To .Rows
            varTmp = .CellValue(intRow, 2)
            If varTmp = "6" Then
                strdata = strdata & .CellValue(intRow, 3) & ";"
                strdata = strdata & .CellValue(intRow, 22)
                Print #1, strdata & Chr(13) + Chr(10);
                strdata = ""
            End If
        Next
        
        '[S_Image]
        Print #1, "[S_Image]" & Chr(13) + Chr(10);
        strdata = ""
        For intRow = 1 To .Rows
            varTmp = .CellValue(intRow, 2)
            If varTmp = "2" Then
                strdata = strdata & .CellValue(intRow, 3) & ";"
                strdata = strdata & "0"
                Print #1, strdata & Chr(13) + Chr(10);
                strdata = ""
            End If
        Next
        
        '[D_Image]
        Print #1, "[D_Image]" & Chr(13) + Chr(10);
        strdata = ""
        For intRow = 1 To .Rows
            varTmp = .CellValue(intRow, 2)
            If varTmp = "3" Then
                strdata = strdata & .CellValue(intRow, 3) & ";"
                varTmp = .CellValue(intRow, 17)
                varTmp = Split(varTmp, "\")
                strdata = strdata & varTmp(UBound(varTmp))
                Print #1, strdata & Chr(13) + Chr(10);
                strdata = ""
            End If
        Next
        
        '[Barcode]
        Print #1, "[Barcode]" & Chr(13) + Chr(10);
        strdata = ""
        For intRow = 1 To .Rows
            varTmp = .CellValue(intRow, 2)
            If varTmp = "4" Then
                strdata = strdata & .CellValue(intRow, 22)
                Print #1, strdata & Chr(13) + Chr(10);
                strdata = ""
            End If
        Next
        
    End With
    
    Close #1
    
    MsgBox Me.Caption & "�� �۾������� �����Ǿ����ϴ�. ", vbOKOnly + vbInformation, Me.Caption

    Exit Sub
    
ErrHandler:

End Sub

Private Sub mnuNew_Click()
    
    Call FrmInitial
    
    Dim sNo1, sNo2 As String
    Dim intCnt As Integer
    Dim strEditObjName As String
    Dim strWLayout As String
    Dim strHLayout As String
    Dim intFVar As Integer
        
    intFVar = 0
    
AgainInput:
    
    sNo1 = Mid(gLayOutValue(gLayOutUse), 1, InStr(gLayOutValue(gLayOutUse), ":") - 1) / 10
    sNo2 = Mid(gLayOutValue(gLayOutUse), InStr(gLayOutValue(gLayOutUse), ":") + 1) / 10
    
'    sNo1 = InputBox("�󺧿��� ���̸� �Է��ϼ��� [���� : cm]", "���� �Է�", "7.5")
'
'    If Len(sNo1) > 0 Then
'        If Not IsNumeric(sNo1) Then
'            MsgBox "���ڸ� �Է��ϼ���.!", vbCritical
'            GoTo AgainInput
'        Else
'            sNo2 = InputBox("�󺧿��� ���̸� �Է��ϼ��� [���� : cm]", "���� �Է�", "3.5")
'            If Len(sNo2) > 0 Then
'                If Not IsNumeric(sNo2) Then
'                    MsgBox "���ڸ� �Է��ϼ���.!", vbCritical
'                    GoTo AgainInput
'                End If
'
'            End If
'        End If
'    End If

    
    If sNo1 <> "" And sNo2 <> "" Then
        txtPaperHSize.Text = sNo1 '/ 10
        txtPaperWSize.Text = sNo2 '/ 10
        
        sNo1 = Round(sNo1 * CM_TOTWIP, 0)
        sNo2 = Round(sNo2 * CM_TOTWIP, 0)
        
        sstType.Tab = 5
        '-- Left
        txtTitle.Text = "LINE_L"    '�׸��(���)
        txtTag.Text = "Control_" & intFVar       '�׸��(����)
        txtXpos.Text = "1"          'X ��ǥ
        txtYpos.Text = "1"          'Y ��ǥ
        txtLineHSize.Text = "1"     '������
        txtLineWSize.Text = sNo1   '������
        chkLineRotate.Value = "1"   '����ȸ��
        chkPrint.Value = "0"        '��¿���
    
        strEditObjName = objMake
        If strEditObjName = "0" Then
            '��ü���� ����
            '-- ��ü�̸� ������Ʈ
            gblCtrlNm = txtTag.Text
            gblCtrlIdx = intFVar
            intFVar = intFVar + 1
            Call MakeSpdSaveList(txtTitle, sstType.Tab)
        End If
        
        '-- Right
        txtTitle.Text = "LINE_R"    '�׸��(���)
        txtTag.Text = "Control_" & intFVar       '�׸��(����)
        txtXpos.Text = sNo2          'X ��ǥ
        txtYpos.Text = "1"          'Y ��ǥ
        txtLineHSize.Text = "1"     '������
        txtLineWSize.Text = sNo1   '������
        chkLineRotate.Value = "1"   '����ȸ��
        chkPrint.Value = "0"        '��¿���
    
        strEditObjName = objMake
        If strEditObjName = "0" Then
            '-- ��ü�̸� ������Ʈ
            gblCtrlNm = txtTag.Text
            gblCtrlIdx = intFVar
            intFVar = intFVar + 1
            '��ü���� ����
            Call MakeSpdSaveList(txtTitle, sstType.Tab)
        End If
    
        '-- Top
        txtTitle.Text = "LINE_T"    '�׸��(���)
        txtTag.Text = "Control_" & intFVar       '�׸��(����)
        txtXpos.Text = "1"          'X ��ǥ
        txtYpos.Text = "1"          'Y ��ǥ
        txtLineHSize.Text = "1"     '������
        txtLineWSize.Text = sNo2   '������
        chkLineRotate.Value = "0"   '����ȸ��
        chkPrint.Value = "0"        '��¿���
    
        strEditObjName = objMake
        If strEditObjName = "0" Then
            '-- ��ü�̸� ������Ʈ
            gblCtrlNm = txtTag.Text
            gblCtrlIdx = intFVar
            intFVar = intFVar + 1
            '��ü���� ����
            Call MakeSpdSaveList(txtTitle, sstType.Tab)
        End If
    
        '-- Bottom
        txtTitle.Text = "LINE_B"    '�׸��(���)
        txtTag.Text = "Control_" & intFVar       '�׸��(����)
        txtXpos.Text = "1"          'X ��ǥ
        txtYpos.Text = sNo1          'Y ��ǥ
        txtLineHSize.Text = "1"     '������
        txtLineWSize.Text = sNo2   '������
        chkLineRotate.Value = "0"   '����ȸ��
        chkPrint.Value = "0"        '��¿���
    
        strEditObjName = objMake
        If strEditObjName = "0" Then
            '-- ��ü�̸� ������Ʈ
            gblCtrlNm = txtTag.Text
            gblCtrlIdx = intFVar
            intFVar = intFVar + 1
            '��ü���� ����
            Call MakeSpdSaveList(txtTitle, sstType.Tab)
        End If
    
    End If
    
End Sub

Private Sub mnuSave_Click()
    Dim i As Integer

    Call MakeLOF
    
End Sub

Private Sub mnuSet_Click()

    frmConfig.Show

End Sub

Private Sub mnuView_Click()

    'If MsgBox("�۾������� �����Ͻðڽ��ϱ�?", vbYesNo + vbInformation, Me.Caption) = vbYes Then
        Call MakeJOB
        
        Call Shell(App.Path & "\" & "NOTEPAD.EXE", vbNormalFocus)
        
        Me.WindowState = 1
        
    'End If

End Sub


Private Sub optDevide_Click(Index As Integer)
    Dim intRow      As Integer
    Dim intCol      As Integer
    Dim strBuf()    As String
    
    gDevide = optDevide(Index).Tag
    
    ' �÷��� �ʱ�ȭ
    Set m_ColCommandButton = Nothing
    Set m_ColCommandButton = New Collection
    
    With spdList
        For intRow = 1 To .Rows
'            .Row = intRow
'            .Col = 1
            Erase strBuf
            If Trim(.CellValue(intRow, 1)) <> "" Then
                ReDim Preserve strBuf(.Cols) As String
                For intCol = 2 To .Cols
                    '.Col = intCol
                    strBuf(intCol - 1) = Trim(.CellValue(intRow, intCol))
                Next
                Call MakeLayout(strBuf)
                Erase strBuf
            End If
        Next
    End With
    
End Sub


Private Sub picDelobj_Click()
    
    Call cmdDelobj_Click

End Sub

Private Sub picFont_Click(Index As Integer)
 
    'Cancel�� True�� �����մϴ�.
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    
    'Flags �Ӽ��� �����մϴ�.
    CommonDialog1.flags = cdlCFEffects Or cdlCFBoth
    
    '��Ʈ �Ӽ��� �����մϴ�.[Default]
    CommonDialog1.FontName = "����"
    CommonDialog1.FontSize = 9
    
    '[�۲�] ��ȭ ���ڸ� ǥ���մϴ�.
    CommonDialog1.ShowFont
    txtFontName(Index).Text = CommonDialog1.FontName
    txtFontSize(Index).Text = CommonDialog1.FontSize
    chkFontBold(Index).Value = IIf(CommonDialog1.FontBold = True, 1, 0)
    chkFontItalic(Index).Value = IIf(CommonDialog1.FontItalic = True, 1, 0)
    chkFontUnder(Index).Value = IIf(CommonDialog1.FontUnderline = True, 1, 0)

    Exit Sub

ErrHandler:
  '" ����ڰ� [���] ���߸� �������ϴ�.
  Exit Sub

End Sub

Private Sub picImage_Click(Index As Integer)

    Dim sFile As String
    sFile = ShowOpen("JPG����(*.jpg)|*.jpg", App.Path & "\" & gImage)
    If sFile <> "" Then
        txtImageName(Index).Text = sFile
        If Index = 0 Then
            Didim_SImg.Picture = LoadPicture(txtImageName(Index).Text)
            txtImageWSize(Index).Text = Round(Didim_SImg.Width / gScaleCal, 0)
            txtImageHSize(Index).Text = Round(Didim_SImg.Height / gScaleCal, 0)
            
            txtImageWSize(Index + 2).Text = txtImageWSize(Index).Text
            txtImageHSize(Index + 2).Text = txtImageHSize(Index).Text
            
            txtImageDevide(Index).SetFocus
        Else
            Didim_DImg.Picture = LoadPicture(txtImageName(Index).Text)
            txtImageWSize(Index).Text = Round(Didim_DImg.Width / gScaleCal, 0)
            txtImageHSize(Index).Text = Round(Didim_DImg.Height / gScaleCal, 0)
        
            txtImageWSize(Index + 2).Text = txtImageWSize(Index).Text
            txtImageHSize(Index + 2).Text = txtImageHSize(Index).Text
            
            txtImageDevide(Index).SetFocus
        End If
    Else
'        MsgBox "You pressed cancel"
    End If

End Sub

Private Sub picMake_Click()
    
    '-- Mode Set [����]
    intMode = 3
    
    Call objNewMake
        
End Sub

Private Sub picPrint_Click()
    Call cmdPrint_Click
End Sub

Private Sub picSet_Click()

    '-- Mode Set [���밡��]
    If intMode = 1 Then
        Call objSet
    End If
    
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'
'    If Button = 1 Then
'        Picture1.Cls '=============>�ٽ� �׸���
''        Picture1.CurrentX = X
''        Picture1.CurrentY = Y
'        DrawX = X '=========>��������ǥ���
'        DrawY = Y
'
'        Picture1.DrawMode = 10
'
'        Ot_X = X
'        Ot_Y = Y
'    End If

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

'    If Button = 1 Then
'        Picture1.DrawWidth = 1
'        Picture1.DrawStyle = 2
'
'        Picture1.Line (DrawX, DrawY)-(Ot_X, Ot_Y), vbBlack, B
'        Picture1.Line (DrawX, DrawY)-(X, Y), vbBlack, B
'
'        Ot_X = X
'        Ot_Y = Y
'    End If
    
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

'    If Button = 1 Then
'        Picture1.Line (DrawX, DrawY)-(Ot_X, Ot_Y), vbBlue, B
'        Picture1.DrawMode = 13
'        Picture1.DrawWidth = 1
'        Picture1.DrawStyle = 0 '========>�ܻ�(���� ������ �״�� ����)
'        Picture1.Line (DrawX, DrawY)-(X, Y), vbBlue, B
'    End If

End Sub

'-- ��Ʈ�� �ʱ�ȭ
Private Sub CtrlInitial()
        
    txtPaperHSize.Text = ""
    txtPaperWSize.Text = ""
        
    '-- Tab 0
    txtFontName(0).Text = ""
    txtFontSize(0).Text = ""
    chkFontBold(0).Value = 0
    chkFontUnder(0).Value = 0
    chkFontItalic(0).Value = 0
    txtContent(0).Text = ""
    
    '-- Tab 1
    txtFontName(1).Text = ""
    txtFontSize(1).Text = ""
    chkFontBold(1).Value = 0
    chkFontUnder(1).Value = 0
    chkFontItalic(1).Value = 0
    txtContent(1).Text = ""
    
    '-- Tab 2
    txtImageName(0).Text = ""
    txtImageWSize(0).Text = ""
    txtImageHSize(0).Text = ""
    txtImageWSize(2).Text = ""
    txtImageHSize(2).Text = ""
    
    chkIStatic.Value = 0
    
    '-- Tab 3
    txtImageName(1).Text = ""
    txtImageWSize(1).Text = ""
    txtImageHSize(1).Text = ""
    txtImageWSize(3).Text = ""
    txtImageHSize(3).Text = ""
    
    '-- Tab 4
    txtBarDevide.Text = ""
    txtBarWSize.Text = ""
    txtBarHSize.Text = ""
    txtBarData.Text = ""
    chkBarRotate.Value = 0
    
    '-- Tab 5
    txtLineHSize.Text = ""
    txtLineWSize.Text = ""
    chkLineRotate.Value = 0
    
    gblCtrlNm = ""
    gblCtrlIdx = 0
    
    
End Sub

'-- ȭ�� �ʱ�ȭ
Private Sub FrmInitial()
    Dim x As Printer
    Dim prtSelectPrinter As Printer
    Dim boolPrinter_Select_Fales As Boolean
    Dim strDefault As String
    Dim Buffer As String
    Dim aryPrinter() As String
    Dim strBuffer As String
    Dim i As Integer
    Dim j As Integer
 
    ' Ŭ���� �ʱ�ȭ
    Set ClsEventMonitor = New ClassEventMonitor
    Set m_ColCommandButton = New Collection

    Call CtrlInitial
    
    '����
    cboType.Clear
    cboType.AddItem "S_Text"
    cboType.AddItem "D_Text"
    cboType.AddItem "S_Image"
    cboType.AddItem "D_Image"
    cboType.AddItem "Barcode"
    cboType.AddItem "Line"
    cboType.AddItem "RFID"
    
    cboType.ListIndex = 0
    
    '���ڵ� Ÿ��
    cboBarType.Clear
    cboBarType.AddItem "None"
    cboBarType.AddItem "2of5[����]"               '5
    cboBarType.AddItem "Interleaved2of5[����]"    '6
    cboBarType.AddItem "3of9[����]"               '0
    cboBarType.AddItem "Codabar[����]"            '9
    cboBarType.AddItem "3of9X[����]"              '1
    cboBarType.AddItem "Code128A[����]"           '11
    cboBarType.AddItem "Code128B[����]"           '12
    cboBarType.AddItem "Code128C[����]"           '13
    cboBarType.AddItem "UPCA[����]"               '15
    cboBarType.AddItem "MSI[����]"                '7
    cboBarType.AddItem "Code93[����]"             '3
    cboBarType.AddItem "ExtendedCode93[����]"     '4
    cboBarType.AddItem "EAN13[����]"              '17
    cboBarType.AddItem "EAN8[����]"               '18
    cboBarType.AddItem "PostNet[����]"            '23
    cboBarType.AddItem "ANSI3of9[�ű�]"           '
    cboBarType.AddItem "ANSI3of9X[�ű�]"          '
    cboBarType.AddItem "Code128Auto[����]"        '10
    cboBarType.AddItem "UCCEAN128[����]"          '27
    cboBarType.AddItem "UPCE[����]"               '16
    cboBarType.AddItem "RoyalMail[�ű�]"          '
    cboBarType.AddItem "MSICode2[����]"           '8  ??MSIPlessey
    cboBarType.AddItem "DUN14[����]"              '28
    
    cboBarType.ListIndex = 7
    
    With spdList
        .BeginUpdate
        .Clear
        
'        .AddColumn "��������", , , "pic"
'        .AddColumn "First", , , , , , True
'        .AddColumn "Second", , , "sec", , [UniGrid Align Left], True
'        .AddColumn "Third", , , , , , True
'        .AddColumn "Forth - RTL", , , , , , True, , True
'        .AddColumn "Fifth", , , , , , True
'        .AddColumn "Boolean", , , "bool", , , True, , , [UniGrid Column Boolean]
'        .AddColumn "Numeric", , [UniGrid Align Right], "num", , , True, , , [UniGrid Column Numeric]
        
        .AddColumn "��������", 70, 2, "01", , , True:    gColWidth(1) = 60
        .AddColumn "�׸񱸺�", 70, 2, "02", , , True:    gColWidth(2) = 60
        .AddColumn "�׸��", 90, 0, "03", , , True:      gColWidth(3) = 50
        .AddColumn "X1��ǥ", 70, 2, "04", , , True, , , [UniGrid Column Text]:    gColWidth(4) = 50
        .AddColumn "X2��ǥ", 70, 2, "05", , , True:      gColWidth(5) = 50
        .AddColumn "Y1��ǥ", 70, 2, "06", , , True:      gColWidth(6) = 50
        .AddColumn "Y2��ǥ", 70, 2, "07", , , True:      gColWidth(7) = 50
        .AddColumn "��Ʈ��", 100, 2, "08", , , True:      gColWidth(8) = 80
        .AddColumn "��Ʈũ��", 70, 2, "09", , , True:     gColWidth(9) = 50
        .AddColumn "��Ʈ����", 70, 2, "10", , , True:       gColWidth(10) = 40
        .AddColumn "��Ʈ��Ʋ��", 70, 2, "11", , , True:     gColWidth(11) = 50
        .AddColumn "��Ʈ����", 70, 2, "12", , , True:       gColWidth(12) = 50
        .AddColumn "��Ʈȸ��", 70, 2, "13", , , True:       gColWidth(13) = 50
        .AddColumn "���ڵ�", 90, 2, "14", , , True:  gColWidth(14) = 70
        .AddColumn "���ڵ���", 0, 2, "15", , , True:          gColWidth(15) = 0
        .AddColumn "���ڵ�ȸ��", 100, 2, "16", , , True:       gColWidth(16) = 50
        .AddColumn "�̹������", 200, 0, "17", , , True: gColWidth(17) = 100
        .AddColumn "����ȸ��", 100, 2, "18", , , True:    gColWidth(18) = 60
        .AddColumn "���εβ�", 80, 2, "19", , , True:       gColWidth(19) = 40
        .AddColumn "������", 80, 2, "20", , , True:         gColWidth(20) = 40
        .AddColumn "��¿���", 100, 2, "21", , , True:    gColWidth(21) = 60
        .AddColumn "��°�", 140, 2, "22", , , True:     gColWidth(22) = 100
        .AddColumn "X��ǥ ������", 0, 2, "23", , , True: gColWidth(23) = 0
        .AddColumn "Y��ǥ ������", 0, 2, "24", , , True: gColWidth(24) = 0
        .AddColumn "��������", 0, 2, "25", , , True:     gColWidth(25) = 0
        .AddColumn "������", 0, 2, "26", , , True:       gColWidth(26) = 0
        .AddColumn "�����ǰ���", 110, 2, "27", , , True:  gColWidth(27) = 70
        .AddColumn "��������", 0, 2, "28", , , True:     gColWidth(28) = 0
        .AddColumn "Tag", 0, 2, "29", , , True:          gColWidth(29) = 0

        .EndUpdate
    
        
    End With
    
    '-- ������
    cmbPrinter.Clear
    For Each x In Printers
        cmbPrinter.AddItem x.DeviceName
    Next
    
    strBuffer = Space(1024)
 
    i = GetProfileString("windows", "Device", "", strBuffer, Len(strBuffer))
    aryPrinter = Split(strBuffer, ",")
    strDefault = Trim(aryPrinter(0))
 
    For Each prtSelectPrinter In Printers
        j = j + 1
        If UCase(Trim(prtSelectPrinter.DeviceName)) = UCase(Trim(strDefault)) Then
            Set Printer = prtSelectPrinter
            boolPrinter_Select_Fales = True
            cmbPrinter.ListIndex = j - 1
            Exit For
        End If
    Next
    
    If optHW(0).Value = True Then   '-- ����
        txtPaperHSize.Text = ""
        txtPaperWSize.Text = ""
    Else                            '-- ����
    End If
    
    '-- Mode Set
    intMode = 0

    '-- ���ڵ� �̹����� �ʱ�ȭ
    strBarImgName = ""
    gOpenFileNm = ""
    
'    txtSpdEdit.Text = ""
'    txtSpdEdit.Visible = False
    
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Sub Form_Load()
    Dim x As Printer
    Dim prtSelectPrinter As Printer
    Dim boolPrinter_Select_Fales As Boolean
    Dim strDefault As String
    Dim Buffer As String
    Dim aryPrinter() As String
    Dim strBuffer As String
    Dim i As Integer
    Dim j As Integer
'    Dim strLicense As String
'    Dim strKey  As String
'
'    strLicense = "License"
'
'    strKey = GetString(HKEY_CURRENT_USER, REG_POSITION, strLicense)
'
'    If strKey = "" Or Not IsDate(strKey) And strKey < Format(Now) Then
'        MsgBox "���̼��� �Ⱓ�� ����Ǿ��ų� �����ϴ�." & vbNewLine & "�����ڿ��� �����Ͻʽÿ�", vbCritical + vbOKOnly, Me.Caption
'        End
'    End If
        
    ' ���� ���� ǥ��
    Me.Caption = Me.Caption & " [Ver " & App.Major & "." & App.Minor & "." & App.Revision & "]"

    'Combo1.ListIndex = 1
    
    Call MDIForm_Tool
    
    Call FrmInitial

    Call GetSetup
        
    txtDevide.Text = gDevide
    
    
    '==== API ���� ���� ���� =================================================
    ReDim CustomColors(0 To 16 * 4 - 1) As Byte
    
    For i = LBound(CustomColors) To UBound(CustomColors)
        CustomColors(i) = 0
    Next i
    '==== API ���� ���� ���� =================================================
    
    Me.ScaleMode = gScaleMode
    
    Me.Top = 0
    Me.Left = 0
    Me.ScaleWidth = 1272
    Me.ScaleHeight = 890
    
    vscPaint.Max = 50 * gDevide
    hscPaint.Max = 50 * gDevide
    
    gLastOpen = ""

'    Picture1.ScaleMode = vbTwips
    
    
End Sub

Private Function ShowOpen(Ufilter As String, Upath As String) As String
    
    OFName.lStructSize = Len(OFName)
    OFName.hwndOwner = Me.hwnd
    OFName.hInstance = App.hInstance
    OFName.lpstrFilter = Ufilter
    OFName.lpstrFile = Space$(254)
    OFName.nMaxFile = 255
    OFName.lpstrFileTitle = Space$(254)
    OFName.nMaxFileTitle = 255
    OFName.lpstrInitialDir = Upath
    OFName.lpstrTitle = "Open File"
    OFName.flags = 0

    If GetOpenFileName(OFName) Then
        ShowOpen = Trim$(OFName.lpstrFile)
        'ShowOpen = Mid(ShowOpen, 1, Len(ShowOpen) - 1)
    Else
        ShowOpen = ""
    End If
    
End Function

Private Function ShowSave(Ufilter As String, Upath As String) As String
    
    OFName.lStructSize = Len(OFName)
    OFName.hwndOwner = Me.hwnd
    OFName.hInstance = App.hInstance
    OFName.lpstrFilter = Ufilter
    OFName.lpstrFile = Space$(254)
    OFName.nMaxFile = 255
    OFName.lpstrFileTitle = Space$(254)
    OFName.nMaxFileTitle = 255
    OFName.lpstrInitialDir = Upath
    OFName.lpstrTitle = "Open File"
    OFName.flags = 0

    If GetSaveFileName(OFName) Then
        ShowSave = Trim$(OFName.lpstrFile)
    Else
        ShowSave = ""
    End If
    
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Sub Form_Unload(Cancel As Integer)

    ' �÷��� �ʱ�ȭ
    Set m_ColCommandButton = Nothing
    Set ClsEventMonitor = Nothing

End Sub

'���ڿ��� byte�� �ǵ��� �ش�.
Function LengthByte(ByVal Var As String) As Long
    Dim Cnt As Long
    Dim num As Long
    Dim TMP As String
    
    Cnt = 0: num = 0
    If Var = "" Then Exit Function
    Do
        Cnt = Cnt + 1: TMP = Mid(Var, Cnt, 1): num = num + 1
        If Asc(TMP) < 0 Then num = num + 1
    Loop Until Cnt >= Len(Var)
    LengthByte = num
End Function

'-- ������ LOF ������ �������忡 ǥ���Ѵ�,
'-- �뵵 : ����,����� ����Ѵ�.
Private Sub SetList(varBuf As Variant)
    Dim intCnt As Integer
    Dim intCol As Integer
    Dim intRow As Integer
    Dim sChain As String
    Dim sEI    As String
    
    With spdList
        sEI = ""
        For intCnt = 0 To UBound(varBuf) - 1
            sEI = sEI & varBuf(intCnt) & vbTab
        Next
        
        If varBuf(1) = "4" Then
            sEI = sEI & gblCtrlNm 'strBarImgName
        Else
            sEI = sEI & Trim(txtTag.Text)
        End If
        
        .AddRow sEI
        '.AddRow = sEI
    End With

End Sub

Private Function BarIdxMapper(idx As Variant) As String
    
    Select Case idx
        Case 0:     BarIdxMapper = 3
        Case 1:     BarIdxMapper = 5
        Case 2:     BarIdxMapper = ""
        Case 3:     BarIdxMapper = 11
        Case 4:     BarIdxMapper = 12
        Case 5:     BarIdxMapper = 1
        Case 6:     BarIdxMapper = 2
        Case 7:     BarIdxMapper = 10
        Case 8:     BarIdxMapper = 22
        Case 9:     BarIdxMapper = 4
        Case 10:    BarIdxMapper = 18
        Case 11:    BarIdxMapper = 6
        Case 12:    BarIdxMapper = 7
        Case 13:    BarIdxMapper = 8
        Case 14:    BarIdxMapper = ""
        Case 15:    BarIdxMapper = 9
        Case 16:    BarIdxMapper = 20
        Case 17:    BarIdxMapper = 13
        Case 18:    BarIdxMapper = 14
        Case 19:    BarIdxMapper = ""
        Case 20:    BarIdxMapper = ""
        Case 21:    BarIdxMapper = ""
        Case 22:    BarIdxMapper = ""
        Case 23:    BarIdxMapper = 15
        Case 24:    BarIdxMapper = ""
        Case 25:    BarIdxMapper = ""
        Case 26:    BarIdxMapper = ""
        Case 27:    BarIdxMapper = ""
        Case 28:    BarIdxMapper = ""
        Case Else:  BarIdxMapper = ""
    End Select

End Function

Private Function BarIdxReMapper(idx As Variant) As String
    
    Select Case idx
        Case 3:     BarIdxReMapper = 0
        Case 5:     BarIdxReMapper = 1
        Case "":     BarIdxReMapper = 2
        Case 11:     BarIdxReMapper = 3
        Case 12:     BarIdxReMapper = 4
        Case 1:     BarIdxReMapper = 5
        Case 2:     BarIdxReMapper = 6
        Case 10:     BarIdxReMapper = 7
        Case 22:     BarIdxReMapper = 8
        Case 4:     BarIdxReMapper = 9
        Case 18:    BarIdxReMapper = 10
        Case 6:    BarIdxReMapper = 11
        Case 7:    BarIdxReMapper = 12
        Case 8:    BarIdxReMapper = 13
        Case "":    BarIdxReMapper = 14
        Case 9:    BarIdxReMapper = 15
        Case 20:    BarIdxReMapper = 16
        Case 13:    BarIdxReMapper = 17
        Case 14:    BarIdxReMapper = 18
        Case 19:    BarIdxReMapper = 19
        Case 20:    BarIdxReMapper = 20
        Case 21:    BarIdxReMapper = 21
        Case "":    BarIdxReMapper = 22
        Case 15:    BarIdxReMapper = 23
        Case 24:    BarIdxReMapper = 24
        Case 25:    BarIdxReMapper = 25
        Case 26:    BarIdxReMapper = 26
        Case 27:    BarIdxReMapper = 27
        Case 28:    BarIdxReMapper = 28
        Case Else:  BarIdxReMapper = ""
    End Select

End Function

Private Sub PaintLine()
    Dim obj                 As Object
    Dim ClsEventObject      As ClassEventObject
    Dim i As Integer
    Dim strTagName As String
    
    'strTagName = txtTag.Text
    '-- ���ζ��α׸���
    For i = 1 To 100
'ReMake:
        txtTag.Text = "LineW_" & i
        Set ClsEventObject = New ClassEventObject
        Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectLine, txtTag.Text)
        If Not obj Is Nothing Then
            obj.x1 = 0
            obj.x2 = 1000
            obj.y1 = i * 15
            obj.y2 = i * 15
            obj.BorderColor = &H8000000F '&HE0E0E0
            obj.BorderStyle = 1
            obj.BorderWidth = 1
        Else
            Set ClsEventObject = Nothing
            Set obj = Nothing
'            GoTo ReMake
            
            Exit Sub
        End If
            
        obj.Visible = True
        Set obj.Container = Picture1
        m_ColCommandButton.Add ClsEventObject
        Set ClsEventObject = Nothing
    Next
    
    '-- ���ζ��α׸���
    For i = 1 To 100
        txtTag.Text = "LineH_" & i
        Set ClsEventObject = New ClassEventObject
        Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectLine, txtTag.Text)
        If Not obj Is Nothing Then
            obj.x1 = i * 15
            obj.x2 = i * 15
            obj.y1 = 0
            obj.y2 = 1000
            obj.BorderColor = &H8000000F '&HE0E0E0
            obj.BorderStyle = 1
            obj.BorderWidth = 1
        Else
            Set ClsEventObject = Nothing
            Exit Sub
        End If
            
        obj.Visible = True
        Set obj.Container = Picture1
        m_ColCommandButton.Add ClsEventObject
        Set ClsEventObject = Nothing
    Next
    

End Sub

'-- ���к��� ������Ʈ ������ �� �׸� ǥ���Ѵ�.
'   ����[varBuf(1)] 0:SText,1:DText,2:SImage,3:DImage,4:Barcode,5:Line
Private Sub MakeLayout(varBuf As Variant)
    Dim strEditObjName      As String
    Dim i As Integer
    Dim strFVar As String
    Dim strTmp
    
MakeAgain:
    
    sstType.Tab = varBuf(1)
    
    txtPaperHSize.Text = varBuf(25)
    txtPaperWSize.Text = varBuf(25)
    
    strFVar = ""
    For i = 1 To Len(varBuf(0))
        If Asc(Mid(varBuf(0), i, 1)) <> 63 Then
           strFVar = strFVar & Mid(varBuf(0), i, 1)
        Else
            'Stop
        End If
    Next
    
    Select Case varBuf(1)
        Case 0  '## Static Label ##
            'txtTag.Text = Replace(varBuf(2), "-", "_")          '�׸��(����)
            txtTag.Text = "Control_" & strFVar
            txtTitle.Text = varBuf(2)                           '�׸��(���)
            txtXpos.Text = varBuf(3)                            'X ��ǥ
            txtYpos.Text = varBuf(5)                            'Y ��ǥ
            txtFontName(0).Text = varBuf(7)                     '��Ʈ��
            txtFontSize(0).Text = varBuf(8)                     '��Ʈũ��
            chkFontBold(0).Value = varBuf(9)                    '    ����
            chkFontUnder(0).Value = varBuf(11)                  '    ����
            chkFontItalic(0).Value = varBuf(10)                 '    ����
                        
            '-- �ؽ�Ʈ �ڽ� �ɼǼ���
            txtContent(0).Font.Name = varBuf(7)                    '��Ʈ��
            txtContent(0).Text = varBuf(21)                     'Text
            
            'txtContent1.Text = varBuf(21)                     'Text
'            txtContent(0).Font.Charset = 163
            chkTStatic.Value = varBuf(26)                       '�����ǰ���
            chkPrint.Value = IIf(varBuf(20) = "1", "0", "1")    '��¾���
        
        Case 1  '## Dynamic Label ##
            'txtTag.Text = Replace(varBuf(2), "-", "_")          '�׸��(����)
            txtTag.Text = "Control_" & strFVar
            txtTitle.Text = varBuf(2)                           '�׸��(���)
            txtXpos.Text = varBuf(3)                            'X ��ǥ
            txtYpos.Text = varBuf(5)                            'Y ��ǥ
            txtFontName(1).Text = varBuf(7)                     '��Ʈ��
            txtFontSize(1).Text = varBuf(8)                     '��Ʈũ��
            chkFontBold(1).Value = varBuf(9)                    '    ����
            chkFontUnder(1).Value = varBuf(11)                  '    ����
            chkFontItalic(1).Value = varBuf(10)                 '    ����
            txtContent(1).Text = varBuf(21)                     'Text
'            txtContent(1).Font.Charset = ""
'            txtContent(1).Font.Charset = 163
            chkPrint.Value = IIf(varBuf(20) = "1", "0", "1")    '��¾���
        
        Case 2  '## Static Image ##
            'txtTag.Text = Replace(varBuf(2), "-", "_")          '�׸��(����)
            txtTag.Text = "Control_" & strFVar
            txtTitle.Text = varBuf(2)                           '�׸��(���)
            txtXpos.Text = varBuf(3)                            'X ��ǥ
            txtYpos.Text = varBuf(5)                            'Y ��ǥ
            txtImageName(0).Text = varBuf(16)                   '�̹������
            txtImageWSize(0).Text = varBuf(4)                   '      ����SIZE
            txtImageHSize(0).Text = varBuf(6)                   '      ����SIZE
            txtImageWSize(2).Text = varBuf(4)                   '      ����SIZE
            txtImageHSize(2).Text = varBuf(6)                   '      ����SIZE
            
            chkIStatic.Value = varBuf(26)                       '�����ǰ���
            chkPrint.Value = IIf(varBuf(20) = "1", "0", "1")    '��¾���
                        
        Case 3  '## Dynamic Image ##
            'txtTag.Text = Replace(varBuf(2), "-", "_")          '�׸��(����)
            txtTag.Text = "Control_" & strFVar
            txtTitle.Text = varBuf(2)                           '�׸��(���)
            txtXpos.Text = varBuf(3)                            'X ��ǥ
            txtYpos.Text = varBuf(5)                            'Y ��ǥ
            txtImageName(1).Text = varBuf(16)                   '�̹������
            txtImageWSize(1).Text = varBuf(4)                   '      ����SIZE
            txtImageHSize(1).Text = varBuf(6)                   '      ����SIZE
            txtImageWSize(3).Text = varBuf(4)                   '      ����SIZE
            txtImageHSize(3).Text = varBuf(6)                   '      ����SIZE
            chkPrint.Value = IIf(varBuf(20) = "1", "0", "1")    '��¾���
            
        Case 4
            'txtTag.Text = Replace(varBuf(2), "-", "_")          '�׸��(����)
            txtTag.Text = "Control_" & strFVar
            txtTitle.Text = varBuf(2)                           '�׸��(���)
            
            
            '-- ���ڵ� Ÿ�� ���� ���α׷��� �ű����α׷� Mapping
            '-- �ε�ø� ����ȴ�.
            If intMode = 0 Then
                strTmp = BarIdxMapper(varBuf(13))
                If strTmp = "" Then
                    cboBarType.ListIndex = 7                   '���ڵ� Ÿ��
                Else
                    cboBarType.ListIndex = strTmp                   '���ڵ� Ÿ��
                End If
            Else
                cboBarType.ListIndex = varBuf(13)                   '���ڵ� Ÿ��
            End If
            
            txtXpos.Text = varBuf(3)                            'X ��ǥ
            txtYpos.Text = varBuf(5)                            'Y ��ǥ
            txtBarData.Text = varBuf(21)                        '���ڵ�Data
            txtBarWSize.Text = varBuf(4)                        '      ����SIZE
            txtBarHSize.Text = varBuf(6)                        '      ����SIZE
            chkBarRotate.Value = IIf(varBuf(15) = "2", "1", "0") '     ȸ��
            chkPrint.Value = IIf(varBuf(20) = "1", "0", "1")    '��¾���
        
        Case 5
            'txtTag.Text = Replace(varBuf(2), "-", "_")          '�׸��(����)
            txtTag.Text = "Control_" & strFVar
            txtTitle.Text = varBuf(2)                           '�׸��(���)
            txtXpos.Text = varBuf(3)                            'X ��ǥ
            txtYpos.Text = varBuf(5)                            'Y ��ǥ
            chkLineRotate.Value = IIf(varBuf(17) = "0", "0", "1")   '����ȸ��
            txtLineHSize.Text = varBuf(18)                      '������
            txtLineWSize.Text = varBuf(19)                      '������
            chkPrint.Value = IIf(varBuf(20) = "1", "0", "1")    '��¾���
            
            If varBuf(17) = "0" Then
                hscPaint.Max = varBuf(19) / 4
            Else
                vscPaint.Max = varBuf(19) / 4
            End If
            
        Case 6  '## Dynamic Label - RFID ##
            'txtTag.Text = Replace(varBuf(2), "-", "_")          '�׸��(����)
            txtTag.Text = "Control_" & strFVar
            txtTitle.Text = varBuf(2)                           '�׸��(���)
            txtXpos.Text = varBuf(3)                            'X ��ǥ
            txtYpos.Text = varBuf(5)                            'Y ��ǥ
            txtFontName(2).Text = varBuf(7)                     '��Ʈ��
            txtFontSize(2).Text = varBuf(8)                     '��Ʈũ��
            chkFontBold(2).Value = varBuf(9)                    '    ����
            chkFontUnder(2).Value = varBuf(11)                  '    ����
            chkFontItalic(2).Value = varBuf(10)                 '    ����
            txtContent(2).Text = varBuf(21)                     'Text
            chkPrint.Value = IIf(varBuf(20) = "1", "0", "1")    '��¾���
            
    End Select
    
    '-- ��ü�̸� ������Ʈ
    gblCtrlNm = txtTag.Text
    gblCtrlIdx = strFVar
    
    '-- ��ü����
    strEditObjName = objMake
    
    If strEditObjName = "0" Then
        '��ü���� ����
    Else
        '��ü���� ����
        varBuf(2) = strEditObjName
        GoTo MakeAgain
    End If

End Sub


Private Sub SetLayout(intTabidx As Integer)

    '����[varBuf(1)] 0:SText,1:DText,2:SImage,3:DImage,4:Barcode,5:Line
    
    Dim intCnt As Integer
    Dim intCol As Integer
    Dim intRow As Integer
    Dim strIdx As String
    Dim strTitle As String
    
    With spdList
        For intRow = 1 To .Rows
            '�׸񱸺�,�׸�� ��
            .Row = intRow
'                If .CellValue(i, 29) = obj.Name Then
'                    .CellValue(i, 4) = frmLabelDesign.txtXpos.Text
'                    .CellValue(i, 6) = frmLabelDesign.txtYpos.Text
'                    Exit For
'                End If
            
            strIdx = .CellValue(intRow, 2)
            strTitle = .CellValue(intRow, 29)
            '.Col = 2: strIdx = Trim(.Text)
            '.Col = 29: strTitle = Trim(.Text)
'            If findSameCtrlNm(3, txtTitle.Text) Then
'                MsgBox "������ �׸���� ����� �� �����ϴ�.", vbInformation, Me.Caption
'                Exit For
'            End If

'            If intTabidx = strIdx And Trim(txtTag.Text) = Trim(strTitle) Then
            If Trim(txtTag.Text) = Trim(strTitle) Then
                Select Case intTabidx
                    Case 0
                        .CellValue(intRow, 2) = intTabidx
                        .CellValue(intRow, 3) = txtTitle.Text
                        .CellValue(intRow, 4) = txtXpos.Text
                        .CellValue(intRow, 6) = txtYpos.Text
                        .CellValue(intRow, 8) = txtFontName(0).Text
                        .CellValue(intRow, 9) = txtFontSize(0).Text
                        .CellValue(intRow, 10) = IIf(chkFontBold(0).Value = "0", "0", "1")
                        .CellValue(intRow, 11) = IIf(chkFontItalic(0).Value = "0", "0", "1")
                        .CellValue(intRow, 12) = IIf(chkFontUnder(0).Value = "0", "0", "1")
                        .CellValue(intRow, 22) = Trim(txtContent(0).Text)
                        .CellValue(intRow, 21) = IIf(chkPrint.Value = "1", "0", "1")     '��¿���
                        .CellValue(intRow, 27) = IIf(chkTStatic.Value = "0", "0", "1")     '�����ǰ���
                        
                    Case 1
                        .CellValue(intRow, 2) = intTabidx
                        .CellValue(intRow, 3) = txtTitle.Text
                        .CellValue(intRow, 4) = txtXpos.Text
                        .CellValue(intRow, 6) = txtYpos.Text
                        .CellValue(intRow, 8) = txtFontName(1).Text
                        .CellValue(intRow, 9) = txtFontSize(1).Text
                        .CellValue(intRow, 10) = IIf(chkFontBold(1).Value = "0", "0", "1")
                        .CellValue(intRow, 11) = IIf(chkFontItalic(1).Value = "0", "0", "1")
                        .CellValue(intRow, 12) = IIf(chkFontUnder(1).Value = "0", "0", "1")
                        .CellValue(intRow, 22) = Trim(txtContent(1).Text)
                        .CellValue(intRow, 21) = IIf(chkPrint.Value = "1", "0", "1")     '��¿���
                        
                    Case 2
                        .CellValue(intRow, 2) = intTabidx
                        .CellValue(intRow, 3) = txtTitle.Text
                        .CellValue(intRow, 4) = txtXpos.Text
                        .CellValue(intRow, 5) = txtImageWSize(0).Text
                        .CellValue(intRow, 6) = txtYpos.Text
                        .CellValue(intRow, 7) = txtImageHSize(0).Text
                        .CellValue(intRow, 17) = txtImageName(0).Text
                        .CellValue(intRow, 21) = IIf(chkPrint.Value = "1", "0", "1")     '��¿���
                        .CellValue(intRow, 27) = IIf(chkIStatic.Value = "0", "0", "1")     '�����ǰ���
            
                    Case 3
                        .CellValue(intRow, 2) = intTabidx
                        .CellValue(intRow, 3) = txtTitle.Text
                        .CellValue(intRow, 4) = txtXpos.Text
                        .CellValue(intRow, 5) = txtImageWSize(1).Text
                        .CellValue(intRow, 6) = txtYpos.Text
                        .CellValue(intRow, 7) = txtImageHSize(1).Text
                        .CellValue(intRow, 17) = txtImageName(1).Text
                        .CellValue(intRow, 21) = IIf(chkPrint.Value = "1", "0", "1")     '��¿���

                    Case 4
                        .CellValue(intRow, 3) = txtTitle.Text
                        .CellValue(intRow, 4) = txtXpos.Text
                        .CellValue(intRow, 5) = txtBarWSize.Text
                        .CellValue(intRow, 6) = txtYpos.Text
                        .CellValue(intRow, 7) = txtBarHSize.Text
                        .CellValue(intRow, 14) = cboBarType.ListIndex   '-- ���ڵ� ����
                        .CellValue(intRow, 16) = IIf(chkBarRotate.Value = "0", "0", "2")    '-- ���ڵ� ȸ��
                        .CellValue(intRow, 22) = Trim(txtBarData.Text)    '-- ���ڵ� ��°�
                        .CellValue(intRow, 21) = IIf(chkPrint.Value = "1", "0", "1")       '��¿���

                    Case 5
                        .CellValue(intRow, 2) = intTabidx
                        .CellValue(intRow, 3) = txtTitle.Text
                        .CellValue(intRow, 4) = txtXpos.Text
                        .CellValue(intRow, 5) = txtLineWSize.Text 'txtYpos.Text
                        .CellValue(intRow, 6) = txtYpos.Text
                        .CellValue(intRow, 7) = txtYpos.Text 'txtLineWSize.Text
                        .CellValue(intRow, 9) = txtLineHSize.Text
                        .CellValue(intRow, 18) = IIf(chkLineRotate.Value = "0", "0", "1")  '����ȸ��
                        .CellValue(intRow, 19) = txtLineHSize.Text                         '���εβ�
                        .CellValue(intRow, 20) = txtLineWSize.Text                         '������
                        .CellValue(intRow, 21) = IIf(chkPrint.Value = "1", "0", "1")       '��¿���
                    
                    Case 6
                        .CellValue(intRow, 2) = intTabidx
                        .CellValue(intRow, 3) = txtTitle.Text
                        .CellValue(intRow, 4) = txtXpos.Text
                        .CellValue(intRow, 6) = txtYpos.Text
                        .CellValue(intRow, 8) = txtFontName(2).Text
                        .CellValue(intRow, 9) = txtFontSize(2).Text
                        .CellValue(intRow, 10) = IIf(chkFontBold(2).Value = "0", "0", "1")
                        .CellValue(intRow, 11) = IIf(chkFontItalic(2).Value = "0", "0", "1")
                        .CellValue(intRow, 12) = IIf(chkFontUnder(2).Value = "0", "0", "1")
                        .CellValue(intRow, 22) = Trim(txtContent(2).Text)
                        .CellValue(intRow, 21) = IIf(chkPrint.Value = "1", "0", "1")     '��¿���

                End Select
                
                Exit Sub
            End If
        Next
    End With

End Sub


Public Function toUTF8(ByVal szSource As String) As String
On Error GoTo ErrHandler

Dim szChar As String
Dim WideChar As Long
Dim nLength As Integer
Dim i As Integer

    nLength = Len(szSource)
    
    For i = 1 To nLength
        szChar = Mid(szSource, i, 1)
        
        If Asc(szChar) < 0 Then
            WideChar = CLng(AscB(MidB(szChar, 2, 1))) * 256 + AscB(MidB(szChar, 1, 1))
        
            If (WideChar And &HFF80) = 0 Then
                toUTF8 = toUTF8 & Hex(WideChar)
            ElseIf (WideChar And &HF000) = 0 Then
                toUTF8 = toUTF8 & _
                Hex(CInt((WideChar And &HFFC0) / 64) Or &HC0) & _
                Hex(WideChar And &H3F Or &H80)
            Else
                toUTF8 = toUTF8 & _
                Hex(CInt((WideChar And &HF000) / 4096) Or &HE0) & _
                Hex(CInt((WideChar And &HFFC0) / 64) And &H3F Or &H80) & _
                Hex(WideChar And &H3F Or &H80)
        
            End If
        Else
            toUTF8 = toUTF8 & Hex(Asc(szChar))
        End If
    Next

Exit Function

ErrHandler:
    toUTF8 = ""

End Function

Public Function URLEncode(URLStr As String) As String

Dim sURL        As String   '** �Է¹��� URL ���ڿ�
Dim sBuffer     As String   '** URL ���ڵ� ó�� �� URL �� ���� ���� ���ڿ�
Dim sTemp       As String   '** �ӽ� ���ڿ�
Dim cChar       As String   '** URL ���ڿ� �� ���� ���ؽ��� ����
Dim lErrNum     As Long     '** ���� ��ȣ
Dim sErrSource  As String   '** ���� �ҽ�
Dim sErrDesc    As String   '** �ҷ� ����
Dim sMsg        As String   '** ���� �޼���
Dim Index       As Integer

On Error GoTo ErrorHanddle:

    sURL = Trim(URLStr) '** URL ���ڿ��� ��´�.
    sBuffer = "" '** �ӽ� ���ۿ� ���ڿ� ���� �ʱ�ȭ.

    '******************************************************
    '* URL ���ڵ� �۾�
    '******************************************************

    For Index = 1 To Len(sURL)
        '** ���� �ε����� ���ڸ� ��´�.
        cChar = Mid(sURL, Index, 1)
        
        If cChar = "0" Or (cChar >= "1" And cChar <= "9") Or (cChar >= "a" And cChar <= "z") Or (cChar >= "A" And cChar <= "Z") Or _
                          cChar = "-" Or cChar = "_" Or cChar = "." Or cChar = "*" Then
            '** URL �� ���Ǵ� ���ڵ� :: ���� ���ڿ��� �߰��Ѵ�.
            sBuffer = sBuffer & cChar
        ElseIf cChar = " " Then
            '** ���� ���� :: + �� ��ü�Ͽ� ���� ���ڿ��� �߰��Ѵ�.
            sBuffer = sBuffer & "+"
        Else
            '** URL �� ������ �ʴ� ���ڵ� :: % �� ���ڵ��ؼ� ���� ���ڿ��� �߰��Ѵ�.
            sTemp = CStr(Hex(Asc(cChar)))
            If Len(sTemp) = 4 Then
                sBuffer = sBuffer & "%" & Left(sTemp, 2) & "%" & Mid(sTemp, 3, 2)
            ElseIf Len(sTemp) = 2 Then
                sBuffer = sBuffer & "%" & sTemp
            End If
        End If
    Next

    '** ����� �����Ѵ�.
    URLEncode = sBuffer

Exit Function

ErrorHanddle:

    '** ������ �߻��ϸ� ���� ���ڸ� �����Ѵ�.
    URLEncode = ""
    
    '** ���� ������ ��´�.
    lErrNum = Err.Number
    sErrSource = Err.Source
    sErrDesc = Err.Description
    
    '** �̺�Ʈ �α׿� ������ ����Ѵ�.
    sMsg = vbCrLf & vbCrLf & _
    "Error Object : EgoCube.URLTools," & vbCrLf & _
    "Error Method : Public Function URLEncode(URLStr As String) As String," & vbCrLf & _
    "Error Number : " & lErrNum & "," & vbCrLf & _
    "Error Source : " & sErrSource & "," & vbCrLf & _
    "Error Description : " & sErrDesc
    
    App.LogEvent sMsg, vbLogEventTypeError
    
    '** ������ �߻���Ų��.
    Err.Raise lErrNum, sErrSource, sErrDesc
    

Exit Function


End Function

Private Sub mnuOpen_Click()
    Dim strSrcfile  As Variant
    Dim varBuffer() As Variant
    Dim varBuf      As Variant
    Dim lngBufLen   As Long
    Dim i           As Long
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim j           As Long
    Dim bytBuff()   As Byte
    
    Static ChkSumCnt As Long
    Dim strTxt As String
    
    Dim FileNumber As Long
    Dim FileName As String
    Dim FileCount As Long
    Dim LineCount As Long
    Dim FileOpenNumber As Integer
    Dim Data As String
    Dim splitdata() As String
    
    Dim utf8() As Byte
    Dim ucs2 As Variant
    Dim chars As Long
    Dim varTmp As Variant
    Dim sFile As String
    
    ' ���ʱ�ȭ
    Call FrmInitial
    
    'Cancel�� True�� �����մϴ�.
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler

    '��� �Ӽ��� �����մϴ�.
    If gLastOpen = "" Then
        CommonDialog1.InitDir = App.Path & "\" & gLayOut
    Else
        CommonDialog1.InitDir = gLastOpen
    End If
    
    CommonDialog1.Filter = "LayoutFile(*.lof)|*.lof"

    '[����] ��ȭ ���ڸ� ǥ���մϴ�.
    CommonDialog1.ShowOpen
    strSrcfile = CommonDialog1.FileName

    '�÷��� �ʱ�ȭ
    Set m_ColCommandButton = Nothing
    Set m_ColCommandButton = New Collection

    'LOF ���� ����
    FileName = CommonDialog1.FileName
    
'    '-- �ֱ� ������ ��η� ����â ����
'    If gLastOpen = "" Then
'        FileName = ShowOpen("LayoutFile(*.lof)|*.lof", App.Path & "\" & gLayOut)
'    Else
'        FileName = ShowOpen("LayoutFile(*.lof)|*.lof", gLastOpen & gLayOut)
'    End If
    
    If FileName <> "" Then
        varTmp = Split(FileName, "\")
        Me.Caption = varTmp(UBound(varTmp))
        FileOpenNumber = FreeFile()
        LineCount = 0
    
        gOpenFileNm = FileName
        
        Open FileName For Binary As #1   'UTF-8 ��������
        ReDim utf8(LOF(1))
        
        Get #1, , utf8
            
    ''��Ƽ����Ʈ���� �����ڵ� ��ȯ ���
    ''  // sTime�̶� ANSI �������� bstr�̶� �̸��� �����ڵ�(BSTRŸ��) ������ ��ȯ
    ''  char sTime[] = '�����ڵ� ��ȯ ����';
    ''  BSTR bstr;
    ''  // sTime�� �����ڵ�� ��ȯ�ϱ⿡ �ռ� ���� �װ��� �����ڵ忡���� ���̸� �˾ƾ� �Ѵ�.
    ''  int nLen = MultiByteToWideChar(CP_ACP, 0, sTime, lstrlen(sTime), NULL, NULL)
    ''  // �� ���̸�ŭ �޸𸮸� �Ҵ��Ѵ�.
    ''  bstr = SysAllocStringLen(NULL, nLen);
    ''  // ���� ��ȯ�� �����Ѵ�.
    ''  MultiByteToWideChar(CP_ACP, 0, sTime, lstrlen(sTime), bstr, nLen);
    
        chars = MultiByteToWideChar(CP_UTF8, 0, VarPtr(utf8(0)), LOF(1), 0, 0)
        ucs2 = Space(chars)
        chars = MultiByteToWideChar(CP_UTF8, 0, VarPtr(utf8(0)), LOF(1), StrPtr(ucs2), chars)
        varBuf = Split(ucs2, Chr(13))
        
        Close #1
                   
        '������ LOF���� ���ۿ� ����
        For i = 0 To UBound(varBuf)
            ReDim Preserve varBuffer(i)
            varBuffer(LineCount) = varBuf(i)
            LineCount = LineCount + 1
        Next
        
        
        '������ LOF���� ȭ��׸���/�������徲��
        For i = 0 To UBound(varBuffer) - 1
            If varBuffer(i) <> "" Then
                varBuf = Split(varBuffer(i), "^")
                'Debug.Print varBuffer(i)
                Call MakeLayout(varBuf)
                Call SetList(varBuf)
            End If
        Next
                
        Call PaintLine
        
    '    intMode = 1
        
        For i = Len(FileName) To 1 Step -1
            'Debug.Print Mid(FileName, i, 1)
            If Mid(FileName, i, 1) = "\" Then
                gLastOpen = Mid(FileName, 1, i)
                Exit For
            End If
        Next
    
    End If
    
    Exit Sub

ErrHandler:

End Sub



Private Sub picUndo_Click()
    Dim Moveobj As Variant
    Dim x, y As Long
    
    Moveobj = LMousePos.obj
    x = LMousePos.fromx
    y = LMousePos.fromy

    Me.Controls(Moveobj).Left = x
    Me.Controls(Moveobj).Top = y
End Sub

Private Sub SetControl(intRow As Long)

Dim strTmp As String
Dim intTab As Integer

    With spdList
        '-- ����
        sstType.Tab = Trim(.CellValue(intRow, 2))
        intTab = sstType.Tab
        txtTitle.Text = Trim(.CellValue(intRow, 3))
        txtTag.Text = Trim(.CellValue(intRow, 29))
        '-- ��ġ
        txtXpos.Text = Trim(.CellValue(intRow, 4))
        txtYpos.Text = Trim(.CellValue(intRow, 6))
        '-- ����,����(�β�)
        Select Case intTab
            Case 2: txtImageWSize(0).Text = Trim(.CellValue(intRow, 5))
                    txtImageHSize(0).Text = Trim(.CellValue(intRow, 7))
                    txtImageWSize(2).Text = Trim(.CellValue(intRow, 5))
                    txtImageHSize(2).Text = Trim(.CellValue(intRow, 5))
            Case 3: txtImageWSize(1).Text = Trim(.CellValue(intRow, 5))
                    txtImageHSize(1).Text = Trim(.CellValue(intRow, 7))
                    txtImageWSize(3).Text = Trim(.CellValue(intRow, 5))
                    txtImageHSize(3).Text = Trim(.CellValue(intRow, 7))
            Case 4: txtBarWSize.Text = Trim(.CellValue(intRow, 5))
                    txtBarHSize.Text = Trim(.CellValue(intRow, 7))
        End Select
        '-- ��Ʈ
        Select Case intTab
            Case 0: txtFontName(0).Text = Trim(.CellValue(intRow, 8))
                    txtFontSize(0).Text = Trim(.CellValue(intRow, 9))
                    chkFontBold(0).Value = IIf(Trim(.CellValue(intRow, 10)) = "0", "0", "1")   '��Ʈ����
                    chkFontUnder(0).Value = IIf(Trim(.CellValue(intRow, 11)) = "0", "0", "1")  '��Ʈ����
                    chkFontItalic(0).Value = IIf(Trim(.CellValue(intRow, 12)) = "0", "0", "1") '��Ʈ����
                    chkFontItalic(0).Value = IIf(Trim(.CellValue(intRow, 13)) = "0", "0", "1") '��Ʈȸ��
            Case 1: txtFontName(1).Text = Trim(.CellValue(intRow, 8))
                    txtFontSize(1).Text = Trim(.CellValue(intRow, 9))
                    chkFontBold(1).Value = IIf(Trim(.CellValue(intRow, 10)) = "0", "0", "1")   '��Ʈ����
                    chkFontUnder(1).Value = IIf(Trim(.CellValue(intRow, 11)) = "0", "0", "1")  '��Ʈ����
                    chkFontItalic(1).Value = IIf(Trim(.CellValue(intRow, 12)) = "0", "0", "1") '��Ʈ����
                    'chkFontItalic(0).Value = IIf(Trim(.CellValue(intRow, 13)) = "0", "0", "1") '��Ʈȸ��
            Case 6: txtFontName(2).Text = Trim(.CellValue(intRow, 8))
                    txtFontSize(2).Text = Trim(.CellValue(intRow, 9))
                    chkFontBold(2).Value = IIf(Trim(.CellValue(intRow, 10)) = "0", "0", "1")   '��Ʈ����
                    chkFontUnder(2).Value = IIf(Trim(.CellValue(intRow, 11)) = "0", "0", "1")  '��Ʈ����
                    chkFontItalic(2).Value = IIf(Trim(.CellValue(intRow, 12)) = "0", "0", "1") '��Ʈ����
        
        End Select
        '-- ���ڵ�
        '-- ���ڵ� Ÿ�� ���� ���α׷��� �ű����α׷� Mapping
        strTmp = BarIdxMapper(Trim(.CellValue(intRow, 14)))
        If strTmp = "" Then
            cboBarType.ListIndex = 7
        Else
            cboBarType.ListIndex = strTmp
        End If
        txtBarDevide.Text = Trim(.CellValue(intRow, 15))
        chkBarRotate.Value = IIf(Trim(.CellValue(intRow, 16)) = "0", 0, 2)
        
        '-- �̹���
        If intTab = 3 Then
            txtImageName(0).Text = Trim(.CellValue(intRow, 17))
        ElseIf intTab = 4 Then
            txtImageName(1).Text = Trim(.CellValue(intRow, 17))
        End If
        
        '-- ����
        chkLineRotate.Value = IIf(Trim(.CellValue(intRow, 18)) = "0", 0, 1)
        txtLineHSize.Text = Trim(.CellValue(intRow, 19))
        txtLineWSize.Text = Trim(.CellValue(intRow, 20))
        
        '-- ��¿���
        chkPrint.Value = IIf(Trim(.CellValue(intRow, 21)) = "1", 0, 1)
        '-- ��°�
        Select Case intTab
            Case 0: txtContent(0).Text = Trim(.CellValue(intRow, 22))
            Case 1: txtContent(1).Text = Trim(.CellValue(intRow, 22))
            Case 4: txtBarData.Text = Trim(.CellValue(intRow, 22))
            Case 6: txtContent(2).Text = Trim(.CellValue(intRow, 22))
        End Select
        
        '-- �����ǰ���
        If intTab = 0 Then
            chkTStatic.Value = IIf(Trim(.CellValue(intRow, 27)) = "0", 0, 1)
        ElseIf intTab = 2 Then
            chkIStatic.Value = IIf(Trim(.CellValue(intRow, 27)) = "0", 0, 1)
        End If
        
    End With

End Sub


Private Sub spdList_Click()
    If spdList.Rows = 0 Then
        Exit Sub
    End If
    
    Call SetControl(spdList.Row)
    
End Sub


Private Sub spdList_HScroll()
'    txtSpdEdit.Text = ""
'    txtSpdEdit.Visible = False

End Sub

'Private Sub spdList_DblClick()
'
'    Dim intRow As Integer
'    Dim intCol As Integer
'    Dim lngLeft As Long
'    Dim lngTop  As Long
'
'    lngLeft = 0
'    lngTop = 0
'
'    With spdList
'        For intRow = 1 To .Rows
'            If intRow = .Row Then
'                lngTop = .Top + (intRow * 15)
'                txtSpdEdit.Left = .Left + lngLeft
'                'txtSpdEdit.Top = lngTop
'                Exit For
'            Else
'                lngLeft = lngLeft + .ColWidth(intRow)
'            End If
'        Next
'
'    End With
'
'End Sub

Private Sub spdList_KeyPress(KeyAscii As Integer)
    Dim varTmp As Variant

    If KeyAscii = 13 Then

        Call SetControl(spdList.Row)

        intMode = 1

        Call cmdSet_Click

    End If

End Sub



Private Sub spdList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim intCol      As Integer
    Dim SelRow      As Integer
    Dim sumColWidth As Long
    Dim calY        As Double
    Dim intRow      As Integer
    Dim modY        As Long

    Dim lngRow As Long
    Dim lngCol As Long
    Dim i As Integer
    
    'Debug.Print spdList.GetColByX(x)
    If spdList.Row <= 0 Then
        Exit Sub
    End If
    
    txtSpdEdit.Visible = True
    lngCol = spdList.GetColByX(x)
    Select Case lngCol
        Case 1, 2, 3 ', 8, 14, 15, 16, 17
'            txtSpdEdit.Visible = False
            Exit Sub
        Case Else
            txtSpdEdit.Text = spdList.CellValue(spdList.Row, lngCol)
            txtSpdEdit.Tag = lngCol
            
            sumColWidth = 0
            For i = 1 To lngCol
                sumColWidth = sumColWidth + spdList.ColWidth(i)
            Next
            
            'txtSpdEdit.Left = sumColWidth
            'txtSpdEdit.Top = spdList.Top + y
            
            txtSpdEdit.SelStart = 0
            txtSpdEdit.SelLength = Len(txtSpdEdit.Text)
            txtSpdEdit.SetFocus
    End Select
    

    
'    Dim intCol      As Integer
'    Dim SelRow      As Integer
'    Dim sumColWidth As Long
'    Dim calY        As Double
'    Dim intRow      As Integer
'    Dim modY        As Long
'
'    If spdList.Rows = 0 Then
'        Exit Sub
'    End If
'
'    SelRow = spdList.Row
'
'    If spdList.Row <= 0 Then
'        Exit Sub
'    End If
'
'    With spdList
'        sumColWidth = 0
'        For intCol = 1 To 29
'            sumColWidth = sumColWidth + gColWidth(intCol)
'            If sumColWidth >= x Then
'
'                Select Case intCol
'                Case 1, 2, 3, 8, 14, 15, 16, 17
'                    txtSpdEdit.Visible = False
'                    Exit Sub
'                Case Else
'                    txtSpdEdit.Visible = True
'                    txtSpdEdit.Text = spdList.CellValue(spdList.Row, intCol)
'                    txtSpdEdit.Tag = intCol
'                    txtSpdEdit.SelStart = 0
'                    txtSpdEdit.SelLength = Len(txtSpdEdit.Text)
'
'                    txtSpdEdit.Left = (spdList.Left + sumColWidth) - 5
'
'                    'calY = Round(y / 17.5, 0)
'                    calY = y / 17.5
'                    calY = calY * 17.5
'
'                    modY = calY Mod 17.5
'
'                    calY = calY - modY
''                    For intRow = 1 To .Rows
''                        If (17.5 * spdList.Row) <= y * (spdList.Row) Then
''                            'Stop
''                            calY = 17.5 * spdList.Row
''                            Exit For
''                        End If
''                    Next
'
'                    'calY = y / 17                    'calY = Mid(calY, 1, InStr(calY, ".") - 1)
'                    txtSpdEdit.Top = spdList.Top + calY  'y
'                    txtSpdEdit.SetFocus
'
'                    Exit For
'                End Select
'            End If
'        Next
'
'    End With
    

    
End Sub



'Private Sub spdList_VScroll()
'    txtSpdEdit.Text = ""
'    txtSpdEdit.Visible = False
'
'End Sub

'Private Sub spdList_LeaveRow(ByVal Row As Long, ByVal RowWasLast As Boolean, ByVal RowChanged As Boolean, ByVal AllCellsHaveData As Boolean, ByVal NewRow As Long, ByVal NewRowIsLast As Long, Cancel As Boolean)
'
'    Call SetControl(NewRow)
'
'End Sub

'Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If X > (Command2.Width - 100) And Y > (Command2.Height - 100) And Button = vbLeftButton Then
'        drageMode = True
'    Else
'        drageMode = False
'    End If
'    If drageMode Then
'        Command2.Height = Y
'        Command2.Width = X
'    End If
'End Sub


Private Sub sstType_Click(PreviousTab As Integer)
Dim intTab As Integer

    intTab = sstType.Tab
    txtTitle.Enabled = True
    
    Select Case intTab
        Case 0
            txtTitle.Text = "S_TEXT" & gblCtrlIdx
            'cmdFont(0).SetFocus
        Case 1
            txtTitle.Text = "D_TEXT" & gblCtrlIdx
            'cmdFont(1).SetFocus
        Case 2
            txtTitle.Text = "S_Image" & gblCtrlIdx
            'cmdImage(0).SetFocus
        Case 3
            txtTitle.Text = "D_Image" & gblCtrlIdx
            'cmdImage(1).SetFocus
        Case 4
            txtTitle.Text = "BARCODE" & gblCtrlIdx
            'cboBarType.SetFocus
        Case 5
            txtTitle.Text = "LINE" & gblCtrlIdx
            'txtLineHSize.SetFocus
            txtLineHSize.Text = "1"
        Case 6
            txtTitle.Text = "RFID"
            'txtTitle.Enabled = False
            txtLineHSize.Text = "1"
    End Select
    
    'txtTag.Text = ""
    txtXpos.Text = 10
    txtYpos.Text = 10
    
    cboType.ListIndex = intTab

End Sub



Private Sub tlbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case TLBKEY_NEW
            Call mnuNew_Click
        Case TLBKEY_OPEN
            Call mnuOpen_Click
        Case TLBKEY_SAVE
            Call mnuSave_Click
        Case TLBKEY_MAKE
            Call mnuMake_Click
        Case TLBKEY_VIEW
            Call mnuView_Click
        Case TLBKEY_EDIT
            Call mnuSet_Click
        Case TLBKEY_EDIT
            Call mnuSet_Click
        Case TLBKEY_EXIT
            Call mnuClose_Click
    End Select

End Sub

Private Sub tmrMove_Timer()
    
    Call objMove(intMoveIdx)

End Sub


Private Sub txtBarHSize_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Not IsNumeric(Trim(txtBarHSize.Text)) Then
            MsgBox "���ڸ� �Է��� �����մϴ�.", vbOKOnly + vbInformation, Me.Caption
            txtBarHSize.SetFocus
        End If
    End If
    
End Sub

Private Sub txtBarWSize_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        If Not IsNumeric(Trim(txtBarWSize.Text)) Then
            MsgBox "���ڸ� �Է��� �����մϴ�.", vbOKOnly + vbInformation, Me.Caption
            txtBarWSize.SetFocus
        End If
    End If

End Sub


''Private Sub txtDevide_KeyPress(KeyAscii As Integer)
''    Dim intRow      As Integer
''    Dim intCol      As Integer
''    Dim strBuf()    As String
''
''    If KeyAscii = 13 Then
''        If IsNumeric(txtDevide.Text) Then
''            gDevide = txtDevide.Text
''
''            ' �÷��� �ʱ�ȭ
''            Set m_ColCommandButton = Nothing
''            Set m_ColCommandButton = New Collection
''
''            With spdList
''                For intRow = 1 To .MaxRows
''                    .Row = intRow
''                    .Col = 1
''                    Erase strBuf
''                    If Trim(.Text) <> "" Then
''                        ReDim Preserve strBuf(.MaxCols) As String
''                        For intCol = 2 To .MaxCols
''                            .Col = intCol
''                            strBuf(intCol - 1) = Trim(.Text)
''                        Next
''                        Call MakeLayout(strBuf)
''                        Erase strBuf
''                    End If
''                Next
''            End With
''        Else
''            MsgBox "���ڸ� �Է��� �����մϴ�.", vbInformation, Me.Caption
''            txtDevide.SetFocus
''            Exit Sub
''        End If
''    End If
''End Sub


Private Sub txtFontSize_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        If Not IsNumeric(Trim(txtFontSize(Index).Text)) Then
            MsgBox "���ڸ� �Է��� �����մϴ�.", vbOKOnly + vbInformation, Me.Caption
            txtFontSize(Index).SetFocus
        End If
    End If

End Sub


Private Sub txtImageDevide_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        Call cmdImageDevSet_Click(Index)
    End If
    
End Sub

Private Sub txtImageHSize_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Not IsNumeric(Trim(txtImageHSize(Index).Text)) Then
            MsgBox "���ڸ� �Է��� �����մϴ�.", vbOKOnly + vbInformation, Me.Caption
            txtImageHSize(Index).SetFocus
        End If
    End If

End Sub

Private Sub txtImageWSize_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Not IsNumeric(Trim(txtImageWSize(Index).Text)) Then
            MsgBox "���ڸ� �Է��� �����մϴ�.", vbOKOnly + vbInformation, Me.Caption
            txtImageWSize(Index).SetFocus
        End If
    End If

End Sub

Private Sub txtLineHSize_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        If Not IsNumeric(Trim(txtLineHSize.Text)) Then
            MsgBox "���ڸ� �Է��� �����մϴ�.", vbOKOnly + vbInformation, Me.Caption
            txtLineHSize.SetFocus
        End If
    End If

End Sub

Private Sub txtLineWSize_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        If Not IsNumeric(Trim(txtLineWSize.Text)) Then
            MsgBox "���ڸ� �Է��� �����մϴ�.", vbOKOnly + vbInformation, Me.Caption
            txtLineWSize.SetFocus
        End If
    End If

End Sub

Private Sub txtPaperHSize_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        If Not IsNumeric(Trim(txtPaperHSize.Text)) Then
            MsgBox "���ڸ� �Է��� �����մϴ�.", vbOKOnly + vbInformation, Me.Caption
            txtPaperHSize.SetFocus
        End If
    End If

End Sub

Private Sub txtPaperWSize_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Not IsNumeric(Trim(txtPaperWSize.Text)) Then
            MsgBox "���ڸ� �Է��� �����մϴ�.", vbOKOnly + vbInformation, Me.Caption
            txtPaperWSize.SetFocus
        End If
    End If

End Sub



Private Sub txtSpdEdit_KeyPress(KeyAscii As Integer, Shift As Integer)
    
    Dim varTmp As Variant

    If KeyAscii = 27 Then
        txtSpdEdit.Tag = ""
        txtSpdEdit.Text = ""
'        txtSpdEdit.Visible = False
    End If
    
    If KeyAscii = 13 Then
        
        spdList.CellValue(spdList.Row, CLng(txtSpdEdit.Tag)) = txtSpdEdit.Text
        
        Call SetControl(spdList.Row)

        intMode = 1

        Call cmdSet_Click
        
        txtSpdEdit.Tag = ""
        txtSpdEdit.Text = ""
'        txtSpdEdit.Visible = False

    End If


End Sub

Private Sub txtSpdEdit_LostFocus()
    
'    txtSpdEdit.Text = ""
'    txtSpdEdit.Visible = False
    
End Sub

Private Sub txtXpos_Change()
    
    If txtXpos.Text <> "" And IsNumeric(txtXpos.Text) Then
        txtXmm.Text = txtXpos.Text / 3.779
    Else
        MsgBox "���ڸ� �Է��� �����մϴ�.", vbOKOnly + vbInformation, Me.Caption
        txtXpos.SetFocus
    End If

End Sub

Private Sub txtXpos_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Not IsNumeric(Trim(txtXpos.Text)) Then
            MsgBox "���ڸ� �Է��� �����մϴ�.", vbOKOnly + vbInformation, Me.Caption
            txtXpos.SetFocus
        End If
    End If

End Sub

Private Sub txtYpos_Change()
    
    If txtYpos.Text <> "" And IsNumeric(txtYpos.Text) Then
        txtYmm.Text = txtYpos.Text / 3.779
    Else
        MsgBox "���ڸ� �Է��� �����մϴ�.", vbOKOnly + vbInformation, Me.Caption
        txtXpos.SetFocus
    End If
    
End Sub

Private Sub txtYpos_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Not IsNumeric(Trim(txtYpos.Text)) Then
            MsgBox "���ڸ� �Է��� �����մϴ�.", vbOKOnly + vbInformation, Me.Caption
            txtYpos.SetFocus
        End If
    End If

End Sub

Private Sub vscPaint_Change()
  Dim lngPicPaintTop As Long
  
  On Error GoTo ErrorHandler
  
    If Abs(gblPicVval) < CLng(vscPaint.Value) * 10 Then
        lngPicPaintTop = -(CLng(vscPaint.Value) * 10)
        Picture1.Top = Picture1.Top + lngPicPaintTop
        Picture1.Height = Picture1.Height + Abs(lngPicPaintTop)
    Else
        lngPicPaintTop = CLng(vscPaint.Value) * 10
        If lngPicPaintTop = 0 Then
            Picture1.Top = 0
            Picture1.Height = Picture1.Height + Abs(lngPicPaintTop)
        Else
            Picture1.Top = Picture1.Top + lngPicPaintTop
            Picture1.Height = Picture1.Height + Abs(lngPicPaintTop)
        End If
    End If
  
    gblPicVval = lngPicPaintTop
      
  
    Exit Sub

ErrorHandler:
'  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description

End Sub
'
'Private Sub vscPaint_Scroll()
'  Dim lngPicPaintTop As Long
'
'  On Error GoTo ErrorHandler
'
'    If Abs(gblPicVval) < CLng(vscPaint.Value) * 10 Then
'        lngPicPaintTop = -(CLng(vscPaint.Value) * 10)
'        Picture1.Top = Picture1.Top + lngPicPaintTop
'        Picture1.Height = Picture1.Height + Abs(lngPicPaintTop)
'    Else
'        lngPicPaintTop = CLng(vscPaint.Value) * 10
'        Picture1.Top = Picture1.Top + lngPicPaintTop
'        Picture1.Height = Picture1.Height + Abs(lngPicPaintTop)
'    End If
'
'    gblPicVval = lngPicPaintTop
'
'
'    Exit Sub
'
'ErrorHandler:
''  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
'
'End Sub
