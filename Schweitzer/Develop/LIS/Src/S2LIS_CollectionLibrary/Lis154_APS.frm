VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frm154NurCol 
   BackColor       =   &H00DBE6E6&
   ClientHeight    =   9060
   ClientLeft      =   -315
   ClientTop       =   420
   ClientWidth     =   14535
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9060
   ScaleWidth      =   14535
   WindowState     =   2  '�ִ�ȭ
   Begin VB.CommandButton cmdWardHelp 
      BackColor       =   &H00F7FDFD&
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1305
      Style           =   1  '�׷���
      TabIndex        =   38
      Top             =   105
      Width           =   255
   End
   Begin MedControls1.LisLabel lblWardId 
      Height          =   240
      Left            =   1560
      TabIndex        =   37
      Top             =   105
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   423
      BackColor       =   8421504
      ForeColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      AutoSize        =   -1  'True
      Caption         =   "61W"
   End
   Begin MedControls1.LisLabel lblBar 
      Height          =   285
      Left            =   4425
      TabIndex        =   24
      Top             =   2610
      Width           =   9960
      _ExtentX        =   17568
      _ExtentY        =   503
      BackColor       =   8421504
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "��ü ä�� ����Ʈ"
      LeftGab         =   100
   End
   Begin VB.Frame fraOrder 
      BackColor       =   &H00DBE6E6&
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   4425
      TabIndex        =   28
      Top             =   2820
      Width           =   9990
      Begin VB.PictureBox picOrdDiv 
         Appearance      =   0  '���
         BackColor       =   &H00DBE6E6&
         BorderStyle     =   0  '����
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2145
         ScaleHeight     =   300
         ScaleWidth      =   3690
         TabIndex        =   33
         Top             =   180
         Width           =   3690
         Begin VB.Shape Shape2 
            BackColor       =   &H00553755&
            BackStyle       =   1  '�������� ����
            BorderColor     =   &H00C0C0C0&
            Height          =   165
            Left            =   2325
            Shape           =   3  '����
            Top             =   60
            Width           =   330
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '����
            Caption         =   "�ӻ󺴸�"
            ForeColor       =   &H00404040&
            Height          =   225
            Left            =   2610
            TabIndex        =   36
            Top             =   60
            Width           =   720
         End
         Begin VB.Shape Shape3 
            BackColor       =   &H005E3F00&
            BackStyle       =   1  '�������� ����
            BorderColor     =   &H00C0C0C0&
            Height          =   165
            Left            =   90
            Shape           =   3  '����
            Top             =   60
            Width           =   330
         End
         Begin VB.Shape Shape4 
            BackColor       =   &H00496835&
            BackStyle       =   1  '�������� ����
            BorderColor     =   &H00C0C0C0&
            Height          =   165
            Left            =   1185
            Shape           =   3  '����
            Top             =   60
            Width           =   330
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '����
            Caption         =   "�غκ���"
            ForeColor       =   &H00404040&
            Height          =   225
            Left            =   390
            TabIndex        =   35
            Top             =   60
            Width           =   720
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  '����
            Caption         =   "��������"
            ForeColor       =   &H00404040&
            Height          =   225
            Left            =   1485
            TabIndex        =   34
            Top             =   60
            Width           =   720
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  '�������� ����
            BorderColor     =   &H00808080&
            FillColor       =   &H00E0E0E0&
            Height          =   300
            Index           =   2
            Left            =   0
            Shape           =   4  '�ձ� �簢��
            Top             =   0
            Width           =   3525
         End
      End
      Begin VB.CheckBox chkSelAll 
         BackColor       =   &H00DBE6E6&
         Caption         =   "��ü ����(&A)"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H004A4189&
         Height          =   315
         Left            =   120
         TabIndex        =   30
         Top             =   180
         Width           =   1470
      End
      Begin VB.CheckBox chkChangeColTm 
         BackColor       =   &H00DBE6E6&
         Caption         =   "ä���ð����� : "
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H004A4189&
         Height          =   300
         Left            =   6465
         TabIndex        =   29
         Top             =   180
         Width           =   1500
      End
      Begin FPSpread.vaSpread tblOrdSheet 
         Height          =   5070
         Left            =   90
         TabIndex        =   31
         Tag             =   "10114"
         Top             =   495
         Width           =   9825
         _Version        =   196608
         _ExtentX        =   17330
         _ExtentY        =   8943
         _StockProps     =   64
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         EditEnterAction =   5
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GridColor       =   14737632
         MaxCols         =   35
         MaxRows         =   19
         ProcessTab      =   -1  'True
         Protect         =   0   'False
         ScrollBars      =   2
         ShadowColor     =   14737632
         ShadowDark      =   14737632
         ShadowText      =   0
         SpreadDesigner  =   "Lis154_APS.frx":0000
         StartingColNumber=   2
         VirtualRows     =   24
         VisibleCols     =   5
         VisibleRows     =   19
      End
      Begin MSComCtl2.DTPicker dtpColDtTm 
         Height          =   300
         Left            =   8010
         TabIndex        =   32
         Top             =   165
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   14737632
         CalendarTitleBackColor=   14737632
         CustomFormat    =   "yyyy-MM-dd H:mm"
         Format          =   71106563
         UpDown          =   -1  'True
         CurrentDate     =   36851.6291666667
      End
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ä   �� (&S)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   10500
      Style           =   1  '�׷���
      TabIndex        =   27
      Tag             =   "0"
      Top             =   8535
      Width           =   1260
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ȭ������(&C)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   11790
      Style           =   1  '�׷���
      TabIndex        =   26
      Tag             =   "0"
      Top             =   8535
      Width           =   1260
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00E0E0E0&
      Caption         =   "�� ��(&X)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   13095
      Style           =   1  '�׷���
      TabIndex        =   25
      Tag             =   "0"
      Top             =   8535
      Width           =   1260
   End
   Begin VB.CheckBox chkCollect 
      BackColor       =   &H00808080&
      Caption         =   "ä����� �˻�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E8F7F7&
      Height          =   225
      Left            =   2685
      TabIndex        =   0
      Top             =   120
      Width           =   1620
   End
   Begin MedControls1.LisLabel LisLabel5 
      Height          =   285
      Left            =   4425
      TabIndex        =   1
      Top             =   75
      Width           =   9960
      _ExtentX        =   17568
      _ExtentY        =   503
      BackColor       =   8421504
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "ȯ�� �⺻����"
      LeftGab         =   100
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   285
      Left            =   135
      TabIndex        =   2
      Top             =   75
      Width           =   4200
      _ExtentX        =   7408
      _ExtentY        =   503
      BackColor       =   8421504
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "ȯ�ڰ˻�"
      LeftGab         =   100
   End
   Begin MSComctlLib.ListView lvwPtList 
      Height          =   8115
      Left            =   120
      TabIndex        =   18
      Top             =   915
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   14314
      View            =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16775406
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Frame fraSearch 
      BackColor       =   &H00DBE6E6&
      Height          =   645
      Left            =   135
      TabIndex        =   3
      Tag             =   "136"
      Top             =   270
      Width           =   4215
      Begin VB.OptionButton optSort 
         BackColor       =   &H00DBE6E6&
         Caption         =   "&ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   1995
         TabIndex        =   6
         Tag             =   "15304"
         Top             =   300
         Width           =   510
      End
      Begin VB.OptionButton optSort 
         BackColor       =   &H00DBE6E6&
         Caption         =   "&Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   2505
         TabIndex        =   5
         Tag             =   "15305"
         Top             =   285
         Value           =   -1  'True
         Width           =   825
      End
      Begin VB.TextBox txtSearchKey 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         MaxLength       =   10
         TabIndex        =   4
         Top             =   240
         Width           =   1830
      End
      Begin VB.Label lblReset 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         BackStyle       =   0  '����
         Caption         =   "Reset"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3570
         MouseIcon       =   "Lis154_APS.frx":0EB0
         MousePointer    =   99  '����� ����
         TabIndex        =   7
         Top             =   285
         Width           =   495
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  '�������� ����
         BorderColor     =   &H00808080&
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  '�ܻ�
         Height          =   285
         Index           =   1
         Left            =   3465
         Shape           =   4  '�ձ� �簢��
         Top             =   255
         Width           =   675
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DBE6E6&
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2265
      Left            =   4425
      TabIndex        =   8
      Top             =   285
      Width           =   9975
      Begin VB.TextBox txtPtId 
         Alignment       =   2  '��� ����
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   990
         MaxLength       =   10
         TabIndex        =   10
         Top             =   315
         Width           =   1425
      End
      Begin VB.TextBox txtMesg 
         BackColor       =   &H00F7FDF8&
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   2190
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  '����
         TabIndex        =   9
         ToolTipText     =   "�˻� ����ũ�� �Է��ϼ���."
         Top             =   1230
         Width           =   7365
      End
      Begin MedControls1.LisLabel LisLabel1 
         Height          =   285
         Left            =   630
         TabIndex        =   11
         Top             =   1290
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   503
         BackColor       =   15728622
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   "�� Remark"
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   195
         Index           =   0
         Left            =   3180
         TabIndex        =   12
         Top             =   360
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   344
         BackColor       =   14411494
         ForeColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   "��     ��"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel3 
         Height          =   255
         Index           =   5
         Left            =   6120
         TabIndex        =   13
         Top             =   360
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   450
         BackColor       =   14411494
         ForeColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   "�� / ����"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   344
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   "ȯ�� ID"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   195
         Index           =   2
         Left            =   3180
         TabIndex        =   15
         Top             =   780
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   344
         BackColor       =   14411494
         ForeColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   "�� �� ��"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel3 
         Height          =   255
         Index           =   7
         Left            =   6120
         TabIndex        =   16
         Top             =   780
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   450
         BackColor       =   14411494
         ForeColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   "��      ��"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   17
         Top             =   765
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   344
         BackColor       =   14411494
         ForeColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   "ó �� ��"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblPtNm 
         Height          =   300
         Left            =   4050
         TabIndex        =   20
         Top             =   300
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   529
         BackColor       =   15662589
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   "��̰�"
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblSexAge 
         Height          =   300
         Left            =   7065
         TabIndex        =   21
         Top             =   345
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         BackColor       =   15662589
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   "��̰�"
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblDoctNm 
         Height          =   300
         Left            =   1005
         TabIndex        =   22
         Top             =   720
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   529
         BackColor       =   15662589
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   "��̰�"
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblDeptNm 
         Height          =   300
         Left            =   4065
         TabIndex        =   23
         Top             =   735
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   529
         BackColor       =   15662589
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   "��̰�"
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblLocation 
         Height          =   300
         Left            =   7095
         TabIndex        =   19
         Top             =   765
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   529
         BackColor       =   15662589
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   "��̰�"
         Appearance      =   0
         LeftGab         =   100
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00EFFFEE&
         BackStyle       =   1  '�������� ����
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Height          =   930
         Index           =   0
         Left            =   210
         Shape           =   4  '�ձ� �簢��
         Top             =   1170
         Width           =   9585
      End
   End
   Begin MSCommLib.MSComm msComm 
      Left            =   4425
      Top             =   8475
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
End
Attribute VB_Name = "frm154NurCol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890
Option Explicit

'-------------------------------------
'�غκ���/�������� ä�� ����
'-------------------------------------
'#Const AllowAPSResultCollect = True
'#Const AllowBBSResultCollect = True
'-------------------------------------

Private IsFirst As Boolean

Private OrdFg As Boolean

Private MyPatient As New clsPatient
Private MySql As New clsLISSqlCollection
Private objLISCollect As New clsLISCollectioin
Private SelAllFg As Boolean

Private mvarEmpId As String
Private mvarWardId As String
Private mvarDeptCd As String
Private mvarHosilID As String
Private mvarRoomID As String

Public PtFg As Boolean
Public MsgFg As Boolean

Private strBlgCd As String      '������ �ǹ� �ڵ�
Private strErBldCd As String    '�����ϰ�� �˻��� �ǹ��ڵ�
Private strGBldCd As String     '�Ϲ��ϰ�� �˻��� �ǹ��ڵ�
Private blnCleared As Boolean

Public Event LastFormUnload()



Private Const lngMaxRows = 19
Private Const lngRowHeight = 12



'EmpId
Public Property Let EmpId(ByVal vData As String)
    mvarEmpId = vData
End Property
Public Property Get EmpId() As String
    EmpId = mvarEmpId
End Property

'WardId
Public Property Let WardId(ByVal vData As String)
    mvarWardId = vData
End Property
Public Property Get WardId() As String
    WardId = mvarWardId
End Property
'DeptCd
Public Property Let DeptCd(ByVal vData As String)
    mvarDeptCd = vData
End Property
Public Property Get DeptCd() As String
    DeptCd = mvarDeptCd
End Property

'HosilId
Public Property Let HosilId(ByVal vData As String)
    mvarHosilID = vData
End Property
Public Property Get HosilId() As String
    HosilId = mvarHosilID
End Property

'RoomID
Public Property Let RoomId(ByVal vData As String)
    mvarRoomID = vData
End Property
Public Property Get RoomId() As String
    RoomId = mvarRoomID
End Property



Private Sub chkChangeColTm_Click()
    
    Dim blnValue As Boolean
    
    blnValue = IIf(chkChangeColTm.Value = 1, True, False)
    dtpColDtTm.Enabled = blnValue
    If chkChangeColTm.Value = 1 Then dtpColDtTm.SetFocus
    
End Sub

Private Sub chkCollect_Click()
    Call txtSearchKey_KeyPress(vbKeyReturn)
End Sub

Private Sub chkSelAll_Click()
   
    SelAllFg = True
    With tblOrdSheet
        .Col = 1: .Col2 = 1
        .Row = 1: .Row2 = .DataRowCnt
        .BlockMode = True
        .Value = chkSelAll.Value
        .BlockMode = False
    End With
    SelAllFg = False
   
End Sub


Private Sub cmdClear_Click()
    Call ClearRtn
    txtPtId.Text = ""
    On Error GoTo Err_Trap
    txtPtId.SetFocus
Err_Trap:
End Sub



Private Sub cmdExit_Click()
    Unload Me
    Set MyPatient = Nothing
    Set MySql = Nothing
    If IsLastForm Then RaiseEvent LastFormUnload
   
End Sub
Private Function CollectionTargetChk() As Boolean
    Dim ii As Integer
    
    With tblOrdSheet
        For ii = 1 To .DataRowCnt
            .Row = ii
            .Col = enCOLLIST.tcCHECK
            If .Value = 1 Then
                CollectionTargetChk = True
                Exit For
            End If
        Next
    End With
End Function
Private Sub tblordersheet()
    With tblOrdSheet
        .SortBy = SortByRow
        .SortKey(1) = enCOLLIST.tcORDDIV
        .SortKeyOrder(1) = SortKeyOrderAscending
        .Col = 1: .Col2 = .MaxCols
        .Row = 1: .Row2 = .MaxRows
        .Action = ActionSort
    End With
End Sub
'& ä�� Ŭ���� MyCollect �� �̿��Ͽ� �ش� ȯ�ڵ��� ó���� ä�������Ѵ�.
Private Sub cmdSave_Click()
   
    Dim APSColSuccess As Boolean
    Dim BBSColSuccess As Boolean
    Dim LISColSuccess As Boolean
    
    Dim iCheckOrder   As Integer
    Dim lngBarCnt     As Long
    Dim lngSelCnt     As Long
    Dim BarCount      As Long
    Dim SelCount      As Long
    
    Dim strColDt As String, strColTm As String, strColID As String
    Dim strBlgCd As String, strErBldCd As String, strGBldCd As String
 
    Dim objPrgBar As clsProgressBar
    Dim ii As Integer
    
    If CollectionTargetChk = False Then
       MsgBox "ä���� �׸��� �����ϼ���..", vbInformation, "�׸���"
       tblOrdSheet.SetFocus
       Exit Sub
    End If
    
    iCheckOrder = objLISCollect.CheckSameOrder(tblOrdSheet, 1)     '�ߺ�ó�� Check
    If iCheckOrder > 0 Then GoTo OrdCheck1
    
    MouseRunning
     
'    Set objPrgBar = New clsProgress
'    With objPrgBar
'        .Caption = "��ȣ��ä��"
'        .Msg = "���õ� �˻��׸� ���� ä��ó�����Դϴ�."
'        .max = tblOrdSheet.DataRowCnt + 100
'        .min = 0
'        .Mode = 1
'        .Visible = True
'    End With
'    DoEvents
    
    Set objPrgBar = New clsProgressBar
    With objPrgBar
        .SetMyForm Me
        .Choice = True
        .XPos = lblBar.Left + 5 'optCondition(1).Left + optCondition(1).Width + 20
        .YPos = lblBar.Top + 5 'optCondition(1).Top + optCondition(1).Height - 260
        .XWidth = lblBar.Width - 10 'fraWSHeader.Width - (optCondition(1).Width * 2)
'        .ForeColor = &H864B24
        .ForeColor = &HFA8B10       'DCM_LightBlue   '&H864B24
        .Appearance = aPlate
        .BorderStyle = bsNone
        .YHeight = lblBar.Height - 10 ' 260
        .Msg = "���õ� �˻��׸� ���� ä��ó�����Դϴ�."
        .max = 90
        .min = 0
        .Value = 10
        DoEvents
    End With

    DoEvents

        '----------------------------------------------------------
    '������ ������ ���ؼ� �������� �ҷ��� �����Ѵ�.(2001/06/08)
    '----------------------------------------------------------
    
    
    
    Call tblordersheet
    
    Dim objDic As New clsDictionary
    objDic.Clear
    objDic.FieldInialize "orddiv", "first,last,coldt,coltm"
    With tblOrdSheet
        For ii = 1 To .DataRowCnt
            .Row = ii: .Col = enCOLLIST.tcORDDIV
            Select Case .Value
                Case BBS_ORDDIV
                    If objDic.Exists(.Value) Then
                        objDic.KeyChange BBS_ORDDIV
                        objDic.Fields("last") = .Row
                    Else
                        .Col = enCOLLIST.tcREQDTTM
                        objDic.AddNew BBS_ORDDIV, .Row & COL_DIV & "" & COL_DIV & _
                                      Format(.Text, "yyyymmdd") & COL_DIV & Format(.Text, "HHmm")
                    End If
                Case APS_ORDDIV
                    If objDic.Exists(.Value) Then
                        objDic.KeyChange APS_ORDDIV
                        objDic.Fields("last") = .Row
                    Else
                        objDic.AddNew APS_ORDDIV, .Row & COL_DIV & "" & COL_DIV & "" & COL_DIV & ""
                    End If
                Case LIS_ORDDIV
                    If objDic.Exists(.Value) Then
                        objDic.KeyChange LIS_ORDDIV
                        objDic.Fields("last") = .Row
                    Else
                        objDic.AddNew LIS_ORDDIV, .Row & COL_DIV & "" & COL_DIV & "" & COL_DIV & ""
                    End If
            End Select
        Next
        objDic.MoveFirst
        Do Until objDic.EOF
            If objDic.Fields("last") = "" Then
                objDic.Fields("last") = objDic.Fields("first")
            End If
            objDic.MoveNext
        Loop
    End With
    With objDic
        .MoveFirst
        Do Until .EOF
            Select Case .Fields("orddiv")
                Case APS_ORDDIV: iCheckOrder = objLISCollect.ChkSpcnm(tblOrdSheet, .Fields("first"), .Fields("last"))
                Case LIS_ORDDIV: iCheckOrder = objLISCollect.ChkSpcnm(tblOrdSheet, .Fields("first"), .Fields("last"))
            End Select
            If iCheckOrder > 0 Then GoTo OrdCheck2
            .MoveNext
        Loop
    End With
    
    '-------------------------------------------------------------
    '�������� ä���� �����Ѵ�(���������� ������ü üũ�� �ʿ����)
    '-------------------------------------------------------------
    With objDic
        .MoveFirst
        BBSColSuccess = True: APSColSuccess = True: LISColSuccess = True
        Do Until .EOF
            Select Case .Fields("orddiv")
                Case BBS_ORDDIV: BBSColSuccess = CollectForBBS_NEW(.Fields("first"), .Fields("last"), _
                                                                    Format(dbconn.GetSysDate, "yyyymmdd"), _
                                                                    Format(dbconn.GetSysDate, "HHmmss"), objPrgBar)
                Case APS_ORDDIV: APSColSuccess = CollectForAPS_New(.Fields("first"), .Fields("last"), objPrgBar)
                Case LIS_ORDDIV: LISColSuccess = CollectForLIS_New(.Fields("first"), .Fields("last"), objPrgBar)
            End Select
            .MoveNext
        Loop
    End With
        
    '����Ÿ���̽��� ��¥/�ð����� System Date/Time�� ����...
    Date = Format(dbconn.GetSysDate, CS_DateLongFormat)
    Time = Format(dbconn.GetSysDate, CS_TimeLongFormat)
     
    SelCount = 0: lngBarCnt = 0
    
    '�غκ��� ä��
'    APSColSuccess = CollectForAPS(lngSelCnt, lngBarCnt, objPrgBar)
'    SelCount = SelCount + lngSelCnt: BarCount = BarCount + lngBarCnt
    '�������� ä��
'    BBSColSuccess = CollectForBBS(lngSelCnt, lngBarCnt, objPrgBar)
'    SelCount = SelCount + lngSelCnt: BarCount = BarCount + lngBarCnt
'    '�ӻ󺴸� ä��
'    LISColSuccess = CollectForLIS(lngSelCnt, lngBarCnt, objPrgBar)
'    SelCount = SelCount + lngSelCnt: BarCount = BarCount + lngBarCnt
    
'    objPrgBar.Visible = False
    
    If APSColSuccess And BBSColSuccess And LISColSuccess Then
        'Form Feed : 2001.2.8 �߰�(2002.1.28 ����)
'        Call BarcodeLabel_FormFeed

'        Dim objBar As New clsBarcode
'        Set objBar.MyDB = dbconn
'        Set objBar.TableInfo = New clsTables
'        Call objBar.Get_PortNo
'        objBar.Label_FormFeed
'        Set objBar = Nothing
    Else
        Set objPrgBar = Nothing
        MsgBox "ä��ó���� ������ �߻��߽��ϴ� !!" & vbCrLf & _
               "������Ͻ� �� ������ ��ӵǸ� ����Ƿ� �����ٶ��ϴ�.", _
               vbCritical, "����"
    End If
    
    MouseDefault
    Set objPrgBar = Nothing
    Set objDic = Nothing
ExitPos:
    Call cmdClear_Click
    txtPtId.SetFocus
    Set objDic = Nothing
    Exit Sub

OrdCheck1:
    tblOrdSheet.Row = iCheckOrder
    tblOrdSheet.Col = 1
    tblOrdSheet.Action = ActionActiveCell
    MsgBox "�ߺ�ó���Դϴ�. Ȯ���ϰ� �ٽ� ä���Ͻʽÿ�.", vbExclamation, "�ߺ�ó��"
    tblOrdSheet.SetFocus
    
    Exit Sub

OrdCheck2:
    tblOrdSheet.Row = iCheckOrder
    tblOrdSheet.Col = 1
    tblOrdSheet.Action = ActionActiveCell
    MsgBox "������ü ������ �����ϴ�. ����Ƿ� �����ϼ���.", vbInformation + vbOKOnly, "����"
''    MsgBox "������ü ������ �����ϴ�. ����Ƿ� �����ϼ���. (��" & ObjSysInfo.helpline & ")", vbCritical, "����"
    tblOrdSheet.SetFocus
    Set objDic = Nothing
    Exit Sub

End Sub

'** �غκ��� ä����ƾ
Private Function CollectForAPS_New(ByVal FRowCnt As Integer, ByVal LRowCnt As Integer, _
                                   ByRef objProgress As Object) As Boolean

    CollectForAPS_New = True
    
    Dim strColDt As String
    Dim strColTm As String
    Dim strOrgBuild As String
    Dim strBuildCd As String
    
    Dim tmpClipData As String
    Dim tmpTotData As Variant
    Dim tmpRowData As Variant
    
    Dim lngColCnt As Long
    Dim lngSelCnt As Long
    Dim I As Long
    
    Dim strApsData As String
    
    Dim objColSave As clsAPSCollectionSave
    
    strColDt = Format(dtpColDtTm, "yyyyMMdd")
    strColTm = Format(dtpColDtTm, "HHmmss")
    strOrgBuild = ObjSysInfo.BuildingCd
    strBuildCd = APS_BUILDCD
    
'    Call GetBarInfo("A")
    lngSelCnt = 0
    
    With tblOrdSheet
        .Row = FRowCnt: .Col = enCOLLIST.tcWARDID: mvarWardId = .Value
                        .Col = enCOLLIST.tcROOMID: mvarHosilID = .Value
                        .Col = enCOLLIST.tcDEPTCD: mvarDeptCd = .Value
        .Col = 1: .Col2 = .MaxCols
        .Row = FRowCnt: .Row2 = LRowCnt
                        
        .BlockMode = True
        tmpClipData = .ClipValue
        tmpTotData = Split(tmpClipData, vbCrLf)
        .BlockMode = False
        
        For I = 0 To UBound(tmpTotData) - 1
          
            tmpRowData = Split(tmpTotData(I), vbTab)
            If objProgress.max > objProgress.Value Then objProgress.Value = objProgress.Value + 1
            If tmpRowData(0) = 0 Then GoTo Skip       '���ÿ���
            'If tmpRowData(enCOLLIST.tcORDDIV - 1) <> APS_ORDDIV Then GoTo Skip   'ó�汸��
          
            '�غκ���-----------------------------------------------------------------------------
            'strSS�� ����
            '   1:ó������,2:ó���ȣ,3:ó��SEQ,4:��ü�ڵ�,5:����ID,6:������ڵ�
            '   7:ó����,8:��ü��,9:�˻��,10:����������,11:���޿���,12:���ڵ�������
            
            lngColCnt = lngColCnt + 1
          
            If lngColCnt > 1 Then strApsData = strApsData & LINE_DIV
            strApsData = strApsData & tmpRowData(enCOLLIST.tcORDDATE - 1) & COL_DIV _
                                    & tmpRowData(enCOLLIST.tcORDNUM - 1) & COL_DIV _
                                    & tmpRowData(enCOLLIST.tcORDSEQ - 1) & COL_DIV _
                                    & tmpRowData(enCOLLIST.tcSPCCD - 1) & COL_DIV _
                                    & tmpRowData(enCOLLIST.tcWARDID - 1) & COL_DIV _
                                    & tmpRowData(enCOLLIST.tcDEPTCD - 1) & COL_DIV _
                                    & tmpRowData(enCOLLIST.tcORDDOCT - 1) & COL_DIV _
                                    & tmpRowData(enCOLLIST.tcSPCABBR - 1) & COL_DIV _
                                    & tmpRowData(enCOLLIST.tcABBRNM - 1) & COL_DIV _
                                    & tmpRowData(enCOLLIST.tcFRZFG - 1) & COL_DIV _
                                    & tmpRowData(enCOLLIST.tcSTATFG - 1) & COL_DIV _
                                    & tmpRowData(enCOLLIST.tcBARCNT - 1)
          
Skip:
       Next
    
    End With
    
    If lngColCnt = 0 Then
        CollectForAPS_New = True
        Exit Function
    End If
    

    Set objColSave = New clsAPSCollectionSave
'    Set objColSave.Barcode = objAPSbarcode     '2001.1.30 kmk�߰�
    If objColSave.SaveOrderCollection(Save_Nurse_Collection, txtPtId.Text, _
                                      strColDt, strColTm, gEmpId, _
                                      strBuildCd, strOrgBuild, 1, strColDt, _
                                      strColTm, strApsData, , "N", mvarWardId, mvarHosilID, _
                                      MyPatient.PtNm, objProgress) = False Then
       'MsgBox "ä��ó���� ���� �ʾҽ��ϴ�. �����ڿ��� ���� �ϼ���!", vbExclamation, "��ȣ��ä�����"
       CollectForAPS_New = False
       Exit Function
    End If
         
    Set objColSave = Nothing
    
End Function



'** �������� ä����ƾ
Private Function CollectForBBS_NEW(ByVal FRowCnt As Integer, ByVal LRowCnt As Integer, _
                                   ByVal ColDt As String, ByVal ColTm As String, _
                                   ByRef objProgress As Object) As Boolean

    
    Dim dicBBS As New clsDictionary
    Dim objCollect  As New clsBBSCollection
    Dim objBar      As New clsDictionary
    
    Dim tmpClipData As String
    
    Dim tmpTotData  As Variant
    Dim tmpRowData  As Variant
    
    Dim strColDt As String      'ä����
    Dim strColTm As String      'ä���Ͻ�
    
    Dim I As Long
    Dim lngColCnt As Integer
    Dim strStatFg As String
    
    lngColCnt = 0
    HosilId = medGetP(lblLocation.Caption, 2, "-")
    
    With tblOrdSheet
        .Row = FRowCnt: .Col = enCOLLIST.tcWARDID: mvarWardId = .Value
                        .Col = enCOLLIST.tcROOMID: mvarHosilID = .Value
                        .Col = enCOLLIST.tcDEPTCD: mvarDeptCd = .Value
        .Col = 1: .Col2 = .MaxCols
        .Row = FRowCnt: .Row2 = LRowCnt
        .BlockMode = True
        tmpClipData = .ClipValue
        tmpTotData = Split(tmpClipData, vbCrLf)
        .BlockMode = False
'        If chkChangeColTm.Value = 1 Then
'            strColDt = Format(dtpColDtTm.Value, "yyyymmdd")
'            strColTm = Format(dtpColDtTm.Value, "hhmmss")
'        Else
            strColDt = ColDt
            strColTm = ColTm
'        End If
        
        .Col = 7: strStatFg = IIf(Trim(.Value) = "Y", "1", "0")
        
        For I = 0 To UBound(tmpTotData) - 1

            tmpRowData = Split(tmpTotData(I), vbTab)
            If objProgress.max > objProgress.Value Then objProgress.Value = objProgress.Value + 1
            If tmpRowData(0) = 0 Then GoTo Skip       '���ÿ���
          
            lngColCnt = lngColCnt + 1
            
            '��������-----------------------------------------------------------------------------
                
                dicBBS.Clear
                dicBBS.FieldInialize "ptid", "ptnm,coldt,coltm,colid,bussdiv,buildcd,hosilid,statfg"
                dicBBS.AddNew txtPtId.Text, Join(Array(lblPtNm.Caption, strColDt, strColTm, _
                              gEmpId, enBussDiv.BussDiv_InPatient, strBlgCd, mvarHosilID, strStatFg), COL_DIV)

Skip:
       Next
    
    End With
    
    If lngColCnt = 0 Then
        CollectForBBS_NEW = True
        Exit Function
    End If
          
    objCollect.WardId = mvarWardId
    CollectForBBS_NEW = objCollect.Set_Collect(dicBBS, , objProgress)
    
    If CollectForBBS_NEW Then
        Set objBar = objCollect.BldDic
        If objBar.RecordCount > 0 Then
        '���ڵ� ���
            'lngBarCnt = 1
            BarCodePrintForBBS objBar
        Else
            'lngBarCnt = 0
            Set objProgress = Nothing
            MsgBox "��ü�� �̹� �����ϹǷ� ���ڵ尡 ��µ��� �ʽ��ϴ�.", vbInformation + vbOKOnly, "���ڵ����"
        End If
        If objCollect.Spc72Chk Then
            MsgBox "�ش� ȯ�ڴ� 72�ð����� ä���� ��ü�� �����մϴ�.", vbInformation + vbOKOnly, "���ڵ����"
        End If
    End If
    
    Set objCollect = Nothing
    Set objBar = Nothing
    Set dicBBS = Nothing

End Function

Private Function CollectForLIS_New(ByVal FRowCnt As Long, _
                               ByVal LRowCnt As Long, _
                               ByRef objProgress As Object) As Boolean

    Dim tmpDate As String, tmpTime As String
    Dim tmpStatFg As String
    Dim SqlStmt As String
    Dim tmpRs As Object
    Dim tmpData() As String
    Dim ColSuccess As Boolean
    Dim I As Integer
    Dim SelCount As Integer
    
    Dim CollectCnt As Integer

    '����Ÿ���̽��� ��¥/�ð����� System Date/Time�� ����...
    Date = dbconn.GetSysDate
    Time = dbconn.GetSysDate

    CollectCnt = 0
    Call objLISCollect.InitRtn

    With tblOrdSheet

        ReDim tmpData(0 To 20)
        .Row = FRowCnt: .Col = enCOLLIST.tcWARDID: mvarWardId = .Value
                        .Col = enCOLLIST.tcROOMID: mvarHosilID = .Value
                        .Col = enCOLLIST.tcDEPTCD: mvarDeptCd = .Value
        For I = FRowCnt To LRowCnt
            
            If objProgress.max > objProgress.Value Then objProgress.Value = objProgress.Value + 1
            
            .Row = I
            
            .Col = enCOLLIST.tcCHECK
            If .Value <> 1 Then GoTo Skip

'            .Col = enCOLLIST.tcORDDIV
'            If .Value <> LIS_ORDDIV Then GoTo Skip
            
            CollectCnt = CollectCnt + 1
            .Col = enCOLLIST.tcBUILDCD:  tmpData(0) = .Value        'Delivery Location
            .Col = enCOLLIST.tcWORKAREA: tmpData(1) = .Value        'WorkArea
            .Col = enCOLLIST.tcSPCCD:    tmpData(2) = .Value        'SpcCd
            .Col = enCOLLIST.tcSTORECD:  tmpData(3) = .Value        'StoreCd
            .Col = enCOLLIST.tcSTATFLAG: tmpData(4) = .Value        'StatFg
            .Col = enCOLLIST.tcREQDTTM:  tmpData(5) = .Value        'ReqColDate

            .Col = enCOLLIST.tcTESTDIV:  tmpData(6) = .Value        'TestDiv
            .Col = enCOLLIST.tcMULTIFG:  tmpData(7) = .Value        'MultiFg
            .Col = enCOLLIST.tcSPCGRP:   tmpData(8) = .Value        'SpcGrp
            .Col = enCOLLIST.tcORDDATE:  tmpData(9) = .Value        'OrdDt
            .Col = enCOLLIST.tcORDNUM:   tmpData(10) = .Value       'OrdNo
            .Col = enCOLLIST.tcORDSEQ:   tmpData(11) = .Value       'OrdSeq
            .Col = enCOLLIST.tcTESTCD:   tmpData(12) = .Value       'OrdCd
            .Col = enCOLLIST.tcDEPTCD:   tmpData(13) = .Value       '�����
            .Col = enCOLLIST.tcORDDOCT:  tmpData(14) = .Value       'ó����
            .Col = enCOLLIST.tcMAJDODT:  tmpData(15) = .Value       '��ġ��
            .Col = enCOLLIST.tcABBRNM:   tmpData(16) = .Value       '�˻� ����
            .Col = enCOLLIST.tcBARCNT:   tmpData(17) = .Value       '��������
            .Col = enCOLLIST.tcLABDIV:   tmpData(18) = .Value       'LabDiv
            .Col = enCOLLIST.tcSPCABBR:  tmpData(19) = .Value       '��ü����
            .Col = enCOLLIST.tcLABRANGE: tmpData(20) = .Value       '�̻���������ȣ����
            
            Call objLISCollect.AddLabCollect(tmpData)
Skip:
        Next
    End With

    If CollectCnt = 0 Then
        CollectForLIS_New = True
        Exit Function
    End If

    With objLISCollect

        ReDim tmpData(0 To 16)

        tmpData(0) = Mid(Format(dbconn.GetSysDate, "YYYY"), 4)  '��ü�⵵
        tmpData(1) = MyPatient.PtId                            'ȯ��ID
        tmpData(2) = MyPatient.PtNm
        tmpData(3) = MyPatient.Sex                             '����
        If IsDate(Format(MyPatient.Dob, CS_DateLongMask)) Then                         'ȯ���Ϸ�
            tmpData(4) = DateDiff("y", Format(MyPatient.Dob, CS_DateLongMask), dbconn.GetSysDate)
        Else
            tmpData(4) = Mid(MyPatient.Dob, 1, 4) & "-01-01"
            If IsDate(tmpData(4)) Then
                tmpData(4) = DateDiff("y", tmpData(4), dbconn.GetSysDate)
            Else
                tmpData(4) = 0
            End If
        End If
        tmpData(5) = MyPatient.BedInDt                           '�Կ���
        tmpData(6) = Format(dbconn.GetSysDate, CS_DateDbFormat)  '�Է���
        tmpData(7) = Format(dbconn.GetSysDate, CS_TimeDbFormat)  '�Է½ð�
        tmpData(8) = gEmpId                                      '�Է���
        tmpData(9) = ""                                          '��������ȣ
        tmpData(10) = Format(dbconn.GetSysDate, CS_DateDbFormat) 'ä����
        tmpData(11) = gEmpId                                     'ä����
        tmpData(12) = mvarWardId                                 '����ID
        tmpData(13) = mvarHosilID                                '����ID
        tmpData(14) = ""                                         'ħ��ID
        tmpData(15) = ""                                         'ħ��ID
        tmpData(16) = ObjSysInfo.BuildingCd                                '** ä���� ����Ǵ� �ǹ��ڵ�

        Call .SetColData(tmpData)
    End With

'    Call GetBarInfo(LIS_ORDDIV)
'    Set objLISCollect.Barcode = objLISbarcode      '2001.1.30 kmk�߰�
    
'
    ' ä�� ����
    ColSuccess = objLISCollect.DoCollection(objProgress)
    If Not ColSuccess Then
        Set objProgress = Nothing
        MsgBox "ä��ó���� ������ �߻��߽��ϴ� !!"
        MouseDefault  '0
        CollectForLIS_New = False
        Exit Function
    End If

    'lgnBarCnt = objLISCollect.ColCount
    CollectForLIS_New = True

    
End Function






Private Sub BarCodePrintForBBS(objDic As clsDictionary)

 
    'Dim objSql As New clsGetSqlStatement
    Dim objSql As New clsBBSCollection
    Dim objBar As New clsBarcode
    Dim strPtid As String
    Dim strPtnm As String
    Dim strColDt As String
    Dim strColTm As String
    Dim strSpcNo As String
    Dim strW_Dept As String
    Dim strBuildNm As String        '�ǹ��̸�
    Dim strAccSeq As String         'SpcYy-SpcNo ������ ��ü��ȣ
    Dim strHosilid As String
    Dim strStatFg  As String
    
    Set objBar.MyDB = dbconn
    Set objBar.TableInfo = New clsTables
    
    strW_Dept = mvarWardId
    If strW_Dept = "" Then
        strW_Dept = mvarDeptCd
    End If
    
    If lblLocation.Caption <> "" Then
        If lblLocation.Caption <> "--" Then strW_Dept = strW_Dept & "/" & mvarHosilID
    End If
    
    'strBuildNm = objSql.TestBldNm(strBlgCd)
    If P_ApplyBuildingInfo Then
        strBuildNm = ObjSysInfo.BuildingNm
    Else
        strBuildNm = "����"
    End If
    
    objDic.MoveFirst
    Do Until objDic.EOF
        strPtid = medGetP(objDic.GetString, 1, COL_DIV)
        strPtnm = medGetP(objDic.GetString, 2, COL_DIV)
        strSpcNo = medGetP(objDic.GetString, 3, COL_DIV)
        strColDt = Mid(medGetP(objDic.GetString, 4, COL_DIV), 1, 4)
        strColTm = Mid(medGetP(objDic.GetString, 5, COL_DIV), 1, 4)
        strStatFg = medGetP(objDic.GetString, 7, COL_DIV)
        strColTm = Format(strColTm, "0#:##")
        
        '��ü��ȣ ��� : 2001.2.8 �߰�
        strAccSeq = Mid(strSpcNo, 1, 2) & "-" & Format(Mid(strSpcNo, 3), "########0")
        strAccSeq = Format(strAccSeq, String(11, "@"))
        '
        objBar.Label_PrintOut strBuildNm, "XM", "", strAccSeq, strSpcNo, strPtid, _
                                            strPtnm, "", "", strStatFg, strW_Dept, strColDt, strColTm, _
                                            "", 1
        objDic.MoveNext
    Loop
    Set objBar = Nothing
    Set objSql = Nothing

End Sub

Private Sub cmdWardHelp_Click()

    Dim objDeptHelp As New clsS2DLP
    
    lvwPtList.ListItems.Clear
    
    With objDeptHelp
        .Caption = "��������Ʈ"
        .HeadName = "����,������"
        .ListPop , 2000, 1500, ObjLISComCode.WardId
        
        mvarWardId = medGetP(.SelectedString, 1, ";")
        lblWardId.Caption = mvarWardId
        If Trim(mvarWardId) <> "" Then
            chkCollect.Enabled = True
            chkCollect.Value = 0
        Else
            chkCollect.Enabled = False
            lblWardId.Caption = "��������"
        End If
    End With
    Set objDeptHelp = Nothing

End Sub

Private Sub dtpColDtTm_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then tblOrdSheet.SetFocus
End Sub

Private Sub dtpColDtTm_LostFocus()

    Dim Resp As VbMsgBoxResult
    If Format(dtpColDtTm.Value, "YYYYMMDD HH:MM") < Format(Now, "YYYYMMDD HH:MM") Then
        Resp = MsgBox("ä���ð��� ����ð����� �����Դϴ�. �����Ͻðڽ��ϱ�?", _
               vbQuestion + vbYesNo, "ä���ð�����")
        If Resp = vbNo Then
            dtpColDtTm.Value = Format(dbconn.GetSysDate, "YY-MM-DD HH:MM")
        End If
        chkChangeColTm.Value = 0
    End If
    
End Sub

Private Sub Form_Activate()
'    medMain.lblSubMenu.Caption = Me.Caption

    If Not IsFirst Then Exit Sub
    IsFirst = False
    
    
    If P_IncludeAPSSystem Or P_IncludeBBSSystem Then
        picOrdDiv.Visible = True
    Else
        picOrdDiv.Visible = False
    End If
    
    'Me.Show
    medInitLvwHead lvwPtList, "ȯ��ID,ȯ�ڼ���,�ֹε�Ϲ�ȣ,�������,����/����", _
                       "50,50,800,300,100"
    txtSearchKey.Text = ""
    Call ClearRtn
    If Trim(mvarWardId) <> "" Then
        lblWardId.Caption = Trim(mvarWardId)
        chkCollect.Enabled = True
    Else
        lblWardId.Caption = "��������"
        chkCollect.Enabled = False
    End If
    
On Error GoTo Err_Trap
    txtPtId.Text = ""
    txtPtId.SetFocus
    SelAllFg = False
    PtFg = False
    MsgFg = False
    optSort(1).Value = True

Err_Trap: End Sub

Private Sub Form_Load()
    IsFirst = True

End Sub

Private Sub lblReset_Click()
    lvwPtList.ListItems.Clear
    txtSearchKey.Text = ""
End Sub



Private Sub lvwPtList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static lngOrder As Long
    With lvwPtList
        lngOrder = (lngOrder + 1) Mod 2
        .SortKey = ColumnHeader.Index - 1
        .SortOrder = Choose(lngOrder + 1, lvwAscending, lvwDescending)
        .Sorted = True
    End With
End Sub

Private Sub lvwPtList_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    'ȯ������ Display
    If Item = "" Then Exit Sub
    DoEvents
    With Item
        txtPtId.Text = MedGetPtid(.Text)                'ȯ��ID
        Call txtPtId_KeyPress(vbKeyReturn)
    End With
    
End Sub

Private Sub tblOrdSheet_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

    Dim I As Integer
    Dim ButtonValue As Variant
    Dim SvOrdDt As String
    Dim SvOrdNo As String

    If SelAllFg Then Exit Sub
    
    With tblOrdSheet
       .Row = Row
       .Col = Col:   ButtonValue = .Value
       
       If .Value = 0 Then Exit Sub
       
       .Col = 9:      SvOrdDt = .Value
       .Col = 10:    SvOrdNo = .Value
       
       For I = 1 To .MaxRows
          If I <> Row Then
             .Row = I
             .Col = 9
             If .Value = SvOrdDt Then
                .Col = 10
                If .Value = SvOrdNo Then
                   .Col = 1
                   If .Value <> ButtonValue Then .Value = ButtonValue
                End If
             End If
          End If
       Next
    End With

End Sub

Private Sub txtPtId_LostFocus()
    
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
On Error GoTo Err_Trap

    If Screen.ActiveForm.ActiveControl.Name = cmdClear.Name Then Exit Sub
    If Screen.ActiveForm.ActiveControl.Name = cmdExit.Name Then Exit Sub
    If Screen.ActiveForm.ActiveControl.Name = txtSearchKey.Name Then Exit Sub
    If Screen.ActiveForm.ActiveControl.Name = lvwPtList.Name Then Exit Sub
    If Screen.ActiveForm.ActiveControl.Name = optSort(0).Name Then Exit Sub
    
    If blnCleared Then Call txtPtId_KeyPress(vbKeyReturn)
    Exit Sub

Err_Trap:
    Resume Next
    
End Sub

Private Sub txtSearchKey_Change()
    If optSort(0).Value = True Then
        Dim lngLen As Long
        
        If PROJECT_HOSCD = "04" Then
            With txtSearchKey
                lngLen = Len(Trim(.Text))
                If lngLen = 2 Then
                    .Text = .Text & "-"
                    .SelStart = Len(.Text)
                End If
            End With
        End If
    End If
End Sub

Private Sub txtSearchKey_GotFocus()

    With txtSearchKey
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub

'% ȯ��ID �Ǵ� �������� �˻� ����Ʈ �ۼ�.
Private Sub txtSearchKey_KeyPress(KeyAscii As Integer)
    
    Dim objPtInfo As New clsHosComSQLStmt
    Dim DrRs As New DrRecordSet
    Dim itmX As ListItem
    Dim lngSearch As Long
    
    If optSort(0).Value = True Then
        If PROJECT_HOSCD = "04" Then
            If Len(txtSearchKey) <> 2 Then
                If KeyAscii = vbKeyInsert Then KeyAscii = 0
            End If
            
            If KeyAscii = vbKeyBack Then
                With txtSearchKey
                    If .Text = "" Then Exit Sub
                    If Mid(.Text, Len(.Text)) = "-" Then
                        .Text = Mid(.Text, 1, Len(.Text) - 2)
                        .SelStart = Len(.Text)
                        KeyAscii = 0
                    End If
                End With
            End If
        End If
    End If
    
    If KeyAscii = vbKeyReturn Then
        lngSearch = IIf(optSort(0).Value, 1, 2) + 4 'True:ȯ��ID, False:ȯ�ڸ�
        
        If lngSearch = 1 And Not IsNumeric(MedSetPtid(txtSearchKey.Text)) Then Exit Sub
        
        If chkCollect.Value = 0 Then
            If txtSearchKey.Text = "" Then Exit Sub
            If optSort(0).Value = True Then
                DrRs.RsOpen , objPtInfo.SqlPtntSearch(lngSearch, MedSetPtid(txtSearchKey))
            Else
                DrRs.RsOpen , objPtInfo.SqlPtntSearch(lngSearch, txtSearchKey)
            End If
        Else
            If optSort(0).Value = True Then
                DrRs.RsOpen , objPtInfo.SqlPtntSearch(lngSearch, MedSetPtid(txtSearchKey), mvarWardId)
            Else
                DrRs.RsOpen , objPtInfo.SqlPtntSearch(lngSearch, txtSearchKey, mvarWardId)
            End If
        End If
        
        lvwPtList.ListItems.Clear
        If DrRs.EOF = False Then
            With lvwPtList
                Do Until DrRs.EOF
                    Set itmX = .ListItems.Add(, , MedGetPtid(DrRs.Fields("ptid").Value))
                    itmX.SubItems(1) = DrRs.Fields("ptnm").Value
                    itmX.SubItems(2) = DrRs.Fields("SSN").Value
                    itmX.SubItems(3) = Format(DrRs.Fields("DOB").Value, CS_DateLongMask)
                    itmX.SubItems(4) = IIf((Mid(DrRs.Fields("ssn").Value, 8, 1) Mod 2) = 1, "��", "��")
                    If IsDate(itmX.SubItems(3)) Then
                        itmX.SubItems(4) = itmX.SubItems(4) & " / " & DateDiff("yyyy", itmX.SubItems(3), dbconn.GetSysDate)
                    End If
                    If .ListItems.Count >= 1000 Then DrRs.MoveLast
                    DrRs.MoveNext
                Loop
            End With
        Else
            MsgBox "���ǿ� �´� �ڷᰡ �����ϴ�. Ȯ���� �˻��ϼ���", vbInformation + vbOKOnly, Me.Caption
        End If
        DrRs.RsClose:    Set DrRs = Nothing
    
    End If
    
    Set objPtInfo = Nothing
    
End Sub

'% ���� ���� ����
Private Sub optSort_Click(Index As Integer)
   If txtSearchKey.Text <> "" Then
      Call txtSearchKey_KeyPress(vbKeyReturn)
   End If
    txtSearchKey.SetFocus
End Sub

'% ȯ��ID�� ����Ǹ� ȭ��Clear
Private Sub txtPtId_Change()
    Dim lngLen As Long
    
    If Not blnCleared Then
       Call ClearRtn
    End If
    
    If PROJECT_HOSCD = "04" Then
        With txtPtId
            lngLen = Len(Trim(.Text))
            If lngLen = 2 Then
                .Text = .Text & "-"
                .SelStart = Len(.Text)
            End If
        End With
    End If
   
End Sub

'% ȯ�� ID
Private Sub txtPtId_GotFocus()
   With txtPtId
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
End Sub

'% ȯ������ �˻�

Private Sub txtPtId_KeyPress(KeyAscii As Integer)
    
    If Trim(txtPtId.Text) = "" Then Exit Sub
    
    
    If PROJECT_HOSCD = "04" Then
        
        If Len(txtPtId) <> 2 Then
            If KeyAscii = vbKeyInsert Then KeyAscii = 0
        End If
        
        If KeyAscii = vbKeyBack Then
            With txtPtId
                If .Text = "" Then Exit Sub
                If Mid(.Text, Len(.Text)) = "-" Then
                    .Text = Mid(.Text, 1, Len(.Text) - 2)
                    .SelStart = Len(.Text)
                    KeyAscii = 0
                End If
            End With
        End If
    End If
    
    
    If KeyAscii = vbKeyReturn Then

'On Error GoTo Err_Trap
        
        Select Case PROJECT_HOSCD
            Case "04": txtPtId = MedGetPtid(txtPtId)
            Case Else: If IsNumeric(txtPtId.Text) Then txtPtId.Text = Format(txtPtId.Text, P_PatientIdFormat)
        End Select
           
        If Not blnCleared Then Call ClearRtn
        DoEvents
        
'        With MyPatient
            Call MyPatient.ClearData   'Ŭ���� �� ���� �ʱ�ȭ
            If MyPatient.PtntQuery(MedSetPtid(txtPtId.Text)) Then
                lblPtNm.Caption = MyPatient.PtNm     '����
                lblSexAge.Caption = MyPatient.SexNm & " / " & MyPatient.Age & " " & MyPatient.AgeDiv      '����
                lblDeptNm.Caption = MyPatient.DeptNm '�����
                lblLocation.Caption = MyPatient.WardId & "-" & MyPatient.RoomId & "-" & MyPatient.BedID   '����
                DoEvents
                PtFg = True
                
                MouseRunning
                Call DisplayOrder
                MouseDefault
                
                cmdSave.Enabled = True
            Else
                txtPtId.Text = ""
                MsgFg = True
                MsgBox "��ϵ��� ���� ȯ��ID�Դϴ�.. �ٽ� �Է��ϼ���.."
                MsgFg = False
                txtPtId.SetFocus
                PtFg = False
                Call txtPtId_GotFocus
                Exit Sub
            End If
'        End With
        If OrdFg Then
            tblOrdSheet.SetFocus
        Else
            Call cmdClear_Click
            cmdSave.Enabled = False
            txtPtId.SetFocus
            Call txtPtId_GotFocus
        End If
        Exit Sub
            Resume Next
            chkSelAll.SetFocus
    End If
Err_Trap:

'Resume Next
End Sub

'% �˻��� ó���� ���̺� ���÷��� �Ѵ�.
Private Sub DisplayOrder()
   
    Dim I As Integer
    Dim SqlStmt As String
    Dim tmpRs As DrRecordSet
    Dim SvOrdDt As String, SvOrdNo As String
    Dim SvSpcNm As String, SvOrdDoct As String
    Dim tmpDate As String, tmpTime As String
    Dim tmpStatFg As String
    Dim tmpTestFg As String
    Dim strErChk As String
    Dim strOrdDiv As String
    
    Dim objProInSts As clsProgressBar
    'Dim objGetSql As New clsGetSqlStatement
    Dim objGetSql As New clsBBSCollection
   
On Error GoTo NoData

    '
    TestBuilding_Search     '�������� ��ü���� ���
   
    Set objProInSts = New clsProgressBar
    With objProInSts
        .SetMyForm Me
        .Choice = True
        .XPos = lblBar.Left + 5 'optCondition(1).Left + optCondition(1).Width + 20
        .YPos = lblBar.Top + 5 'optCondition(1).Top + optCondition(1).Height - 260
        .XWidth = lblBar.Width - 10 'fraWSHeader.Width - (optCondition(1).Width * 2)
'        .ForeColor = &H864B24
        .ForeColor = &HFA8B10       'DCM_LightBlue   '&H864B24
        .Appearance = aPlate
        .BorderStyle = bsNone
        .YHeight = lblBar.Height - 10 ' 260
        .Msg = "�ش�ȯ���� ó�� ������ �˻� ���Դϴ�...."
        .max = 90
        .min = 0
        .Value = 10
        DoEvents
    End With

    DoEvents
    txtMesg.Text = ""
    
    ' ó�泻�� �˻�
    tmpDate = Format(Now, CS_DateDbFormat)
    tmpTime = Format(Now, CS_TimeDbFormat)
    
    If gUsingInWardMenu Then
        strOrdDiv = "W"
    Else
        strOrdDiv = Mid(ObjSysInfo.projectid, 1, 1)
    End If
    SqlStmt = MySql.SqlReadWardOrder(MedSetPtid(txtPtId.Text), tmpDate, tmpTime, , enBussDiv.BussDiv_InPatient, , strOrdDiv)
    Set tmpRs = OpenRecordSet(SqlStmt)
    If tmpRs.EOF Then
        tmpRs.RsClose
        Set tmpRs = Nothing
        Set objProInSts = Nothing
       
        MsgBox MyPatient.PtNm & " ���� ó�泻���� �����ϴ�", vbInformation, "��ȣ�� ä��"
        If Not blnCleared Then Call ClearRtn
        Exit Sub
    End If
    
    With tblOrdSheet
       
        .ReDraw = False
        .MaxRows = 0
        If tmpRs.RecordCount < lngMaxRows Then
            .MaxRows = lngMaxRows
            .Row = tmpRs.RecordCount + 1
            .Row2 = lngMaxRows
            .Col = 1: .Col2 = .MaxCols
            .BlockMode = True
            .Lock = True
            .Protect = True
            .BlockMode = False
        Else
            .MaxRows = tmpRs.RecordCount   '����Ÿ �Ǽ�
        End If
       
        objProInSts.max = tmpRs.RecordCount
        
        'Locking Cells
        .Row = -1
        .Col = 2: .Col2 = .MaxCols
        .BlockMode = True
        .Lock = True
        .Protect = True
        .BlockMode = False
             
        For I = 1 To tmpRs.RecordCount

            objProInSts.Value = I

            .Row = I

            If SvOrdDt <> Trim("" & tmpRs.Fields("OrdDt").Value) Then
                .Col = enCOLLIST.tcORDDT:   .Text = Format("" & tmpRs.Fields("OrdDt").Value, CS_DateShortMask)    'ó����
                .Col = enCOLLIST.tcORDNO:   .Text = Trim("" & tmpRs.Fields("OrdNo").Value)     'ó���ȣ
                .Col = enCOLLIST.tcSPCNM:   .Text = Trim("" & tmpRs.Fields("SpcNm").Value)     '��ü
                .Col = enCOLLIST.tcDOCTNM:  .Text = Trim("" & tmpRs.Fields("DoctNm").Value)     'ó����
                SvOrdDt = Trim("" & tmpRs.Fields("OrdDt").Value)
                SvOrdNo = Trim("" & tmpRs.Fields("OrdNo").Value)    'ó���ȣ
                SvSpcNm = Trim("" & tmpRs.Fields("SpcNm").Value)    '��ü
                SvOrdDoct = Trim("" & tmpRs.Fields("DoctNm").Value) 'ó����
            End If
            If SvOrdNo <> Trim("" & tmpRs.Fields("OrdNo").Value) Then
                .Col = enCOLLIST.tcORDNO:   .Text = Trim("" & tmpRs.Fields("OrdNo").Value)     'ó���ȣ
                .Col = enCOLLIST.tcSPCNM:   .Text = Trim("" & tmpRs.Fields("SpcNm").Value)     '��ü
                .Col = enCOLLIST.tcDOCTNM:  .Text = Trim("" & tmpRs.Fields("DoctNm").Value)    'ó����
                SvOrdNo = Trim("" & tmpRs.Fields("OrdNo").Value)    'ó���ȣ
                SvSpcNm = Trim("" & tmpRs.Fields("SpcNm").Value)    '��ü
                SvOrdDoct = Trim("" & tmpRs.Fields("DoctNm").Value) 'ó����
            End If
            If SvSpcNm <> Trim("" & tmpRs.Fields("SpcNm").Value) Then
                .Col = enCOLLIST.tcSPCNM:   .Text = Trim("" & tmpRs.Fields("SpcNm").Value)     '��ü
                SvSpcNm = Trim("" & tmpRs.Fields("SpcNm").Value)
            End If
            If SvOrdDoct <> Trim("" & tmpRs.Fields("DoctNm").Value) Then
                .Col = enCOLLIST.tcDOCTNM: .Text = Trim("" & tmpRs.Fields("DoctNm").Value)    'ó����
                SvOrdDoct = Trim("" & tmpRs.Fields("DoctNm").Value)
            End If

            tmpStatFg = medGetP("" & tmpRs.Fields("StatFlags").Value, 1, ";")   '�ǹ��� ���ް��� ����
            tmpTestFg = medGetP("" & tmpRs.Fields("StatFlags").Value, 2, ";")   '�ǹ��� �˻簡�� ����
'
            Select Case tmpRs.Fields("orddiv")
            Case APS_ORDDIV:
                .Col = enCOLLIST.tcSTATFG:  .Text = Trim("" & tmpRs.Fields("StatFg").Value)      '���޿���  --> ������ ó��...
                .Col = enCOLLIST.tcBUILDCD: .Text = CentralLab
                .Col = enCOLLIST.tcBUILDNM: .Text = CentralLabNm
            
            Case BBS_ORDDIV:
                strErChk = objGetSql.ER_Chk(txtPtId.Text, SvOrdDt)
                .Col = enCOLLIST.tcSTATFG: .Value = Trim("" & tmpRs.Fields("StatFg").Value)     '���޿���  --> ������ ó��...
                .Col = enCOLLIST.tcBUILDCD: .Value = IIf(strErChk = "1", strErBldCd, strGBldCd)
                If ObjLISComCode.Building.Exists(.Value) Then
                    ObjLISComCode.Building.KeyChange (.Value)
                End If
                .Col = enCOLLIST.tcBUILDNM: .Value = ObjLISComCode.Building.Fields("buildnm")
            
            Case LIS_ORDDIV:

            '***�ǹ����� ���
                If P_ApplyBuildingInfo Then
    
                   If Trim(tmpRs.Fields("StatFg").Value) = "1" Then
    
                       '**���ް˻� ����
                       If Mid(tmpStatFg, ObjSysInfo.BuildingNo, 1) = "1" Then
    
                           '** �߾�/���̼��Ϳ��� ���ް˻簡 �߻��ϸ�.. --> ���޼��ͷ�...
                           If ObjSysInfo.BuildingCd = CentralLab Or _
                              ObjSysInfo.BuildingCd = AneLab Then
                               .Col = enCOLLIST.tcBUILDCD: .Text = EmergencyLab
                               .Col = enCOLLIST.tcBUILDNM: .Text = EmergencyLabNm
    
                           '** �ش�ǹ����� ���ް˻� ������
                           Else
                               .Col = enCOLLIST.tcBUILDCD: .Text = ObjSysInfo.BuildingCd
                               .Col = enCOLLIST.tcBUILDNM: .Text = ObjSysInfo.BuildingNm
                           End If
                           .Col = enCOLLIST.tcSTATFLAG: .Text = "1"       'StatFg
                           GoTo DataSet
                       Else
                       '*******************************************************************************************************
                       '** ����/���弾�� : ���ް˻簡 �������� ������� ���޽ǿ��� �˻簡 �����ϸ� ���޽Ƿ�, �ƴϸ� �߾�����...
                       '*******************************************************************************************************
                           '** ����/���弾�Ϳ��� ���ް˻簡 �߻��ϸ�..
                           If ObjSysInfo.BuildingCd = WomLab Or ObjSysInfo.BuildingCd = HrtLab Then
                               '** ���޽ǿ��� ���ް˻� ���� --> ���޼��ͷ�...
                               If Mid(tmpStatFg, EmergencyNo, 1) = "1" Then
                                   .Col = enCOLLIST.tcBUILDCD: .Text = EmergencyLab
                                   .Col = enCOLLIST.tcBUILDNM: .Text = EmergencyLabNm
                                   .Col = enCOLLIST.tcSTATFLAG:   .Text = "1"   'StatFg
                                   GoTo DataSet
                               End If
                           End If
                       '*******************************************************************************************************
                       End If
                   End If
    
                   .Col = enCOLLIST.tcSTATFLAG: .Text = "0"          'StatFg
    
                   '**�Ϲݰ˻簡��
                   If Mid(tmpTestFg, ObjSysInfo.BuildingNo, 1) = "1" Then
                       .Col = enCOLLIST.tcBUILDCD: .Text = ObjSysInfo.BuildingCd
                       .Col = enCOLLIST.tcBUILDNM: .Text = ObjSysInfo.BuildingNm
    
                   '**�Ϲݰ˻� �Ұ��� --> �߾Ӱ˻�Ƿ�...
                   Else
                       .Col = enCOLLIST.tcBUILDCD: .Text = CentralLab
                       .Col = enCOLLIST.tcBUILDNM: .Text = CentralLabNm
                   End If
    
            '***�ǹ����� ������� ����
                Else
                    .Col = enCOLLIST.tcBUILDCD:  .Text = ObjSysInfo.BuildingCd
                    .Col = enCOLLIST.tcBUILDNM:  .Text = ObjSysInfo.BuildingNm
                    .Col = enCOLLIST.tcSTATFLAG: .Text = Trim(tmpRs.Fields("StatFg").Value)
                End If
            
            End Select
          
DataSet:
            .Col = enCOLLIST.tcTESTNM:  .Text = Trim("" & tmpRs.Fields("TestNm").Value)     'ó���
                    Select Case tmpRs.Fields("orddiv")
                        Case APS_ORDDIV: .ForeColor = &H5E3F00     '&HDF6A3E     '&H00DF6A3E&�ణ �Ķ���
                        Case BBS_ORDDIV: .ForeColor = &H496835     '&H6C6181     '&H81815A     '�ణ���   &H00845584&�����
                        Case LIS_ORDDIV: .ForeColor = &H553755
                    End Select
            .Col = enCOLLIST.tcSTATFG:  .Text = IIf("" & tmpRs.Fields("StatFg").Value = "0", "", "Y") '���޿���
                                        .ForeColor = DCM_Red                                '������
            .Col = enCOLLIST.tcREQDTTM: .Text = Format("" & tmpRs.Fields("ReqDt").Value, CS_DateMask) & " " & _
                                         Format("" & tmpRs.Fields("ReqTm").Value, CS_TimeLongMask)      '���ä���Ͻ�
            .Col = enCOLLIST.tcORDDATE: .Text = Trim("" & tmpRs.Fields("OrdDt").Value)      'ó����
            .Col = enCOLLIST.tcORDNUM:  .Text = Trim("" & tmpRs.Fields("OrdNo").Value)      'ó���ȣ
            .Col = enCOLLIST.tcORDSEQ:  .Text = Trim("" & tmpRs.Fields("OrdSeq").Value)     'ó��Seq
            .Col = enCOLLIST.tcTESTCD:  .Text = Trim("" & tmpRs.Fields("OrdCd").Value)      '�˻��ڵ�

            Call ObjLISComCode.LisItem.KeyChange(.Text)
            .Col = enCOLLIST.tcLABDIV:  .Text = ObjLISComCode.LisItem.Fields("labdiv")      'LabDiv

            .Col = enCOLLIST.tcSPCCD:   .Text = Trim("" & tmpRs.Fields("SpcCd").Value)      '��ü�ڵ�

            Call ObjLISComCode.LisSpc.KeyChange(.Text)
            .Col = enCOLLIST.tcSPCABBR:  .Text = Trim("" & tmpRs.Fields("spcnm5").Value)         '��ü����
            .Col = enCOLLIST.tcLABRANGE: .Text = ObjLISComCode.LisSpc.Fields("labrange")    '�̻���������ȣ����

            .Col = enCOLLIST.tcWORKAREA: .Text = Trim("" & tmpRs.Fields("WorkArea").Value)  'WorkArea
            .Col = enCOLLIST.tcSTORECD:  .Text = Trim("" & tmpRs.Fields("StoreCd").Value)   '�����ڵ�
            .Col = enCOLLIST.tcTESTDIV:  .Text = Trim("" & tmpRs.Fields("TestDiv").Value)   '�˻籸��
            .Col = enCOLLIST.tcMULTIFG:  .Text = Trim("" & tmpRs.Fields("MultiFg").Value)   '������ü����
            .Col = enCOLLIST.tcSPCGRP:   .Text = Trim("" & tmpRs.Fields("SpcGrp").Value)    '��ü��
            .Col = enCOLLIST.tcORDDOCT:  .Text = Trim("" & tmpRs.Fields("OrdDoct").Value)   'ó����
                                         'ó���Ǹ�
                                         If .Text <> "" And lblDoctNm.Caption = "" Then
                                            lblDoctNm.Caption = Trim("" & tmpRs.Fields("DoctNm").Value)
                                         End If
            .Col = enCOLLIST.tcMAJDODT:  .Text = Trim("" & tmpRs.Fields("MajDoct").Value)   '��ġ��
            .Col = enCOLLIST.tcDEPTCD:   .Text = Trim("" & tmpRs.Fields("DeptCd").Value)    '�����
                                         '�������
                                         If .Text <> "" And lblDeptNm.Caption = "" Then
                                            If ObjLISComCode.DeptCd.Exists(.Text) Then
                                                ObjLISComCode.DeptCd.KeyChange (.Text)
                                                lblDeptNm.Caption = ObjLISComCode.DeptCd.Fields("deptnm")
                                            End If
                                         End If
            .Col = enCOLLIST.tcABBRNM:  .Text = Trim("" & tmpRs.Fields("AbbrNm5").Value)    '����
            .Col = enCOLLIST.tcBARCNT:  .Text = Trim("" & tmpRs.Fields("LabelCnt").Value)   '��������
            .Col = enCOLLIST.tcPAYDT:   .Text = Trim("" & tmpRs.Fields("ReceptNo").Value)   '��������ȣ
                                        .ForeColor = vbRed

            .Col = enCOLLIST.tcWARDID:  .Text = Trim("" & tmpRs.Fields("WardId").Value)     '����
                                        mvarWardId = .Text
            .Col = enCOLLIST.tcROOMID:  .Text = Trim("" & tmpRs.Fields("hosilid").Value)     '����
                                        mvarHosilID = .Text
            .Col = enCOLLIST.tcBEDID:   .Text = Trim("" & tmpRs.Fields("roomid").Value)      '����
                                        mvarRoomID = .Text

            .Col = enCOLLIST.tcFRZFG:   .Text = Trim("" & tmpRs.Fields("FzFg").Value)       '��������
            .Col = enCOLLIST.tcORDDIV:  .Text = Trim("" & tmpRs.Fields("OrdDiv").Value)     'ó�汸��
            
            If mvarWardId <> "" Then
                lblLocation.Caption = mvarWardId & "-" & mvarHosilID & "-" & mvarRoomID
            End If

            '����μ� Remark
            If Trim("" & tmpRs.Fields("Mesg").Value) <> "" Then
                txtMesg.Text = txtMesg.Text & "# " & Format(Trim("" & tmpRs.Fields("OrdNo").Value), "##") & " - "
                txtMesg.Text = txtMesg.Text & Trim("" & tmpRs.Fields("TestNm").Value) & vbCrLf
                txtMesg.Text = txtMesg.Text & Trim("" & tmpRs.Fields("Mesg").Value) & vbCrLf
            End If

            tmpRs.MoveNext
        Next

        .RowHeight(-1) = lngRowHeight
        .ReDraw = True
       
    End With
    OrdFg = True
    fraOrder.Enabled = True
    blnCleared = False
    Set objProInSts = Nothing
    
NoData:
    tmpRs.RsClose
    Set tmpRs = Nothing
   
End Sub

Private Sub TestBuilding_Search()
    
    'Dim objSql As New clsGetSqlStatement
    Dim objSql As New clsBBSCollection
    Dim strTmp As String
    
    With objSql
        If mvarWardId = "" Then
            strBlgCd = ObjSysInfo.BuildingCd
        Else
            strBlgCd = .Get_BuildingCd(UCase(mvarWardId))
        End If
        strTmp = .TestBuildCd(strBlgCd)
        strErBldCd = medGetP(strTmp, 1, COL_DIV)
        strGBldCd = medGetP(strTmp, 2, COL_DIV)
    End With
    
'    With tblCount
'        .Row = 1: .Col = 1: .Value = strErbldcd
'        objBBS901.TestBldNm dbConn, strErbldcd
'        .Row = 1: .Col = 2: .Value = objBBS901.ErBuildNM
'        .Row = 2: .Col = 1: .Value = strGbldcd
'        objBBS901.TestBldNm dbConn, strGbldcd
'        .Row = 2: .Col = 2: .Value = objBBS901.GBuildNM
'    End With
    Set objSql = Nothing
    
End Sub


Private Sub ClearRtn()
   
    lblPtNm.Caption = ""
    lblSexAge.Caption = ""
    lblDeptNm.Caption = ""
    lblLocation.Caption = ""
    lblDoctNm.Caption = ""
    txtMesg.Text = ""
    chkSelAll.Value = 0
    chkChangeColTm.Value = 0
    dtpColDtTm.Value = dbconn.GetSysDate
    dtpColDtTm.Enabled = False
    fraOrder.Enabled = False
    'optSort(0).Value = True
    With tblOrdSheet
        .Row = -1
        .Col = -1
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
    End With
    cmdSave.Enabled = False
    OrdFg = False
    PtFg = False
    MsgFg = False
    Set MyPatient = Nothing
    DoEvents
    
    Set objLISCollect = Nothing
    Set objLISCollect = New clsLISCollectioin
   
    Set MyPatient = New clsPatient
    Set MyPatient.objDB = dbconn
    DoEvents
   
    blnCleared = True
   
End Sub


Public Sub Call_PtId_KeyPress()

    Call txtPtId_KeyPress(vbKeyReturn)

End Sub

Public Function BarcodeLabel_FormFeed()

   Dim StrX As String
   
   Call ObjLISComCode.BarInfo.GetBarConfig
   
   StrX = ""
   StrX = StrX & "\1B@z" & vbCrLf
   StrX = StrX & "\1B@f09" & vbCrLf
   StrX = StrX & "\1Ba09" & ObjLISComCode.BarInfo.LabelLength & ObjLISComCode.BarInfo.LabelTotLength & vbCrLf
   StrX = StrX & "\1Bf09" & vbCrLf
   StrX = StrX & "\1Bq0001" & vbCrLf
   
   MSComm.CommPort = IIf(Val(ObjLISComCode.BarInfo.PortNo) <= 0, 1, Val(ObjLISComCode.BarInfo.PortNo))
   MSComm.Settings = "9600,N,8,1"
   MSComm.InputLen = 8192
   
   MSComm.PortOpen = True
   MSComm.Output = StrX
   MSComm.PortOpen = False

End Function

'------===================-----------==============================----------------------
'���� �Ʒ��� ������ �����.....(����,�ӻ�,�غ� ��� New�� ������ ���ο� ����� �������.)
'------===================-----------==============================----------------------


'** �������� ä����ƾ
Private Function CollectForBBS(ByRef lngSelCnt As Long, _
                               ByRef lngBarCnt As Long, _
                               ByRef objProgress As Object) As Boolean

'    CollectForBBS = True
'
'
'#If AllowBBSResultCollect Then
'
'    Dim dicBBS As New clsDictionary
'    'Dim objGathering As New clsGetSqlStatement
'
'    'Dim objCollect As New clsSpcAddPaper
'    Dim objCollect As New clsBBSCollection
'    Dim objBar As New clsDictionary
'
'    Dim tmpClipData As String
'    Dim tmpTotData As Variant
'    Dim tmpRowData As Variant
'
'    Dim strColDt As String      'ä����
'    Dim strColTm As String      'ä���Ͻ�
'
'    lngSelCnt = 0
'
'    With tblOrdSheet
'
'        .Col = 1: .Col2 = .MaxCols
'        .Row = 1: .Row2 = .DataRowCnt
'        .BlockMode = True
'        tmpClipData = .ClipValue
'        tmpTotData = Split(tmpClipData, vbCrLf)
'        .BlockMode = False
'
'        For I = 0 To UBound(tmpTotData) - 1
'
'            tmpRowData = Split(tmpTotData(I), vbTab)
'
'            If tmpRowData(0) = 0 Then GoTo Skip       '���ÿ���
'            If tmpRowData(enCOLLIST.tcORDDIV - 1) <> BBS_ORDDIV Then GoTo Skip   'ó�汸��
'
'            lngSelCnt = lngSelCnt + 1
'
'            '��������-----------------------------------------------------------------------------
'            If lngSelCnt <= 0 Then
'                With objCollect 'objGathering
'                    '����ID�� ������ �ǹ��ڵ带 �ҷ��´�..
'                    Dim strTestBuildCd As String
'                    'Call .setDbConn(dbconn)
'                    strBlgCd = .Get_BuildingCd(tmpRowData(enCOLLIST.tcWARDID - 1))         '���� �ǹ� �ڵ�
'                    strTestBuildCd = .TestBuildCd(strBlgCd)             '�ǹ��ڵ带 ������ �����˻縦 ������ �ǹ�.
'                    strErBldCd = medGetP(strTestBuildCd, 1, COL_DIV)    '���ް˻� �ǹ��ڵ�
'                    strGBldCd = medGetP(strTestBuildCd, 2, COL_DIV)     '�Ϲݰ˻� �ǹ��ڵ�
'                End With
'                dicBBS.Clear
'                dicBBS.FieldInialize "ptid", "ptnm,coldt,coltm,colid,bussdiv,buildcd"
'                dicBBS.AddNew txtPtId.Text, Join(Array(lblPtNm.Caption, strColDt, strColTm, _
'                              ObjSysInfo.EmpId, enBussDiv.BussDiv_InPatient, strBlgCd), COL_DIV)
'            End If
'            lngSelCnt = lngSelCnt + 1
'
'Skip:
'       Next
'
'    End With
'
'    If lngSelCnt = 0 Then
'        CollectForBBS = True
'        Exit Function
'    End If
'
'
'    CollectForBBS = objCollect.Set_Collect(dicBBS, , objProgress)
'
'    If CollectForBBS Then
'        Set objBar = objCollect.BldDic
'        If objBar.RecordCount > 0 Then
'        '���ڵ� ���
'            lngBarCnt = 1
'            BarCodePrintForBBS objBar
'        Else
'            lngBarCnt = 0
'            MsgBox "��ü�� �̹� �����ϹǷ� ���ڵ尡 ��µ��� �ʽ��ϴ�.", vbInformation + vbOKOnly, "���ڵ����"
'        End If
'    End If
'
'    Set objGathering = Nothing
'    Set objCollect = Nothing
'    Set objBar = Nothing
'    Set dicBBS = Nothing
'
'#End If

End Function


Private Function CollectForLIS(ByRef lngSelCnt As Long, _
                               ByRef lgnBarCnt As Long, _
                               ByRef objProgress As Object) As Boolean

'    Dim tmpDate As String, tmpTime As String
'    Dim tmpStatFg As String
'    Dim SqlStmt As String
'    Dim tmpRs As Object
'    Dim tmpData() As String
'    Dim ColSuccess As Boolean
'    Dim I As Integer
'    Dim SelCount As Integer
'
'    '����Ÿ���̽��� ��¥/�ð����� System Date/Time�� ����...
'    Date = dbconn.GetSysDate
'    Time = dbconn.GetSysDate
'
'    lngSelCnt = 0
'    Call objLISCollect.InitRtn
'
'    With tblOrdSheet
'
'        ReDim tmpData(0 To 20)
'
'        For I = 1 To .DataRowCnt
'
'            If objProgress.Max > objProgress.Value Then objProgress.Value = objProgress.Value + 1
'
'            .Row = I
'
'            .Col = enCOLLIST.tcCHECK
'            If .Value <> 1 Then GoTo Skip
'
'            .Col = enCOLLIST.tcORDDIV
'            If .Value <> LIS_ORDDIV Then GoTo Skip
'
'            lngSelCnt = lngSelCnt + 1
'            .Col = enCOLLIST.tcBUILDCD:  tmpData(0) = .Value        'Delivery Location
'            .Col = enCOLLIST.tcWORKAREA: tmpData(1) = .Value        'WorkArea
'            .Col = enCOLLIST.tcSPCCD:    tmpData(2) = .Value        'SpcCd
'            .Col = enCOLLIST.tcSTORECD:  tmpData(3) = .Value        'StoreCd
'            .Col = enCOLLIST.tcSTATFLAG: tmpData(4) = .Value        'StatFg
'            .Col = enCOLLIST.tcREQDTTM:  tmpData(5) = .Value        'ReqColDate
'
'            .Col = enCOLLIST.tcTESTDIV:  tmpData(6) = .Value        'TestDiv
'            .Col = enCOLLIST.tcMULTIFG:  tmpData(7) = .Value        'MultiFg
'            .Col = enCOLLIST.tcSPCGRP:   tmpData(8) = .Value        'SpcGrp
'            .Col = enCOLLIST.tcORDDATE:  tmpData(9) = .Value        'OrdDt
'            .Col = enCOLLIST.tcORDNUM:   tmpData(10) = .Value       'OrdNo
'            .Col = enCOLLIST.tcORDSEQ:   tmpData(11) = .Value       'OrdSeq
'            .Col = enCOLLIST.tcTESTCD:   tmpData(12) = .Value       'OrdCd
'            .Col = enCOLLIST.tcDEPTCD:   tmpData(13) = .Value       '�����
'            .Col = enCOLLIST.tcORDDOCT:  tmpData(14) = .Value       'ó����
'            .Col = enCOLLIST.tcMAJDODT:  tmpData(15) = .Value       '��ġ��
'            .Col = enCOLLIST.tcABBRNM:   tmpData(16) = .Value       '�˻� ����
'            .Col = enCOLLIST.tcBARCNT:   tmpData(17) = .Value       '��������
'            .Col = enCOLLIST.tcLABDIV:   tmpData(18) = .Value       'LabDiv
'            .Col = enCOLLIST.tcSPCABBR:  tmpData(19) = .Value       '��ü����
'            .Col = enCOLLIST.tcLABRANGE: tmpData(20) = .Value       '�̻���������ȣ����
'
'            Call objLISCollect.AddLabCollect(tmpData)
'Skip:
'        Next
'    End With
'
'    If lngSelCnt = 0 Then
'        CollectForLIS = True
'        Exit Function
'    End If
'
'    With objLISCollect
'
'        ReDim tmpData(0 To 16)
'
'        tmpData(0) = Mid(Format(dbconn.GetSysDate, "YYYY"), 4)  '��ü�⵵
'        tmpData(1) = MyPatient.PtId                            'ȯ��ID
'        tmpData(2) = MyPatient.PtNm
'        tmpData(3) = MyPatient.Sex                             '����
'        If IsDate(Format(MyPatient.Dob, CS_DateLongMask)) Then                         'ȯ���Ϸ�
'            tmpData(4) = DateDiff("y", Format(MyPatient.Dob, CS_DateLongMask), dbconn.GetSysDate)
'        Else
'            tmpData(4) = DateDiff("y", Mid(MyPatient.Dob, 1, 4) & "-01-01", dbconn.GetSysDate)
'        End If
'        tmpData(5) = MyPatient.BedInDt                         '�Կ���
'        tmpData(6) = Format(dbconn.GetSysDate, CS_DateDbFormat) '�Է���
'        tmpData(7) = Format(dbconn.GetSysDate, CS_TimeDbFormat) '�Է½ð�
'        tmpData(8) = ObjMyUser.EmpId                            '�Է���
'        tmpData(9) = ""                                         '��������ȣ
'        tmpData(10) = Format(dbconn.GetSysDate, CS_DateDbFormat) 'ä����
'        tmpData(11) = ObjMyUser.EmpId                           'ä����
'        tmpData(12) = medGetP(lblLocation.Caption, 1, "-")      '����ID
'        tmpData(13) = medGetP(lblLocation.Caption, 2, "-")      '����ID
'        tmpData(14) = ""                                        'ħ��ID
'        tmpData(15) = ""                                        'ħ��ID
'        tmpData(16) = ObjSysInfo.BuildingCd                                '** ä���� ����Ǵ� �ǹ��ڵ�
'
'        Call .SetColData(tmpData)
'    End With
'
''    Call GetBarInfo(LIS_ORDDIV)
''    Set objLISCollect.Barcode = objLISbarcode      '2001.1.30 kmk�߰�
'
''
'    ' ä�� ����
'    ColSuccess = objLISCollect.DoCollection(objProgress)
'    If Not ColSuccess Then
'        MsgBox "ä��ó���� ������ �߻��߽��ϴ� !!"
'        MouseDefault  '0
'        CollectForLIS = False
'        Exit Function
'    End If
'
'    lgnBarCnt = objLISCollect.ColCount
'    CollectForLIS = True

    
End Function




'** �غκ��� ä����ƾ
Private Function CollectForAPS(ByRef lngSelCnt As Long, ByRef lngBarCnt As Long, _
                               ByRef objProgress As Object) As Boolean

'    CollectForAPS = True
'
'#If AllowAPSResultCollect Then
'
'    Dim strColDt As String
'    Dim strColTm As String
'    Dim strOrgBuild As String
'    Dim strBuildCd As String
'
'    Dim tmpClipData As String
'    Dim tmpTotData As Variant
'    Dim tmpRowData As Variant
'
'    Dim objColSave As clsAPSCollectionSave
'
'    strColDt = Format(dtpColDtTm, "yyyyMMdd")
'    strColTm = Format(dtpColDtTm, "HHmmss")
'    strOrgBuild = ObjSysInfo.BuildingCd
'    strBuildCd = APS_BUILDCD
'
'    Call GetBarInfo("A")
'    lngSelCnt = 0
'
'    With tblOrdSheet
'
'        .Col = 1: .Col2 = .MaxCols
'        .Row = 1: .Row2 = .DataRowCnt
'        .BlockMode = True
'        tmpClipData = .ClipValue
'        tmpTotData = Split(tmpClipData, vbCrLf)
'        .BlockMode = False
'
'        For I = 0 To UBound(tmpTotData) - 1
'
'            tmpRowData = Split(tmpTotData(I), vbTab)
'
'            If tmpRowData(0) = 0 Then GoTo Skip       '���ÿ���
'            If tmpRowData(enCOLLIST.tcORDDIV - 1) <> APS_ORDDIV Then GoTo Skip   'ó�汸��
'
''            �غκ���-----------------------------------------------------------------------------
''            strSS�� ����
''1:               ó������ , 2: ó���ȣ , 3: ó��SEQ , 4: ��ü�ڵ� , 5: ����ID , 6: ������ڵ�
''7:               ó���� , 8: ��ü�� , 9: �˻�� , 10: ���������� , 11: ���޿��� , 12: ���ڵ�������
'
'            lngSelCnt = lngSelCnt + 1
'
'            If lngSelCnt > 1 Then strApsData = strApsData & LINE_DIV
'            strApsData = strApsData & tmpRowData(enCOLLIST.tcORDDATE - 1) & COL_DIV _
'                                    & tmpRowData(enCOLLIST.tcORDNUM - 1) & COL_DIV _
'                                    & tmpRowData(enCOLLIST.tcORDSEQ - 1) & COL_DIV _
'                                    & tmpRowData(enCOLLIST.tcSPCCD - 1) & COL_DIV _
'                                    & tmpRowData(enCOLLIST.tcWARDID - 1) & COL_DIV _
'                                    & tmpRowData(enCOLLIST.tcDEPTCD - 1) & COL_DIV _
'                                    & tmpRowData(enCOLLIST.tcORDDOCT - 1) & COL_DIV _
'                                    & tmpRowData(enCOLLIST.tcSPCABBR - 1) & COL_DIV _
'                                    & tmpRowData(enCOLLIST.tcABBRNM - 1) & COL_DIV _
'                                    & tmpRowData(enCOLLIST.tcFRZFG - 1) & COL_DIV _
'                                    & tmpRowData(enCOLLIST.tcSTATFLAG - 1) & COL_DIV _
'                                    & tmpRowData(enCOLLIST.tcBARCNT - 1)
'
'          'ó��Header(LAB101)�� �ִ� WardId, RoomId�� ����...
'          If tmpRowData(26) <> "" Then
'                MyPatient.WardId = tmpRowData(26)
'                MyPatient.RoomId = tmpRowData(27)
'          End If
'Skip:
'       Next
'
'    End With
'
'    If lngSelCnt = 0 Then
'        CollectForAPS = True
'        Exit Function
'    End If
'
'
'    Set objColSave = New clsAPSCollectionSave
'    Set objColSave.Barcode = objAPSbarcode     '2001.1.30 kmk�߰�
'    If objColSave.SaveOrderCollection(Save_Nurse_Collection, txtPtId.Text, _
'                                      strColDt, strColTm, ObjSysInfo.EmpId, _
'                                      strBuildCd, strOrgBuild, lngBarCnt, strColDt, _
'                                      strColTm, strApsData, , "N", mvarWardId, mvarHosilID, _
'                                      MyPatient.PtNm, objProgress) = False Then
'       MsgBox "ä��ó���� ���� �ʾҽ��ϴ�. �����ڿ��� ���� �ϼ���!", vbExclamation, "��ȣ��ä�����"
'       CollectForAPS = False
'       Exit Function
'    End If
'
'    Set objColSave = Nothing
'
'#End If
    
End Function



















































'------===================-----------------
'���� �Ʒ��� �������� �ּ�ó�� �Ǿ��ִ�����
'------===================-----------------










''123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890
'Private blnOrdFg As Boolean
'Private blnInitFg As Boolean
'
'Private MyPatient As New clsPatient
'Private objMySql As New clsLISSqlCollection
'Private objLISCollect As New clsLISCollectioin
'Private SelAllFg As Boolean
'
'Private mvarEmpId As String
'Private mvarWardId As String
'
'Public PtFg As Boolean
'Public MsgFg As Boolean
'
'Private strBlgCd As String      '������ �ǹ� �ڵ�
'Private strErBldCd As String    '�����ϰ�� �˻��� �ǹ��ڵ�
'Private strGBldCd As String     '�Ϲ��ϰ�� �˻��� �ǹ��ڵ�
'Private blnCleared As Boolean
'
'Private Const lngMaxRows = 20
'Private Const lngRowHeight = 12
'
'Public Event LastFormUnload()
'
'
''EmpId
'Public Property Let EmpId(ByVal vData As String)
'    mvarEmpId = vData
'End Property
'Public Property Get EmpId() As String
'    EmpId = mvarEmpId
'End Property
'
''WardId
'Public Property Let WardId(ByVal vData As String)
'    mvarWardId = vData
'End Property
'Public Property Get WardId() As String
'    WardId = mvarWardId
'End Property
'
'
'Private Sub chkChangeColTm_Click()
'
'    Dim blnValue As Boolean
'
'    blnValue = IIf(chkChangeColTm.Value = 1, True, False)
'    dtpColDtTm.Enabled = blnValue
'    If chkChangeColTm.Value = 1 Then dtpColDtTm.SetFocus
'
'End Sub
'
'Private Sub chkCollect_Click()
'    Call txtSearchKey_KeyPress(vbKeyReturn)
'End Sub
'
'Private Sub chkSelAll_Click()
'
'   SelAllFg = True
'   With tblOrdSheet
'      .Col = 1: .Col2 = 1
'      .Row = 1: .Row2 = .DataRowCnt
'      .BlockMode = True
'      .Value = chkSelAll.Value
'      .BlockMode = False
'   End With
'   SelAllFg = False
'
'End Sub
'
'
'Private Sub cmdClear_Click()
'   Call ClearRtn
'   txtPtId.Text = ""
'   On Error GoTo Err_Trap
'   txtPtId.SetFocus
'Err_Trap:
'End Sub
'
'
'Private Sub cmdExit_Click()
'    Unload Me
'    'Set frm154SendPt = Nothing
'    Set MyPatient = Nothing
'    Set objMySql = Nothing
'    If IsLastForm Then RaiseEvent LastFormUnload
'
'End Sub
'
''& ä�� Ŭ���� objLISCollect �� �̿��Ͽ� �ش� ȯ�ڵ��� ó���� ä�������Ѵ�.
'Private Sub cmdSave_Click()
'
'
'    Dim iCount As Integer
'    Dim iCheckOrder As Integer
'    Dim objProgress As New clsProgressBar
'
'    iCheckOrder = objLISCollect.CheckSameOrder(tblOrdSheet, 1)     '�ߺ�ó�� Check
'    If iCheckOrder > 0 Then GoTo OrdCheck1
'
'    iCheckOrder = objLISCollect.CheckSameOrder(tblOrdSheet, 2)     '������ü Check
'    If iCheckOrder > 0 Then GoTo OrdCheck2
'
'    MouseRunning  '13
'
'    iCount = 0
'
'    Set objProgress = New clsProgressBar
'    With objProgress
'        .SetMyForm Me
'        .Choice = True
'        .XPos = lblBar.Left + 5 'optCondition(1).Left + optCondition(1).Width + 20
'        .YPos = lblBar.Top + 5 'optCondition(1).Top + optCondition(1).Height - 260
'        .XWidth = lblBar.Width - 10 'fraWSHeader.Width - (optCondition(1).Width * 2)
''        .ForeColor = &H864B24
'        .ForeColor = &HFA8B10       'DCM_LightBlue   '&H864B24
'        .Appearance = aPlate
'        .BorderStyle = bsNone
'        .YHeight = lblBar.Height - 10 ' 260
'        .Msg = "Worksheet ����� �˻� ���Դϴ�...."
'        .Max = 90
'        .Min = 0
'        .Value = 10
'        DoEvents
'    End With
'
'    DoEvents
'
'    objProgress.Msg = "ä�� Procedure�� �����ϰ� �ֽ��ϴ�."
'    DoEvents
'
'    'ä����ƾ ����
'    If Not DoCollection(iCount, objProgress) Then
'        Set objProgress = Nothing
'        MouseDefault  '0
'        'Call ClearRtn
'        Exit Sub
'    End If
'
'    Set objProgress = Nothing
'    MouseDefault  '0
'
'    If iCount = 0 Then
'        MsgBox "ä���� �׸��� �����ϼ���..", vbInformation, "�ܷ�ä��"
'        tblOrdSheet.SetFocus
'        Exit Sub
'    ElseIf iCount > 0 Then
'        Dim strMsg As String
'
'        If P_UseBarcodeSystem Then ObjLISComCode.BarInfo.Label_FormFeed    '�� �ǵ�..
'
'        strMsg = "���������� ä���Ǿ����ϴ�."
'        If P_UseBarcodeSystem Then strMsg = strMsg & vbCrLf & "Barcode " & CStr(iCount) & " �� ����..." & vbCrLf
'
'        MsgBox strMsg, vbInformation, "ä���Ϸ�"
'
'        Call cmdClear_Click
'        txtPtId.SetFocus
'    End If
'
'    Exit Sub
'
'OrdCheck1:
'    tblOrdSheet.Row = iCheckOrder
'    tblOrdSheet.Col = 1
'    tblOrdSheet.Action = ActionActiveCell
'    MsgBox "�ߺ�ó���Դϴ�. Ȯ���ϰ� �ٽ� ä���Ͻʽÿ�.", vbExclamation, "�޼���"
'    tblOrdSheet.SetFocus
'    Exit Sub
'
'OrdCheck2:
'    tblOrdSheet.Row = iCheckOrder
'    tblOrdSheet.Col = 1
'    tblOrdSheet.Action = ActionActiveCell
'    MsgBox "������ü ������ �����ϴ�. ����Ƿ� �����ϼ���. (��" & ObjSysInfo.helpline & ")", vbCritical, "����"
'    tblOrdSheet.SetFocus
'    Exit Sub
'
'End Sub
'
'
'Private Function DoCollection(ByRef iCount As Integer, ByRef objProgress As Object) As Boolean
'
'    Dim tmpDate As String, tmpTime As String
'    Dim tmpStatFg As String
'    Dim SqlStmt As String
'    Dim tmpRs As Object
'    Dim tmpData() As String
'    Dim ColSuccess As Boolean
'    Dim i As Integer
'    Dim SelCount As Integer
'
'    '����Ÿ���̽��� ��¥/�ð����� System Date/Time�� ����...
'    Date = dbconn.GetSysDate
'    Time = dbconn.GetSysDate
'
'    SelCount = 0
'    Call objLISCollect.InitRtn
'
'    With tblOrdSheet
'
'        ReDim tmpData(0 To 20)
'
'        For i = 1 To .DataRowCnt
'            .Row = i
'            .Col = enCOLLIST.tcCHECK
'
'            If objProgress.Max > objProgress.Value Then objProgress.Value = objProgress.Value + 1
'            If .Value <> 1 Then GoTo Skip
'
'            SelCount = SelCount + 1
'            .Col = enCOLLIST.tcBUILDCD:  tmpData(0) = .Value        'Delivery Location
'            .Col = enCOLLIST.tcWORKAREA: tmpData(1) = .Value        'WorkArea
'            .Col = enCOLLIST.tcSPCCD:    tmpData(2) = .Value        'SpcCd
'            .Col = enCOLLIST.tcSTORECD:  tmpData(3) = .Value        'StoreCd
'            .Col = enCOLLIST.tcSTATFLAG: tmpData(4) = .Value        'StatFg
'            .Col = enCOLLIST.tcREQDTTM:  tmpData(5) = .Value        'ReqColDate
'
'            .Col = enCOLLIST.tcTESTDIV:  tmpData(6) = .Value        'TestDiv
'            .Col = enCOLLIST.tcMULTIFG:  tmpData(7) = .Value        'MultiFg
'            .Col = enCOLLIST.tcSPCGRP:   tmpData(8) = .Value        'SpcGrp
'            .Col = enCOLLIST.tcORDDATE:  tmpData(9) = .Value        'OrdDt
'            .Col = enCOLLIST.tcORDNUM:   tmpData(10) = .Value       'OrdNo
'            .Col = enCOLLIST.tcORDSEQ:   tmpData(11) = .Value       'OrdSeq
'            .Col = enCOLLIST.tcTESTCD:   tmpData(12) = .Value       'OrdCd
'            .Col = enCOLLIST.tcDEPTCD:   tmpData(13) = .Value       '�����
'            .Col = enCOLLIST.tcORDDOCT:  tmpData(14) = .Value       'ó����
'            .Col = enCOLLIST.tcMAJDODT:  tmpData(15) = .Value       '��ġ��
'            .Col = enCOLLIST.tcABBRNM:   tmpData(16) = .Value       '�˻� ����
'            .Col = enCOLLIST.tcBARCNT:   tmpData(17) = .Value       '��������
'            .Col = enCOLLIST.tcLABDIV:   tmpData(18) = .Value       'LabDiv
'            .Col = enCOLLIST.tcSPCABBR:  tmpData(19) = .Value       '��ü����
'            .Col = enCOLLIST.tcLABRANGE: tmpData(20) = .Value       '�̻���������ȣ����
'            Call objLISCollect.AddLabCollect(tmpData)
'Skip:
'        Next
'    End With
'
'    If SelCount = 0 Then
'        DoCollection = True
'        Exit Function
'    End If
'
'    iCount = iCount + SelCount
'    With objLISCollect
''        Call .SetDatabase(DbConn)
'
'        ReDim tmpData(0 To 16)
'
'        tmpData(0) = Mid(Format(dbconn.GetSysDate, "YYYY"), 4)  '��ü�⵵
'        tmpData(1) = MyPatient.PtId                            'ȯ��ID
'        tmpData(2) = MyPatient.PtNm
'        tmpData(3) = MyPatient.Sex                             '����
'        If IsDate(Format(MyPatient.Dob, CS_DateLongMask)) Then                         'ȯ���Ϸ�
'            tmpData(4) = DateDiff("y", Format(MyPatient.Dob, CS_DateLongMask), dbconn.GetSysDate)
'        Else
'            tmpData(4) = DateDiff("y", Mid(MyPatient.Dob, 1, 4) & "-01-01", dbconn.GetSysDate)
'        End If
'        tmpData(5) = MyPatient.BedInDt                         '�Կ���
'        tmpData(6) = Format(dbconn.GetSysDate, CS_DateDbFormat) '�Է���
'        tmpData(7) = Format(dbconn.GetSysDate, CS_TimeDbFormat) '�Է½ð�
'        tmpData(8) = ObjMyUser.EmpId                            '�Է���
'        tmpData(9) = ""                                         '��������ȣ
'        tmpData(10) = Format(dbconn.GetSysDate, CS_DateDbFormat) 'ä����
'        tmpData(11) = ObjMyUser.EmpId                           'ä����
'        tmpData(12) = medGetP(lblLocation.Caption, 1, "-")      '����ID
'        tmpData(13) = medGetP(lblLocation.Caption, 2, "-")      '����ID
'        tmpData(14) = ""                                        'ħ��ID
'        tmpData(15) = ""                                        'ħ��ID
'        tmpData(16) = ObjSysInfo.BuildingCd                                '** ä���� ����Ǵ� �ǹ��ڵ�
'
'        Call .SetColData(tmpData)
'    End With
'
'    ' ä�� ����
'    ColSuccess = objLISCollect.DoCollection(objProgress)
'    If Not ColSuccess Then
'        MsgBox "ä��ó���� ������ �߻��߽��ϴ� !!"
'        MouseDefault  '0
'        DoCollection = False
'        Exit Function
'    End If
'
'    iCount = objLISCollect.ColCount
'    DoCollection = True
'
'End Function
'
''% ä�������� ���ڵ�� ���...
'Sub Print_BarcodeLabel(Optional ByVal AccFg As Boolean = False)
'
'    Dim BarcodeBuffer As String
'    Dim i As Integer
'
'    lngLabelCnt = lngLabelCnt + objLISCollect.ColCount
'    For i = 1 To objLISCollect.ColCount
'        Call objLISCollect.GetBarcodeLabel(i, AccFg)
'    Next
'
'End Sub
'
'
'Private Sub dtpColDtTm_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then tblOrdSheet.SetFocus
'End Sub
'
'Private Sub dtpColDtTm_LostFocus()
'
'    Dim Resp As VbMsgBoxResult
'    If Format(dtpColDtTm.Value, "YYYYMMDD HH:MM") < Format(Now, "YYYYMMDD HH:MM") Then
'        Resp = MsgBox("ä���ð��� ����ð����� �����Դϴ�. �����Ͻðڽ��ϱ�?", _
'               vbQuestion + vbYesNo, "ä���ð�����")
'        If Resp = vbNo Then
'            dtpColDtTm.Value = Format(dbconn.GetSysDate, "YY-MM-DD HH:MM")
'        End If
'        chkChangeColTm.Value = 0
'    End If
'
'End Sub
'
'Private Sub Form_Activate()
''    medMain.lblSubMenu.Caption = Me.Caption
'    If blnInitFg Then Exit Sub
'    medInitLvwHead lvwPtList, "ȯ��ID,ȯ�ڼ���,�ֹε�Ϲ�ȣ,�������,����/����", _
'                       "50,50,800,300,100"
'    txtSearchKey.Text = ""
'    Call ClearRtn
'On Error GoTo Err_Trap
'    SelAllFg = False
'    PtFg = False
'    MsgFg = False
'    optSort(1).Value = True
'    txtPtId.Text = ""
'    txtPtId.SetFocus
'    blnInitFg = True
'Err_Trap:
'End Sub
'
'Private Sub Form_Load()
'    Me.Show
'    blnInitFg = False
'End Sub
'
'Private Sub lblReset_Click()
'    lvwPtList.ListItems.Clear
'    txtSearchKey.Text = ""
'End Sub
'
'
'Private Sub lvwPtList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'    Static lngOrder As Long
'    With lvwPtList
'        lngOrder = (lngOrder + 1) Mod 2
'        .SortKey = ColumnHeader.Index - 1
'        .SortOrder = Choose(lngOrder + 1, lvwAscending, lvwDescending)
'        .Sorted = True
'    End With
'End Sub
'
'Private Sub lvwPtList_ItemClick(ByVal Item As MSComctlLib.ListItem)
'
'    'ȯ������ Display
'    If Item = "" Then Exit Sub
'    DoEvents
'    With Item
'        txtPtId.Text = Trim(.Text)                'ȯ��ID
'        Call txtPtId_KeyPress(vbKeyReturn)
'    End With
'
'End Sub
'
'Private Sub tblOrdSheet_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
'
'   Dim i As Integer
'   Dim ButtonValue As Variant
'   Dim SvOrdDt As String
'   Dim SvIrdNo As String
'
'   If SelAllFg Then Exit Sub
'
'   With tblOrdSheet
'      .Row = Row
'      .Col = Col:   ButtonValue = .Value
'
'      If .Value = 0 Then Exit Sub
'
'      .Col = 9:      SvOrdDt = .Value
'      .Col = 10:    SvOrdNo = .Value
'
'      For i = 1 To .MaxRows
'         If i <> Row Then
'            .Row = i
'            .Col = 9
'            If .Value = SvOrdDt Then
'               .Col = 10
'               If .Value = SvOrdNo Then
'                  .Col = 1
'                  If .Value <> ButtonValue Then .Value = ButtonValue
'               End If
'            End If
'         End If
'      Next
'   End With
'
'End Sub
'
'Private Sub txtPtId_LostFocus()
'
'    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
'
'    If Screen.ActiveControl.Name = cmdClear.Name Then Exit Sub
'    If Screen.ActiveControl.Name = cmdExit.Name Then Exit Sub
'    If Screen.ActiveControl.Name = txtSearchKey.Name Then Exit Sub
'    If Screen.ActiveControl.Name = lvwPtList.Name Then Exit Sub
'    If Screen.ActiveControl.Name = optSort(0).Name Then Exit Sub
'
'    If blnCleared Then Call txtPtId_KeyPress(vbKeyReturn)
'
'End Sub
'
'Private Sub txtSearchKey_GotFocus()
'
'    With txtSearchKey
'        .SelStart = 0
'        .SelLength = Len(.Text)
'    End With
'
'End Sub
'
''% ȯ��ID �Ǵ� �������� �˻� ����Ʈ �ۼ�.
'Private Sub txtSearchKey_KeyPress(KeyAscii As Integer)
'
'    Dim objPtInfo As New clsHosComSQLStmt
'    Dim DrRS As New DrRecordSet
'    Dim itmx As ListItem
'    Dim lngSearch As Long
'
'    'Set objPtInfo.objDb = dbConn
'    If KeyAscii = vbKeyReturn Then
'        lngSearch = IIf(optSort(0).Value, 1, 2)  'True:ȯ��ID, False:ȯ�ڸ�
'        If lngSearch = 1 And Not IsNumeric(txtSearchKey.Text) Then Exit Sub
'        If chkCollect.Value = 0 Then
'            If txtSearchKey.Text = "" Then Exit Sub
'            DrRS.RsOpen , objPtInfo.SqlPtntSearch(lngSearch, txtSearchKey)
'        Else
'            If mvarWardId = "" And txtSearchKey.Text = "" Then Exit Sub
'            DrRS.RsOpen , objPtInfo.SqlPtntSearch(lngSearch, txtSearchKey, mvarWardId)
'        End If
'        lvwPtList.ListItems.Clear
'        If DrRS.EOF = False Then
'            With lvwPtList
'                Do Until DrRS.EOF
'                    Set itmx = .ListItems.Add(, , DrRS.Fields("ptid").Value)
'                    itmx.SubItems(1) = DrRS.Fields("ptnm").Value
'                    itmx.SubItems(2) = DrRS.Fields("SSN").Value
'                    itmx.SubItems(3) = Format(DrRS.Fields("DOB").Value, CS_DateLongMask)
'                    itmx.SubItems(4) = IIf((Mid(DrRS.Fields("ssn").Value, 8, 1) Mod 2) = 1, "��", "��")
'                    If IsDate(itmx.SubItems(3)) Then
'                        itmx.SubItems(4) = itmx.SubItems(4) & " / " & DateDiff("yyyy", itmx.SubItems(3), Now)
'                    Else
'                        itmx.SubItems(4) = itmx.SubItems(4) & " / ?"
'                    End If
'                    DrRS.MoveNext
'                Loop
'            End With
'        Else
'            MsgBox "���ǿ� �´� �ڷᰡ �����ϴ�. Ȯ���� �˻��ϼ���", vbInformation + vbOKOnly, Me.Caption
'        End If
'        DrRS.RsClose:    Set DrRS = Nothing
'
'    End If
'
'    Set objPtInfo = Nothing
'
'End Sub
'
''% ���� ���� ����
'Private Sub optSort_Click(Index As Integer)
'   If txtSearchKey.Text <> "" Then
'      Call txtSearchKey_KeyPress(vbKeyReturn)
'   End If
'    txtSearchKey.SetFocus
'End Sub
'
''% ȯ��ID�� ����Ǹ� ȭ��Clear
'Private Sub txtPtId_Change()
'   If Not blnCleared Then
'      Call ClearRtn
'   End If
'End Sub
'
''% ȯ�� ID
'Private Sub txtPtId_GotFocus()
'   With txtPtId
'      .SelStart = 0
'      .SelLength = Len(.Text)
'   End With
'End Sub
'
''% ȯ������ �˻�
'
'Private Sub txtPtId_KeyPress(KeyAscii As Integer)
'
'    If Trim(txtPtId.Text) = "" Then Exit Sub
'
'    If KeyAscii = vbKeyReturn Then
'
''On Error GoTo Err_Trap
'
'        If Not blnCleared Then Call ClearRtn
'        DoEvents
'
''        With MyPatient
'        Call MyPatient.ClearData   'Ŭ���� �� ���� �ʱ�ȭ
'        If MyPatient.PtntQuery(txtPtId.Text) Then
'            lblPtNm.Caption = MyPatient.PtNm     '����
'            lblSexAge.Caption = MyPatient.SexNm & " / " & MyPatient.Age & " " & MyPatient.AgeDiv      '����
'            'lblAge.Caption = .Age       '����
'            'lblAgeDiv.Caption = .AgeDiv '���̴���
'            lblDeptNm.Caption = MyPatient.DeptNm '�����
'            If Trim(MyPatient.WardId) <> vbNullString Then
'                lblLocation.Caption = MyPatient.WardId & "-" & MyPatient.RoomId & "-" & MyPatient.BedID   '����
'            End If
'            'lblBedinDt.Caption = Format(.BedInDt, CS_DateMask)
'            'lblBedoutDt.Caption = Format(.BedOutDt, CS_DateMask)
'            DoEvents
'            PtFg = True
'            MouseRunning
'            Call DisplayOrder
'            MouseDefault
'            cmdSave.Enabled = True
'        Else
'            txtPtId.Text = ""
'            MsgFg = True
'            MsgBox "��ϵ��� ���� ȯ��ID�Դϴ�.. �ٽ� �Է��ϼ���.."
'            MsgFg = False
'            txtPtId.SetFocus
'            PtFg = False
'            Call txtPtId_GotFocus
'            Exit Sub
'        End If
''        End With
'        If blnOrdFg Then
'            tblOrdSheet.SetFocus
'        Else
'            Call cmdClear_Click
'            cmdSave.Enabled = False
'            txtPtId.SetFocus
'            Call txtPtId_GotFocus
'        End If
'        blnCleared = False
'        Exit Sub
'
'    End If
'Err_Trap:
'
''Resume Next
'End Sub
'
''% �˻��� ó���� ���̺� ���÷��� �Ѵ�.
'Private Sub DisplayOrder()
'
'    Dim i As Integer
'    Dim SqlStmt As String
'    Dim tmpRs As DrRecordSet
'    Dim SvOrdDt As String, SvOrdNo As String
'    Dim SvSpcNm As String, SvOrdDoct As String
'    Dim tmpDate As String, tmpTime As String
'    Dim tmpStatFg As String
'    Dim tmpTestFg As String
'    Dim strWardId As String, strRoomId As String, strBedId As String
'    Dim objProInSts As New clsProgressBar
'
'    MouseRunning
'
'    With objProInSts
'        .SetStsBar Me
'        .Choice = True
'        .Appearance = aPlate
'        .XPos = lblBar.Left + 5
'        .YPos = lblBar.Top + 5
'        .XWidth = lblBar.Width - 10
'        .YHeight = lblBar.Height - 10
'        .Min = 0
'        .ForeColor = DCM_MidBlue
'    End With
'
'    ' ó�泻�� �˻�
'    tmpDate = Format(dbconn.GetSysDate, CS_DateDbFormat)
'    tmpTime = Format(dbconn.GetSysDate, CS_TimeDbFormat)
'
'    SqlStmt = objMySql.SqlReadOrder(txtPtId.Text, tmpDate, tmpTime, , BussDiv_InPatient, , LIS_ORDDIV)
'    'SqlStmt = objMySql.SqlReadOrder(txtPtId.Text, tmpDate, tmpTime, , BussDiv_OutPatient, txtReceptNo.Text)
'    Set tmpRs = OpenRecordSet(SqlStmt)
'    If tmpRs.EOF Then
'        MsgBox MyPatient.PtNm & " ���� ó�泻���� �����ϴ�"
'        MouseDefault
'        blnOrdFg = False
'        GoTo NoData
'    End If
'
'    With tblOrdSheet
'
'        .ReDraw = False
'        .MaxRows = 0
'        objProInSts.Max = tmpRs.RecordCount
'
'        If tmpRs.RecordCount < lngMaxRows Then
'            .MaxRows = lngMaxRows
'        Else
'            .MaxRows = tmpRs.RecordCount   '����Ÿ �Ǽ�
'        End If
'
'        txtMesg.Text = ""
'        lblDeptNm.Caption = .Text
'
'        For i = 1 To tmpRs.RecordCount
'
'            objProInSts.Value = i
'
'            .Row = i
'
'            If SvOrdDt <> Trim("" & tmpRs.Fields("OrdDt").Value) Then
'                .Col = enCOLLIST.tcORDDT:   .Text = Format("" & tmpRs.Fields("OrdDt").Value, CS_DateShortMask)    'ó����
'                .Col = enCOLLIST.tcORDNO:   .Text = Trim("" & tmpRs.Fields("OrdNo").Value)     'ó���ȣ
'                .Col = enCOLLIST.tcSPCNM:   .Text = Trim("" & tmpRs.Fields("SpcNm").Value)     '��ü
'                .Col = enCOLLIST.tcDOCTNM:  .Text = Trim("" & tmpRs.Fields("DoctNm").Value)     'ó����
'                SvOrdDt = Trim("" & tmpRs.Fields("OrdDt").Value)
'                SvOrdNo = Trim("" & tmpRs.Fields("OrdNo").Value)    'ó���ȣ
'                SvSpcNm = Trim("" & tmpRs.Fields("SpcNm").Value)    '��ü
'                SvOrdDoct = Trim("" & tmpRs.Fields("DoctNm").Value) 'ó����
'            End If
'            If SvOrdNo <> Trim("" & tmpRs.Fields("OrdNo").Value) Then
'                .Col = enCOLLIST.tcORDNO:   .Text = Trim("" & tmpRs.Fields("OrdNo").Value)     'ó���ȣ
'                .Col = enCOLLIST.tcSPCNM:   .Text = Trim("" & tmpRs.Fields("SpcNm").Value)     '��ü
'                .Col = enCOLLIST.tcDOCTNM:  .Text = Trim("" & tmpRs.Fields("DoctNm").Value)    'ó����
'                SvOrdNo = Trim("" & tmpRs.Fields("OrdNo").Value)    'ó���ȣ
'                SvSpcNm = Trim("" & tmpRs.Fields("SpcNm").Value)    '��ü
'                SvOrdDoct = Trim("" & tmpRs.Fields("DoctNm").Value) 'ó����
'            End If
'            If SvSpcNm <> Trim("" & tmpRs.Fields("SpcNm").Value) Then
'                .Col = enCOLLIST.tcSPCNM:   .Text = Trim("" & tmpRs.Fields("SpcNm").Value)     '��ü
'                SvSpcNm = Trim("" & tmpRs.Fields("SpcNm").Value)
'            End If
'            If SvOrdDoct <> Trim("" & tmpRs.Fields("DoctNm").Value) Then
'                .Col = enCOLLIST.tcDOCTNM: .Text = Trim("" & tmpRs.Fields("DoctNm").Value)    'ó����
'                SvOrdDoct = Trim("" & tmpRs.Fields("DoctNm").Value)
'            End If
'
'            tmpStatFg = medGetP("" & tmpRs.Fields("StatFlags").Value, 1, ";")   '�ǹ��� ���ް��� ����
'            tmpTestFg = medGetP("" & tmpRs.Fields("StatFlags").Value, 2, ";")   '�ǹ��� �˻簡�� ����
'
'        '***�ǹ����� ���
'            If P_ApplyBuildingInfo Then
'
'               If Trim(tmpRs.Fields("StatFg").Value) = "1" Then
'
'                   '**���ް˻� ����
'                   If Mid(tmpStatFg, ObjSysInfo.BuildingNo, 1) = "1" Then
'
'                       '** �߾�/���̼��Ϳ��� ���ް˻簡 �߻��ϸ�.. --> ���޼��ͷ�...
'                       If ObjSysInfo.BuildingCd = CentralLab Or _
'                          ObjSysInfo.BuildingCd = AneLab Then
'                           .Col = enCOLLIST.tcBUILDCD: .Text = EmergencyLab
'                           .Col = enCOLLIST.tcBUILDNM: .Text = EmergencyLabNm
'
'                       '** �ش�ǹ����� ���ް˻� ������
'                       Else
'                           .Col = enCOLLIST.tcBUILDCD: .Text = ObjSysInfo.BuildingCd
'                           .Col = enCOLLIST.tcBUILDNM: .Text = ObjSysInfo.BuildingNm
'                       End If
'                       .Col = enCOLLIST.tcSTATFG: .Text = "1"      'StatFg
'                       GoTo DataSet
'                   Else
'                   '*******************************************************************************************************
'                   '** ����/���弾�� : ���ް˻簡 �������� ������� ���޽ǿ��� �˻簡 �����ϸ� ���޽Ƿ�, �ƴϸ� �߾�����...
'                   '*******************************************************************************************************
'                       '** ����/���弾�Ϳ��� ���ް˻簡 �߻��ϸ�..
'                       If ObjSysInfo.BuildingCd = WomLab Or ObjSysInfo.BuildingCd = HrtLab Then
'                           '** ���޽ǿ��� ���ް˻� ���� --> ���޼��ͷ�...
'                           If Mid(tmpStatFg, EmergencyNo, 1) = "1" Then
'                               .Col = enCOLLIST.tcBUILDCD: .Text = EmergencyLab
'                               .Col = enCOLLIST.tcBUILDNM: .Text = EmergencyLabNm
'                               .Col = enCOLLIST.tcSTATFG:  .Text = "1"   'StatFg
'                               GoTo DataSet
'                           End If
'                       End If
'                   '*******************************************************************************************************
'                   End If
'               End If
'
'               .Col = enCOLLIST.tcSTATFG: .Text = "0"         'StatFg
'
'               '**�Ϲݰ˻簡��
'               If Mid(tmpTestFg, ObjSysInfo.BuildingNo, 1) = "1" Then
'                   .Col = enCOLLIST.tcBUILDCD: .Text = ObjSysInfo.BuildingCd
'                   .Col = enCOLLIST.tcBUILDNM: .Text = ObjSysInfo.BuildingNm
'
'               '**�Ϲݰ˻� �Ұ��� --> �߾Ӱ˻�Ƿ�...
'               Else
'                   .Col = enCOLLIST.tcBUILDCD: .Text = CentralLab
'                   .Col = enCOLLIST.tcBUILDNM: .Text = CentralLabNm
'               End If
'
'        '***�ǹ����� ������� ����
'            Else
'                .Col = enCOLLIST.tcBUILDCD: .Text = ObjSysInfo.BuildingCd
'                .Col = enCOLLIST.tcBUILDNM: .Text = ObjSysInfo.BuildingNm
'            End If
'
'DataSet:
'            .Col = enCOLLIST.tcTESTNM:  .Text = Trim("" & tmpRs.Fields("TestNm").Value)    'ó���
'                    Select Case tmpRs.Fields("orddiv")
'                        Case "A": .ForeColor = &H5E3F00     '&HDF6A3E     '&H00DF6A3E&�ణ �Ķ���
'                        Case "B": .ForeColor = &H496835     '&H6C6181     '&H81815A     '�ణ���   &H00845584&�����
'                        Case "L": .ForeColor = &H553755
'                    End Select
'            .Col = enCOLLIST.tcSTATFG:  .Text = IIf("" & tmpRs.Fields("StatFg").Value = "0", "", "Y") '���޿���
'                                        .ForeColor = DCM_Red                                '������
'            .Col = enCOLLIST.tcREQDTTM: .Text = Format("" & tmpRs.Fields("ReqDt").Value, CS_DateMask) & " " & _
'                                         Format("" & tmpRs.Fields("ReqTm").Value, CS_TimeLongMask)      '���ä���Ͻ�
'            .Col = enCOLLIST.tcORDDATE: .Text = Trim("" & tmpRs.Fields("OrdDt").Value)     'ó����
'            .Col = enCOLLIST.tcORDNUM:  .Text = Trim("" & tmpRs.Fields("OrdNo").Value)     'ó���ȣ
'            .Col = enCOLLIST.tcORDSEQ:  .Text = Trim("" & tmpRs.Fields("OrdSeq").Value)    'ó��Seq
'            .Col = enCOLLIST.tcTESTCD:  .Text = Trim("" & tmpRs.Fields("OrdCd").Value)     '�˻��ڵ�
'
'            Call ObjLISComCode.LisItem.KeyChange(.Text)
'            .Col = enCOLLIST.tcLABDIV:  .Text = ObjLISComCode.LisItem.Fields("labdiv")         'LabDiv
'
'            .Col = enCOLLIST.tcSPCCD:   .Text = Trim("" & tmpRs.Fields("SpcCd").Value)     '��ü�ڵ�
'
'            Call ObjLISComCode.LisSpc.KeyChange(.Text)
'            .Col = enCOLLIST.tcSPCABBR:  .Text = ObjLISComCode.LisSpc.Fields("spcabbr")        '��ü����
'            .Col = enCOLLIST.tcLABRANGE: .Text = ObjLISComCode.LisSpc.Fields("labrange")       '�̻���������ȣ����
'
'            .Col = enCOLLIST.tcWORKAREA: .Text = Trim("" & tmpRs.Fields("WorkArea").Value) 'WorkArea
'            .Col = enCOLLIST.tcSTORECD: .Text = Trim("" & tmpRs.Fields("StoreCd").Value)   '�����ڵ�
'            .Col = enCOLLIST.tcTESTDIV: .Text = Trim("" & tmpRs.Fields("TestDiv").Value)   '�˻籸��
'            .Col = enCOLLIST.tcMULTIFG: .Text = Trim("" & tmpRs.Fields("MultiFg").Value)   '������ü����
'            .Col = enCOLLIST.tcSPCGRP:  .Text = Trim("" & tmpRs.Fields("SpcGrp").Value)    '��ü��
'            .Col = enCOLLIST.tcORDDOCT: .Text = Trim("" & tmpRs.Fields("OrdDoct").Value)   'ó����
'            .Col = enCOLLIST.tcMAJDODT: .Text = Trim("" & tmpRs.Fields("MajDoct").Value)   '��ġ��
'            .Col = enCOLLIST.tcDEPTCD:  .Text = Trim("" & tmpRs.Fields("DeptCd").Value)    '�����
'                                        '�������
'                                        If .Text <> "" And lblDeptNm.Caption = "" Then
'                                            If ObjLISComCode.DeptCd.Exists(.Text) Then
'                                                ObjLISComCode.DeptCd.KeyChange (.Text)
'                                                lblDeptNm.Caption = ObjLISComCode.DeptCd.Fields("deptnm")
'                                            End If
'                                        End If
'            .Col = enCOLLIST.tcABBRNM:  .Text = Trim("" & tmpRs.Fields("AbbrNm5").Value)   '����
'            .Col = enCOLLIST.tcBARCNT:  .Text = Trim("" & tmpRs.Fields("LabelCnt").Value)  '��������
'            .Col = enCOLLIST.tcPAYDT:   .Text = Trim("" & tmpRs.Fields("ReceptNo").Value)  '��������ȣ
'                                        .ForeColor = vbRed
'
'            .Col = enCOLLIST.tcWARDID:  .Text = Trim("" & tmpRs.Fields("WardId").Value)    '����
'                                        strWardId = .Text
'            .Col = enCOLLIST.tcROOMID:  .Text = Trim("" & tmpRs.Fields("RoomId").Value)    '����
'                                        strRoomId = .Text
'            .Col = enCOLLIST.tcBEDID:   .Text = Trim("" & tmpRs.Fields("BedId").Value)     '����
'                                        strBedId = .Text
'
'            If strWardId <> "" And lblLocation.Caption = "" Then
'                lblLocation.Caption = strWardId & "-" & strRoomId & "-" & strBedId
'            End If
'
'            '����μ� Remark
'            If Trim("" & tmpRs.Fields("Mesg").Value) <> "" Then
'                txtRemark.Text = txtRemark.Text & "# " & Format(Trim("" & tmpRs.Fields("OrdNo").Value), "##") & " - "
'                txtRemark.Text = txtRemark.Text & Trim("" & tmpRs.Fields("TestNm").Value) & vbCrLf
'                txtRemark.Text = txtRemark.Text & Trim("" & tmpRs.Fields("Mesg").Value) & vbCrLf
'            End If
'
'            tmpRs.MoveNext
'        Next
'
'        .RowHeight(-1) = lngRowHeight
'        .ReDraw = True
'
'    End With
'    blnOrdFg = True
'    fraOrder.Enabled = True
'
'NoData:
'    MouseDefault
'    tmpRs.RsClose
'    Set tmpRs = Nothing
'    Set objProInSts = Nothing
'
'End Sub
'
'
''% �˻��� ó���� ���̺� ���÷��� �Ѵ�.
'Private Sub DisplayOrder_back()
'
'    Dim i As Integer
'    Dim SqlStmt As String
''    Dim objGetSql As New clsGetSqlStatement
'    Dim tmpRs As DrRecordSet
'    Dim SvOrdDt As String, SvOrdNo As String
'    Dim SvSpcNm As String, SvOrdDoct As String
'    Dim tmpDate As String, tmpTime As String
'    Dim tmpStatFg As String
'    Dim tmpTestFg As String
'    Dim strErChk As String
'
'On Error GoTo NoData
'
'    '
'    TestBuilding_Search     '�������� ��ü���� ���
'
'    ' ó�泻�� �˻�
'    tmpDate = Format(Now, CS_DateDbFormat)
'    tmpTime = Format(Now, CS_TimeDbFormat)
'
'    SqlStmt = objMySql.SqlReadOrder(txtPtId.Text, tmpDate, tmpTime, , CS_BussIn)
'    Set tmpRs = OpenRecordSet(SqlStmt)
'    If tmpRs.EOF Then
'       tmpRs.RsClose
'       Set tmpRs = Nothing
'
'       MsgBox MyPatient.PtNm & " ���� ó�泻���� �����ϴ�"
'       If Not blnCleared Then Call ClearRtn
'       Exit Sub
'    End If
'
'    With tblOrdSheet
'
'       .ReDraw = False
'       .MaxRows = 0
'       If tmpRs.RecordCount < 20 Then
'          .MaxRows = 20
'          .Row = tmpRs.RecordCount + 1
'          .Row2 = 20
'          .Col = 1: .Col2 = .MaxCols
'          .BlockMode = True
'          .Lock = True
'          .Protect = True
'          .BlockMode = False
'       Else
'          .MaxRows = tmpRs.RecordCount   '����Ÿ �Ǽ�
'       End If
'
'       'Locking Cells
'       .Row = -1
'       .Col = 2: .Col2 = .MaxCols
'       .BlockMode = True
'       .Lock = True
'       .Protect = True
'       .BlockMode = False
'
'       For i = 1 To tmpRs.RecordCount
'          .Row = i
'          '.Col = 1: .Value = 0
'          If SvOrdDt <> Trim("" & tmpRs.Fields("OrdDt").Value) Then
'             .Col = 2: .Value = Format("" & tmpRs.Fields("OrdDt").Value, CS_DateMask)   'ó����
'             .Col = 3: .Value = Trim("" & tmpRs.Fields("OrdNo").Value)      'ó���ȣ
'             .Col = 5: .Value = Trim("" & tmpRs.Fields("SpcNm").Value)      '��ü
'             .Col = 6: .Value = Trim("" & tmpRs.Fields("DoctNm").Value)     'ó����
'             SvOrdDt = Trim("" & tmpRs.Fields("OrdDt").Value)
'             SvOrdNo = Trim("" & tmpRs.Fields("OrdNo").Value)               'ó���ȣ
'             SvSpcNm = Trim("" & tmpRs.Fields("SpcNm").Value)               '��ü
'             SvOrdDoct = Trim("" & tmpRs.Fields("DoctNm").Value)            'ó����
'          End If
'          If SvOrdNo <> Trim("" & tmpRs.Fields("OrdNo").Value) Then
'             .Col = 3: .Value = Trim("" & tmpRs.Fields("OrdNo").Value)      'ó���ȣ
'             .Col = 5: .Value = Trim("" & tmpRs.Fields("SpcNm").Value)      '��ü
'             .Col = 6: .Value = Trim("" & tmpRs.Fields("DoctNm").Value)     'ó����
'             SvOrdNo = Trim("" & tmpRs.Fields("OrdNo").Value)               'ó���ȣ
'             SvSpcNm = Trim("" & tmpRs.Fields("SpcNm").Value)               '��ü
'             SvOrdDoct = Trim("" & tmpRs.Fields("DoctNm").Value)            'ó����
'          End If
'          If SvSpcNm <> Trim("" & tmpRs.Fields("SpcNm").Value) Then
'             .Col = 5: .Value = Trim("" & tmpRs.Fields("SpcNm").Value)      '��ü
'             SvSpcNm = Trim("" & tmpRs.Fields("SpcNm").Value)
'          End If
'          If SvOrdDoct <> Trim("" & tmpRs.Fields("DoctNm").Value) Then
'             .Col = 6: .Value = Trim("" & tmpRs.Fields("DoctNm").Value)     'ó����
'             SvOrdDoct = Trim("" & tmpRs.Fields("DoctNm").Value)
'          End If
'
'          If .Value <> "" And lblDoctNm.Caption = "" Then lblDoctNm.Caption = SvOrdDoct
'
'          tmpStatFg = medGetP("" & tmpRs.Fields("StatFlags").Value, 1, ";") '�ǹ��� ���ް��� ����
'          tmpTestFg = medGetP("" & tmpRs.Fields("StatFlags").Value, 2, ";") '�ǹ��� �˻簡�� ����
'
'          Select Case tmpRs.Fields("orddiv")
'          Case "A":
'            .Col = 21: .Value = Trim("" & tmpRs.Fields("StatFg").Value)     '���޿���  --> ������ ó��...
'            .Col = 25: .Value = CentralLab
'            .Col = 26: .Value = CentralLabNm
'          Case "B":
'            strErChk = objGetSql.ER_Chk(txtPtId.Text, SvOrdDt)
'            .Col = 21: .Value = Trim("" & tmpRs.Fields("StatFg").Value)     '���޿���  --> ������ ó��...
'            .Col = 25: .Value = IIf(strErChk = "1", strErBldCd, strGBldCd)
'            If ObjLISComCode.Building.Exists(.Value) Then
'                ObjLISComCode.Building.KeyChange (.Value)
'            End If
'            .Col = 26: .Value = ObjLISComCode.Building.Fields("buildnm")
'          Case "L":
'              If Trim(tmpRs.Fields("StatFg").Value) = "1" Then
'                 If Mid(tmpStatFg, ObjSysInfo.BuildingNo, 1) = "1" Then     '**���ް˻� ����
'                    If ObjSysInfo.BuildingCd = CentralLab Or ObjSysInfo.BuildingCd = AneLab Then   '** �߾�/���̼��Ϳ��� ���ް˻簡 �߻��ϸ�..
'                       .Col = 25: .Value = EmergencyLab                     '--> ���޼��ͷ�...
'                       .Col = 26: .Value = EmergencyLabNm
'                    Else
'                       .Col = 25: .Value = ObjSysInfo.BuildingCd            '**�ش�ǹ����� ���ް˻� ������
'                       .Col = 26: .Value = ObjSysInfo.BuildingNm
'                    End If
'                    .Col = 21: .Value = "1"  'StatFg
'                    GoTo DataSet
'                 Else
'                 '*******************************************************************************************************
'                 '** ����/���弾�� : ���ް˻簡 �������� ������� ���޽ǿ��� �˻簡 �����ϸ� ���޽Ƿ�, �ƴϸ� �߾�����...
'                 '*******************************************************************************************************
'                    If ObjSysInfo.BuildingCd = WomLab Or ObjSysInfo.BuildingCd = HrtLab Then     '** ����/���弾�Ϳ��� ���ް˻簡 �߻��ϸ�..
'                         If Mid(tmpStatFg, EmergencyNo, 1) = "1" Then    '** ���޽ǿ��� ���ް˻� ����
'                           .Col = 25: .Value = EmergencyLab              '--> ���޼��ͷ�...
'                           .Col = 26: .Value = EmergencyLabNm
'                           .Col = 21: .Value = "1"  'StatFg
'                           GoTo DataSet
'                         End If
'                    End If
'                 '*******************************************************************************************************
'
'                 End If
'              End If
'
'              .Col = 21: .Value = "0"        'StatFg
'              If Mid(tmpTestFg, ObjSysInfo.BuildingNo, 1) = "1" Then
'                 .Col = 25: .Value = ObjSysInfo.BuildingCd          '**�Ϲݰ˻簡��
'                 .Col = 26: .Value = ObjSysInfo.BuildingNm          '**�Ϲݰ˻簡��
'              Else
'                 .Col = 25: .Value = CentralLab                     '**�Ϲݰ˻� �Ұ��� --> �߾Ӱ˻�Ƿ�...
'                 .Col = 26: .Value = CentralLabNm
'              End If
'
'          End Select
'
'DataSet:
'          .Col = 4: .Value = Trim("" & tmpRs.Fields("TestNm").Value)        'ó���
'                    '.ForeColor = &H8E6000           '�ణ �Ķ���   &H00898963& '�ణ���  &H00553755&
'                    Select Case tmpRs.Fields("orddiv")
'                        Case "A": .ForeColor = &H5E3F00     '&HDF6A3E     '&H00DF6A3E&�ణ �Ķ���
'                        Case "B": .ForeColor = &H496835     '&H6C6181     '&H81815A     '�ణ���   &H00845584&�����
'                        Case "L": .ForeColor = &H553755
'                    End Select
'          .Col = 7: .Value = Choose(Val("" & tmpRs.Fields("StatFg").Value) + 1, "", "Y")     '���޿���
'                    .ForeColor = &HFF&       '������
'          .Col = 8: .Value = Format("" & tmpRs.Fields("ReqDt").Value, CS_DateMask) & " " & _
'                                   Format("" & tmpRs.Fields("ReqTm").Value, CS_TimeLongMask)        '���ä���Ͻ�
'          .Col = 9: .Value = Trim("" & tmpRs.Fields("OrdDt").Value)         'ó����
'          .Col = 10: .Value = Trim("" & tmpRs.Fields("OrdNo").Value)        'ó���ȣ
'          .Col = 11: .Value = Trim("" & tmpRs.Fields("OrdSeq").Value)       'ó��Seq
'          .Col = 12: .Value = Trim("" & tmpRs.Fields("OrdCd").Value)        '�˻��ڵ�
'          .Col = 13: .Value = Trim("" & tmpRs.Fields("SpcCd").Value)        '��ü�ڵ�
'          .Col = 14: .Value = Trim("" & tmpRs.Fields("WorkArea").Value)     'WorkArea
'          .Col = 15: .Value = Trim("" & tmpRs.Fields("StoreCd").Value)      '�����ڵ�
'          .Col = 16: .Value = Trim("" & tmpRs.Fields("TestDiv").Value)      '�˻籸��
'          .Col = 17: .Value = Trim("" & tmpRs.Fields("MultiFg").Value)      '������ü����
'          .Col = 18: .Value = Trim("" & tmpRs.Fields("SpcGrp").Value)       '��ü��
'          .Col = 19: .Value = Trim("" & tmpRs.Fields("OrdDoct").Value)      'ó����
'          .Col = 20: .Value = Trim("" & tmpRs.Fields("MajDoct").Value)      '��ġ��
'          '.Col = 21: .Value = Trim("" & tmpRs.Fields("StatFg").Value)      '���޿���  --> ������ ó��...
'          .Col = 22: .Value = Trim("" & tmpRs.Fields("DeptCd").Value)       '�����
''                    If .Value <> "" And lblDeptNm.Caption = "" Then
''                        Dim MyResult As New clsLISResultReview
''                        lblDeptNm.Caption = MyResult.GetDeptNm(Trim("" & tmpRs.Fields("DeptCd").Value))
''                        Set MyResult = Nothing
''                    End If
'          .Col = 23: .Value = Trim("" & tmpRs.Fields("AbbrNm5").Value)      '����
'          .Col = 24: .Value = Trim("" & tmpRs.Fields("LabelCnt").Value)     '��������
'          .Col = 27: .Value = Trim("" & tmpRs.Fields("WardId").Value)       '����ID
'                    If .Value <> "" Then MyPatient.WardId = .Value
'          .Col = 28: .Value = Trim("" & tmpRs.Fields("HosilId").Value)      'ȣ��ID
'                    If .Value <> "" Then MyPatient.RoomId = .Value
'          .Col = 29: .Value = Trim("" & tmpRs.Fields("BedInDt").Value)      '�Կ���
'                    If .Value <> "" Then MyPatient.BedInDt = .Value
'          .Col = 30: .Value = Trim("" & tmpRs.Fields("FzFg").Value)         '��������
'          .Col = 31: .Value = Trim("" & tmpRs.Fields("OrdDiv").Value)       'ó�汸��
'          .Col = 32: .Value = Trim("" & tmpRs.Fields("SpcNm5").Value)       '��ü����
'
'          If Trim("" & tmpRs.Fields("Mesg").Value) <> "" Then
'             'txtMesg.Text = txtMesg.Text & Format(Trim("" & tmpRs.Fields("OrdDt").Value), "####/##/##") & "  "
'             'txtMesg.Text = txtMesg.Text & Format(Trim("" & tmpRs.Fields("OrdTm").Value), "0#:##") & " - "
'             txtMesg.Text = txtMesg.Text & "# " & Format(Trim("" & tmpRs.Fields("OrdNo").Value), "##") & " - "
'             txtMesg.Text = txtMesg.Text & Trim("" & tmpRs.Fields("TestNm").Value) & vbCrLf
'             txtMesg.Text = txtMesg.Text & Trim("" & tmpRs.Fields("Mesg").Value) & vbCrLf
'          End If
'
'          tmpRs.MoveNext
'       Next
'
'       lblLocation.Caption = MyPatient.WardId & " - " & MyPatient.RoomId
'
'       .ReDraw = True
'
'    End With
'    blnOrdFg = True
'    fraOrder.Enabled = True
'    blnCleared = False
'
'NoData:
'    Resume Next
'    tmpRs.RsClose
'    Set tmpRs = Nothing
'
'End Sub
'
'Private Sub TestBuilding_Search()
'
''    Dim objSql As New clsGetSqlStatement
'    Dim strTmp As String
'
'    objSql.setDbConn dbconn
'
'    With objSql
'        If txtWardId = "" Then
'            strBlgCd = ObjSysInfo.BuildingCd
'        Else
'            strBlgCd = .Get_BuildingCd(UCase(txtWardId))
'        End If
'        strTmp = .TestBuildCd(strBlgCd)
'        strErBldCd = medGetP(strTmp, 1, COL_DIV)
'        strGBldCd = medGetP(strTmp, 2, COL_DIV)
'    End With
'
''    With tblCount
''        .Row = 1: .Col = 1: .Value = strErbldcd
''        objBBS901.TestBldNm dbConn, strErbldcd
''        .Row = 1: .Col = 2: .Value = objBBS901.ErBuildNM
''        .Row = 2: .Col = 1: .Value = strGbldcd
''        objBBS901.TestBldNm dbConn, strGbldcd
''        .Row = 2: .Col = 2: .Value = objBBS901.GBuildNM
''    End With
'    Set objSql = Nothing
'
'End Sub
'
'
'Private Sub ClearRtn()
'
'   lblPtNm.Caption = ""
'   lblSexAge.Caption = ""
''   lblAge.Caption = ""
''   lblAgeDiv.Caption = ""
'   lblDeptNm.Caption = ""
'   lblLocation.Caption = ""
'   lblDoctNm.Caption = ""
'   txtMesg.Text = ""
'   chkSelAll.Value = 0
'   chkChangeColTm.Value = 0
'   dtpColDtTm.Value = dbconn.GetSysDate
'   dtpColDtTm.Enabled = False
'   fraOrder.Enabled = False
'   'optSort(0).Value = True
'   With tblOrdSheet
'      .Row = -1
'      .Col = -1
'      .BlockMode = True
'      .Action = ActionClearText
'      .BlockMode = False
'   End With
'   cmdSave.Enabled = False
''   fraStatus.Visible = False
''   lblStatus.Caption = ""
'   blnOrdFg = False
'   PtFg = False
'   MsgFg = False
'   Set MyPatient = Nothing
'   DoEvents
'
'   Set objLISCollect = Nothing
'   Set objLISCollect = New clsLISCollectioin
'   'Set objLISCollect.MyOraSE = OraSe
''   Call objLISCollect.SetDatabase(DbConn)
'
'   Set MyPatient = New clsPatient
'   'Set MyPatient.MyOraSE = OraSe
'   Set MyPatient.objDB = dbconn
'   DoEvents
'
'   blnCleared = True
'
'End Sub
'
'Public Function CheckSameOrder(ByVal intSel As Integer) As Integer
'
'   Dim i As Integer, j As Integer
'   Dim SaveCode As String
'   Dim SaveSpc As String
'   Dim SaveDate As String
'   Dim SaveStatFg As String
'
'   CheckSameOrder = 0
'   With tblOrdSheet
'      For i = 1 To .DataRowCnt
'         .Row = i
'         .Col = 1
'         If .Value <> intSel Then GoTo Skip1
'         .Col = 12:  SaveCode = .Value
'         .Col = 8:  SaveDate = .Value
'         .Col = 13:  SaveSpc = .Value
'         .Col = 21:  SaveStatFg = .Value
'         For j = i + 1 To .DataRowCnt
'            .Row = j
'            .Col = 1
'            If .Value <> intSel Then GoTo Skip2
'            .Col = 12
'            If .Value = SaveCode Then
'               .Col = 8
'               If .Value = SaveDate Then
'                    .Col = 13
'                    If .Value = SaveSpc Then
'                        .Col = 21
'                        If .Value = SaveStatFg Then
'                           CheckSameOrder = j
'                           Exit Function
'                        End If
'                    End If
'               End If
'            End If
'Skip2:
'         Next
'Skip1:
'      Next
'   End With
'
'End Function
'
'Public Sub Call_PtId_KeyPress()
'
'   Call txtPtId_KeyPress(vbKeyReturn)
'
'End Sub
'
'
'
'
'


