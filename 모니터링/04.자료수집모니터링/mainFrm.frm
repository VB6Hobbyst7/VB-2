VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Begin VB.Form mainFrm 
   AutoRedraw      =   -1  'True
   Caption         =   "�ѱ��ؾ�������ȸ �ڷ���� ��Ȳ ����͸�"
   ClientHeight    =   10500
   ClientLeft      =   870
   ClientTop       =   630
   ClientWidth     =   12270
   Icon            =   "mainFrm.frx":0000
   LinkTopic       =   "�ڷ���� ��Ȳ"
   LockControls    =   -1  'True
   ScaleHeight     =   11473.57
   ScaleMode       =   0  '�����
   ScaleWidth      =   12270
   StartUpPosition =   2  'ȭ�� ���
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '�Ʒ� ����
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   10245
      Width           =   12270
      _ExtentX        =   21643
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   18
            MinWidth        =   18
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   18494
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "���� 3:46"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab mntTab 
      Height          =   9855
      Left            =   120
      TabIndex        =   1
      Top             =   90
      Width           =   12285
      _ExtentX        =   21669
      _ExtentY        =   17383
      _Version        =   393216
      Tabs            =   8
      TabsPerRow      =   7
      TabHeight       =   520
      TabCaption(0)   =   "��Ȳ"
      TabPicture(0)   =   "mainFrm.frx":058A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(2)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(5)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(4)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fpSpread_Tot_Usn"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "fpSpread_Tot_Ag"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "fpSpread_Tot_DtVPN"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "fpSpread_Tot_DtCDMA"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "fpSpread_Tot_Tw"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "fpSpread_Tot_Rt"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Frm_Time"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "����������"
      TabPicture(1)   =   "mainFrm.frx":05A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Panel_Time_Dt"
      Tab(1).Control(1)=   "SSPanel6(0)"
      Tab(1).Control(2)=   "fpSpread_DtLog"
      Tab(1).Control(3)=   "DT_Timer"
      Tab(1).Control(4)=   "Frame1"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "�����ؾ��������"
      TabPicture(2)   =   "mainFrm.frx":05C2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Panel_Time_Tw"
      Tab(2).Control(1)=   "SSPanel6(1)"
      Tab(2).Control(2)=   "fpSpread_TwLog"
      Tab(2).Control(3)=   "Frame2"
      Tab(2).Control(4)=   "TW_Timer"
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "�ؾ������"
      TabPicture(3)   =   "mainFrm.frx":05DE
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Panel_Time_Rt"
      Tab(3).Control(1)=   "SSPanel6(2)"
      Tab(3).Control(2)=   "fpSpread_RtLog"
      Tab(3).Control(3)=   "Frame3"
      Tab(3).Control(4)=   "RT_Timer"
      Tab(3).ControlCount=   5
      TabCaption(4)   =   "�ؼ���������"
      TabPicture(4)   =   "mainFrm.frx":05FA
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Panel_Time_Ag"
      Tab(4).Control(1)=   "SSPanel6(3)"
      Tab(4).Control(2)=   "fpSpread_AgLog"
      Tab(4).Control(3)=   "Frame4"
      Tab(4).Control(4)=   "AG_Timer"
      Tab(4).ControlCount=   5
      TabCaption(5)   =   "USN"
      TabPicture(5)   =   "mainFrm.frx":0616
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Panel_Time_Usn"
      Tab(5).Control(1)=   "SSPanel6(4)"
      Tab(5).Control(2)=   "fpSpread_UsnLog"
      Tab(5).Control(3)=   "USN_Timer"
      Tab(5).Control(4)=   "Frame5"
      Tab(5).ControlCount=   5
      TabCaption(6)   =   "�ؾ���� View"
      TabPicture(6)   =   "mainFrm.frx":0632
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "ss_Tw"
      Tab(6).Control(1)=   "Fm_Tw"
      Tab(6).ControlCount=   2
      TabCaption(7)   =   "�ؾ���� View"
      TabPicture(7)   =   "mainFrm.frx":064E
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Fm_RTID"
      Tab(7).Control(1)=   "fpSpread2"
      Tab(7).ControlCount=   2
      Begin VB.Frame Frame1 
         Caption         =   "���������� �ڷ� ���� �α� �˻�����"
         Height          =   735
         Left            =   -74640
         TabIndex        =   75
         Top             =   695
         Width           =   11655
         Begin VB.ComboBox cmbxSechDTID 
            Height          =   300
            ItemData        =   "mainFrm.frx":066A
            Left            =   1155
            List            =   "mainFrm.frx":066C
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   82
            Top             =   285
            Width           =   1575
         End
         Begin VB.TextBox txtSechDTStDate 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "yyyy-MM-dd"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1042
               SubFormatType   =   3
            EndProperty
            Height          =   300
            Left            =   3510
            MaxLength       =   10
            TabIndex        =   81
            Top             =   300
            Width           =   1215
         End
         Begin VB.TextBox txtSechDTEdDate 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "yyyy-MM-dd"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1042
               SubFormatType   =   3
            EndProperty
            Height          =   300
            Left            =   5655
            MaxLength       =   10
            TabIndex        =   80
            Top             =   300
            Width           =   1170
         End
         Begin VB.CommandButton btnDtSearch 
            Caption         =   "�˻�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   7560
            TabIndex        =   79
            Top             =   300
            Width           =   735
         End
         Begin VB.ComboBox cmbxSechDTStHour 
            Height          =   300
            Left            =   4770
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   78
            Top             =   300
            Width           =   615
         End
         Begin VB.ComboBox cmbxSechDTEdHour 
            Height          =   300
            IMEMode         =   1  '�Է� ���� ����
            Left            =   6870
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   77
            Top             =   285
            Width           =   615
         End
         Begin VB.ComboBox cmbxDtRownum 
            Height          =   300
            Left            =   10470
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   76
            Top             =   285
            Width           =   1095
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "�Ⱓ : "
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2835
            TabIndex        =   86
            Top             =   345
            Width           =   735
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "������ : "
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   315
            TabIndex        =   85
            Top             =   345
            Width           =   855
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "~"
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
            Left            =   5430
            TabIndex        =   84
            Top             =   300
            Width           =   165
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "��°Ǽ� : "
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
            Left            =   9360
            TabIndex        =   83
            Top             =   360
            Width           =   1065
         End
      End
      Begin VB.Timer DT_Timer 
         Left            =   -66600
         Top             =   8780
      End
      Begin VB.Timer TW_Timer 
         Left            =   -66600
         Top             =   8780
      End
      Begin VB.Frame Frame2 
         Caption         =   "�����ؾ�������� �ڷ� ���� �α� �˻�����"
         Height          =   735
         Left            =   -74640
         TabIndex        =   64
         Top             =   695
         Width           =   11655
         Begin VB.ComboBox cmbxTwRownum 
            Height          =   300
            Left            =   10470
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   69
            Top             =   285
            Width           =   1095
         End
         Begin VB.CommandButton btnTwSearch 
            Caption         =   "�˻�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   6330
            TabIndex        =   68
            Top             =   285
            Width           =   735
         End
         Begin VB.TextBox txtSechTWEdDate 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "yyyy-MM-dd"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1042
               SubFormatType   =   3
            EndProperty
            Height          =   300
            Left            =   5085
            MaxLength       =   10
            TabIndex        =   67
            Top             =   300
            Width           =   1170
         End
         Begin VB.TextBox txtSechTWStDate 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "yyyy-MM-dd"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1042
               SubFormatType   =   3
            EndProperty
            Height          =   300
            Left            =   3510
            MaxLength       =   10
            TabIndex        =   66
            Top             =   300
            Width           =   1215
         End
         Begin VB.ComboBox cmbxSechTWID 
            Height          =   300
            Left            =   1155
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   65
            Top             =   285
            Width           =   1575
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "��°Ǽ� : "
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
            Left            =   9360
            TabIndex        =   73
            Top             =   360
            Width           =   1065
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "~"
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
            Left            =   4860
            TabIndex        =   72
            Top             =   300
            Width           =   165
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "������ : "
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   315
            TabIndex        =   71
            Top             =   345
            Width           =   855
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "�Ⱓ : "
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2835
            TabIndex        =   70
            Top             =   345
            Width           =   735
         End
      End
      Begin VB.Timer RT_Timer 
         Left            =   -65160
         Top             =   8780
      End
      Begin VB.Frame Frame3 
         Caption         =   "�ؾ������ �ڷ� ���� �α� �˻�����"
         Height          =   735
         Left            =   -74640
         TabIndex        =   54
         Top             =   695
         Width           =   11655
         Begin VB.ComboBox cmbxRtRownum 
            Height          =   300
            Left            =   10470
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   59
            Top             =   285
            Width           =   1095
         End
         Begin VB.CommandButton btnRtSearch 
            Caption         =   "�˻�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   6330
            TabIndex        =   58
            Top             =   285
            Width           =   735
         End
         Begin VB.TextBox txtSechRTEdDate 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "yyyy-MM-dd"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1042
               SubFormatType   =   3
            EndProperty
            Height          =   300
            Left            =   5025
            MaxLength       =   10
            TabIndex        =   57
            Top             =   285
            Width           =   1170
         End
         Begin VB.TextBox txtSechRTStDate 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "yyyy-MM-dd"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1042
               SubFormatType   =   3
            EndProperty
            Height          =   300
            Left            =   3510
            MaxLength       =   10
            TabIndex        =   56
            Top             =   300
            Width           =   1215
         End
         Begin VB.ComboBox cmbxSechRTID 
            Height          =   300
            Left            =   1155
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   55
            Top             =   285
            Width           =   1575
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "��°Ǽ� : "
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
            Left            =   9360
            TabIndex        =   63
            Top             =   360
            Width           =   1065
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "~"
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
            Left            =   4800
            TabIndex        =   62
            Top             =   285
            Width           =   165
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "������ : "
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   315
            TabIndex        =   61
            Top             =   345
            Width           =   855
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "�Ⱓ : "
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2835
            TabIndex        =   60
            Top             =   345
            Width           =   735
         End
      End
      Begin VB.Timer AG_Timer 
         Left            =   -65160
         Top             =   8780
      End
      Begin VB.Frame Frame4 
         Caption         =   "�ؼ��������� �ڷ� ���� �α� �˻�����"
         Height          =   735
         Left            =   -74640
         TabIndex        =   44
         Top             =   695
         Width           =   11655
         Begin VB.ComboBox cmbxAgRownum 
            Height          =   300
            Left            =   10470
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   49
            Top             =   285
            Width           =   1095
         End
         Begin VB.CommandButton btnAgSearch 
            Caption         =   "�˻�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   6285
            TabIndex        =   48
            Top             =   300
            Width           =   735
         End
         Begin VB.TextBox txtSechAGEdDate 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "yyyy-MM-dd"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1042
               SubFormatType   =   3
            EndProperty
            Height          =   300
            Left            =   5010
            MaxLength       =   10
            TabIndex        =   47
            Top             =   300
            Width           =   1170
         End
         Begin VB.TextBox txtSechAGStDate 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "yyyy-MM-dd"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1042
               SubFormatType   =   3
            EndProperty
            Height          =   300
            Left            =   3510
            MaxLength       =   10
            TabIndex        =   46
            Top             =   300
            Width           =   1215
         End
         Begin VB.ComboBox cmbxSechAGID 
            Height          =   300
            Left            =   1155
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   45
            Top             =   285
            Width           =   1575
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "��°Ǽ� : "
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
            Left            =   9360
            TabIndex        =   53
            Top             =   360
            Width           =   1065
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "~"
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
            Left            =   4785
            TabIndex        =   52
            Top             =   300
            Width           =   165
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "���� ID : "
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
            Left            =   315
            TabIndex        =   51
            Top             =   345
            Width           =   930
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "�Ⱓ : "
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2835
            TabIndex        =   50
            Top             =   345
            Width           =   735
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "USN �ڷ� ���� �α� �˻�����"
         Height          =   735
         Left            =   -74640
         TabIndex        =   34
         Top             =   695
         Width           =   11655
         Begin VB.ComboBox cmbxSechUSNID 
            Height          =   300
            Left            =   1155
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   39
            Top             =   285
            Width           =   1575
         End
         Begin VB.TextBox txtSechUSNStDate 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "yyyy-MM-dd"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1042
               SubFormatType   =   3
            EndProperty
            Height          =   300
            Left            =   3510
            MaxLength       =   10
            TabIndex        =   38
            Top             =   300
            Width           =   1215
         End
         Begin VB.TextBox txtSechUSNEdDate 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "yyyy-MM-dd"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1042
               SubFormatType   =   3
            EndProperty
            Height          =   300
            Left            =   5010
            MaxLength       =   10
            TabIndex        =   37
            Top             =   300
            Width           =   1170
         End
         Begin VB.CommandButton btnUsnSearch 
            Caption         =   "�˻�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   6285
            TabIndex        =   36
            Top             =   300
            Width           =   735
         End
         Begin VB.ComboBox cmbxUsnRownum 
            Height          =   300
            Left            =   10470
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   35
            Top             =   285
            Width           =   1095
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "�Ⱓ : "
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2835
            TabIndex        =   43
            Top             =   345
            Width           =   735
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "���� ID : "
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
            Left            =   315
            TabIndex        =   42
            Top             =   345
            Width           =   930
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "~"
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
            Left            =   4785
            TabIndex        =   41
            Top             =   300
            Width           =   165
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "��°Ǽ� : "
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
            Left            =   9360
            TabIndex        =   40
            Top             =   360
            Width           =   1065
         End
      End
      Begin VB.Timer USN_Timer 
         Left            =   -65250
         Top             =   8765
      End
      Begin VB.Frame Frm_Time 
         Height          =   585
         Left            =   60
         TabIndex        =   25
         Top             =   9165
         Width           =   11955
         Begin VB.CommandButton cmdRefresh_Tot 
            Caption         =   "������ ����"
            Height          =   375
            Left            =   8190
            Style           =   1  '�׷���
            TabIndex        =   26
            Top             =   150
            Width           =   1455
         End
         Begin VB.Timer Tot_Timer 
            Left            =   4890
            Top             =   90
         End
         Begin Threed.SSPanel SSPanel6 
            Height          =   375
            Index           =   5
            Left            =   90
            TabIndex        =   27
            Top             =   150
            Width           =   3060
            _ExtentX        =   5398
            _ExtentY        =   661
            _Version        =   196610
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "���� "
            Alignment       =   4
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.CheckBox chkTimer_Tot 
               Caption         =   "�ڵ�����"
               Height          =   255
               Left            =   180
               TabIndex        =   30
               Top             =   60
               Width           =   1095
            End
            Begin VB.TextBox schTime_Tot 
               Alignment       =   2  '��� ����
               BackColor       =   &H00404040&
               ForeColor       =   &H00FFFFFF&
               Height          =   300
               Left            =   1275
               TabIndex        =   29
               Text            =   "30"
               Top             =   50
               Width           =   495
            End
            Begin VB.ComboBox cmbInterval_Tot 
               Height          =   300
               Left            =   1950
               Style           =   2  '��Ӵٿ� ���
               TabIndex        =   28
               Top             =   37
               Width           =   630
            End
         End
         Begin Threed.SSPanel Panel_Time_Tot 
            Height          =   255
            Left            =   3270
            TabIndex        =   31
            Top             =   240
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   450
            _Version        =   196610
            ForeColor       =   16777215
            BackColor       =   4210752
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "----�� --�� --�� -- --:--:-- ( --��--��--�� )"
            BevelOuter      =   1
            Alignment       =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
      End
      Begin VB.Frame Fm_Tw 
         Caption         =   "�����ؾ�������� �ڷ� �˻�����"
         Height          =   735
         Left            =   -74640
         TabIndex        =   13
         Top             =   695
         Width           =   11655
         Begin VB.ComboBox CboTw_NM 
            Height          =   300
            ItemData        =   "mainFrm.frx":066E
            Left            =   1155
            List            =   "mainFrm.frx":0670
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   14
            Top             =   285
            Width           =   1575
         End
         Begin VB.ComboBox cboTwhh_To 
            Height          =   300
            IMEMode         =   1  '�Է� ���� ����
            Left            =   7260
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   20
            Top             =   285
            Width           =   615
         End
         Begin VB.ComboBox cboTwhh_From 
            Height          =   300
            Left            =   5280
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   19
            Top             =   300
            Width           =   615
         End
         Begin VB.CommandButton Cmd_Tw 
            Caption         =   "�˻�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   7890
            TabIndex        =   18
            Top             =   300
            Width           =   735
         End
         Begin VB.TextBox txtTwDate_To 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "yyyy-MM-dd"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1042
               SubFormatType   =   3
            EndProperty
            Height          =   300
            Left            =   6105
            MaxLength       =   10
            TabIndex        =   17
            Top             =   300
            Width           =   1170
         End
         Begin VB.TextBox txtTwDate_From 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "yyyy-MM-dd"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1042
               SubFormatType   =   3
            EndProperty
            Height          =   300
            Left            =   4080
            MaxLength       =   10
            TabIndex        =   16
            Top             =   300
            Width           =   1215
         End
         Begin VB.ComboBox CboTw_ID 
            Height          =   300
            ItemData        =   "mainFrm.frx":0672
            Left            =   1800
            List            =   "mainFrm.frx":0674
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   15
            Top             =   270
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "~"
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
            Left            =   5910
            TabIndex        =   24
            Top             =   300
            Width           =   165
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "������ : "
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   315
            TabIndex        =   23
            Top             =   345
            Width           =   855
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "�Ⱓ : "
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3465
            TabIndex        =   22
            Top             =   345
            Width           =   735
         End
         Begin VB.Label Lb_State 
            AutoSize        =   -1  'True
            Caption         =   "�˻����� : �����"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   8910
            TabIndex        =   21
            Top             =   360
            Width           =   1440
         End
      End
      Begin VB.Frame Fm_RTID 
         Caption         =   "�ؾ������ �ڷ� �˻�����"
         Height          =   735
         Left            =   -74640
         TabIndex        =   2
         Top             =   695
         Width           =   11655
         Begin VB.TextBox txtRTIDDate_From 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "yyyy-MM-dd"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1042
               SubFormatType   =   3
            EndProperty
            Height          =   300
            Left            =   4290
            MaxLength       =   10
            TabIndex        =   9
            Top             =   300
            Width           =   1215
         End
         Begin VB.TextBox txtRTIDDate_To 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "yyyy-MM-dd"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1042
               SubFormatType   =   3
            EndProperty
            Height          =   300
            Left            =   6675
            MaxLength       =   10
            TabIndex        =   8
            Top             =   300
            Width           =   1170
         End
         Begin VB.CommandButton Cmd_RTID 
            Caption         =   "�˻�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   8640
            TabIndex        =   7
            Top             =   300
            Width           =   735
         End
         Begin VB.ComboBox cboRTIDhh_From 
            Height          =   300
            Left            =   5550
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   6
            Top             =   300
            Width           =   615
         End
         Begin VB.ComboBox cboRTIDhh_To 
            Height          =   300
            IMEMode         =   1  '�Է� ���� ����
            Left            =   7890
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   5
            Top             =   285
            Width           =   615
         End
         Begin VB.ComboBox CboRTID_ID 
            Height          =   300
            ItemData        =   "mainFrm.frx":0676
            Left            =   1650
            List            =   "mainFrm.frx":0678
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   4
            Top             =   270
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.ComboBox CboRTID_NM 
            Height          =   300
            ItemData        =   "mainFrm.frx":067A
            Left            =   1155
            List            =   "mainFrm.frx":067C
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   3
            Top             =   285
            Width           =   1575
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "�Ⱓ : "
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3555
            TabIndex        =   12
            Top             =   345
            Width           =   735
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "������ : "
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   315
            TabIndex        =   11
            Top             =   345
            Width           =   855
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "~"
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
            Left            =   6330
            TabIndex        =   10
            Top             =   300
            Width           =   165
         End
      End
      Begin FPSpreadADO.fpSpread fpSpread_Tot_Rt 
         Height          =   1785
         Left            =   3780
         TabIndex        =   32
         Top             =   7230
         Width           =   4485
         _Version        =   458752
         _ExtentX        =   7911
         _ExtentY        =   3149
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpreadDesigner  =   "mainFrm.frx":067E
      End
      Begin FPSpreadADO.fpSpread fpSpread_Tot_Tw 
         Height          =   5955
         Left            =   3780
         TabIndex        =   33
         Top             =   870
         Width           =   4485
         _Version        =   458752
         _ExtentX        =   7911
         _ExtentY        =   10504
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpreadDesigner  =   "mainFrm.frx":0876
      End
      Begin FPSpreadADO.fpSpread fpSpread_DtLog 
         Height          =   7095
         Left            =   -74640
         TabIndex        =   74
         Top             =   1580
         Width           =   11655
         _Version        =   458752
         _ExtentX        =   20558
         _ExtentY        =   12515
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpreadDesigner  =   "mainFrm.frx":0A80
      End
      Begin FPSpreadADO.fpSpread fpSpread_Tot_DtCDMA 
         Height          =   3855
         Left            =   4920
         TabIndex        =   87
         Top             =   1100
         Visible         =   0   'False
         Width           =   2895
         _Version        =   458752
         _ExtentX        =   5106
         _ExtentY        =   6800
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpreadDesigner  =   "mainFrm.frx":0C54
      End
      Begin FPSpreadADO.fpSpread fpSpread_Tot_DtVPN 
         Height          =   8325
         Left            =   60
         TabIndex        =   88
         Top             =   855
         Width           =   3465
         _Version        =   458752
         _ExtentX        =   6112
         _ExtentY        =   14684
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpreadDesigner  =   "mainFrm.frx":0E28
      End
      Begin FPSpreadADO.fpSpread fpSpread_Tot_Ag 
         Height          =   4065
         Left            =   8490
         TabIndex        =   89
         Top             =   900
         Width           =   3525
         _Version        =   458752
         _ExtentX        =   6218
         _ExtentY        =   7170
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         MaxRows         =   1
         SpreadDesigner  =   "mainFrm.frx":1020
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   375
         Index           =   0
         Left            =   -74670
         TabIndex        =   90
         Top             =   8765
         Width           =   3060
         _ExtentX        =   5398
         _ExtentY        =   661
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "���� "
         Alignment       =   4
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.ComboBox cmbInterval_Dt 
            Height          =   300
            Left            =   1920
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   93
            Top             =   35
            Width           =   630
         End
         Begin VB.TextBox schTime_Dt 
            Alignment       =   2  '��� ����
            BackColor       =   &H00404040&
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   1335
            TabIndex        =   92
            Text            =   "30"
            Top             =   30
            Width           =   495
         End
         Begin VB.CheckBox chkTimer_Dt 
            Caption         =   "�ڵ�����"
            Height          =   255
            Left            =   120
            TabIndex        =   91
            Top             =   58
            Width           =   1095
         End
      End
      Begin Threed.SSPanel Panel_Time_Dt 
         Height          =   255
         Left            =   -71400
         TabIndex        =   94
         Top             =   8840
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   450
         _Version        =   196610
         ForeColor       =   16777215
         BackColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "----�� --�� --�� -- --:--:-- ( --��--��--�� )"
         BevelOuter      =   1
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin FPSpreadADO.fpSpread fpSpread_TwLog 
         Height          =   7095
         Left            =   -74640
         TabIndex        =   95
         Top             =   1580
         Width           =   11655
         _Version        =   458752
         _ExtentX        =   20558
         _ExtentY        =   12515
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpreadDesigner  =   "mainFrm.frx":1210
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   375
         Index           =   1
         Left            =   -74670
         TabIndex        =   96
         Top             =   8765
         Width           =   3060
         _ExtentX        =   5398
         _ExtentY        =   661
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "���� "
         Alignment       =   4
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkTimer_Tw 
            Caption         =   "�ڵ�����"
            Height          =   255
            Left            =   120
            TabIndex        =   99
            Top             =   58
            Width           =   1095
         End
         Begin VB.TextBox schTime_Tw 
            Alignment       =   2  '��� ����
            BackColor       =   &H00404040&
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   1350
            TabIndex        =   98
            Text            =   "30"
            Top             =   45
            Width           =   495
         End
         Begin VB.ComboBox cmbInterval_Tw 
            Height          =   300
            Left            =   1920
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   97
            Top             =   35
            Width           =   630
         End
      End
      Begin Threed.SSPanel Panel_Time_Tw 
         Height          =   255
         Left            =   -71400
         TabIndex        =   100
         Top             =   8840
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   450
         _Version        =   196610
         ForeColor       =   16777215
         BackColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "----�� --�� --�� -- --:--:-- ( --��--��--�� )"
         BevelOuter      =   1
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin FPSpreadADO.fpSpread fpSpread_RtLog 
         Height          =   7095
         Left            =   -74640
         TabIndex        =   101
         Top             =   1580
         Width           =   11655
         _Version        =   458752
         _ExtentX        =   20558
         _ExtentY        =   12515
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpreadDesigner  =   "mainFrm.frx":13E4
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   375
         Index           =   2
         Left            =   -74670
         TabIndex        =   102
         Top             =   8765
         Width           =   3060
         _ExtentX        =   5398
         _ExtentY        =   661
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "���� "
         Alignment       =   4
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkTimer_Rt 
            Caption         =   "�ڵ�����"
            Height          =   255
            Left            =   120
            TabIndex        =   105
            Top             =   58
            Width           =   1095
         End
         Begin VB.TextBox schTime_Rt 
            Alignment       =   2  '��� ����
            BackColor       =   &H00404040&
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   1335
            TabIndex        =   104
            Text            =   "30"
            Top             =   30
            Width           =   495
         End
         Begin VB.ComboBox cmbInterval_Rt 
            Height          =   300
            Left            =   1920
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   103
            Top             =   35
            Width           =   630
         End
      End
      Begin Threed.SSPanel Panel_Time_Rt 
         Height          =   255
         Left            =   -71400
         TabIndex        =   106
         Top             =   8840
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   450
         _Version        =   196610
         ForeColor       =   16777215
         BackColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "----�� --�� --�� -- --:--:-- ( --��--��--�� )"
         BevelOuter      =   1
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin FPSpreadADO.fpSpread fpSpread_AgLog 
         Height          =   7095
         Left            =   -74640
         TabIndex        =   107
         Top             =   1580
         Width           =   11655
         _Version        =   458752
         _ExtentX        =   20558
         _ExtentY        =   12515
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpreadDesigner  =   "mainFrm.frx":15B8
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   375
         Index           =   3
         Left            =   -74670
         TabIndex        =   108
         Top             =   8765
         Width           =   3060
         _ExtentX        =   5398
         _ExtentY        =   661
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "���� "
         Alignment       =   4
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.CheckBox chkTimer_Ag 
            Caption         =   "�ڵ�����"
            Height          =   255
            Left            =   120
            TabIndex        =   111
            Top             =   58
            Width           =   1095
         End
         Begin VB.TextBox schTime_Ag 
            Alignment       =   2  '��� ����
            BackColor       =   &H00404040&
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   1320
            TabIndex        =   110
            Text            =   "30"
            Top             =   30
            Width           =   495
         End
         Begin VB.ComboBox cmbInterval_Ag 
            Height          =   300
            Left            =   1920
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   109
            Top             =   35
            Width           =   630
         End
      End
      Begin Threed.SSPanel Panel_Time_Ag 
         Height          =   255
         Left            =   -71400
         TabIndex        =   112
         Top             =   8840
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   450
         _Version        =   196610
         ForeColor       =   16777215
         BackColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "----�� --�� --�� -- --:--:-- ( --��--��--�� )"
         BevelOuter      =   1
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin FPSpreadADO.fpSpread fpSpread_Tot_Usn 
         Height          =   3615
         Left            =   8460
         TabIndex        =   113
         Top             =   5400
         Width           =   3525
         _Version        =   458752
         _ExtentX        =   6218
         _ExtentY        =   6376
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         MaxRows         =   1
         SpreadDesigner  =   "mainFrm.frx":178C
      End
      Begin FPSpreadADO.fpSpread fpSpread_UsnLog 
         Height          =   7095
         Left            =   -74640
         TabIndex        =   114
         Top             =   1580
         Width           =   11655
         _Version        =   458752
         _ExtentX        =   20558
         _ExtentY        =   12515
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpreadDesigner  =   "mainFrm.frx":197C
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   375
         Index           =   4
         Left            =   -74670
         TabIndex        =   115
         Top             =   8765
         Width           =   3060
         _ExtentX        =   5398
         _ExtentY        =   661
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "���� "
         Alignment       =   4
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.ComboBox cmbInterval_Usn 
            Height          =   300
            Left            =   1920
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   118
            Top             =   35
            Width           =   630
         End
         Begin VB.TextBox schTime_Usn 
            Alignment       =   2  '��� ����
            BackColor       =   &H00404040&
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   1335
            TabIndex        =   117
            Text            =   "30"
            Top             =   30
            Width           =   495
         End
         Begin VB.CheckBox chkTimer_Usn 
            Caption         =   "�ڵ�����"
            Height          =   255
            Left            =   120
            TabIndex        =   116
            Top             =   58
            Width           =   1095
         End
      End
      Begin Threed.SSPanel Panel_Time_Usn 
         Height          =   255
         Left            =   -71400
         TabIndex        =   119
         Top             =   8840
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   450
         _Version        =   196610
         ForeColor       =   16777215
         BackColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "----�� --�� --�� -- --:--:-- ( --��--��--�� )"
         BevelOuter      =   1
         Alignment       =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin FPSpreadADO.fpSpread fpSpread2 
         Height          =   7095
         Left            =   -74640
         TabIndex        =   120
         Top             =   1580
         Width           =   11655
         _Version        =   458752
         _ExtentX        =   20558
         _ExtentY        =   12515
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxRows         =   3
         SpreadDesigner  =   "mainFrm.frx":1B50
         Appearance      =   1
      End
      Begin FPSpreadADO.fpSpread ss_Tw 
         Height          =   8145
         Left            =   -74640
         TabIndex        =   121
         Top             =   1575
         Width           =   11655
         _Version        =   458752
         _ExtentX        =   20558
         _ExtentY        =   14367
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         GridColor       =   12632319
         MaxCols         =   26
         MaxRows         =   22
         OperationMode   =   3
         SelectBlockOptions=   2
         ShadowColor     =   12648447
         ShadowDark      =   12632319
         SpreadDesigner  =   "mainFrm.frx":1D32
         Appearance      =   1
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         AutoSize        =   -1  'True
         Caption         =   "����������-VPN"
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
         Index           =   0
         Left            =   375
         TabIndex        =   127
         Top             =   575
         Width           =   2835
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         AutoSize        =   -1  'True
         Caption         =   "����������-CDMA"
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
         Index           =   1
         Left            =   4935
         TabIndex        =   126
         Top             =   830
         Visible         =   0   'False
         Width           =   2880
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         AutoSize        =   -1  'True
         Caption         =   "�ؼ���������"
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
         Index           =   3
         Left            =   8955
         TabIndex        =   125
         Top             =   615
         Width           =   2880
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         AutoSize        =   -1  'True
         Caption         =   "�ؾ������"
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
         Index           =   4
         Left            =   4515
         TabIndex        =   124
         Top             =   7005
         Width           =   2880
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         AutoSize        =   -1  'True
         Caption         =   "USN"
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
         Index           =   5
         Left            =   10110
         TabIndex        =   123
         Top             =   5130
         Width           =   450
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         AutoSize        =   -1  'True
         Caption         =   "�����ؾ��������"
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
         Index           =   2
         Left            =   4815
         TabIndex        =   122
         Top             =   600
         Width           =   2865
      End
   End
   Begin VB.Menu mnuCfg 
      Caption         =   "����"
      Begin VB.Menu mnuDbConf 
         Caption         =   "ȯ�漳��"
      End
   End
   Begin VB.Menu mnuExit 
      Caption         =   "����"
   End
   Begin VB.Menu mnuPop 
      Caption         =   "�˾��޴�"
      Visible         =   0   'False
      Begin VB.Menu mnuPopVisible 
         Caption         =   "�����(&H)"
      End
   End
End
Attribute VB_Name = "mainFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents tt As TrayTool
Attribute tt.VB_VarHelpID = -1
Private strTimer_Value As String

Private Sub btnDtSearch_Click()
    '�˻�
    Call goDtSearch
End Sub

Private Sub btnTwSearch_Click()
    '�˻�
    Call goTwSearch
End Sub
Private Sub btnRtSearch_Click()
    '�˻�
    Call goRtSearch
End Sub
Private Sub btnAgSearch_Click()
    '�˻�
    Call goAgSearch
End Sub
Private Sub btnUsnSearch_Click()
    '�˻�
    Call goUsnSearch
End Sub

Private Sub CboTw_NM_Click()
    
    CboTw_ID.ListIndex = CboTw_NM.ListIndex
    
End Sub

Private Sub CboRTID_NM_Click()
    
    CboRTID_ID.ListIndex = CboRTID_NM.ListIndex
    
End Sub

Private Sub chkTimer_Tw_Click()
    If chkTimer_Tw.Value Then
        dTimes = 0
        TW_Timer.Interval = 1000
        TW_Timer.Enabled = True

    Else
        dTimes = 0
        TW_Timer.Interval = 0
        TW_Timer.Enabled = False

        Panel_Time_Tw.Caption = "----�� --�� --�� -- --:--:-- ( --��--��--�� )"

    End If
End Sub
Private Sub chkTimer_Usn_Click()

    If chkTimer_Usn.Value Then
        dTimes = 0
        USN_Timer.Interval = 1000
        USN_Timer.Enabled = True

    Else
        dTimes = 0
        USN_Timer.Interval = 0
        USN_Timer.Enabled = False

        Panel_Time_Usn.Caption = "----�� --�� --�� -- --:--:-- ( --��--��--�� )"

    End If
End Sub

Private Sub Cmd_Tw_Click()
    
On Error GoTo Err
    
    Dim dtRs As ADODB.Recordset
    Set dtRs = New ADODB.Recordset
    
    Dim stDateHour As String
    Dim edDateHour As String
    Dim i As Integer
    Dim intDate As Integer
        
    '�ʱ�ȭ
    ss_Tw.MaxRows = 0   '�������� �ο찪
    intDate = 0
        
    '�˻����� üũ
    If Not IsDate(mainFrm.txtTwDate_From.Text) Then
        MsgBox "�˻� ���� �������ڸ� Ȯ�����ּ���."
        mainFrm.txtTwDate_From.SetFocus
        Exit Sub
    ElseIf IsNumeric(mainFrm.cboTwhh_From.Text) = False Then
        MsgBox "�˻� ���� �����Ͻø� Ȯ�����ּ���."
        mainFrm.cboTwhh_From.SetFocus
        Exit Sub
    Else
        If mainFrm.cboTwhh_From.Text > 24 Then
            MsgBox "�˻� ���� �����Ͻô� ���ڸ� �Է����ּ���."
            mainFrm.cboTwhh_From.SetFocus
            Exit Sub
        End If
    End If
    
    If Not IsDate(mainFrm.txtTwDate_To.Text) Then
        MsgBox "�˻� ���� �����Ͻø� Ȯ�����ּ���."
        mainFrm.txtTwDate_To.SetFocus
        Exit Sub
    ElseIf IsNumeric(mainFrm.cboTwhh_To.Text) = False Then
        MsgBox "�˻� ���� �����Ͻø� Ȯ�����ּ���."
        mainFrm.cboTwhh_To.SetFocus
        Exit Sub
    Else
        If mainFrm.cboTwhh_To.Text > 24 Then
            MsgBox "�˻� ���� �����Ͻô� ���ڸ� �Է����ּ���."
            mainFrm.cboTwhh_To.SetFocus
            Exit Sub
        End If
    End If
    
    stDateHour = mainFrm.txtTwDate_From.Text & Space(1) & mainFrm.cboTwhh_From
    edDateHour = mainFrm.txtTwDate_To.Text & Space(1) & mainFrm.cboTwhh_To
    
    '12�ð� �ʰ� �˻��� ���ϰ� ����(�ʹ� ������ ���� ������....)
    intDate = DateDiff("h", stDateHour & ":00:00", edDateHour & ":00:00")
    
    If intDate > 12 Then
        MsgBox "�˻��ð��� " & intDate & "�ð��̻� ���̳��ϴ�." & vbCrLf & "12�ð����� �˻� �����մϴ�.", vbCritical, "�˻��ð� ����"
        Exit Sub
    End If
    
    Lb_State.Caption = "�˻����� : �˻����Դϴ�...."
    
    'DB����
    Set AdoDBConn = New ADODB.Connection
    AdoDBConn.Open strAdoDBConn

    strQry = ""
    strQry = strQry & vbCrLf & "SELECT  BUOY_LON,       BUOY_LAT,       OBS_TIME,       WIND_S,     WIND_D,     WIND_G, "
    strQry = strQry & vbCrLf & "        AIR_TEMP,       AIR_PRES,       BUOY_ORIENTATION,           WH,         WP, "
    strQry = strQry & vbCrLf & "        CURRENT_S,      CURRENT_D,      W_TEMP,         CONDUCTIVITY,           SAL, "
    strQry = strQry & vbCrLf & "        EQUIP_ID,       WIND_D_RAW,     Visibility,     BATTERY,    REFERENCE,  TRACK_SEQ, "
    strQry = strQry & vbCrLf & "        MAX_WAVE_PERIOD,                MAX_WAVE_HEIGHT,            WAVE_DIRECT,           ORIGINAL_WIND_D "
    strQry = strQry & vbCrLf & "FROM    T_WRN_TW_BUOY "
    strQry = strQry & vbCrLf & "WHERE   STATION_ID = '" & CboTw_ID.Text & "' "
    strQry = strQry & vbCrLf & "  AND   TO_CHAR(OBS_TIME, 'YYYY-MM-DD HH24') BETWEEN '" & stDateHour & "' AND '" & edDateHour & "' "
    strQry = strQry & vbCrLf & "ORDER BY OBS_TIME DESC"
    
    dtRs.CursorLocation = adUseClient
    dtRs.Open strQry, AdoDBConn

    Do Until dtRs.EOF = True
                
        ss_Tw.MaxRows = ss_Tw.MaxRows + 1
        ss_Tw.Row = ss_Tw.MaxRows
        
        For i = 1 To dtRs.Fields.Count
        
            If IsNull(Trim(dtRs.Fields(i - 1))) = False Then
                ss_Tw.Col = i:     ss_Tw.Text = Trim(dtRs.Fields(i - 1))
            End If
            
        Next i
        
        DoEvents
        
        dtRs.MoveNext
    Loop
    
    Lb_State.Caption = "�˻����� : �˻��Ϸ�"
    
    If dtRs.State = adStateOpen Then dtRs.Close
    If Not dtRs Is Nothing Then Set dtRs = Nothing
    'DB��������
    If AdoDBConn.State = adStateOpen Then
       AdoDBConn.Close
    End If
    
    If Not AdoDBConn Is Nothing Then
        Set AdoDBConn = Nothing
    End If
    
    Exit Sub
    
Err:
    Call Sub_MsgBox(Err.Description, 2)
End Sub

Private Sub cmdRefresh_Tot_Click()
    
    'DB����
    Set AdoDBConn = New ADODB.Connection
    AdoDBConn.Open strAdoDBConn
    
    getJowiVpnList      '����������-VPN
    'getJowiCdmaList     '����������-CDMA
    getTWList           '�����ؾ��������
    getRTList           '�ؾ������
    'getAGList           '�ؼ���������
    'getUSNList          'USN
    
    'DB��������
    If AdoDBConn.State = adStateOpen Then
       AdoDBConn.Close
    End If
    
    If Not AdoDBConn Is Nothing Then
        Set AdoDBConn = Nothing
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim bUnload As Boolean
    Dim i As Integer
    
    Select Case UnloadMode
    Case vbFormControlMenu
        Cancel = -1
        Me.WindowState = vbMinimized
    Case vbFormCode
        Dim msg As String
        
        If bUnload = True Then
            msg = "�ڷ� ���� ����͸� ���α׷��� �����Ͻðڽ��ϱ�?"
            If MsgBox(msg, vbCritical + vbOKCancel, mainFrm.Caption) = vbOK Then
                
                Call UnloadForm
            Else
                Cancel = -1
                Exit Sub
            End If
        Else
            msg = "�ڷ� ���� ����͸� ���α׷��� �����Ͻðڽ��ϱ�?"
            If MsgBox(msg, vbQuestion + vbOKCancel, mainFrm.Caption) = vbOK Then
                Call UnloadForm
                
            Else
                Cancel = -1
                Exit Sub
            End If
        End If
        
        
    Case Else
        Call UnloadForm
    End Select
End Sub
Private Function UnloadForm() As Boolean
    Dim tmr As Long
    Dim i As Integer
    If Not tt Is Nothing Then Set tt = Nothing
    
    Tot_Timer.Enabled = False
    RT_Timer.Enabled = False
    AG_Timer.Enabled = False
    DT_Timer.Enabled = False
    TW_Timer.Enabled = False
    
    tmr = Timer
    Do While tmr > Timer - 1
        DoEvents
    Loop
    
    End
End Function
Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        If gExeFlag = 3 Then
            '// Test3�϶��� ó��
            Call tt.StartTrayProc(Me.hWnd, Me.Icon.Handle, "�ڷ���� ����͸�")
        End If
        tt.Hide
        
        '// �ּ�ȭ �ɶ� �����
        mnuPopVisible.Caption = "���̱�(&S)"
        mnuPopVisible.Tag = "Show"
    
    ElseIf Me.WindowState = vbNormal Then
        If gExeFlag = 3 Then
            '// Test3�϶��� ó��
            Call tt.StopTrayProc
        Else
            tt.Restore
        End If

        '// �����ɶ� ���̱�
        mnuPopVisible.Caption = "�����(&H)"
        mnuPopVisible.Tag = "Hide"
    End If
End Sub
Private Sub Form_Load()

'CCY On Error GoTo ErrorHandler
        
    ' �� �ν��Ͻ��� �ۼ�
    Set tt = New TrayTool
    
    ' Ʈ���̷� ����
    Call tt.StartTrayProc(Me.hWnd, Me.Icon.Handle, "�ڷ���� ����͸�")

    ' �⺻���´� â ������ ����
    mnuPopVisible.Caption = "�����(&H)"
    mnuPopVisible.Tag = "Hide"
    
'���ð� �ʱ�ȭ Start
    '����������(VPN) ���ð�
    strJowiVPNCautionMin = 5
    '����������(CDMA) ���ð�
    strJowiCDMACautionMin = 30
    '�ؾ�������� ���ð�
    strTwCautionMin = 60
    '�ؼ��������� ���ð�
    strAgCautionMin = 60
    '�ؾ������ ���ð�
    strRtCautionMin = 40
    'USN ���ð�
    strUsnCautionMin = 40
'���ð� �ʱ�ȭ End
   
    'ȯ�漳������ �ε�
    Call GetConfig
    
    'chkTimer_Tot = 1
    chkTimer_Tot.Value = strTimer_Value
    
    '�� ������
    Call Sub_mntTab
    Call Sub_SS
    
    '-------------------------
    '�������� �ʱ�ȭ
    '-------------------------
    '��Ȳ�� ��������
    Call Init_fpSpread_Tot_DtVPN     '��Ȳ-����VPN
    'Init_fpSpread_Tot_DtCDMA   '��Ȳ-����CDMA
    Call Init_fpSpread_Tot_Tw    '��Ȳ-�����ؾ��������
    Call Init_fpSpread_Tot_Rt    '��Ȳ-�ؾ������
    Init_fpSpread_Tot_Ag   '��Ȳ-�ؼ���������
    Init_fpSpread_Tot_Usn  '��Ȳ-USN
    
    '�� �α��� ��������
    Call Init_fpSpread_DtLog '����������
    Call Init_fpSpread_TwLog '�����ؾ��������
    Call Init_fpSpread_RtLog '�ؾ������
    'Init_fpSpread_AgLog '�ؼ���������
    'Init_fpSpread_UsnLog    'USN
    
    '-- ���͹� ���� ����
    cmbInterval_Tot.AddItem "��"
    cmbInterval_Tot.AddItem "��"
    cmbInterval_Tot.ListIndex = 0
    
    strAdoDBConn = "Provider=MSDAORA.1;Password=" & CfgDb.ID & ";User ID=" & CfgDb.PW & ";Data Source=" & CfgDb.DataSource & ";Persist Security Info=True"
    'strAdoDBConn = "Provider=OraOLEDB.Oracle.1;Data Source=" & CfgDb.DataSource & ";User Id=" & CfgDb.PW & ";Password=" & CfgDb.ID

    Set AdoDBConn = New ADODB.Connection
    
    AdoDBConn.Open strAdoDBConn
        
    Call getJowiVpnList
    'getJowiCdmaList
    Call getTWList
    Call getRTList
    'getAGList
    'getUSNList

'��-���������� â ���� Start
    '�˻�â�������� ���� ����
    Call setDtStationID
    '-- ���͹� ���� ����
    cmbInterval_Dt.AddItem "��"
    cmbInterval_Dt.AddItem "��"
    cmbInterval_Dt.ListIndex = 0
    
    '-- ��°Ǽ�����
    cmbxDtRownum.AddItem "30"
    cmbxDtRownum.AddItem "50"
    cmbxDtRownum.AddItem "100"
    cmbxDtRownum.AddItem "500"
    cmbxDtRownum.AddItem "1000"
    cmbxDtRownum.AddItem "ALL"
    cmbxDtRownum.ListIndex = 0
'��-���������� â ���� End
    
'��-�����ؾ�������� â ���� Start
    '�˻�â�������� ���� ����
    Call setTwStationID
    '-- ���͹� ���� ����
    cmbInterval_Tw.AddItem "��"
    cmbInterval_Tw.AddItem "��"
    cmbInterval_Tw.ListIndex = 0
    
    '-- ��°Ǽ�����
    cmbxTwRownum.AddItem "30"
    cmbxTwRownum.AddItem "50"
    cmbxTwRownum.AddItem "100"
    cmbxTwRownum.AddItem "500"
    cmbxTwRownum.AddItem "1000"
    cmbxTwRownum.AddItem "ALL"
    cmbxTwRownum.ListIndex = 0
'��-�����ؾ�������� â ���� End
    
'��-�ؾ������ â ���� Start
    '�˻�â�������� ���� ����
    Call setRtStationID
    
    '-- ���͹� ���� ����
    cmbInterval_Rt.AddItem "��"
    cmbInterval_Rt.AddItem "��"
    cmbInterval_Rt.ListIndex = 0
        
    '-- ��°Ǽ�����
    cmbxRtRownum.AddItem "30"
    cmbxRtRownum.AddItem "50"
    cmbxRtRownum.AddItem "100"
    cmbxRtRownum.AddItem "500"
    cmbxRtRownum.AddItem "1000"
    cmbxRtRownum.AddItem "ALL"
    cmbxRtRownum.ListIndex = 0
'��-�����ؾ�������� â ���� End

'��-�����ؾ�������� â ���� Start
    '�˻�â�������� ���� ����
   
    Call setAgStationID
    '-- ���͹� ���� ����
    cmbInterval_Ag.AddItem "��"
    cmbInterval_Ag.AddItem "��"
    cmbInterval_Ag.ListIndex = 0
    
    '-- ��°Ǽ�����
    cmbxAgRownum.AddItem "30"
    cmbxAgRownum.AddItem "50"
    cmbxAgRownum.AddItem "100"
    cmbxAgRownum.AddItem "500"
    cmbxAgRownum.AddItem "1000"
    cmbxAgRownum.AddItem "ALL"
    cmbxAgRownum.ListIndex = 0
'��-�����ؾ�������� â ���� End
LogWrite "-1-"
'��-USN â ���� Start
    '�˻�â�������� ���� ����

    Call setUsnStationID
    '-- ���͹� ���� ����
    cmbInterval_Usn.AddItem "��"
    cmbInterval_Usn.AddItem "��"
    cmbInterval_Usn.ListIndex = 0

    '-- ��°Ǽ�����
    cmbxUsnRownum.AddItem "30"
    cmbxUsnRownum.AddItem "50"
    cmbxUsnRownum.AddItem "100"
    cmbxUsnRownum.AddItem "500"
    cmbxUsnRownum.AddItem "1000"
    cmbxUsnRownum.AddItem "ALL"
    cmbxUsnRownum.ListIndex = 0
LogWrite "-2-"
'��-USN â ���� End
       
    'DB��������
    If AdoDBConn.State = adStateOpen Then
       AdoDBConn.Close
    End If
    
ErrorHandler:
    If Err.Number <> 0 Then
        Call LogWrite("Form_Load : " & Err.Number & "-" & Err.Description)
        Err.Clear
        Call MsgBox("���α׷� ���������� �д� �� ���ܰ� �߻��Ͽ����ϴ�. ȯ�漳���� ���ּ���.", vbCritical + vbOKOnly, mainFrm.Caption)
        frmDbConfig.Show , mainFrm
        'Unload Me
    End If
End Sub

Public Sub GetConfig()
'"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
'" 1. ����       : ���밳����
'" 2. ���������� : ȯ�漳�� �ε�
'" 3. �ۼ���     : ��â��
'" 4. �ۼ���     : 2008/09/17
'" 5. ���ϰ�     :
'" 6. ���� �̷�  :
'"
'"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""

On Error GoTo ErrorHandler

    Dim Fnum As Long
    Dim i As Long
    Dim DataLine As String
    Dim Str() As String
    Dim QCItemCount As Integer
    
    QCItemCount = 0
    
    Fnum = FreeFile

    Open App.Path & "\Config.bin" For Input As #Fnum
    Do
        If EOF(Fnum) = True Then
            Exit Do
        End If
        Line Input #Fnum, DataLine
        Select Case Trim(DataLine)
           
            Case "[DataBaseInfo]"

                'DBInfo
                Line Input #Fnum, DataLine
                Str = Split(DataLine, ",", -1)
                CfgDb.ID = Str(0)
                CfgDb.PW = Str(1)
                CfgDb.DataSource = Str(2)

            Case "[CAUTION]"

                '���ð�����
                Dim cnt As Integer
                For cnt = 0 To 4
                    Line Input #Fnum, DataLine
                    Str = Split(DataLine, "=", -1)
                    If Str(0) = "DT_VPN" Then
                        '����������(VPN) ���ð�
                        strJowiVPNCautionMin = Str(1)
                    ElseIf Str(0) = "DT_CDMA" Then
                        '����������(CDMA) ���ð�
                        strJowiCDMACautionMin = Str(1)
                    ElseIf Str(0) = "TW" Then
                        '�ؾ�������� ���ð�
                        strTwCautionMin = Str(1)
                    ElseIf Str(0) = "AG" Then
                        '�ؼ��������� ���ð�
                        strAgCautionMin = Str(1)
                    ElseIf Str(0) = "RT" Then
                        '�ؾ������ ���ð�
                        strRtCautionMin = Str(1)
                    ElseIf Str(0) = "USN" Then
                        '�ؾ������ ���ð�
                        strUsnCautionMin = Str(1)
                    End If
                Next cnt
            Case "[OPTION]"
                Line Input #Fnum, DataLine
                Str = Split(DataLine, "=", -1)
                If Str(0) = "VPNTimer_Value" Then   '�������� ����͸� �ڵ����� �ɼ�
                    strTimer_Value = Str(1)
                End If

        End Select
        

        DoEvents
    Loop

    Close #Fnum
    
    Exit Sub
    
ErrorHandler:
'    If Err.Number <> 0 Then
''        Call LogWrite("GetConfig : " & Err.Number & "-" & Err.Description)
'        Err.Clear
'        Call MsgBox("���α׷� ���������� �д� �� ���ܰ� �߻��Ͽ����ϴ�. ȯ�漳���� ���ּ���.", vbCritical + vbOKOnly, mainFrm.Caption)
'        frmDbConfig.Show , mainFrm
'        'Unload Me
'    End If
End Sub

Public Sub chkTimer_Tot_Click()

    If chkTimer_Tot.Value Then
        dTimes = 0
        Tot_Timer.Interval = 1000
        Tot_Timer.Enabled = True
        
        'chkTimer_Tot.Caption = "�ڵ�����"
'        mnuStopService.Enabled = True
'        mnuStartService.Enabled = False
    Else
        dTimes = 0
        Tot_Timer.Interval = 0
        Tot_Timer.Enabled = False
        
        'chkTimer_Tot.Caption = "��������"
        Panel_Time_Tot.Caption = "----�� --�� --�� -- --:--:-- ( --��--��--�� )"
'        mnuStopService.Enabled = False
'        mnuStartService.Enabled = True
    End If
End Sub

Private Sub Label1_Click(Index As Integer)
    
    frmSubMonitor.Show vbModal

End Sub

Private Sub mnuDbConf_Click()
    frmDbConfig.Show , mainFrm
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub Tot_Timer_Timer()

On Error GoTo ErrorHandler

    Dim dd As Integer, mm As Integer, ss As Double ' ss As Integer
    Dim MyStr As String
    Dim Ret As Double
    Dim i As Integer
    Dim l_ret As Boolean
    Dim Fnum As Long
    Dim TimeInterval As Long
    
    If cmbInterval_Tot.Text = "��" Then
        TimeInterval = Val(schTime_Tot.Text) * 60
    ElseIf cmbInterval_Tot.Text = "��" Then
        TimeInterval = Val(schTime_Tot.Text)
    End If
    
    l_ret = True
    
    'FTP ���ε� ���� ����
    dTimes = dTimes + 1
    dProcessTimes = dProcessTimes + 1
    
            
    If dTimes > TimeInterval Then
        If Not Busy Then
            dTimes = 0
            '���â �ʱ�ȭ
'            Init_fpSpread_Tot_Ag
'            Init_fpSpread_Tot_DtCDMA
'            Init_fpSpread_Tot_DtVPN
'            Init_fpSpread_Tot_Rt
'            Init_fpSpread_Tot_Tw
                        
'            'DB����
'            Set AdoDBConn = New ADODB.Connection
'            AdoDBConn.Open strAdoDBConn
'
'
'            '�����ȸ
'            getJowiVpnList
'            'getJowiCdmaList
'            getTWList
'            getRTList
'            'getAGList
'
'            'DB���� ����
'            If AdoDBConn.State = adStateOpen Then
'               AdoDBConn.Close
'            End If
'
'            If Not AdoDBConn Is Nothing Then
'                Set AdoDBConn = Nothing
'            End If

            Call cmdRefresh_Tot_Click
        End If
    Else
        Call CalScale_10to60(((TimeInterval) - dTimes) / 60 / 60, dd, mm, ss)
        MyStr = "20" & Format(DateAdd("s", (TimeInterval) - dTimes, Now), "Medium Date")
        MyStr = MyStr & " " & Format(DateAdd("s", (TimeInterval) - dTimes, Now), "AMPM hh:mm:ss")
        Panel_Time_Tot.Caption = MyStr & " ( " & dd & "��" & mm & "��" & ss & "��" & " )"
    End If

Exit Sub
    
ErrorHandler:

End Sub

Private Sub mntTab_Click(PreviousTab As Integer)
    Dim intIndex As Integer

    '���� ���õ� Tab�� ���� ����.
    intIndex = mntTab.Tab
    
    Select Case intIndex
        Case 0
            '����
'            Tot_Timer.Interval = 1000
'            DT_Timer.Interval = 0
'            TW_Timer.Interval = 0
''            RT_Timer.Interval = 0

            chkTimer_Tot = strTimer_Value
            chkTimer_Dt = 0
            chkTimer_Tw = 0
            chkTimer_Rt = 0
            chkTimer_Ag = 0
            chkTimer_Usn = 0
            
        Case 1
            '����������
'            Tot_Timer.Interval = 0
'            DT_Timer.Interval = 1000
'            TW_Timer.Interval = 0
''            RT_Timer.Interval = 0
            chkTimer_Tot = 0
            chkTimer_Dt = 1
            chkTimer_Tw = 0
            chkTimer_Rt = 0
            chkTimer_Ag = 0
            chkTimer_Usn = 0
            
            Call startDTView
            
        Case 2
            '�����ؾ��������
'            Tot_Timer.Interval = 0
'            DT_Timer.Interval = 0
'            TW_Timer.Interval = 1000
''            RT_Timer.Interval = 0

            chkTimer_Tot = 0
            chkTimer_Dt = 0
            chkTimer_Tw = 1
            chkTimer_Rt = 0
            chkTimer_Ag = 0
            chkTimer_Usn = 0
            Call startTWView
                        
        Case 3
            '�ؾ������
'            Tot_Timer.Interval = 0
'            DT_Timer.Interval = 0
''            DT_Timer.Interval = 0
''            RT_Timer.Interval = 1000

            chkTimer_Tot = 0
            chkTimer_Dt = 0
            chkTimer_Tw = 0
            chkTimer_Rt = 1
            chkTimer_Ag = 0
            chkTimer_Usn = 0
            
            Call startRTView
            
        Case 4
            '�ؼ���������
'            Tot_Timer.Interval = 0
'            DT_Timer.Interval = 0
''            DT_Timer.Interval = 0
''            RT_Timer.Interval = 1000

            chkTimer_Tot = 0
            chkTimer_Dt = 0
            chkTimer_Tw = 0
            chkTimer_Rt = 0
            chkTimer_Ag = 1
            chkTimer_Usn = 0
            
            Call startAGView
            
        Case 5
            'USN

            chkTimer_Tot = 0
            chkTimer_Dt = 0
            chkTimer_Tw = 0
            chkTimer_Rt = 0
            chkTimer_Ag = 0
            chkTimer_Usn = 1
            
            Call startUSNView
            
        Case 6
            chkTimer_Tot = 0
            chkTimer_Dt = 0
            chkTimer_Tw = 0
            chkTimer_Rt = 0
            chkTimer_Ag = 0
            chkTimer_Usn = 0
            
            '�˻� �ð� ����
            Call Sub_Tw
    End Select
    
   
End Sub

Public Sub chkTimer_Dt_Click()
    If chkTimer_Dt.Value Then
        dTimes = 0
        DT_Timer.Interval = 1000
        DT_Timer.Enabled = True

        
        'chkTimer_Dt.Caption = "�ڵ�����"
'        mnuStopService.Enabled = True
'        mnuStartService.Enabled = False
    Else
        dTimes = 0
        DT_Timer.Interval = 0
        DT_Timer.Enabled = False

        
        'chkTimer_Dt.Caption = "��������"
        Panel_Time_Dt.Caption = "----�� --�� --�� -- --:--:-- ( --��--��--�� )"
'        mnuStopService.Enabled = False
'        mnuStartService.Enabled = True
    End If
End Sub

Private Sub Dt_Timer_Timer()
    Dim dd As Integer, mm As Integer, ss As Double ' ss As Integer
    Dim MyStr As String
    Dim Ret As Double
    Dim i As Integer
    Dim l_ret As Boolean
    Dim Fnum As Long
    Dim TimeInterval As Long
    
    If cmbInterval_Dt.Text = "��" Then
        TimeInterval = Val(schTime_Dt.Text) * 60
    ElseIf cmbInterval_Dt.Text = "��" Then
        TimeInterval = Val(schTime_Dt.Text)
    End If
    
    l_ret = True
    
    'FTP ���ε� ���� ����
    dTimes = dTimes + 1
    dProcessTimes = dProcessTimes + 1
                
    If dTimes > TimeInterval Then
        If Not Busy Then
            dTimes = 0
            Call goDtSearch
        
        End If
    Else
        Call CalScale_10to60(((TimeInterval) - dTimes) / 60 / 60, dd, mm, ss)
        MyStr = "20" & Format(DateAdd("s", (TimeInterval) - dTimes, Now), "Medium Date")
        MyStr = MyStr & " " & Format(DateAdd("s", (TimeInterval) - dTimes, Now), "AMPM hh:mm:ss")
        Panel_Time_Dt.Caption = MyStr & " ( " & dd & "��" & mm & "��" & ss & "��" & " )"
    End If
End Sub

Private Sub Tw_Timer_Timer()

    Dim dd As Integer, mm As Integer, ss As Double ' ss As Integer
    Dim MyStr As String
    Dim TimeInterval As Long
    
    If cmbInterval_Tw.Text = "��" Then
        TimeInterval = Val(schTime_Tw.Text) * 60
    ElseIf cmbInterval_Tw.Text = "��" Then
        TimeInterval = Val(schTime_Tw.Text)
    End If
        
    'FTP ���ε� ���� ����
    dTimes = dTimes + 1
    dProcessTimes = dProcessTimes + 1
            
    If dTimes > TimeInterval Then
        If Not Busy Then
            dTimes = 0
            goTwSearch
        End If
    Else
        Call CalScale_10to60(((TimeInterval) - dTimes) / 60 / 60, dd, mm, ss)
        MyStr = "20" & Format(DateAdd("s", (TimeInterval) - dTimes, Now), "Medium Date")
        MyStr = MyStr & " " & Format(DateAdd("s", (TimeInterval) - dTimes, Now), "AMPM hh:mm:ss")
        Panel_Time_Tw.Caption = MyStr & " ( " & dd & "��" & mm & "��" & ss & "��" & " )"
    End If
    
End Sub

Public Sub chkTimer_Rt_Click()
    If chkTimer_Rt.Value Then
        dTimes = 0
        RT_Timer.Interval = 1000
        RT_Timer.Enabled = True

        
        'chkTimer_Rt.Caption = "�ڵ�����"
'        mnuStopService.Enabled = True
'        mnuStartService.Enabled = False
    Else
        dTimes = 0
        RT_Timer.Interval = 0
        RT_Timer.Enabled = False

        
        'chkTimer_Rt.Caption = "��������"
        Panel_Time_Rt.Caption = "----�� --�� --�� -- --:--:-- ( --��--��--�� )"
'        mnuStopService.Enabled = False
'        mnuStartService.Enabled = True
    End If
End Sub

Private Sub Rt_Timer_Timer()
    Dim dd As Integer, mm As Integer, ss As Double ' ss As Integer
    Dim MyStr As String
    Dim Ret As Double
    Dim i As Integer
    Dim l_ret As Boolean
    Dim Fnum As Long
    Dim TimeInterval As Long
    
    If cmbInterval_Rt.Text = "��" Then
        TimeInterval = Val(schTime_Rt.Text) * 60
    ElseIf cmbInterval_Rt.Text = "��" Then
        TimeInterval = Val(schTime_Rt.Text)
    End If
    
    l_ret = True
    
    'FTP ���ε� ���� ����
    dTimes = dTimes + 1
    dProcessTimes = dProcessTimes + 1
                
    If dTimes > TimeInterval Then
        If Not Busy Then
            dTimes = 0
            goRtSearch
        End If
    Else
        Call CalScale_10to60(((TimeInterval) - dTimes) / 60 / 60, dd, mm, ss)
        MyStr = "20" & Format(DateAdd("s", (TimeInterval) - dTimes, Now), "Medium Date")
        MyStr = MyStr & " " & Format(DateAdd("s", (TimeInterval) - dTimes, Now), "AMPM hh:mm:ss")
        Panel_Time_Rt.Caption = MyStr & " ( " & dd & "��" & mm & "��" & ss & "��" & " )"
    End If
End Sub


Public Sub chkTimer_Ag_Click()
    If chkTimer_Ag.Value Then
        dTimes = 0
        AG_Timer.Interval = 1000
        AG_Timer.Enabled = True

        
        'chkTimer_Ag.Caption = "�ڵ�����"
'        mnuStopService.Enabled = True
'        mnuStartService.Enabled = False
    Else
        dTimes = 0
        AG_Timer.Interval = 0
        AG_Timer.Enabled = False

        
        'chkTimer_Ag.Caption = "��������"
        Panel_Time_Ag.Caption = "----�� --�� --�� -- --:--:-- ( --��--��--�� )"
'        mnuStopService.Enabled = False
'        mnuStartService.Enabled = True
    End If
End Sub

Private Sub Ag_Timer_Timer()
    Dim dd As Integer, mm As Integer, ss As Double ' ss As Integer
    Dim MyStr As String
    Dim Ret As Double
    Dim i As Integer
    Dim l_ret As Boolean
    Dim Fnum As Long
    Dim TimeInterval As Long
    
    If cmbInterval_Ag.Text = "��" Then
        TimeInterval = Val(schTime_Ag.Text) * 60
    ElseIf cmbInterval_Ag.Text = "��" Then
        TimeInterval = Val(schTime_Ag.Text)
    End If
    
    l_ret = True
    
    'FTP ���ε� ���� ����
    dTimes = dTimes + 1
    dProcessTimes = dProcessTimes + 1
    
            
    If dTimes > TimeInterval Then
        If Not Busy Then
            dTimes = 0
            goAgSearch
        End If
    Else
        Call CalScale_10to60(((TimeInterval) - dTimes) / 60 / 60, dd, mm, ss)
        MyStr = "20" & Format(DateAdd("s", (TimeInterval) - dTimes, Now), "Medium Date")
        MyStr = MyStr & " " & Format(DateAdd("s", (TimeInterval) - dTimes, Now), "AMPM hh:mm:ss")
        Panel_Time_Ag.Caption = MyStr & " ( " & dd & "��" & mm & "��" & ss & "��" & " )"
    End If
End Sub

Private Sub Usn_Timer_Timer()
    Dim dd As Integer, mm As Integer, ss As Double ' ss As Integer
    Dim MyStr As String
    Dim Ret As Double
    Dim i As Integer
    Dim l_ret As Boolean
    Dim Fnum As Long
    Dim TimeInterval As Long
    
    If cmbInterval_Usn.Text = "��" Then
        TimeInterval = Val(schTime_Usn.Text) * 60
    ElseIf cmbInterval_Usn.Text = "��" Then
        TimeInterval = Val(schTime_Usn.Text)
    End If
    
    l_ret = True
    
    'FTP ���ε� ���� ����
    dTimes = dTimes + 1
    dProcessTimes = dProcessTimes + 1
                
    If dTimes > TimeInterval Then
        If Not Busy Then
            dTimes = 0
            goUsnSearch
        End If
    Else
        Call CalScale_10to60(((TimeInterval) - dTimes) / 60 / 60, dd, mm, ss)
        MyStr = "20" & Format(DateAdd("s", (TimeInterval) - dTimes, Now), "Medium Date")
        MyStr = MyStr & " " & Format(DateAdd("s", (TimeInterval) - dTimes, Now), "AMPM hh:mm:ss")
        Panel_Time_Usn.Caption = MyStr & " ( " & dd & "��" & mm & "��" & ss & "��" & " )"
    End If
End Sub

Private Sub Sub_mntTab()
    
'    mntTab.TabVisible(4) = False
'    mntTab.TabVisible(5) = False
    mntTab.TabVisible(7) = False
    mntTab.Tab = 0
    
    '�˻�â�� �ð� ���� ����
    Call setDtSechCondition
    Call setTwSechCondition
    Call setRtSechCondition
    Call setAgSechCondition
    Call setUsnSechCondition
    Call Sub_SetTwDate
    Call Sub_SetRTIDDate
    
End Sub

Private Sub Sub_SS()
    
    '�������� ������
    Dim i As Integer
        
    '�ؾ���� View �������� �ش� ��Ʈ ����
    With ss_Tw
        .Row = 0
        
        For i = 1 To .MaxCols
            .Col = i
            .FontBold = True
        Next i
        
    End With
    
End Sub

