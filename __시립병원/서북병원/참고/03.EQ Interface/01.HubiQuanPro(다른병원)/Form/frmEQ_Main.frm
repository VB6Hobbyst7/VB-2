VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmEQ_Main 
   Caption         =   "Hi Interface EQ"
   ClientHeight    =   10095
   ClientLeft      =   1110
   ClientTop       =   5790
   ClientWidth     =   15000
   BeginProperty Font 
      Name            =   "����ü"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEQ_Main.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10095
   ScaleWidth      =   15000
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   6300
      TabIndex        =   38
      Top             =   60
      Width           =   2235
   End
   Begin VB.TextBox txtBuff 
      Height          =   1755
      Left            =   5820
      MultiLine       =   -1  'True
      TabIndex        =   37
      Top             =   6420
      Width           =   6195
   End
   Begin VB.TextBox txtSerialData 
      Height          =   5175
      Left            =   15180
      MultiLine       =   -1  'True
      ScrollBars      =   2  '����
      TabIndex        =   3
      Top             =   4500
      Width           =   4995
   End
   Begin FPSpread.vaSpread sprDResult 
      Height          =   3075
      Left            =   5700
      TabIndex        =   7
      Top             =   1020
      Width           =   9255
      _Version        =   393216
      _ExtentX        =   16325
      _ExtentY        =   5424
      _StockProps     =   64
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   11
      MaxRows         =   10
      SpreadDesigner  =   "frmEQ_Main.frx":263A
      UserResize      =   1
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "ȭ������(&C)"
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
      Left            =   12660
      Style           =   1  '�׷���
      TabIndex        =   6
      Top             =   60
      Width           =   1275
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "����(&X)"
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
      Left            =   13980
      Style           =   1  '�׷���
      TabIndex        =   5
      Top             =   60
      Width           =   975
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   9660
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   5
      DTREnable       =   -1  'True
      InputLen        =   1
      RThreshold      =   1
   End
   Begin MSComctlLib.ProgressBar prgPatient 
      Height          =   75
      Left            =   60
      TabIndex        =   2
      Top             =   600
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   132
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar staCondition 
      Align           =   2  '�Ʒ� ����
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   9720
      Width           =   15000
      _ExtentX        =   26458
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   5398
            MinWidth        =   3528
            Text            =   "Copyright �� 2010 Medimate Corp."
            TextSave        =   "Copyright �� 2010 Medimate Corp."
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10954
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            Picture         =   "frmEQ_Main.frx":2D72
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1588
            MinWidth        =   1058
            Text            =   "Local DB"
            TextSave        =   "Local DB"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1270
            MinWidth        =   1058
            Text            =   "HIS DB"
            TextSave        =   "HIS DB"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "COM"
            TextSave        =   "COM"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1940
            MinWidth        =   1940
            TextSave        =   "2011-10-16"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "���� 7:14"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin FPSpread.vaSpread sprLResult 
      Height          =   5175
      Left            =   60
      TabIndex        =   4
      Top             =   4500
      Width           =   14895
      _Version        =   393216
      _ExtentX        =   26273
      _ExtentY        =   9128
      _StockProps     =   64
      BackColorStyle  =   1
      ColsFrozen      =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   16
      MaxRows         =   20
      OperationMode   =   2
      SelectBlockOptions=   0
      SpreadDesigner  =   "frmEQ_Main.frx":32E9
   End
   Begin VB.Line Line1 
      X1              =   60
      X2              =   5640
      Y1              =   1380
      Y2              =   1380
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "Sample No"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   17
      Left            =   120
      TabIndex        =   40
      Top             =   1500
      Width           =   945
   End
   Begin VB.Label lblSAMPLENO 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "1234567890"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   1260
      TabIndex        =   39
      Top             =   1500
      Width           =   900
   End
   Begin VB.Label lblDISKNOPOSNO 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "1234567890"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   1260
      TabIndex        =   36
      Top             =   1740
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "Rack/Pos"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   16
      Left            =   120
      TabIndex        =   35
      Top             =   1740
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "�˻�ȸ��"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   20
      Left            =   4260
      TabIndex        =   34
      Top             =   1080
      Width           =   780
   End
   Begin VB.Label lblEXSEQ 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   5160
      TabIndex        =   33
      Top             =   1080
      Width           =   120
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  '�������� ����
      Height          =   255
      Index           =   5
      Left            =   9240
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "Delta"
      Height          =   180
      Index           =   15
      Left            =   9540
      TabIndex        =   32
      Top             =   4260
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "Low"
      Height          =   180
      Index           =   14
      Left            =   7920
      TabIndex        =   31
      Top             =   4260
      Width           =   270
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  '�������� ����
      Height          =   255
      Index           =   4
      Left            =   7620
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "Panic"
      Height          =   180
      Index           =   13
      Left            =   10320
      TabIndex        =   30
      Top             =   4260
      Width           =   450
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  '�������� ����
      Height          =   255
      Index           =   2
      Left            =   10020
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "High"
      Height          =   180
      Index           =   12
      Left            =   8700
      TabIndex        =   29
      Top             =   4260
      Width           =   360
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFF80&
      BackStyle       =   1  '�������� ����
      Height          =   255
      Index           =   1
      Left            =   8400
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "��� ����"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   11
      Left            =   120
      TabIndex        =   28
      Top             =   2460
      Width           =   885
   End
   Begin VB.Label lblSDDT 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "1234567890"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   1260
      TabIndex        =   27
      Top             =   2460
      Width           =   900
   End
   Begin VB.Shape shpCon 
      BackStyle       =   1  '�������� ����
      FillColor       =   &H00FF0000&
      FillStyle       =   0  '�ܻ�
      Height          =   375
      Index           =   1
      Left            =   4800
      Top             =   180
      Width           =   135
   End
   Begin VB.Shape shpCon 
      BackStyle       =   1  '�������� ����
      FillColor       =   &H000000FF&
      FillStyle       =   0  '�ܻ�
      Height          =   375
      Index           =   0
      Left            =   60
      Top             =   60
      Width           =   135
   End
   Begin VB.Label lblRCDT 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "1234567890"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   1260
      TabIndex        =   26
      Top             =   2220
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "��� ����"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   10
      Left            =   120
      TabIndex        =   25
      Top             =   2220
      Width           =   885
   End
   Begin VB.Label lblORDGB 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "1234567890"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   1260
      TabIndex        =   24
      Top             =   3180
      Width           =   900
   End
   Begin VB.Label lblORDDT 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "1234567890"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   1260
      TabIndex        =   23
      Top             =   2940
      Width           =   900
   End
   Begin VB.Label lblSEXAGE 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "1234567890"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   1260
      TabIndex        =   22
      Top             =   3900
      Width           =   900
   End
   Begin VB.Label lblPATNM 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "1234567890"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   1260
      TabIndex        =   21
      Top             =   3660
      Width           =   900
   End
   Begin VB.Label lblPATNO 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "1234567890"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   1260
      TabIndex        =   20
      Top             =   3420
      Width           =   900
   End
   Begin VB.Label lblEXDT 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "1234567890"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   1260
      TabIndex        =   19
      Top             =   1980
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "�ǽð� �˻縮��Ʈ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   8
      Left            =   180
      TabIndex        =   18
      Top             =   4200
      Width           =   1620
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "ó�� ����"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   7
      Left            =   120
      TabIndex        =   17
      Top             =   3180
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "ó�� ����"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   6
      Left            =   120
      TabIndex        =   16
      Top             =   2940
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "ó�� ����"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   5
      Left            =   120
      TabIndex        =   15
      Top             =   1980
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "����/����"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   4
      Left            =   120
      TabIndex        =   14
      Top             =   3900
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "������ ��"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   3
      Left            =   120
      TabIndex        =   13
      Top             =   3660
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "���� ��ȣ"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   12
      Top             =   3420
      Width           =   885
   End
   Begin VB.Label lblBARCD 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "1234567890"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1260
      TabIndex        =   11
      Top             =   1080
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "��ü ��ȣ"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "��ü��ȣ�� ��������"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   180
      TabIndex        =   9
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "��ü��ȣ�� �˻���"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   9
      Left            =   5820
      TabIndex        =   8
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label lbl���� 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "�˻�����"
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
      Left            =   300
      TabIndex        =   1
      Top             =   60
      Width           =   1800
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  '�������� ����
      BorderColor     =   &H00000000&
      FillColor       =   &H00FF8080&
      FillStyle       =   0  '�ܻ�
      Height          =   495
      Index           =   3
      Left            =   60
      Shape           =   4  '�ձ� �簢��
      Top             =   60
      Width           =   4875
   End
   Begin VB.Shape shpDResult 
      BackColor       =   &H00808000&
      BackStyle       =   1  '�������� ����
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFC0C0&
      FillStyle       =   5  '���� �밢��
      Height          =   255
      Left            =   5700
      Shape           =   4  '�ձ� �簢��
      Top             =   720
      Width           =   9255
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808000&
      BackStyle       =   1  '�������� ����
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFC0C0&
      FillStyle       =   5  '���� �밢��
      Height          =   255
      Index           =   0
      Left            =   60
      Shape           =   4  '�ձ� �簢��
      Top             =   720
      Width           =   5595
   End
   Begin VB.Shape shpLResult 
      BackColor       =   &H00808000&
      BackStyle       =   1  '�������� ����
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFC0C0&
      FillStyle       =   5  '���� �밢��
      Height          =   255
      Left            =   60
      Shape           =   4  '�ձ� �簢��
      Top             =   4200
      Width           =   6915
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File    "
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuSetting 
      Caption         =   "ȯ�漳��    "
      Begin VB.Menu mnuSettingSub 
         Caption         =   "��ż���"
         Index           =   0
      End
      Begin VB.Menu mnuSettingSub 
         Caption         =   "HIS DB ��������"
         Index           =   1
      End
      Begin VB.Menu mnuSettingSub 
         Caption         =   "ETC DB ��������"
         Index           =   2
      End
      Begin VB.Menu mnuSettingSub 
         Caption         =   "��Ž�ȣ ����"
         Index           =   3
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuJob 
      Caption         =   "�۾�    "
      Begin VB.Menu mnuJobSub 
         Caption         =   "WorkList �۾�"
         Index           =   0
      End
      Begin VB.Menu mnuJobSub 
         Caption         =   "�˻��� ����"
         Index           =   1
      End
      Begin VB.Menu mnuJobSub 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuJobSub 
         Caption         =   "���۹��"
         Index           =   4
         Begin VB.Menu mnuJobModeAuto 
            Caption         =   "�ڵ�����[Auto]"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuJobModeManual 
            Caption         =   "��������[Manual]"
         End
      End
   End
   Begin VB.Menu mnuCode 
      Caption         =   "�����ڵ�    "
      Begin VB.Menu mnuCodeSub 
         Caption         =   "���˻��ڵ� ����"
         Index           =   0
      End
   End
   Begin VB.Menu mnuInfo 
      Caption         =   "����"
   End
End
Attribute VB_Name = "frmEQ_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lngMeHeight     As Long '/Me.Height�� �ʱⰪ
Dim lngMeWidth      As Long '/Me.Width�� �ʱⰪ
Dim strOneLine      As String


Private Type ConWhere   ' ����� ���� ������ ����ϴ�.
   Nm       As String
   Left     As Long
   Top      As Long
   Width    As Long
   Height   As Long
End Type
Dim CW()    As ConWhere

Public Function CHK_COMM_PORT() As Boolean
    CHK_COMM_PORT = False
    
On Error GoTo RTN_ERR_PORT

RE_CHK:

    MSComm1.CommPort = gtypEQ_INFO.SERIALPORT
    MSComm1.RTSEnable = gtypEQ_INFO.SERIALRTS
    MSComm1.DTREnable = gtypEQ_INFO.SERIALDTR
    MSComm1.Settings = gtypEQ_INFO.SERIALBAUD & "," & gtypEQ_INFO.SERIALPARITY & "," & gtypEQ_INFO.SERIALDATABIT & "," & gtypEQ_INFO.SERIALSTOPBIT

    If MSComm1.PortOpen = False Then MSComm1.PortOpen = True
    
    CHK_COMM_PORT = True
Exit Function

'/----------------------------------------------------------------------------------------------------/

RTN_ERR_PORT:
    If Err = 8002 Then      'Port
        If MsgBox("��������ż��� Info" & vbCrLf & vbCrLf & _
                  "MSComm Port Setting �� �ùٸ��� �ʽ��ϴ�." & vbCrLf & _
                  "(��)�����ϰڽ��ϱ�?", vbQuestion + vbYesNo, "����") = vbNo Then
            
            MsgBox "��������ż��� Info" & vbCrLf & vbCrLf & _
                   "��� ������ ��� �Ϻ� ����� ���ѵ˴ϴ�." & vbCrLf & _
                   "�������� ���α׷� ����� ���� ����� Ȥ�� ���޾�ü�� �����ֽñ� �ٶ��ϴ�.", vbInformation, "Ȯ��"
                
            Exit Function
        Else
            frmEQ����_Set_Port.Show vbModal
            
            GoTo RE_CHK
        End If
    Else
        Resume Next
    End If
End Function

Public Function FUNC_LOC_VIEW(ArgSection As Integer) As Boolean
    Dim stró���ڵ�     As String
    
    FUNC_LOC_VIEW = False
    
On Error GoTo RTN_ERR
    
    
    
'''    If ConnDB_LOC(gstrREG_DB_CONSTR) = True Then
'''        '/����ڵ庰 ó���ڵ� ��������
'''        gstrQuy = "SELECT ORDCD "
'''        gstrQuy = gstrQuy & vbCrLf & "  FROM MM_EMR_EQORD "
'''        gstrQuy = gstrQuy & vbCrLf & " WHERE EQUIPCODE = '" & gtypEQ_INFO.EQUIPCODE & "' "
'''        If ReadSQL_LOC(gstrQuy, ADR_LOC) = False Then Call CloseDB_LOC: End
'''
'''        If Not ADR_LOC Is Nothing Then
'''            Do Until ADR_LOC.EOF
'''                stró���ڵ� = stró���ڵ� & ",'" & Trim(ADR_LOC!ORDCD & "") & "'"
'''
'''                ADR_LOC.MoveNext
'''            Loop
'''            ADR_LOC.Close: Set ADR_LOC = Nothing
'''
'''            stró���ڵ� = Mid(stró���ڵ�, 2)
'''        End If
    
    
    FUNC_LOC_VIEW = True

Exit Function

'/----------------------------------------------------------------------------------------------------/

RTN_ERR:

End Function

Public Sub SUB_COMM_PART_HUBIQUANPRO_BAR(argCOMM_BF As String)
    '/Patient ID �� ���ڵ��� ���� ó���Ѵ�.
    
    '/�Ϲݰ˻� Sample
    'H|humasis|HUBI-QUAN pro|HP90003|46
    'P|110928-0001|20110928091712|P|CARDIAC 3/1|10-006
    'R1|CK-MB|0.00~5.00|MYO|0.00~100.00|TNI|0.00~0.40|
    'R2|CK-MB|>30.00|ng/mL| |
    'R2|MYO|>150.00|ng/mL| |
    'R2|TNI|7.00|ng/mL| |
    'L|1|N
    
    '/QC Sample
    'H|humasis|HUBI-QUAN pro|HP90003|
    'P|110928-0002|20110928095439|CARDIAC 3/1|10-006
    'R|CK-MB|ng/mL|10.53|Low|29.24|High
    'R|MYO|ng/mL|68.66|Low|>150.00|High
    'R|TNI|ng/mL|2.39|Low|6.87|High
    'L|1|N
    
    Dim strCrlf '/�迭����(CRLF ����)
    Dim strPart '/�迭����(| ����)
    Dim strLotNo '/Lot No
    Dim strEXSEQ            As String '/EXSEQ �� ��������
    Dim intLResultRow       As Integer
    Dim intLResultTarRow    As Integer
    Dim intLResultTarCol    As Integer
    Dim intCol              As Integer
    Dim intLineCnt          As Integer
    
    strCrlf = Split(argCOMM_BF, vbCrLf)
    
    For intLineCnt = 0 To UBound(strCrlf) - 1
        strPart = Split(strCrlf(intLineCnt), "|")
        
        Select Case strPart(0)
            Case "H" '/Hearder����(�����ȣ����)
                '/ó������
            
            Case "P" '/Patient����(ȯ������)
                '/strPart(0): P
                '/strPart(1): Test No(Barcode) YYDDMM-XXX1 ��¥�� SEQNO
                '/strPart(2): ��¥�ð�(�˻�����Ͻ�?, ��������Ͻ�?) YYYYMMDDHHMMSS
                '/strPart(3): Patient ID (�Էµ��� ������ P �� ǥ�õ�)
                '/strPart(4): Device Name(�ڵ带 �а� ��ǰ���� ��µȴ�)
                '/strPart(5): Lot No
                
                gtypPAT_RES.SAMPLENO = Trim(strPart(1)) '/Sample No(Test No)
                
                '/Patient ID �� ���ڵ��� ���
                gtypPAT_RES.BARCD = Trim(strPart(3)) '/BARCD(��ü��ȣ(Barcode))
                '/Patient ID �� ���ڵ��� ���
                
                gtypPAT_RES.EXDT = Trim(Left(strPart(2), 8)) '/EXDT(�˻�ó����������(YYYYMMDD) HIEQ->�Ƿ����)
                gtypPAT_RES.EXTM = Trim(Mid(strPart(2), 9)) '/EXTM(�˻�ó�����۽ð�(24HHMMSS) HIEQ->�Ƿ����)
                
                strLotNo = Split(strPart(5), "-")
                gtypPAT_RES.DISKNO = strLotNo(0) '/DISKNO(LotNo �� �պκ�)
                gtypPAT_RES.POSNO = strLotNo(1) '/POSNO(LotNo �� �޺κ�)
                
                Call FUNC_HIS_PATIENT '/HIS ȯ������ ��������
                
                '/��ü��ȣ/SampleNo/Rack/Pos�� ���ǵ� ���¿��� ������ ��.
                If strEXSEQ <> "Y" Then
                    gtypPAT_RES.EXSEQ = FUNC_GET_EXSEQ(gtypPAT_RES.BARCD) '/��ü��ȣ(Barcode)�� �˻�ȸ��
                    strEXSEQ = "Y"
                End If
                
                '/�ǽð� �˻縮��Ʈ �����ֱ�----------------------------------------------------------------------------------------------------/
                intLResultTarRow = 0
                For intLResultRow = 1 To sprLResult.DataRowCnt
                    If Trim(GET_CELL(sprLResult, 1, intLResultRow)) = gtypPAT_RES.BARCD And _
                       Trim(GET_CELL(sprLResult, 2, intLResultRow)) = gtypPAT_RES.EXSEQ And _
                       Trim(GET_CELL(sprLResult, 3, intLResultRow)) = gtypPAT_RES.SAMPLENO And _
                       Trim(GET_CELL(sprLResult, 4, intLResultRow)) = gtypPAT_RES.DISKNO And _
                       Trim(GET_CELL(sprLResult, 5, intLResultRow)) = gtypPAT_RES.POSNO Then
                       
                        intLResultTarRow = intLResultRow
                        Exit For
                    End If
                Next intLResultRow
            
                If intLResultTarRow = 0 Then
                    sprLResult.MaxRows = sprLResult.MaxRows + 1
                    intLResultTarRow = sprLResult.MaxRows
                    
                    Call SET_CELL(sprLResult, 1, intLResultTarRow, gtypPAT_RES.BARCD)
                    Call SET_CELL(sprLResult, 2, intLResultTarRow, gtypPAT_RES.EXSEQ)
                    Call SET_CELL(sprLResult, 3, intLResultTarRow, gtypPAT_RES.SAMPLENO)
                    Call SET_CELL(sprLResult, 4, intLResultTarRow, gtypPAT_RES.DISKNO)
                    Call SET_CELL(sprLResult, 5, intLResultTarRow, gtypPAT_RES.POSNO)
                    Call SET_CELL(sprLResult, 8, intLResultTarRow, IIf(gtypPAT_RES.EXDT <> "", Format(gtypPAT_RES.EXDT, "@@@@-@@-@@"), "") & IIf(gtypPAT_RES.EXTM <> "", " " & Format(gtypPAT_RES.EXTM, "@@:@@:@@"), ""))
                    Call SET_CELL(sprLResult, 13, intLResultTarRow, gtypPAT_RES.PATNO)
                    Call SET_CELL(sprLResult, 14, intLResultTarRow, gtypPAT_RES.PATNM)
                    If gtypPAT_RES.PATSEX <> "" Or gtypPAT_RES.PATAGE <> "" Then
                        Call SET_CELL(sprLResult, 15, intLResultTarRow, gtypPAT_RES.PATSEX & "/" & gtypPAT_RES.PATAGE)
                    End If
                End If
                '/�ǽð� �˻縮��Ʈ �����ֱ�----------------------------------------------------------------------------------------------------/
                
            Case "R" '/�˻�������(QC)
                
            Case "R1" '/�����������(�Ϲݰ˻�)
                '/ó������
            
            Case "R2" '/�˻�������(�Ϲݰ˻�)
                '/strPart(0): R2
                '/strPart(1): ���˻��ڵ�
                '/strPart(2): ���˻���
                '/strPart(3): �������
                
                'R2|CK-MB|>30.00|ng/mL| |
            
                gtypPAT_RES.EQCD = strPart(1) '/EQCD(���˻��ڵ�)
                
                gtypPAT_RES.EQRESULT = strPart(2) '/EQRESULT(�����ð��)
                gtypPAT_RES.Result = FUNC_RESULT_CHANGE(gtypPAT_RES.EQCD, gtypPAT_RES.EQRESULT) '/RESULT(�˻���(������ ���))
                gtypPAT_RES.RCDT = Format(Now, "YYYYMMDD") '/RCDT(�˻�����������(YYYYMMDD) �Ƿ���� ->HIEQ)
                gtypPAT_RES.RCTM = Format(Now, "HHMMSS") '/RCTM(�˻������Žð�(24HHMMSS) �Ƿ���� ->HIEQ)
                gtypPAT_RES.STATEFLAG = "1" '/STATEFLAG(���������� (0:ó��, 1:���))
                
                Call FUNC_HIS_ORDER_VIEW    '/ó�泻�� ��ȸ
                Call FUNC_HIS_RESULT_JUDGMENT   '/��� ����
                
                gtypPAT_RES.SENDFLAG = "0"
                    
                '/�ǽð� �˻縮��Ʈ �����ֱ�----------------------------------------------------------------------------------------------------/
                Call SET_CELL(sprLResult, 9, intLResultTarRow, IIf(gtypPAT_RES.RCDT <> "", Format(gtypPAT_RES.RCDT, "@@@@-@@-@@"), "") & IIf(gtypPAT_RES.RCTM <> "", " " & Format(gtypPAT_RES.RCTM, "@@:@@:@@"), "")) '/RCDT(�˻�����������(YYYYMMDD) �Ƿ���� ->HIEQ)
                Call SET_CELL(sprLResult, 6, intLResultTarRow, IIf(gtypPAT_RES.STATEFLAG = "1", "���", "ó��"))
                Call SET_CELL(sprLResult, 6, intLResultTarRow, IIf(gtypPAT_RES.SENDFLAG = "1", "�Ϸ�", "���"))
                Call SET_CELL(sprLResult, 11, intLResultTarRow, gtypPAT_RES.ORDDT) '/ORDDT(ó������)
                Select Case gtypPAT_RES.ORDGB '/ORDGB(ó������(O.�ܷ�, I.�Կ�, G.�ǰ�����))
                    Case "O": Call SET_CELL(sprLResult, 12, intLResultTarRow, "�ܷ�")
                    Case "I": Call SET_CELL(sprLResult, 12, intLResultTarRow, "�Կ�")
                    Case "G": Call SET_CELL(sprLResult, 12, intLResultTarRow, "����")
                End Select
                
                '/���˻��׸� Column Set/���
                For intCol = gintEQ_StartCol To sprLResult.MaxCols
                    If GET_CELL(sprLResult, intCol, -1000) = gtypPAT_RES.EQCD Then
                        Call SET_CELL(sprLResult, intCol, intLResultTarRow, gtypPAT_RES.Result) '/RESULT(�˻���(������ ���))
                        
                        sprLResult.Col = intCol
                        sprLResult.Row = intLResultTarRow
                        
                        If gtypPAT_RES.AFLAG = "L" Then
                            sprLResult.BackColor = &HFFFF&
                        End If
                        If gtypPAT_RES.AFLAG = "H" Then
                            sprLResult.BackColor = &HFFFF80
                        End If
                        If gtypPAT_RES.DFLAG = "D" Then
                            sprLResult.BackColor = &HFF8080
                        End If
                        If gtypPAT_RES.PFLAG = "P" Then
                            sprLResult.BackColor = &HFF&
                        End If

                        Exit For
                    End If
                Next intCol
                '/�ǽð� �˻縮��Ʈ �����ֱ�----------------------------------------------------------------------------------------------------/
                
                '/�ش� Row�� Focus �̵�
                sprLResult.Col = 1
                sprLResult.Row = intLResultTarRow
                sprLResult.Action = ActionActiveCell
                
                '/�ش� Row �ڷ� ���������� �˻��� ǥ��
                Call sprLResult_LeaveRow(intLResultTarRow - 1, False, False, False, intLResultTarRow, False, False)
                
                '/���� �ڷ� Local ����
                If FUNC_LOC_SAVE_PAT_RES = True Then
                    If mnuJobModeAuto.Checked = True Then '/���۹���� �ڵ������̸�...
                        Call FUNC_HIS_SAVE '/HIS�� ��� ����
                        Call FUNC_LOC_SAVE_SEND(gtypPAT_RES.BARCD, gtypPAT_RES.EXSEQ, gtypPAT_RES.EQCD, gtypPAT_RES.SAMPLENO, gtypPAT_RES.DISKNO, gtypPAT_RES.POSNO, "1") '/HIS�� ��� ����
                    End If
                End If
            Case "L" '/��������
                '/ó������
    
                gtypPAT_RES.BARCD = ""
                gtypPAT_RES.EXSEQ = ""            '/EXSEQ(��ü��ȣ(Barcode)�� �˻�ȸ��)
                gtypPAT_RES.EQCD = ""             '/EQCD(���˻��ڵ�)
                gtypPAT_RES.EXAMCD = ""           '/EXAMCD(ó���ڵ�(HIS or LIS�� �˻��ڵ�))
                gtypPAT_RES.EXDT = ""             '/EXDT(�˻�ó����������(YYYYMMDD) HIEQ->�Ƿ����)
                gtypPAT_RES.EXTM = ""             '/EXTM(�˻�ó�����۽ð�(24HHMMSS) HIEQ->�Ƿ����)
                gtypPAT_RES.RCDT = ""             '/RCDT(�˻�����������(YYYYMMDD) �Ƿ���� ->HIEQ)
                gtypPAT_RES.RCTM = ""             '/RCTM(�˻������Žð�(24HHMMSS) �Ƿ���� ->HIEQ)
                gtypPAT_RES.SDDT = ""             '/SDDT(�˻�����������(YYYYMMDD) HIEQ->HIS)
                gtypPAT_RES.SDTM = ""             '/SDTM(�˻������۽ð�(24HHMMSS) HIEQ->HIS)
                gtypPAT_RES.Result = ""           '/RESULT(�˻���(������ ���))
                gtypPAT_RES.EQRESULT = ""         '/EQRESULT(�����ð��)
                gtypPAT_RES.AFLAG = ""            '/AFLAG(Abnormal(��������ġ ���� (H)High or (L)Low �� ǥ��))
                gtypPAT_RES.PFLAG = ""            '/PFLAG(Panic)
                gtypPAT_RES.DFLAG = ""            '/DFLAG(Delta)
                gtypPAT_RES.SAMPLENO = ""         '/Sample No(AU2700, Uriscan � ���)
                gtypPAT_RES.DISKNO = ""           '/DISKNO(��ũ��ȣ or ����ȣ)
                gtypPAT_RES.POSNO = ""            '/POSNO(��ġ��ȣ)
                gtypPAT_RES.ORDDT = ""            '/ORDDT(ó������)
                gtypPAT_RES.ORDGB = ""            '/ORDGB(ó������(O.�ܷ�, I.�Կ�, G.�ǰ�����))
                gtypPAT_RES.PATNO = ""            '/PATNO(���Ϲ�ȣ)
                gtypPAT_RES.PATNM = ""            '/PATNM(�����ڸ�)
                gtypPAT_RES.PATSEX = ""           '/PATSEX(����)
                gtypPAT_RES.PATAGE = ""           '/PATAGE(����)
                gtypPAT_RES.SENDFLAG = ""         '/SENDFLAG(HIS ���� FLAG (0:���, 1:�Ϸ�))
                gtypPAT_RES.STATEFLAG = ""        '/STATEFLAG(���������� (0:ó��, 1:���))
        End Select
    Next intLineCnt
End Sub

Public Sub SUB_MM_CANCEL()
    lbl���� = ""
    prgPatient.Max = 100
    prgPatient.Value = 100
    
    Call SUB_MM_KEY_CLEAR("1") '/��ü��ȣ�� ��������
    Call SUB_MM_KEY_CLEAR("2") '/��ü��ȣ�� �˻���
    Call SUB_MM_KEY_CLEAR("3") '/�ǽð� �˻縮��Ʈ
    
    mnuSetting.Visible = False '/�����޴� �Ⱥ��̱�

    txtSerialData.Visible = False
End Sub

Public Sub SUB_MM_INITIAL()
    '/Resize�� ���� �ʱ� Size Setting----------------------------------------------------------------------------------------------------/
    For intX = 0 To Me.Count - 1
        Select Case True
            Case TypeOf Me.Controls(intX) Is Timer
            Case TypeOf Me.Controls(intX) Is Menu
            Case TypeOf Me.Controls(intX) Is Line
            Case TypeOf Me.Controls(intX) Is MSComm
            Case TypeOf Me.Controls(intX) Is CommonDialog
            Case Else
                ReDim Preserve CW(intX)
                
                CW(intX).Nm = Me.Controls(intX).Name
                CW(intX).Left = Me.Controls(intX).Left
                CW(intX).Top = Me.Controls(intX).Top
                CW(intX).Width = Me.Controls(intX).Width
                CW(intX).Height = Me.Controls(intX).Height
        End Select
    Next intX
    
    '/Form Size Setting
    lngMeHeight = 10890
    lngMeWidth = 15150
    
    Me.Height = lngMeHeight
    Me.Width = lngMeWidth
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Show
    '/Resize�� ���� �ʱ� Size Setting----------------------------------------------------------------------------------------------------/
    
    '/�ʱ� �ڷ� Setting----------------------------------------------------------------------------------------------------/
    GoSub ADD_ITEM
    '/�ʱ� �ڷ� Setting----------------------------------------------------------------------------------------------------/
    
    '/���� ��Ʈ�� �ʱ�ȭ----------------------------------------------------------------------------------------------------/
    Call SUB_MM_CANCEL
    '/���� ��Ʈ�� �ʱ�ȭ----------------------------------------------------------------------------------------------------/
    
    '/��ũ����Ʈ �۾�����(Y.�����, N.������)
    If gtypEQ_INFO.WORKLISTGB = "Y" Then
        mnuJobSub(0).Visible = True
    Else
        mnuJobSub(0).Visible = False
    End If
    
    '/�۾����(A.�ڵ�, M.����)
    If gtypEQ_INFO.AUTOGB = "Y" Then
        mnuJobModeAuto.Checked = True
        staCondition.Panels.Item(3).Picture = LoadPicture(App.Path & "\Auto.jpg")
        mnuJobModeManual.Checked = False
    Else
        mnuJobModeManual.Checked = True
        staCondition.Panels.Item(3).Picture = LoadPicture(App.Path & "\Manual.jpg")
        mnuJobModeAuto.Checked = False
    End If
    
    Me.Caption = Me.Caption & " For " & App.FileDescription
    Me.Caption = Me.Caption & Space(10) & "(�����: " & gtypUSER.USERNM & " )"
    
    lbl���� = App.FileDescription

    '/�۾� ���� Check----------------------------------------------------------------------------------------------------/
    If ConnDB_HIS = True Then
        Call CloseDB_HIS
        staCondition.Panels.Item(5).Enabled = True '/HISDB Ȱ��ȭ
    Else
        staCondition.Panels.Item(5).Enabled = False '/HISDB ��Ȱ��ȭ
    End If
    
    If CHK_COMM_PORT = True Then
        staCondition.Panels.Item(6).Enabled = True '/COM Port Ȱ��ȭ
    Else
        staCondition.Panels.Item(6).Enabled = False '/COM Port ��Ȱ��ȭ
    End If
    '/�۾� ���� Check----------------------------------------------------------------------------------------------------/
Exit Sub

'/----------------------------------------------------------------------------------------------------/

ADD_ITEM:
    Dim intCnt   As Integer
    
    '/��ü��ȣ�� �˻��� Title Clear
    '/�˻��|���˻��ڵ�|���|Wall
    sprDResult.ClearRange 1, -1, 2, -1, True
    sprDResult.ClearRange 4, -1, 5, -1, True
    sprDResult.ClearRange 7, -1, 8, -1, True
    sprDResult.ClearRange 10, -1, 11, -1, True
    
    If sprLResult.MaxCols > gintEQ_StartCol - 1 Then sprLResult.MaxCols = gintEQ_StartCol - 1
    
    If ConnDB_LOC = False Then
        MsgBox "������Local DataBase Info" & vbCrLf & vbCrLf & _
               "Local DataBase �� ������ �� �����ϴ�." & vbCrLf & _
               "�������� ���α׷� ����� ���� ����� Ȥ�� ���޾�ü�� �����ֽñ� �ٶ��ϴ�.", vbInformation, "Ȯ��"
        End
    Else
        frmEQ_Main.staCondition.Panels.Item(4).Enabled = True '/Main ȭ���� Local ���� Ȱ��ȭ

        gstrQuy = "SELECT * "
        gstrQuy = gstrQuy & vbCrLf & "  FROM EQ_MST "
        gstrQuy = gstrQuy & vbCrLf & " ORDER BY EQSEQ "
        If ReadSQL_LOC(gstrQuy, ADR_LOC) = False Then End
        
        If Not ADR_LOC Is Nothing Then
            Do Until ADR_LOC.EOF
                sprLResult.MaxCols = sprLResult.MaxCols + 1
                sprLResult.Col = sprLResult.MaxCols
                sprLResult.Row = -1
                sprLResult.BackColor = RGB(255, 255, 230)
                sprLResult.CellType = CellTypeStaticText
                sprLResult.TypeHAlign = TypeHAlignRight
                sprLResult.TypeVAlign = TypeVAlignCenter

                sprLResult.Row = -1000
                sprLResult.Text = Trim(ADR_LOC!EQCD & "")
                sprLResult.RowHidden = True

                sprLResult.Row = -999
                sprLResult.Text = Trim(ADR_LOC!EQNM & "")
                
                intCnt = intCnt + 1             '/�ǰ��� �˻縮��Ʈ �˻��׸� �б� ����
        
                '/��ü��ȣ�� �˻��� Column
                Select Case intCnt
                    Case 1 To 10:  sprDResult.Col = 1
                    Case 11 To 20: sprDResult.Col = 4
                    Case 21 To 30: sprDResult.Col = 7
                    Case 31 To 40: sprDResult.Col = 10
                End Select
                
                '/��ü��ȣ�� �˻��� Row
                If (intCnt Mod 10) = 0 Then
                    sprDResult.Row = 10
                Else
                    sprDResult.Row = intCnt Mod 10
                End If
                sprDResult.Text = Trim(ADR_LOC!EQNM & "")
                
                ADR_LOC.MoveNext
            Loop
            
            ADR_LOC.Close: Set ADR_LOC = Nothing
        End If
        
        Call CloseDB_LOC
    End If
Return
End Sub

Public Sub SUB_MM_KEY_CLEAR(ArgSection As String)
    Select Case ArgSection
        Case "1" '/��ü��ȣ�� ��������
            lblBARCD = ""
            lblEXSEQ = ""
            lblSAMPLENO = ""
            lblDISKNOPOSNO = ""
            lblEXDT = ""
            lblRCDT = ""
            lblSDDT = ""
            lblORDDT = ""
            lblORDGB = ""
            lblPATNO = ""
            lblPATNM = ""
            lblSEXAGE = ""
            
        Case "2" '/��ü��ȣ�� �˻���
            '/(lCol As Long, lRow As Long, lCol2 As Long, lRow2 As Long, bDataOnly As Boolean)
            sprDResult.ClearRange 2, -1, 2, -1, True
            sprDResult.ClearRange 5, -1, 5, -1, True
            sprDResult.ClearRange 8, -1, 8, -1, True
            sprDResult.ClearRange 11, -1, 11, -1, True
            
        Case "3": '/�ǽð� �˻縮��Ʈ
            If sprLResult.MaxRows > 0 Then sprLResult.MaxRows = 0
    End Select
End Sub

Public Sub SUB_MM_PRINT()
'''    Dim strFont1  As String
'''    Dim strFont2  As String
'''    Dim strHead1  As String
'''
'''    If sprVIEW.MaxRows = 0 Then MsgBox "����� �ڷᰡ �����ϴ�.", vbInformation, "Ȯ��": Exit Function
'''
'''    If MsgBox("����ϰڽ��ϱ�?", vbQuestion + vbOKCancel, "��¿���") = vbCancel Then Exit Function
'''
'''    strFont1 = "/fn""����ü""/fz""15""/fb1/fi0/fu1/fk0/fs1"
'''    strFont2 = "/fn""����ü""/fz""10""/fb0/fi0/fu0/fk0/fs2"
'''
'''    strHead1 = "/f1/c" & "�ŷ�ó �ڵ�" & "/n/n/n"
'''
'''    With sprVIEW
'''        .PrintAbortMsg = "�ŷ�ó �ڵ� ��� ��..."
'''        .PrintHeader = strFont1 + strHead1 + strFont2
'''        .PrintFooter = "/c" & "PAGE : " & "/P"
'''        .PrintBorder = True
'''        .PrintGrid = True
'''        .PrintColHeaders = True
'''        .PrintRowHeaders = True
'''        .PrintColor = False
'''        .PrintMarginTop = 500
'''        .PrintMarginBottom = 500
'''        .PrintMarginLeft = 500
'''        .PrintMarginRight = 0
'''        .PrintType = PrintTypeAll
'''        .PrintShadows = False
'''        .PrintUseDataMax = False
'''        .Action = ActionSmartPrint
'''    End With
End Sub

Private Sub cmdClear_Click()
    Call SUB_MM_KEY_CLEAR("1") '/��ü��ȣ�� ��������
    Call SUB_MM_KEY_CLEAR("2") '/��ü��ȣ�� �˻���
    Call SUB_MM_KEY_CLEAR("3") '/�ǽð� �˻縮��Ʈ
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    MsgBox InStr(txtBuff, chrCR)
    Call SUB_COMM_PART_HUBIQUANPRO_BAR(txtBuff)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim ShiftDown, AltDown, CtrlDown, Txt
   
    ShiftDown = (Shift And vbShiftMask) > 0
    AltDown = (Shift And vbAltMask) > 0
    CtrlDown = (Shift And vbCtrlMask) > 0
   
    If KeyCode = vbKeyM Then   ' Ű�� ���� ���¸� ����մϴ�.
        If mnuSetting.Visible = True Then
            mnuSetting.Visible = False
        Else
            mnuSetting.Visible = True
        End If
    End If
End Sub

Private Sub Form_Load()
    Call SUB_MM_INITIAL
    
    DoEvents
    DoEvents
    DoEvents
End Sub

Private Sub Form_Resize()
    Dim intCnt  As Integer

On Error Resume Next
    '/object.Move Left, Top, Width, Height
    '/(((Me.Height - lngMeHeight) / 3) * 2) : ���̰� �þ�� ��ü 3��, �����λ� �ش� ��ü ���� �þ ��ü�� 2��
    For intCnt = 0 To UBound(CW)
        Select Case CW(intCnt).Nm
            Case cmdClear.Name:     cmdClear.Move CW(intCnt).Left + (Me.Width - lngMeWidth), CW(intCnt).Top, CW(intCnt).Width, CW(intCnt).Height
            Case cmdExit.Name:      cmdExit.Move CW(intCnt).Left + (Me.Width - lngMeWidth), CW(intCnt).Top, CW(intCnt).Width, CW(intCnt).Height
            Case prgPatient.Name:   prgPatient.Move CW(intCnt).Left, CW(intCnt).Top, CW(intCnt).Width + (Me.Width - lngMeWidth), CW(intCnt).Height
            Case shpDResult.Name:   shpDResult.Move CW(intCnt).Left, CW(intCnt).Top, CW(intCnt).Width + (Me.Width - lngMeWidth), CW(intCnt).Height
            Case sprDResult.Name:   sprDResult.Move CW(intCnt).Left, CW(intCnt).Top, CW(intCnt).Width + (Me.Width - lngMeWidth), CW(intCnt).Height
            Case shpLResult.Name:   shpLResult.Move CW(intCnt).Left, CW(intCnt).Top, CW(intCnt).Width + (Me.Width - lngMeWidth), CW(intCnt).Height
            Case sprLResult.Name:   sprLResult.Move CW(intCnt).Left, CW(intCnt).Top, CW(intCnt).Width + (Me.Width - lngMeWidth), CW(intCnt).Height + (Me.Height - lngMeHeight)
        End Select
    Next intCnt
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CloseDB_LOC
    Call CloseDB_HIS
    Call CloseDB_ETC
    
    If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
    
    Set frmEQ_Main = Nothing
End Sub

Private Sub mnuCodeSub_Click(Index As Integer)
    Select Case Index
        Case 0:
            MsgBox "�Ƿ����� ��� �߿� ���˻��ڵ� ������ �����ϸ�" & vbCrLf & _
                   "�ǵ����� ���� ����� �ʷ��� �� �ֽ��ϴ�." & vbCrLf & vbCrLf & _
                   "���˻��ڵ� ������ ������ �Ŀ� ���α׷��� �� �����Ͻʽÿ�", vbExclamation, "����"
        
            frmEQ����_���˻��ڵ����_��ȸ.Show vbModal
    End Select
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuInfo_Click()
    frmEQ����_Info.Show vbModal
End Sub

Private Sub mnuJobModeAuto_Click()
    If mnuJobModeAuto.Checked = False Then
        If MsgBox("�˻����� ���� HIS(���������ý���)���� ���۹���� �ڵ�����[Auto] (��)�� �ϰڽ��ϱ�?" & vbCrLf & vbCrLf & _
                  "(����: �Ƿ����� ����� �϶��� ���۹���� �ٲ��� ���ʽÿ�!)", vbQuestion + vbOKCancel + vbDefaultButton2, "���۹�� ���� Ȯ��") = vbCancel Then Exit Sub
        mnuJobModeAuto.Checked = True
        staCondition.Panels.Item(3).Picture = LoadPicture(App.Path & "\Auto.jpg")
        mnuJobModeManual.Checked = False
    End If
End Sub

Private Sub mnuJobModeManual_Click()
    If mnuJobModeManual.Checked = False Then
        If MsgBox("�˻����� ���� HIS(���������ý���)���� ���۹���� ��������[Manual] (��)�� �ϰڽ��ϱ�?" & vbCrLf & vbCrLf & _
                  "(����: �Ƿ����� ����� �϶��� ���۹���� �ٲ��� ���ʽÿ�!)", vbQuestion + vbOKCancel + vbDefaultButton2, "���۹�� ���� Ȯ��") = vbCancel Then Exit Sub
        mnuJobModeManual.Checked = True
        staCondition.Panels.Item(3).Picture = LoadPicture(App.Path & "\Manual.jpg")
        mnuJobModeAuto.Checked = False
    End If
End Sub

Private Sub mnuJobSub_Click(Index As Integer)
    Select Case Index
        Case 0: 'frmWorkList.Show vbModal
        Case 1: frmEQ_�˻�������.Show vbModal
    End Select
End Sub

Private Sub mnuSettingSub_Click(Index As Integer)
    Select Case Index
        Case 0: frmEQ����_Set_Port.Show vbModal
        Case 1: gstrArgTemp1 = "HIS": frmEQ����_Set_DB.Show vbModal
        Case 2: gstrArgTemp1 = "ETC": frmEQ����_Set_DB.Show vbModal
        Case 3:
            txtSerialData.Left = 7800
            txtSerialData.Top = 2775
            If txtSerialData.Visible = False Then
                txtSerialData.Visible = True
            Else
                txtSerialData.Visible = False
            End If
        End Select
End Sub

Private Sub MSComm1_OnComm()
    Dim strOneByte  As String
    Dim strBuff     As String
    
    If shpCon(0).FillColor = &HFF& Then
        shpCon(0).FillColor = &HFF0000
    Else
        shpCon(0).FillColor = &HFF&
    End If

    If shpCon(1).FillColor = &HFF& Then
        shpCon(1).FillColor = &HFF0000
    Else
        shpCon(1).FillColor = &HFF
    End If

    strOneByte = MSComm1.Input

    txtBuff = txtBuff & strOneByte          '/��ü ����
    strOneLine = strOneLine & strOneByte    '/�� ���� ��� ����

    Select Case strOneByte
        Case chrLF
            If Mid(strOneLine, 1, 1) = "L" Then
                strBuff = txtBuff
                
                txtBuff = ""
                strOneLine = ""
                
                Call SUB_COMM_PART_HUBIQUANPRO_BAR(strBuff)
            Else
                strOneLine = ""
            End If
    End Select
End Sub

Private Sub sprLResult_LeaveRow(ByVal Row As Long, ByVal RowWasLast As Boolean, ByVal RowChanged As Boolean, ByVal AllCellsHaveData As Boolean, ByVal NewRow As Long, ByVal NewRowIsLast As Long, Cancel As Boolean)
    If Cancel = True Then Exit Sub  '/���Ƿ� Cancel�� True�� ����� ȣ���� ��� ó������ �ʰ� �� �� �ִ�.
    If NewRow = Row Then Exit Sub   '/Row�� �������� ������ ó������ �ʴ´�.
    If NewRow < 1 Then Exit Sub     '/������ Row�� ��ȿ�� �������� ������ ó������ �ʴ´�.
    
    Dim intLResultCol   As Integer
    Dim intCnt          As Integer
    Dim strResult       As String
    
    Call SUB_MM_KEY_CLEAR("1") '/��ü��ȣ�� ��������
            
    lblBARCD = GET_CELL(sprLResult, 1, NewRow)
    lblEXSEQ = GET_CELL(sprLResult, 2, NewRow)
    lblSAMPLENO = GET_CELL(sprLResult, 3, NewRow)
    lblDISKNOPOSNO = GET_CELL(sprLResult, 4, NewRow) & "/" & GET_CELL(sprLResult, 5, NewRow)
    
    lblEXDT = GET_CELL(sprLResult, 8, NewRow)
    lblRCDT = GET_CELL(sprLResult, 9, NewRow)
    lblSDDT = GET_CELL(sprLResult, 10, NewRow)
    
    lblORDDT = GET_CELL(sprLResult, 11, NewRow)
    lblORDGB = GET_CELL(sprLResult, 12, NewRow)
    
    lblPATNO = GET_CELL(sprLResult, 13, NewRow)
    lblPATNM = GET_CELL(sprLResult, 14, NewRow)
    lblSEXAGE = GET_CELL(sprLResult, 15, NewRow)
    
    
    
    Call SUB_MM_KEY_CLEAR("2") '/��ü��ȣ�� �˻���
    
    For intLResultCol = gintEQ_StartCol To sprLResult.MaxCols
        '/�б�----------------------------------------------------------------------------------------------------/
        sprLResult.Col = intLResultCol
        sprLResult.Row = NewRow         '/�ǽð� �˻縮��Ʈ �˻��� Row
        strResult = sprLResult.Text     '/�ǽð� �˻縮��Ʈ �˻��� ��
        '/�б�----------------------------------------------------------------------------------------------------/

        '/����----------------------------------------------------------------------------------------------------/
        intCnt = intCnt + 1             '/�ǰ��� �˻縮��Ʈ �˻��׸� �б� ����

        '/��ü��ȣ�� �˻��� Column
        Select Case intCnt
            Case 1 To 10:  sprDResult.Col = 2
            Case 11 To 20: sprDResult.Col = 5
            Case 21 To 30: sprDResult.Col = 8
            Case 31 To 40: sprDResult.Col = 11
        End Select

        '/��ü��ȣ�� �˻��� Row
        If (intCnt Mod 10) = 0 Then
            sprDResult.Row = 10
        Else
            sprDResult.Row = intCnt Mod 10
        End If

        sprDResult.Text = strResult
        '/����----------------------------------------------------------------------------------------------------/
    Next intLResultCol
End Sub

Private Sub staCondition_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel = "COM" Then
        If txtSerialData.Visible = False Then
            txtSerialData.Visible = True
        Else
            txtSerialData.Visible = False
        End If
    End If
End Sub
