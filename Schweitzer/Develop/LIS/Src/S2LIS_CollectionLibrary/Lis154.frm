VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
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
   ScaleHeight     =   9060
   ScaleWidth      =   14535
   WindowState     =   2  '�ִ�ȭ
   Begin VB.Frame Frame4 
      BackColor       =   &H0080FFFF&
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   7860
      Left            =   5400
      TabIndex        =   44
      Top             =   60
      Width           =   7005
      Begin VB.Frame Frame5 
         Height          =   825
         Left            =   90
         TabIndex        =   81
         Top             =   3750
         Width           =   6810
         Begin VB.CheckBox Check1 
            Caption         =   "ǳ��"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   20
            Left            =   1575
            TabIndex        =   88
            Top             =   225
            Width           =   1335
         End
         Begin VB.CheckBox Check1 
            Caption         =   "�����÷�"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   19
            Left            =   135
            TabIndex        =   87
            Top             =   225
            Width           =   1335
         End
         Begin VB.CheckBox Check1 
            Caption         =   "���÷翣��"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   23
            Left            =   2790
            TabIndex        =   86
            Top             =   225
            Width           =   1335
         End
         Begin VB.CheckBox Check1 
            Caption         =   "������"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   30
            Left            =   135
            TabIndex        =   85
            Top             =   480
            Width           =   1335
         End
         Begin VB.CheckBox Check1 
            Caption         =   "�������ռ�����"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   31
            Left            =   1575
            TabIndex        =   84
            Top             =   480
            Width           =   1665
         End
         Begin VB.CheckBox Check1 
            Caption         =   "��Ÿ"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   32
            Left            =   3420
            TabIndex        =   83
            Top             =   480
            Width           =   1335
         End
         Begin VB.CheckBox Check1 
            Caption         =   "���༺���ϼ���"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   33
            Left            =   4410
            TabIndex        =   82
            Top             =   225
            Width           =   1875
         End
      End
      Begin VB.Frame Frame7 
         Height          =   1215
         Left            =   90
         TabIndex        =   66
         Top             =   2505
         Width           =   6810
         Begin VB.CheckBox Check1 
            Caption         =   "���ռ�����"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   18
            Left            =   5205
            TabIndex        =   80
            Top             =   585
            Width           =   1335
         End
         Begin VB.CheckBox Check1 
            Caption         =   "��ƼǪ��"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   17
            Left            =   3855
            TabIndex        =   79
            Top             =   585
            Width           =   1065
         End
         Begin VB.CheckBox Check1 
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   16
            Left            =   2595
            TabIndex        =   78
            Top             =   585
            Width           =   525
         End
         Begin VB.CheckBox Check1 
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   15
            Left            =   1560
            TabIndex        =   77
            Top             =   585
            Width           =   525
         End
         Begin VB.CheckBox Check1 
            Caption         =   "CRAB(IRAB)"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   13
            Left            =   5205
            TabIndex        =   76
            Top             =   240
            Width           =   1395
         End
         Begin VB.CheckBox Check1 
            Caption         =   "C.diffic"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   12
            Left            =   3855
            TabIndex        =   75
            Top             =   240
            Width           =   885
         End
         Begin VB.CheckBox Check1 
            Caption         =   "VRE"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   11
            Left            =   2595
            TabIndex        =   74
            Top             =   240
            Width           =   885
         End
         Begin VB.CheckBox Check1 
            Caption         =   "MRSA"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   10
            Left            =   1560
            TabIndex        =   73
            Top             =   240
            Width           =   885
         End
         Begin VB.CheckBox Check1 
            Caption         =   "anti_HAVIgM"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   9
            Left            =   135
            TabIndex        =   72
            Top             =   240
            Width           =   1380
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Rotavirus"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   14
            Left            =   135
            TabIndex        =   71
            Top             =   585
            Width           =   1200
         End
         Begin VB.CheckBox Check1 
            Caption         =   "CRE"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   26
            Left            =   135
            TabIndex        =   70
            Top             =   900
            Width           =   1125
         End
         Begin VB.CheckBox Check1 
            Caption         =   "VRSA"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   27
            Left            =   1560
            TabIndex        =   69
            Top             =   900
            Width           =   1005
         End
         Begin VB.CheckBox Check1 
            Caption         =   "CJD"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   28
            Left            =   2595
            TabIndex        =   68
            Top             =   900
            Width           =   1125
         End
         Begin VB.CheckBox Check1 
            Caption         =   "��Ÿ"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   29
            Left            =   3855
            TabIndex        =   67
            Top             =   900
            Width           =   1125
         End
      End
      Begin VB.Frame Frame6 
         Height          =   795
         Left            =   90
         TabIndex        =   59
         Top             =   1680
         Width           =   6810
         Begin VB.CheckBox Check1 
            Caption         =   "anti_HBc IgM"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   8
            Left            =   5205
            TabIndex        =   65
            Top             =   225
            Width           =   1455
         End
         Begin VB.CheckBox Check1 
            Caption         =   "anti_HCV"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   7
            Left            =   3870
            TabIndex        =   64
            Top             =   225
            Width           =   1275
         End
         Begin VB.CheckBox Check1 
            Caption         =   "HIV"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   4
            Left            =   180
            TabIndex        =   63
            Top             =   225
            Width           =   1065
         End
         Begin VB.CheckBox Check1 
            Caption         =   "HBsAg"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   6
            Left            =   2565
            TabIndex        =   62
            Top             =   225
            Width           =   1095
         End
         Begin VB.CheckBox Check1 
            Caption         =   "VDRL"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   5
            Left            =   1305
            TabIndex        =   61
            Top             =   225
            Width           =   900
         End
         Begin VB.CheckBox Check1 
            Caption         =   "��Ÿ"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   25
            Left            =   180
            TabIndex        =   60
            Top             =   510
            Width           =   1125
         End
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   90
         TabIndex        =   58
         Text            =   "Caution ������ ���������ǿ� ��û�Ͽ� �ֽʽÿ�."
         Top             =   6825
         Width           =   6795
      End
      Begin VB.Frame Frame10 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   90
         TabIndex        =   52
         Top             =   870
         Width           =   6795
         Begin VB.CheckBox Check1 
            Caption         =   "AFB"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   165
            TabIndex        =   57
            Top             =   210
            Width           =   1080
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Tb"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   1335
            TabIndex        =   56
            Top             =   210
            Width           =   1185
         End
         Begin VB.CheckBox Check1 
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   2565
            TabIndex        =   55
            Top             =   210
            Width           =   1125
         End
         Begin VB.CheckBox Check1 
            Caption         =   "ȫ��"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   3855
            TabIndex        =   54
            Top             =   210
            Width           =   1125
         End
         Begin VB.CheckBox Check1 
            Caption         =   "����������"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   24
            Left            =   165
            TabIndex        =   53
            Top             =   510
            Width           =   1335
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Drug Allergy"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   90
         TabIndex        =   48
         Top             =   4605
         Width           =   6795
         Begin VB.CheckBox Check1 
            Caption         =   "RadioContrast"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   22
            Left            =   1575
            TabIndex        =   51
            Top             =   225
            Width           =   1650
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Penicillin"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   21
            Left            =   180
            TabIndex        =   50
            Top             =   225
            Width           =   1335
         End
         Begin VB.TextBox txtDrug 
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   180
            TabIndex        =   49
            Text            =   "Text1"
            Top             =   570
            Width           =   6465
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Ư�̼Ұ�"
         Enabled         =   0   'False
         Height          =   975
         Left            =   90
         TabIndex        =   46
         Top             =   5790
         Width           =   6795
         Begin RichTextLib.RichTextBox RichText 
            Height          =   540
            Left            =   150
            TabIndex        =   47
            Top             =   300
            Width           =   6495
            _ExtentX        =   11456
            _ExtentY        =   953
            _Version        =   393217
            ScrollBars      =   2
            TextRTF         =   $"Lis154.frx":0000
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
      Begin VB.CommandButton Command1 
         Caption         =   "�� ��"
         Height          =   495
         Left            =   5250
         TabIndex        =   45
         Top             =   7245
         Width           =   1665
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   18
         Left            =   3720
         TabIndex        =   89
         TabStop         =   0   'False
         Top             =   180
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "���������"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   19
         Left            =   3720
         TabIndex        =   90
         TabStop         =   0   'False
         Top             =   510
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "���������"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblWDt 
         Height          =   300
         Left            =   5040
         TabIndex        =   91
         TabStop         =   0   'False
         Top             =   180
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   529
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblWNm 
         Height          =   300
         Left            =   5040
         TabIndex        =   92
         TabStop         =   0   'False
         Top             =   510
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   529
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FF8080&
      Height          =   3750
      Left            =   2160
      ScaleHeight     =   3690
      ScaleWidth      =   10350
      TabIndex        =   41
      Top             =   2880
      Visible         =   0   'False
      Width           =   10410
      Begin VB.CommandButton Command2 
         Caption         =   "����"
         Height          =   600
         Left            =   8505
         TabIndex        =   42
         Top             =   2925
         Width           =   1500
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H00FFFFFF&
         Caption         =   "HIV"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   120
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   2625
         Left            =   270
         TabIndex        =   43
         Top             =   180
         Width           =   9735
      End
   End
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
      Height          =   330
      Left            =   1020
      Style           =   1  '�׷���
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   30
      Width           =   300
   End
   Begin VB.CheckBox chkCollect 
      BackColor       =   &H00800000&
      Caption         =   "ä���� �˻�"
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
      Height          =   315
      Left            =   2265
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   60
      Width           =   2100
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   195
      Index           =   4
      Left            =   7710
      TabIndex        =   29
      Top             =   120
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   344
      BackColor       =   8388608
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
      Caption         =   "���� ��ȸ"
      Appearance      =   0
   End
   Begin MSComCtl2.DTPicker dtpQDt 
      Height          =   300
      Left            =   5790
      TabIndex        =   28
      Top             =   45
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
      Format          =   24051715
      UpDown          =   -1  'True
      CurrentDate     =   36851.6291666667
   End
   Begin MedControls1.LisLabel lblBar 
      Height          =   315
      Left            =   4425
      TabIndex        =   16
      Top             =   2580
      Width           =   10020
      _ExtentX        =   17674
      _ExtentY        =   556
      BackColor       =   8388608
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
      TabIndex        =   20
      Top             =   2820
      Width           =   10035
      Begin VB.PictureBox picOrdDiv 
         Appearance      =   0  '���
         BackColor       =   &H00DBE6E6&
         BorderStyle     =   0  '����
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2145
         ScaleHeight     =   300
         ScaleWidth      =   3690
         TabIndex        =   25
         Top             =   180
         Width           =   3690
         Begin VB.Label lblBBS 
            AutoSize        =   -1  'True
            BackStyle       =   0  '����
            Caption         =   "��������"
            ForeColor       =   &H00404040&
            Height          =   225
            Left            =   1485
            TabIndex        =   26
            Top             =   45
            Width           =   720
         End
         Begin VB.Shape shp1 
            BackColor       =   &H00553755&
            BackStyle       =   1  '�������� ����
            BorderColor     =   &H00C0C0C0&
            Height          =   165
            Index           =   1
            Left            =   150
            Shape           =   3  '����
            Top             =   60
            Width           =   330
         End
         Begin VB.Label lblLIS 
            AutoSize        =   -1  'True
            BackStyle       =   0  '����
            Caption         =   "�ӻ󺴸�"
            ForeColor       =   &H00404040&
            Height          =   225
            Left            =   435
            TabIndex        =   27
            Top             =   45
            Width           =   720
         End
         Begin VB.Shape shp1 
            BackColor       =   &H00496835&
            BackStyle       =   1  '�������� ����
            BorderColor     =   &H00C0C0C0&
            Height          =   165
            Index           =   2
            Left            =   1185
            Shape           =   3  '����
            Top             =   60
            Width           =   330
         End
         Begin VB.Shape shp1 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  '�������� ����
            BorderColor     =   &H00808080&
            FillColor       =   &H00E0E0E0&
            Height          =   300
            Index           =   0
            Left            =   0
            Shape           =   4  '�ձ� �簢��
            Top             =   0
            Width           =   2325
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
         TabIndex        =   22
         Top             =   180
         Width           =   1470
      End
      Begin VB.CheckBox chkChangeColTm 
         BackColor       =   &H00DBE6E6&
         Caption         =   "ä��ð����� : "
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
         TabIndex        =   21
         Top             =   180
         Width           =   1500
      End
      Begin FPSpread.vaSpread tblOrdSheet 
         Height          =   5070
         Left            =   90
         TabIndex        =   23
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
         MaxCols         =   36
         MaxRows         =   19
         ProcessTab      =   -1  'True
         Protect         =   0   'False
         ScrollBars      =   2
         ShadowColor     =   14737632
         ShadowDark      =   14737632
         ShadowText      =   0
         SpreadDesigner  =   "Lis154.frx":007F
         StartingColNumber=   2
         VirtualRows     =   24
         VisibleCols     =   5
         VisibleRows     =   19
      End
      Begin MSComCtl2.DTPicker dtpColDtTm 
         Height          =   300
         Left            =   8010
         TabIndex        =   24
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
         Format          =   24051715
         UpDown          =   -1  'True
         CurrentDate     =   36851.6291666667
      End
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ä��/���� (&S)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   10500
      Style           =   1  '�׷���
      TabIndex        =   19
      Tag             =   "0"
      Top             =   8535
      Width           =   1320
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
      Height          =   510
      Left            =   11820
      Style           =   1  '�׷���
      TabIndex        =   18
      Tag             =   "0"
      Top             =   8535
      Width           =   1320
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
      Height          =   510
      Left            =   13140
      Style           =   1  '�׷���
      TabIndex        =   17
      Tag             =   "0"
      Top             =   8535
      Width           =   1320
   End
   Begin MedControls1.LisLabel LisLabel5 
      Height          =   315
      Left            =   4425
      TabIndex        =   0
      Top             =   45
      Width           =   10020
      _ExtentX        =   17674
      _ExtentY        =   556
      BackColor       =   8388608
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
   Begin MSComctlLib.ListView lvwPtList 
      Height          =   7680
      Left            =   60
      TabIndex        =   10
      Top             =   1350
      Width           =   4320
      _ExtentX        =   7620
      _ExtentY        =   13547
      View            =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
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
   Begin MedControls1.LisLabel lblWardId 
      Height          =   315
      Left            =   1320
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   60
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   556
      BackColor       =   8388608
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
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   330
      Left            =   75
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   45
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   582
      BackColor       =   8388608
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
      Caption         =   "��������"
   End
   Begin VB.Frame fraSearch 
      BackColor       =   &H00DBE6E6&
      Height          =   1020
      Left            =   75
      TabIndex        =   1
      Tag             =   "136"
      Top             =   285
      Width           =   4305
      Begin VB.CommandButton cmdFind 
         Caption         =   "��ȸ"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3210
         TabIndex        =   93
         Top             =   570
         Width           =   915
      End
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
         Left            =   1935
         TabIndex        =   4
         Tag             =   "15304"
         Top             =   210
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
         Left            =   2535
         TabIndex        =   3
         Tag             =   "15305"
         Top             =   225
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
         Left            =   60
         MaxLength       =   10
         TabIndex        =   2
         Top             =   180
         Width           =   1830
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   195
         Index           =   0
         Left            =   1395
         TabIndex        =   94
         Top             =   660
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   344
         BackColor       =   14411494
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
         Alignment       =   2
         Caption         =   "To"
         Appearance      =   0
      End
      Begin MSComCtl2.DTPicker dtpRcvDt 
         Height          =   300
         Left            =   1785
         TabIndex        =   95
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   24051715
         CurrentDate     =   36467
      End
      Begin MSComCtl2.DTPicker dtpFRcvDt 
         Height          =   300
         Left            =   90
         TabIndex        =   96
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   24051715
         CurrentDate     =   36467
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
         Left            =   3540
         MouseIcon       =   "Lis154.frx":0F3A
         MousePointer    =   99  '����� ����
         TabIndex        =   5
         Top             =   225
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
         Top             =   195
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
      TabIndex        =   6
      Top             =   285
      Width           =   10035
      Begin VB.CommandButton cmdCaution 
         BackColor       =   &H008080FF&
         Caption         =   "Caution"
         Height          =   345
         Left            =   0
         MaskColor       =   &H8000000F&
         Style           =   1  '�׷���
         TabIndex        =   40
         Top             =   0
         Visible         =   0   'False
         Width           =   975
      End
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
         Left            =   1395
         MaxLength       =   10
         TabIndex        =   8
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
         TabIndex        =   7
         ToolTipText     =   "�˻� ����ũ�� �Է��ϼ���."
         Top             =   1230
         Width           =   7365
      End
      Begin MedControls1.LisLabel LisLabel1 
         Height          =   285
         Left            =   630
         TabIndex        =   9
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
      Begin MedControls1.LisLabel lblPtNm 
         Height          =   315
         Left            =   4815
         TabIndex        =   12
         Top             =   300
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
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
         Height          =   315
         Left            =   8160
         TabIndex        =   13
         Top             =   300
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   556
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
         Height          =   315
         Left            =   1410
         TabIndex        =   14
         Top             =   675
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
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
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblDeptNm 
         Height          =   315
         Left            =   4815
         TabIndex        =   15
         Top             =   660
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
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
         Height          =   315
         Left            =   8160
         TabIndex        =   11
         Top             =   660
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   556
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
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   5
         Left            =   225
         TabIndex        =   34
         Top             =   315
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "ȯ��   ID"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblReceptNo 
         Height          =   315
         Left            =   225
         TabIndex        =   35
         Top             =   675
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "ó �� ��"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   7
         Left            =   3645
         TabIndex        =   36
         Top             =   300
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "��      ��"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   8
         Left            =   3645
         TabIndex        =   37
         Top             =   660
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "�� �� ��"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   11
         Left            =   6990
         TabIndex        =   38
         Top             =   300
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "�� / ����"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   12
         Left            =   6990
         TabIndex        =   39
         Top             =   660
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "��      ��"
         Appearance      =   0
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
End
Attribute VB_Name = "frm154NurCol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890
Option Explicit

Private MySql           As clsLISSqlCollection
Private MyPatient       As clsPatient
Private objLISCollect   As clsLISCollectioin

Private mvarEmpId       As String
Private mvarWardId      As String
Private mvarDeptCd      As String
Private mvarHosilID     As String
Private mvarRoomID      As String

Private PtFg            As Boolean
Private MsgFg           As Boolean
Private SelAllFg        As Boolean
Private IsFirst         As Boolean
Private OrdFg           As Boolean
Private blnCleared      As Boolean


Private strBlgCd        As String       '������ �ǹ� �ڵ�
Private strErBldCd      As String       '�����ϰ�� �˻��� �ǹ��ڵ�
Private strGBldCd       As String       '�Ϲ��ϰ�� �˻��� �ǹ��ڵ�

Public Event LastFormUnload()
Private Const lngMaxRows = 19
Private Const lngRowHeight = 12

Private AdoCn_ORACLE    As ADODB.Connection
Private AdoRs_ORACLE    As ADODB.Recordset

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
    If chkCollect.Value = 0 Then lvwPtList.ListItems.Clear: Exit Sub
        
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

Private Sub cmdCaution_Click()
    Dim SQL As String
    Dim iCnt As Integer

    Set AdoCn_ORACLE = New ADODB.Connection

    With AdoCn_ORACLE
        .ConnectionTimeout = 25
'        .Provider = "OraOLEDB.Oracle.1"
        .Provider = "MSDAORA.1"                 ' Oracle "MSDAORA.1"
        .Properties("Data Source").Value = "PMC"
'        .Properties("Initial Catalog").Value = DatabaseName
        .Properties("Persist Security Info") = True
        
        .Properties("User ID").Value = "oral1"
        .Properties("Password").Value = "oral1"
        
'        Screen.MousePointer = vbHourglass
        .Open
    End With
    
    Set AdoRs_ORACLE = New ADODB.Recordset
    
    SQL = ""
    SQL = SQL + "SELECT AFBYN,                                     "
    SQL = SQL + "       TBYN,                                      "
    SQL = SQL + "       SUDUYN,                                    "
    SQL = SQL + "       HONGYN,                                    "
    SQL = SQL + "       HIVYN,                                     "
    SQL = SQL + "       VDRLYN,                                    "
    SQL = SQL + "       HBSAGYN,                                   "
    SQL = SQL + "       HCVYN,                                     "
    SQL = SQL + "       HBCYN,                                     "
    SQL = SQL + "       HAVYN,                                     "
    SQL = SQL + "       MRSAYN,                                    "
    SQL = SQL + "       VREYN,                                     "
    SQL = SQL + "       CDIFFIYN,                                  "
    SQL = SQL + "       FUNGUSYN,                                  "
    SQL = SQL + "       ROTAYN,                                    "
    SQL = SQL + "       OHMYN,                                     "
    SQL = SQL + "       EEEYN,                                     "
    SQL = SQL + "       JANGTIYN,                                  "
    SQL = SQL + "       EEEJILYN,                                  "
    SQL = SQL + "       NEWFLUYN,                                  "
    SQL = SQL + "       PUNGYN,                                    "
    SQL = SQL + "       PENICILN,                                  "
    SQL = SQL + "       INFLUYN,                                    "
    SQL = SQL + "       NEWINFECYN,                                 "
    SQL = SQL + "       BETCYN,                                     "
    SQL = SQL + "       CREYN,                                      "
    SQL = SQL + "       VRSAYN,                                     "
    SQL = SQL + "       CJDYN,                                      "
    SQL = SQL + "       CETCYN,                                     "
    SQL = SQL + "       PERYN,                                      "
    SQL = SQL + "       MENYN,                                      "
    SQL = SQL + "       DETCYN,                                     "
    SQL = SQL + "       MUMPSYN,                                    "
    SQL = SQL + "       RADCONT,                                   "
    SQL = SQL + "       DRUGALGY,                                  "
    SQL = SQL + "       OTHERRMK,                                  "
    SQL = SQL + "       PATNO,                                     "
    SQL = SQL + "       SEQ,                                       "
    SQL = SQL + "       TO_CHAR(EDITDATE,'YYYYMMDD') AS EDITDATE,                      "
    SQL = SQL + "       EDITID,                                                        "
    SQL = SQL + "       FN_USERNAME_SELECT(EDITID) AS EDITNM                          "
    SQL = SQL + "  FROM MDCAUTNT                                                       "
    SQL = SQL + " WHERE PATNO = '" & Trim(txtPtId.Text) & "'                                             "
    SQL = SQL + "   AND SEQ = (SELECT MAX(SEQ) FROM MDCAUTNT WHERE PATNO = '" & Trim(txtPtId.Text) & "') "

    AdoRs_ORACLE.CursorLocation = adUseClient
    AdoRs_ORACLE.Open SQL, AdoCn_ORACLE
    
    With AdoRs_ORACLE
        If .RecordCount > 0 Then
            For iCnt = 0 To 33
                If .Fields(iCnt).Value = "Y" Then
                    Check1(iCnt).Value = 1
                Else
                    Check1(iCnt).Value = 0
                End If
            Next
            
'            '2014-01-26 ���÷翣�� ����
'            If .Fields("INFLUYN").Value = "Y" Then
'                Check1(23).Value = 1
'            Else
'                Check1(23).Value = 0
'            End If
'            '2014-01-26 ���÷翣�� ����
            
            lblWDt.Caption = Format(.Fields("EDITDATE").Value & "", "####-##-##")
            lblWNm.Caption = .Fields("EDITNM").Value & ""
            txtDrug.Text = .Fields("DRUGALGY").Value & ""
            RichText.Text = .Fields("OTHERRMK").Value & ""
            
            Frame4.Visible = True
            If Check1(4).Value = 1 Then
                Picture2.Visible = True
            Else
                Picture2.Visible = False
            End If
        Else
            Frame4.Visible = False
        End If
        .Close
    End With
    Set AdoCn_ORACLE = Nothing
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
    Set MySql = Nothing
    Set MyPatient = Nothing
    Set objLISCollect = Nothing
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

Private Sub cmdFind_Click()
    Dim objPtInfo   As clsPatient
    Dim Rs          As Recordset
    Dim itmX        As ListItem
    Dim lngSearch   As Long
    Dim tmpFrDt     As String
    Dim tmpToDt     As String
   
    Set objPtInfo = New clsPatient
    Set Rs = New Recordset
    
    tmpFrDt = Format(dtpFRcvDt.Value, "yyyymmdd")
    tmpToDt = Format(dtpRcvDt.Value, "yyyymmdd")
        
    Rs.Open objPtInfo.GetSQLCol_Wm(tmpFrDt, tmpToDt, lblWardId.Caption), DBConn
        
    lvwPtList.ListItems.Clear
    If Rs.EOF = False Then
        With lvwPtList
            Do Until Rs.EOF
                Set itmX = .ListItems.Add()
                itmX.Text = Rs.Fields("ptid").Value & ""
                itmX.SubItems(1) = Rs.Fields("ptnm").Value & ""
                itmX.SubItems(2) = Rs.Fields("SSN").Value & ""
                itmX.SubItems(3) = Format(Rs.Fields("DOB").Value & "", CS_DateLongMask)
                itmX.SubItems(4) = IIf((Mid(Rs.Fields("ssn").Value & "", 7, 1) Mod 2) = 1, "��", "��")
                If IsDate(itmX.SubItems(3)) Then
                    itmX.SubItems(4) = itmX.SubItems(4) & " / " & DateDiff("yyyy", itmX.SubItems(3), GetSystemDate)
                End If
                If .ListItems.Count >= 1000 Then Exit Do
                Rs.MoveNext
            Loop
        End With
    Else
        MsgBox "���ǿ� �´� �ڷᰡ �����ϴ�. Ȯ���� �˻��ϼ���", vbInformation + vbOKOnly, Me.Caption
    End If
    Set Rs = Nothing
    
    Set objPtInfo = Nothing
End Sub

'& ä�� Ŭ���� MyCollect �� �̿��Ͽ� �ش� ȯ�ڵ��� ó���� ä�������Ѵ�.
Private Sub cmdSave_Click()
    Dim objPrgBar       As jProgressBar.clsProgress
    Dim APSColSuccess   As Boolean
    Dim BBSColSuccess   As Boolean
    Dim LISColSuccess   As Boolean

    Dim ii              As Integer
    Dim iCheckOrder     As Integer

    If CollectionTargetChk = False Then
       MsgBox "ä���� �׸��� �����ϼ���..", vbInformation, "�׸���"
       tblOrdSheet.SetFocus
       Exit Sub
    End If

    cmdSave.Enabled = False

    iCheckOrder = objLISCollect.CheckSameOrder(tblOrdSheet, 1)     '�ߺ�ó�� Check
    If iCheckOrder > 0 Then GoTo OrdCheck1

    Call MouseRunning


    Set objPrgBar = New jProgressBar.clsProgress
    With objPrgBar
        .Container = Me
        .Left = lblBar.Left + 5
        .Top = lblBar.Top + 5
        .Width = lblBar.Width - 10
        .Height = lblBar.Height - 10
        .Message = "���õ� �˻��׸� ���� ä��ó�����Դϴ�..."
        .Max = 90
        .Value = 10
        DoEvents
    End With

    DoEvents

        '----------------------------------------------------------
    '������ ������ ���ؼ� �������� �ҷ��� �����Ѵ�.(2001/06/08)
    '----------------------------------------------------------
    Call tblordersheet

    Dim objDIC As New clsDictionary
    objDIC.Clear
    objDIC.FieldInialize "orddiv", "first,last,coldt,coltm"
    With tblOrdSheet
        For ii = 1 To .DataRowCnt
            .Row = ii: .Col = enCOLLIST.tcORDDIV
            Select Case .Value
                Case BBS_ORDDIV
                    If objDIC.Exists(.Value) Then
                        objDIC.KeyChange BBS_ORDDIV
                        objDIC.Fields("last") = .Row
                    Else
                        .Col = enCOLLIST.tcREQDTTM
                        objDIC.AddNew BBS_ORDDIV, .Row & COL_DIV & "" & COL_DIV & _
                                      Format(.Text, "yyyymmdd") & COL_DIV & Format(.Text, "HHmm")
                    End If
                Case LIS_ORDDIV
                    If objDIC.Exists(.Value) Then
                        objDIC.KeyChange LIS_ORDDIV
                        objDIC.Fields("last") = .Row
                    Else
                        objDIC.AddNew LIS_ORDDIV, .Row & COL_DIV & "" & COL_DIV & "" & COL_DIV & ""
                    End If
            End Select
        Next
        objDIC.MoveFirst
        Do Until objDIC.EOF
            If objDIC.Fields("last") = "" Then
                objDIC.Fields("last") = objDIC.Fields("first")
            End If
            objDIC.MoveNext
        Loop
    End With
    With objDIC
        .MoveFirst
        Do Until .EOF
            Select Case .Fields("orddiv")
                Case LIS_ORDDIV: iCheckOrder = objLISCollect.ChkSpcnm(tblOrdSheet, .Fields("first"), .Fields("last"))
            End Select
            If iCheckOrder > 0 Then GoTo OrdCheck2
            .MoveNext
        Loop
    End With

    '-------------------------------------------------------------
    '�������� ä���� �����Ѵ�(���������� ������ü üũ�� �ʿ����)
    '-------------------------------------------------------------
    With objDIC
        .MoveFirst
        BBSColSuccess = True: APSColSuccess = True: LISColSuccess = True
        Do Until .EOF
            Select Case .Fields("orddiv")
                Case BBS_ORDDIV:
                                 If chkChangeColTm.Value = 1 Then
                                        BBSColSuccess = CollectForBBS(.Fields("first"), .Fields("last"), _
                                                                           Format(dtpColDtTm.Value, "yyyymmdd"), _
                                                                           Format(dtpColDtTm.Value, "HHmmss"), objPrgBar)
                                 Else
                                        BBSColSuccess = CollectForBBS(.Fields("first"), .Fields("last"), _
                                                                           Format(GetSystemDate, "yyyymmdd"), _
                                                                           Format(GetSystemDate, "HHmmss"), objPrgBar)
                                 End If
                Case LIS_ORDDIV: LISColSuccess = CollectForLIS_New(.Fields("first"), .Fields("last"), objPrgBar, 1)
            End Select
            .MoveNext
        Loop
    End With

    '����Ÿ���̽��� ��¥/�ð����� System Date/Time�� ����...
    Date = Format(GetSystemDate, CS_DateLongFormat)
    Time = Format(GetSystemDate, CS_TimeLongFormat)

    If Not BBSColSuccess And LISColSuccess Then
        Set objPrgBar = Nothing
        MsgBox "ä��ó���� ������ �߻��߽��ϴ� !!" & vbCrLf & _
               "������Ͻ� �� ������ ��ӵǸ� ����� Ȥ�� �ӻ󺴸����� �����ٶ��ϴ�.", _
               vbCritical, "����"
    End If

    MouseDefault
    Set objPrgBar = Nothing
    Set objDIC = Nothing
ExitPos:
    Call cmdClear_Click
    cmdSave.Enabled = True
    txtPtId.SetFocus
    Set objDIC = Nothing
    Exit Sub

OrdCheck1:
    tblOrdSheet.Row = iCheckOrder
    tblOrdSheet.Col = 1
    tblOrdSheet.Action = ActionActiveCell
    cmdSave.Enabled = True
    tblOrdSheet.SetFocus
    Exit Sub

OrdCheck2:
    tblOrdSheet.Row = iCheckOrder
    tblOrdSheet.Col = 1
    tblOrdSheet.Action = ActionActiveCell
    MsgBox "������ü ������ �����ϴ�. ����� Ȥ�� �ӻ󺴸����� �����ϼ���.", vbInformation + vbOKOnly, "����"
    cmdSave.Enabled = True
    tblOrdSheet.SetFocus
    Set objDIC = Nothing
    Exit Sub
End Sub

'** �������� ä����ƾ
Private Function CollectForBBS(ByVal FRowCnt As Integer, ByVal LRowCnt As Integer, _
                               ByVal ColDt As String, ByVal ColTm As String, _
                               ByRef objProgress As Object) As Boolean

    
    Dim dicBBS      As clsDictionary
    Dim objBar      As clsDictionary
    Dim objCollect  As clsBBSCollection
    Dim tmpClipData As String
    Dim tmpTotData  As Variant
    Dim tmpRowData  As Variant
    
    Dim strColDt    As String      'ä����
    Dim strColTm    As String      'ä���Ͻ�
    Dim strStatFg   As String
    
    Dim i           As Long
    Dim lngColCnt   As Integer
    
    Set objCollect = New clsBBSCollection
    
    If objCollect.Blood_Existence(txtPtId.Text, Format(GetSystemDate, "yyyymmdd"), Format(GetSystemDate, "hhmmss")) = False Then
        If objCollect.SetAccessCheck(txtPtId.Text) = True Then
           '��ü�� �̹� �����ϴ� ���
           CollectForBBS = objCollect.SetWardAccess(txtPtId.Text, enBussDiv.BussDiv_InPatient, Format(GetSystemDate, "yyyymmdd"), _
                                    Format(GetSystemDate, "hhmmss"), ObjSysInfo.EmpId)
                
            Set objCollect = Nothing
            Exit Function
        End If
    End If
    
    Set dicBBS = New clsDictionary
    Set objBar = New clsDictionary
'    Set objCollect = New clsBBSCollection
    
    lngColCnt = 0
    HosilId = medGetP(lblLocation.Caption, 2, "-")
    
    With tblOrdSheet
        .Row = FRowCnt: .Col = enCOLLIST.tcWARDID: mvarWardId = .Value
                        .Col = enCOLLIST.tcROOMID: mvarHosilID = .Value
                        .Col = enCOLLIST.tcDEPTCD: mvarDeptCd = .Value
        .Col = 1: .Col2 = .MaxCols
        .Row = FRowCnt: .Row2 = LRowCnt
        .BlockMode = True
        tmpClipData = .ClipValue: tmpTotData = Split(tmpClipData, vbCrLf)
        .BlockMode = False
        strColDt = ColDt: strColTm = ColTm

        .Col = 7: strStatFg = IIf(Trim(.Value) = "Y", "1", "0")
        
        For i = 0 To UBound(tmpTotData) - 1

            tmpRowData = Split(tmpTotData(i), vbTab)
            If objProgress.Max > objProgress.Value Then objProgress.Value = objProgress.Value + 1
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
        CollectForBBS = True
        Exit Function
    End If
          
    objCollect.WardId = mvarWardId
    CollectForBBS = objCollect.Set_Collect(dicBBS, , objProgress)
    
    If CollectForBBS Then
        Set objBar = objCollect.BldDic
        If objBar.RecordCount > 0 Then
        '���ڵ� ���
            BarCodePrintForBBS objBar
        Else
            Set objProgress = Nothing
            
            If objCollect.CheckCol Then
                MsgBox "���������� ó������ �ʾҽ��ϴ�.", vbExclamation
            Else
                MsgBox "����ó�� ��ü�� �̹� �����ϹǷ� ���ڵ尡 ��µ��� �ʽ��ϴ�.", vbInformation + vbOKOnly, "���ڵ����"
            End If
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
                               ByRef objProgress As Object, _
                               Optional pINx As Integer = 0) As Boolean
    Dim tmpData()   As String
    Dim i           As Integer
    Dim SelCount    As Integer
    Dim CollectCnt  As Integer
    
    Dim ColSuccess  As Boolean
    
    CollectCnt = 0
    Call objLISCollect.InitRtn

    With tblOrdSheet

        ReDim tmpData(0 To 20)
        For i = FRowCnt To LRowCnt
            
            If objProgress.Max > objProgress.Value Then objProgress.Value = objProgress.Value + 1
            
            .Row = i
            
            .Col = enCOLLIST.tcCHECK
            If .Value <> 1 Then GoTo Skip

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
' 2009.01.14 �缺�� �̷�ó���� ó���� ���� tcORDDATE�� BEDINDT�� �����.
'            .Col = enCOLLIST.tcORDDT:  tmpData(9) = .Value        'OrdDt
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
        objLISCollect = True
        Exit Function
    End If

    With objLISCollect

        ReDim tmpData(0 To 16)

        tmpData(0) = Mid(Format(GetSystemDate, "YYYY"), 4)  '��ü�⵵
        tmpData(1) = MyPatient.Ptid                         'ȯ��ID
        
'' 2008.10.24 ������ �������ϰ�� ���ڵ忡 ���� ���̴� ��� �߰�.
'
'        If Len(lblDiseaseSang.Caption) > 5 Then
'            tmpData(2) = "*" & Trim(objPatient.PtNm)
'        Else
'            tmpData(2) = objPatient.PtNm
'        End If

' 2011.07.06 ������ �������ϰ�� ���ڵ忡 ���� ���̴� ��� ����

'        If Len(lblDiseaseSang_New.Caption) > 0 Then
'            tmpData(2) = "*" & Trim(objPatient.PtNm)
'        Else
            tmpData(2) = MyPatient.PtNm
'        End If
        
        
        tmpData(3) = MyPatient.Sex                             '����
        If IsDate(Format(MyPatient.Dob, CS_DateLongMask)) Then                          'ȯ���Ϸ�
            tmpData(4) = DateDiff("y", Format(MyPatient.Dob, CS_DateLongMask), GetSystemDate)
        Else
            tmpData(4) = Mid(MyPatient.Dob, 1, 4) & "-01-01"
            If IsDate(tmpData(4)) Then
                tmpData(4) = DateDiff("y", tmpData(4), GetSystemDate)
            Else
                tmpData(4) = 0
            End If
        End If
        tmpData(5) = ""                                          '�Կ���
        tmpData(6) = Format(GetSystemDate, CS_DateDbFormat)  '�Է���
        tmpData(7) = Format(GetSystemDate, CS_TimeDbFormat)  '�Է½ð�
        tmpData(8) = ObjSysInfo.EmpId                            '�Է���
        tmpData(9) = ""                                          '��������ȣ
        tmpData(10) = Format(GetSystemDate, CS_DateDbFormat) 'ä����
        .ColTm = Format(GetSystemDate, "HHMMSS")             'ä����
        tmpData(11) = ObjSysInfo.EmpId                           'ä����
        tmpData(12) = ""                                         '����ID
        tmpData(13) = ""                                         '����ID
        tmpData(14) = ""                                         'ħ��ID
        tmpData(15) = ""                                         'ħ��ID
        tmpData(16) = ObjSysInfo.BuildingCd                      '** ä���� ����Ǵ� �ǹ��ڵ�
        
        

        Call .SetColData(tmpData)
    End With
    ' ä�� ����
    ColSuccess = objLISCollect.DoCollection(objProgress)
    
    '** ��������:���PCx(��������) ���Կ� ���� ä��-���� ��ƾ �ʿ� (�ܷ�ä���ǿ��� �ٷ� ������ �ϱ� ����)
    '   �߰� By M.G.Choi 2007.04.02
    '---------------------------------------------------------------------------------------------------------------------
    If ColSuccess = True And pINx = 1 Then
        objProgress.Message = "���� Procedure�� �����ϰ� �ֽ��ϴ�."
        Dim objAccess   As New clsLISAccession
        
        With objLISCollect
            If .CollectDone Then
                Dim pWorkArea As String
                Dim pAccDt As String
                Dim pAccSeq As Long
                
                For i = 1 To .ColCount
                    objProgress.Message = "���� Procedure�� �����ϰ� �ֽ��ϴ�. (" & CStr(i) & "/" & CStr(.ColCount) & ")"
                    Call .GetLabNumbers(i, pWorkArea, pAccDt, pAccSeq)
                    ColSuccess = objAccess.DoAccession_New(pWorkArea, pAccDt, pAccSeq, ObjMyUser.EmpId)
                    If Not ColSuccess Then Exit For
                    If objProgress.Value = objProgress.Max Then objProgress.Max = objProgress.Max + 1
                    objProgress.Value = objProgress.Value + 1
                    DoEvents
                Next
            End If
        End With
        Set objAccess = Nothing
    End If
    '----------------------------------------------------------------------------------------------------------------------
    
    If Not ColSuccess Then
        Set objProgress = Nothing
        MsgBox "ä��ó���� ������ �߻��߽��ϴ� !!"
        MouseDefault  '0
        CollectForLIS_New = False
        Exit Function
    End If
    CollectForLIS_New = True

End Function

Private Function CollectForLIS(ByVal FRowCnt As Long, ByVal LRowCnt As Long, ByRef objProgress As Object) As Boolean
    Dim tmpData()   As String
    Dim ColSuccess  As Boolean
    
    Dim strTmp1     As String
    Dim strReqDt    As String
    Dim strReqtm    As String
    Dim strReqTm1   As String
    Dim strLastTm   As String
    Dim CollectCnt  As Integer
    Dim i           As Integer
    
    strLastTm = ""

    '����Ÿ���̽��� ��¥/�ð����� System Date/Time�� ����...
    Date = GetSystemDate
    Time = GetSystemDate

    CollectCnt = 0
    Call objLISCollect.InitRtn

    With tblOrdSheet

        ReDim tmpData(0 To 20)
        .Row = FRowCnt: .Col = enCOLLIST.tcWARDID: mvarWardId = .Value
                        .Col = enCOLLIST.tcROOMID: mvarHosilID = .Value
                        .Col = enCOLLIST.tcDEPTCD: mvarDeptCd = .Value
        For i = FRowCnt To LRowCnt
            
            If objProgress.Max > objProgress.Value Then objProgress.Value = objProgress.Value + 1
            
            .Row = i
            
            .Col = enCOLLIST.tcCHECK
            If .Value <> 1 Then GoTo Skip

            CollectCnt = CollectCnt + 1
            .Col = 36: strTmp1 = .Value
            .Col = enCOLLIST.tcBUILDCD:  tmpData(0) = .Value        'Delivery Location
            .Col = enCOLLIST.tcWORKAREA: tmpData(1) = .Value        'WorkArea
            .Col = enCOLLIST.tcSPCCD:    tmpData(2) = .Value        'SpcCd
            .Col = enCOLLIST.tcSTORECD:  tmpData(3) = .Value        'StoreCd
            .Col = enCOLLIST.tcSTATFLAG: tmpData(4) = .Value        'StatFg
            .Col = enCOLLIST.tcREQDTTM:
                
                If strTmp1 = "1" Then
                    strReqDt = medGetP(.Value, 1, " ")
                    If strLastTm = "" Then
                        strReqtm = Val(Mid(medGetP(.Value, 2, " "), 1, 2)) + 1
                        strLastTm = strReqtm
                    Else
                        strReqtm = Val(strLastTm) + 1
                    End If
                    strReqTm1 = Mid(medGetP(.Value, 2, " "), 3)
                    strReqtm = strReqtm & strReqTm1
                    strReqDt = strReqDt & " " & strReqtm
                    tmpData(5) = strReqDt        'ReqColDate
                Else
                    tmpData(5) = .Value        'ReqColDate
                End If
                
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
        CollectForLIS = True
        Exit Function
    End If

    With objLISCollect

        ReDim tmpData(0 To 16)

        tmpData(0) = Mid(Format(GetSystemDate, "YYYY"), 4)  '��ü�⵵
        tmpData(1) = MyPatient.Ptid                            'ȯ��ID
        tmpData(2) = MyPatient.PtNm
        tmpData(3) = MyPatient.Sex                             '����
        If IsDate(Format(MyPatient.Dob, CS_DateLongMask)) Then                         'ȯ���Ϸ�
            tmpData(4) = DateDiff("y", Format(MyPatient.Dob, CS_DateLongMask), GetSystemDate)
        Else
            tmpData(4) = Mid(MyPatient.Dob, 1, 4) & "-01-01"
            If IsDate(tmpData(4)) Then
                tmpData(4) = DateDiff("y", tmpData(4), GetSystemDate)
            Else
                tmpData(4) = 0
            End If
        End If
        tmpData(5) = MyPatient.BedInDt                           '�Կ���
        tmpData(6) = Format(GetSystemDate, CS_DateDbFormat)  '�Է���
        tmpData(7) = Format(GetSystemDate, CS_TimeDbFormat)  '�Է½ð�
        tmpData(8) = gEmpId                                      '�Է���
        tmpData(9) = ""                                          '��������ȣ
        tmpData(10) = Format(GetSystemDate, CS_DateDbFormat) 'ä����
        
        tmpData(11) = gEmpId                                     'ä����
        tmpData(12) = mvarWardId                                 '����ID
        tmpData(13) = mvarHosilID                                '����ID
        tmpData(14) = ""                                         'ħ��ID
        tmpData(15) = ""                                         'ħ��ID
        tmpData(16) = ObjSysInfo.BuildingCd      '** ä���� ����Ǵ� �ǹ��ڵ�
        

        
        Call .SetColData(tmpData)
        
        If chkChangeColTm.Value = 1 Then
            .ColDt = Format(dtpColDtTm.Value, CS_DateDbFormat)
            .ColTm = Format(dtpColDtTm.Value, "HHMMSS")
        Else
            .ColDt = Format(GetSystemDate, CS_DateDbFormat)
            .ColTm = Format(GetSystemDate, "HHMMSS")
        End If
        
        
    End With

    ' ä�� ����
    objLISCollect.SetTrans = True
    ColSuccess = objLISCollect.DoCollection(objProgress)
    If Not ColSuccess Then
        Set objProgress = Nothing
        MsgBox "ä��ó���� ������ �߻��߽��ϴ� !!"
        MouseDefault  '0
        CollectForLIS = False
        Exit Function
    End If
    CollectForLIS = True

    
End Function

Private Sub BarCodePrintForBBS(objDIC As clsDictionary)
    Dim objBar      As clsBarcode
    Dim objSQL      As clsBBSCollection
    
    Dim strPtid     As String
    Dim strPtnm     As String
    Dim strColDt    As String
    Dim strColTm    As String
    Dim strSpcNo    As String
    Dim strW_Dept   As String
    Dim strBuildNm  As String        '�ǹ��̸�
    Dim strAccSeq   As String         'SpcYy-SpcNo ������ ��ü��ȣ
    Dim strHosilid  As String
    Dim strStatFg   As String
    
    Set objBar = New clsBarcode
    Set objSQL = New clsBBSCollection
    
'    Set objBAR.MyDB = dbconn
    Set objBar.TableInfo = New clsTables
    
    strW_Dept = mvarWardId
    If strW_Dept = "" Then
        strW_Dept = mvarDeptCd
    End If
    
    If lblLocation.Caption <> "" Then
        If lblLocation.Caption <> "--" Then strW_Dept = strW_Dept & "/" & mvarHosilID
    End If
    
    If P_ApplyBuildingInfo Then
        strBuildNm = ObjSysInfo.BuildingNm
    Else
        strBuildNm = "����"
    End If
    
    objDIC.MoveFirst
    Do Until objDIC.EOF
        strPtid = medGetP(objDIC.GetString, 1, COL_DIV)
        strPtnm = medGetP(objDIC.GetString, 2, COL_DIV)
        strSpcNo = medGetP(objDIC.GetString, 3, COL_DIV)
        strColDt = Mid(medGetP(objDIC.GetString, 4, COL_DIV), 5)
        strColTm = Mid(medGetP(objDIC.GetString, 5, COL_DIV), 1, 4)
        strStatFg = medGetP(objDIC.GetString, 7, COL_DIV)
        strColDt = Format(strColDt, "00/00")
        strColTm = Format(strColTm, "0#:##")
        
        '��ü��ȣ ��� : 2001.2.8 �߰�
        strAccSeq = Mid(strSpcNo, 1, 2) & "-" & Format(Mid(strSpcNo, 3), "########0")
        strAccSeq = Format(strAccSeq, String(11, "@"))
        '
        objBar.Label_PrintOut strBuildNm, "XM", "", strAccSeq, strSpcNo, strPtid, _
                                            strPtnm, "", "", strStatFg, strW_Dept, strColDt, strColTm, _
                                            "", 1
        objDIC.MoveNext
    Loop
    Set objBar = Nothing
    Set objSQL = Nothing

End Sub

Private Sub cmdWardHelp_Click()

    Dim objDeptHelp As clsPopUpList
'    Dim objWard As clsBasisData
    
    Set objDeptHelp = New clsPopUpList
'    Set objWard = New clsBasisData
    
    lvwPtList.ListItems.Clear
    
    With objDeptHelp
        .Connection = DBConn
        .FormCaption = "��������Ʈ"
        .ColumnHeaderText = "����;������"
        .LoadPopUp GetSQLWard ', 2000, 1500 ', ObjLISComCode.WardId
        
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
'    Set objWard = Nothing

End Sub

Private Sub Command1_Click()
    lblWDt.Caption = ""
    lblWNm.Caption = ""
    txtDrug.Text = ""
    RichText.Text = ""
    Frame4.Visible = False
End Sub

Private Sub Command2_Click()
    Picture2.Visible = False
End Sub

Private Sub dtpColDtTm_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then tblOrdSheet.SetFocus
End Sub

Private Sub dtpColDtTm_LostFocus()

    Dim Resp As VbMsgBoxResult
    If Format(dtpColDtTm.Value, "YYYYMMDD HH:MM") < Format(Now, "YYYYMMDD HH:MM") Then
        Resp = MsgBox("ä��ð��� ����ð����� �����Դϴ�. �����Ͻðڽ��ϱ�?", _
               vbQuestion + vbYesNo, "ä��ð�����")
        If Resp = vbNo Then
            dtpColDtTm.Value = Format(GetSystemDate, "YY-MM-DD HH:MM")
        End If
        chkChangeColTm.Value = 0
    End If
    
End Sub

Private Sub Form_Activate()
    If Not IsFirst Then Exit Sub
    IsFirst = False
    
    If P_IncludeBBSSystem Then
        picOrdDiv.Visible = True
    Else
        picOrdDiv.Visible = False
    End If
    
    medInitLvwHead lvwPtList, "ȯ��ID,ȯ�ڼ���,�ֹε�Ϲ�ȣ,�������,����/����", _
                       "50,50,800,300,100"
    txtSearchKey.Text = ""
    Call ClearRtn
    If Trim(gWardId) <> "" Then
        lblWardId.Caption = Trim(gWardId)
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
    Set MySql = New clsLISSqlCollection
    Set MyPatient = New clsPatient
    Set objLISCollect = New clsLISCollectioin
    Frame4.Visible = False
    
    dtpFRcvDt.Value = GetSystemDate
    dtpRcvDt.Value = GetSystemDate
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call ICSPatientMark
    
    Set MySql = Nothing
    Set MyPatient = Nothing
    Set objLISCollect = Nothing
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
        txtPtId.Text = .Text                'ȯ��ID
        Call txtPtId_KeyPress(vbKeyReturn)
    End With
    
End Sub

Private Sub tblOrdSheet_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

    
    Dim ButtonValue As Variant
    Dim SvOrdDt     As String
    Dim SvOrdNo     As String
    Dim i           As Integer
    
    If SelAllFg Then Exit Sub
    
    With tblOrdSheet
       .Row = Row
       .Col = Col:   ButtonValue = .Value
       
       If .Value = 0 Then Exit Sub
       
       .Col = 9:      SvOrdDt = .Value
       .Col = 10:    SvOrdNo = .Value
       
       For i = 1 To .MaxRows
          If i <> Row Then
             .Row = i
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
    
On Error GoTo Errors

    If Screen.ActiveForm.ActiveControl.Name = cmdClear.Name Then Exit Sub
    If Screen.ActiveForm.ActiveControl.Name = cmdExit.Name Then Exit Sub
    If Screen.ActiveForm.ActiveControl.Name = txtSearchKey.Name Then Exit Sub
    If Screen.ActiveForm.ActiveControl.Name = lvwPtList.Name Then Exit Sub
    If Screen.ActiveForm.ActiveControl.Name = optSort(0).Name Then Exit Sub
    
    If blnCleared Then Call txtPtId_KeyPress(vbKeyReturn)
    Exit Sub

Errors:
    Resume Next
    
End Sub

Private Sub txtSearchKey_GotFocus()

    With txtSearchKey
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub

'% ȯ��ID �Ǵ� �������� �˻� ����Ʈ �ۼ�.
Private Sub txtSearchKey_KeyPress(KeyAscii As Integer)
    
    Dim objPtInfo   As clsPatient
    Dim Rs          As Recordset
    Dim itmX        As ListItem
    Dim lngSearch   As Long
    
    If txtSearchKey.Text = "" Then Exit Sub
    
    Set objPtInfo = New clsPatient ' clsHosComSQLStmt
    Set Rs = New Recordset
    
    If KeyAscii = vbKeyReturn Then
'        lngSearch = IIf(optSort(0).Value, 1, 2) + 4 'True:ȯ��ID, False:ȯ�ڸ�
        
        If chkCollect.Value = 0 And optSort(0).Value Then
            lngSearch = "5"
        ElseIf chkCollect.Value = 0 And optSort(0).Value = False Then
            lngSearch = "6"
        ElseIf chkCollect.Value = 1 And optSort(0).Value Then
            lngSearch = "1"
        ElseIf chkCollect.Value = 1 And optSort(0).Value = False Then
            lngSearch = "2"
        End If
        
        If chkCollect.Value = 0 Then
            Rs.Open objPtInfo.GetSQLPtNt(lngSearch, txtSearchKey.Text), DBConn
        Else
            Rs.Open objPtInfo.GetSQLCol(lngSearch, txtSearchKey.Text, lblWardId.Caption), DBConn
        End If
        
        lvwPtList.ListItems.Clear
        If Rs.EOF = False Then
            With lvwPtList
                Do Until Rs.EOF
                    Set itmX = .ListItems.Add()
                    itmX.Text = Rs.Fields("ptid").Value & ""
                    itmX.SubItems(1) = Rs.Fields("ptnm").Value & ""
                    itmX.SubItems(2) = Rs.Fields("SSN").Value & ""
                    itmX.SubItems(3) = Format(Rs.Fields("DOB").Value & "", CS_DateLongMask)
                    itmX.SubItems(4) = IIf((Mid(Rs.Fields("ssn").Value & "", 7, 1) Mod 2) = 1, "��", "��")
                    If IsDate(itmX.SubItems(3)) Then
                        itmX.SubItems(4) = itmX.SubItems(4) & " / " & DateDiff("yyyy", itmX.SubItems(3), GetSystemDate)
                    End If
                    If .ListItems.Count >= 1000 Then Exit Do
                    Rs.MoveNext
                Loop
            End With
        Else
            MsgBox "���ǿ� �´� �ڷᰡ �����ϴ�. Ȯ���� �˻��ϼ���", vbInformation + vbOKOnly, Me.Caption
        End If
        Set Rs = Nothing
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
    
    If Not blnCleared Then
       Call ClearRtn
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
   
On Error GoTo Errors

    If KeyAscii = vbKeyReturn Then
        If IsNumeric(txtPtId.Text) Then txtPtId.Text = Format(txtPtId.Text, P_PatientIdFormat)
        Call ICSPatientMark(txtPtId.Text, enICSNum.LIS_ALL)
        If Not blnCleared Then Call ClearRtn
        DoEvents
        
        Set MyPatient = Nothing
        Set MyPatient = New clsPatient
        
'        Call MyPatient.ClearData   'Ŭ���� �� ���� �ʱ�ȭ
        If MyPatient.GETPatient(txtPtId.Text) Then
            lblPtNm.Caption = MyPatient.PtNm     '����
            lblSexAge.Caption = MyPatient.SEXNM & " / " & MyPatient.Age & " " & MyPatient.AGEDIV      '����
            lblDeptNm.Caption = MyPatient.DeptNm '�����
            lblLocation.Caption = MyPatient.WardId & "-" & MyPatient.RoomId & "-" & MyPatient.BedID   '����
            DoEvents
            PtFg = True
            Call MouseRunning
            Call DisplayOrder
            Call MouseDefault
'            Call cmdCaution_Click

            cmdSave.Enabled = True
        Else
            txtPtId.Text = ""
            MsgBox "��ϵ��� ���� ȯ��ID�Դϴ�.. �ٽ� �Է��ϼ���.."
            MsgFg = False: PtFg = False
            If txtPtId.Enabled Then txtPtId.SetFocus
            Call txtPtId_GotFocus
            Exit Sub
        End If
        
        If OrdFg Then
            tblOrdSheet.SetFocus
        Else
            Call cmdClear_Click
            cmdSave.Enabled = False
            txtPtId.SetFocus
            Call txtPtId_GotFocus
        End If
        Exit Sub
        
        chkSelAll.SetFocus
    End If
Errors:
    Resume Next
End Sub

'% �˻��� ó���� ���̺� ���÷��� �Ѵ�.
Private Sub DisplayOrder()
    Dim objProInSts As jProgressBar.clsProgress
    Dim objGetSql   As clsBBSCollection
    Dim tmpRs       As Recordset
    Dim SqlStmt     As String
    Dim SvOrdDt     As String
    Dim SvOrdNo     As String
    Dim SvSpcNm     As String
    Dim SvOrdDoct   As String
    Dim tmpDate     As String
    Dim tmpTime     As String
    Dim tmpStatFg   As String
    Dim tmpTestFg   As String
    Dim strErChk    As String
    Dim strOrdDiv   As String
    
    Dim i As Integer
    
On Error GoTo NoData
    
    Call TestBuilding_Search     '�������� ��ü���� ���
   
    Set objGetSql = New clsBBSCollection
    Set objProInSts = New jProgressBar.clsProgress
    
    With objProInSts
        .Container = Me
        .Left = lblBar.Left + 5
        .Top = lblBar.Top + 5
        .Width = lblBar.Width - 10
        .Height = lblBar.Height - 10
        .Message = "�ش�ȯ���� ó�� ������ �˻� ���Դϴ�..."
'        .SetMyForm Me
'        .Choice = True
'        .XPos = lblBar.Left + 5             'optCondition(1).Left + optCondition(1).Width + 20
'        .YPos = lblBar.Top + 5              'optCondition(1).Top + optCondition(1).Height - 260
'        .XWidth = lblBar.Width - 10         'fraWSHeader.Width - (optCondition(1).Width * 2)
'        .ForeColor = &HFA8B10               'DCM_LightBlue   '&H864B24
'        .Appearance = aPlate
'        .BorderStyle = bsNone
'        .YHeight = lblBar.Height - 10       ' 260
'        .Msg = "�ش�ȯ���� ó�� ������ �˻� ���Դϴ�...."
'        .Max = 90
'        .Min = 0
'        .Value = 10
        DoEvents
    End With

    DoEvents
    txtMesg.Text = ""
    
    ' ó�泻�� �˻�
'    tmpDate = Format(DateAdd("d", -2, GetSystemDate), CS_DateDbFormat)
    tmpDate = Format(GetSystemDate, CS_DateDbFormat)
    tmpTime = "000000"
    
'������������ �׽�Ʈ������ ����ҷ��� ���ٰ� �� ��������
'##############################
    gUsingInWardMenu = True
'##############################
    
    If gUsingInWardMenu Then
        strOrdDiv = "W"
    Else
        strOrdDiv = Mid(ObjSysInfo.projectid, 1, 1)
    End If
 
    Call cmdCaution_Click
    
    SqlStmt = MySql.SqlReadWardOrder_New(txtPtId.Text, tmpDate, tmpTime, , enBussDiv.BussDiv_InPatient, , strOrdDiv)
    Set tmpRs = New Recordset
    tmpRs.Open SqlStmt, DBConn
    
    If tmpRs.EOF Then
        MsgBox MyPatient.PtNm & " ���� ó�泻���� �����ϴ�", vbInformation, "��ȣ�� ä��"
        If Not blnCleared Then Call ClearRtn
        GoTo NoData
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
       
        objProInSts.Max = tmpRs.RecordCount
        
        'Locking Cells
        .Row = -1
        .Col = 2: .Col2 = .MaxCols
        .BlockMode = True
        .Lock = True
        .Protect = True
        .BlockMode = False
             
        For i = 1 To tmpRs.RecordCount

            objProInSts.Value = i

            .Row = i

            If SvOrdDt <> Trim("" & tmpRs.Fields("OrdDt").Value) Then
                .Col = enCOLLIST.tcORDDT:   .Text = Format("" & tmpRs.Fields("OrdDt").Value, CS_DateShortMask)    'ó����
                .Col = enCOLLIST.tcORDNO: .Text = Trim("" & tmpRs.Fields("OrdNo").Value)      'ó���ȣ
                .Col = enCOLLIST.tcSPCNM: .Text = Trim("" & tmpRs.Fields("SpcNm").Value)      '��ü
                .Col = enCOLLIST.tcDOCTNM: .Text = GetDoctNm(tmpRs.Fields("orddoct").Value & "") 'Trim("" & tmpRs.Fields("DoctNm").Value)     'ó����
                SvOrdDt = Trim("" & tmpRs.Fields("OrdDt").Value)
                SvOrdNo = Trim("" & tmpRs.Fields("OrdNo").Value)    'ó���ȣ
                SvSpcNm = Trim("" & tmpRs.Fields("SpcNm").Value)    '��ü
                SvOrdDoct = GetDoctNm(tmpRs.Fields("orddoct").Value & "") 'Trim("" & tmpRs.Fields("DoctNm").Value) 'ó����
            End If
            If SvOrdNo <> Trim("" & tmpRs.Fields("OrdNo").Value) Then
                .Col = enCOLLIST.tcORDNO: .Text = Trim("" & tmpRs.Fields("OrdNo").Value)      'ó���ȣ
                .Col = enCOLLIST.tcSPCNM: .Text = Trim("" & tmpRs.Fields("SpcNm").Value)      '��ü
                .Col = enCOLLIST.tcDOCTNM: .Text = GetDoctNm(tmpRs.Fields("orddoct").Value & "") 'Trim("" & tmpRs.Fields("DoctNm").Value)    'ó����
                SvOrdNo = Trim("" & tmpRs.Fields("OrdNo").Value)    'ó���ȣ
                SvSpcNm = Trim("" & tmpRs.Fields("SpcNm").Value)    '��ü
                SvOrdDoct = GetDoctNm(tmpRs.Fields("orddoct").Value & "") 'Trim("" & tmpRs.Fields("DoctNm").Value) 'ó����
            End If
            If SvSpcNm <> Trim("" & tmpRs.Fields("SpcNm").Value) Then
                .Col = enCOLLIST.tcSPCNM: .Text = Trim("" & tmpRs.Fields("SpcNm").Value)      '��ü
                SvSpcNm = Trim("" & tmpRs.Fields("SpcNm").Value)
            End If
            If SvOrdDoct <> GetDoctNm(tmpRs.Fields("orddoct").Value & "") Then 'Trim("" & tmpRs.Fields("DoctNm").Value) Then
                .Col = enCOLLIST.tcDOCTNM: .Text = GetDoctNm(tmpRs.Fields("orddoct").Value & "") 'Trim("" & tmpRs.Fields("DoctNm").Value)    'ó����
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
'                Dim objBld As clsBasisData
                Dim strBld As String
                
'                Set objBld = New clsBasisData
                strBld = GetBuildNm(.Value)
'                Set objBld = Nothing
                
'                If ObjLISComCode.Building.Exists(.Value) Then
'                    ObjLISComCode.Building.KeyChange (.Value)
'                End If
                .Col = enCOLLIST.tcBUILDNM: .Value = strBld 'ObjLISComCode.Building.Fields("buildnm")
            
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
                    .Col = enCOLLIST.tcSTATFLAG: .Text = Trim(tmpRs.Fields("StatFg").Value & "")
                End If
            
            End Select
          
DataSet:
            .Col = enCOLLIST.tcTESTNM:  .Text = Trim("" & tmpRs.Fields("TestNm").Value)     'ó���
            '������ ó�� ������
            
            Select Case tmpRs.Fields("orddiv")
                Case APS_ORDDIV: .ForeColor = &H5E3F00     '&HDF6A3E     '&H00DF6A3E&�ణ �Ķ���
                Case BBS_ORDDIV: .ForeColor = &H496835     '&H6C6181     '&H81815A     '�ణ���   &H00845584&�����
                Case LIS_ORDDIV: .ForeColor = &H553755
            End Select
            If Trim("" & tmpRs.Fields("WorkArea").Value) = "OR" Or Trim("" & tmpRs.Fields("WorkArea").Value) = "RI" Then
                .Col = enCOLLIST.tcTESTNM: .ForeColor = DCM_LightRed
            End If
            .Col = enCOLLIST.tcSTATFG:  .Text = IIf("" & tmpRs.Fields("StatFg").Value = "1", "Y", "N") '���޿���
                                        .ForeColor = DCM_Red                                '������
            .Col = enCOLLIST.tcREQDTTM: .Text = Format("" & tmpRs.Fields("ReqDt").Value, CS_DateMask) & " " & _
                                         Format("" & tmpRs.Fields("ReqTm").Value, CS_TimeLongMask)      '���ä���Ͻ�
            .Col = enCOLLIST.tcORDDATE: .Text = Trim("" & tmpRs.Fields("OrdDt").Value)      'ó����
            .Col = enCOLLIST.tcORDNUM:  .Text = Trim("" & tmpRs.Fields("OrdNo").Value)      'ó���ȣ
            .Col = enCOLLIST.tcORDSEQ:  .Text = Trim("" & tmpRs.Fields("OrdSeq").Value)     'ó��Seq
            .Col = enCOLLIST.tcTESTCD:  .Text = Trim("" & tmpRs.Fields("OrdCd").Value)      '�˻��ڵ�

'            Call ObjLISComCode.LisItem.KeyChange(.Text)
            .Col = enCOLLIST.tcLABDIV:  .Text = GetLabDiv(Trim("" & tmpRs.Fields("OrdCd").Value)) 'ObjLISComCode.LisItem.Fields("labdiv")      'LabDiv

            .Col = enCOLLIST.tcSPCCD:   .Text = Trim("" & tmpRs.Fields("SpcCd").Value)      '��ü�ڵ�

'            Call ObjLISComCode.LisSpc.KeyChange(.Text)
            Dim strSpcAbbr As String
            Dim strLabRng As String
            Call GetSpcInfo(.Text, strSpcAbbr, strLabRng)
            .Col = enCOLLIST.tcSPCABBR:  .Text = Trim("" & tmpRs.Fields("spcnm5").Value)         '��ü����
            .Col = enCOLLIST.tcLABRANGE: .Text = strLabRng 'ObjLISComCode.LisSpc.Fields("labrange")    '�̻���������ȣ����

            .Col = enCOLLIST.tcWORKAREA: .Text = Trim("" & tmpRs.Fields("WorkArea").Value)  'WorkArea
            .Col = enCOLLIST.tcSTORECD:  .Text = Trim("" & tmpRs.Fields("StoreCd").Value)   '�����ڵ�
            .Col = enCOLLIST.tcTESTDIV:  .Text = Trim("" & tmpRs.Fields("TestDiv").Value)   '�˻籸��
            .Col = enCOLLIST.tcMULTIFG:  .Text = Trim("" & tmpRs.Fields("MultiFg").Value)   '������ü����
            .Col = enCOLLIST.tcSPCGRP:   .Text = Trim("" & tmpRs.Fields("SpcGrp").Value)    '��ü��
            .Col = enCOLLIST.tcORDDOCT:  .Text = Trim("" & tmpRs.Fields("OrdDoct").Value)   'ó����
                                         'ó���Ǹ�
                                         If .Text <> "" And lblDoctNm.Caption = "" Then
                                            lblDoctNm.Caption = GetDoctNm(tmpRs.Fields("orddoct").Value & "") 'Trim("" & tmpRs.Fields("DoctNm").Value)
                                         End If
            .Col = enCOLLIST.tcMAJDODT:  .Text = Trim("" & tmpRs.Fields("MajDoct").Value)   '��ġ��
            .Col = enCOLLIST.tcDEPTCD:   .Text = Trim("" & tmpRs.Fields("DeptCd").Value)    '�����
                                         '�������
                                         If .Text <> "" And lblDeptNm.Caption = "" Then
'                                            Dim objDept As clsBasisData
                                            Dim strDept As String
'                                            Set objDept = Nothing
'                                            Set objDept = New clsBasisData
                                            strDept = GetDeptNm(.Text)
'                                            Set objDept = Nothing
                                            
'                                            If ObjLISComCode.DeptCd.Exists(.Text) Then
'                                                ObjLISComCode.DeptCd.KeyChange (.Text)
                                                lblDeptNm.Caption = strDept ' ObjLISComCode.DeptCd.Fields("deptnm")
'                                            End If
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
    
NoData:
    Set tmpRs = Nothing
    Set objProInSts = Nothing
   
End Sub

Private Function GetLabDiv(ByVal vTestCd As String) As String
    Dim Rs As Recordset
    Dim strSQL As String
    
    strSQL = " select a.testcd,a.applydt,b.field2 from " & T_LAB001 & " a, " & T_LAB032 & " b "
    strSQL = strSQL & " where " & DBW("b.cdindex=", LC3_WorkArea)
    strSQL = strSQL & " and a.workarea=b.cdval1"
    strSQL = strSQL & " and " & DBW("a.testcd=", vTestCd)
    
    Set Rs = New Recordset
    Rs.Open strSQL, DBConn
    If Rs.EOF = False Then
        GetLabDiv = Rs.Fields("field2").Value & ""
    End If
    Set Rs = Nothing
End Function

Private Sub GetSpcInfo(ByVal vSpcCd As String, ByRef vSpcAbbr As String, _
                            ByRef vLabRng As String)
    Dim Rs As Recordset
    Dim strSQL As String
    
    strSQL = " select  a.cdval1 spccd, a.field4 spcnm, a.field3 spcabbr, a.field5 spcbarnm,  " & _
            " a.field1 multifg, a.field2 spcgrp, b.field2 labrange " & _
            " from " & T_LAB032 & " b, " & T_LAB032 & " a " & _
            " where " & DBW("a.cdindex =", LC3_Specimen) & _
            " and " & DBW("a.cdval1=", vSpcCd) & _
            " and    " & DBJ("b.cdindex ='C217'") & _
            " and    " & DBJ("b.cdval1  =* a.field2")

    Set Rs = New Recordset
    Rs.Open strSQL, DBConn
    
    If Rs.EOF = False Then
        vSpcAbbr = Rs.Fields("spcabbr").Value & ""
        vLabRng = Rs.Fields("labrange").Value & ""
    End If
    Set Rs = Nothing
End Sub

Private Sub TestBuilding_Search()
    
    Dim objSQL As clsBBSCollection
    Dim strtmp As String
    
    Set objSQL = New clsBBSCollection
    
    With objSQL
        If P_ApplyBuildingInfo Then
            If mvarWardId = "" Then
                strBlgCd = ObjSysInfo.BuildingCd
            Else
                strBlgCd = .Get_BuildingCd(UCase(mvarWardId))
            End If
        Else
            strBlgCd = "10"
        End If
        
        strtmp = .TestBuildCd(strBlgCd)
        strErBldCd = medGetP(strtmp, 1, COL_DIV)
        strGBldCd = medGetP(strtmp, 2, COL_DIV)
    End With
    
    Set objSQL = Nothing
    
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
    dtpColDtTm.Value = GetSystemDate
    dtpColDtTm.Enabled = False
    

    dtpQDt.Visible = False
    LisLabel4(4).Visible = False
    
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
'    Set MyPatient.objDB = dbconn
    DoEvents
   
    blnCleared = True
   
End Sub


Public Sub Call_PtId_KeyPress()

    Call txtPtId_KeyPress(vbKeyReturn)

End Sub
