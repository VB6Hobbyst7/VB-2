VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmOrderCode 
   Caption         =   "��� �ڵ� ����"
   ClientHeight    =   7755
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11685
   LinkTopic       =   "Form1"
   ScaleHeight     =   7755
   ScaleWidth      =   11685
   StartUpPosition =   2  'ȭ�� ���
   Begin FPSpread.vaSpread vasList 
      Height          =   6945
      Left            =   60
      TabIndex        =   0
      Top             =   690
      Width           =   7245
      _Version        =   393216
      _ExtentX        =   12779
      _ExtentY        =   12250
      _StockProps     =   64
      ColHeaderDisplay=   1
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����ü"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   20
      RetainSelBlock  =   0   'False
      ScrollBarExtMode=   -1  'True
      ScrollBars      =   2
      SpreadDesigner  =   "frmOrderCode.frx":0000
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   585
      Left            =   60
      TabIndex        =   12
      Top             =   60
      Width           =   11565
      _Version        =   65536
      _ExtentX        =   20399
      _ExtentY        =   1032
      _StockProps     =   15
      Caption         =   "   Cobas4800 ��� �ڵ� ����"
      ForeColor       =   4194304
      BackColor       =   16056319
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      Alignment       =   1
   End
   Begin VB.Frame Frame1 
      Height          =   6945
      Left            =   7320
      TabIndex        =   1
      Top             =   690
      Width           =   4275
      Begin IF_Cobas4800�����ϼ���.MDButton cmdClose 
         Height          =   585
         Left            =   3120
         TabIndex        =   38
         Top             =   6150
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "����"
      End
      Begin IF_Cobas4800�����ϼ���.MDButton cmdCancel 
         Height          =   585
         Left            =   2100
         TabIndex        =   37
         Top             =   6150
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Clear"
      End
      Begin IF_Cobas4800�����ϼ���.MDButton cmdDelete 
         Height          =   585
         Left            =   1080
         TabIndex        =   36
         Top             =   6150
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "����"
      End
      Begin IF_Cobas4800�����ϼ���.MDButton cmdSave 
         Height          =   585
         Left            =   90
         TabIndex        =   35
         Top             =   6150
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "����"
      End
      Begin VB.Frame Frame2 
         Caption         =   "[����]"
         Height          =   1635
         Left            =   120
         TabIndex        =   42
         Top             =   4200
         Visible         =   0   'False
         Width           =   3975
         Begin VB.CheckBox chkPos 
            Caption         =   "�ε�ȣ����"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2310
            TabIndex        =   47
            Top             =   1155
            Width           =   1485
         End
         Begin VB.TextBox txtPos 
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   11.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1410
            TabIndex        =   46
            Top             =   1080
            Width           =   795
         End
         Begin VB.CheckBox chkNeg 
            Caption         =   "�ε�ȣ����"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2310
            TabIndex        =   45
            Top             =   645
            Width           =   1455
         End
         Begin VB.ComboBox cboArr 
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmOrderCode.frx":1DA6
            Left            =   690
            List            =   "frmOrderCode.frx":1DB3
            TabIndex        =   44
            Top             =   210
            Width           =   1215
         End
         Begin VB.TextBox txtNeg 
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   11.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1410
            TabIndex        =   43
            Top             =   585
            Width           =   795
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "�缺���ذ�"
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
            Left            =   150
            TabIndex        =   50
            Top             =   1155
            Width           =   1125
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "�������ذ�"
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
            Left            =   150
            TabIndex        =   49
            Top             =   675
            Width           =   1125
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "�缺"
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
            Left            =   150
            TabIndex        =   48
            Top             =   270
            Width           =   450
         End
      End
      Begin VB.ComboBox cboGubun 
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmOrderCode.frx":1DC9
         Left            =   1410
         List            =   "frmOrderCode.frx":1DD9
         TabIndex        =   40
         Top             =   6150
         Visible         =   0   'False
         Width           =   2565
      End
      Begin VB.PictureBox picEquip 
         Appearance      =   0  '���
         BackColor       =   &H80000005&
         BorderStyle     =   0  '����
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3720
         Picture         =   "frmOrderCode.frx":1E23
         ScaleHeight     =   330
         ScaleWidth      =   330
         TabIndex        =   28
         Top             =   1140
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CheckBox Check2 
         Height          =   285
         Left            =   1290
         TabIndex        =   34
         Top             =   6000
         Value           =   1  'Ȯ��
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.ComboBox cmbType 
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmOrderCode.frx":4451
         Left            =   1410
         List            =   "frmOrderCode.frx":4453
         TabIndex        =   32
         Top             =   1980
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.CheckBox Check1 
         Height          =   285
         Left            =   1440
         TabIndex        =   27
         Top             =   2400
         Value           =   1  'Ȯ��
         Width           =   435
      End
      Begin VB.TextBox txtSeq 
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1410
         TabIndex        =   24
         Top             =   2760
         Width           =   945
      End
      Begin VB.TextBox txtDelta 
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         TabIndex        =   21
         Top             =   5880
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.TextBox txtPHigh 
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1770
         TabIndex        =   19
         Top             =   5820
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.TextBox txtPLow 
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   630
         TabIndex        =   17
         Top             =   5910
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.TextBox txtRefHigh 
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2790
         TabIndex        =   15
         Top             =   3210
         Width           =   945
      End
      Begin VB.TextBox txtRefLow 
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1410
         TabIndex        =   13
         Top             =   3210
         Width           =   945
      End
      Begin VB.TextBox txtMuch 
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1410
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   240
         Width           =   2655
      End
      Begin VB.TextBox txtName 
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1410
         TabIndex        =   9
         Top             =   1530
         Width           =   2655
      End
      Begin VB.TextBox txtDec 
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1410
         TabIndex        =   7
         Top             =   3660
         Width           =   945
      End
      Begin VB.TextBox txtCode 
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1410
         TabIndex        =   5
         Top             =   1104
         Width           =   2655
      End
      Begin VB.TextBox txtEquipCode 
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1410
         TabIndex        =   3
         Top             =   672
         Width           =   2655
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "ó�汸��"
         BeginProperty Font 
            Name            =   "������"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   390
         TabIndex        =   41
         Top             =   5760
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "CutOff"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2670
         TabIndex        =   39
         Top             =   5880
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "Flag"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   330
         TabIndex        =   33
         Top             =   5880
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "��üŸ��"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   31
         Top             =   2010
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "�˻翩��"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   26
         Top             =   2430
         Width           =   1020
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "��    ��"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   25
         Top             =   2820
         Width           =   1050
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   480
         TabIndex        =   23
         Top             =   5550
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "��    Ÿ"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   60
         TabIndex        =   22
         Top             =   6090
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1440
         TabIndex        =   20
         Top             =   6030
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "��    ��"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   150
         TabIndex        =   18
         Top             =   5820
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2490
         TabIndex        =   16
         Top             =   3285
         Width           =   135
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "�������"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   14
         Top             =   3285
         Width           =   1020
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "��񱸺�"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   10
         Top             =   315
         Width           =   1020
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "�� �� ��"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   8
         Top             =   1605
         Width           =   1035
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "�� Ȯ ��"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   6
         Top             =   3750
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "�˻��ڵ�"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   4
         Top             =   1170
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "����ڵ�"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   2
         Top             =   750
         Width           =   1020
      End
   End
   Begin Threed.SSFrame ssfrmExam 
      Height          =   5205
      Left            =   570
      TabIndex        =   29
      Top             =   2220
      Visible         =   0   'False
      Width           =   6465
      _Version        =   65536
      _ExtentX        =   11404
      _ExtentY        =   9181
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin FPSpread.vaSpread vasExam 
         Height          =   4905
         Left            =   180
         TabIndex        =   30
         Top             =   210
         Width           =   6225
         _Version        =   393216
         _ExtentX        =   10980
         _ExtentY        =   8652
         _StockProps     =   64
         ColHeaderDisplay=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   4
         RowHeaderDisplay=   0
         ScrollBars      =   2
         SpreadDesigner  =   "frmOrderCode.frx":4455
      End
   End
End
Attribute VB_Name = "frmOrderCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub ClearText()

    txtEquipCode = ""
    txtCode = ""
    txtName = ""
    txtDec = "1"
    txtRefLow = ""
    txtRefHigh = ""
    txtPLow = ""
    txtPHigh = ""
    txtDelta = ""
    txtSeq = ""
    cmbType = ""
    
    cboArr.ListIndex = 0
    txtNeg = ""
    chkNeg.Value = 0
    txtPos = ""
    chkPos.Value = 0
    
    cboGubun.Text = ""
    
    cmdSave.Caption = "����"
End Sub

Sub DisplayList()
    ClearSpread vasList
    
'    SQL = "SELECT equipcode, examcode, examname, resprec, reflow, refhigh, paniclow, panichigh, deltavalue " & CR & _
'          "  From equipexam " & CR & _
'          " WHERE equipno = '" & gEquip & "' " & CR & _
'          " Order by seqno, EquipCode "
          
    SQL = "SELECT equipcode, examcode, examname, resprec, reflow, refhigh, paniclow, panichigh, deltavalue, examflag, seqno, examtype, " & CR & _
          "  CutOffFlag, NegValue, NegEqual, PosValue, PosEqual, cutoff, ordgubun " & CR & _
          "  From equipexam " & CR & _
          " WHERE equipno = '" & gEquip & "' " & CR & _
          " Order by seqno "
          
    db_select_Vas gLocal, SQL, vasList
    
    vasList.MaxRows = vasList.DataRowCnt
End Sub

Function ExistOfEquipCode(asEquipCode As String, Optional asSuga As String = "") As Integer
'����ڵ�� �����ڵ忡 �ش��ϴ� ����Ÿ ���� Ȯ�� �ϴ� procedure

    ExistOfEquipCode = -1
    
    If asEquipCode = "" Then
        Exit Function
    End If
    
    SQL = "SELECT equipcode, examcode, examname, resprec, reflow, refhigh, paniclow, panichigh, deltavalue   " & CR & _
          "  From equipexam " & CR & _
          " WHERE equipno = '" & gEquip & "' " & CR & _
          "   AND equipcode = '" & asEquipCode & "' "
    If Trim(asSuga) <> "" Then
        SQL = SQL & CR & _
          "   AND examcode = '" & asSuga & "' "
    End If
    res = db_select_Col(gLocal, SQL)
    If res = 0 Then
        ExistOfEquipCode = 0
        Exit Function
    ElseIf res = -1 Then
        ExistOfEquipCode = -1
        Exit Function
    End If
    
    If Trim(gReadBuf(0)) <> asEquipCode Or Trim(gReadBuf(1)) <> asSuga Then
        Exit Function
    End If
        
    ExistOfEquipCode = 1
End Function

Function Select_Suga_Info(asSuga As String) As Integer
    Select_Suga_Info = -1
    
    If Trim(asSuga) = "" Then
        Exit Function
    End If
    
'    If Not Connect_Server Then
'        cn_Server_Flag = False
'        Exit Function
'    Else
'        cn_Server_Flag = True
'    End If
    
    'Connect_Server_Neosoft
    
    SQL = " Select LABM_ID, LABM_NAME " & CR & _
          " from CC_LABM " & CR & _
          " where LABM_ID = '" & Trim(asSuga) & "' "

    res = db_select_Col_Neo(gServer, SQL)
    
'    If cn_Server_Flag Then DisConnect_Server
    
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    ElseIf res = 0 Then
        Select_Suga_Info = 0
        Exit Function
    End If
    If Trim(gReadBuf(0)) <> asSuga Then
        Select_Suga_Info = 0
        Exit Function
    End If
    
    txtDec = ""
    txtName = Trim(gReadBuf(1))
    txtRefLow = ""
    txtRefHigh = ""
    txtPLow = ""
    txtPHigh = ""
    
    txtDelta = ""
    
    Select_Suga_Info = 1
End Function

Private Sub Check1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtSeq.SetFocus
    End If
End Sub

Private Sub cmdCancel_Click()
    ClearText
    txtEquipCode.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    If Trim(txtEquipCode) = "" Then
        txtEquipCode.SetFocus
        Exit Sub
    End If
    
    
'    If Trim(txtCode) = "" Then
'        txtCode.SetFocus
'        Exit Sub
'    End If
        
    db_BeginTran gLocal
    
    SQL = "Delete From equipexam " & CR & _
          "Where equipno = '" & gEquip & "' " & CR & _
          "  and equipcode = '" & Trim(txtEquipCode) & "' " & CR & _
          "  and examcode = '" & Trim(txtCode) & "' "
    res = SendQuery(gLocal, SQL)
    If res = -1 Then
        db_RollBack gLocal
        Exit Sub
    End If
    
    db_Commit gLocal

    DisplayList
    
    cmdCancel_Click

End Sub

Private Sub cmdSave_Click()
    Dim lsFlag          As String
    Dim lsResFlag       As String
    Dim liSeqNo         As Integer
    Dim lsCutOffFlag    As String
    
    Dim lsNegFlag As String
    Dim lsPosFlag As String
    Dim lsValFlag As String
    
    If Trim(txtEquipCode) = "" Then
        txtEquipCode.SetFocus
        MsgBox "����ڵ带 �Է��ϼ���", vbInformation
        Exit Sub
    End If
    
'    If Trim(txtCode) = "" Then
'        txtCode.SetFocus
'        MsgBox "�˻��ڵ带 �Է��ϼ���", vbInformation
'        Exit Sub
'    End If
    
    If Trim(txtDec) = "" Then
        txtDec.Text = 1
'        txtDec.SetFocus
'        Exit Sub
    End If
    
'    If IsNumeric(txtSeq) Then
'        liSeqNo = CInt(txtSeq)
'    Else
'        liSeqNo = 0
'    End If
    
    If Check1.Value = 1 Then
        lsFlag = "1"
    Else
        lsFlag = "0"
    End If
    
    If Check2.Value = 1 Then
        lsResFlag = "1"
    Else
        lsResFlag = "0"
    End If
    
    lsCutOffFlag = cboArr.ListIndex

    lsNegFlag = chkNeg.Value
    lsPosFlag = chkPos.Value
    'lsValFlag = chkValFlag.Value
    lsValFlag = 0
    
    db_BeginTran gLocal
    'examcode, examname, resprec, refmlow, refmhigh, refwlow, refwhigh
    res = ExistOfEquipCode(Trim(txtEquipCode), Trim(txtCode))
    If res = 1 Then
        '"Set resprec = '" & Trim(lsResFlag) & "', " & vbCrLf & _

        SQL = "Update equipexam " & CR & _
              "Set resprec = '" & Trim(txtDec) & "', " & vbCrLf & _
              "    examname = '" & Trim(txtName) & "', " & vbCrLf & _
              "    reflow = '" & Trim(txtRefLow) & "', " & vbCrLf & _
              "    refhigh = '" & Trim(txtRefHigh) & "', " & vbCrLf & _
              "    paniclow = '" & Trim(txtPLow) & "', " & vbCrLf & _
              "    panichigh = '" & Trim(txtPHigh) & "', " & vbCrLf & _
              "    deltavalue = '" & Trim(txtDelta) & "', " & vbCrLf & _
              "    examflag = " & lsFlag & ", " & vbCrLf & _
              "    examtype = '" & Trim(cmbType) & "', " & vbCrLf & _
              "    CutOffFlag = " & lsCutOffFlag & ", " & vbCrLf & _
              "    NegValue = '" & Trim(txtNeg) & "', " & vbCrLf & _
              "    NegEqual = " & lsNegFlag & ", " & vbCrLf & _
              "    PosValue = '" & Trim(txtPos) & "', " & vbCrLf & _
              "    cutoff = '" & Trim(lsValFlag) & "', " & vbCrLf & _
              "    PosEqual = " & lsPosFlag & ", " & vbCrLf & _
              "    ordgubun = '" & Left(cboGubun.Text, 1) & "', " & vbCrLf & _
              "    seqno = '" & Trim(txtSeq.Text) & "' " & vbCrLf & _
              "Where equipno = '" & gEquip & "' " & vbCrLf & _
              "  and equipcode = '" & Trim(txtEquipCode) & "' " & vbCrLf & _
              "  and examcode = '" & Trim(txtCode) & "' "
    ElseIf res = 0 Then
        SQL = "Insert Into equipexam (equipno,equipcode, examcode, examname, resprec, reflow, refhigh, paniclow, panichigh, deltavalue, examflag, seqno, examtype, " & CR & _
              " CutOffFlag, NegValue, NegEqual, PosValue, PosEqual, cutoff, ordgubun) " & CR & _
              "Values ('" & gEquip & "', '" & Trim(txtEquipCode) & "', '" & Trim(txtCode) & "', '" & Trim(txtName.Text) & "', '" & Trim(txtDec) & "', '" & Trim(txtRefLow) & "', '" & Trim(txtRefHigh) & "', '" & Trim(txtPLow) & "', '" & Trim(txtPHigh) & "', '" & Trim(txtDelta) & "', " & lsFlag & ", '" & Trim(txtSeq.Text) & "', '" & Trim(cmbType) & "', " & CR & _
              "        '" & lsCutOffFlag & "', '" & Trim(txtNeg) & "', " & lsNegFlag & ", '" & Trim(txtPos) & "', " & lsPosFlag & ",  '" & Trim(lsValFlag) & "','" & Left(cboGubun.Text, 1) & "') "
    End If

    res = SendQuery(gLocal, SQL)
    If res = -1 Then
        db_RollBack gLocal
        SaveQuery SQL
        Exit Sub
    End If
    
    db_Commit gLocal
    
    'gEquip = txtMuch
    
    DisplayList
    
    cmdCancel_Click
End Sub

Private Sub Form_Load()
    Me.Height = 8600
    Me.Width = 11970
            
    ClearText
    DisplayList
    
    cmbType.AddItem "SE", 0
    cmbType.AddItem "TU", 1
    cmbType.AddItem "UR", 2
    cmbType.AddItem "SF", 3
    
    
    txtMuch = gEquip
End Sub


Private Sub txtDelta_GotFocus()
    SelectFocus txtDelta
End Sub

Private Sub txtDelta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdSave.SetFocus
    End If
End Sub

Private Sub txtEquipCode_GotFocus()
    SelectFocus txtEquipCode
End Sub

Private Sub txtEquipCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txtEquipCode = "" Then
            txtEquipCode.SetFocus
            Exit Sub
        End If
        txtCode.SetFocus
    End If
End Sub

Private Sub txtDec_GotFocus()
    SelectFocus txtDec
End Sub

Private Sub txtDec_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txtDec = "" Then
            txtDec.SetFocus
'            Exit Sub
        End If
        
        'txtRefLow.SetFocus
        cmdSave.SetFocus
    End If
End Sub

Private Sub txtcode_GotFocus()
    SelectFocus txtCode
End Sub

Private Sub txtcode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtCode = UCase(txtCode)
        res = ExistOfEquipCode(Trim(txtEquipCode), Trim(txtCode))
        If res = -1 Then
            txtCode.SetFocus
            Exit Sub
        ElseIf res = 0 Then
            cmdSave.Caption = "����"
            'res = Select_Suga_Info(txtCode)
            'If res <= 0 Then
            '    MsgBox "�˻��ȣ�� �������� �ʽ��ϴ�", vbExclamation
            '    txtCode.SetFocus
            '    Exit Sub
            'End If
            
        ElseIf res = 1 Then
            cmdSave.Caption = "����"
            txtName = Trim(gReadBuf(2))
            'txtDec = Trim(gReadBuf(3))
            txtRefLow = Trim(gReadBuf(5))
            txtRefHigh = Trim(gReadBuf(6))
        End If
        
        txtName.SetFocus
    End If
End Sub

Private Sub txtPHigh_GotFocus()
    SelectFocus txtPHigh
End Sub

Private Sub txtPHigh_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtDelta.SetFocus
    End If
End Sub

Private Sub txtPLow_GotFocus()
    SelectFocus txtPLow
End Sub

Private Sub txtPLow_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtPHigh.SetFocus
    End If
End Sub

Private Sub txtRefhigh_GotFocus()
    SelectFocus txtRefHigh
End Sub

Private Sub txtRefhigh_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        'txtPLow.SetFocus
        cmdSave.SetFocus
    End If
End Sub

'Private Sub Check2_GotFocus()
'    SelectFocus Check2
'End Sub

'Private Sub Check2_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'        'txtPLow.SetFocus
'        cmdSave.SetFocus
'    End If
'End Sub



Private Sub txtRefLow_GotFocus()
    SelectFocus txtRefLow
End Sub

Private Sub txtRefLow_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtRefHigh.SetFocus
    End If
End Sub

Private Sub txtMuch_GotFocus()
    SelectFocus txtMuch
End Sub

Private Sub txtMuch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Trim(txtMuch.Text) = "" Then
            txtMuch.SetFocus
            Exit Sub
        End If
        txtEquipCode.SetFocus
    End If
End Sub

Private Sub cmbType_GotFocus()
    SelectFocus cmbType
End Sub

Private Sub cmbType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Trim(cmbType.Text) = "" Then
            cmbType.SetFocus
            Exit Sub
        End If

        'txtDec.SetFocus
        'cmdSave.SetFocus
        
        Check1.SetFocus
    End If
End Sub

Private Sub txtName_GotFocus()
    SelectFocus txtName
End Sub

Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Trim(txtName.Text) = "" Then
            txtName.SetFocus
            Exit Sub
        End If
        
        'cmbType.SetFocus
        
        txtSeq.SetFocus
    End If
End Sub



Private Sub txtSeq_GotFocus()
    SelectFocus txtSeq
End Sub

Private Sub txtSeq_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Trim(txtSeq.Text) = "" Then
            txtSeq.SetFocus
            Exit Sub
        End If

        'txtDec.SetFocus
        txtRefLow.SetFocus
        'txtRefLow.SetFocus
    End If
End Sub

Private Sub vasExam_Click(ByVal Col As Long, ByVal Row As Long)
    If Row = 0 Then
        vasSort vasExam, Col
    End If
End Sub

Private Sub vasExam_DblClick(ByVal Col As Long, ByVal Row As Long)
    If Row < 1 Or Row > vasExam.DataRowCnt Then Exit Sub
    
    txtCode = Trim(GetText(vasExam, Row, 1))
    txtName = Trim(GetText(vasExam, Row, 2))
    txtSeq = Trim(GetText(vasExam, Row, 4))
    Check1.Value = 1
    
    ssfrmExam.Visible = False
    
    cmdSave.SetFocus
End Sub

Private Sub vasList_Click(ByVal Col As Long, ByVal Row As Long)
    Dim i As Integer
    
    If Row = 0 Then
        Select Case Col
        Case 1
            vasSort vasList, 1, 2
        Case 2
            vasSort vasList, 2, 1
        End Select
        Exit Sub
    End If
    
    If Row < 1 Or Row > vasList.DataRowCnt Then
        cmdSave.Caption = "����"
        ClearText
        Exit Sub
    End If
    
    txtEquipCode = Trim(GetText(vasList, Row, 1))
    txtCode = Trim(GetText(vasList, Row, 2))
    txtName = Trim(GetText(vasList, Row, 3))
    txtDec = Trim(GetText(vasList, Row, 4))
    txtRefLow = Trim(GetText(vasList, Row, 5))
    txtRefHigh = Trim(GetText(vasList, Row, 6))
    txtPLow = Trim(GetText(vasList, Row, 7))
    txtPHigh = Trim(GetText(vasList, Row, 8))
    txtDelta = Trim(GetText(vasList, Row, 9))
    
    cmbType = Trim(GetText(vasList, Row, 12))
    
    If Trim(GetText(vasList, Row, 4)) = "1" Then
        Check2.Value = 1
    Else
        Check2.Value = 0
    End If
    
    If Trim(GetText(vasList, Row, 10)) = "1" Then
        Check1.Value = 1
    Else
        Check1.Value = 0
    End If
    txtSeq = Trim(GetText(vasList, Row, 11))
    
    If Trim(GetText(vasList, Row, 13)) = "1" Then
        cboArr.ListIndex = 1
    ElseIf Trim(GetText(vasList, Row, 13)) = "2" Then
        cboArr.ListIndex = 2
    Else
        cboArr.ListIndex = 0
    End If
    txtNeg.Text = Trim(GetText(vasList, Row, 14))
    If Trim(GetText(vasList, Row, 15)) = "1" Then
        chkNeg.Value = 1
    Else
        chkNeg.Value = 0
    End If
    
    txtPos.Text = Trim(GetText(vasList, Row, 16))
    If Trim(GetText(vasList, Row, 17)) = "1" Then
        chkPos.Value = 1
    Else
        chkPos.Value = 0
    End If
    
'    If Trim(GetText(vasList, Row, 18)) = "1" Then
'        chkValFlag.Value = 1
'    Else
'        chkValFlag.Value = 0
'    End If
    
    cboGubun.Text = ""
    For i = 0 To cboGubun.ListCount - 1
        If Left(cboGubun.List(i), 1) = Trim(GetText(vasList, Row, 19)) Then
            cboGubun.ListIndex = i
            Exit For
        End If
    Next i
    
    cmdSave.Caption = "����"
End Sub
