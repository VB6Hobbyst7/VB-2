VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frm167CollectionM 
   Appearance      =   0  '���
   BackColor       =   &H80000005&
   ClientHeight    =   9105
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   14535
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   9105
   ScaleWidth      =   14535
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ȭ������(&C)"
      Height          =   510
      Left            =   4335
      Style           =   1  '�׷���
      TabIndex        =   49
      Tag             =   "0"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00E0E0E0&
      Caption         =   "�� ��(&X)"
      Height          =   510
      Left            =   13140
      Style           =   1  '�׷���
      TabIndex        =   47
      Tag             =   "0"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdSaveNurse 
      BackColor       =   &H00E0E0E0&
      Caption         =   "���κ�ä�� (&P)"
      Height          =   510
      Left            =   11820
      Style           =   1  '�׷���
      TabIndex        =   46
      Tag             =   "0"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "�ϰ�ä��(&S)"
      Height          =   510
      Left            =   3015
      Style           =   1  '�׷���
      TabIndex        =   45
      Tag             =   "15101"
      Top             =   8535
      Width           =   1320
   End
   Begin Crystal.CrystalReport CReport 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame fraWard 
      BackColor       =   &H00DBE6E6&
      Height          =   8280
      Left            =   7260
      TabIndex        =   12
      Top             =   45
      Width           =   7200
      Begin MedControls1.LisLabel LisLabel8 
         Height          =   300
         Left            =   45
         TabIndex        =   13
         Top             =   120
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   529
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
         Caption         =   "��� �ɼ�"
         LeftGab         =   100
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00DBE6E6&
         Height          =   5895
         Left            =   45
         ScaleHeight     =   5835
         ScaleWidth      =   7050
         TabIndex        =   22
         Top             =   2265
         Width           =   7110
         Begin MedControls1.LisLabel lblColNm 
            Height          =   330
            Left            =   345
            TabIndex        =   23
            Top             =   555
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   582
            BackColor       =   13752531
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            Caption         =   ""
            Appearance      =   0
            LeftGab         =   100
         End
         Begin MedControls1.LisLabel lblPtCount 
            Height          =   330
            Left            =   345
            TabIndex        =   24
            Top             =   1440
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   582
            BackColor       =   13752531
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            Caption         =   ""
            Appearance      =   0
            LeftGab         =   100
         End
         Begin FPSpread.vaSpread tblCount 
            Height          =   4530
            Left            =   2415
            TabIndex        =   25
            Tag             =   "15109"
            Top             =   0
            Width           =   2970
            _Version        =   196608
            _ExtentX        =   5239
            _ExtentY        =   7990
            _StockProps     =   64
            AllowDragDrop   =   -1  'True
            AllowMultiBlocks=   -1  'True
            AllowUserFormulas=   -1  'True
            BackColorStyle  =   1
            BorderStyle     =   0
            DisplayRowHeaders=   0   'False
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
            MaxCols         =   3
            MaxRows         =   30
            Protect         =   0   'False
            ScrollBars      =   2
            ShadowColor     =   14737632
            ShadowDark      =   12632256
            ShadowText      =   0
            SpreadDesigner  =   "frm167.frx":0000
            VisibleCols     =   3
            VisibleRows     =   15
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   360
            Index           =   7
            Left            =   345
            TabIndex        =   60
            Top             =   180
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   635
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
            Caption         =   "�� ä����"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   360
            Index           =   8
            Left            =   345
            TabIndex        =   61
            Top             =   1065
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   635
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
            Caption         =   "�� ȯ�ڼ�"
            Appearance      =   0
         End
         Begin VB.Label Label4 
            BackColor       =   &H00DBE6E6&
            Caption         =   "��"
            Height          =   255
            Left            =   1620
            TabIndex        =   26
            Tag             =   "20104"
            Top             =   1515
            Width           =   270
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            X1              =   2400
            X2              =   2400
            Y1              =   0
            Y2              =   4770
         End
      End
      Begin VB.Frame fraPrtOption 
         BackColor       =   &H00DBE6E6&
         Height          =   1245
         Left            =   45
         TabIndex        =   14
         Top             =   345
         Width           =   7110
         Begin VB.CheckBox chkTestdiv 
            BackColor       =   &H00DBE6E6&
            Caption         =   "�˻��ڵ����"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   3330
            TabIndex        =   48
            Top             =   150
            Width           =   1425
         End
         Begin VB.CheckBox chkPrintFg 
            BackColor       =   &H00DBE6E6&
            Caption         =   "��¾���"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   615
            TabIndex        =   18
            Top             =   150
            Width           =   1305
         End
         Begin VB.OptionButton optOption 
            BackColor       =   &H00DBE6E6&
            Caption         =   "���ڵ�Lable And ä�� ����Ʈ"
            Height          =   315
            Index           =   0
            Left            =   2040
            TabIndex        =   17
            Top             =   495
            Width           =   2715
         End
         Begin VB.OptionButton optOption 
            BackColor       =   &H00DBE6E6&
            Caption         =   "���ڵ� Only"
            Height          =   315
            Index           =   1
            Left            =   600
            TabIndex        =   16
            Top             =   465
            Value           =   -1  'True
            Width           =   1365
         End
         Begin VB.TextBox txtCopy 
            Alignment       =   1  '������ ����
            Height          =   345
            Left            =   3000
            TabIndex        =   15
            Top             =   810
            Width           =   750
         End
         Begin MSComCtl2.UpDown UpDown1 
            Height          =   360
            Left            =   3765
            TabIndex        =   19
            Top             =   795
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   635
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MedControls1.LisLabel lblColList 
            Height          =   255
            Left            =   600
            TabIndex        =   20
            Top             =   840
            Width           =   2205
            _ExtentX        =   3889
            _ExtentY        =   450
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
            Alignment       =   1
            Caption         =   "ä������Ʈ ������"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel lblPage 
            Height          =   255
            Left            =   4080
            TabIndex        =   21
            Top             =   870
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   450
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
            Alignment       =   1
            Caption         =   "��"
            Appearance      =   0
         End
      End
      Begin MSComctlLib.ProgressBar pbrPtCnt 
         Height          =   150
         Left            =   210
         TabIndex        =   27
         Top             =   2025
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   265
         _Version        =   393216
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel9 
         Height          =   300
         Left            =   45
         TabIndex        =   28
         Top             =   1620
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   529
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
         Caption         =   "���� ��Ȳ"
         LeftGab         =   100
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  '�������� ����
         BorderColor     =   &H00808080&
         FillColor       =   &H00D8DEDA&
         FillStyle       =   0  '�ܻ�
         Height          =   330
         Index           =   1
         Left            =   45
         Shape           =   4  '�ձ� �簢��
         Top             =   1935
         Width           =   7080
      End
   End
   Begin VB.Frame fraQuery 
      BackColor       =   &H00DBE6E6&
      Height          =   8280
      Left            =   0
      TabIndex        =   29
      Top             =   45
      Width           =   7275
      Begin VB.CheckBox ChkMornFg 
         BackColor       =   &H00800000&
         Caption         =   "�ӻ󺴸� ��ħä��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   255
         Left            =   2145
         TabIndex        =   44
         Top             =   165
         Width           =   2010
      End
      Begin MedControls1.LisLabel LisLabel7 
         Height          =   300
         Left            =   4290
         TabIndex        =   32
         Top             =   135
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   529
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
         Caption         =   "ä�� �Ͻ�"
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel LisLabel1 
         Height          =   300
         Left            =   0
         TabIndex        =   33
         Top             =   150
         Width           =   4245
         _ExtentX        =   7488
         _ExtentY        =   529
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
         Caption         =   "���� ����"
         LeftGab         =   100
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00DBE6E6&
         Height          =   1230
         Left            =   4275
         TabIndex        =   40
         Top             =   360
         Width           =   2955
         Begin VB.OptionButton optApplyColTm 
            BackColor       =   &H00DBE6E6&
            Caption         =   "���� Row�� ����"
            Height          =   285
            Index           =   1
            Left            =   1215
            TabIndex        =   42
            Top             =   300
            Width           =   1710
         End
         Begin VB.OptionButton optApplyColTm 
            BackColor       =   &H00DBE6E6&
            Caption         =   "��ü����"
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   41
            Top             =   300
            Width           =   1035
         End
         Begin MSComCtl2.DTPicker dtpColDtTm 
            Height          =   375
            Left            =   930
            TabIndex        =   43
            Top             =   675
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
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
            CustomFormat    =   "yyy-MM-dd  HH:mm"
            Format          =   106954755
            UpDown          =   -1  'True
            CurrentDate     =   36328.5416666667
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   360
            Index           =   11
            Left            =   60
            TabIndex        =   68
            TabStop         =   0   'False
            Top             =   675
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   635
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
            Caption         =   "ä���Ͻ�"
            Appearance      =   0
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00DBE6E6&
         Height          =   1215
         Left            =   0
         TabIndex        =   34
         Top             =   375
         Width           =   4275
         Begin VB.CommandButton cmdGetOrders 
            BackColor       =   &H00E0E0E0&
            Caption         =   "��ȸ(&F)"
            Height          =   465
            Left            =   3270
            Style           =   1  '�׷���
            TabIndex        =   37
            Tag             =   "0"
            Top             =   690
            Width           =   930
         End
         Begin VB.CommandButton cmdWardList 
            BackColor       =   &H0098A7A5&
            Caption         =   "��"
            Height          =   360
            Left            =   2265
            Style           =   1  '�׷���
            TabIndex        =   36
            Tag             =   "WardID"
            Top             =   255
            Width           =   360
         End
         Begin VB.TextBox txtWardID 
            Appearance      =   0  '���
            BackColor       =   &H00FFFFFF&
            Height          =   345
            Left            =   870
            MaxLength       =   9
            TabIndex        =   35
            Top             =   270
            Width           =   1395
         End
         Begin MSComCtl2.DTPicker dtpToTime 
            Height          =   360
            Left            =   870
            TabIndex        =   38
            Top             =   750
            Width           =   2310
            _ExtentX        =   4075
            _ExtentY        =   635
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
            CustomFormat    =   "yyy-MM-dd    HH:mm:ss"
            Format          =   106954755
            CurrentDate     =   36328
         End
         Begin MedControls1.LisLabel lblWardNm 
            Height          =   345
            Left            =   2640
            TabIndex        =   39
            Top             =   270
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   609
            BackColor       =   13622494
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
            Caption         =   ""
            Appearance      =   0
            LeftGab         =   100
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   360
            Index           =   0
            Left            =   90
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   255
            Width           =   720
            _ExtentX        =   1270
            _ExtentY        =   635
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
            Caption         =   "����ID"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   360
            Index           =   6
            Left            =   90
            TabIndex        =   59
            TabStop         =   0   'False
            Top             =   735
            Width           =   720
            _ExtentX        =   1270
            _ExtentY        =   635
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
            Caption         =   "ó����"
            Appearance      =   0
         End
      End
      Begin FPSpread.vaSpread tblPtList 
         Height          =   6300
         Left            =   15
         TabIndex        =   30
         Top             =   1935
         Width           =   7215
         _Version        =   196608
         _ExtentX        =   12726
         _ExtentY        =   11113
         _StockProps     =   64
         BackColorStyle  =   1
         ColsFrozen      =   4
         DisplayRowHeaders=   0   'False
         EditEnterAction =   5
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   14411494
         MaxCols         =   23
         MaxRows         =   50
         ShadowColor     =   14737632
         ShadowDark      =   14737632
         SpreadDesigner  =   "frm167.frx":0428
         TextTip         =   4
         ScrollBarTrack  =   3
      End
      Begin MedControls1.LisLabel LisLabel6 
         Height          =   300
         Left            =   0
         TabIndex        =   31
         Top             =   1605
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   529
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
   End
   Begin VB.Frame fraNurse 
      BackColor       =   &H00DBE6E6&
      Height          =   8280
      Left            =   5595
      TabIndex        =   0
      Top             =   45
      Width           =   8865
      Begin MSComCtl2.DTPicker DTPNurse 
         Height          =   300
         Left            =   6990
         TabIndex        =   56
         Top             =   1605
         Width           =   1740
         _ExtentX        =   3069
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
         Format          =   106954755
         UpDown          =   -1  'True
         CurrentDate     =   36851.6291666667
      End
      Begin VB.CheckBox chkChangeColTm 
         BackColor       =   &H00800000&
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
         Height          =   255
         Left            =   5580
         TabIndex        =   57
         Top             =   1620
         Width           =   1500
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00800000&
         BorderStyle     =   0  '����
         Height          =   270
         Left            =   1290
         TabIndex        =   52
         Top             =   1605
         Width           =   4890
         Begin VB.CheckBox chkSelAll 
            BackColor       =   &H00800000&
            Caption         =   "��ü(&A)"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H004A4189&
            Height          =   240
            Left            =   45
            TabIndex        =   55
            Top             =   30
            Width           =   1050
         End
         Begin VB.Label lblBBS 
            AutoSize        =   -1  'True
            BackStyle       =   0  '����
            Caption         =   "��������"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00496835&
            Height          =   180
            Left            =   2415
            TabIndex        =   54
            Top             =   60
            Width           =   795
         End
         Begin VB.Shape shp1 
            BackColor       =   &H00496835&
            BackStyle       =   1  '�������� ����
            BorderColor     =   &H00C0C0C0&
            Height          =   165
            Index           =   2
            Left            =   2085
            Shape           =   3  '����
            Top             =   45
            Width           =   330
         End
         Begin VB.Label lblLIS 
            AutoSize        =   -1  'True
            BackStyle       =   0  '����
            Caption         =   "�ӻ󺴸�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00553755&
            Height          =   180
            Left            =   1320
            TabIndex        =   53
            Top             =   45
            Width           =   795
         End
         Begin VB.Shape shp1 
            BackColor       =   &H00553755&
            BackStyle       =   1  '�������� ����
            BorderColor     =   &H00C0C0C0&
            Height          =   165
            Index           =   1
            Left            =   1065
            Shape           =   3  '����
            Top             =   60
            Width           =   330
         End
      End
      Begin MedControls1.LisLabel lblBar 
         Height          =   285
         Left            =   60
         TabIndex        =   1
         Top             =   1605
         Width           =   8820
         _ExtentX        =   15558
         _ExtentY        =   503
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
         Caption         =   "ó�� ����Ʈ"
         LeftGab         =   100
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
         Left            =   1245
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  '����
         TabIndex        =   50
         ToolTipText     =   "�˻� ����ũ�� �Է��ϼ���."
         Top             =   7350
         Width           =   7245
      End
      Begin MedControls1.LisLabel LisLabel5 
         Height          =   285
         Left            =   75
         TabIndex        =   4
         Top             =   135
         Width           =   8715
         _ExtentX        =   15372
         _ExtentY        =   503
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
      Begin MedControls1.LisLabel LisLabel2 
         Height          =   285
         Left            =   180
         TabIndex        =   51
         Top             =   7365
         Width           =   990
         _ExtentX        =   1746
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
      Begin VB.Frame Frame3 
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
         Height          =   1245
         Left            =   60
         TabIndex        =   5
         Top             =   345
         Width           =   8730
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
            Height          =   375
            Left            =   945
            MaxLength       =   10
            TabIndex        =   6
            Top             =   255
            Width           =   1500
         End
         Begin MedControls1.LisLabel lblPtNm 
            Height          =   375
            Left            =   3900
            TabIndex        =   7
            Top             =   255
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   661
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
            Height          =   360
            Left            =   7035
            TabIndex        =   8
            Top             =   255
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   635
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
            Height          =   375
            Left            =   960
            TabIndex        =   9
            Top             =   735
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   661
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
            Height          =   360
            Left            =   3900
            TabIndex        =   10
            Top             =   720
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   635
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
            Height          =   360
            Left            =   7050
            TabIndex        =   11
            Top             =   705
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   635
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
            Height          =   360
            Index           =   9
            Left            =   75
            TabIndex        =   62
            TabStop         =   0   'False
            Top             =   255
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   635
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
            Caption         =   "ȯ�� ID"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   360
            Index           =   10
            Left            =   75
            TabIndex        =   63
            TabStop         =   0   'False
            Top             =   735
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   635
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
            Height          =   360
            Index           =   3
            Left            =   3030
            TabIndex        =   64
            TabStop         =   0   'False
            Top             =   255
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   635
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
            Caption         =   "��     ��"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   360
            Index           =   5
            Left            =   3030
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   720
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   635
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
            Height          =   360
            Index           =   1
            Left            =   6150
            TabIndex        =   66
            TabStop         =   0   'False
            Top             =   255
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   635
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
            Height          =   360
            Index           =   4
            Left            =   6150
            TabIndex        =   67
            TabStop         =   0   'False
            Top             =   720
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   635
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
         Height          =   5415
         Left            =   60
         TabIndex        =   2
         Top             =   1815
         Width           =   8790
         Begin FPSpread.vaSpread tblOrdSheet 
            Height          =   5025
            Left            =   45
            TabIndex        =   3
            Tag             =   "10114"
            Top             =   135
            Width           =   8715
            _Version        =   196608
            _ExtentX        =   15372
            _ExtentY        =   8864
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
            SpreadDesigner  =   "frm167.frx":0F5F
            StartingColNumber=   2
            VirtualRows     =   24
            VisibleCols     =   5
            VisibleRows     =   19
         End
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00EFFFEE&
         BackStyle       =   1  '�������� ����
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Height          =   930
         Index           =   0
         Left            =   75
         Shape           =   4  '�ձ� �簢��
         Top             =   7290
         Width           =   8505
      End
   End
End
Attribute VB_Name = "frm167CollectionM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private objMySql                As clsLISSqlCollection
Private objLISCollect           As clsLISCollectioin
Private MyPatient               As clsPatient
Private WithEvents objMyList    As clsPopUpList
Attribute objMyList.VB_VarHelpID = -1

Private sWorkDt                 As String
Private sWorkTm                 As String
Private mvarWardId              As String
Private mvarDeptCd              As String
Private mvarHosilID             As String
Private mvarRoomID              As String

Private IsFirst                 As Boolean
Private blnCleanFg              As Boolean
Private blnCollectFg            As Boolean             'ä������(�Ѱ��̶��...�Ǹ� True)
Private blnCleared              As Boolean
Private PtFg                    As Boolean
Private MsgFg                   As Boolean
Private OrdFg                   As Boolean
Private SelAllFg                As Boolean

Private intPtCount              As Integer
Private intErrCount             As Integer

Private Const lngMaxRows = 19
Private Const lngRowHeight = 12

Public Event LastFormUnload()

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

Private Sub cmdClear_Click()
    Call ClearRtn(1)
    Call ClearRtn(2)
    If txtPtId.Enabled = True Then txtPtId.Text = ""
    fraQuery.ZOrder 0
    fraWard.ZOrder 0
On Error GoTo Err_Trap
    txtWardID.SetFocus
Err_Trap:
End Sub
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
Private Sub cmdSaveNurse_Click()
    Dim objPrgBar       As jProgressBar.clsProgress
    Dim objDIC          As clsDictionary
    Dim BBSColSuccess   As Boolean
    Dim LISColSuccess   As Boolean
    
    Dim iCheckOrder     As Integer
    Dim ii              As Integer
    
    Dim lngBarCnt       As Long
    Dim lngSelCnt       As Long
    Dim BarCount        As Long
    Dim SelCount        As Long

    If CollectionTargetChk = False Then
       MsgBox "ä���� �׸��� �����ϼ���..", vbInformation, "�׸���"
       tblOrdSheet.SetFocus
       Exit Sub
    End If
    
    iCheckOrder = objLISCollect.CheckSameOrder(tblOrdSheet, 1)       '�ߺ�ó�� Check
    If iCheckOrder > 0 Then GoTo OrdCheck1
    
    Call MouseRunning

    Set objPrgBar = New jProgressBar.clsProgress
    With objPrgBar
        .Container = Me
        .Left = fraNurse.Left + lblBar.Left
        .Top = lblBar.Top - 70
        .Width = lblBar.Width - 10
        .Height = lblBar.Height - 10
        .Message = "���õ� �˻��׸� ���� ä��ó�����Դϴ�..."
'        .SetMyForm Me
'        .Choice = True
'        .XPos = fraNurse.Left + lblBar.Left
'        .YPos = lblBar.Top - 70
'        .XWidth = lblBar.Width - 10
'        .ForeColor = &HFA8B10
'        .Appearance = aPlate
'        .BorderStyle = bsNone
'        .YHeight = lblBar.Height - 10
'        .Msg = "���õ� �˻��׸� ���� ä��ó�����Դϴ�."
'        .Max = 90
'        .Min = 0
'        .Value = 10
        DoEvents
    End With

    DoEvents

    '----------------------------------------------------------
    '������ ������ ���ؼ� �������� �ҷ��� �����Ѵ�.(2001/06/08)
    '----------------------------------------------------------
    
    Call tblordersheet
    
    Set objDIC = New clsDictionary
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
        BBSColSuccess = True:  LISColSuccess = True
        Do Until .EOF
            Select Case .Fields("orddiv")
                Case BBS_ORDDIV: BBSColSuccess = CollectForBBS(.Fields("first"), .Fields("last"), _
                                                                    Format(GetSystemDate, "yyyymmdd"), _
                                                                    Format(GetSystemDate, "HHmmss"), objPrgBar)
                Case LIS_ORDDIV: LISColSuccess = CollectForLIS(.Fields("first"), .Fields("last"), objPrgBar)
            End Select
            .MoveNext
        Loop
    End With
        

    SelCount = 0: lngBarCnt = 0
    

    
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

    Call cmdGetOrders_Click
    Set objDIC = Nothing
    Exit Sub

OrdCheck1:
    tblOrdSheet.Row = iCheckOrder
    tblOrdSheet.Col = 1
    tblOrdSheet.Action = ActionActiveCell
    tblOrdSheet.SetFocus
    Set objDIC = Nothing
    Exit Sub

OrdCheck2:
    tblOrdSheet.Row = iCheckOrder
    tblOrdSheet.Col = 1
    tblOrdSheet.Action = ActionActiveCell
    MsgBox "������ü ������ �����ϴ�. ����� Ȥ�� �ӻ󺴸����� �����ϼ���.", vbInformation + vbOKOnly, "����"
    tblOrdSheet.SetFocus
    Set objDIC = Nothing
    Exit Sub


End Sub

Private Sub dtpColDtTm_Change()

    Dim Resp As VbMsgBoxResult

    If blnCleanFg Then Exit Sub
    If dtpColDtTm.Value < Now Then
        If dtpColDtTm.Tag = "1" Then
            dtpColDtTm.Tag = "0"
        Else
            Resp = MsgBox("ä���ð��� ����ð����� �����Դϴ�. �����Ͻðڽ��ϱ�?", _
                   vbQuestion + vbYesNo, "ä���ð�����")
            If Resp = vbYes Then
                dtpColDtTm.Tag = "1"
            Else
                dtpColDtTm.Tag = "0"
                dtpColDtTm.Value = GetSystemDate
            End If
        End If
    End If

    With tblPtList
        If optApplyColTm(0).Value Then  '��ü
            .Row = 1: .Row2 = .DataRowCnt
            .Col = 19: .Col2 = 19
            .BlockMode = True
            .Text = Format(dtpColDtTm.Value, "YY-MM-DD")
            .BlockMode = False
            .Col = 20: .Col2 = 20
            .BlockMode = True
            .Text = Format(dtpColDtTm.Value, "HH:MM")
            .BlockMode = False
        Else
            If .ActiveRow > .DataRowCnt Then Exit Sub
            .Row = .ActiveRow
            .Col = 19: .Text = Format(dtpColDtTm.Value, "YY-MM-DD")
            .Col = 20: .Text = Format(dtpColDtTm.Value, "HH:MM")
        End If
    End With

End Sub

Private Sub dtpToTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Activate()
    If Not IsFirst Then Exit Sub
    IsFirst = False
    
    dtpToTime.Value = Format(GetSystemDate, "YYYY-MM-DD HH:MM:SS")
    dtpColDtTm.Value = GetSystemDate
    blnCleanFg = True

    txtWardID.Text = "": lblWardNm.Caption = "": txtCopy.Text = 1
    intErrCount = 0
    
On Error GoTo Err_Trap
    
    txtWardID.SetFocus
    chkPrintFg.Value = 0
    optOption(1).Value = True
    Exit Sub
Err_Trap:
    Resume Next
End Sub
Private Sub Form_Load()
    IsFirst = True
    
    If P_MornCollection = False Then
        ChkMornFg.Visible = False
        LisLabel7.Visible = False
        Frame1.Visible = False
        LisLabel1.Width = 5595
        Frame2.Width = 5595
        cmdGetOrders.Left = cmdGetOrders.Left + 100
    End If
    
    Set MyPatient = New clsPatient
    Set objMySql = New clsLISSqlCollection
    Set objLISCollect = New clsLISCollectioin
    
End Sub


'& ��� Option ����
Private Sub chkPrintFg_Click()
    If chkPrintFg.Value = 1 Then
        optOption(0).Value = False
        optOption(1).Value = False
    Else
        optOption(1).Value = True
    End If
End Sub

'% ����
Private Sub cmdExit_Click()
    Set objMySql = Nothing
    Set objLISCollect = Nothing
    Set MyPatient = Nothing
    Set objMyList = Nothing

    Unload Me
    If IsLastForm Then RaiseEvent LastFormUnload
End Sub

'% �ϰ�ä�� ����
Private Sub cmdSave_Click()
    Dim Resp        As VbMsgBoxResult
    Dim intSelCount As Integer
    Dim sBuildCd    As String
    Dim sBuildNm    As String

    Dim strSavePtId As String
    Dim i           As Integer
    Dim j           As Integer


    blnCollectFg = False
    Set objLISCollect = New clsLISCollectioin

    sWorkDt = Format(GetSystemDate, CS_DateDbFormat)
    sWorkTm = Format(GetSystemDate, CS_TimeDbFormat)

    tblCount.Row = 0
    intErrCount = 0
    intSelCount = 0
    strSavePtId = ""

    Call SetLock(True)

    Me.MousePointer = 11

    With tblPtList
        pbrPtCnt.Visible = True
        pbrPtCnt.Max = .DataRowCnt * 3 * 101
        pbrPtCnt.Min = 0
        lblPtCount.Caption = ""

        For i = 1 To .DataRowCnt
            .Row = i

            '* ���ܹ�ư Check
            .Col = 1: If .Value = 1 Then GoTo Skip

            intSelCount = intSelCount + 1

            '* ä������
            .Col = 15   'for LIS
            If Trim(.Value) <> "" Then Call DoCollectionForLIS(i)
            
            If pbrPtCnt.Value + 100 >= pbrPtCnt.Max Then pbrPtCnt.Max = pbrPtCnt.Value + 100
            pbrPtCnt.Value = pbrPtCnt.Value + 100
            DoEvents

            .Col = 17   'for BBS
            If Trim(.Value) <> "" Then Call DoCollectionForBBS(i)
            
            If pbrPtCnt.Value + 100 >= pbrPtCnt.Max Then pbrPtCnt.Max = pbrPtCnt.Value + 100
            pbrPtCnt.Value = pbrPtCnt.Value + 100
            DoEvents

            '* ȯ�ڼ� Count
            .Row = i: .Col = 3
            If strSavePtId <> Trim(.Text) Then
               lblPtCount.Caption = Val(lblPtCount.Caption) + 1
               strSavePtId = .Text
            End If

            '* ä�� Class Initialize
            objLISCollect.InitRtn
            DoEvents
Skip:
        Next

        'ä����
        lblColNm.Caption = ObjSysInfo.EmpId

    End With

    If intSelCount = 0 Then
         Screen.MousePointer = vbDefault  '0
         Call cmdClear_Click
         MsgBox "ó���� ����Ÿ�� �����ϴ�..", vbInformation, "Message"
         Exit Sub
    End If

On Error GoTo Errors

    If blnCollectFg = True Then
    
        pbrPtCnt.Value = pbrPtCnt.Max
        DoEvents
    
        MouseDefault
    
        If intErrCount > 0 Then
             MsgBox CStr(intErrCount) & "���� ������ �߻��߽��ϴ�.."
        Else
        
             If optOption(0).Value Then
                 Call medClearTable(tblPtList)
                 Resp = MsgBox("��� ���������� ä��ó�� �Ǿ����ϴ�.." & vbCrLf & _
                               "ä�븮��Ʈ�� ���� ����Ͻðڽ��ϱ� ? ", vbYesNo, "ä�븮��Ʈ ���")
                 If Resp = vbYes Then
                     For i = 1 To tblCount.DataRowCnt
                         tblCount.Row = i
                         tblCount.Col = 3:  sBuildCd = tblCount.Value
                         tblCount.Col = 1:  sBuildNm = tblCount.Value
                         
                         Call PrintColList(txtWardID.Text, lblWardNm.Caption, sWorkDt, sWorkTm, sBuildCd, sBuildNm)
                     Next
                 End If
             Else
                 Call MsgBox("��� ���������� ä��ó�� �Ǿ����ϴ�..", vbInformation, "�޼���")
             End If
    
             Call ClearRtn(0)
             Call cmdGetOrders_Click
             
        End If
    Else
        Call ClearRtn(0)
        txtWardID.SetFocus
    End If
    pbrPtCnt.Visible = False
    Me.MousePointer = 0
Errors:

End Sub

Private Sub SetLock(ByVal blnLock As Boolean)

    txtWardID.Enabled = Not blnLock
    txtWardID.BackColor = IIf(blnLock, &H8000000F, vbWhite)
    cmdWardList.Enabled = Not blnLock
    dtpToTime.Enabled = Not blnLock
    cmdGetOrders.Enabled = Not blnLock

End Sub

Private Sub DoCollectionForBBS(ByVal Row As Long)
'
    Dim objBar          As clsDictionary
    Dim objDIC          As clsDictionary
    Dim objBBSCollect   As clsBBSCollection
    Dim strPtid         As String       'ȯ��id
    Dim strPtnm         As String       'ȯ�ڸ�
    Dim strColDt        As String      'ä����
    Dim strColTm        As String      'ä���Ͻ�
    Dim strHosilid      As String
    Dim strStatFg       As String
    Dim lngErCnt        As Long
    Dim lngGcnt         As Long
    Dim lngBldRow       As Long
    Dim j               As Long

    Set objDIC = New clsDictionary
    Set objBBSCollect = New clsBBSCollection
    
    Call objBBSCollect.SetWardCol(txtWardID.Text, sWorkDt, sWorkTm)
    
'    Dim objBld As clsBasisData
    Dim strBld As String
    
    With tblPtList
        .Row = Row
        .Col = 3:   strPtid = .Value
        .Col = 4:   strPtnm = .Value
        .Col = 5
                    If .Value = "��" Then   '����
                        lngErCnt = lngErCnt + 1
                    Else
                        lngGcnt = lngGcnt + 1
                    End If
        .Col = 23:  strStatFg = IIf(.Value = "1", "1", "")
        .Col = 12:  strHosilid = Trim(.Value)
        .Col = 19:  strColDt = Format(.Text, LIS_LabDayFormat)
        .Col = 20:  strColTm = Format(.Text, "HHMMss")
        
        objDIC.Clear
        objDIC.FieldInialize "ptid", "ptnm,coldt,coltm,colid,bussdiv,buildcd,hosilid,statfg"

        objDIC.AddNew strPtid, Join(Array(strPtnm, strColDt, strColTm, ObjSysInfo.EmpId, _
                                    enBussDiv.BussDiv_InPatient, ObjSysInfo.BuildingCd, strHosilid, strStatFg), COL_DIV)
        
        If objDIC.RecordCount > 0 Then
            objBBSCollect.WardId = txtWardID.Text
            If objBBSCollect.Set_Collect(objDIC, ObjSysInfo.BuildingCd, , True) Then     '�ϰ�ä����������
'                Call ObjLISComCode.Building.KeyChange(objsysinfo.BuildingCd)
'                Set objBld = Nothing
'                Set objBld = New clsBasisData
                strBld = GetBuildNm(ObjSysInfo.BuildingCd)
'                Set objBld = Nothing
                
                lngBldRow = 0
                For j = 1 To tblCount.DataRowCnt
                    tblCount.Row = j: tblCount.Col = 3
                    If tblCount.Value = ObjSysInfo.BuildingCd Then
                        lngBldRow = j
                        Exit For
                    End If
                Next

                If lngBldRow = 0 Then lngBldRow = tblCount.DataRowCnt + 1
                tblCount.Row = lngBldRow
                tblCount.Col = 1: tblCount.Text = strBld 'ObjLISComCode.Building.Fields("buildnm")
                tblCount.Col = 2: tblCount.Text = Val(tblCount.Text) + 1
                tblCount.Col = 3: tblCount.Text = ObjSysInfo.BuildingCd

                
                Set objBar = New clsDictionary
                Set objBar = objBBSCollect.BldDic
                If objBar.RecordCount > 0 Then
                    BarCode_Print objBar
                    blnCollectFg = True
                End If
            End If
        End If
    End With

    Set objBBSCollect = Nothing
    Set objDIC = Nothing
    Set objBar = Nothing
End Sub

Private Sub BarCode_Print(objDIC As clsDictionary)
    
    Dim objBar      As clsBarcode
    Dim strPtid     As String
    Dim strPtnm     As String
    Dim strColDt    As String
    Dim strColTm    As String
    Dim strSpcNo    As String
    Dim strAccSeq   As String         'SpcYy-SpcNo ������ ��ü��ȣ
    Dim HosilId     As String
    Dim strStatFg   As String
    Dim strBarW_H   As String
    
    
    Set objBar = New clsBarcode
    
'    Set objBAR.MyDB = dbconn
    Set objBar.TableInfo = New clsTables
    Set objBar.FieldInfo = New clsFields
    
    objDIC.MoveFirst

    Do Until objDIC.EOF
        strPtid = medGetP(objDIC.GetString, 1, COL_DIV)
        strPtnm = medGetP(objDIC.GetString, 2, COL_DIV)
        strSpcNo = medGetP(objDIC.GetString, 3, COL_DIV)
        strColDt = medGetP(objDIC.GetString, 4, COL_DIV)
        strColDt = Format(Mid(strColDt, 5, 4), "##/##")
        strColTm = Mid(medGetP(objDIC.GetString, 5, COL_DIV), 1, 4)
        strColTm = Format(strColTm, "##:##")
        HosilId = medGetP(objDIC.GetString, 6, COL_DIV)
        strStatFg = medGetP(objDIC.GetString, 7, COL_DIV)
        
        If HosilId <> "" Then
            strBarW_H = txtWardID.Text & "/" & HosilId
        Else
            strBarW_H = txtWardID.Text
        End If
        
        '��ü��ȣ ��� : 2001.2.8 �߰�
        strAccSeq = Mid(strSpcNo, 1, 2) & "-" & Format(Mid(strSpcNo, 3), "########0")
        strAccSeq = Format(strAccSeq, String(11, "@"))
        '���ڵ� ���
        'ObjBBSComCode.BarInfo
        objBar.Label_PrintOut _
                        BBSName, "XM", "", strAccSeq, strSpcNo, strPtid, _
                        strPtnm, "", "", strStatFg, strBarW_H, _
                        strColDt, strColTm, "", Val(txtCopy)

        objDIC.MoveNext
    Loop
    
    Set objBar = Nothing
End Sub

Private Sub DoCollectionForLIS(ByVal Row As Long)
    Dim Rs          As Recordset
    Dim tmpData()   As String
    Dim tmpDate     As String
    Dim tmpTime     As String
    Dim SqlStmt     As String
    Dim tmpDeptCd   As String
    Dim tmpOrdDoct  As String
    Dim tmpMajDoct  As String
    Dim sWorkarea   As String
    Dim sAccdt      As String
    Dim sBuildCd    As String
    
    Dim blnSuccess  As Boolean
    Dim blnMornCol  As Boolean
    
    Dim iAccseq     As Long
    Dim i           As Integer
    Dim j           As Integer
    
    Dim lngBldRow   As Long

    Call objLISCollect.SetWardCol(sWorkDt, sWorkTm, Trim(txtWardID.Text))
    objLISCollect.MornFg = ChkMornFg.Value      '��ħä������

    ReDim tmpData(0 To 16)
    
    With tblPtList
        .Row = Row
                    tmpData(0) = Mid(Format(Now, "YYYY"), 4)
        .Col = 3:   tmpData(1) = .Value               'ȯ��ID
        .Col = 4:   tmpData(2) = .Value              'ȯ�ڸ�
        .Col = 14:  tmpData(3) = .Value               'ȯ�ڼ���
        .Col = 7:
                    If IsDate(Format(.Value, CS_DateMask)) Then
                        tmpData(4) = DateDiff("y", Format(.Value, CS_DateMask), Now)  'ȯ���Ϸ�
                    Else
                        tmpData(4) = Mid(.Value, 1, 4) & "-01-01"
                        If IsDate(tmpData(4)) Then
                            tmpData(4) = DateDiff("y", tmpData(4), GetSystemDate)
                        Else
                            tmpData(4) = 0
                        End If
                    End If
        .Col = 8:   tmpData(5) = .Value                                             '�Կ���
                    tmpData(6) = Format(GetSystemDate, CS_DateDbFormat)         '�Է���
                    tmpData(7) = Format(GetSystemDate, CS_TimeDbFormat)         '�Է½ð�
                    tmpData(8) = ObjSysInfo.EmpId                                   '�Է���
                    tmpData(9) = ""                                                 '��������ȣ
                    tmpData(10) = Format(GetSystemDate, CS_DateDbFormat)        'ä����
                    objLISCollect.ColTm = Format(GetSystemDate, "HHMMSS")
                    tmpData(11) = ObjSysInfo.EmpId                                  'ä����
        .Col = 2:   tmpData(12) = .Value                                            '����ID
        .Col = 12:  tmpData(13) = .Value                                            '����ID
        .Col = 13:  tmpData(14) = .Value                                            'ȣ��ID
                    tmpData(15) = ""                                                'ħ��ID
                    tmpData(16) = ObjSysInfo.BuildingCd                             '** ä���� ����Ǵ� �ǹ��ڵ�
                    
        Call objLISCollect.SetColData(tmpData)
        
        .Col = 22:  blnMornCol = Choose(Val(.Text) + 1, False, True)
        
        .Col = 9:   tmpDeptCd = .Value                        '�����
        .Col = 10:  tmpOrdDoct = .Value                       'ó����
        .Col = 11:  tmpMajDoct = .Value                       '��ġ��
    End With

    tmpDate = Format(dtpToTime.Value, CS_DateDbFormat)
    tmpTime = Format(dtpToTime.Value, CS_TimeDbFormat)

    ' ó�泻�� �˻�
    If blnMornCol Then
        SqlStmt = objMySql.SqlReadOrderForMornCol(objLISCollect.Ptid, tmpDate, tmpTime)
    Else
        SqlStmt = objMySql.SqlReadWardOrder(objLISCollect.Ptid, tmpDate, tmpTime, , _
                                            enBussDiv.BussDiv_InPatient, , LIS_ORDDIV)
    End If
    
    Set Rs = New Recordset
    Rs.Open SqlStmt, DBConn
    
    If Rs.EOF Then
        blnSuccess = False
        GoTo Err_Trap
    End If

    ReDim tmpData(0 To 20)
    With Rs
        
        For i = 1 To .RecordCount
            tmpData(0) = ObjSysInfo.BuildingCd: sBuildCd = tmpData(0)
            tmpData(1) = Trim("" & .Fields("WorkArea").Value)   'WorkArea
            tmpData(2) = Trim("" & .Fields("SpcCd").Value)      'SpcCd
            tmpData(3) = Trim("" & .Fields("StoreCd").Value)    'StoreCd
            tmpData(4) = Trim("" & .Fields("StatFg").Value)
            tmpData(5) = Format("" & Rs.Fields("ReqDt").Value, CS_DateMask) & " " & _
                         Format("" & Rs.Fields("ReqTm").Value, CS_TimeLongMask)        '���ä���Ͻ�
            tmpData(6) = Trim("" & .Fields("TestDiv").Value)    'TestDiv
            tmpData(7) = Trim("" & .Fields("MultiFg").Value)    'MultiFg
            tmpData(8) = Trim("" & .Fields("SpcGrp").Value)     'SpcGrp
            tmpData(9) = Trim("" & .Fields("OrdDt").Value)      'OrdDt
            tmpData(10) = Trim("" & .Fields("OrdNo").Value)     'OrdNo
            tmpData(11) = Trim("" & .Fields("OrdSeq").Value)    'OrdSeq
            tmpData(12) = Trim("" & .Fields("OrdCd").Value)     'OrdCd
            tmpData(13) = tmpDeptCd
            tmpData(14) = tmpOrdDoct
            tmpData(15) = tmpMajDoct
            tmpData(16) = Trim("" & .Fields("AbbrNm5").Value)   'ó�� ����
            tmpData(17) = Trim("" & .Fields("LabelCnt").Value)  '��������
            
'            Call ObjLISComCode.LisItem.KeyChange(tmpData(12))
            tmpData(18) = GetLabDiv(tmpData(12)) 'ObjLISComCode.LisItem.Fields("labdiv")    'LabDiv
            
            Call GetSpcInfo(tmpData(2), tmpData(19), tmpData(20))
'            Call ObjLISComCode.LisSpc.KeyChange(tmpData(2))
'            tmpData(19) = ObjLISComCode.LisSpc.Fields("spcbarnm")    '��ü����
'            tmpData(20) = ObjLISComCode.LisSpc.Fields("labrange")   '�̻���������ȣ����
            
            Call objLISCollect.SetAddLabCollect(tmpData)
            .MoveNext
        Next
    End With

    ' ä�� ����
    
    If Rs.RecordCount > 0 Then
        blnSuccess = objLISCollect.DoCollection(pbrPtCnt)
        blnCollectFg = True
    Else
        GoTo Skip
    End If

Err_Trap:
'    Dim objBld As clsBasisData
    Dim strBld As String
    
'    Set objBld = New clsBasisData
    strBld = GetBuildNm(ObjSysInfo.BuildingCd)
'    Set objBld = Nothing
    
    If Not blnSuccess Then
        tblPtList.Row = Row
        tblPtList.Col = -1
        tblPtList.ForeColor = vbRed       '������
        intErrCount = intErrCount + 1
    Else
         '* Delivery Location �� Count
         For i = 1 To objLISCollect.ColCount
            Call objLISCollect.GetLabNumbers(i, sWorkarea, sAccdt, iAccseq, sBuildCd)
'            Call ObjLISComCode.Building.KeyChange(sBuildCd)
           
            lngBldRow = 0
            For j = 1 To tblCount.DataRowCnt
                tblCount.Row = j: tblCount.Col = 3
                If tblCount.Value = sBuildCd Then
                    lngBldRow = j
                    Exit For
                End If
            Next

            If lngBldRow = 0 Then lngBldRow = tblCount.DataRowCnt + 1
            tblCount.Row = lngBldRow
            tblCount.Col = 1: tblCount.Text = strBld 'ObjLISComCode.Building.Fields("buildnm")
            tblCount.Col = 2: tblCount.Text = Val(tblCount.Text) + 1
            tblCount.Col = 3: tblCount.Text = ObjSysInfo.BuildingCd
        Next

    End If
Skip:
    Set Rs = Nothing

End Sub

Private Function GetLabDiv(ByVal vTestCd As String) As String
    Dim Rs As Recordset
    Dim strSQL As String
    
    strSQL = " select a.testcd,a.applydt,b.field2 from " & T_LAB001 & " a, " & T_LAB032 & " b"
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
    vSpcAbbr = Rs.Fields("spcbarnm").Value & ""
    vLabRng = Rs.Fields("labrange").Value & ""
    End If
    Set Rs = Nothing
End Sub

'% �������� ���� �Կ����� ȯ�ڵ��� ó���� �˻��Ѵ�.
Private Sub cmdGetOrders_Click()
    Dim Rs          As Recordset
    Dim Resp        As VbMsgBoxResult
    Dim objProgress As clsProgress
    Dim SqlStmt     As String
    Dim tmpDate     As String
    Dim tmpTime     As String

    Dim i           As Integer

    If Trim(txtWardID.Text) = "" Then
        MsgBox "����ID�� �Է��ϼ���.", vbInformation, "��������"
        txtWardID.SetFocus
        Exit Sub
    End If
    
    Call ClearRtn(2): txtPtId.Text = ""
    fraWard.ZOrder 0
    fraQuery.ZOrder 0
    cmdSave.Enabled = True
    cmdSaveNurse.Enabled = False

    Set objLISCollect = New clsLISCollectioin
    
    If Not objLISCollect.Archive_WardColData(txtWardID.Text) Then
        MsgBox "�����ϰ�ä�� ���� Archive�� ������ �߻��߽��ϴ�." & vbCrLf & _
                "����� Ȥ�� �ӻ󺴸����� �����ٶ��ϴ�. (��" & ObjSysInfo.HelpLine & ")", vbCritical, "�����߻�"
    End If

    Set objLISCollect = Nothing
    '---------------------------------------------------------------------------------------------
    
    If ChkMornFg.Value = 1 Then
        Resp = MsgBox("�ӻ󺴸� ��ħä�� �۾��� �����Ͻðڽ��ϱ�?", vbQuestion + vbYesNo, "��ħä��")
        If Resp = vbNo Then Exit Sub
    End If
    
    Call TableClear(1)

    tmpDate = Format(dtpToTime.Value, CS_DateDbFormat)
    tmpTime = Format(dtpToTime.Value, CS_TimeDbFormat)

    Call MouseRunning
    Set objProgress = New clsProgress
    With objProgress
        .Container = MainFrm.stsbar
        .Message = Trim(txtWardID.Text) & " ���� ȯ�ڵ��� ó���� �˻����Դϴ�.."
'        .Caption = "�����ϰ�ä��"
'        .Msg = Trim(txtWardID.Text) & " ���� ȯ�ڵ��� ó���� �˻����Դϴ�.."
'        .Mode = 1
    End With

    If ChkMornFg.Value = 1 Then
        SqlStmt = objMySql.SqlOrderForMornCol(tmpDate, tmpTime, txtWardID.Text)
    Else
        SqlStmt = objMySql.SqlWardOrder(tmpDate, tmpTime, txtWardID.Text)
    End If
    
    Set Rs = New Recordset
    Rs.Open SqlStmt, DBConn
    
    If Rs.EOF Then
'        MsgBox "ó�� �˻��� ������ �߻��߽��ϴ�. " & _
'               "����� Ȥ�� �ӻ󺴸����� �����ٶ��ϴ�.", vbExclamation, "�����߻�"
        GoTo Errors
    End If

    If Rs.EOF Then
        If BBSAddSpecimenChk(Format(GetSystemDate, "yyyymmdd"), Trim(txtWardID.Text)) = False Then
            MsgBox "ä������ �����ϴ�..", vbInformation, "����ä��"
            cmdSave.Enabled = False
            GoTo Errors
        End If
    Else
        Call DisplayOrders(Rs, objProgress)
    End If

    'ó�泻�� Display
    cmdSave.Enabled = True
    blnCleanFg = False

    DoEvents

    tblPtList.SetFocus

Errors:
    Set Rs = Nothing
    Set objProgress = Nothing

    MouseDefault

End Sub

Private Sub DisplayOrders(ByVal objRs As Recordset, Optional ByRef objPrgBar As Object = Nothing)

    Dim objGetSql   As clsBBSCollection

    Dim tmpPtId     As String
    Dim tmpStatFg   As String
    Dim tmpSpcCd    As String
    Dim tmpOrdDiv   As String
    Dim i           As Long
    
    Set objGetSql = New clsBBSCollection
    
    With tblPtList
        '���α׷����� ó��..
        If Not objPrgBar Is Nothing Then
            objPrgBar.Min = 0
            objPrgBar.Max = objRs.RecordCount * 100 + 1
            objPrgBar.Value = 0
            objPrgBar.Visible = True
            DoEvents
        End If

        .MaxRows = 0
        .MaxRows = IIf(objRs.RecordCount < 29, 29, objRs.RecordCount)
        .Row = 1

        intPtCount = 0

        For i = 1 To objRs.RecordCount

            If tmpPtId <> Trim(objRs.Fields("PtId")) Then

                If Not objPrgBar Is Nothing Then objPrgBar.Value = objPrgBar.Value + 50
                DoEvents

                intPtCount = intPtCount + 1
                .Row = intPtCount
                .Col = 2: .Text = "" & objRs.Fields("WardId").Value   '����ID
                .Col = 3: .Text = "" & objRs.Fields("PtId").Value     'ȯ��ID
                .Col = 4: .Text = "" & objRs.Fields("PtNm").Value     '����
                .Col = 7: .Text = "" & objRs.Fields("DOB").Value      '�������
                .Col = 8: .Text = "" & objRs.Fields("BedInDt").Value  '�Կ���
                .Col = 14:
                .Text = Trim("" & objRs.Fields("sex").Value)
                If IsNumeric(.Text) Then
                    .Text = Choose((Val(.Text) Mod 2) + 1, "F", "M")
                End If
                tmpPtId = "" & objRs.Fields("PtId").Value
            End If

            .Col = 9:   .Text = "" & objRs.Fields("DeptCd").Value       '�����
            .Col = 10:  .Text = "" & objRs.Fields("OrdDoct").Value     'ó����
            .Col = 11:  .Text = "" & objRs.Fields("MajDoct").Value     '��ġ��
            .Col = 12:  .Text = "" & objRs.Fields("HosilId").Value     '����ID
            .Col = 13:  .Text = "" & objRs.Fields("RoomId").Value      'ȣ��ID

            tmpStatFg = "" & objRs.Fields("StatFg").Value             '���޿���
            tmpOrdDiv = "" & objRs.Fields("orddiv").Value             'ó�汸��
            tmpSpcCd = "" & objRs.Fields("SpcCd").Value               '��ü
            
            If tmpOrdDiv = BBS_ORDDIV Then .Col = 23: .Value = tmpStatFg
            
            If chkTestdiv.Value = 1 Then                              '�˻��ڵ�� ���
                If tmpSpcCd = "" And tmpOrdDiv = BBS_ORDDIV Then tmpSpcCd = BBSName
            Else                                                      '�˻������ ���
                If tmpOrdDiv = LIS_ORDDIV Then
                    Dim tmpSpcNm As String
                    Dim tmpLabRng As String
                    
                    Call GetSpcInfo(tmpSpcCd, tmpSpcNm, tmpLabRng)
                    
                    If tmpSpcNm <> "" Then
                        tmpSpcCd = tmpSpcNm
                    Else
                        tmpSpcCd = objMySql.Get_SpcNm(tmpSpcCd, tmpOrdDiv)
                    End If
                
'                    If ObjLISComCode.LisSpc.Exists(tmpSpcCd) Then
'                        ObjLISComCode.LisSpc.KeyChange (tmpSpcCd)
'                        tmpSpcCd = ObjLISComCode.LisSpc.Fields("spcbarnm")
'                    Else
'                        tmpSpcCd = objMySql.Get_SpcNm(tmpSpcCd, tmpOrdDiv)
'                    End If
                Else
                    tmpSpcCd = objMySql.Get_SpcNm(tmpSpcCd, tmpOrdDiv)
                End If
                
                If tmpSpcCd = "" And tmpOrdDiv = BBS_ORDDIV Then tmpSpcCd = BBSName
                
            End If
            If tmpStatFg = "1" Then     '���ް�ü
                .Col = 5
                If InStr(1, .Text, tmpSpcCd) = 0 Then
                    .Text = .Text & tmpSpcCd & ", "
                End If
                '��ħä������
                .Col = 22: .Text = "0"
            Else
                .Col = 6
                If InStr(1, .Text, tmpSpcCd) = 0 Then
                    .Text = .Text & tmpSpcCd & ", "
                End If
                If ChkMornFg.Value = 1 Then
                    .Col = 22: .Text = "1"
                Else
                    .Col = 22: .Text = "0"
                End If
            End If

            Select Case tmpOrdDiv
            Case LIS_ORDDIV:   '�ӻ�
                .Col = 15: .ForeColor = vbRed: .Text = "��"     'ó�汸�С��
            Case BBS_ORDDIV:   '����
                .Col = 17: .ForeColor = vbRed: .Text = "��"     'ó�汸�С��
                If objGetSql.Blood_Existence(tmpPtId, Format(GetSystemDate, CS_DateDbFormat), _
                                            Format(GetSystemDate, "HHmm")) = True Then
                    .Col = 18: .ForeColor = vbBlue: .Value = "�ű�"
                Else
                    .Col = 18: .ForeColor = DCM_Gray: .Value = "����"
                End If

            End Select

            .Col = 19: .Text = Format(dtpColDtTm.Value, "YY-MM-DD")
            .Col = 20: .Text = Format(dtpColDtTm.Value, "HH:MM")

            objRs.MoveNext
        Next

        If Not objPrgBar Is Nothing Then objPrgBar.Value = objPrgBar.Max
        DoEvents

        pbrPtCnt.Min = 0
        pbrPtCnt.Max = .DataRowCnt * 10
        pbrPtCnt.Value = 0

        dtpColDtTm.Value = GetSystemDate

    End With

    Set objGetSql = Nothing

End Sub

Private Function BBSAddSpecimenChk(ByVal OrdDt As String, ByVal qWardId As String) As Boolean
'���������� ä��������߿� ��ü �߰� ����ڰ� ���ԵǾ� �ִ��� �Ǵ��ؼ� �����ش�.
'��ü �߰� ����ڴ� �̹� ������ ȯ�ڸ� �������� �ҷ��´�.
'�߰���û���� ������ ���� ��¥�� �������� �۰ų� ���� �͸��� ������� �Ѵ�.

    Dim objGetSql   As clsBBSCollection
    Dim Rs          As Recordset
    Dim strErChk    As String
    Dim strPtid     As String
    Dim cnt         As Integer

    BBSAddSpecimenChk = True

    Set objGetSql = New clsBBSCollection
    
    Set Rs = objGetSql.Get_SpcAdd(UCase(qWardId))

    
    If Not Rs.EOF Then
        With tblPtList
            Do Until Rs.EOF
                If DupCheck("" & Rs.Fields("ptid").Value) = False Then
                    If .DataRowCnt <= .MaxRows Then .MaxRows = .MaxRows + 1
                    .Row = .DataRowCnt + 1
                    .ForeColor = vbBlue
                    .Col = 2:   .Value = qWardId
                    .Col = 3:   .Value = "" & Rs.Fields("ptid").Value
                                strPtid = Trim("" & Rs.Fields("ptid").Value)
                    .Col = 4:   .Value = "" & Rs.Fields("ptnm").Value
                                strErChk = objGetSql.ER_Chk(strPtid, "" & Rs.Fields("orddt").Value)
                    .Col = 5:   .Value = IIf(strErChk = "1", "��", "")
                    .Col = 6:   .Value = IIf(strErChk = "0", "��", "")
                    .Col = 7:   .Value = "" & Rs.Fields("dob").Value
                    .Col = 8:   .Value = "" & Rs.Fields("bedindt").Value
                    .Col = 14:  .Text = Choose((Val("" & Rs.Fields("Sex")) Mod 2) + 1, "F", "M") '����
                    Select Case "" & Rs.Fields("orddiv").Value
                        Case "L":   '�ӻ�
                            .Col = 15: .ForeColor = vbRed: .Text = "��"     'ó�汸�С��
                        Case "B":   '����
                            .Col = 17: .ForeColor = vbRed: .Text = "��"     'ó�汸�С��
                    End Select
                    .Col = 18:  .Value = "�߰�"
                    .Col = 19:  .Text = Format(dtpColDtTm.Value, "YY-MM-DD")
                    .Col = 20:  .Value = Format(dtpColDtTm.Value, "HH:MM:SS")

                    .Col = 9:   .Text = "" & Rs.Fields("DeptCd").Value       '�����
                    .Col = 10:  .Text = "" & Rs.Fields("OrdDoct").Value     'ó����
                    .Col = 11:  .Text = "" & Rs.Fields("MajDoct").Value     '��ġ��
                    .Col = 12:  .Text = "" & Rs.Fields("RoomId").Value      '����ID
                    .Col = 13:  .Text = "" & Rs.Fields("HosilId").Value     'ȣ��ID
                    cnt = cnt + 1
                Else
                    '�߰�ä����, �Ϲ�ä���� ���ÿ� �߻��Ѱ��
                    .Col = 21:  .Value = "*"
                End If
                Rs.MoveNext
            Loop
        End With
    Else
        BBSAddSpecimenChk = False
    End If

    If cnt = 0 Then BBSAddSpecimenChk = False

    Set Rs = Nothing
    Set objGetSql = Nothing

End Function

Private Function DupCheck(ByVal pBldNo As String) As Boolean
'�ߺ����� üũ�Ѵ�.

    Dim strClip As String

    With tblPtList
        .Row = 1: .Row2 = .MaxRows
        .Col = 3: .Col2 = 3
        .BlockMode = True
        strClip = .ClipValue
        .BlockMode = False

        If InStr(strClip, pBldNo) Then
            DupCheck = True
        Else
            DupCheck = False
        End If
    End With

End Function


' ���ؽð��� ����Ǹ� Clear
Private Sub dtpToTime_Change()

    If Not blnCleanFg Then Call TableClear(1)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call ICSPatientMark
    Set objMySql = Nothing
    Set objLISCollect = Nothing
    Set MyPatient = Nothing
    Set objMyList = Nothing
End Sub

Private Sub optApplyColTm_Click(Index As Integer)

    Dim Resp As VbMsgBoxResult

    If dtpColDtTm.Value < Now Then
        If dtpColDtTm.Tag = "1" Then
            dtpColDtTm.Tag = "0"
        Else
            Resp = MsgBox("ä��ð��� ����ð����� �����Դϴ�. �����Ͻðڽ��ϱ�?", _
                   vbQuestion + vbYesNo, "ä��ð�����")
            If Resp = vbYes Then
                dtpColDtTm.Tag = "1"
            Else
                dtpColDtTm.Tag = "0"
                dtpColDtTm.Value = Format(GetSystemDate, "YY-MM-DD HH:MM")
            End If
        End If
    End If

    With tblPtList
        If optApplyColTm(0).Value Then  '��ü
            .Row = 1: .Row2 = .DataRowCnt
            .Col = 19: .Col2 = 19
            .BlockMode = True
            .Text = Format(dtpColDtTm.Value, "YY-MM-DD")
            .BlockMode = False
            .Col = 20: .Col2 = 20
            .BlockMode = True
            .Text = Format(dtpColDtTm.Value, "HH:MM")
            .BlockMode = False
        Else
            If .ActiveRow > .DataRowCnt Then Exit Sub
            .Row = .ActiveRow
            .Col = 19: .Text = Format(dtpColDtTm.Value, "YY-MM-DD")
            .Col = 20: .Text = Format(dtpColDtTm.Value, "HH:MM")
            optApplyColTm(1).Value = False
        End If
    End With

End Sub

Private Sub optOption_Click(Index As Integer)

    Select Case Index
    Case 0, 2: txtCopy.Text = 1
                txtCopy.Enabled = True
    Case 1: txtCopy.Text = 0
                txtCopy.Enabled = False
    End Select

End Sub

Private Sub cmdWardList_Click()
'% �����ڵ� ����Ʈ�� �˾��Ѵ�.
'    Dim objWard As clsBasisData
    
    Set objMyList = New clsPopUpList
'    Set objWard = New clsBasisData
    
    txtWardID.Text = "": lblWardNm.Caption = ""
    With objMyList
        .Connection = DBConn
        .FormCaption = "���� ��ȸ"
        .ColumnHeaderText = "�����ڵ�;������"
        Call .LoadPopUp(GetSQLWardList) ', 2700, Frame2.Left + cmdWardList.Left) ', ObjLISComCode.WardId)
        If .SelectedString <> "" Then
            txtWardID.Text = medGetP(.SelectedString, 1, ";")
            lblWardNm.Caption = medGetP(.SelectedString, 2, ";")
        End If
    End With
'    Set objWard = Nothing
    Set objMyList = Nothing
End Sub

Private Sub tblOrdSheet_DblClick(ByVal Col As Long, ByVal Row As Long)
    cmdSave.Enabled = True
    cmdSaveNurse.Enabled = False
    fraWard.ZOrder 0
    fraQuery.ZOrder 0
End Sub

Private Sub tblPtList_Click(ByVal Col As Long, ByVal Row As Long)
    If Row < 1 Then Exit Sub
    If Col < 2 Then Exit Sub
    tblPtList.Row = Row
    tblPtList.Col = 3
    If tblPtList.Value = "" Then Exit Sub
    
    txtPtId.Text = tblPtList.Value
    cmdSave.Enabled = False
    Call GetPtCollection(txtPtId.Text)
    Call DisplayOrder
    fraNurse.ZOrder 0
End Sub
Private Sub tblPtList_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
    Dim tmpRs       As Recordset
    Dim tmpToolTip  As String
    
    Dim strSQL      As String
    Dim strPtid     As String
    Dim strOrdDate  As String
    Dim strOrdDiv   As String
    Dim strWardId   As String
    Dim strBBSORDCd As String
    Dim strLISORDCd As String

    If Row = 0 Then Exit Sub

    tmpToolTip = vbCrLf

    With tblPtList
        .Row = Row

        .Col = 2: If Trim(.Value) = "" Then Exit Sub

        .Col = 4: tmpToolTip = tmpToolTip & "  " & .Value & vbCrLf & vbCrLf    'ȯ�ڸ�
        .Col = 5: tmpToolTip = tmpToolTip & "  ���ް�ü : " & .Value & vbCrLf  '���ް�ü
        .Col = 6: tmpToolTip = tmpToolTip & "  �Ϲݰ�ü : " & .Value & vbCrLf  '�Ϲݰ�ü
        
        '-- ToolTip �߰����� : �˻��׸� Display
        ' - ȯ��ID
        .Col = 3: strPtid = Trim(.Value)
        strOrdDate = Format(dtpToTime.Value, CS_DateDbFormat)
        strWardId = Trim(txtWardID.Text)
        
        strSQL = objMySql.WardMn_ORDCD(strPtid, strOrdDate, strWardId)
        
        Set tmpRs = New Recordset
        tmpRs.Open strSQL, DBConn
        
        If tmpRs.BOF = False Then
            Do Until tmpRs.EOF = True
                strOrdDiv = Trim(tmpRs.Fields("orddiv").Value & "")
                
                Select Case strOrdDiv
                    Case "B"
                        strBBSORDCd = strBBSORDCd & tmpRs.Fields("abbrnm5").Value & "," '�������� �˻��׸�
                        
                    Case "L"
                        strLISORDCd = strLISORDCd & tmpRs.Fields("abbrnm5").Value & "," '�ӻ󺴸� �˻��׸�
                        
                End Select
                
                tmpRs.MoveNext
            Loop
        End If
        
        If strBBSORDCd <> "" Then
                tmpToolTip = tmpToolTip & "  �������� : " & strBBSORDCd & vbCrLf  '�������� �˻��׸�
        ElseIf strLISORDCd <> "" Then
                tmpToolTip = tmpToolTip & "  �ӻ󺴸� : " & strLISORDCd & vbCrLf  '�ӻ󺴸� �˻��׸�
        End If
        
        MultiLine = 1
        TipText = tmpToolTip
        TipWidth = 5000
        .TextTipDelay = 1000
        Call .SetTextTipAppearance("����ü", 9, False, False, &HEEFDF2, &H996666)
        ShowTip = True
    End With
    
    Set tmpRs = Nothing
End Sub

'% ��� ������ ����Ǹ� Clear
Private Sub txtWardId_Change()
    If Not blnCleanFg Then Call TableClear(1)
End Sub

Private Sub ClearRtn(ByVal intOpt As Integer)
    Select Case intOpt
        Case 1
            txtWardID.Enabled = True
            txtWardID.BackColor = &H80000005
            cmdWardList.Enabled = True
            dtpToTime.Enabled = True
            cmdGetOrders.Enabled = True
            cmdSave.Enabled = False
        
            sWorkDt = "": sWorkTm = ""
            txtWardID.Text = ""
            lblWardNm.Caption = ""
            dtpToTime.Value = Format(GetSystemDate, "YYYY-MM-DD hh:mm:ss")
            dtpColDtTm.Value = GetSystemDate
            dtpColDtTm.Tag = "0"
            pbrPtCnt.Value = 0
            chkPrintFg = 0
            optOption(1).Value = True
            optApplyColTm(0).Value = True
            intErrCount = 0
            Call TableClear(intOpt)
        Case Else
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
            
             With tblOrdSheet
                 .Row = -1
                 .Col = -1
                 .BlockMode = True
                 .Action = ActionClearText
                 .BlockMode = False
             End With
             
             cmdSaveNurse.Enabled = True
    End Select
End Sub


'% Table���� Clear�Ѵ�
Private Sub TableClear(ByVal intOpt As Integer)
    tblPtList.MaxRows = 0
    tblPtList.MaxRows = 50
    If intOpt = 1 Then
        lblColNm.Caption = ""
        lblPtCount.Caption = ""
        tblCount.MaxRows = 0
        tblCount.MaxRows = 50
        blnCleanFg = True
    End If
End Sub

'% ���� ID
Private Sub txtWardId_GotFocus()

    With txtWardID
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub txtWardId_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If objMyList Is Nothing Then Call cmdWardList_Click
    End If
End Sub


Private Sub txtWardId_KeyPress(KeyAscii As Integer)

    On Error GoTo Err_Trap

    KeyAscii = Asc(UCase(Chr(KeyAscii)))

    If KeyAscii = vbKeyReturn Then
        If txtWardID.Text = "" Then
            lblWardNm.Caption = ""
            Exit Sub
        Else
'            Dim objWard As clsBasisData
            Dim Rs As Recordset
            Dim strWard As String
            
'            Set objWard = New clsBasisData
            Set Rs = New Recordset
            
            strWard = GetSQLWard(txtWardID.Text)
            
            Rs.Open strWard, DBConn
            
            If Rs.EOF = False Then
                ObjSysInfo.BuildingCd = Rs.Fields("bldgb").Value & ""
                ObjSysInfo.BuildingNm = Rs.Fields("bldnm").Value & ""
                ObjSysInfo.BuildingNo = Rs.Fields("bldno").Value & ""
                txtWardID.Tag = txtWardID.Text
            Else
                MsgBox "���� �ڵ带 Ȯ���ϼ���.", vbInformation
                txtWardID.Text = ""
                lblWardNm.Caption = ""
                txtWardID.SetFocus
                Call txtWardId_KeyDown(vbKeyDown, 0)
            End If
            Set Rs = Nothing
'            Set objWard = Nothing
            
'            With ObjLISComCode.WardId
'                If .Exists(txtWardID.Text) Then
'                    Call .KeyChange(txtWardID.Text)
'                    lblWardNm.Caption = .Fields("WardNm")
'                    objsysinfo.BuildingCd = .Tags("bldgb")
'                    objsysinfo.BuildingNm = .Tags("bldnm")
'                    objsysinfo.BuildingNo = .Tags("bldno")
'                    dtpToTime.SetFocus
'                Else
'                    MsgBox "���� �ڵ带 Ȯ���ϼ���..", vbInformation, "�ڵ��Է¿���"
'                    txtWardID.Text = ""
'                    lblWardNm.Caption = ""
'                    txtWardID.SetFocus
'                    Call txtWardId_KeyDown(vbKeyDown, 0)
'                    Exit Sub
'                End If
'            End With
        End If
    End If
    Exit Sub

Err_Trap:
    Resume Next

End Sub

Private Sub PrintColList(ByVal pWardId As String, ByVal pWardNm As String, _
                         ByVal pWorkDt As String, ByVal pWorkTm As String, _
                         ByVal pBuildCd As String, ByVal pBuildNm As String)

    Dim MyReport    As clsWardColList
    Dim strTitleNm  As String
    
    Set MyReport = New clsWardColList
    strTitleNm = IIf(ChkMornFg.Value = 0, "���� ä�� ����Ʈ", "������ ��ħä�� ����Ʈ")

    With MyReport
        .WardId = pWardId
        .WardNm = pWardNm
        .WorkDt = pWorkDt
        .WorkTm = pWorkTm
        .BuildCd = pBuildCd
        .BuildNm = pBuildNm
        .TestDiv = chkTestdiv.Value
        .TitleNm = strTitleNm
        .SetCrpt CReport
        Call .Print_ColList
    End With

    Set MyReport = Nothing

End Sub


Public Sub Call_WardId_KeyPress()

    Call txtWardId_KeyPress(vbKeyReturn)

End Sub

Public Sub Call_cmdGetOrders_click()

     Call cmdGetOrders_Click

End Sub

Private Sub txtWardId_LostFocus()

On Error GoTo Err_Trap

    If ActiveControl.Name = cmdWardList.Name Then Exit Sub
    If ActiveControl.Name = cmdClear.Name Then Exit Sub
    If ActiveControl.Name = cmdExit.Name Then Exit Sub
    If txtWardID.Text = "" Then
        lblWardNm.Caption = ""
        Exit Sub
    Else
        Call txtWardId_KeyPress(vbKeyReturn)
    End If
    
    Exit Sub
Err_Trap:
    Resume Next

End Sub
'=======================================================================================
'���������� ����ä���κ���.
'=======================================================================================
'% ȯ��ID�� ����Ǹ� ȭ��Clear
Private Sub GetPtCollection(ByVal sPtid As String)
    If IsNumeric(txtPtId.Text) Then txtPtId.Text = Format(txtPtId.Text, P_PatientIdFormat)
       
    If Not blnCleared Then Call ClearRtn(2)
    DoEvents
    
    Call ICSPatientMark(txtPtId.Text, enICSNum.LIS_ALL)
    
'    Call MyPatient.ClearData   'Ŭ���� �� ���� �ʱ�ȭ
    If MyPatient.GETPatient(txtPtId.Text) Then
        lblPtNm.Caption = MyPatient.PtNm     '����
        lblSexAge.Caption = MyPatient.SEXNM & " / " & MyPatient.Age & " " & MyPatient.AGEDIV      '����
        lblDeptNm.Caption = MyPatient.DeptNm '�����
        lblLocation.Caption = MyPatient.WardId & "-" & MyPatient.RoomId & "-" & MyPatient.BedID   '����
        DoEvents
        PtFg = True
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
    If Not OrdFg Then
        cmdSave.Enabled = False
        txtPtId.SetFocus
        Call txtPtId_GotFocus
    End If
End Sub

Private Sub txtPtId_Change()
    If Not blnCleared Then
       Call ClearRtn(2)
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
    
    If KeyAscii = vbKeyReturn Then
        Call GetPtCollection(txtPtId.Text)
        Call DisplayOrder
    End If
End Sub

'% �˻��� ó���� ���̺� ���÷��� �Ѵ�.
Private Sub DisplayOrder()
    Dim Rs          As Recordset
    Dim objProInSts As jProgressBar.clsProgress
    Dim objGetSql   As clsBBSCollection
    Dim i           As Integer
    Dim SqlStmt     As String
    
    Dim SvOrdDt     As String
    Dim SvOrdNo     As String
    Dim SvSpcNm     As String
    Dim SvOrdDoct   As String
    Dim tmpDate     As String
    Dim tmpTime     As String

    Dim strErChk    As String
    Dim strOrdDiv   As String
   
On Error GoTo NoData

   
    Set objProInSts = New jProgressBar.clsProgress
    With objProInSts
        .Container = Me
        .Left = fraNurse.Left + lblBar.Left
        .Top = lblBar.Top - 70
        .Width = lblBar.Width
        .Height = lblBar.Height
        .Message = "�ش�ȯ���� ó�� ������ �˻� ���Դϴ�..."
'        .SetMyForm Me
'        .Choice = True
'        .XPos = fraNurse.Left + lblBar.Left
'        .YPos = lblBar.Top - 70
'        .XWidth = lblBar.Width
'        .ForeColor = &HFA8B10
'        .Appearance = aPlate
'        .BorderStyle = bsNone
'        .YHeight = lblBar.Height '
'        .Msg = "�ش�ȯ���� ó�� ������ �˻� ���Դϴ�...."
'        .Max = 90
'        .Min = 0
'        .Value = 10
        DoEvents
    End With

    DoEvents
    txtMesg.Text = ""
    
    ' ó�泻�� �˻�

    tmpDate = Format(dtpToTime.Value, CS_DateDbFormat)
    tmpTime = Format(dtpToTime.Value, CS_TimeDbFormat)
    

    strOrdDiv = "W"
    SqlStmt = objMySql.SqlReadWardOrder(txtPtId.Text, tmpDate, tmpTime, , enBussDiv.BussDiv_InPatient, , strOrdDiv)
    Set Rs = New Recordset
    Rs.Open SqlStmt, DBConn
    
    If Rs.EOF Then
        Set Rs = Nothing
        Set objProInSts = Nothing
       
        MsgBox MyPatient.PtNm & " ���� ó�泻���� �����ϴ�", vbInformation, "��ȣ�� ä��"
        If Not blnCleared Then Call ClearRtn(2): txtPtId.Text = ""
        Exit Sub
    End If
    
'    Dim objBld As clsBasisData
    Dim strBld As String
    
'    Set objBld = New clsBasisData
    strBld = GetBuildNm(ObjSysInfo.BuildingCd)
'    Set objBld = Nothing
    
    With tblOrdSheet
       
        .ReDraw = False
        .MaxRows = 0
        If Rs.RecordCount < lngMaxRows Then
            .MaxRows = lngMaxRows
            .Row = Rs.RecordCount + 1
            .Row2 = lngMaxRows
            .Col = 1: .Col2 = .MaxCols
            .BlockMode = True
            .Lock = True
            .Protect = True
            .BlockMode = False
        Else
            .MaxRows = Rs.RecordCount   '����Ÿ �Ǽ�
        End If
       
        objProInSts.Max = Rs.RecordCount

        .Row = -1
        .Col = 2: .Col2 = .MaxCols
        .BlockMode = True
        .Lock = True
        .Protect = True
        .BlockMode = False
: SelAllFg = True
        For i = 1 To Rs.RecordCount

            objProInSts.Value = i
            
            .Row = i
            .Col = enCOLLIST.tcCHECK: .Value = 1
            If SvOrdDt <> Trim("" & Rs.Fields("OrdDt").Value) Then
                .Col = enCOLLIST.tcORDDT:   .Text = Format("" & Rs.Fields("OrdDt").Value, CS_DateShortMask)    'ó����
                .Col = enCOLLIST.tcORDNO:   .Text = Trim("" & Rs.Fields("OrdNo").Value)     'ó���ȣ
                .Col = enCOLLIST.tcSPCNM:   .Text = Trim("" & Rs.Fields("SpcNm").Value)     '��ü
                .Col = enCOLLIST.tcDOCTNM:  .Text = Trim("" & Rs.Fields("DoctNm").Value)     'ó����
                SvOrdDt = Trim("" & Rs.Fields("OrdDt").Value)
                SvOrdNo = Trim("" & Rs.Fields("OrdNo").Value)    'ó���ȣ
                SvSpcNm = Trim("" & Rs.Fields("SpcNm").Value)    '��ü
                SvOrdDoct = Trim("" & Rs.Fields("DoctNm").Value) 'ó����
            End If
            If SvOrdNo <> Trim("" & Rs.Fields("OrdNo").Value) Then
                .Col = enCOLLIST.tcORDNO:   .Text = Trim("" & Rs.Fields("OrdNo").Value)     'ó���ȣ
                .Col = enCOLLIST.tcSPCNM:   .Text = Trim("" & Rs.Fields("SpcNm").Value)     '��ü
                .Col = enCOLLIST.tcDOCTNM:  .Text = Trim("" & Rs.Fields("DoctNm").Value)    'ó����
                SvOrdNo = Trim("" & Rs.Fields("OrdNo").Value)    'ó���ȣ
                SvSpcNm = Trim("" & Rs.Fields("SpcNm").Value)    '��ü
                SvOrdDoct = Trim("" & Rs.Fields("DoctNm").Value) 'ó����
            End If
            If SvSpcNm <> Trim("" & Rs.Fields("SpcNm").Value) Then
                .Col = enCOLLIST.tcSPCNM:   .Text = Trim("" & Rs.Fields("SpcNm").Value)     '��ü
                SvSpcNm = Trim("" & Rs.Fields("SpcNm").Value)
            End If
            If SvOrdDoct <> Trim("" & Rs.Fields("DoctNm").Value) Then
                .Col = enCOLLIST.tcDOCTNM: .Text = Trim("" & Rs.Fields("DoctNm").Value)    'ó����
                SvOrdDoct = Trim("" & Rs.Fields("DoctNm").Value)
            End If

            Select Case Rs.Fields("orddiv")
                Case BBS_ORDDIV:
                    Set objGetSql = New clsBBSCollection
                    strErChk = objGetSql.ER_Chk(txtPtId.Text, SvOrdDt)
                    .Col = enCOLLIST.tcSTATFG:  .Value = Trim("" & Rs.Fields("StatFg").Value)     '���޿���  --> ������ ó��...
                    .Col = enCOLLIST.tcBUILDCD: .Value = ObjSysInfo.BuildingCd 'IIf(strErChk = "1", objsysinfo.BuildingCd, objsysinfo.BuildingCd)
'                    If ObjLISComCode.Building.Exists(.Value) Then
'                        ObjLISComCode.Building.KeyChange (.Value)
'                    End If
                    .Col = enCOLLIST.tcBUILDNM: .Value = strBld ' ObjLISComCode.Building.Fields("buildnm")
                    Set objGetSql = Nothing
                Case LIS_ORDDIV:
                    .Col = enCOLLIST.tcBUILDCD:  .Text = ObjSysInfo.BuildingCd
                    .Col = enCOLLIST.tcBUILDNM:  .Text = ObjSysInfo.BuildingNm
                    .Col = enCOLLIST.tcSTATFLAG: .Text = Trim(Rs.Fields("StatFg").Value)
            End Select
            .Col = enCOLLIST.tcTESTNM:  .Text = Trim("" & Rs.Fields("TestNm").Value)     'ó���
                    Select Case Rs.Fields("orddiv")
                        Case BBS_ORDDIV: .ForeColor = &H496835     '&H6C6181     '&H81815A     '�ణ���   &H00845584&�����
                        Case LIS_ORDDIV: .ForeColor = &H553755
                    End Select
            
            .Col = enCOLLIST.tcSTATFG:  .Text = IIf("" & Rs.Fields("StatFg").Value = "0", "", "Y") '���޿���
                                        .ForeColor = DCM_Red                                '������
            .Col = enCOLLIST.tcREQDTTM: .Text = Format("" & Rs.Fields("ReqDt").Value, CS_DateMask) & " " & _
                                         Format("" & Rs.Fields("ReqTm").Value, CS_TimeLongMask)      '���ä���Ͻ�
            .Col = enCOLLIST.tcORDDATE: .Text = Trim("" & Rs.Fields("OrdDt").Value)      'ó����
            .Col = enCOLLIST.tcORDNUM:  .Text = Trim("" & Rs.Fields("OrdNo").Value)      'ó���ȣ
            .Col = enCOLLIST.tcORDSEQ:  .Text = Trim("" & Rs.Fields("OrdSeq").Value)     'ó��Seq
            .Col = enCOLLIST.tcTESTCD:  .Text = Trim("" & Rs.Fields("OrdCd").Value)      '�˻��ڵ�
            
            Dim strLabDiv As String
            strLabDiv = GetLabDiv(.Text)
            
'            Call ObjLISComCode.LisItem.KeyChange(.Text)
            .Col = enCOLLIST.tcLABDIV:  .Text = strLabDiv 'ObjLISComCode.LisItem.Fields("labdiv")      'LabDiv

            .Col = enCOLLIST.tcSPCCD:   .Text = Trim("" & Rs.Fields("SpcCd").Value)      '��ü�ڵ�
            
            Dim strSpcAbbr As String
            Dim strLabRng As String
            Call GetSpcInfo(.Text, strSpcAbbr, strLabRng)
            
'            Call ObjLISComCode.LisSpc.KeyChange(.Text)
            .Col = enCOLLIST.tcSPCABBR:  .Text = Trim("" & Rs.Fields("spcnm5").Value)         '��ü����
            .Col = enCOLLIST.tcLABRANGE: .Text = strLabRng 'ObjLISComCode.LisSpc.Fields("labrange")    '�̻���������ȣ����

            .Col = enCOLLIST.tcWORKAREA: .Text = Trim("" & Rs.Fields("WorkArea").Value)  'WorkArea
            .Col = enCOLLIST.tcSTORECD:  .Text = Trim("" & Rs.Fields("StoreCd").Value)   '�����ڵ�
            .Col = enCOLLIST.tcTESTDIV:  .Text = Trim("" & Rs.Fields("TestDiv").Value)   '�˻籸��
            .Col = enCOLLIST.tcMULTIFG:  .Text = Trim("" & Rs.Fields("MultiFg").Value)   '������ü����
            .Col = enCOLLIST.tcSPCGRP:   .Text = Trim("" & Rs.Fields("SpcGrp").Value)    '��ü��
            .Col = enCOLLIST.tcORDDOCT:  .Text = Trim("" & Rs.Fields("OrdDoct").Value)   'ó����
                                         'ó���Ǹ�
                                         If .Text <> "" And lblDoctNm.Caption = "" Then
                                            lblDoctNm.Caption = Trim("" & Rs.Fields("DoctNm").Value)
                                         End If
            .Col = enCOLLIST.tcMAJDODT:  .Text = Trim("" & Rs.Fields("MajDoct").Value)   '��ġ��
            .Col = enCOLLIST.tcDEPTCD:   .Text = Trim("" & Rs.Fields("DeptCd").Value)    '�����
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
                                                lblDeptNm.Caption = strDept 'ObjLISComCode.DeptCd.Fields("deptnm")
'                                            End If
                                         End If
            .Col = enCOLLIST.tcABBRNM:  .Text = Trim("" & Rs.Fields("AbbrNm5").Value)    '����
            .Col = enCOLLIST.tcBARCNT:  .Text = Trim("" & Rs.Fields("LabelCnt").Value)   '��������
            .Col = enCOLLIST.tcPAYDT:   .Text = Trim("" & Rs.Fields("ReceptNo").Value)   '��������ȣ
                                        .ForeColor = vbRed

            .Col = enCOLLIST.tcWARDID:  .Text = Trim("" & Rs.Fields("WardId").Value)     '����
                                        mvarWardId = .Text
            .Col = enCOLLIST.tcROOMID:  .Text = Trim("" & Rs.Fields("hosilid").Value)     '����
                                        mvarHosilID = .Text
            .Col = enCOLLIST.tcBEDID:   .Text = Trim("" & Rs.Fields("roomid").Value)      '����
                                        mvarRoomID = .Text

            .Col = enCOLLIST.tcFRZFG:   .Text = Trim("" & Rs.Fields("FzFg").Value)       '��������
            .Col = enCOLLIST.tcORDDIV:  .Text = Trim("" & Rs.Fields("OrdDiv").Value)     'ó�汸��
            
            If mvarWardId <> "" Then
                lblLocation.Caption = mvarWardId & "-" & mvarHosilID & "-" & mvarRoomID
            End If

            '����μ� Remark
            If Trim("" & Rs.Fields("Mesg").Value) <> "" Then
                txtMesg.Text = txtMesg.Text & "# " & Format(Trim("" & Rs.Fields("OrdNo").Value), "##") & " - "
                txtMesg.Text = txtMesg.Text & Trim("" & Rs.Fields("TestNm").Value) & vbCrLf
                txtMesg.Text = txtMesg.Text & Trim("" & Rs.Fields("Mesg").Value) & vbCrLf
            End If

            Rs.MoveNext
        Next

        .RowHeight(-1) = lngRowHeight
        .ReDraw = True
        
   
    End With
    OrdFg = True
    fraOrder.Enabled = True
    blnCleared = False
    Set objProInSts = Nothing
    
NoData:
    Set Rs = Nothing
   
End Sub

Private Function CollectForLIS(ByVal FRowCnt As Long, ByVal LRowCnt As Long, ByRef objProgress As Object) As Boolean
    Dim tmpData()   As String
    Dim tmpDate     As String
    Dim tmpTime     As String
    Dim tmpStatFg   As String
    Dim SqlStmt     As String

    Dim ColSuccess  As Boolean
    Dim i           As Integer
    Dim SelCount    As Integer
    Dim strTmp1     As String
    Dim strReqDt    As String
    Dim strReqtm    As String
    Dim strReqTm1   As String
    Dim strLastTm   As String
    Dim CollectCnt  As Integer
    
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
        tmpData(8) = ObjSysInfo.EmpId                                      '�Է���
        tmpData(9) = ""                                          '��������ȣ
        tmpData(10) = Format(GetSystemDate, CS_DateDbFormat) 'ä����
        tmpData(11) = ObjSysInfo.EmpId                           'ä����
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
    
    Dim i           As Long
    Dim lngColCnt   As Integer
    Dim strStatFg   As String
    
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

        
        .Col = 7: strStatFg = IIf(Trim(.Value) = "Y", "1", "0")
        
        For i = 0 To UBound(tmpTotData) - 1

            tmpRowData = Split(tmpTotData(i), vbTab)
            If objProgress.Max > objProgress.Value Then objProgress.Value = objProgress.Value + 1
            If tmpRowData(0) = 0 Then GoTo Skip       '���ÿ���
          
            lngColCnt = lngColCnt + 1
            
            '��������-----------------------------------------------------------------------------
                
                dicBBS.Clear
                dicBBS.FieldInialize "ptid", "ptnm,coldt,coltm,colid,bussdiv,buildcd,hosilid,statfg"
                dicBBS.AddNew txtPtId.Text, Join(Array(lblPtNm.Caption, ColDt, ColTm, _
                              ObjSysInfo.EmpId, enBussDiv.BussDiv_InPatient, ObjSysInfo.BuildingCd, mvarHosilID, strStatFg), COL_DIV)
Skip:
       Next
    
    End With
    
    If lngColCnt = 0 Then
        CollectForBBS = True
        Set objCollect = Nothing
        Set objBar = Nothing
        Set dicBBS = Nothing
        Exit Function
    End If
          
    objCollect.WardId = mvarWardId
    CollectForBBS = objCollect.Set_Collect(dicBBS, , objProgress)
    
    If CollectForBBS Then
        Set objBar = objCollect.BldDic
        If objBar.RecordCount > 0 Then
            BarCodePrintForBBS objBar '���ڵ� ���
        Else
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

Private Sub BarCodePrintForBBS(objDIC As clsDictionary)
    Dim objBar      As clsBarcode
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
'    Set objBAR.MyDB = dbconn
    Set objBar.TableInfo = New clsTables
    Set objBar.FieldInfo = New clsFields
    
    strW_Dept = mvarWardId
    If strW_Dept = "" Then strW_Dept = mvarDeptCd
    
    If lblLocation.Caption <> "" Then
        If lblLocation.Caption <> "--" Then strW_Dept = strW_Dept & "/" & mvarHosilID
    End If
    

    strBuildNm = BBSName
    
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

End Sub
Private Sub tblOrdSheet_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

    Dim i As Integer
    Dim ButtonValue As Variant
    Dim SvOrdDt As String
    Dim SvOrdNo As String

    If SelAllFg Then Exit Sub
    
    With tblOrdSheet
       .Row = Row
       .Col = Col:   ButtonValue = .Value
       
       If .Value = 0 Then Exit Sub
       
       .Col = 9:      SvOrdDt = .Value
       .Col = 10:     SvOrdNo = .Value
       
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
