VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MEDCONTROLS1.OCX"
Object = "{9167B9A7-D5FA-11D2-86CA-00104BD5476F}#5.0#0"; "DRCTL1.OCX"
Begin VB.Form frm404AllResult_APS 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   9105
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   14745
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00808080&
   BeginProperty Font 
      Name            =   "����ü"
      Size            =   11.25
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9105
   ScaleWidth      =   14745
   WindowState     =   2  '�ִ�ȭ
   Begin VB.PictureBox picFootNote 
      Appearance      =   0  '���
      BackColor       =   &H00EFFEFE&
      ForeColor       =   &H80000008&
      Height          =   1185
      Left            =   11100
      ScaleHeight     =   1155
      ScaleWidth      =   4995
      TabIndex        =   62
      Top             =   7905
      Width           =   5025
      Begin RichTextLib.RichTextBox txtSamCmt 
         Height          =   1080
         Left            =   30
         TabIndex        =   63
         Top             =   30
         Width           =   4905
         _ExtentX        =   8652
         _ExtentY        =   1905
         _Version        =   393217
         BackColor       =   15728382
         BorderStyle     =   0
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"Lis404_APS.frx":0000
         MouseIcon       =   "Lis404_APS.frx":0304
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
   End
   Begin FPSpread.vaSpread tblResult 
      Height          =   6165
      Left            =   11085
      TabIndex        =   64
      Top             =   1725
      Width           =   5025
      _Version        =   196608
      _ExtentX        =   8864
      _ExtentY        =   10874
      _StockProps     =   64
      AllowCellOverflow=   -1  'True
      AutoCalc        =   0   'False
      AutoClipboard   =   0   'False
      BackColorStyle  =   3
      DisplayColHeaders=   0   'False
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
      GridShowHoriz   =   0   'False
      GridShowVert    =   0   'False
      GridSolid       =   0   'False
      MaxCols         =   11
      OperationMode   =   1
      Protect         =   0   'False
      ScrollBars      =   2
      ShadowColor     =   12632256
      ShadowDark      =   12632256
      ShadowText      =   0
      SpreadDesigner  =   "Lis404_APS.frx":0466
      UnitType        =   0
      UserResize      =   0
      VisibleCols     =   8
      VisibleRows     =   22
      TextTip         =   4
   End
   Begin VB.PictureBox picResult 
      BackColor       =   &H00E0E0E0&
      Height          =   8145
      Left            =   11055
      ScaleHeight     =   8085
      ScaleWidth      =   5100
      TabIndex        =   51
      Top             =   960
      Width           =   5160
      Begin VB.Frame fraLisResult 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  '����
         Caption         =   "Frame2"
         Height          =   735
         Left            =   -15
         TabIndex        =   52
         Top             =   0
         Width           =   5145
         Begin VB.CheckBox chkSamCmt 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Sample Comment"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   3240
            TabIndex        =   53
            Tag             =   "40205"
            Top             =   435
            Value           =   1  'Ȯ��
            Width           =   1815
         End
         Begin VB.Label Label1 
            Alignment       =   2  '��� ����
            AutoSize        =   -1  'True
            BackStyle       =   0  '����
            Caption         =   "������/�Ұ�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   225
            TabIndex        =   56
            Top             =   135
            Width           =   1170
         End
         Begin VB.Label lblSpecimenNm 
            AutoSize        =   -1  'True
            BackStyle       =   0  '����
            Caption         =   "Serum"
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
            Left            =   810
            TabIndex        =   55
            Top             =   540
            Width           =   645
         End
         Begin VB.Label lblSpecimen 
            AutoSize        =   -1  'True
            BackStyle       =   0  '����
            Caption         =   "��ü : "
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
            Left            =   165
            TabIndex        =   54
            Tag             =   "157"
            Top             =   540
            Width           =   630
         End
         Begin VB.Shape Shape3 
            BorderColor     =   &H00FFFFFF&
            FillColor       =   &H00DF6A3E&
            FillStyle       =   0  '�ܻ�
            Height          =   360
            Left            =   60
            Shape           =   4  '�ձ� �簢��
            Top             =   45
            Width           =   1470
         End
      End
      Begin RichTextLib.RichTextBox rtfResult 
         Height          =   8070
         Left            =   -30
         TabIndex        =   57
         Top             =   45
         Visible         =   0   'False
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   14235
         _Version        =   393217
         BackColor       =   16777207
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         RightMargin     =   9000
         TextRTF         =   $"Lis404_APS.frx":2057
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
   End
   Begin FPSpread.vaSpread tblOrdSheet 
      Height          =   6855
      Left            =   1380
      TabIndex        =   4
      Top             =   2010
      Width           =   9675
      _Version        =   196608
      _ExtentX        =   17066
      _ExtentY        =   12091
      _StockProps     =   64
      AllowCellOverflow=   -1  'True
      BackColorStyle  =   1
      DisplayColHeaders=   0   'False
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
      GrayAreaBackColor=   14411494
      GridColor       =   14013909
      GridShowVert    =   0   'False
      MaxCols         =   45
      OperationMode   =   1
      ScrollBars      =   2
      ShadowColor     =   16252927
      ShadowDark      =   14737632
      ShadowText      =   0
      SpreadDesigner  =   "Lis404_APS.frx":23A2
      TextTip         =   4
   End
   Begin VB.PictureBox picPtList 
      Align           =   3  '���� ����
      BackColor       =   &H00E0E0E0&
      DragMode        =   1  '�ڵ�
      Height          =   8145
      Left            =   0
      ScaleHeight     =   8085
      ScaleWidth      =   1320
      TabIndex        =   25
      Top             =   960
      Visible         =   0   'False
      Width           =   1380
      Begin VB.CheckBox chkVerified 
         BackColor       =   &H00D7E6E6&
         Caption         =   "���� ������� ��� �˻�"
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
         ForeColor       =   &H00553755&
         Height          =   225
         Left            =   1800
         TabIndex        =   68
         Top             =   405
         Width           =   2460
      End
      Begin VB.CheckBox chkAllWard 
         BackColor       =   &H00D7E6E6&
         Caption         =   "��ü����"
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
         Left            =   2370
         TabIndex        =   67
         Top             =   120
         Width           =   1035
      End
      Begin MSComctlLib.ListView lvwPtList 
         Height          =   6840
         Left            =   30
         TabIndex        =   31
         Top             =   1260
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   12065
         View            =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16643054
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
         BackColor       =   &H00E0E0E0&
         Height          =   645
         Left            =   45
         TabIndex        =   26
         Tag             =   "136"
         Top             =   630
         Width           =   4200
         Begin VB.OptionButton optSort 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&ID"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   1995
            TabIndex        =   29
            Tag             =   "15304"
            Top             =   300
            Width           =   495
         End
         Begin VB.OptionButton optSort 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&Name"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   2505
            TabIndex        =   28
            Tag             =   "15305"
            Top             =   285
            Value           =   -1  'True
            Width           =   810
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
            TabIndex        =   27
            Text            =   "��"
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
            MouseIcon       =   "Lis404_APS.frx":6AAF
            MousePointer    =   99  '����� ����
            TabIndex        =   30
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
      Begin VB.Label lblWardId 
         BackColor       =   &H00DBE6E6&
         BackStyle       =   0  '����
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
         ForeColor       =   &H00553755&
         Height          =   180
         Left            =   3465
         MouseIcon       =   "Lis404_APS.frx":6DB9
         MousePointer    =   99  '����� ����
         TabIndex        =   69
         ToolTipText     =   "Click�Ͻø� �����ð��� ������ �� �ֽ��ϴ�."
         Top             =   120
         Width           =   720
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00808080&
         FillColor       =   &H00E8F7F7&
         FillStyle       =   0  '�ܻ�
         Height          =   270
         Left            =   3420
         Shape           =   4  '�ձ� �簢��
         Top             =   90
         Width           =   795
      End
      Begin VB.Label lblPtList 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "Patient List"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   105
         TabIndex        =   32
         Tag             =   "106"
         Top             =   135
         Width           =   1185
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  '�� ����
      Appearance      =   0  '���
      BackColor       =   &H00FCEFE9&
      BorderStyle     =   0  '����
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   0
      ScaleHeight     =   960
      ScaleWidth      =   14745
      TabIndex        =   5
      Top             =   0
      Width           =   14745
      Begin VB.CheckBox chkPtList 
         BackColor       =   &H00FCEFE9&
         Caption         =   "ȯ�ڰ˻� ����Ʈ(&S)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H004A4189&
         Height          =   255
         Left            =   315
         TabIndex        =   8
         Tag             =   "40101"
         Top             =   675
         Width           =   2445
      End
      Begin VB.TextBox txtPtId 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1170
         MaxLength       =   10
         TabIndex        =   0
         Top             =   150
         Width           =   1575
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00E0E0E0&
         Caption         =   "����(&X)"
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
         Left            =   13335
         Style           =   1  '�׷���
         TabIndex        =   7
         Tag             =   "128"
         Top             =   495
         Width           =   1320
      End
      Begin VB.CommandButton cmdReport 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Report"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   13335
         Style           =   1  '�׷���
         TabIndex        =   6
         Tag             =   "40102"
         Top             =   60
         Width           =   1320
      End
      Begin MedControls1.LisLabel lblPtNm 
         Height          =   270
         Left            =   4620
         TabIndex        =   9
         Top             =   45
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   476
         BackColor       =   16703181
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
         Caption         =   ""
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblVerifierNm 
         Height          =   270
         Left            =   10905
         TabIndex        =   10
         Top             =   45
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   476
         BackColor       =   16703181
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
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblVerifyDt 
         Height          =   270
         Left            =   10905
         TabIndex        =   11
         Top             =   360
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   476
         BackColor       =   16703181
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
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblLocation 
         Height          =   270
         Left            =   7890
         TabIndex        =   12
         Top             =   360
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   476
         BackColor       =   16703181
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
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblDeptNm 
         Height          =   270
         Left            =   7890
         TabIndex        =   13
         Top             =   45
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   476
         BackColor       =   16703181
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
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel lblDisease 
         Height          =   225
         Left            =   4620
         TabIndex        =   65
         Top             =   675
         Width           =   8265
         _ExtentX        =   14579
         _ExtentY        =   397
         BackColor       =   16703181
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
         LeftGab         =   100
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '����
         Caption         =   "��  ��  �� : "
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005A5A5A&
         Height          =   180
         Left            =   3525
         TabIndex        =   66
         Tag             =   "103"
         Top             =   705
         Width           =   1080
      End
      Begin VB.Label lblSexAge 
         BackStyle       =   0  '����
         Caption         =   "����/���� : "
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005A5A5A&
         Height          =   180
         Left            =   3510
         TabIndex        =   23
         Tag             =   "108"
         Top             =   390
         Width           =   1125
      End
      Begin VB.Label lblName 
         BackStyle       =   0  '����
         Caption         =   "��      �� : "
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005A5A5A&
         Height          =   180
         Left            =   3600
         TabIndex        =   22
         Tag             =   "103"
         Top             =   90
         Width           =   1080
      End
      Begin VB.Label lblPtId 
         BackStyle       =   0  '����
         Caption         =   "ȯ�� ID : "
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005A5A5A&
         Height          =   180
         Left            =   315
         TabIndex        =   21
         Tag             =   "105"
         Top             =   210
         Width           =   900
      End
      Begin VB.Label lblRptTm 
         BackStyle       =   0  '����
         Caption         =   "�����Ͻ� : "
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005A5A5A&
         Height          =   180
         Left            =   9870
         TabIndex        =   20
         Tag             =   "40108"
         Top             =   405
         Width           =   1110
      End
      Begin VB.Label lblSex 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '����
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
         Height          =   255
         Left            =   4770
         TabIndex        =   19
         Top             =   405
         Width           =   585
      End
      Begin VB.Label lblAge 
         Alignment       =   1  '������ ����
         Appearance      =   0  '���
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '����
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
         Height          =   165
         Left            =   5700
         TabIndex        =   18
         Top             =   390
         Width           =   345
      End
      Begin VB.Label lblAgeDiv 
         Appearance      =   0  '���
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '����
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
         Height          =   180
         Left            =   6270
         TabIndex        =   17
         Top             =   405
         Width           =   60
      End
      Begin VB.Label lblVerifier 
         BackStyle       =   0  '����
         Caption         =   "�� �� �� : "
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005A5A5A&
         Height          =   180
         Left            =   9855
         TabIndex        =   16
         Tag             =   "40111"
         Top             =   90
         Width           =   1125
      End
      Begin VB.Label lblDept 
         BackStyle       =   0  '����
         Caption         =   "�� �� �� : "
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005A5A5A&
         Height          =   180
         Left            =   6915
         TabIndex        =   15
         Tag             =   "40304"
         Top             =   105
         Width           =   975
      End
      Begin VB.Label lblLocation1 
         BackStyle       =   0  '����
         Caption         =   "��     �� : "
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005A5A5A&
         Height          =   180
         Left            =   6900
         TabIndex        =   14
         Tag             =   "102"
         Top             =   405
         Width           =   1005
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H00808080&
         Height          =   285
         Left            =   1155
         Shape           =   4  '�ձ� �簢��
         Top             =   135
         Width           =   1605
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FEDECD&
         Caption         =   "              /"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4620
         TabIndex        =   24
         Top             =   360
         Width           =   1965
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   2
         FillColor       =   &H00FCEFE9&
         FillStyle       =   0  '�ܻ�
         Height          =   960
         Left            =   0
         Shape           =   4  '�ձ� �簢��
         Top             =   0
         Width           =   14775
      End
   End
   Begin VB.PictureBox picOrder 
      BackColor       =   &H00E0E0E0&
      Height          =   1035
      Left            =   1380
      ScaleHeight     =   975
      ScaleWidth      =   9585
      TabIndex        =   33
      Top             =   975
      Width           =   9645
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Height          =   765
         Left            =   1380
         TabIndex        =   34
         Top             =   -60
         Width           =   8190
         Begin VB.CommandButton cmdRefresh 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Re&fresh"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7005
            Style           =   1  '�׷���
            TabIndex        =   3
            Tag             =   "128"
            Top             =   315
            Width           =   1065
         End
         Begin VB.OptionButton optQueryKey 
            BackColor       =   &H00E0E0E0&
            Caption         =   "������"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   225
            TabIndex        =   38
            Tag             =   "15304"
            Top             =   195
            Width           =   945
         End
         Begin VB.OptionButton optQueryKey 
            BackColor       =   &H00E0E0E0&
            Caption         =   "������"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   225
            TabIndex        =   37
            Tag             =   "15305"
            Top             =   450
            Width           =   1005
         End
         Begin VB.CheckBox chkToolTip 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Show &ToolTip"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5520
            TabIndex        =   36
            Top             =   435
            Width           =   1455
         End
         Begin VB.CheckBox chkRefVal 
            BackColor       =   &H00E0E0E0&
            Caption         =   "����ġ ��ȸ"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5520
            TabIndex        =   35
            Top             =   180
            Width           =   1410
         End
         Begin MSComCtl2.DTPicker dtpFromDate 
            Height          =   285
            Left            =   1920
            TabIndex        =   1
            Top             =   270
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   503
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
            CustomFormat    =   "yyy-MM-dd"
            Format          =   24641539
            CurrentDate     =   36328
         End
         Begin MSComCtl2.DTPicker dtpToDate 
            Height          =   285
            Left            =   3765
            TabIndex        =   2
            Top             =   270
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   503
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
            CustomFormat    =   "yyy-MM-dd"
            Format          =   24641539
            CurrentDate     =   36328
         End
         Begin VB.Label lblTo 
            AutoSize        =   -1  'True
            BackStyle       =   0  '����
            Caption         =   "To"
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
            Left            =   3525
            TabIndex        =   40
            Tag             =   "40110"
            Top             =   315
            Width           =   255
         End
         Begin VB.Label lblFrom 
            AutoSize        =   -1  'True
            BackStyle       =   0  '����
            Caption         =   "From"
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
            Left            =   1410
            TabIndex        =   39
            Tag             =   "40105"
            Top             =   300
            Width           =   495
         End
      End
      Begin MedControls1.LisLabel lblKeyDate 
         Height          =   315
         Left            =   30
         TabIndex        =   41
         Top             =   720
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         BackColor       =   14737632
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
         Alignment       =   1
         Caption         =   "ó����"
      End
      Begin MedControls1.LisLabel LisLabel2 
         Height          =   315
         Left            =   945
         TabIndex        =   42
         Top             =   720
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         BackColor       =   14737632
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
         Alignment       =   1
         Caption         =   "��ü"
      End
      Begin MedControls1.LisLabel LisLabel3 
         Height          =   315
         Left            =   2055
         TabIndex        =   43
         Top             =   720
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   556
         BackColor       =   14737632
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
         Alignment       =   1
         Caption         =   "�˻��"
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Left            =   4140
         TabIndex        =   44
         Top             =   720
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         BackColor       =   14737632
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
         Alignment       =   1
         Caption         =   "���"
      End
      Begin MedControls1.LisLabel LisLabel5 
         Height          =   315
         Left            =   5355
         TabIndex        =   45
         Top             =   720
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   556
         BackColor       =   14737632
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
         Alignment       =   1
         Caption         =   "����"
      End
      Begin MedControls1.LisLabel LisLabel6 
         Height          =   315
         Left            =   6315
         TabIndex        =   46
         Top             =   720
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   556
         BackColor       =   14737632
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
         Alignment       =   1
         Caption         =   "HL"
      End
      Begin MedControls1.LisLabel LisLabel7 
         Height          =   315
         Left            =   6765
         TabIndex        =   47
         Top             =   720
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   556
         BackColor       =   14737632
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
         Alignment       =   1
         Caption         =   "DP"
      End
      Begin MedControls1.LisLabel LisLabel8 
         Height          =   315
         Left            =   7335
         TabIndex        =   48
         Top             =   720
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         BackColor       =   14737632
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
         Alignment       =   1
         Caption         =   "����ġ"
      End
      Begin MedControls1.LisLabel LisLabel9 
         Height          =   315
         Left            =   8850
         TabIndex        =   49
         Top             =   705
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         BackColor       =   14737632
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
         Alignment       =   1
         Caption         =   "More"
      End
      Begin VB.Label lblRefresh 
         Alignment       =   2  '��� ����
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "�Ϲ� ���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   225
         TabIndex        =   50
         Top             =   120
         Width           =   930
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00DF6A3E&
         FillStyle       =   0  '�ܻ�
         Height          =   360
         Index           =   0
         Left            =   45
         Shape           =   4  '�ձ� �簢��
         Top             =   30
         Width           =   1260
      End
   End
   Begin DRcontrol1.DrFrame fraTextResult 
      Height          =   8430
      Left            =   2610
      TabIndex        =   58
      Top             =   450
      Visible         =   0   'False
      Width           =   9750
      _ExtentX        =   17198
      _ExtentY        =   14870
      BorderStyle     =   0   'False
      Appearance      =   0
      Title           =   ""
      DelLine         =   0
      BackColor       =   15593969
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.PictureBox Picture3 
         Appearance      =   0  '���
         BackColor       =   &H00FFF8EE&
         ForeColor       =   &H80000008&
         Height          =   7590
         Left            =   165
         ScaleHeight     =   7560
         ScaleWidth      =   9375
         TabIndex        =   59
         Top             =   630
         Width           =   9405
         Begin RichTextLib.RichTextBox txtRstCmt1 
            Height          =   7320
            Left            =   90
            TabIndex        =   60
            Top             =   75
            Width           =   9165
            _ExtentX        =   16166
            _ExtentY        =   12912
            _Version        =   393217
            BackColor       =   16775671
            BorderStyle     =   0
            Enabled         =   -1  'True
            ReadOnly        =   -1  'True
            ScrollBars      =   2
            Appearance      =   0
            TextRTF         =   $"Lis404_APS.frx":70C3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����ü"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Label lblRstCmt1 
         BackStyle       =   0  '����
         Caption         =   "Result "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   240
         TabIndex        =   61
         Tag             =   "40204"
         Top             =   195
         Width           =   2205
      End
   End
End
Attribute VB_Name = "frm404AllResult_APS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'% ������ �������� ����

Option Explicit

'-------------------------------------
'�غκ���/�������� �����ȸ ����
'-------------------------------------
'#Const AllowAPSResultReview = True
'#Const AllowBBSResultReview = True
'-------------------------------------


'Private objResult As New clsAPSResult
'Private objSql As New clsAPSSqlResult
Private MyPatient As New clsPatient   'ȯ�� Ŭ����
Private MySql As New clsLISSqlReview     'Sql�� Ŭ����
Private ClearFg As Boolean
Private OrderFg As Boolean
Private ResultFg As Boolean
Private MsgFg As Boolean
Private OldRow As Long
Private OldBackColor As Long
Private TopLeftShow As Boolean
Private TopLeftShow1 As Boolean
Private TopLeftShow2 As Boolean
Private aryMesg() As String

Private WithEvents objMyList As clsS2DLP
Attribute objMyList.VB_VarHelpID = -1
'Private WithEvents objTextForm As Form

Private mvarDeptCd As String

Public PtFg As Boolean
Public QueryFg As Boolean
Private StopFg As Boolean

Public Event LastFormUnload()
Public Event ThisFormUnload()


Public Property Get DeptCd() As String
    DeptCd = mvarDeptCd
End Property
Public Property Let DeptCd(ByVal vData As String)
    mvarDeptCd = vData
End Property

Private Sub chkAllWard_Click()
    If chkAllWard.Value = 0 Then
        chkVerified.Value = 0
        chkVerified.Enabled = False
    Else
        chkVerified.Enabled = True
    End If

End Sub

'% ȯ�ڸ���Ʈ Display ����
Private Sub chkPtList_Click()
    If chkPtList.Value = 1 Then
        lblWardId.Caption = mvarDeptCd
        picPtList.Visible = True
        picPtList.Width = 4290
        picOrder.Left = picPtList.Width
        tblOrdSheet.Left = picOrder.Left
        picResult.Left = picOrder.Left + picOrder.Width
        tblResult.Left = picResult.Left + 50
        picFootNote.Left = picResult.Left + 50
        txtSearchKey.SetFocus
    ElseIf chkPtList.Value = 0 Then
        picPtList.Visible = False
        picOrder.Left = 0
        tblOrdSheet.Left = picOrder.Left
        picResult.Left = picOrder.Left + picOrder.Width
        tblResult.Left = picResult.Left + 50
        picFootNote.Left = picResult.Left + 50
    End If
    Exit Sub

End Sub

Private Sub chkRefVal_Click()

    Dim tmpTestCd As String
    Dim tmpSpcCd As String
    Dim tmpVfyDt As String
    Dim tmpSex As String
    Dim tmpAgeDay As String
    Dim tmpRs1 As New DrRecordSet
    Dim tmpRefFromVal As Double
    Dim tmpRefToVal As Double
    Dim tmpRefCd As String
    Dim I As Long, j As Long
    Dim SqlStmt As String
    
    With tblOrdSheet
        For I = 1 To .MaxRows
            '����ġ �˻�
            .Row = I
            .Col = 8: If .Value <> CS_QuestionMark Then GoTo RefSkip
            
            .Col = 25:  tmpSex = Trim(.Value)
            .Col = 26:  tmpAgeDay = Trim(.Value)
            .Col = 27:  tmpTestCd = Trim(.Value)
            .Col = 28:  tmpSpcCd = Trim(.Value)
            .Col = 29:  tmpVfyDt = Trim(.Value)
                        If tmpVfyDt = "" Then tmpVfyDt = Format(Now, CS_DateDbFormat)
         
            SqlStmt = MySql.SqlGetReference(tmpTestCd, tmpSpcCd, tmpVfyDt, "B", tmpAgeDay)
            Set tmpRs1 = OpenRecordSet(SqlStmt)
            If tmpRs1.EOF Then
                '"B"(Both)�� �ش��ϴ� ����ġ�� ���� ��� ȯ�ڼ����� �ش��ϴ� ����Ÿ �˻�
                '--> ���� Both�� ��ϵ�.
                tmpRs1.RsClose
                SqlStmt = MySql.SqlGetReference(tmpTestCd, tmpSpcCd, tmpVfyDt, tmpSex, tmpAgeDay)
                Set tmpRs1 = OpenRecordSet(SqlStmt)
            End If
            If tmpRs1.EOF Then
                tmpRefCd = Space(5)
            Else
                tmpRefFromVal = Val("" & tmpRs1.Fields("RefValFrom").Value)
                tmpRefToVal = Val("" & tmpRs1.Fields("RefValTo").Value)
                tmpRefCd = Trim("" & tmpRs1.Fields("RefCd").Value)
                If tmpRefFromVal <> 0 Or tmpRefToVal <> 0 Then _
                   tmpRefCd = tmpRefFromVal & "  -  " & tmpRefToVal
            End If
            tmpRs1.RsClose
            For j = I To .MaxRows
                .Row = j
                .Col = 27   '����ġ
                If Trim(.Value) = tmpTestCd Then _
                    .Col = 8: .Value = tmpRefCd: .ForeColor = DCM_Green
            Next
         
            DoEvents
        
RefSkip:
        Next
    End With
    Set tmpRs1 = Nothing
    
End Sub

'% �ؽ�Ʈ ������� �ڽ� Display ����
'Private Sub chkRstCmt_Click()
'    If chkRstCmt.Value = 1 And picRstText.Visible = False Then
'        picRstText.Visible = True
'        tblResult.Height = tblResult.Height - picRstText.Height
'    ElseIf chkRstCmt.Value = 0 And picRstText.Visible = True Then
'        picRstText.Visible = False
'        tblResult.Height = tblResult.Height + picRstText.Height
'    End If
'End Sub

'% ǲ��Ʈ, ��ü����ũ �ڽ� Display ����
Private Sub chkSamCmt_Click()
    If chkSamCmt.Value = 1 Then
        picFootNote.Visible = True
        tblResult.Height = tblResult.Height - picFootNote.Height
        'picRstText.Top = picRstText.Top - picFootNote.HEIGHT
    ElseIf chkSamCmt.Value = 0 Then
        picFootNote.Visible = False
        tblResult.Height = tblResult.Height + picFootNote.Height
        'picRstText.Top = picRstText.Top + picFootNote.HEIGHT
    End If
End Sub

'%����
Private Sub cmdExit_Click()
    Unload Me
    RaiseEvent ThisFormUnload
    Set objMyList = Nothing
    If IsLastForm Then RaiseEvent LastFormUnload
End Sub

Private Sub chkVerified_Click()
    Call txtSearchKey_KeyPress(vbKeyReturn)
End Sub

Private Sub cmdRefresh_Click()
   '% ó����ȸ
   OldRow = 0
   OrderFg = False
   Call dtpToDate_KeyDown(vbKeyReturn, 0)
End Sub

'% ����Ʈ ���
Private Sub cmdReport_Click()
   
   Dim MyData As New clsResults
   Dim MyReport As New clsResultReport
   Dim strLastRst As String
   'Dim MyReport As New clsBatchReport
   Dim I As Integer
   
   Screen.MousePointer = vbArrowHourglass
   
   With tblOrdSheet
'        MyReport.DateCaption = lblKeyDate.Caption
        For I = 1 To .DataRowCnt
            .Row = I
            .Col = 1:  MyData.OrdDt = .Value   'ó����
            .Col = 2: MyData.SpcNm = .Value  '��ü��
            .Col = 27
            If Trim(.Value) <> "" Then
                .Col = 3: MyData.TestNm = Mid(.Value, 1, 25) '�˻��
            Else
                .Col = 3: MyData.TestNm = .Value
            End If
            .Col = 29: MyData.VfyDt = .Value     '������
            .Col = 4: MyData.RstCd = .Value     '���
            .Col = 5: MyData.RstUnit = .Value    '����
            .Col = 6: MyData.HLDiv = .Value      'High/Low
            .Col = 7: MyData.DPDiv = .Value      'Delta/Panic
            .Col = 8: MyData.RefRng = .Value    '����ġ
            .Col = 34: MyData.TxtFg = .Value     '�Ұ߿���
            .Col = 17: MyData.WorkArea = .Value   'WorkArea
            .Col = 18: MyData.AccDt = .Value        'AccDt
            .Col = 19: MyData.AccSeq = .Value     'AccSeq
            .Col = 20: strLastRst = .Value     '�ֱٰ��
            .Col = 21:
                       If Trim(strLastRst) <> "" Then
                          MyData.LastRst = strLastRst & " (" & Mid(.Value, 4, 5) & ")"
                       Else
                          MyData.LastRst = strLastRst
                       End If
            .Col = 27: MyData.TestCd = .Value       '�˻��ڵ�
            .Col = 28: MyData.SpcCd = .Value       '��ü�ڵ�
            .Col = 30: MyData.TestDiv = .Value      'TestDiv
            .Col = 32: MyData.OrdDate = .Value
            .Col = 33: MyData.SpcName = .Value
            .Col = 35: MyData.FootNoteFg = .Value  'footnotefg
            .Col = 36: MyData.RmkCd = .Value         'Remark �ڵ�
            .Col = 37: MyData.SenFg = .Value         '����������
            Call MyReport.Add(MyData)
Skip:
        Next
   End With
   MyReport.ptid = MedSetPtid(txtPtId.Text)
   MyReport.PtNm = lblPtNm.Caption
   MyReport.PtSex = lblSex.Caption
   MyReport.PtAge = lblAge.Caption & " " & lblAgeDiv.Caption
   MyReport.FromDt = Format(dtpFromDate.Value, CS_DateLongFormat)
   MyReport.ToDt = Format(dtpToDate.Value, CS_DateLongFormat)
'   MyReport.VfyDt = lblVerifyDt.Caption
'   MyReport.VfyNM = lblVerifierNm.Caption
'   MyReport.MdfDt = "2001/05/23"            '������
'   MyReport.Dept = lblDeptNm.Caption
'   MyReport.Ward = lblLocation.Caption
   
   Call MyReport.Print_Report
   
   Screen.MousePointer = vbDefault
            
'      .ReDraw = False
'
'      .DisplayColHeaders = True
'      .Row = 0
'      .RowHeight(0) = 20
'      .PrintAbortMsg = "�����... ����Ϸ��� Cancel ��ư�� ��������"
'      .PrintJobName = "Result Print"
'      .PrintHeader = "/l ��  ȯ�ں� �˻���/n/n" & _
'                            "/l   ȯ �� : " & txtPtId.Text & Space(3) & lblPtNm.Caption & Space(3) & lblSex.Caption & " / " & lblAge.Caption & " " & lblAgeDiv.Caption & "/n" & _
'                              "/l   �� �� : " & Format(dtpFromDate.Value, CS_DateFormat) & "  ~  " & Format(dtpToDate.Value, CS_DateFormat) & "/n/n"
'      .PrintFooter = "/cPage /p"
'      .PrintBorder = True
'      .PrintColor = False
'      .PrintGrid = True
'      .PrintMarginTop = 100
'      .PrintMarginBottom = 100
'      .PrintMarginLeft = 1700
'      .PrintMarginRight = 100
'      .PrintType = PrintTypeAll    'SS_PRINT_ALL
'      .PrintRowHeaders = True
'      .PrintColHeaders = True
'      .PrintBorder = True
'      '.GridSolid = False
'      .PrintGrid = False
'      .PrintShadows = False
'      .PrintUseDataMax = True
'      ' Perform the printing action
'      .Action = ActionPrint
'
'      .DisplayColHeaders = False
'      .ReDraw = True

End Sub

'% ��ȸ�Ⱓ �Է� (From Date)
Private Sub dtpFromDate_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo Err_Trap
   If KeyCode = vbKeyReturn Then dtpToDate.SetFocus
Err_Trap:
End Sub

'% ��ȸ�Ⱓ �Է� (To Date)
Private Sub dtpToDate_KeyDown(KeyCode As Integer, Shift As Integer)
   
    '% ó����ȸ
    Dim I As Integer
    Dim ResultExist As Boolean
   
    If KeyCode <> vbKeyReturn Then Exit Sub
   
    On Error GoTo Err_Trap
   
    If Format(dtpToDate.Value, CS_DateDbFormat) < Format(dtpFromDate.Value, CS_DateDbFormat) Then
        MsgBox "�Ⱓ �Է� �����Դϴ�. ��¥�� �����Ͻʽÿ�..", vbExclamation, "�Է¿���"
        dtpFromDate.SetFocus
        Exit Sub
    End If
   
    cmdRefresh.Enabled = False
    dtpFromDate.Enabled = False
    dtpToDate.Enabled = False
   
    Call FieldClear
    Call TableClear
    Call ResultClear
   
    'Status Bar Popup
'    Dim objPrgBar As New clsProgress
'
'    DoEvents
'    With objPrgBar
'        .Mode = 0
'        .CaptionOn = False
'        .Msg = lblPtNm.Caption & " ���� �˻� ��������� �˻����Դϴ�..."
'        .Min = 0
'        .Max = 100
'        .Value = 0
'        .Visible = True
'    End With
    
    Dim objPrgBar As New clsProgressBar
    
    With objPrgBar
        .Choice = True
        .Appearance = aPlate
        .SetMyForm Me
        .XWidth = tblOrdSheet.Width
        .XPos = tblOrdSheet.Left
        .YPos = tblOrdSheet.Top - 280
        .YHeight = 280
        .ForeColor = &H864B24
        .Msg = lblPtNm.Caption & " ���� �˻� ��������� �˻����Դϴ�..."
        .Value = 1
    End With
    
    DoEvents
   
    With tblOrdSheet
        '.ReDraw = False
        .MaxRows = 0
           
        ResultExist = False
        ResultExist = ResultExist Or DisplayOrders("3", objPrgBar)
        
        '.ReDraw = True
        .Col = 1: .Row = 1: .Action = ActionActiveCell
    End With
    
    'objPrgBar.Visible = False
    Set objPrgBar = Nothing
    
    cmdRefresh.Enabled = True
    dtpFromDate.Enabled = True
    dtpToDate.Enabled = True
   
    If Not ResultExist Then
        MsgBox "�� ȯ�ڴ� �Է��Ͻ� �Ⱓ���ȿ� ����� ����� �����ϴ�."
        dtpFromDate.SetFocus
        Exit Sub
    End If
   
End_Pos:
    ClearFg = False
    ResultFg = False
    OrderFg = True
    'tblOrdSheet.SetFocus
    cmdReport.Enabled = True
    Exit Sub
    
Err_Trap:
    MsgBox Err.Description, vbExclamation, "�����߻�"
    GoTo End_Pos
    
    'Resume Next
End Sub

'% ȯ��ID, ó����(ä����)�� �������� ó�泻���� �˻��Ѵ�.

Private Function DisplayOrders(ByVal pTestDiv As String, ByRef barStatus As Object) As Boolean

    Dim I As Integer, j As Integer
    Dim SqlStmt As String
    Dim ColCnt As Integer
    Dim tmpTestNm As String
    Dim tmpRs As New DrRecordSet
    Dim SvKeyDt As String, SvSpcNm As String
    Dim pWorkArea As String, pAccDt As String, pAccSeq As String
    Dim strKeyFld As String
    Dim strNotice As String, strTmp As String
   
    'barStatus.Value = (pTestDiv + 1) * 30
    'lblStatus.Caption = lblPtNm.Caption & " ���� " & Choose(pTestDiv + 1, "�Ϲ�", "Ư��", "�̻���") & "�˻� ��������� �˻����Դϴ�..."
   
    If StopFg Then Exit Function
    
    Me.Enabled = False
    QueryFg = True
   
    MouseRunning
    barStatus.Value = 20
    
    'ó����/������ ����
    strKeyFld = IIf(optQueryKey(1).Value, "examdt", "rcvdt")
    SqlStmt = MySql.SqlQueryAllResults(MedSetPtid(txtPtId.Text), strKeyFld, Format(dtpFromDate.Value, CS_DateDbFormat), _
                                    Format(dtpToDate.Value, CS_DateDbFormat), pTestDiv)
    barStatus.Value = 40
    
    'Query
    ColCnt = tmpRs.OpenCursor(DBConn, SqlStmt)
    
    SvKeyDt = "": SvSpcNm = ""
    
    DoEvents
   
    ReDim aryMesg(0)
    DisplayOrders = False
    
    If ColCnt = 0 Then GoTo NoData
    
    With tblOrdSheet
      
        '.ReDraw = False
      
        While (tmpRs.FetchCursor(ColCnt))
         
            If StopFg Then
                tmpRs.CloseCursor
                StopFg = False
                GoTo NoData
            End If
         
            If barStatus.Value >= barStatus.Max Then barStatus.Max = barStatus.Max + 50
            barStatus.Value = barStatus.Value + 1

            DoEvents
        
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            .RowHeight(.MaxRows) = 11.5
            
            If SvKeyDt <> Trim("" & tmpRs.GetValue("KeyDate")) Then
                .Col = 1:   .Value = Trim("" & tmpRs.GetValue("KeyDate"))
                            .FontBold = True: .ForeColor = vbBlack       '-- ó����
                .Col = 2:   .Value = Trim("" & tmpRs.GetValue("SpcNm"))
                            .FontBold = True: .ForeColor = DCM_LightRed  '-- ��ü��
                SvKeyDt = Trim("" & tmpRs.GetValue("KeyDate"))
                SvSpcNm = Trim("" & tmpRs.GetValue("SpcNm"))
            Else
                .Col = 1:   .Value = "":
                            .FontBold = True: .ForeColor = vbBlack       '-- ó����
                            If SvSpcNm <> Trim("" & tmpRs.GetValue("SpcNm")) Then
                                .Col = 2:
                                .Value = Trim("" & tmpRs.GetValue("SpcNm"))
                                .FontBold = True: .ForeColor = DCM_LightRed  '-- ��ü��
                                SvSpcNm = Trim("" & tmpRs.GetValue("SpcNm"))
                            Else
                                .Col = 2:
                                .Value = "":
                                .FontBold = True: .ForeColor = DCM_LightRed  '-- ��ü��
                            End If
            End If
            
            .Col = 32:  .Value = Trim("" & tmpRs.GetValue("KeyDate"))   'ó����
            .Col = 33:  .Value = Trim("" & tmpRs.GetValue("SpcNm"))     '��ü��
            
            .Col = 3:   '-- �˻��
                        .ForeColor = DCM_MidBlue
                        tmpTestNm = Mid(Trim("" & tmpRs.GetValue("TestLongNm")), 1, 33)
                        If (Trim("" & tmpRs.GetValue("DetailFg")) = "" And _
                            Trim("" & tmpRs.GetValue("DetailItem")) = "") Or _
                            Trim("" & tmpRs.GetValue("RstDiv")) = "*" Then
                            
                            .Value = tmpTestNm & " " & String(35 - Len(tmpTestNm), ".")
                        Else
                            .Value = Space(4) & tmpTestNm & " " & String(35 - Len("  " & tmpTestNm), ".")
                        End If
                        
            .Col = 4:   '-- �����(�ڵ��� ���..)
                        .ForeColor = DCM_Brown   '����
                        If Trim("" & tmpRs.GetValue("VfyDt")) = "" Then
                            .Value = "��Ȯ"
                            .ForeColor = DCM_MidGray: .FontBold = False:
                        Else
                            If Trim("" & tmpRs.GetValue("RstCdNm")) = "" Then
                                .TypeHAlign = TypeHAlignCenter
                                .Value = Trim("" & tmpRs.GetValue("RstCd"))
                            Else
                                .CellType = CellTypeEdit
                                .TypeHAlign = TypeHAlignLeft
                                .Value = " " & Trim("" & tmpRs.GetValue("RstCdNm"))
                            End If
                            If Trim("" & tmpRs.GetValue("SenFg")) = "Y" Then
                                .Value = "Growth"
                            ElseIf Trim("" & tmpRs.GetValue("RstCd")) = "" Then
                                .Value = Space(3)
                            End If
                        End If
                        
            .Col = 5:   '-- �������
                        .Value = Trim("" & tmpRs.GetValue("RstUnit"))
            
            .Col = 6    '-- High / Low
                        .Value = ""
                        If Trim("" & tmpRs.GetValue("VfyDt")) <> "" Then
                            If Trim("" & tmpRs.GetValue("HLDiv")) = HLDIV_HIGH_CD Then .Value = HLDIV_HIGH_FG: .ForeColor = DCM_LightRed
                            If Trim("" & tmpRs.GetValue("HLDiv")) = HLDIV_LOW_CD Then .Value = HLDIV_LOW_FG:  .ForeColor = DCM_LightBlue
                            If Trim("" & tmpRs.GetValue("HLDiv")) = "*" Then .Value = "*": .ForeColor = vbRed
                        End If
            
            .Col = 7:   '-- Delta/Panic
                        .Value = Trim("" & tmpRs.GetValue("DPDiv")): .ForeColor = vbRed
            
            .Col = 8:   '-- ����ġ
                        If Trim("" & tmpRs.GetValue("RstDiv")) <> "*" And Trim("" & tmpRs.GetValue("TestDiv")) < "4" Then .Value = CS_QuestionMark
            
            .Col = 9:   '-- More Result...
                        .Value = "": .ForeColor = DCM_LightBlue
                        If Trim("" & tmpRs.GetValue("TxtFg")) > "0" Then .Value = CS_FingerMark
                        If Trim("" & tmpRs.GetValue("TxtFg")) = "Y" Then .Value = CS_FingerMark
                        If Trim("" & tmpRs.GetValue("SenFg")) = "Y" Then .Value = CS_FingerMark
                        If (Trim("" & tmpRs.GetValue("DetailFg")) = "" And _
                            Trim("" & tmpRs.GetValue("DetailItem")) = "") Or _
                            Trim("" & tmpRs.GetValue("RstDiv")) = "*" Then
                            If Trim("" & tmpRs.GetValue("FootNoteFg")) = "1" Then .Value = CS_FingerMark
                            If Trim("" & tmpRs.GetValue("RmkCd")) <> "" Then .Value = CS_FingerMark
                        End If
                        If Trim("" & tmpRs.GetValue("DcFg")) = "1" Then .Value = .Value & "*"
                        If Trim("" & tmpRs.GetValue("TestDiv")) = "4" Then .Value = CS_FingerMark    '�غκ���
                        If Trim("" & tmpRs.GetValue("TestDiv")) = "5" Then .Value = CS_FingerMark    '��������
                        If Trim("" & tmpRs.GetValue("OrdDiv")) = CMT_ORDDIV Then .Value = CS_FingerMark    '���հ���
         
            .Col = 10: .Value = Trim("" & tmpRs.GetValue("OrdDate"))        '-- ó����
            .Col = 11: .Value = Trim("" & tmpRs.GetValue("OrdNo"))          '-- ó���ȣ
            .Col = 12: .Value = Trim("" & tmpRs.GetValue("OrdDoct"))        '-- ó����
            .Col = 13: .Value = Trim("" & tmpRs.GetValue("ColDtTm"))        '-- ä���Ͻ�
            .Col = 14: .Value = Trim("" & tmpRs.GetValue("ColId"))          '-- ä����
            .Col = 15: .Value = Trim("" & tmpRs.GetValue("RcvDtTm"))        '-- �����Ͻ�
            .Col = 16: .Value = Trim("" & tmpRs.GetValue("RcvId"))          '-- ������
            .Col = 17: .Value = Trim("" & tmpRs.GetValue("WorkArea")):  pWorkArea = .Value  'WorkArea
            .Col = 18: .Value = Trim("" & tmpRs.GetValue("AccDt")):     pAccDt = .Value     'AccDt
            .Col = 19: .Value = Trim("" & tmpRs.GetValue("AccSeq")):    pAccSeq = .Value    'AccSeq
            .Col = 20: .Value = Trim("" & tmpRs.GetValue("LastRst"))        '-- �ֱٰ��
            .Col = 21: .Value = Trim("" & tmpRs.GetValue("LstVfyDtTm"))     '-- �ֱٰ���Ͻ�
            .Col = 22: .Value = Trim("" & tmpRs.GetValue("LastVfyId"))      '-- �ֱٰ�� ������
            .Col = 23: .Value = Trim("" & tmpRs.GetValue("VfyDtTm"))        '-- �����Ͻ�
            .Col = 24: .Value = Trim("" & tmpRs.GetValue("VfyId"))          '-- ������
            .Col = 25: .Value = Trim("" & tmpRs.GetValue("Sex"))            '-- Sex
            .Col = 26: .Value = Trim("" & tmpRs.GetValue("AgeDay"))         '-- AgeDay
            .Col = 27: .Value = Trim("" & tmpRs.GetValue("TestCd"))         '-- �˻��ڵ�
            .Col = 28: .Value = Trim("" & tmpRs.GetValue("SpcCd"))          '-- ��ü�ڵ�
            .Col = 29: .Value = Trim("" & tmpRs.GetValue("VfyDt"))          '-- ������
            .Col = 30: .Value = Trim("" & tmpRs.GetValue("TestDiv"))        '-- �˻籸��
            .Col = 31: .Value = Trim("" & tmpRs.GetValue("DeptCd"))         '-- �����
            .Col = 34: .Value = Trim("" & tmpRs.GetValue("TxtFg"))          '-- �Ұ߰������
            .Col = 35: .Value = Trim("" & tmpRs.GetValue("FootNoteFg"))     '-- Footnote ����
            .Col = 36: .Value = Trim("" & tmpRs.GetValue("RmkCd"))          '-- Remark �ڵ�
            .Col = 37: .Value = Trim("" & tmpRs.GetValue("SenFg"))          '-- ������ ����
            .Col = 38: .Value = Trim("" & tmpRs.GetValue("OrdDiv"))         '-- ó�汸��
            .Col = 39: .Value = Trim("" & tmpRs.GetValue("UnitQty"))        '-- ��������
            .Col = 40: .Value = Trim("" & tmpRs.GetValue("ReqDt"))          '-- ����������
            .Col = 41: .Value = Trim("" & tmpRs.GetValue("ReqTm"))          '-- ���������ð�
            .Col = 42: .Value = Trim("" & tmpRs.GetValue("WardId"))         '-- ����
            .Col = 43: .Value = Trim("" & tmpRs.GetValue("HosilId"))        '-- ȣ��
            .Col = 44: .Value = Trim("" & tmpRs.GetValue("RoomId"))        '-- ȣ��
            .Col = 45: .Value = Trim("" & tmpRs.GetValue("Notice"))        '-- ȣ��
            
            ReDim Preserve aryMesg(UBound(aryMesg) + 1)
            aryMesg(UBound(aryMesg)) = Trim("" & tmpRs.GetValue("Mesg"))    '-- �����Remark
         
            DisplayOrders = True
            
            If Trim("" & tmpRs.GetValue("Notice")) <> "" Then
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                .Col = 3
                .TypeEditMultiLine = False
                .ForeColor = vbBlack
                .Value = "�� Clinical Notice "  '& vbCrLf & Trim("" & tmpRs.GetValue("Notice"))
                .RowHeight(.MaxRows) = .MaxTextRowHeight(.MaxRows)
                strNotice = Trim("" & tmpRs.GetValue("Notice"))
                strNotice = Replace(strNotice, vbCr, "")
                strTmp = medShift(strNotice, vbLf)
                While strTmp <> ""
                    .MaxRows = .MaxRows + 1
                    .Row = .MaxRows
                    .Col = 3
                    .TypeEditMultiLine = False
                    .ForeColor = &H747474
                    .Value = strTmp
                    strTmp = medShift(strNotice, vbLf)
                Wend
            End If
      
        Wend
      
        tmpRs.CloseCursor
        
        .Row = -1: .Col = 3: .Col2 = 5
        .BlockMode = True
        .AllowCellOverflow = True
        .BlockMode = False
      
        '.RowHeight(-1) = 11.5
        .ReDraw = True
      
        If chkRefVal.Value = 0 Then GoTo ExitPos
      
        Dim tmpTestCd As String
        Dim tmpSpcCd As String
        Dim tmpVfyDt As String
        Dim tmpSex As String
        Dim tmpAgeDay As String
        Dim tmpRs1 As New DrRecordSet
        Dim tmpRefFromVal As Double
        Dim tmpRefToVal As Double
        Dim tmpRefCd As String
      
        barStatus.Value = barStatus.Max - 10
        barStatus.Msg = "�ӻ� ����ġ�� �˻��ϰ� �ֽ��ϴ�.."
        DoEvents
        For I = 1 To .MaxRows
            If barStatus.Value < barStatus.Max Then barStatus.Value = barStatus.Value + 1
            '����ġ �˻�
            .Row = I
            .Col = 8: If .Value <> CS_QuestionMark Then GoTo RefSkip
            
            .Col = 25:  tmpSex = Trim(.Value)
            .Col = 26:  tmpAgeDay = Trim(.Value)
            .Col = 27:  tmpTestCd = Trim(.Value)
            .Col = 28:  tmpSpcCd = Trim(.Value)
            .Col = 29:  tmpVfyDt = Trim(.Value)
                        If tmpVfyDt = "" Then tmpVfyDt = Format(Now, CS_DateDbFormat)
         
            SqlStmt = MySql.SqlGetReference(tmpTestCd, tmpSpcCd, tmpVfyDt, "B", tmpAgeDay)
            Set tmpRs1 = OpenRecordSet(SqlStmt)
            If tmpRs1.EOF Then
                '"B"(Both)�� �ش��ϴ� ����ġ�� ���� ��� ȯ�ڼ����� �ش��ϴ� ����Ÿ �˻�
                '--> ���� Both�� ��ϵ�.
                tmpRs1.RsClose
                SqlStmt = MySql.SqlGetReference(tmpTestCd, tmpSpcCd, tmpVfyDt, tmpSex, tmpAgeDay)
                Set tmpRs1 = OpenRecordSet(SqlStmt)
            End If
            If tmpRs1.EOF Then
                tmpRefCd = Space(5)
            Else
                tmpRefFromVal = Val("" & tmpRs1.Fields("RefValFrom").Value)
                tmpRefToVal = Val("" & tmpRs1.Fields("RefValTo").Value)
                tmpRefCd = Trim("" & tmpRs1.Fields("RefCd").Value)
                If tmpRefFromVal <> 0 Or tmpRefToVal <> 0 Then _
                   tmpRefCd = tmpRefFromVal & "  -  " & tmpRefToVal
            End If
            tmpRs1.RsClose
            For j = I To .MaxRows
                .Row = j
                .Col = 27   '����ġ
                If Trim(.Value) = tmpTestCd Then _
                    .Col = 8: .Value = tmpRefCd: .ForeColor = DCM_Green
            Next
         
            DoEvents

RefSkip:
        Next
      
ExitPos:
        barStatus.Value = barStatus.Max
        DoEvents
        medSleep (500)
'        barStatus.Visible = False
      
        If .MaxRows < 33 Then .MaxRows = 33
      
    End With
   
NoData:
    QueryFg = False
    Me.Enabled = True
    MouseDefault
    DoEvents
    Set tmpRs = Nothing
    Set tmpRs1 = Nothing
   
End Function


Private Sub Form_Activate()
    'medMain.lblSubMenu.Caption = Me.Caption
    'If Screen.ActiveControl Is Nothing Then Exit Sub
    MsgFg = False
    Call chkPtList_Click
    If Trim(gPatientId) <> "" Then txtPtId.Text = MedGetPtid(gPatientId)
On Error GoTo Err_Trap
    txtPtId.SetFocus
    DoEvents
Err_Trap:
    If Trim(txtPtId.Text) <> "" Then SendKeys "{TAB}"
End Sub

Private Sub Form_Terminate()
    StopFg = True
End Sub

Private Sub lblReset_Click()
    lvwPtList.ListItems.Clear
    txtSearchKey.Text = ""
End Sub

Private Sub lblWardId_Click()

    Set objMyList = New clsS2DLP
    
    With objMyList
        .Caption = "���� ��ȸ"
        .Tag = "WardId"
        .HeadName = "�����ڵ�,������"
        Call .ListPop(, 1640, 10550, ObjLISComCode.WardId)
    End With
    'Set objMyList = Nothing
End Sub
Private Sub objMyList_SendCode(ByVal SelString As String)
    If objMyList.Tag = "WardId" Then
        lblWardId.Caption = Trim(medGetP(SelString, 1, ";"))
        lblWardId.Tag = "1"
        mvarDeptCd = lblWardId.Caption
        chkVerified.Enabled = True
        If chkVerified.Value = 1 Then Call txtSearchKey_KeyPress(vbKeyReturn)
    End If
    
End Sub
Private Sub lvwPtList_ItemClick(ByVal Item As MSComctlLib.ListItem)

    Dim tmpStr As String
    
    On Error GoTo Err_Trap
    
    If Item.Text = "" Then Exit Sub
    txtPtId.SetFocus
    DoEvents
    txtPtId.Text = MedGetPtid(Item.Text)
    Call txtPtId_KeyPress(vbKeyReturn)
    Exit Sub
Err_Trap:
     Resume Next

End Sub


Private Sub optQueryKey_Click(Index As Integer)
    lblKeyDate.Caption = optQueryKey(Index).Caption
End Sub

Private Sub optQueryKey_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub picResult_Resize()
    rtfResult.Width = picResult.Width - 50
End Sub

Private Sub rtfResult_DblClick()

    Dim sLabNo()
    Dim strTag As String
    Dim strLabNo As String
    Dim aryLabNo As Variant
    
    Screen.MousePointer = vbArrowHourglass
    DoEvents
    
    strTag = rtfResult.Tag
    strLabNo = medGetP(strTag, 1, COL_DIV)
    aryLabNo = Split(strLabNo, "-")
    If aryLabNo(3) = BBS_ORDDIV Then Exit Sub
    
'    Set objTextForm = frmAPS905
'        .fraTextResult.Visible = False
    frmAPS905.rtfResultText.Visible = True
    frmAPS905.OrdDiv = aryLabNo(3)
    If aryLabNo(3) = APS_ORDDIV Then
        Call frmAPS905.GetResultText(aryLabNo(0), aryLabNo(1), aryLabNo(2))
    ElseIf aryLabNo(3) = LIS_ORDDIV Then
        frmAPS905.Caption = medGetP(strTag, 2, COL_DIV)
        frmAPS905.rtfResultText.TextRTF = rtfResult.TextRTF
    End If
'    frmAPS905.Top = Me.Top
'    frmAPS905.Left = Me.Left + 7000
    
    Screen.MousePointer = vbDefault

    DoEvents
    
    frmAPS905.WindowState = 0
    frmAPS905.Show vbModal

End Sub

'% ó�� ����(Click)�ϸ� �ش� ��� ���÷���...
Private Sub tblOrdSheet_Click(ByVal Col As Long, ByVal Row As Long)
   
    Dim pWorkArea As String
    Dim pAccDt As String
    Dim pAccSeq As String
    Dim pTestDiv As String
    Dim strOrdDiv As String
    Dim strWardId As String
    Dim strHosilId As String
    Dim tmpResult As New clsLISResultReview
    Dim strRoomId As String '�߰�����(����)
     
    If Row = 0 Then Exit Sub
    If OldRow = Row Then Exit Sub
    
    With tblOrdSheet
      
        .Row = Row
        .Col = 3:  If .Value = "" Then Exit Sub
        
        .Col = 17: pWorkArea = .Value
        .Col = 18: pAccDt = .Value
        .Col = 19: pAccSeq = .Value
        .Col = 30: pTestDiv = .Value
        .Col = 38: strOrdDiv = .Value
        .Col = 42: strWardId = .Value
        .Col = 43: strHosilId = .Value
        .Col = 44: strRoomId = .Value   '�߰�����
        
        If strWardId <> "" Then
            lblLocation.Caption = strWardId & " - " & strHosilId
            If Trim(strRoomId) <> "" Then lblLocation.Caption = lblLocation.Caption & " - " & strRoomId
        Else
            lblLocation.Caption = ""
        End If
      
        If (pWorkArea = "" Or pAccDt = "" Or pAccSeq = "") And strOrdDiv <> BBS_ORDDIV And strOrdDiv <> POC_ORDDIV And strOrdDiv <> CMT_ORDDIV Then
            MsgBox "������ȣ�� �����ϴ�. (����Ƿ� �����ٶ� ��" & ObjSysInfo.HelpLine & ")", vbExclamation, "�����߻�"
            Exit Sub
        End If
      
        If OldRow > 0 Then
            .Row = OldRow
            .Col = -1
            .BackColor = OldBackColor
        End If
         
        .Row = Row
        .Col = -1
        OldRow = Row
        OldBackColor = .BackColor
        .BackColor = &HD9ECFF ' &HFCEFE9   ' &HF5FFF4       '���λ�
      
      
        .Col = 8: '����ġ
        If Trim(.Value) = CS_QuestionMark Then Call GetRefValue(Row)
        DoEvents
        
        .Col = 23:  lblVerifyDt.Caption = .Value                        '�����Ͻ�
        .Col = 24:  lblVerifierNm.Caption = tmpResult.GetEmpNm(.Value)  '������
        .Col = 31:  lblDeptNm.Caption = tmpResult.GetDeptNm(.Value)     '�����
        Set tmpResult = Nothing
        
        Call ResultClear
        cmdReport.Enabled = True
        .Col = 33:   lblSpecimenNm.Caption = .Value '��ü
        
        tblResult.ReDraw = False
        
        MouseRunning
        
        Select Case strOrdDiv
        Case APS_ORDDIV
            'fraLisResult.Visible = False
            'rtfResult.Visible = True
            rtfResult.Tag = pWorkArea & "-" & pAccDt & "-" & pAccSeq & "-" & strOrdDiv
            rtfResult.Text = ""
            DoEvents
            Call DisplayAPSResult(pWorkArea, pAccDt, Val(pAccSeq))
        Case BBS_ORDDIV
            Screen.MousePointer = vbArrowHourglass
            fraLisResult.Visible = False
            tblResult.Visible = False
            picFootNote.Visible = False
            rtfResult.Visible = True
            rtfResult.Tag = pWorkArea & "-" & pAccDt & "-" & pAccSeq & "-" & strOrdDiv
            DoEvents
            Call DisplayBBSResult(pWorkArea, pAccDt, Val(pAccSeq), Row)
            Screen.MousePointer = vbDefault
        Case LIS_ORDDIV
            Screen.MousePointer = vbArrowHourglass

            rtfResult.Tag = pWorkArea & "-" & pAccDt & "-" & pAccSeq & "-" & strOrdDiv
            fraLisResult.Visible = True
            tblResult.Visible = True
            picFootNote.Visible = True
            rtfResult.Visible = False
            DoEvents
            Call DisplayLISResult(pWorkArea, pAccDt, Val(pAccSeq), pTestDiv)
            
            Screen.MousePointer = vbDefault
        Case CMT_ORDDIV
            
            Call DisplayLABCommrnt(Row)

        End Select
        
        tblResult.ReDraw = True
        MouseDefault
        
        'Debug.Print "Show :", ",", Now
        
        tblResult.TopRow = 1
        ResultFg = True
      
    End With
    
End Sub

'% Lab No.�� �������� �˻��� ��������� ���̺� Display�Ѵ�.
Private Sub DisplayAPSResult(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Integer)

If P_IncludeAPSSystem Then

    Dim I As Integer, j As Integer
    Dim ResultBuffer As String
    Dim RstTxtBuffer As String
    Dim SamTxtBuffer As String
    Dim strWAccDt As String
    Dim strAccSeq As String
    Dim rs As New DrRecordSet
    Dim objResult As New clsAPSResult
    Dim objSql As New clsAPSSqlResult
    Dim strRsEntryType  As String
    
    With objResult
    
        strWAccDt = Trim(pWorkArea) & Trim(Mid(pAccDt, 3, 2))
        strAccSeq = Trim(Format(pAccSeq, "00000#"))
        
        Call .LoadResult(strWAccDt, strAccSeq, , False, False)
        
        strRsEntryType = .RstEntryType
        
        If strRsEntryType = "" Then Exit Sub
        
        If .stscd < "6" Then Exit Sub   ' �ǵ�
        
'        Call .LoadResult(strWAccDt, strAccSeq, strRsEntryType)
'
'        ObjLISComCode.PTHDOCT.Exists (.PTHDOCT)
'        If ObjLISComCode.PTHDOCT.Exists(.PTHDOCT) = True Then
'            ObjLISComCode.PTHDOCT.KeyChange .PTHDOCT
'            lblVerifierNm.Caption = ObjLISComCode.PTHDOCT.Fields("pthdoctnm")   'Ȯ����
'        Else
'            lblVerifierNm.Caption = ""
'        End If
'
'        lblDeptNm.Caption = .DeptCdNm
'
'        '��� ��ȸ
'        Call LoadResultText(.WorkArea, .AccDt, .AccSeq)
'        rtfResult.Visible = True
'        DoEvents
        Call rtfResult_DblClick
        DoEvents

    End With

End If

End Sub

Private Sub LoadResultText(ByVal pWorkArea As String, ByVal pAccDt As String, _
                           ByVal pAccSeq As String)
    
    If P_IncludeAPSSystem Then
        Dim objText As clsAPSScreenResult
    
        Set objText = New clsAPSScreenResult
        'objText.setDbConn DBConn
        
        Call objText.LoadScreenResult(pWorkArea, pAccDt, pAccSeq, rtfResult)
    
        Set objText = Nothing
    End If
    
End Sub

'% Lab No.�� �������� �˻��� ��������� ���̺� Display�Ѵ�.
Private Sub DisplayBBSResult(ByVal pWorkArea As String, ByVal pAccDt As String, _
                             ByVal pAccSeq As Integer, ByVal iRow As Long)

        Dim strTransResult As String
        Dim strUnitQty As String
        Dim strReqDtTm As String
        Dim strReason As String
        Dim strOrdDt As String
        Dim strOrdNo As String
        Dim lngAssignCnt As Long
        Dim lngDeliveryCnt As Long
        Dim ObjABO As New clsABO
        Dim objTransReason As New clsQueryOrder
        Dim objA As New clsGetSqlStatement
        Dim objRs As New DrRecordSet
        
        With tblOrdSheet
            .Row = iRow
            .Col = 39: strUnitQty = .Value
            .Col = 40: strReqDtTm = Format(.Value, CS_DateMask)
            .Col = 41: strReqDtTm = strReqDtTm & " " & Format(Mid(.Value, 1, 4), CS_TimeShortMask)
            .Col = 10: strOrdDt = Format(.Value, CS_DateDbFormat)
            .Col = 11: strOrdNo = .Value
        End With
        
        strReason = objTransReason.GetTransReason( _
                    MedSetPtid(txtPtId.Text), strOrdDt, strOrdNo)
        Set objTransReason = Nothing
        
        With objA
    '        .setDbConn DbConn
            Set objRs = OpenRecordSet(.Order_Status_LIst(strOrdDt, strOrdDt, False, "", MedSetPtid(txtPtId.Text)))
            '�������ϱ�
            lngAssignCnt = Val("" & objRs.Fields("assigncnt").Value) - Val("" & objRs.Fields("assigncancelcnt").Value)
            lngDeliveryCnt = Val("" & objRs.Fields("deliverycnt").Value)
            objRs.RsClose
            Set objRs = Nothing
        End With
        
        ObjABO.ptid = MedSetPtid(txtPtId.Text)  '�������� ������.
        ObjABO.GetABO
        With rtfResult
            .Visible = False
            .Text = vbCrLf & Space(13) & "�� ���� �����Ȳ ��" & vbCrLf & vbCrLf
            .Text = .Text & Space(3) & "�� �� �� �� : " & ObjABO.ABO & ObjABO.Rh & vbCrLf & vbCrLf
            .Text = .Text & Space(3) & "�� �� �� �� : " & strReqDtTm & vbCrLf & vbCrLf
            .Text = .Text & Space(3) & "�� �������� : " & strReason & vbCrLf & vbCrLf
            .Text = .Text & Space(3) & "�� ��    �� : " & strUnitQty & vbCrLf & vbCrLf
            .Text = .Text & Space(3) & "�� Assign   : " & lngAssignCnt & vbCrLf & vbCrLf
            .Text = .Text & Space(3) & "�� ������ : " & lngDeliveryCnt & vbCrLf & vbCrLf
            
            .SelStart = 15: .SelLength = Len(.Text)
            .SelFontName = "����"
            .SelFontSize = 13
            .SelBold = True
            
            .SelStart = 30: .SelLength = Len(.Text)
            .SelFontName = "����ü"
            .SelFontSize = 10
            .SelBold = True
            '.SelColor = &H553755 &HE48372 '�ణ �Ķ���
            
            Call HighlightText(rtfResult, "�� ���� �����Ȳ ��", True, , &H4A4189)
            Call HighlightText(rtfResult, "�� �� �� �� :", False, , &H553755)
            Call HighlightText(rtfResult, ObjABO.ABO & ObjABO.Rh, False, , &H7477EF, 15) '�ణ ������
            Call HighlightText(rtfResult, "�� �� �� �� :", False, , &H553755)
            Call HighlightText(rtfResult, strReqDtTm, False, , &HE48372)
            Call HighlightText(rtfResult, "�� �������� :", False, , &H553755)
            Call HighlightText(rtfResult, "�� ��    �� :", False, , &H553755)
            Call HighlightText(rtfResult, "�� Assign   :", False, , &H553755)
            Call HighlightText(rtfResult, "�� ������ :", False, , &H553755)
            
            .Visible = True
        
        End With
        
        Set ObjABO = Nothing
    
End Sub

'% Lab No.�� �������� �˻��� ��������� ���̺� Display�Ѵ�.
Private Sub DisplayLISResult(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Integer, _
                          ByVal pTestDiv As String, Optional pQuery As Boolean = True)
   
    Dim I As Integer, j As Integer
    Dim MyResult As New clsLISResultReview
    Dim ResultBuffer As String
    Dim RstTxtBuffer As String
    Dim SamTxtBuffer As String
   
    With MyResult
      
        MouseRunning
        
        Call .ResultMore(pWorkArea, pAccDt, pAccSeq, pTestDiv)
      
        If .ResultCnt = 0 Then
            MouseDefault
            Exit Sub
        End If
      
        lblDeptNm.Caption = .DeptNm
'      lblLocation.Caption = .WardId & "-" & .HosilID
      
        For I = 1 To .RstRow
            tblResult.Row = I + .OffSet
            For j = 1 To 8
                tblResult.Col = j
                If .Get_ForeColor(j, I) <> 0 Then tblResult.ForeColor = .Get_ForeColor(j, I)
            Next
        Next
      
        '������� Display
        tblResult.Row = 1
        tblResult.Row2 = tblResult.MaxRows
        tblResult.Col = 2
        tblResult.Col2 = tblResult.MaxCols
        tblResult.BlockMode = True
        tblResult.AllowCellOverflow = True
        tblResult.Clip = .ResultClipText    '& .SenClipText             'ResultBuffer
        tblResult.BlockMode = False
      
        '�̻��� ������ ����� ��� �׻����� ������ Sort / Align Left
        'If .SortFg Then
        If .SortFg Then
            For I = 1 To .SensiCount
                tblResult.SortBy = SortByRow
                tblResult.SortKey(1) = 2  '�׻�����
                tblResult.SortKeyOrder(1) = SortKeyOrderAscending
                tblResult.Col = -1
                tblResult.Row = .AntiSortStartRow(I)   '+ .OffSet
                tblResult.Row2 = .AntiSortEndRow(I)    '+ .OffSet
                tblResult.Action = ActionSort
                tblResult.Row = .SortStartRow - 1 '+ .OffSet
                tblResult.Col = 2
                tblResult.FontUnderline = True
            Next
        Else
            tblResult.Col = 6
            tblResult.Row = -1
            tblResult.ForeColor = DCM_LightRed
            tblResult.FontBold = True
        End If
        If .TestDiv = TST_MicTest Then
            '�̻��� ��� : �ո��÷� Align Left
            tblResult.Row = -1
            tblResult.Col = -1
            tblResult.BlockMode = True
            tblResult.AllowCellOverflow = True
            tblResult.TypeHAlign = TypeHAlignLeft
            tblResult.BlockMode = False
            tblResult.ColWidth(2) = 17
            'tblResult.ColWidth(3) = 60
            For I = 1 To 5
                If .MicFg(I) Then
                    tblResult.ColWidth(I + 2) = 9
                Else
                    tblResult.ColWidth(I + 2) = 4
                End If
            Next
            tblResult.ColWidth(8) = 20
            tblResult.Col = 3: tblResult.Col2 = 7
            tblResult.Row = -1
            tblResult.BlockMode = True
            tblResult.FontBold = False
            tblResult.BlockMode = False
        Else
            '�Ϲݰ�� : ����÷� Align Center
            tblResult.Row = 1: tblResult.Row2 = tblResult.MaxRows
            tblResult.Col = 3: tblResult.Col2 = 7
            tblResult.BlockMode = True
            tblResult.TypeHAlign = TypeHAlignCenter
            tblResult.BlockMode = False
            tblResult.ColWidth(2) = 13
            tblResult.ColWidth(3) = 9
            tblResult.ColWidth(4) = 9
            tblResult.ColWidth(5) = 3
            tblResult.ColWidth(6) = 5
            tblResult.ColWidth(7) = 13
        End If
      
       '�ؽ�Ʈ��� Display
    
       'If .TextFg Then
       '   txtRstCmt.Text = .RstTextBuffer        'RstTxtBuffer
       '   txtRstCmt1.Text = .RstTextBuffer        'RstTxtBuffer
       '   chkRstCmt.Value = 1
       '   chkRstCmt.Enabled = True
       '   Call HighlightText(txtRstCmt, "<< �˻� �Ұ� >>", True)
       '   Call HighlightText(txtRstCmt, "<< Supplemental Report >>", False)
       '   Call HighlightText(txtRstCmt1, "<< �˻� �Ұ� >>", True)
       '   Call HighlightText(txtRstCmt1, "<< Supplemental Report >>", False)
       'Else
       '   chkRstCmt.Value = 0
       '   chkRstCmt.Enabled = False
       'End If
       
        '��ü����ũ & ǲ��Ʈ Display
        If .CommentFg Then
            txtSamCmt.Text = .SamTextBuffer
            'txtSamCmt1.Text = .SamTextBuffer
            chkSamCmt.Value = 1
            chkSamCmt.Enabled = True
            Call HighlightText(txtSamCmt, "<< Remark >>", True)
            Call HighlightText(txtSamCmt, "<< Foot Note >>", False)
            'Call HighlightText(txtSamCmt1, "<< Remark >>", True)
            'Call HighlightText(txtSamCmt1, "<< Foot Note >>", False)
        Else
            chkSamCmt.Value = 0
            chkSamCmt.Enabled = False
            picFootNote.Visible = False
        End If
      
        'Ư���˻� ��� Display
        If .SpecialFg Then
            rtfResult.TextRTF = .SpeTextBuffer
            rtfResult.Tag = rtfResult.Tag & COL_DIV & .SpeRstTitle
            Call rtfResult_DblClick
        End If
        
    End With
   
   
    With tblResult
        .Col = 2: .Col2 = 5 '.MaxCols
        .Row = 1: .Row2 = .MaxRows
        .BlockMode = True
        txtRstCmt1.Text = .Clip
        .BlockMode = False
    End With
    Call HighlightText(txtRstCmt1, "<< �˻� �Ұ� >>", True)
    Call HighlightText(txtRstCmt1, "<< Supplemental Report >>", False)
    Call HighlightText(txtRstCmt1, "[ Susceptibility test ]", False)
    Call HighlightText(txtRstCmt1, "Antibiotics", False, , &HDF6A3E)
    Call HighlightText(txtRstCmt1, "1      ", False, , &HDF6A3E)
    Call HighlightText(txtRstCmt1, "2      ", False, , &HDF6A3E)
    Call HighlightText(txtRstCmt1, "3      ", False, , &HDF6A3E)
   
    MouseDefault
   
End Sub


'% �� �ε�
Private Sub Form_Load()
'    Me.Show
    txtSearchKey.Text = ""
    chkPtList.Value = 0
    Call chkPtList_Click
    OrderFg = False
    ResultFg = False
    ClearFg = True
    PtFg = False
    OldRow = 0
    medInitLvwHead lvwPtList, "ȯ��ID,ȯ�ڼ���,�ֹε�Ϲ�ȣ,�������,����/����", _
                       "50,50,800,300,100"
   
    TopLeftShow = False
    optSort(1).Value = True
    
    If gUsingInWardMenu Then
'        dtpFromDate.Value = DateAdd("d", -2, Now)
        If P_AllResultReview Then
            dtpFromDate.Value = Format(P_ReviewStartDate, CS_DateLongMask)
        Else
            dtpFromDate.Value = DateAdd("m", -2, Now)
        End If
        optQueryKey(1).Value = True
    Else
        dtpFromDate.Value = DateAdd("d", -4, Now)
        optQueryKey(0).Value = True
    End If
    dtpToDate.Value = Now
    
    'Set MyPatient.MyOraSE = OraSe
    Set MyPatient.objDb = DBConn
    cmdReport.Enabled = False
    
    On Error GoTo Err_Trap
Err_Trap:
End Sub


'% ���� ���� ����
Private Sub optSort_Click(Index As Integer)
    If Not picPtList.Visible Then Exit Sub
    If txtSearchKey.Text <> "" Then
        Call txtSearchKey_KeyPress(vbKeyReturn)
    End If
    txtSearchKey.SetFocus
End Sub


'% ó�����̺� Set Focus
Private Sub tblOrdSheet_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo Err_Trap
    If OrderFg Then tblOrdSheet.SetFocus
Err_Trap:
End Sub

'ó�泻�� ���̺� ToolTip �����ֱ�...
Private Sub tblOrdSheet_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)

    Dim tmpToolTip As String
    Dim MyResult As New clsLISResultReview
    Dim strSql As String
    Dim rsMod As DrRecordSet
    Dim tmpColNm As String
   
    If Not OrderFg Then Exit Sub
   
    tmpToolTip = vbCrLf
   
    With tblOrdSheet
        .Row = Row
       
        If Col = 4 Then
            .Col = 4
            If Len(.Value) > 20 Then
                MultiLine = 1
                TipWidth = 4000
                tmpToolTip = vbCrLf & Space(3) & .Value & Space(3) & vbCrLf
                TipText = tmpToolTip
                ShowTip = True
                Exit Sub
            End If
        End If
       
        .Col = 3:    If Trim(.Value) = "" Then Exit Sub
       
        If chkToolTip.Value = 0 Then GoTo Skip
       
        .Col = 10:  tmpToolTip = tmpToolTip & "  ó    �� : " & .Value       'ó����
        .Col = 12:  tmpToolTip = tmpToolTip & "  by  " & MyResult.GetDoctNm(.Value)                 'ó����
        .Col = 11:  tmpToolTip = tmpToolTip & " ( # " & Format(.Value, "##") & " )" & vbCrLf        'ó���ȣ
        .Col = 13:  tmpToolTip = tmpToolTip & "  ä    �� : " & .Value       'ä���Ͻ�
        .Col = 14:  tmpColNm = MyResult.GetDoctNm(.Value)       'ä����-��ȣ��
                    If Trim(tmpColNm) = "" Then
                        tmpColNm = MyResult.GetEmpNm(.Value)    'ä����-������
                    End If
                    tmpToolTip = tmpToolTip & "  by  " & tmpColNm & vbCrLf   'ä����
        .Col = 15:  tmpToolTip = tmpToolTip & "  ��    �� : " & .Value       '�����Ͻ�
        .Col = 16:  tmpToolTip = tmpToolTip & "  by  " & MyResult.GetEmpNm(.Value) & vbCrLf   '������
        .Col = 4:
                    If .Value <> "��Ȯ" Then
                        .Col = 23:   tmpToolTip = tmpToolTip & "  ������� : " & .Value       '�����Ͻ�
                        .Col = 24:   tmpToolTip = tmpToolTip & "  by  " & MyResult.GetEmpNm(.Value) & vbCrLf   '������
                    End If
       .Col = 20:
                    If .Value <> "" Then
                        tmpToolTip = tmpToolTip & vbCrLf & "  �ֱٰ�� : [ " & .Value & " ] " '& vbCrLf        '�ֱٰ��
                        '.Col = 21:   tmpToolTip = tmpToolTip & "             " & .Value  '�ֱٰ���Ͻ�
                        .Col = 21
                        tmpToolTip = tmpToolTip & Mid(.Value, 1, 9) '�ֱٰ���Ͻ�
                        .Col = 22
                        tmpToolTip = tmpToolTip & "  by " & MyResult.GetEmpNm(.Value) & vbCrLf  '�ֱٰ�� ������
                    End If
       '������ ���...
       .Col = 38:
                    Dim strModRst As String
                    Dim pWorkArea As String
                    Dim pAccDt As String
                    Dim pAccSeq As String
                  
                    .Col = 17: pWorkArea = .Value
                    .Col = 18: pAccDt = .Value
                    .Col = 19: pAccSeq = .Value
                  
                    .Col = 27
                    strSql = MySql.SqlGetOldResult(pWorkArea, pAccDt, pAccSeq, .Value)
                    Set rsMod = OpenRecordSet(strSql)
                    If Not rsMod.EOF Then
                        tmpToolTip = tmpToolTip & vbCrLf & "  ��������� : " & vbCrLf
                        'While (Not rsMod.EOF)
                            strModRst = "             [ " & Trim(rsMod.Fields("RstCd").Value) & " ] "
                            strModRst = strModRst & Format(Mid(rsMod.Fields("vfydt").Value, 3, 6), "0#-##-##") & Space(2)
                            strModRst = strModRst & " by " & rsMod.Fields("EmpNm").Value & vbCrLf
                            tmpToolTip = tmpToolTip & strModRst
                        '    rsMod.MoveNext
                        'Wend
                    End If
                    rsMod.RsClose
                    Set rsMod = Nothing
       
Skip:
        If UBound(aryMesg) >= Row Then
           If aryMesg(Row) <> "" Then tmpToolTip = tmpToolTip & vbCrLf & "  " & aryMesg(Row) & vbCrLf
        End If
     
Skip1:
        MultiLine = 1
        TipText = tmpToolTip
        TipWidth = 5000
        .TextTipDelay = 500
        Call .SetTextTipAppearance("����ü", 9, False, False, &HEEFDF2, &H996666)
        If chkToolTip.Value = 1 Then
            ShowTip = True
        Else
            ShowTip = False
        End If
       
    End With
   
End Sub

Private Sub tblOrdSheet_TopLeftChange(ByVal OldLeft As Long, ByVal OldTop As Long, ByVal NewLeft As Long, ByVal NewTop As Long)
   
    Dim tmpStr As Variant
   
    With tblOrdSheet
        If Not TopLeftShow1 Then
            .Row = OldTop
            '.Col = 1:  .ForeColor = .BackColor
            .Col = 1:  .Value = "" ' .ForeColor = .BackColor
        End If
        If Not TopLeftShow2 Then
            .Row = OldTop
            '.Col = 2:  .ForeColor = .BackColor
            .Col = 2:  .Value = ""  '.ForeColor = .BackColor
        End If
        
        .Row = NewTop
        .Col = 1:
        Call .GetText(32, NewTop, tmpStr)
        'If .ForeColor <> .BackColor Then
        If .Value = tmpStr Then
            TopLeftShow1 = True
        Else
            TopLeftShow1 = False
            '.Col = 1:  .ForeColor = vbBlack
            .Col = 1:  .Value = tmpStr
        End If
        .Col = 2:
        Call .GetText(33, NewTop, tmpStr)
        'If .ForeColor <> .BackColor Then
        If .Value = tmpStr Then
            TopLeftShow2 = True
        Else
            TopLeftShow2 = False
            '.Col = 2:  .ForeColor = &H7477EF   '�ణ ������
            .Col = 2:  .Value = tmpStr
        End If
    End With
   
End Sub

Private Sub tblResult_DblClick(ByVal Col As Long, ByVal Row As Long)
    fraTextResult.Visible = True
    fraTextResult.ZOrder 0

End Sub

'% ������̺� Set Focus
Private Sub tblResult_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo Err_Trap
    If ResultFg Then tblResult.SetFocus
Err_Trap:
End Sub

'������� ���̺� ToolTip �����ֱ�...
Private Sub tblResult_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)

    Dim tmpToolTip As String
    
    If Not ResultFg Then Exit Sub
    
    tmpToolTip = vbCrLf
   
    With tblResult
    
        .Row = Row
        .Col = 2:
                   If .Value = "" Then
                      ShowTip = False
                      Exit Sub
                   End If
        .Col = 8:  tmpToolTip = tmpToolTip & "  " & .Value & vbCrLf   'ó���(Long)
        .Col = 9:  If .Value = "" Then GoTo Skip
                   tmpToolTip = tmpToolTip & vbCrLf & "  �ֱٰ�� : " & .Value & vbCrLf   '�ֱٰ��
        .Col = 10: tmpToolTip = tmpToolTip & "  �����Ͻ� : " & .Value & vbCrLf  '�ֱٰ����
      
Skip:
        MultiLine = 1
        If Trim(Replace(tmpToolTip, vbCrLf, "", 1, -1, vbBinaryCompare)) = "" Then
          ShowTip = False
          Exit Sub
        End If
        TipText = tmpToolTip
        TipWidth = 4000
        .TextTipDelay = 1000
        Call .SetTextTipAppearance("����ü", 9, False, False, &HEEFDF2, &H996666)
        ShowTip = True
    End With
   
End Sub

'% ȯ��ID�� ����Ǹ� ȭ��Clear
Private Sub txtPtId_Change()
    If Not ClearFg Then
        Call ClearRtn
    End If
    StopFg = True
   Dim lngLen As Long
    
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
'��굿�������ΰ�� �⵵(2�ڸ�) "-" ��ȣ(6�ڸ�)
    
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
        optQueryKey(0).SetFocus
    End If
End Sub


'% �ؽ�Ʈ��� �ڽ� ����Ŭ�� - Larger Box Popup
Private Sub txtRstCmt_DblClick()
    fraTextResult.Visible = True
    fraTextResult.ZOrder 0
End Sub

Private Sub txtPtId_LostFocus()
    
'    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    If ActiveControl.Name = cmdExit.Name Then Exit Sub
    If ActiveControl.Name = chkPtList.Name Then Exit Sub
    If MsgFg Then Exit Sub
      
    On Error GoTo Err_Trap
    
    If txtPtId.Text = "" Then
        'txtPtId.SetFocus
        Exit Sub
    End If
    Select Case PROJECT_HOSCD
        Case "04": txtPtId = MedGetPtid(txtPtId)
        Case Else
            If IsNumeric(txtPtId.Text) Then
                txtPtId.Text = Format(txtPtId.Text, P_PatientIdFormat)
            End If
    End Select
    Dim strWardId As String
    
    With MyPatient
        If .PtntQuery(MedSetPtid(txtPtId.Text)) Then
            lblPtNm.Caption = .PtNm
            lblSex.Caption = .SexNm
            lblAge.Caption = .Age
            lblAgeDiv.Caption = .AgeDiv
            lblDeptNm.Caption = .DeptNm
            'lblBedinDt.Caption = Format(.BedInDt, CS_DateMask)
            'lblBedoutDt.Caption = Format(.BedOutDt, CS_DateMask)
            strWardId = .WardId
            If strWardId <> "" Then
                If .RoomId <> "" Then strWardId = strWardId & "-" & .RoomId
'                If gUsingInWardMenu Then
'                    If P_AllResultReview Then
'                        dtpFromDate.Value = Format(P_ReviewStartDate, CS_DateLongMask)
'                    Else
'                        dtpFromDate.Value = DateAdd("d", -2, Now)
'                    End If
'                    'optQueryKey(2).Value = True
'                End If
            Else
'                If gUsingInWardMenu Then
'                    If P_AllResultReview Then
'                        dtpFromDate.Value = Format(P_ReviewStartDate, CS_DateLongMask)
'                    Else
'                        dtpFromDate.Value = DateAdd("m", -2, Now)
'                    End If
'                    'optQueryKey(2).Value = True
'                End If
            End If
            lblLocation.Caption = strWardId
            If .BedOutDt <> "" Then
                Dim strTmp1 As String
'            '�ֱ��� ó����� ������ �´�.
                Select Case PROJECT_HOSCD
                    Case "02"
                        strTmp1 = MySql.GetDeptInfo(MedSetPtid(txtPtId.Text))
                        If strTmp1 <> "" Then
                            lblLocation.Caption = ""
                            lblDeptNm.Caption = medGetP(strTmp1, 1, COL_DIV)
                            'lblDoctNm.Caption = medGetP(strTmp1, 2, COL_DIV)
                        End If
                    Case Else
                End Select
            End If
            Dim objDisease  As New S2LIS_ReportLib.clsDisease
            objDisease.ptid = MedSetPtid(txtPtId.Text)
            lblDisease.Caption = objDisease.Disease
            Set objDisease = Nothing
                
            gPatientId = MedSetPtid(txtPtId.Text)
            PtFg = True
        Else
            MsgFg = True
            MsgBox "��ϵ��� ���� ȯ��ID�Դϴ�.. �ٽ� �Է��ϼ���.."
            MsgFg = False
            Me.Enabled = True
            txtPtId.SetFocus
            PtFg = False
            Call txtPtId_GotFocus
            Exit Sub
        End If
    End With
    StopFg = False

On Error GoTo Err_Trap
    If ActiveControl.Name <> cmdRefresh.Name Then
       If dtpFromDate.Enabled Then dtpFromDate.SetFocus
    End If
    If ClearFg Then Call cmdRefresh_Click
    ClearFg = False
    Exit Sub
Err_Trap:
    Resume Next

End Sub

'% �ؽ�Ʈ��� �ڽ�1 ����Ŭ�� - Invisible
Private Sub txtRstCmt1_DblClick()
    fraTextResult.Visible = False
End Sub

'% ǲ��Ʈ �ڽ� ����Ŭ�� - Larger Box Popup
Private Sub txtSamCmt_DblClick()
    fraTextResult.Visible = True
    fraTextResult.ZOrder 0
End Sub

'% ǲ��Ʈ �ڽ�1 ����Ŭ�� -Invisible
Private Sub txtSamCmt1_DblClick()
   'fraTextResult.Visible = False
End Sub

'% Popup Frame ����Ŭ�� - Invisible
Private Sub fraTextResult_DblClick()
   fraTextResult.Visible = False
End Sub


Private Sub txtSearchKey_Change()
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
End Sub

'% ȯ�� �˻� (ID �Ǵ� ��������...)
Private Sub txtSearchKey_GotFocus()

    With txtSearchKey
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

'% ȯ��ID �Ǵ� �������� �˻� ����Ʈ �ۼ�.
Private Sub txtSearchKey_KeyPress(KeyAscii As Integer)
    
'    Dim objPtInfo As New clsHosComSQLStmt
'    Dim DrRs As New DrRecordSet
'    Dim itmx As ListItem
'    Dim lngSearch As Long
'    Dim ColCnt As Long
'    Dim RowCnt As Long
'
'    If PROJECT_HOSCD = "04" Then
'        If Len(txtPtId) <> 2 Then
'            If KeyAscii = vbKeyInsert Then KeyAscii = 0
'        End If
'
'        If KeyAscii = vbKeyBack Then
'            With txtPtId
'                If .Text = "" Then Exit Sub
'                If Mid(.Text, Len(.Text)) = "-" Then
'                    .Text = Mid(.Text, 1, Len(.Text) - 2)
'                    .SelStart = Len(.Text)
'                    KeyAscii = 0
'                End If
'            End With
'        End If
'    End If
'
'    If KeyAscii = vbKeyReturn Then
'        lngSearch = IIf(optSort(0).Value, 1, 2)  'True:ȯ��ID, False:ȯ�ڸ�
'
'        If lngSearch = 1 And Not IsNumeric(MedSetPtid(txtSearchKey.Text)) Then Exit Sub
'
'        If chkVerified.Value = 0 Then
'            If lngSearch = 2 And Len(txtSearchKey.Text) < 2 Then
'                MsgBox "2���� �̻� �Է��Ͻ��� �˻��Ͻʽÿ�.", vbInformation, "ȯ�ڰ˻�"
'                txtSearchKey.SetFocus
'                Exit Sub
'            End If
'            If lngSearch = 1 Then
'                ColCnt = DrRs.OpenCursor(, objPtInfo.SqlPtntSearch(lngSearch, MedSetPtid(txtSearchKey)))
'            Else
'                ColCnt = DrRs.OpenCursor(, objPtInfo.SqlPtntSearch(lngSearch, txtSearchKey))
'            End If
'        Else
'            If lngSearch = 1 Then
'                ColCnt = DrRs.OpenCursor(, objPtInfo.SqlPtntSearch(lngSearch, MedSetPtid(txtSearchKey), _
'                              mvarDeptCd, Format(DBConn.GetSysDate, CS_DateDbFormat)))
'            Else
'                ColCnt = DrRs.OpenCursor(, objPtInfo.SqlPtntSearch(lngSearch, txtSearchKey, _
'                              mvarDeptCd, Format(DBConn.GetSysDate, CS_DateDbFormat)))
'            End If
'        End If
'        lvwPtList.ListItems.Clear
'        If ColCnt > 0 Then
'            RowCnt = 0
'            With lvwPtList
'                Do While (DrRs.FetchCursor(ColCnt))
'                    RowCnt = RowCnt + 1
'                    Set itmx = .ListItems.Add(, , MedGetPtid("" & DrRs.GetValue("ptid")))
'                    itmx.SubItems(1) = "" & DrRs.GetValue("ptnm")
'                    itmx.SubItems(2) = "" & DrRs.GetValue("SSN")
'                    itmx.SubItems(3) = "" & DrRs.GetValue("DOB")
'                    If Not IsDate(itmx.SubItems(3)) Then
'                        itmx.SubItems(3) = Mid(itmx.SubItems(3), 1, 4) & "-01-01"
'                    End If
'                    If IsNumeric(Mid("" & DrRs.GetValue("ssn"), 8, 1)) Then
'                        itmx.SubItems(4) = IIf((Mid("" & DrRs.GetValue("ssn"), 8, 1) Mod 2) = 1, "��", "��")
'                    Else
'                        itmx.SubItems(4) = "��"
'                    End If
'                    If IsDate(itmx.SubItems(3)) Then
'                        itmx.SubItems(4) = itmx.SubItems(4) & " / " & DateDiff("yyyy", itmx.SubItems(3), Now)
'                    Else
'                        itmx.SubItems(4) = itmx.SubItems(4) & " / ? "
'                    End If
'                    If RowCnt > 1000 Then Exit Do
'                Loop
'            End With
'        Else
'            MsgBox "���ǿ� �´� �ڷᰡ �����ϴ�. Ȯ���� �˻��ϼ���", vbInformation + vbOKOnly, Me.Caption
'        End If
'        DrRs.CloseCursor:     Set DrRs = Nothing
'
'    End If
'
'    Set objPtInfo = Nothing
        
    Dim objPtInfo As New clsHosComSQLStmt
    Dim DrRs As New DrRecordSet
    Dim itmx As ListItem
    Dim lngSearch As Long
    Dim ColCnt As Long
    Dim RowCnt As Long
    
    
    If PROJECT_HOSCD = "04" Then
        
        If Len(txtPtId) <> 2 Then
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
    
    If KeyAscii = vbKeyReturn Then
        lngSearch = IIf(optSort(0).Value, 1, 2)  'True:ȯ��ID, False:ȯ�ڸ�
        
        If lngSearch = 1 And Not IsNumeric(MedSetPtid(txtSearchKey.Text)) Then Exit Sub
        
        If chkVerified.Value = 0 Then
            If lngSearch = 2 And Len(txtSearchKey.Text) < 2 Then
                MsgBox "2���� �̻� �Է��Ͻ��� �˻��Ͻʽÿ�.", vbInformation, "ȯ�ڰ˻�"
                txtSearchKey.SetFocus
                Exit Sub
            End If
            If optSort(0).Value = True Then
                ColCnt = DrRs.OpenCursor(, objPtInfo.SqlPtntSearch(lngSearch, MedSetPtid(txtSearchKey)))
            Else
                ColCnt = DrRs.OpenCursor(, objPtInfo.SqlPtntSearch(lngSearch, txtSearchKey))
            End If
        Else
            If optSort(0).Value = True Then
                ColCnt = DrRs.OpenCursor(, objPtInfo.SqlPtntSearch(lngSearch, MedSetPtid(txtSearchKey), _
                              mvarDeptCd, Format(DBConn.GetSysDate, CS_DateDbFormat)))
            Else
                ColCnt = DrRs.OpenCursor(, objPtInfo.SqlPtntSearch(lngSearch, txtSearchKey, _
                              mvarDeptCd, Format(DBConn.GetSysDate, CS_DateDbFormat)))
            End If
        End If
        lvwPtList.ListItems.Clear
        If ColCnt > 0 Then
            RowCnt = 0
            With lvwPtList
                Do While (DrRs.FetchCursor(ColCnt))
                    RowCnt = RowCnt + 1
                    Set itmx = .ListItems.Add(, , "" & MedGetPtid(DrRs.GetValue("ptid")))
                    itmx.SubItems(1) = "" & DrRs.GetValue("ptnm")
                    itmx.SubItems(2) = "" & DrRs.GetValue("SSN")
                    itmx.SubItems(3) = "" & DrRs.GetValue("DOB")
                    If Not IsDate(itmx.SubItems(3)) Then
                        itmx.SubItems(3) = Mid(itmx.SubItems(3), 1, 4) & "-01-01"
                    End If
                    If IsNumeric(Mid("" & DrRs.GetValue("ssn"), 8, 1)) Then
                        itmx.SubItems(4) = IIf((Mid("" & DrRs.GetValue("ssn"), 8, 1) Mod 2) = 1, "��", "��")
                    Else
                        itmx.SubItems(4) = "��"
                    End If
                    If IsDate(itmx.SubItems(3)) Then
                        itmx.SubItems(4) = itmx.SubItems(4) & " / " & DateDiff("yyyy", itmx.SubItems(3), Now)
                    Else
                        itmx.SubItems(4) = itmx.SubItems(4) & " / ? "
                    End If
                    If RowCnt > 1000 Then Exit Do
                Loop
            End With
        Else
            MsgBox "���ǿ� �´� �ڷᰡ �����ϴ�. Ȯ���� �˻��ϼ���", vbInformation + vbOKOnly, Me.Caption
        End If
        DrRs.CloseCursor:     Set DrRs = Nothing
    
    End If
    
    Set objPtInfo = Nothing
    
End Sub

'% Clear ��ƾ
Private Sub ClearRtn()
    lblPtNm.Caption = ""
    lblSex.Caption = ""
    lblAge.Caption = ""
    lblAgeDiv.Caption = ""
    lblDeptNm.Caption = ""
    lblLocation.Caption = ""
    lblDisease.Caption = ""
    'lblBedinDt.Caption = ""
    'lblBedoutDt.Caption = ""
    Call FieldClear
    Call TableClear
    ClearFg = True
    OrderFg = False
    MsgFg = False
    StopFg = False
    QueryFg = False
    OldRow = 0
    cmdReport.Enabled = False
End Sub

Private Sub FieldClear()

    'lblDoctNm.Caption = ""
    'lblCollectorNm.Caption = ""
    'lblReceiverNm.Caption = ""
    lblVerifierNm.Caption = ""
    'lblOrdDt.Caption = ""
    'lblCollectDt.Caption = ""
    'lblReceiveDt.Caption = ""
    lblVerifyDt.Caption = ""
    txtSamCmt.Text = ""
    'txtRstCmt.Text = ""
    'txtSamCmt1.Text = ""
    txtRstCmt1.Text = ""
    'lblWorkArea.Caption = ""
    'lblSpecimenNm.Caption = ""

End Sub

Private Sub TableClear()
    tblOrdSheet.MaxRows = 0
    tblOrdSheet.MaxRows = 100
    tblResult.MaxRows = 0
    tblResult.MaxRows = 100
    OldRow = 0
    TopLeftShow = False
End Sub

'% ��� Part Clear
Private Sub ResultClear()
   
'    ResultBuffer = ""
'    RstTxtBuffer = ""
'    SamTxtBuffer = ""
   
    'txtRstCmt.Text = ""
    txtSamCmt.Text = ""
    txtRstCmt1.Text = ""
    'txtSamCmt1.Text = ""
      
    'lblWorkArea.Caption = ""
    lblSpecimenNm.Caption = ""
    rtfResult.Text = ""
            
    fraLisResult.Visible = True
    tblResult.Visible = True
    picFootNote.Visible = True
    rtfResult.Visible = False
   
    ResultFg = False
   
    With tblResult
        '������̺� Clear
        .Row = -1:  .Col = -1
        .BlockMode = True
        .FontBold = False
        .Action = ActionClearText
        .ForeColor = &H747474
        .BlockMode = False
        '�˻��/��� �÷� Bold
        .Row = -1: .Col = 2: .Col2 = 3
        .BlockMode = True
        .FontBold = True
        .BlockMode = False
        'High/Low field font ����
        .Row = -1: .Col = 5: .Col2 = 5
        .BlockMode = True
        .FontName = "����"
        .BlockMode = False
        .RowsFrozen = 0
    End With
    
    cmdReport.Enabled = False

End Sub

Private Sub GetRefValue(ByVal iRow As Integer)

    Dim tmpTestCd As String
    Dim tmpSpcCd As String
    Dim tmpVfyDt As String
    Dim tmpSex As String
    Dim tmpAgeDay As String
    Dim tmpRs1 As New DrRecordSet
    Dim tmpRefFromVal As Double
    Dim tmpRefToVal As Double
    Dim tmpRefCd As String
    Dim SqlStmt As String
      
    With tblOrdSheet
        '����ġ �˻�
        .Row = iRow
        .Col = 8: If .Value <> CS_QuestionMark Then Exit Sub
        
        .Col = 25:    tmpSex = Trim(.Value)
        .Col = 26:    tmpAgeDay = Trim(.Value)
        .Col = 27:    tmpTestCd = Trim(.Value)
        .Col = 28:    tmpSpcCd = Trim(.Value)
        .Col = 29:    tmpVfyDt = Trim(.Value)
                      If tmpVfyDt = "" Then tmpVfyDt = Format(Now, CS_DateDbFormat)
        
        'Debug.Print tmpTestCd, ",", tmpSpcCd
        SqlStmt = MySql.SqlGetReference(tmpTestCd, tmpSpcCd, tmpVfyDt, "B", tmpAgeDay)
        Set tmpRs1 = OpenRecordSet(SqlStmt)
        If tmpRs1.EOF Then  '"B"(Both)�� �ش��ϴ� ����ġ�� ���� ��� ȯ�ڼ����� �ش��ϴ� ����Ÿ �˻� --> ���� Both�� ��ϵ�.
           tmpRs1.RsClose
           SqlStmt = MySql.SqlGetReference(tmpTestCd, tmpSpcCd, tmpVfyDt, tmpSex, tmpAgeDay)
           Set tmpRs1 = OpenRecordSet(SqlStmt)
        End If
        If tmpRs1.EOF Then
           tmpRefCd = Space(5)
        Else
           tmpRefFromVal = Val("" & tmpRs1.Fields("RefValFrom").Value)
           tmpRefToVal = Val("" & tmpRs1.Fields("RefValTo").Value)
           tmpRefCd = Trim("" & tmpRs1.Fields("RefCd").Value)
           If tmpRefFromVal <> 0 Or tmpRefToVal <> 0 Then tmpRefCd = tmpRefFromVal & "  -  " & tmpRefToVal
        End If
        tmpRs1.RsClose
        .Col = 8: .ForeColor = &H8000&
        If Trim(tmpRefCd) = "" Then
            .Value = "����"
        Else
            .Value = tmpRefCd:
        End If
    End With
    
    Set tmpRs1 = Nothing
    
End Sub

Public Sub Call_ToDate_LostFocus()

    If Not gUsingInWardMenu Then
        If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
        If ActiveControl.Name = cmdExit.Name Then Exit Sub
        If ActiveControl.Name = chkPtList.Name Then Exit Sub
    End If
'    Call dtpToDate_KeyDown(vbKeyReturn, 0)
    Call cmdRefresh_Click
   
End Sub


Public Sub Call_PtId_KeyPress()

    On Error GoTo Err_Trap
    
    If txtPtId.Text = "" Then
        If Screen.ActiveForm.Name = Me.Name Then txtPtId.SetFocus
        Exit Sub
    End If
    
      With MyPatient
         If .PtntQuery(MedSetPtid(txtPtId.Text)) Then
            lblPtNm.Caption = .PtNm
            lblSex.Caption = .SexNm
            lblAge.Caption = .Age
            lblAgeDiv.Caption = .AgeDiv
            lblDeptNm.Caption = .DeptNm
            'lblBedinDt.Caption = Format(.BedInDt, CS_DateMask)
            'lblBedoutDt.Caption = Format(.BedOutDt, CS_DateMask)
            txtPtId.SetFocus
            PtFg = True
'            DoEvents
            
            Dim objDisease  As New S2LIS_ReportLib.clsDisease
            objDisease.ptid = MedSetPtid(txtPtId.Text)
            lblDisease.Caption = objDisease.Disease
            Set objDisease = Nothing
                
            gPatientId = MedSetPtid(txtPtId.Text)
            ClearFg = False
         Else
            MsgFg = True
            MsgBox "��ϵ��� ���� ȯ��ID�Դϴ�.. �ٽ� �Է��ϼ���.."
            Me.Enabled = True
            txtPtId.SetFocus
            MsgFg = False
            PtFg = False
            Call txtPtId_GotFocus
            Exit Sub
         End If
      End With
      If ClearFg Then Call dtpToDate.SetFocus
      
      StopFg = False
      Exit Sub
Err_Trap:
    Resume Next

End Sub

Private Sub DisplayLABCommrnt(ByVal iRow As Long)

    Dim sBedinDt As String
    
    tblOrdSheet.Row = iRow
    tblOrdSheet.Col = 10
    sBedinDt = tblOrdSheet.Value
    
    With frmLabReport
        .ZOrder 0
        DoEvents
        .ptid = MedSetPtid(txtPtId.Text)
        .BedinDt = Format(sBedinDt, CS_DateDbFormat)
        Call .StartQuery
        .Show 1
    End With
    
    'Call medAlwaysOn(frmLabReport, 1)

End Sub


