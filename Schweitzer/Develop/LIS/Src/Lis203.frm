VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frm203InstDataEntry 
   BackColor       =   &H00DBE6E6&
   Caption         =   "�ڵ�ȭ��� ������"
   ClientHeight    =   9225
   ClientLeft      =   1980
   ClientTop       =   5805
   ClientWidth     =   14505
   BeginProperty Font 
      Name            =   "����"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Lis203.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   14505
   Tag             =   "20300"
   Visible         =   0   'False
   WindowState     =   2  '�ִ�ȭ
   Begin VB.Frame frmSMS 
      BackColor       =   &H00F8E4D8&
      Caption         =   "SMS����"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   5415
      Left            =   6330
      TabIndex        =   59
      Top             =   1590
      Width           =   4515
      Begin VB.TextBox txtTestCd 
         Appearance      =   0  '���
         BackColor       =   &H00F1F5F4&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5100
         MaxLength       =   15
         TabIndex        =   74
         Tag             =   "opt"
         Top             =   1350
         Width           =   1305
      End
      Begin VB.TextBox txtTransDt 
         Appearance      =   0  '���
         BackColor       =   &H00F1F5F4&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1140
         MaxLength       =   25
         TabIndex        =   73
         Tag             =   "opt"
         Top             =   4170
         Width           =   3195
      End
      Begin VB.TextBox txtDtNo 
         Appearance      =   0  '���
         BackColor       =   &H00F1F5F4&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2040
         MaxLength       =   15
         TabIndex        =   72
         Tag             =   "opt"
         Top             =   1410
         Width           =   2295
      End
      Begin VB.TextBox txtDeptNm 
         Appearance      =   0  '���
         BackColor       =   &H00F1F5F4&
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1140
         MaxLength       =   15
         TabIndex        =   71
         Tag             =   "opt"
         Top             =   2580
         Width           =   3195
      End
      Begin VB.TextBox txtDetpCd 
         Appearance      =   0  '���
         BackColor       =   &H00F1F5F4&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1140
         MaxLength       =   15
         TabIndex        =   70
         Tag             =   "opt"
         Top             =   2580
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtDtNm 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00F1F5F4&
         Height          =   360
         Left            =   2040
         MaxLength       =   15
         TabIndex        =   69
         Tag             =   "opt"
         Top             =   1020
         Width           =   975
      End
      Begin VB.TextBox txtDtId 
         Appearance      =   0  '���
         BackColor       =   &H00F1F5F4&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3030
         MaxLength       =   15
         TabIndex        =   68
         Tag             =   "opt"
         Top             =   1020
         Width           =   1305
      End
      Begin VB.TextBox txtTransNo 
         Appearance      =   0  '���
         BackColor       =   &H00F1F5F4&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1140
         MaxLength       =   15
         TabIndex        =   67
         Tag             =   "opt"
         Top             =   630
         Width           =   3195
      End
      Begin VB.TextBox txtTransNm 
         Appearance      =   0  '���
         BackColor       =   &H00F1F5F4&
         Height          =   360
         Left            =   2460
         MaxLength       =   15
         TabIndex        =   66
         Tag             =   "opt"
         Top             =   300
         Width           =   1875
      End
      Begin VB.TextBox txtTransId 
         Appearance      =   0  '���
         BackColor       =   &H00F1F5F4&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1140
         MaxLength       =   15
         TabIndex        =   65
         Tag             =   "opt"
         Top             =   300
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancle 
         BackColor       =   &H00F4F0F2&
         Caption         =   "���"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3030
         Style           =   1  '�׷���
         TabIndex        =   64
         Tag             =   "135"
         Top             =   4680
         Width           =   1320
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00F4F0F2&
         Caption         =   "����"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1680
         Style           =   1  '�׷���
         TabIndex        =   63
         Tag             =   "135"
         Top             =   4680
         Width           =   1320
      End
      Begin VB.TextBox txtExDtId 
         Appearance      =   0  '���
         BackColor       =   &H00F1F5F4&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3030
         MaxLength       =   15
         TabIndex        =   62
         Tag             =   "opt"
         Top             =   1800
         Width           =   1305
      End
      Begin VB.TextBox txtExDtNm 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00F1F5F4&
         Height          =   360
         Left            =   2040
         MaxLength       =   15
         TabIndex        =   61
         Tag             =   "opt"
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox txtExDtNo 
         Appearance      =   0  '���
         BackColor       =   &H00F1F5F4&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2040
         MaxLength       =   15
         TabIndex        =   60
         Tag             =   "opt"
         Top             =   2190
         Width           =   2295
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   345
         Index           =   7
         Left            =   180
         TabIndex        =   75
         Top             =   300
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   609
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
         Caption         =   "������"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   1905
         Index           =   8
         Left            =   180
         TabIndex        =   76
         Top             =   1020
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   3360
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
         Caption         =   "������"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   9
         Left            =   180
         TabIndex        =   77
         Top             =   2970
         Width           =   915
         _ExtentX        =   1614
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
         Caption         =   "�޽���"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   10
         Left            =   180
         TabIndex        =   78
         Top             =   4200
         Width           =   915
         _ExtentX        =   1614
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
         Caption         =   "�����Ͻ�"
         Appearance      =   0
      End
      Begin RichTextLib.RichTextBox rtfMessage 
         Height          =   1170
         Left            =   1140
         TabIndex        =   79
         Top             =   2970
         Width           =   3210
         _ExtentX        =   5662
         _ExtentY        =   2064
         _Version        =   393217
         BackColor       =   16776172
         ScrollBars      =   2
         TextRTF         =   $"Lis203.frx":08CA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   345
         Index           =   11
         Left            =   180
         TabIndex        =   80
         Top             =   630
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   609
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
         Caption         =   "������ȣ"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   765
         Index           =   13
         Left            =   1140
         TabIndex        =   81
         Top             =   1020
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   1349
         BackColor       =   14737632
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
         Caption         =   "ó����"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   765
         Index           =   14
         Left            =   1140
         TabIndex        =   82
         Top             =   1800
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   1349
         BackColor       =   14737632
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
         Caption         =   "��ġ��"
         Appearance      =   0
      End
   End
   Begin VB.CommandButton cmdOrderView 
      BackColor       =   &H00F4F0F2&
      Caption         =   "ó�溰��ȸ(&C)"
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
      Left            =   5480
      Style           =   1  '�׷���
      TabIndex        =   55
      Top             =   90
      Visible         =   0   'False
      Width           =   1500
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   330
      Index           =   5
      Left            =   10290
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   6405
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   582
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
      Caption         =   "��ġ���"
      Appearance      =   0
   End
   Begin VB.TextBox txtBatchRst 
      Appearance      =   0  '���
      BackColor       =   &H00F1F5F4&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   11235
      MaxLength       =   15
      TabIndex        =   53
      Tag             =   "opt"
      Top             =   6405
      Width           =   1785
   End
   Begin VB.CommandButton cmdApply 
      BackColor       =   &H00ACCDD0&
      Caption         =   "����"
      Enabled         =   0   'False
      Height          =   330
      Left            =   13035
      Style           =   1  '�׷���
      TabIndex        =   52
      Top             =   6405
      Width           =   690
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00F4F0F2&
      Caption         =   "���"
      Height          =   330
      Left            =   13740
      Style           =   1  '�׷���
      TabIndex        =   51
      Top             =   6405
      Width           =   690
   End
   Begin VB.CommandButton cmdSpecial 
      BackColor       =   &H00DBE6E6&
      Caption         =   "Ư  ��"
      Height          =   285
      Left            =   12750
      Style           =   1  '�׷���
      TabIndex        =   46
      Top             =   100
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.CommandButton cmdMicro 
      BackColor       =   &H00DBE6E6&
      Caption         =   "�̻���"
      Height          =   285
      Left            =   13590
      Style           =   1  '�׷���
      TabIndex        =   45
      Top             =   100
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.CommandButton cmdRmk 
      BackColor       =   &H008080FF&
      Caption         =   "ó����"
      Height          =   285
      Left            =   11850
      Style           =   1  '�׷���
      TabIndex        =   44
      Top             =   100
      Visible         =   0   'False
      Width           =   900
   End
   Begin MedControls1.LisLabel lblDisease 
      Height          =   270
      Left            =   8355
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   435
      Width           =   6050
      _ExtentX        =   10663
      _ExtentY        =   476
      BackColor       =   16777215
      ForeColor       =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      Appearance      =   0
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "ȭ������(&C)"
      CausesValidation=   0   'False
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
      TabIndex        =   39
      Tag             =   "124"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "����(&X)"
      CausesValidation=   0   'False
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
      TabIndex        =   38
      Tag             =   "128"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "Ȯ��(&S)"
      CausesValidation=   0   'False
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
      TabIndex        =   37
      Tag             =   "135"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.Frame fraCul 
      BackColor       =   &H00DBE6E6&
      BorderStyle     =   0  '����
      Height          =   555
      Left            =   6420
      TabIndex        =   34
      Top             =   8535
      Width           =   4065
      Begin VB.CommandButton cmdSMS 
         BackColor       =   &H008080FF&
         Caption         =   "SMS"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1620
         Style           =   1  '�׷���
         TabIndex        =   56
         Tag             =   "135"
         Top             =   30
         Width           =   1080
      End
      Begin VB.CommandButton cmdCul 
         BackColor       =   &H00F4F0F2&
         Caption         =   "���������ȸ"
         CausesValidation=   0   'False
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
         Left            =   2730
         Style           =   1  '�׷���
         TabIndex        =   36
         Tag             =   "135"
         Top             =   0
         Width           =   1320
      End
      Begin VB.CheckBox chkCul 
         BackColor       =   &H00DBE6E6&
         Caption         =   "�κ�"
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
         Left            =   120
         TabIndex        =   35
         Top             =   90
         Width           =   960
      End
   End
   Begin VB.Frame fraEQP 
      BackColor       =   &H00DBE6E6&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Left            =   75
      TabIndex        =   9
      Top             =   -60
      Width           =   5355
      Begin VB.CheckBox chkStatFg 
         BackColor       =   &H00DBE6E6&
         Caption         =   "���ް�ü�켱"
         ForeColor       =   &H006B72A9&
         Height          =   225
         Left            =   2655
         TabIndex        =   2
         Top             =   570
         Width           =   1470
      End
      Begin VB.CommandButton cmdEqp 
         BackColor       =   &H00DEDBDD&
         Caption         =   "��"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2250
         MousePointer    =   14  'ȭ��ǥ�� ����ǥ
         Style           =   1  '�׷���
         TabIndex        =   27
         Top             =   195
         Width           =   285
      End
      Begin VB.CommandButton cmdTrans 
         BackColor       =   &H00DEDBDD&
         Caption         =   "��"
         CausesValidation=   0   'False
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
         Left            =   2250
         MousePointer    =   14  'ȭ��ǥ�� ����ǥ
         Style           =   1  '�׷���
         TabIndex        =   18
         Top             =   555
         Width           =   285
      End
      Begin VB.CommandButton cmdQuery 
         BackColor       =   &H00FFF7FC&
         Caption         =   "&Query"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4095
         MaskColor       =   &H00808080&
         Style           =   1  '�׷���
         TabIndex        =   3
         Top             =   690
         Width           =   1185
      End
      Begin VB.TextBox txtEqpCd 
         Appearance      =   0  '���
         BackColor       =   &H00F1F5F4&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   960
         TabIndex        =   0
         Text            =   "MEDS-770"
         Top             =   210
         Width           =   1290
      End
      Begin MSMask.MaskEdBox mskSpcNo 
         Height          =   315
         Left            =   990
         TabIndex        =   1
         Top             =   570
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   15857140
         AutoTab         =   -1  'True
         MaxLength       =   12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "&&-#########"
         PromptChar      =   "_"
      End
      Begin MedControls1.LisLabel lblEqpCdNm 
         Height          =   315
         Left            =   2550
         TabIndex        =   23
         Top             =   210
         Width           =   2730
         _ExtentX        =   4815
         _ExtentY        =   556
         BackColor       =   15265000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Appearance      =   0
      End
      Begin VB.CheckBox chkUr 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Urine"
         ForeColor       =   &H006B72A9&
         Height          =   225
         Left            =   2655
         TabIndex        =   30
         Top             =   795
         Visible         =   0   'False
         Width           =   1605
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   0
         Left            =   30
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   210
         Width           =   915
         _ExtentX        =   1614
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
         Caption         =   "����ڵ�"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   4
         Left            =   30
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   570
         Width           =   915
         _ExtentX        =   1614
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
         Caption         =   "��ü��ȣ"
         Appearance      =   0
      End
   End
   Begin VB.ComboBox cboRelTest 
      BackColor       =   &H00F1F5F4&
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8355
      Style           =   2  '��Ӵٿ� ���
      TabIndex        =   26
      Top             =   720
      Width           =   6090
   End
   Begin VB.CommandButton cmdPrevious 
      BackColor       =   &H00FFF7FC&
      Caption         =   "<< (&P)"
      Height          =   375
      Left            =   5430
      Style           =   1  '�׷���
      TabIndex        =   25
      Top             =   630
      Width           =   810
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00FFF7FC&
      Caption         =   "(&N) >>"
      Height          =   375
      Left            =   6240
      Style           =   1  '�׷���
      TabIndex        =   24
      Top             =   630
      Width           =   765
   End
   Begin VB.PictureBox picRst 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4770
      Left            =   3435
      ScaleHeight     =   4710
      ScaleWidth      =   10965
      TabIndex        =   10
      Top             =   1605
      Width           =   11025
      Begin MSComctlLib.ProgressBar prgRst 
         Height          =   240
         Left            =   0
         TabIndex        =   11
         ToolTipText     =   "�ڷḦ �������� �����ϴ�."
         Top             =   4485
         Visible         =   0   'False
         Width           =   10785
         _ExtentX        =   19024
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin FPSpread.vaSpread ssRst 
         CausesValidation=   0   'False
         Height          =   4515
         Left            =   0
         TabIndex        =   6
         Tag             =   "20001"
         Top             =   0
         Width           =   10965
         _Version        =   196608
         _ExtentX        =   19341
         _ExtentY        =   7964
         _StockProps     =   64
         BackColorStyle  =   1
         BorderStyle     =   0
         ColHeaderDisplay=   0
         DisplayRowHeaders=   0   'False
         EditEnterAction =   8
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
         GrayAreaBackColor=   15857140
         GridColor       =   13290186
         MaxCols         =   19
         MaxRows         =   15
         Protect         =   0   'False
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         SpreadDesigner  =   "Lis203.frx":0967
         VisibleCols     =   10
         VisibleRows     =   13
         TextTip         =   2
      End
      Begin VB.Label lblSpreadLoading 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '����
         Caption         =   "��� ��ٷ� �ּ���. ��� �����͸� �ε��ϰ� �����ϴ�."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2970
         TabIndex        =   12
         Top             =   1890
         Width           =   4605
      End
   End
   Begin MSComctlLib.ListView lvwPatient 
      Height          =   555
      Left            =   3420
      TabIndex        =   5
      Tag             =   "20113"
      Top             =   1020
      Width           =   11025
      _ExtentX        =   19447
      _ExtentY        =   979
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   15857140
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvwEQP 
      Height          =   7170
      Left            =   75
      TabIndex        =   4
      Top             =   990
      Width           =   3345
      _ExtentX        =   5900
      _ExtentY        =   12647
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   15857140
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Frame fraText 
      BackColor       =   &H00DBE6E6&
      Caption         =   " Text Result"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   8940
      TabIndex        =   13
      Tag             =   "20002"
      Top             =   6690
      Width           =   5520
      Begin VB.CommandButton cmdTextTemplete 
         BackColor       =   &H00DEDBDD&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5160
         Picture         =   "Lis203.frx":121F
         Style           =   1  '�׷���
         TabIndex        =   14
         Top             =   870
         Width           =   315
      End
      Begin RichTextLib.RichTextBox rtfText 
         Height          =   945
         Left            =   90
         TabIndex        =   8
         Top             =   225
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   1667
         _Version        =   393217
         BackColor       =   15663102
         Enabled         =   0   'False
         ScrollBars      =   3
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"Lis203.frx":1751
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox rtfFlagText 
         Height          =   525
         Left            =   570
         TabIndex        =   57
         Top             =   1200
         Width           =   4905
         _ExtentX        =   8652
         _ExtentY        =   926
         _Version        =   393217
         BackColor       =   16777215
         ScrollBars      =   3
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"Lis203.frx":19C4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   525
         Index           =   12
         Left            =   90
         TabIndex        =   58
         Top             =   1200
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   926
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
         Caption         =   "Flag"
         Appearance      =   0
      End
   End
   Begin VB.Frame fraComment 
      BackColor       =   &H00DBE6E6&
      Caption         =   "Comment by Accession No"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   3450
      TabIndex        =   15
      Tag             =   "20003"
      Top             =   6690
      Width           =   5445
      Begin VB.CommandButton cmdRemarkTemplete 
         BackColor       =   &H00DEDBDD&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5055
         Picture         =   "Lis203.frx":1A66
         Style           =   1  '�׷���
         TabIndex        =   19
         Top             =   1440
         Width           =   315
      End
      Begin VB.CommandButton cmdCommentTemplete 
         BackColor       =   &H00DEDBDD&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5050
         Picture         =   "Lis203.frx":1F98
         Style           =   1  '�׷���
         TabIndex        =   16
         Top             =   870
         Width           =   315
      End
      Begin RichTextLib.RichTextBox rtfComment 
         Height          =   960
         Left            =   90
         TabIndex        =   7
         Top             =   225
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   1693
         _Version        =   393217
         BackColor       =   15857140
         Enabled         =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"Lis203.frx":24CA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox rtfRemark 
         Height          =   360
         Left            =   90
         TabIndex        =   20
         Top             =   1410
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   635
         _Version        =   393217
         BackColor       =   16776172
         Enabled         =   0   'False
         ScrollBars      =   2
         TextRTF         =   $"Lis203.frx":26FC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Remark"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   21
         Top             =   1200
         Width           =   1545
      End
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   345
      Left            =   75
      TabIndex        =   29
      Top             =   8175
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   609
      BackColor       =   13752531
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "���۵� �����Ǽ�"
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   270
      Index           =   1
      Left            =   7050
      TabIndex        =   31
      Top             =   150
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   476
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
      Caption         =   "�� �� ó"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   270
      Index           =   2
      Left            =   7050
      TabIndex        =   32
      Top             =   435
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   476
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
      Height          =   270
      Index           =   3
      Left            =   7050
      TabIndex        =   33
      Top             =   735
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   476
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
      Caption         =   "���ð˻� ���"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel lblTelno 
      Height          =   270
      Left            =   8355
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   150
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   476
      BackColor       =   16777215
      ForeColor       =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      Appearance      =   0
   End
   Begin VB.ListBox lstEQCode 
      Appearance      =   0  '���
      BackColor       =   &H00FFF7FC&
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   1035
      TabIndex        =   22
      Top             =   480
      Visible         =   0   'False
      Width           =   5235
   End
   Begin VB.Frame fraMesg 
      BackColor       =   &H00DBE6E6&
      Height          =   2655
      Left            =   10335
      TabIndex        =   47
      Top             =   630
      Visible         =   0   'False
      Width           =   4095
      Begin VB.TextBox txtMesg 
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1785
         Left            =   15
         MultiLine       =   -1  'True
         ScrollBars      =   2  '����
         TabIndex        =   49
         Top             =   390
         Width           =   4050
      End
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Ȯ��"
         Height          =   420
         Left            =   2940
         Style           =   1  '�׷���
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   2175
         Width           =   1095
      End
      Begin MedControls1.LisLabel LisLabel7 
         Height          =   300
         Index           =   2
         Left            =   15
         TabIndex        =   50
         Top             =   90
         Width           =   4050
         _ExtentX        =   7144
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
         Caption         =   "ó�� ������ ��ȸ"
         Appearance      =   0
         LeftGab         =   200
      End
   End
   Begin VB.Label lblErr 
      AutoSize        =   -1  'True
      BackColor       =   &H00DDF0F5&
      BackStyle       =   0  '����
      Caption         =   "������ �߻��ߴ�."
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00313D46&
      Height          =   180
      Left            =   240
      TabIndex        =   28
      Top             =   8760
      Width           =   1380
   End
   Begin VB.Label lblAccNoCnt 
      Alignment       =   1  '������ ����
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  '���� ����
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2490
      TabIndex        =   17
      Tag             =   "20304"
      Top             =   8175
      Width           =   915
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFF9F7&
      BackStyle       =   1  '�������� ����
      BorderColor     =   &H00C0C0C0&
      Height          =   330
      Left            =   60
      Shape           =   4  '�ձ� �簢��
      Top             =   8685
      Width           =   6255
   End
End
Attribute VB_Name = "frm203InstDataEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents clsTemplete   As frm230TempSearch
Attribute clsTemplete.VB_VarHelpID = -1
Private WithEvents objCodeList  As clsPopUpList
Attribute objCodeList.VB_VarHelpID = -1
Private WithEvents objCuM       As frmTmpCumulative
Attribute objCuM.VB_VarHelpID = -1

Private objLab306       As clsEquipTransfer
Private objPtInfo       As clsPatientInfo

Private insForm         As Form
Private gintTemplete    As Integer
Private blnFirst        As Boolean
Private gblnNewObj      As Boolean
Private blnDayCount     As Boolean
Private gstrPtAddInfo   As String
Private gblnModify      As Boolean
Private gstrModifyData  As String
Private gstrMsk         As String

Private IndexPointer    As Integer
Private MsgFg           As Boolean
Private LeaveCellFg As Boolean

Private strCombo        As String
Private blnRstChange    As Boolean
Private blnExpect       As Boolean

Private AdoCn_SQL       As ADODB.Connection
Private AdoRs_SQL       As ADODB.Recordset

Private AdoCn_ORACLE    As ADODB.Connection
Private AdoRs_ORACLE    As ADODB.Recordset
Dim strRcvDt    As String
Dim strTmpAge   As String
Dim strTmpSex   As String

Private Sub cmdApply_Click()
    Dim strBatchRst As String
    Dim i           As Long
    
    If Trim(txtBatchRst.Text) = "" Then
        txtBatchRst.SetFocus
        Exit Sub
    End If
    
    '** �߿����
    ' * �÷� �� ������ ���ġ �ʰ� �ִ� 14�� ��ġ��� ��� �� ����Ѵ�.
    strBatchRst = Trim(txtBatchRst.Text)
    i = 1
    With ssRst
        If .DataRowCnt = 0 Then
            Exit Sub
        End If
        
        For i = 1 To .DataRowCnt
            .Row = i
            
            .Col = 1
            If .BackColor = vbGrayText Then
                GoTo Skip
            End If
            
            .Col = 2:
            'If Trim(.Value) = "" Then
            If Len(Trim(.Value & "")) = 0 Then  '��������°͸� ó���Ѵ�. 2014-08-13 PSK
                .Value = strBatchRst
                .ForeColor = DCM_LightRed
                
                '-- Batch Result
                .Col = .MaxCols: .Value = strBatchRst
                
                '-- Batch Result Check Flag
                .Col = 14: .Value = 1
            End If
            
Skip:
            
        Next
    End With
    
End Sub

Private Sub cmdCancel_Click()
    Dim i       As Long
    
    i = 1
    With ssRst
        If .DataRowCnt = 0 Then
            Exit Sub
        End If
        
        For i = 1 To .DataRowCnt
            .Row = i
            
            '-- Batch Result
            .Col = 14:
            If Trim(.Value) = 1 Then
                .Value = ""
                
                '-- Batch Result
                .Col = 2: .Value = ""
                
                '-- Batch Result Check Flag
                .Col = 14: .Value = ""
            End If
        Next
    End With
    
End Sub

Private Sub cmdCancle_Click()
    frmSMS.Visible = False
End Sub

Private Sub cmdOrderView_Click()
' 2008.12.17. �缺�� �۾����Դϴ�.
' 2009.01.09 �缺�� ȯ��ID �Ķ���� �߰�
' 2009.04.13 �缺�� �߰�
    Dim i As Integer
    Dim pFrmName As String
'    Dim cxxx  As S2LIS_ReviewLib.clsLISResultReview
    pFrmName = "frm401ResultView"
    
    If ObjMyUser(pFrmName) Is Nothing Then GoTo PermissionDenied
    If Not ObjMyUser(pFrmName).CanRead Then GoTo PermissionDenied

    medMain.lblSubMenu.Caption = "ó������ȸ" 'medGetP(Button.Tag, 1, "(")
    
    
'   gPatientId = lblPtId.Caption
'  s2lis_reviewlib.PtId = lblPtId.Caption
    
'    gUsingInWardMenu = True
    frmLisReview.ButtonKey = "LIS155A" 'Button.Key
    frmLisReview.PtId = lvwPatient.ListItems(1).SubItems(1)   'lblPtId.Caption
    frmLisReview.Show
    frmLisReview.ZOrder 0
    frmLisReview.ShowThisForm

    Exit Sub

PermissionDenied:
   
'    blnFormShow = False
    MsgBox "�� ȭ���� ����� �� �ִ� ������ �����ϴ�.", vbExclamation, "Security Check!"

End Sub

Private Sub cmdSMS_Click()
    Dim SSQL As String
    
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
    
    frmSMS.Visible = True
    txtTransId.Text = Trim(ObjSysInfo.EmpId)
    txtTransNm.Text = GetEmpNm(Trim(ObjSysInfo.EmpId))
'    txtTransNo.Text = mskAccNo.Text
    txtDtNo.Text = ""
    txtTransDt.Text = Format(Now, "YYYY-MM-DD HH:MM:DD")
    
    rtfMessage.Text = rtfMessage.Text & vbCRLF & "Critical value ���óġ����" & vbCr ' & rtfComment.Text
    If txtDtId.Text <> "" Then
'        SSQL = ""
'        SSQL = SSQL & vbCr & "SELECT TELNO,EMPNO FROM S2COM098"
'        SSQL = SSQL & vbCr & " WHERE replace(EMPNM,' ','') LIKE '%" & txtDtNm.Text & "'"
                       
'        SSQL = ""
'        SSQL = SSQL & vbCr & "SELECT hphoneno AS TELNO, empno AS EMPNO from gainsamt"
'        SSQL = SSQL & vbCr & " WHERE replace(EMPNM,' ','') LIKE '%" & txtDtNm.Text & "'"

'        SSQL = ""
'        SSQL = SSQL & vbCr & "SELECT hphoneno AS TELNO, empno AS EMPNO from gainsamt"
'        SSQL = SSQL & vbCr & " WHERE replace(EMPNO,' ','') = (select orddoct from s2lab201 where workarea = '" & objPtInfo.WorkArea & "' and accdt =  '" & objPtInfo.AccDt & "' and accseq = '" & objPtInfo.AccSeq & "')"

        SSQL = ""
        SSQL = SSQL & vbCr & "SELECT hphoneno AS TELNO, EMPNM AS EMPNM from gainsamt"
        SSQL = SSQL & vbCr & " WHERE replace(EMPNO,' ','') = '" & txtDtId.Text & "' "

        Set AdoRs_ORACLE = New ADODB.Recordset
    
        AdoRs_ORACLE.CursorLocation = adUseClient
        AdoRs_ORACLE.Open SSQL, AdoCn_ORACLE
    
        If AdoRs_ORACLE.RecordCount > 0 Then
            txtDtNo.Text = AdoRs_ORACLE.Fields("TELNO") & ""
            txtDtNm.Text = AdoRs_ORACLE.Fields("EMPNM") & ""
        End If
'
'        Set AdoCn_ORACLE = Nothing
    End If
    
    If txtExDtId.Text <> "" Then
'        SSQL = ""
'        SSQL = SSQL & vbCr & "SELECT TELNO,EMPNO FROM S2COM098"
'        SSQL = SSQL & vbCr & " WHERE replace(EMPNM,' ','') LIKE '%" & txtDtNm.Text & "'"
                       
'        SSQL = ""
'        SSQL = SSQL & vbCr & "SELECT hphoneno AS TELNO, empno AS EMPNO from gainsamt"
'        SSQL = SSQL & vbCr & " WHERE replace(EMPNM,' ','') LIKE '%" & txtDtNm.Text & "'"

'        SSQL = ""
'        SSQL = SSQL & vbCr & "SELECT hphoneno AS TELNO, empno AS EMPNO from gainsamt"
'        SSQL = SSQL & vbCr & " WHERE replace(EMPNO,' ','') = (select orddoct from s2lab201 where workarea = '" & objPtInfo.WorkArea & "' and accdt =  '" & objPtInfo.AccDt & "' and accseq = '" & objPtInfo.AccSeq & "')"

        SSQL = ""
        SSQL = SSQL & vbCr & "SELECT hphoneno AS TELNO, EMPNM AS EMPNM from gainsamt"
        SSQL = SSQL & vbCr & " WHERE replace(EMPNO,' ','') = '" & txtExDtId.Text & "' "

        Set AdoRs_ORACLE = New ADODB.Recordset
    
        AdoRs_ORACLE.CursorLocation = adUseClient
        AdoRs_ORACLE.Open SSQL, AdoCn_ORACLE
    
        If AdoRs_ORACLE.RecordCount > 0 Then
            txtExDtNo.Text = AdoRs_ORACLE.Fields("TELNO") & ""
            txtExDtNm.Text = AdoRs_ORACLE.Fields("EMPNM") & ""
        End If
        
        Set AdoCn_ORACLE = Nothing
    End If
'    Dim SSQL As String
'
'    Set AdoCn_ORACLE = New ADODB.Connection
'
'    With AdoCn_ORACLE
'        .ConnectionTimeout = 25
''        .Provider = "OraOLEDB.Oracle.1"
'        .Provider = "MSDAORA.1"                 ' Oracle "MSDAORA.1"
'        .Properties("Data Source").Value = "PMC"
''        .Properties("Initial Catalog").Value = DatabaseName
'        .Properties("Persist Security Info") = True
'
'        .Properties("User ID").Value = "oral1"
'        .Properties("Password").Value = "oral1"
'
''        Screen.MousePointer = vbHourglass
'        .Open
'    End With
'
'    frmSMS.Visible = True
'    txtTransId.Text = Trim(ObjSysInfo.EmpId)
'    txtTransNm.Text = GetEmpNm(Trim(ObjSysInfo.EmpId))
''    txtTransNo.Text = ""
'    txtDtNo.Text = ""
'    txtTransDt.Text = Format(Now, "YYYY-MM-DD HH:MM:SS")
'
'    rtfMessage.Text = rtfMessage.Text & vbCRLF & "Critical value ���óġ����" & vbCr 'rtfComment.Text
'
'    If txtDtNm.Text <> "" Then
''        SSQL = ""
''        SSQL = SSQL & vbCr & "SELECT TELNO,EMPNO FROM S2COM098"
''        SSQL = SSQL & vbCr & " WHERE replace(EMPNM,' ','') LIKE '%" & txtDtNm.Text & "'"
'
''        SSQL = ""
''        SSQL = SSQL & vbCr & "SELECT hphoneno AS TELNO, empno AS EMPNO from gainsamt"
''        SSQL = SSQL & vbCr & " WHERE replace(EMPNM,' ','') LIKE '%" & txtDtNm.Text & "'"
'
'        SSQL = ""
'        SSQL = SSQL & vbCr & "SELECT hphoneno AS TELNO, empno AS EMPNO from gainsamt"
'        SSQL = SSQL & vbCr & " WHERE replace(EMPNO,' ','') = (select orddoct from s2lab201 where workarea = '" & objPtInfo.WorkArea & "' and accdt =  '" & objPtInfo.AccDt & "' and accseq = '" & objPtInfo.AccSeq & "')"
'
'        Set AdoRs_ORACLE = New ADODB.Recordset
'
'        AdoRs_ORACLE.CursorLocation = adUseClient
'        AdoRs_ORACLE.Open SSQL, AdoCn_ORACLE
'
'        If AdoRs_ORACLE.RecordCount > 0 Then
'            txtDtNo.Text = AdoRs_ORACLE.Fields("TELNO") & ""
'            txtDtId.Text = AdoRs_ORACLE.Fields("EMPNO") & ""
'        End If
'
'        Set AdoCn_ORACLE = Nothing
'    End If
End Sub

Private Sub cmdSpecial_Click()
    frmRealTestShow.DrFrame1.ZOrder 0
    frmRealTestShow.LisLabel7(0).Caption = "Ư���˻� ���ð˻� ����Ʈ"
    frmRealTestShow.Show
    Call frmRealTestShow.SpecialTest(lvwPatient.ListItems.Item(1).Text, lvwPatient.ListItems.Item(1).SubItems(1), cboRelTest, "1")
End Sub
Private Sub cmdMicro_Click()
    frmRealTestShow.DrFrame2.ZOrder 0
    frmRealTestShow.LisLabel7(0).Caption = "�̻��� ���ð˻� ����Ʈ"
    frmRealTestShow.Show
    Call frmRealTestShow.SpecialTest(lvwPatient.ListItems.Item(1).Text, lvwPatient.ListItems.Item(1).SubItems(1), cboRelTest, "2")
End Sub

Private Sub chkStatFg_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cmdClear_Click()
    Call ClearData
End Sub

Private Sub cmdEqp_Click()

    If lstEQCode.ListCount = 0 Then
        MsgBox "������ ��� �����ϴ�.", vbCritical
        Exit Sub
    End If
    lstEQCode.Visible = True
    Set objCodeList = Nothing
    lstEQCode.ZOrder 0
    lstEQCode.SetFocus

End Sub

Private Sub cmdExit_Click()
    
    Dim intYesNo As VbMsgBoxResult
    '
    If gblnModify = True Then
        objPtInfo.FootNote = rtfComment.Text
        objPtInfo.Result.Item(ssRst.ActiveRow).TextRst = rtfText.Text
        If DataFetch <> gstrModifyData Then
            intYesNo = MsgBox("�ڷᰡ �����Ǿ����ϴ�." & vbNewLine & "������ �ڷḦ �����Ͻðڽ��ϱ�?", _
                vbYesNo, "������")
            If intYesNo = vbYes Then Call cmdSave_Click    '����Ÿ ����
        End If
        gblnModify = False: gstrModifyData = ""
    End If

    Set clsTemplete = Nothing
    Set objLab306 = Nothing
    Set objPtInfo = Nothing
    
    Unload Me
    Set frm203InstDataEntry = Nothing
    
End Sub

Private Sub cmdNext_Click()
    
    Dim objLvwItem As MSComctlLib.ListItem

    If lvwEQP.ListItems.Count > IndexPointer Then
        Set objLvwItem = lvwEQP.ListItems.Item(IndexPointer + 1)
        lvwEQP_ItemClick objLvwItem
    End If

End Sub

Private Sub cmdPrevious_Click()
    
    Dim objLvwItem As MSComctlLib.ListItem

    If IndexPointer > 1 Then
        Set objLvwItem = lvwEQP.ListItems.Item(IndexPointer - 1)
        lvwEQP_ItemClick objLvwItem
    End If

End Sub

Private Sub cmdQuery_Click()
    
    Dim objLvwItem As MSComctlLib.ListItem
    Dim i As Integer
    '
    If txtEqpCd.Text = "" Then
        Exit Sub
    End If
   '
    If mskSpcNo.ClipText = "" Then
        Exit Sub
    End If
   '
    Set objLab306 = New clsEquipTransfer
    With objLab306
        '2011.12.26
        '�½�ȣ ��� ��� ��ȸ�� ������ư Ȱ��ȭ
        lvwEQP.ListItems.Clear '�߰�
        
        .LoadTable txtEqpCd.Text, mskSpcNo.FormattedText, chkStatFg.Value
        medDataLoadLvw lvwEQP, vbNewLine, vbTab, .GetStrEqpTrans
        DoEvents

        For i = 1 To lvwEQP.ListItems.Count
            Set objLvwItem = lvwEQP.ListItems.Item(i)
            If Trim(objLvwItem.SubItems(4)) = "1" Then objLvwItem.ForeColor = vbRed
        Next

        If .RecordCount > 0 Then
            EditData
            DisplayCount
            lvwEQP.FlatScrollBar = True
            Set objLvwItem = lvwEQP.ListItems.Item(1)
            lvwEQP_ItemClick objLvwItem
            cmdOrderView.Visible = True
        Else
            ClearData
            MsgBox "�ش� ����Ÿ�� �����ϴ�."
            cmdOrderView.Visible = False
        End If
    End With
   '
End Sub

Private Sub cmdRemarkTemplete_Click()
    
'    Dim SqlStmt As String
'
'    Set objCodeList = Nothing
'    Set objCodeList = New clsPopUpList
'
'    SqlStmt = "SELECT cdval1, text1 FROM " & T_LAB034 & " WHERE cdindex = '" & LC4_Remark & "' "
    Dim RS      As Recordset
    Dim strWorkArea As String
    Dim SqlStmt As String
    
    Set objCodeList = Nothing
    Set objCodeList = New clsPopUpList
    strWorkArea = medGetP(lvwPatient.ListItems.Item(1).Text, 1, "-")
    
    SqlStmt = "SELECT cdval1, text1 FROM " & T_LAB034 & " WHERE " & DBW("cdindex = ", LC4_Remark) & " and " & DBW("field1=", strWorkArea)
    Set RS = New Recordset
    RS.Open SqlStmt, DBConn
    If RS.EOF Then
        SqlStmt = "SELECT cdval1, text1 FROM " & T_LAB034 & " WHERE " & DBW("cdindex = ", LC4_Remark)
    End If
    Set RS = Nothing
    
    With objCodeList
        .Connection = DBConn
        .FormCaption = "Remark"
        .ColumnHeaderText = "Code;Remark"
'        .HideColumnHeaders = True
        .ColumnHeaderWidth = "840.189;5309.858"
        .FormHeight = 3105
        .FormWidth = 6605
        .HideSearchTool = True
        .SelectByClick = True
        .Tag = "Remark"
        .LoadPopUp SqlStmt
        
'        .ListCaption = "Remark"
'        .ListColHeader = "Code" & vbTab & "Remark"
'        .Top = Me.cmdRemarkTemplete.Top + 5700
'        .Left = Me.cmdRemarkTemplete.Left + 2000
'        .Width = 6250
'        .Height = 3000
'        .Tag = "Remark"
'        .CaptionOn = True
'        .MultiSel = False
'        .PopupList SqlStmt, 2
'        .ListAdd vbTab & "< �� �� > ", 2, 1
    End With

End Sub
Private Function DiffSaveCheck() As Boolean
    '===================================================================
    'DIFF COUNT CHECK
    '�����Ϳ� DIFF COUNT �ڵ忡 ��ϵ� �ڵ��� ���� 100�� �ƴϸ� �ȵȴ�.
    'S2LAB032 �� CDINDEX=LC3_WBCDiffCode �̸� �˻��ڵ�� CDVAL1��
    '�ش� CDVAL1�� ��� ���� ���� 100�� �ƴϸ� �ȵ˴ϴ�.
    '===================================================================
    Dim objDic As New clsDictionary
    Dim SSQL   As String
    Dim RS     As Recordset
    Dim ii     As Long
    
    Dim sValue As String
    Dim tValue As String
    Dim blnCheck As Boolean
    
    objDic.Clear
    objDic.FieldInialize "testcd", "rstcd"
    SSQL = objPtInfo.DiffCountSQL
    tValue = "0"
    
    blnCheck = False
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    If Not RS.EOF Then
        Do Until RS.EOF
            objDic.AddNew RS.Fields("cdval1").Value & "", ""
            RS.MoveNext
        Loop
        For ii = 1 To ssRst.MaxRows
            With objPtInfo.Result.Item(ii)
                If objDic.Exists(.TestCd) Then
                    If .SpcCd = P_DiffSpcCd Then
                        blnCheck = True
                        objDic.KeyChange .TestCd
                        ssRst.Row = ii
                        ssRst.Col = objPtInfo.SSCol("RESULT")
                        objDic.Fields("rstcd") = ssRst.Value
                    End If
                End If
            End With
        Next
        objDic.MoveFirst
        Do Until objDic.EOF
            tValue = CDbl(tValue) + Val(objDic.Fields("rstcd"))
            objDic.MoveNext
        Loop
        
        If blnCheck = True And CDbl(tValue) <> 100 Then
            MsgBox "Diff Count ����Է¿����Դϴ�." & _
                   "�Է� ���հ�� " & tValue & " �Դϴ�.", vbCritical + vbOKOnly, "������ ����"
            Set RS = Nothing
            Set objDic = Nothing
            Exit Function
        End If
    End If
    DiffSaveCheck = True
    Set RS = Nothing
    Set objDic = Nothing
End Function
Private Sub cmdSave_Click()
    
    Dim ii As Long
    Dim blnDBSuccess As Boolean
    Dim objLvwItem As MSComctlLib.ListItem
    Dim intLvwCount As Integer
    Dim strYesNo    As String
    Dim strWorkArea As String
    Dim strAccDt    As String
    Dim strAccSeq   As String
    
    'WBC DIFF COUNT ���üũ
    If P_DiffFg Then
        If DiffSaveCheck = False Then
            strYesNo = MsgBox("�������� �Ͻðڽ��ϱ�?.", vbInformation + vbYesNo, "������")
            If strYesNo = vbNo Then Exit Sub
        End If
    End If
    
   
    With objPtInfo
        .FootNote = rtfComment.Text
        .Result.Item(ssRst.ActiveRow).TextRst = rtfText.Text
    End With
    '/*
    For ii = 1 To ssRst.MaxRows
        With objPtInfo.Result.Item(ii)
            ssRst.Row = ii
            ssRst.Col = objPtInfo.SSCol("RESULT")
             If UCase(ssRst.Value) = UCase(CS_EqpError) Then
                ssRst.Action = ActionActiveCell
                Exit Sub
            End If
            If .TxtType = "2" And .RstDiv = "R" Then
                If .TextRst = "" Or ssRst.Value = "" Then
                'If (ssRst.Value <> "" AND .TextRst = "") _
                    Or (ssRst.Value = "" AND .TextRst <> "") Then
                      '�˻�� �Ϲݰ���� �ؽ�Ʈ ����� ���� �Է¿�. ������� ó��.
                    ssRst.Col = objPtInfo.SSCol("EC")
                    ssRst.Value = 1
                End If
            End If
        End With
    Next ii
   '
    blnDBSuccess = objPtInfo.DataEntry 'objPtInfo                  '�������� �����Ѵ�.
    If blnDBSuccess = False Then
        lblErr.Caption = "�ڷẸ�� Error!"
        Exit Sub
    Else
        lblErr.Caption = "�ڷᰡ ���������� �����Ǿ����ϴ�."
    End If

    If objPtInfo.StsCd = enStsCd.StsCd_LIS_FinRst Then
        If P_RealPrinter = True Then
        '����� ���޽� ������ ������
'            DoEvents
            With lvwEQP
                strWorkArea = medGetP(.SelectedItem.ListSubItems(3).Text, 1, "-")
                strAccDt = Mid(Format(GetSystemDate, "YYYY"), 1, 2) & medGetP(.SelectedItem.ListSubItems(3).Text, 2, "-")
                strAccSeq = medGetP(.SelectedItem.ListSubItems(3).Text, 3, "-")
        
                Call PrintEROP24(strWorkArea, strAccDt, strAccSeq)
            End With
'            DoEvents
        End If
        
        ssRst.MaxRows = 0
        lvwPatient.ListItems.Clear
        rtfText.Text = ""
        rtfComment.Text = ""
        rtfRemark.Text = ""
        With lvwEQP
            intLvwCount = .ListItems.Count
            For ii = 1 To .ListItems.Count
                If .ListItems.Item(ii).Selected = True Then
                    .ListItems.Remove (ii)
                    Exit For
                End If
            Next ii
            If intLvwCount = .ListItems.Count Then
                For ii = 1 To .ListItems.Count
                    If .ListItems.Item(ii).SubItems(1) = objPtInfo.AccNo Then
                        .ListItems.Remove (ii)
                        Exit For
                    End If
                Next ii
            End If
        End With
        IndexPointer = IndexPointer - 1
        If lvwEQP.ListItems.Count = IndexPointer Then IndexPointer = IndexPointer - 1
    End If
   '
    If lvwEQP.ListItems.Count = IndexPointer Then
        Set objLvwItem = lvwEQP.ListItems(IndexPointer)
        objLvwItem.SubItems(2) = " "
       IndexPointer = 0
    End If
    '
    If lvwEQP.ListItems.Count = 0 Then
        ClearData
    Else
        gblnModify = False
        Call cmdNext_Click
        'Set objLvwItem = lvwEQP.ListItems.Item(1)
        'lvwEQP_ItemClick objLvwItem
    End If
   '
End Sub

Private Sub PrintEROP24(ByVal LastWorkArea As String, ByVal LastAccDt As String, ByVal LastAccSeq As String)
    Dim RS          As Recordset
    Dim objReport   As clsBatchReport
    Dim objSQL      As clsLISSqlReport
    Dim objDisease  As S2LIS_ReportLib.clsDisease
    Dim picESign    As Object
    Dim strSQL      As String
    Dim strEmpId    As String
    Dim strAge      As String
    Dim strWardID   As String
        
    Set objSQL = New clsBatchReport
    Set objReport = New clsLISSqlReport
    Set objDisease = New S2LIS_ReportLib.clsDisease
    
   '�ڵ���±�� ����Ŭ��������
    strSQL = " SELECT a.ptid,a.workarea,a.accdt,a.accseq,a.stscd,a.vfydt,a.vfytm, " & _
             "        d." & F_PTNM & " as ptnm, " & F_DOB2("d") & " as dob, d." & F_SEX & " as sex, " & _
             "        c.deptcd, c.wardid, c.majdoct " & _
             "   FROM " & T_HIS001 & " d, " & T_LAB101 & " c, " & T_LAB102 & " b, " & T_LAB201 & " a " & _
             "  WHERE " & DBW("a.workarea", LastWorkArea, 2) & _
             "    AND " & DBW("a.accdt", LastAccDt, 2) & _
             "    AND " & DBW("a.accseq", LastAccSeq, 2) & _
             "    AND a.reqtotcnt=a.reqinputcnt " & _
             "    AND a.workarea=b.workarea AND a.accdt=b.accdt AND a.accseq=b.accseq " & _
             "    AND b.ptid=c.ptid AND b.orddt=c.orddt AND b.ordno=c.ordno " & _
             "    AND c.deptcd in ('ER','24') AND b.ptid = d." & F_PTID
    Set RS = New Recordset
    RS.Open strSQL, DBConn
    
    If RS.RecordCount > 0 Then '
        
        With objReport
            .PtId = RS.Fields("ptid").Value & ""
            .ptnm = RS.Fields("ptnm").Value & ""
            .PtSex = RS.Fields("sex").Value & ""
            strAge = RS.Fields("dob").Value & ""
            If Len(strAge) = 6 Then strAge = strAge & "01"
            .PtAge = DateDiff("yyyy", Format(strAge, CS_DateMask), GetSystemDate)
                
            .FromDt = RS.Fields("vfydt").Value & ""
            .ToDt = RS.Fields("vfydt").Value & ""
            
            If Trim(RS.Fields("bussdiv").Value & "") = "1" Then
                .Dept = RS.Fields("deptcd").Value & ""
            Else
                .Dept = RS.Fields("wardid").Value & ""
            End If
            
            .VfyDt = RS.Fields("vfydt").Value & ""
'            If objLisComCode.DeptCd.Exists(Rs.Fields("deptcd").Value & "") Then
'                objLisComCode.DeptCd.KeyChange (Rs.Fields("deptcd").Value & "")
                .DeptNm = GetDeptNm(RS.Fields("deptcd").Value & "") 'objLisComCode.DeptCd.Fields("deptnm")
'            End If
            
            strWardID = Trim(RS.Fields("wardid").Value & "")
            
            .WardId = strWardID
            
'            If objLisComCode.WardId.Exists(strWardID) = True Then
'                objLisComCode.WardId.KeyChange (strWardID)
                objReport.Ward = GetWardNm(strWardID) 'objLisComCode.WardId.Fields("wardnm")
'            End If
            .Doct = GetEmpNm(RS.Fields("majdoct").Value & "")

            
            .VfyNM = ObjMyUser.EmpLngNm
            .MdfDt = ""
            objDisease.PtId = RS.Fields("ptid").Value
            .ICD = objDisease.Disease
            Call .ReportForOneERPatient(RS.Fields("ptid").Value & "", RS.Fields("vfydt").Value & "", _
                                   RS.Fields("workarea").Value & "", RS.Fields("accdt").Value & "", _
                                   RS.Fields("accseq").Value & "", picESign, RS.Fields("vfydt").Value & "", _
                                   RS.Fields("vfydt").Value & "")
        End With
    End If
    
    If P_PrinterChkFg = True Then
        strSQL = " update " & T_LAB302 & _
                 "    set " & _
                            DBW("rptfg", "Y", 3) & _
                            DBW("rptdt", Format(GetSystemDate, "YYYYMMDD"), 3) & _
                            DBW("rptid", ObjSysInfo.EmpId, 2) & _
                 "  WHERE " & DBW("a.workarea", LastWorkArea, 2) & _
                 "    AND " & DBW("a.accdt", LastAccDt, 2) & _
                 "    AND " & DBW("a.accseq", LastAccSeq, 2)
        
        DBConn.Execute strSQL
    End If
    
NoData:
    Set RS = Nothing
    Set objSQL = Nothing
    Set objReport = Nothing
    Set objDisease = Nothing
End Sub

'Private Function GetEmpNm(ByVal vEmpID As String) As String
'    Dim objData As New clsBasisData
'
'    GetEmpNm = objData.GetEmpNm(vEmpID)
'    Set objData = Nothing
'End Function

'Private Function GetDeptNm(ByVal vDeptCd As String) As String
'    Dim objData As New clsBasisData
'
'    GetDeptNm = objData.GetDeptNm(vDeptCd)
'    Set objData = Nothing
'End Function

'Private Function GetWardNm(ByVal vWardId As String) As String
'    Dim objData As New clsBasisData
'
'    GetWardNm = objData.GetWardNm(vWardId)
'    Set objData = Nothing
'End Function

Private Sub cmdTextTemplete_Click()
    If rtfText.Enabled = False Then Exit Sub
    Call CallTemplete(2, 0)
End Sub

Private Sub cmdCommentTemplete_Click()
    If ssRst.MaxRows < 1 Then Exit Sub
    Call CallTemplete(3, 0)
End Sub

Private Sub cmdTrans_Click()
   '
    If objPtInfo Is Nothing Then
       Set objPtInfo = New clsPatientInfo
    End If
    lstEQCode.Visible = False
    
    '2001-11-07 �߰� : ���� ������۳��� ���� (�Ⱓ : 1����)
    Screen.MousePointer = vbArrowHourglass
    lblErr.Caption = "������ ������� ������ �����ϰ� �ֽ��ϴ�."
    Call objPtInfo.EqpHistoryDelete(txtEqpCd.Text, Format(DateAdd("d", -30, Now), CS_DateDbFormat))
    lblErr.Caption = ""
    Screen.MousePointer = vbDefault
    
    TrasferListPop txtEqpCd.Text
   '
End Sub

Private Sub Command1_Click()
    Dim ServerName   As String
    Dim DatabaseName As String
    Dim UserName     As String
    Dim Password     As String
    Dim strTransCd   As String
    Dim strDoctCd    As String
    Dim strTransDt   As String
    Dim strTransStatus As String
    Dim strTansEtc   As String
    Dim strMessage   As String
    Dim strTransNo   As String
    Dim strDoctNo    As String
    Dim strSQL       As String
    Dim strDeptNm    As String
    Dim strTranNm    As String
    Dim strSMSIP     As String
    Dim strBackNo    As String
    Dim strTmpTestCd As String
    Dim strMaDtId  As String
    Dim strMaTransNo As String
    
    Set AdoCn_ORACLE = New ADODB.Connection
    
    On Error Resume Next    '2013-09-11 PSK
    
    With AdoCn_ORACLE
        .ConnectionTimeout = 25
'        .Provider = "OraOLEDB.Oracle.1"
        .Provider = "MSDAORA.1"                 ' Oracle "MSDAORA.1"
        .Properties("Data Source").Value = "PMC"
        .Properties("Persist Security Info") = True
        .Properties("User ID").Value = "oral1"
        .Properties("Password").Value = "oral1"
        .Open
    End With
           
    Set AdoRs_ORACLE = New ADODB.Recordset
        
    strSQL = ""
    strSQL = "SELECT * FROM S2lab032  "
    strSQL = strSQL + " WHERE cdindex = 'C232'"
    strSQL = strSQL + "   AND cdval1 = 'SVR1'  "

    AdoRs_ORACLE.CursorLocation = adUseClient
    AdoRs_ORACLE.Open strSQL, AdoCn_ORACLE
    
    With AdoRs_ORACLE
        If .RecordCount > 0 Then
            strSMSIP = AdoRs_ORACLE.Fields("FIELD4") & ""
        Else
            strSMSIP = "172.16.200.37"
        End If
        .Close
    End With
    
    Set AdoCn_SQL = New ADODB.Connection

    ServerName = strSMSIP
    DatabaseName = "medicalCRM_jesus"
    UserName = "jesus"
    Password = "jesus"
   
    With AdoCn_SQL
        .ConnectionTimeout = 10
        .Provider = "SQLOLEDB"
        .Properties("Data Source").Value = ServerName
        .Properties("Initial Catalog").Value = DatabaseName
        .Properties("User ID").Value = UserName
        .Properties("Password").Value = Password
        Screen.MousePointer = vbHourglass
        .Open
    End With
    Screen.MousePointer = vbDefault
    
'    If txtDtNo.Text = "" Then
'        MsgBox "���Ź�ȣ�� �Է��ϼ���.", vbCritical + vbOKOnly, "���Ź�ȣ��� Message"
'        txtDtNo.SetFocus
'        Exit Sub
'    End If
    
    strTransCd = ObjSysInfo.EmpId
    strTransNo = txtTransNo.Text
    strDoctCd = txtDtId.Text
    strMaDtId = txtExDtId.Text
    strMaTransNo = txtExDtNo.Text
    strTransDt = Format(Now, "YYYY-MM-DD HH:MM:SS")
    strDoctNo = txtDtNo.Text
    strTransStatus = "1"
    strTansEtc = "LIS"
    strDeptNm = txtDeptNm.Text
    strTranNm = txtTransNm.Text
    strMessage = rtfMessage.Text & vbCRLF & "- " & strTranNm
    strBackNo = "063-230-8753"
    strTmpTestCd = txtTestCd.Text
    
    If Len(strMessage) > 80 Then
        MsgBox "�޽����� ũ�⸦ �ٿ��ּ���.", vbCritical + vbOKOnly, "�޽���������� Message"
        rtfMessage.SetFocus
        Exit Sub
    End If
    
    strSQL = ""
    strSQL = strSQL & " INSERT INTO em_tran (TRAN_ID, TRAN_PHONE, TRAN_CALLBACK, TRAN_MSG, TRAN_DATE, TRAN_STATUS, TRAN_ETC1)"
    strSQL = strSQL & " values('" & strTransCd & "' ,"
    strSQL = strSQL & "        '" & strDoctNo & "' ,"
    strSQL = strSQL & "        '" & strBackNo & "' ,"
    strSQL = strSQL & "        '" & strMessage & "' ,"
    strSQL = strSQL & "        '" & strTransDt & "' ,"
    strSQL = strSQL & "        '" & strTransStatus & "' ,"
    strSQL = strSQL & "        '" & strTansEtc & "')"
    
    AdoCn_SQL.Execute strSQL
    
    ' �˻��ڵ� �߰�
    ' 2019-05-03 SMS ��ȸ �˻��ڵ�� ��ȸ ��
    
    strSQL = ""
    strSQL = strSQL & " INSERT INTO S2COM102 (TRANSDT, TRANSID, TELNO, DOCTID, DOCTNM, DEPTNM, TRANSMSG, RCVSTAT, REMARK, RCVDT, TESTCD)"
    strSQL = strSQL & " values('" & strTransDt & "' ,"
    strSQL = strSQL & "        '" & strTransCd & "' ,"
    strSQL = strSQL & "        '" & strDoctNo & "' ,"
    strSQL = strSQL & "        '" & Trim(txtDtId.Text) & "' ,"
    strSQL = strSQL & "        '" & Trim(txtDtNm.Text) & "' ,"
    strSQL = strSQL & "        '" & strDeptNm & "' ,"
    strSQL = strSQL & "        '" & strMessage & "' ,"
    strSQL = strSQL & "        '����' ,"
    strSQL = strSQL & "        '" & strTransNo & "',"
    strSQL = strSQL & "        '" & strRcvDt & "',"
    strSQL = strSQL & "        '" & strTmpTestCd & "')"
    
    AdoCn_ORACLE.Execute strSQL
    
    strSQL = ""
    strSQL = strSQL & " INSERT INTO MDNOTIFT (RECVID, NOTIDATE, SEQNO, NOTITYPE, SENDDATE, TITLE, CONTENTS, SENDID, WORKAREA)"
    strSQL = strSQL & " (select '" & strDoctCd & "' ,"
    strSQL = strSQL & "        TO_DATE(TO_CHAR(sysdate, 'yyyymmdd'),'yyyymmdd'),"
    strSQL = strSQL & "        NVL(Max(SEQNO), 0) + 1,"
    strSQL = strSQL & "        '7' ,"
    strSQL = strSQL & "        SYSDATE ,"
    strSQL = strSQL & "        '[CVR(�̻�������)]' ,"
    strSQL = strSQL & "        '" & strMessage & "' ,"
    strSQL = strSQL & "        '" & strTransCd & "', '" & strTransNo & "' from mdnotift where recvid = '" & strDoctCd & "' and notidate = TO_DATE(TO_CHAR(sysdate, 'yyyymmdd'),'yyyymmdd'))"
    
    AdoCn_ORACLE.Execute strSQL
    
    If Trim(txtDtId.Text) <> Trim(txtExDtId.Text) Then
        strSQL = ""
        strSQL = strSQL & " INSERT INTO em_tran (TRAN_ID, TRAN_PHONE, TRAN_CALLBACK, TRAN_MSG, TRAN_DATE, TRAN_STATUS, TRAN_ETC1)"
        strSQL = strSQL & " values('" & strTransCd & "' ,"
        strSQL = strSQL & "        '" & txtExDtNo.Text & "' ,"
        strSQL = strSQL & "        '" & strBackNo & "' ,"
        strSQL = strSQL & "        '" & strMessage & "' ,"
        strSQL = strSQL & "        '" & strTransDt & "' ,"
        strSQL = strSQL & "        '" & strTransStatus & "' ,"
        strSQL = strSQL & "        '" & strTansEtc & "')"
        
        AdoCn_SQL.Execute strSQL
        
        ' �˻��ڵ� �߰�
        ' 2019-05-03 SMS ��ȸ �˻��ڵ�� ��ȸ ��
        
        strSQL = ""
        strSQL = strSQL & " INSERT INTO S2COM102 (TRANSDT, TRANSID, TELNO, DOCTID, DOCTNM, DEPTNM, TRANSMSG, RCVSTAT, REMARK, RCVDT, TESTCD)"
        strSQL = strSQL & " values('" & strTransDt & "' ,"
        strSQL = strSQL & "        '" & strTransCd & "' ,"
        strSQL = strSQL & "        '" & txtExDtNo.Text & "' ,"
        strSQL = strSQL & "        '" & Trim(txtExDtId.Text) & "' ,"
        strSQL = strSQL & "        '" & Trim(txtExDtNm.Text) & "' ,"
        strSQL = strSQL & "        '" & strDeptNm & "' ,"
        strSQL = strSQL & "        '" & strMessage & "' ,"
        strSQL = strSQL & "        '����' ,"
        strSQL = strSQL & "        '" & strTransNo & "',"
        strSQL = strSQL & "        '" & strRcvDt & "',"
        strSQL = strSQL & "        '" & strTmpTestCd & "')"
        
        AdoCn_ORACLE.Execute strSQL
        
        strSQL = ""
        strSQL = strSQL & " INSERT INTO MDNOTIFT (RECVID, NOTIDATE, SEQNO, NOTITYPE, SENDDATE, TITLE, CONTENTS, SENDID, WORKAREA)"
        strSQL = strSQL & " (select '" & strMaDtId & "' ,"
        strSQL = strSQL & "        TO_DATE(TO_CHAR(sysdate, 'yyyymmdd'),'yyyymmdd'),"
        strSQL = strSQL & "        NVL(Max(SEQNO), 0) + 1,"
        strSQL = strSQL & "        '7' ,"
        strSQL = strSQL & "        SYSDATE ,"
        strSQL = strSQL & "        '[CVR(�̻�������)]' ,"
        strSQL = strSQL & "        '" & strMessage & "' ,"
        strSQL = strSQL & "        '" & strTransCd & "', '" & strTransNo & "' from mdnotift where recvid = '" & strMaDtId & "' and notidate = TO_DATE(TO_CHAR(sysdate, 'yyyymmdd'),'yyyymmdd'))"
        
        AdoCn_ORACLE.Execute strSQL
    End If
    
    strRcvDt = ""
    
    frmSMS.Visible = False
    Set AdoCn_SQL = Nothing
    Set AdoCn_ORACLE = Nothing

'    Dim ServerName   As String
'    Dim DatabaseName As String
'    Dim UserName     As String
'    Dim Password     As String
'    Dim strTransCd   As String
'    Dim strDoctCd    As String
'    Dim strTransDt   As String
'    Dim strTransStatus As String
'    Dim strTansEtc   As String
'    Dim strMessage   As String
'    Dim strTransNo   As String
'    Dim strDoctNo    As String
'    Dim strSQL       As String
'    Dim strDeptNm    As String
'    Dim strTranNm    As String
'    Dim strSMSIP     As String
'    Dim strBackNo    As String
'
'    Set AdoCn_ORACLE = New ADODB.Connection
'
'    On Error Resume Next    '2013-09-11 PSK
'
'    With AdoCn_ORACLE
'        .ConnectionTimeout = 25
''        .Provider = "OraOLEDB.Oracle.1"
'        .Provider = "MSDAORA.1"                 ' Oracle "MSDAORA.1"
'        .Properties("Data Source").Value = "PMC"
'        .Properties("Persist Security Info") = True
'        .Properties("User ID").Value = "oral1"
'        .Properties("Password").Value = "oral1"
'        .Open
'    End With
'
'    Set AdoRs_ORACLE = New ADODB.Recordset
'
'    strSQL = ""
'    strSQL = "SELECT * FROM S2lab032  "
'    strSQL = strSQL + " WHERE cdindex = 'C232'"
'    strSQL = strSQL + "   AND cdval1 = 'SVR1'  "
'
'    AdoRs_ORACLE.CursorLocation = adUseClient
'    AdoRs_ORACLE.Open strSQL, AdoCn_ORACLE
'
'    With AdoRs_ORACLE
'        If .RecordCount > 0 Then
'            strSMSIP = AdoRs_ORACLE.Fields("FIELD4") & ""
'        Else
'            strSMSIP = "172.16.200.37"
'        End If
'        .Close
'    End With
'
'
'    Set AdoCn_SQL = New ADODB.Connection
'
'    ServerName = strSMSIP
'    DatabaseName = "medicalCRM_jesus"
'    UserName = "jesus"
'    Password = "jesus"
'
'    With AdoCn_SQL
'        .ConnectionTimeout = 10
'        .Provider = "SQLOLEDB"
'        .Properties("Data Source").Value = ServerName
'        .Properties("Initial Catalog").Value = DatabaseName
'
'        .Properties("User ID").Value = UserName
'        .Properties("Password").Value = Password
'
'        Screen.MousePointer = vbHourglass
'        .Open
'    End With
'    Screen.MousePointer = vbDefault
'
''    If txtDtNo.Text = "" Then
''        MsgBox "���Ź�ȣ�� �Է��ϼ���.", vbCritical + vbOKOnly, "���Ź�ȣ��� Message"
''        txtDtNo.SetFocus
''        Exit Sub
''    End If
'
'    strTransCd = ObjSysInfo.EmpId
'    strTransNo = txtTransNo.Text
'    strDoctCd = txtDtId.Text
'    strTransDt = Format(Now, "YYYY-MM-DD HH:MM:SS")
'    strDoctNo = txtDtNo.Text
'    strTransStatus = "1"
'    strTansEtc = "LIS"
'    strBackNo = "063-230-8753"
'    strDeptNm = txtDeptNm.Text
'    strTranNm = txtTransNm.Text
'    strMessage = rtfMessage.Text & vbCRLF & "- " & strTranNm
'
'    If Len(strMessage) > 80 Then
'        MsgBox "�޽����� ũ�⸦ �ٿ��ּ���.", vbCritical + vbOKOnly, "�޽���������� Message"
'        rtfMessage.SetFocus
'        Exit Sub
'    End If
'
'    strSQL = ""
'    strSQL = strSQL & " INSERT INTO em_tran (TRAN_ID, TRAN_PHONE, TRAN_CALLBACK, TRAN_MSG, TRAN_DATE, TRAN_STATUS, TRAN_ETC1)"
'    strSQL = strSQL & " values('" & strTransCd & "' ,"
'    strSQL = strSQL & "        '" & strDoctNo & "' ,"
'    strSQL = strSQL & "        '" & strBackNo & "' ,"
'    strSQL = strSQL & "        '" & strMessage & "' ,"
'    strSQL = strSQL & "        '" & strTransDt & "' ,"
'    strSQL = strSQL & "        '" & strTransStatus & "' ,"
'    strSQL = strSQL & "        '" & strTansEtc & "')"
'
'    AdoCn_SQL.Execute strSQL
'
'    strSQL = ""
'    strSQL = strSQL & " INSERT INTO S2COM102 (TRANSDT, TRANSID, TELNO, DOCTID, DOCTNM, DEPTNM, TRANSMSG, RCVSTAT, REMARK, RCVDT)"
'    strSQL = strSQL & " values('" & strTransDt & "' ,"
'    strSQL = strSQL & "        '" & strTransCd & "' ,"
'    strSQL = strSQL & "        '" & strDoctNo & "' ,"
'    strSQL = strSQL & "        '" & Trim(txtDtNm.Text) & "' ,"
'    strSQL = strSQL & "        '' ,"
'    strSQL = strSQL & "        '" & strDeptNm & "' ,"
'    strSQL = strSQL & "        '" & strMessage & "' ,"
'    strSQL = strSQL & "        '����' ,"
'    strSQL = strSQL & "        '" & strTransNo & "',"
'    strSQL = strSQL & "        '" & strRcvDt & "')"
'
'    AdoCn_ORACLE.Execute strSQL
'
''    strSQL = ""
''    strSQL = strSQL & " INSERT INTO MDNOTIFT (RECVID, NOTIDATE, SEQNO, NOTITYPE, SENDDATE, TITLE, CONTENTS, SENDID)"
''    strSQL = strSQL & " (select '" & strDoctCd & "' ,"
''    strSQL = strSQL & "        TO_DATE(TO_CHAR(sysdate, 'yyyymmdd'),'yyyymmdd'),"
''    strSQL = strSQL & "        NVL(Max(SEQNO), 0) + 1,"
''    strSQL = strSQL & "        '7' ,"
''    strSQL = strSQL & "        SYSDATE ,"
''    strSQL = strSQL & "        '[CVR(�̻�������)]' ,"
''    strSQL = strSQL & "        '" & strMessage & "' ,"
''    strSQL = strSQL & "        '" & strTransCd & "' from mdnotift where recvid = '" & strDoctCd & "' and notidate = TO_DATE(TO_CHAR(sysdate, 'yyyymmdd'),'yyyymmdd'))"
''
''    AdoCn_ORACLE.Execute strSQL
'
'     strSQL = ""
'    strSQL = strSQL & " INSERT INTO MDNOTIFT (RECVID, NOTIDATE, SEQNO, NOTITYPE, SENDDATE, TITLE, CONTENTS, SENDID, WORKAREA)"
'    strSQL = strSQL & " (select '" & strDoctCd & "' ,"
'    strSQL = strSQL & "        TO_DATE(TO_CHAR(sysdate, 'yyyymmdd'),'yyyymmdd'),"
'    strSQL = strSQL & "        NVL(Max(SEQNO), 0) + 1,"
'    strSQL = strSQL & "        '7' ,"
'    strSQL = strSQL & "        SYSDATE ,"
'    strSQL = strSQL & "        '[CVR(�̻�������)]' ,"
'    strSQL = strSQL & "        '" & strMessage & "' ,"
'    strSQL = strSQL & "        '" & strTransCd & "', '" & strTransNo & "' from mdnotift where recvid = '" & strDoctCd & "' and notidate = TO_DATE(TO_CHAR(sysdate, 'yyyymmdd'),'yyyymmdd'))"
'
'    AdoCn_ORACLE.Execute strSQL
'
'    strRcvDt = ""
'    frmSMS.Visible = False
'    Set AdoCn_SQL = Nothing
'    Set AdoCn_ORACLE = Nothing
    
End Sub

Private Sub Form_Activate()
   '
    If blnFirst = False Then
        Call LoadLvwHead
        blnFirst = True
        ClearData
    End If
    '
    If objLab306 Is Nothing Then
        Set objLab306 = New clsEquipTransfer
        objLab306.LoadTable "", ""
    End If
   '
    '��������� ���ð˻�(�̻���/Ư����ȸ����)
    If P_RealTestMicSpecial = True Then fraCul.Visible = True

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        If lstEQCode.Visible Then
            lstEQCode.Visible = False
            txtEqpCd.SetFocus
        End If
        If Not objCodeList Is Nothing Then
            Set objCodeList = Nothing
            mskSpcNo.SetFocus
        End If
    End If
End Sub

Private Sub Form_Load()

    chkUr.Visible = True
    chkUr.Caption = "��ü"
        
    blnFirst = False
    gblnModify = False
    
     '
    gstrMsk = String(P_SpcYyLength, "_") & "-" & String(P_SpcNoLength, "_")      ' "_-________"
    prgRst.Align = vbAlignBottom
    prgRst.Visible = False
    ssRst.RowHeight(-1) = 12.5
    '
    Set objPtInfo = New clsPatientInfo
    
    Me.Show
    Call cmdClear_Click
    Call ShowEqpList
    KeyPreview = True
    frmSMS.Visible = False
End Sub

Private Sub ShowEqpList()
    Dim FNo As Long
    Dim FName As String
    Dim i As Long
    Dim strData As String
    Dim strTemp As String
    
    FNo = FreeFile
    
    On Error GoTo ErrList
    
    If Dir(App.Path & "\LIS.dat") = "" Then
        MsgBox "������ ��� �����ϴ�.", vbExclamation
        Exit Sub
    End If
    
    Open App.Path & "\LIS.dat" For Input As #FNo
    
    lstEQCode.Clear
    Do While Not EOF(1)
        Line Input #FNo, strTemp
        
        strData = DECrypt(strTemp)
        
        lstEQCode.AddItem Trim(Mid(strData, 1, 10)) & vbTab & Trim(Mid(strData, 11)) & vbTab, i
        'lstEQCode.BackColor = vbRed
        i = i + 1
    Loop
    Close #FNo
    
    If lstEQCode.ListCount = 0 Then
        MsgBox "������ ��� �����ϴ�.", vbCritical
    End If
    
    Exit Sub
ErrList:
    MsgBox Err.Description, vbExclamation
    On Error Resume Next
    Close #FNo
End Sub

Private Sub clsTemplete_CopyTemplete()
   '
    If ssRst.MaxRows < 1 Then Exit Sub
    With objPtInfo
        Select Case gintTemplete
            Case 1:
                If clsTemplete.rtfText.Text <> "" Then
                    rtfRemark.Text = clsTemplete.rtfText.Text
                    .RmkCd = frm230TempSearch.lblCode.Caption
                    .RmkNm = rtfRemark.Text
                Else
                    rtfRemark.Text = ""
                    .RmkCd = ""
                    .RmkNm = ""
                End If
            Case 2:
                rtfText.Text = clsTemplete.rtfText.Text
                .Result.Item(ssRst.ActiveRow).TextRst = rtfText.Text
                rtfText.SetFocus
            Case 3:
                rtfComment.Text = clsTemplete.rtfText.Text
                .FootNote = rtfComment.Text
                rtfComment.SetFocus
        End Select
    End With
    
End Sub

Private Sub CallTemplete(ByVal pintPrg As Integer, ByVal pintMode As Integer)
    
    Dim strTitle As String
    Dim strWorkArea As String
   
    strWorkArea = medGetP(lvwPatient.ListItems.Item(1).Text, 1, "-")
    Set clsTemplete = frm230TempSearch
    strTitle = Choose(pintPrg, "Remark", "Text Result", "Foot Note")
    With clsTemplete
        .qField1 = strWorkArea
        .Show
        If pintMode = 0 Then
            .lblName.Caption = "Edit " & strTitle
        Else
            .lblName.Caption = "Modify " & strTitle
        End If
        .Caption = strTitle & " " & "Templete Editor"
        .lblInfo.Caption = pintMode & "$" & pintPrg
        Select Case pintPrg
            Case 1:
                .lblCode.Caption = objPtInfo.RmkCd
                .rtfText = rtfRemark.Text
            Case 2:
                .rtfText = rtfText.Text
            Case 3:
                .rtfText = rtfComment.Text
        End Select
    End With
    gintTemplete = pintPrg
    
End Sub

Private Sub LoadLvwHead()
    
    Dim colHead As ColumnHeader
    Dim intMode As Integer
    
    '������ ���� ���
    intMode = 1         'Korea
    'intMode = 2         'English
    If intMode = 1 Then
        medInitLvwHead lvwPatient, "������ȣ,ȯ��ID,ȯ�ڼ���,��/����,�������,����,��ġ��,��ü,��������,���(�ܺ�QC)", _
           "400,-100,100,-450,0,50,-150,150,0"
'        medInitLvwHead lvwPatient, "������ȣ,ȯ��ID,ȯ�ڼ���,��/����,�������,����,��ġ��,��ü,��������", _
'           "400,-100,100,-450,0,50,-150,150"
        medInitLvwHead lvwEQP, "��ü��ġ,��ü��ȣ,,������ȣ,���޿���", _
           "650,700,-250,-500,-100"
        'InitLvwHead lvwEQP, "�����Ͻ�,��ü��ȣ,,������ȣ,���޿���", _
           "800,700,-250,-500,-100"
    Else
        medInitLvwHead lvwPatient, "Accession#,Patient ID,Patient Name,Sex/Age,Location,Physician", _
           "200,-100,200,-400,0,100,0"
        medInitLvwHead lvwEQP, "Transfer Date/Time,Spc No,,Accession No,StatFg", _
           "-185,185,0,400,-100"
    End If
   '
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objCuM = Nothing
    Set clsTemplete = Nothing
    Set objCodeList = Nothing
    Set objLab306 = Nothing
    Set objPtInfo = Nothing
    Call ICSPatientMark
End Sub

Private Sub lstEQCode_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Call lstEQCode_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub lstEQCode_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If lstEQCode.Enabled Then lstEQCode.SetFocus
End Sub

Private Sub mskSpcNo_GotFocus()
   '
    FocusMe Me.mskSpcNo
   '
End Sub

Private Sub mskSpcNo_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyDown Then Call cmdTrans_Click
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"

End Sub

Private Sub mskSpcNo_KeyPress(KeyAscii As Integer)
    Dim Char As String
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
End Sub

Private Sub mskSpcNo_Validate(Cancel As Boolean)
'    Dim ii As Integer
'    Dim strTmp As String
'
'   If Trim(mskSpcNo.ClipText) = "" Then
'      'Cancel = True
'      lblErr.Caption = ""
'      Exit Sub
'   End If
   '
'   strTmp = mskSpcNo.Text
'   If medGetP(mskSpcNo.Text, 3, "-") = "____" Then
'      mskSpcNo.Text = medGetP(strTmp, 1, "-") & "-" & _
'      medGetP(strTmp, 2, "-") & "-1___"
'   End If
   '
'   txtEqpCd.SetFocus

   '
End Sub

Private Sub objCodeList_SelectedItem(ByVal pSelectedItem As String)
Dim strTmp As String
   '
'   If Not IsNull(pSelectedItem) And pSelectedItem <> "" Then
      Select Case objCodeList.Tag
         Case "Transfer":
            If Not IsNull(pSelectedItem) And pSelectedItem <> "" Then
                strTmp = medGetP(pSelectedItem, 2, ";")
                strTmp = strTmp & String(12 - Len(strTmp), "_")
                mskSpcNo.Text = strTmp
                cmdQuery.SetFocus
            End If
         Case "Remark":
            objPtInfo.RmkCd = objCodeList.SelectedItems(0)
            objPtInfo.RmkNm = objCodeList.SelectedItems(1)
            rtfRemark.Text = objPtInfo.RmkNm
'            objPtInfo.RmkCd = medGetP(pSelectedItem, 1, ";")
'            If Trim(objPtInfo.RmkCd) <> "" Then
'                objPtInfo.RmkNm = medGetP(pSelectedItem, 2, ";")
'            Else
'                objPtInfo.RmkNm = ""
'            End If
'            rtfRemark.Text = objPtInfo.RmkNm
      End Select
'   End If
   Set objCodeList = Nothing
End Sub

Private Sub rtfText_LostFocus()
   '
    objPtInfo.Result.Item(ssRst.ActiveRow).TextRst = rtfText.Text
   '
End Sub

Private Sub ssRst_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim varTmp As Variant
    Dim strTest As String
    Dim strResult As String
    Dim strTestCd As String
    Dim strSQL          As String
    Dim RS              As New Recordset
    
    txtTestCd.Text = ""
    
    With ssRst
        .GetText 1, Row, varTmp: strTest = Trim(varTmp)
        .GetText 2, Row, varTmp: strResult = Trim(varTmp)
    End With
    rtfMessage.Text = rtfMessage.Text & strTest & " : " & strResult & vbCRLF
    
    strTestCd = objPtInfo.Result.Item(ssRst.ActiveRow).TestCd
    
    ' 2019-05-03 �˻��ڵ� �߰�
    txtTestCd.Text = strTestCd
    
    Select Case strTestCd
        Case "B2021", "LB2021"
            If UCase(strResult) = "NEGATIVE" Then
                Call cmdSMS_Click
            End If
        Case "B2061"
            If UCase(strResult) = "POSITIVE" Then
                Call cmdSMS_Click
            End If
        Case "ABOC22"
            If UCase(strResult) = "NEGATIVE" Then
                strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '00051'"
                RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
    
                If RS.RecordCount > 0 Then
                    RS.MoveFirst
                    rtfComment.Text = RS.Fields("text1") & ""
                End If
                RS.Close
            End If
        Case "B2602"
            If Val(strTmpAge) < 20 Or InStr(UCase(strTmpAge), "D") > 0 Then
                If strTmpSex = "M" Then
                    strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = 'ALP20M'"
                Else
                    strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = 'ALP20F'"
                End If
                RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
    
                If RS.RecordCount > 0 Then
                    RS.MoveFirst
                    rtfComment.Text = RS.Fields("text1") & ""
                End If
                RS.Close
            End If
        Case "C4612"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '3029-1'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C35901"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '3063'"
            RS.Open strSQL, DBConn
            
            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "CZ394"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '0031'"
            RS.Open strSQL, DBConn
            
            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "CZ394D"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '0031'"
            RS.Open strSQL, DBConn
            
            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "E435A"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '3067'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "E435B"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '3067'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "E435C"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '3067'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "E435D"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '3067'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "E435E"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '3067'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "E435F"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '3067'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "E435G"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '3067'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C404"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '3062'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "S641"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '3064'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C3260"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '3065'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C3530"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '3068'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C3630"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '3069'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
' �ܺΰ˻� �ڸ�Ʈ ó�� 2016.03.28
        Case "27LB"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8116'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "27LC"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8191'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "27LN"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8051'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "B145"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8025'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "B1712"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8173'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "B2700"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8106'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C208"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8196'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C3241"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8004'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "B3380"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8230'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C3460"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8207'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C3470"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8002'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C3580"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8017'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C3600"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8210'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C3823"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8101'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C3931"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8192'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C432"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8182'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C450339"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8050'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C4503390"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8012'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C450401"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8206'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C45041"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8171'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C450410"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8111'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C450412"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8131'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C450413"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8183'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C450415"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8109'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C450416"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8250'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C452361"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8172'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C452363"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8200'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C452364"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8153'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C452399"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8015'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C4523991"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8047'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C468242"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8108'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C468244"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8112'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C468245"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8114'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C468246"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8151'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
' �ܺΰ˻� �ʼ��׸�
        Case "C468249"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8251'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C468246"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8151'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C468250"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8184'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C468251"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8013'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C468253"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8215'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C468344"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8113'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C468349"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8252'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C468350"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8185'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C468351"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8014'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C468353"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8216'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C472261"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8234'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C472361"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8235'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C474281"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8103'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "E300"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8100'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "E426"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8119'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "X146"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8170'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "S095"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8020'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "X700"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8048'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "X701"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8048'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "X728"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8030'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "X730"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8029'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "Z133"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8031'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "Z134"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8032'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "Z982"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '8162'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C23202"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '6001'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C23201"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '6001'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
'2016-07-11 �߰�
        Case "C4912"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '6020'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C4913"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '6020'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C468342"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '6012-1'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C468343"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '6017-1'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C549"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '6010-1'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C46901"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '6004'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "M724"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '6032-1'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "M724TB"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '6032-1'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C46501"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '6009'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C468241"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '6014'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "B568"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '3080'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "B0260"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '0029-1'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C474381"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '6011-1'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "H976"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '6019-1'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C4690596"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '6003'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "M724"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '6032-1'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "M724TB"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '6032-1'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C23202A"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '6001'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C23201A"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '6001'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
' 2018.04.25 ���������� COMMENT �߰�
        Case "S876"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '1500'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "S606"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '1501'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "RV1201G"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '1510'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "H977"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '1502'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "PNBPCR5G"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '1525'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "AFBR5957"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '1540'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "CREPCR5"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '1541'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C5154"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '1503'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "CY7512", "B4064", "S729TB", "X560TB", "B4021AC", "B4021D", "B4052A"

            rtfComment.Text = "������ȸ ����Դϴ�."
' 2018-09-13 �߰�
        Case "B2640", "C3942T", "C3842", "27EGER", "C4712ER"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '3070'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
' 2018-09-13 �߰�
        Case "B2640", "C3942T"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '3070'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
' �ڵ庯�� "C3842", "27EGER", "C4712ER
' 2020-01-28 ����
        Case "C3842", "27EGER", "C4712ER"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '3070-2'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "X274"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '3066'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C2285"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '3062-1'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C3400"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '3062-2'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "C2532"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '3062-3'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
        Case "CZ394HIM", "CZ394ERM"
            strSQL = "SELECT text1 FROM S2LAB034 A WHERE CDVAL1 = '3059-1'"
            RS.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly

            If RS.RecordCount > 0 Then
                RS.MoveFirst
                rtfComment.Text = RS.Fields("text1") & ""
            End If
            RS.Close
    End Select
    
End Sub

Private Sub ssRst_EditChange(ByVal Col As Long, ByVal Row As Long)
    ssRst.Row = Row
    ssRst.Col = objPtInfo.SSCol("MAXCOL")
    ssRst.Value = ""
End Sub

Private Sub txtBatchRst_GotFocus()
    With txtBatchRst
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtBatchRst_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cmdApply.SetFocus
    End If
End Sub

Private Sub txtEqpCd_Change()
    mskSpcNo.Text = gstrMsk
End Sub

Private Sub txtEqpCd_GotFocus()
   '
    FocusMe Me.txtEqpCd
   '
End Sub

Private Sub txtEqpCd_KeyDown(KeyCode As Integer, Shift As Integer)

    If lstEQCode.ListCount = 0 Then Exit Sub
    If KeyCode = vbKeyDown Then
        lstEQCode.Visible = True
        Set objCodeList = Nothing
        lstEQCode.ListIndex = 0
        lstEQCode.ZOrder 0
        lstEQCode.SetFocus
    End If

End Sub

Private Sub txtEqpCd_KeyPress(KeyAscii As Integer)
    
    Dim Char As String
    
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
    If KeyAscii = vbKeyEscape Then Exit Sub
    If KeyAscii = vbKeyReturn Then
         Call lstEQCode_KeyDown(vbKeyReturn, 0)
         lstEQCode.Visible = False
         Exit Sub
    End If

    lstEQCode.Visible = True
    Set objCodeList = Nothing
    lstEQCode.ZOrder 0
    Call medCodeHelp(KeyAscii, lstEQCode, txtEqpCd.Text, txtEqpCd, mskSpcNo)

End Sub

Private Sub txtEqpCd_Validate(Cancel As Boolean)
   '

    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    If ActiveControl.Name = cmdClear.Name Then Exit Sub
    If ActiveControl.Name = cmdExit.Name Then Exit Sub
    
    IndexPointer = 0
    lblEqpCdNm.Caption = ""
    If Trim(txtEqpCd.Text) = "" Then
        'Cancel = True
        Exit Sub
    End If
   '
    Dim strEqpNm As String
    
    strEqpNm = objPtInfo.GetEqpName(txtEqpCd.Text)
    If Trim(strEqpNm) = "" Then
        MsgBox "�ڵ� �Է� Error!", vbCritical
        'txtEqpCd.Text = ""
        Cancel = True
        FocusMe Me.txtEqpCd
        Exit Sub
    End If
   '
    lblEqpCdNm.Caption = strEqpNm
    
   '
End Sub

Private Sub lstEQCode_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        txtEqpCd.Text = medGetP(lstEQCode.Text, 1, vbTab)
        lblEqpCdNm.Caption = medGetP(lstEQCode.Text, 2, vbTab)
        lstEQCode.Visible = False
        mskSpcNo.SetFocus
    End If

End Sub

Private Sub lvwEQP_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   '
    Dim i As Long
    With lvwEQP
        If .ListItems.Count > 0 Then
            .SortKey = ColumnHeader.Index - 1
            .Sorted = True
            For i = 1 To .ListItems.Count
                If .ListItems(i).SubItems(2) = "��" Then
                    .ListItems(i).Selected = True
                    IndexPointer = i
                    Exit For
                End If
            Next
        End If
    End With
    If ssRst.Enabled Then
        ssRst.SetFocus
    End If

End Sub

Private Sub lvwEQP_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    Dim strLvw As String
    Dim intYesNo As VbMsgBoxResult
    Dim objLvwItem As MSComctlLib.ListItem
    Dim strCurrentData As String
    Dim ii As Integer
   '
    Dim strWorkArea As String
    Dim strAccDt As String
    Dim strAccSeq As String
    Dim strSQL   As String
    Dim M2LAB302 As New ADODB.Recordset
    Dim M2LABFLAG As New ADODB.Recordset
    
    If gblnModify = True Then
        objPtInfo.FootNote = rtfComment.Text
        objPtInfo.Result.Item(ssRst.ActiveRow).TextRst = rtfText.Text
        If DataFetch <> gstrModifyData Then
             intYesNo = MsgBox("�ڷᰡ �����Ǿ����ϴ�." & vbNewLine & "������ �ڷḦ �����Ͻðڽ��ϱ�?", _
                                vbYesNo, "������")
            If intYesNo = vbYes Then Call cmdSave_Click    '����Ÿ ����
        End If
        gblnModify = False: gstrModifyData = ""
    End If
   '
    lvwPatient.ListItems.Clear
    ssRst.MaxRows = 0
    rtfText.Text = ""
    rtfComment.Text = ""
    rtfRemark.Text = ""
    blnExpect = False
    cboRelTest.Clear
    CmdTemplete False
    DoEvents
   '
    If objPtInfo Is Nothing Then
        Set objPtInfo = New clsPatientInfo
    Else
        Set objPtInfo = Nothing
        Set objPtInfo = New clsPatientInfo
    End If
    '
    If IndexPointer > 0 Then
        Set objLvwItem = lvwEQP.ListItems(IndexPointer)
        objLvwItem.SubItems(2) = " "
    End If

    IndexPointer = Item.Index
    Item.SubItems(2) = "��"
    Item.Selected = True
    Item.EnsureVisible
    
    If IndexPointer = lvwEQP.ListItems.Count Then
        cmdNext.Enabled = False
    Else
        cmdNext.Enabled = True
    End If
    If IndexPointer = 1 Then
        cmdPrevious.Enabled = False
    Else
        cmdPrevious.Enabled = True
    End If
    strLvw = LvwClickData(Item)
    PtResultLoad medGetP(strLvw, 4, vbTab)
    
    '/* ���ο� ��ũ��Ʈ�� ����� ��ȸ�ؼ� ���� ó�� �ε�� ���¿��� �����ϰ�
    '  ���������� �����Ͱ� ���ߴ��� Ȯ���ϱ� ���� gblnModify,gstrModifyData�� �̿�.
    '  gblnModify = True : ������ ��������,gstrModifyData : ������ ������
    DoEvents
    If ssRst.MaxRows > 0 Then
        gblnModify = True
        gstrModifyData = DataFetch()
        
        With lvwEQP
            strWorkArea = medGetP(.SelectedItem.ListSubItems(3).Text, 1, "-")
            strAccDt = Mid(Format(GetSystemDate, "YYYY"), 1, 2) & medGetP(.SelectedItem.ListSubItems(3).Text, 2, "-")
            strAccSeq = medGetP(.SelectedItem.ListSubItems(3).Text, 3, "-")
            
            strSQL = "         SELECT a.testcd, b.testnm, c.r_cnt, a.WORKAREA, a.ACCDT, a.ACCSEQ " & vbCRLF
            strSQL = strSQL & "FROM   S2LAB302 a, S2LAB001 b, " & vbCRLF
            strSQL = strSQL & "       ( SELECT workarea, accdt, accseq, nvl(COUNT(*),0) AS r_cnt FROM S2LAB302 where testcd = 'B004118' GROUP BY  workarea, accdt, accseq ) c " & vbCRLF
            strSQL = strSQL & "Where  b.TestCd = a.TestCd " & vbCRLF
            strSQL = strSQL & "AND    a.workarea = '" & strWorkArea & "' " & vbCRLF
            strSQL = strSQL & "AND    a.accdt = '" & strAccDt & "' " & vbCRLF
            strSQL = strSQL & "AND    a.accseq = '" & strAccSeq & "' " & vbCRLF
            strSQL = strSQL & "AND    c.workarea = a.workarea " & vbCRLF
            strSQL = strSQL & "AND    c.accdt = a.accdt " & vbCRLF
            strSQL = strSQL & "AND    c.accseq = a.accseq " & vbCRLF
            strSQL = strSQL & "AND    a.TESTCD = 'B004102' " & vbCRLF
            strSQL = strSQL & "AND    LTRIM(a.rstcd) In ('1-4��','1���̸�') "
            Debug.Print strSQL
            M2LAB302.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
            If Not M2LAB302.EOF() And Not M2LAB302.BOF() Then
               M2LAB302.MoveFirst
               If Len(Trim(M2LAB302!TestCd & "")) > 0 And M2LAB302!r_cnt > 0 Then
                  If Len(Trim(rtfComment.Text & "")) = 0 Then
                     rtfComment.Text = "inadequate RBC number :" & vbCRLF & "RBC 1~4/H.P.F���Ϸ� ���� �ʹ� ���� ���" & vbCRLF & "����� ��Ȯ���� ������ �����Ƿ� RBC ���� ������ �������� �ʽ��ϴ�." & vbCRLF
                  End If
               End If
            End If
            M2LAB302.Close: Set M2LAB302 = Nothing
            '======================================================================================
            strSQL = ""
            strSQL = " SELECT rsttxt FROM s2labflag WHERE workarea = '" & strWorkArea & "'AND accdt = '" & strAccDt & "'  AND accseq = '" & strAccSeq & "'  "
            
            M2LABFLAG.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
            
             If Not M2LABFLAG.EOF() And Not M2LABFLAG.BOF() Then
                M2LABFLAG.MoveFirst
                If Len(Trim(M2LABFLAG!RSTTXT & "")) > 0 Then
                    rtfFlagText.Text = M2LABFLAG!RSTTXT & ""
                End If
             Else
                rtfFlagText.Text = ""
             End If
             M2LABFLAG.Close: Set M2LABFLAG = Nothing
            
        End With
        ssRst.SetFocus
    End If
End Sub

Private Function DataFetch() As String
    Dim ii As Integer
    
    DataFetch = ""
    With ssRst
        .Col = objPtInfo.SSCol("RESULT"): .COL2 = objPtInfo.SSCol("EC")
        .Row = 1: .Row2 = .MaxRows
        DataFetch = .Clip & "$"
    End With
    With objPtInfo
        DataFetch = DataFetch & .FootNote & "$" & .RmkCd & "$"
        For ii = 1 To ssRst.MaxRows
            DataFetch = DataFetch & .Result.Item(ii).TextRst
        Next ii
    End With
    
End Function

Private Sub ClearData()
    gblnModify = False
    txtEqpCd.Text = ""
    lblEqpCdNm.Caption = ""
    lblErr.Caption = ""
    lblDisease.Caption = ""
    lblTelno.Caption = ""
    fraEQP.Enabled = True
    If blnFirst = True Then
       txtEqpCd.SetFocus
    End If
    chkUr.Value = 0
    IndexPointer = 0
    mskSpcNo.Text = gstrMsk
    ssRst.MaxRows = 0
    ssRst.Enabled = False
    txtEqpCd.BackColor = vbWhite
    cmdQuery.Enabled = True
    cmdSave.Enabled = False
    CmdTemplete False
   '
    lvwEQP.ListItems.Clear
    lvwPatient.ListItems.Clear
    mskSpcNo.BackColor = vbWhite
    lvwEQP.BackColor = DCM_LightGray
    lvwPatient.BackColor = DCM_LightGray
    rtfComment.BackColor = DCM_LightGray
    rtfText.BackColor = DCM_LightGray
   '
    fraComment.Enabled = False
    fraText.Enabled = False
    MsgFg = False
    LeaveCellFg = False
   '
    lblAccNoCnt.Caption = "0"
    rtfComment.Text = ""
    rtfText.Text = ""
    rtfRemark.Text = ""
    
    cmdRmk.Visible = False
    fraMesg.Visible = False
    
    cmdApply.Enabled = False
    txtBatchRst.Text = ""
    
End Sub

Private Sub EditData()
   '
    ssRst.Enabled = True
    '
    txtEqpCd.BackColor = DCM_LightGray
    'cmdQuery.Enabled = False
    cmdSave.Enabled = True
    '
    fraComment.Enabled = True
    fraText.Enabled = True
    '
'    fraEQP.Enabled = False
    mskSpcNo.BackColor = DCM_LightGray
    lvwEQP.BackColor = vbWhite
    lvwPatient.BackColor = vbWhite
    rtfComment.BackColor = &HF1F5F4     'vbWhite
    rtfText.BackColor = &HEEFFFE    'vbWhite
   '
End Sub

Private Sub DisplayCount()
    lblAccNoCnt.Caption = lvwEQP.ListItems.Count
End Sub

Private Sub PtResultLoad(ByVal strAccNo As String)
    Dim objLvwItem  As MSComctlLib.ListItem
    Dim intLvwCount As Integer
    Dim ii          As Integer
    Dim valPtInfo   As Variant
    
    lvwPatient.ListItems.Clear
    MouseRunning
    Set objPtInfo.PrgBar = prgRst
    objPtInfo.PrgBarInit
    ssRst.Visible = False
    
    If fraMesg.Visible Then fraMesg.Visible = False
    If cmdRmk.Visible Then cmdRmk.Visible = False
    
    strTmpAge = ""
    With objPtInfo
        If chkUr.Value = 1 Then
            .PtType = RESULT_BY_ACCESSION
        Else
            .PtType = RESULT_BY_EQUIPMENT
        End If
        
        .AccNo = strAccNo      '/* ������ȣ, �ݵ�� ���� �ؾ� ��./
        .LoadTable txtEqpCd.Text, ObjMyUser.EmpId
        If .TestCount > 0 Then
            CmdTemplete True
            If lvwPatient.Enabled = False Then
               lvwPatient.Enabled = True
            End If
            If .PtType = RESULT_BY_ACCESSION Then
                medDataLoadLvw lvwPatient, vbNewLine, vbTab, .GetEQPStringPtInfo
            Else
                medDataLoadLvw lvwPatient, vbNewLine, vbTab, .GetStringPtInfo
            End If
              
            valPtInfo = Split(.GetStringPtInfo, vbTab)
            
            If chkUr.Value = 1 Then
                txtDeptNm.Text = valPtInfo(4)
    '            txtDtNm.Text = valPtInfo(5)
                txtDtId.Text = objPtInfo.OrdDoct
                txtExDtId.Text = objPtInfo.MajDoct
                txtTransNo.Text = strAccNo
                strRcvDt = valPtInfo(7)
                
                rtfMessage.Text = "ȯ�ڸ� : " & valPtInfo(1) & "(" & valPtInfo(0) & ")" & vbCRLF
                strTmpAge = Trim(medGetP(valPtInfo(2), 2, "/"))
                strTmpSex = Trim(medGetP(valPtInfo(2), 1, "/"))
            Else
                txtDeptNm.Text = valPtInfo(5)
        '            txtDtNm.Text = valPtInfo(5)
                txtDtId.Text = objPtInfo.OrdDoct
                txtExDtId.Text = objPtInfo.MajDoct
                txtTransNo.Text = valPtInfo(0)
                strRcvDt = valPtInfo(8)
                
                rtfMessage.Text = "ȯ�ڸ� : " & valPtInfo(2) & "(" & valPtInfo(1) & ")" & vbCRLF
                strTmpAge = Trim(medGetP(valPtInfo(3), 2, "/"))
                strTmpSex = Trim(medGetP(valPtInfo(3), 1, "/"))
            End If

              
            Dim objDisease  As New S2LIS_ReportLib.clsDisease
            objDisease.PtId = lvwPatient.ListItems(1).SubItems(1)
            lblDisease.Caption = objDisease.Disease
            lblDisease.ToolTipText = objDisease.Disease
            Set objDisease = Nothing
            
            '========================================================================================
            '��������
            Call ICSPatientMark(lvwPatient.ListItems(1).SubItems(1), enICSNum.LIS_ALL)
            '����/����� ����ó(ȯ��ID,CONTROL)
            Call GetPtTelInfo(objPtInfo.Result.Item(1).WorkArea, objPtInfo.Result.Item(1).AccDt, objPtInfo.Result.Item(1).AccSeq, lblTelno)
            '========================================================================================
            
            rtfRemark.Text = .RmkNm
            rtfComment.Text = .FootNote
            If objPtInfo.Result.Item(1).TxtType <> "0" Then
                rtfText.Text = objPtInfo.Result.Item(1).TextRst
                rtfText.Enabled = True
                rtfText.BackColor = &HEEFFFE      'vbWhite
                cmdTextTemplete.Enabled = True
            Else
                rtfText.Enabled = False
                rtfText.BackColor = DCM_LightGray
                cmdTextTemplete.Enabled = False
            End If
            
            .GetResultSpread ssRst, RESULT_BY_EQUIPMENT
            
            '���ð˻��� ��� ...
            Dim MyResult    As New clsLISResultReview
            Dim RS          As Recordset
            Dim SSQL        As String
            Call MyResult.GetRelTest(cboRelTest, medGetP(strAccNo, 1, vbTab))
            
            '------------------------�߰� ����----------------------------
            
            strCombo = ""
            For ii = 0 To cboRelTest.ListCount - 1
                strCombo = strCombo & cboRelTest.List(ii) & COL_DIV
            Next
            If strCombo <> "" Then strCombo = Mid(strCombo, 1, Len(strCombo) - 1)
            Call frmRealTestShow.ComboDisplay(objPtInfo.Result.Item(1).TestCd, strCombo, cboRelTest, cmdSpecial, cmdMicro)
            
            'ó�渮��ũ ��ȸ(�ִ��� ���θ� ��ȸ)
            SSQL = MyResult.GetOrderRemark(objPtInfo.Result.Item(1).WorkArea, objPtInfo.Result.Item(1).AccDt, objPtInfo.Result.Item(1).AccSeq)
            Set RS = New Recordset
            RS.Open SSQL, DBConn
            If Not RS.EOF Then cmdRmk.Visible = True
            
            cboRelTest.ListIndex = 0
            Set RS = Nothing
            Set MyResult = Nothing
            '------------------------�߰� ����----------------------------
            cmdApply.Enabled = True
        Else
            MsgBox "�ش� ������ȣ�� ����� ��� Ȯ�Ή���ϴ�.", vbCritical + vbOKOnly, "������ Message"
            lblErr.Caption = "�ش� ������ȣ�� ����� ��� Ȯ�Ή���ϴ�."
            ssRst.MaxRows = 0
            lvwPatient.ListItems.Clear
            rtfText.Text = ""
            rtfComment.Text = ""
            rtfRemark.Text = ""
            With lvwEQP
                intLvwCount = .ListItems.Count
                For ii = 1 To .ListItems.Count
                    If .ListItems.Item(ii).Selected = True Then
                        .ListItems.Remove (ii)
                        Exit For
                    End If
                Next ii
                If intLvwCount = .ListItems.Count Then
                    For ii = 1 To .ListItems.Count
                        If .ListItems.Item(1).SubItems(1) = objPtInfo.AccNo Then
                            .ListItems.Remove (ii)
                            Exit For
                        End If
                    Next ii
                End If
            End With
            IndexPointer = IndexPointer - 1
            If lvwEQP.ListItems.Count = IndexPointer Then IndexPointer = IndexPointer - 1
        '
            If lvwEQP.ListItems.Count = IndexPointer Then
                Set objLvwItem = lvwEQP.ListItems(IndexPointer)
                objLvwItem.SubItems(2) = " "
                IndexPointer = 0
            End If
            If lvwEQP.ListItems.Count = 0 Then
                ClearData
            Else
                gblnModify = False
                Call cmdNext_Click
                'Set objLvwItem = lvwWS.ListItems.Item(1)
                'lvwWS_ItemClick objLvwItem
            End If
            
            cmdApply.Enabled = False
            
        End If
    End With
    
    With ssRst
        For ii = 1 To .DataRowCnt
            .Row = ii
            .Col = 5: .ForeColor = DCM_LightRed: .FontBold = True
        Next
    End With
    
    Dim i       As Integer

    With ssRst
        For i = 1 To .DataRowCnt
            .Col = 2: .Row = i
            If IsNumeric(.Text) Then
                Call objPtInfo.NumValCheck(i)
            Else
                Call ssRst_LeaveCell(2, i, 2, i, False)
            End If
        Next
        .Col = 2: .Row = 1
        If IsNumeric(.Text) Then
            Call objPtInfo.NumValCheck(1)
        Else
            Call ssRst_LeaveCell(2, 1, 2, 1, False)
        End If
    End With
    
    ssRst.Visible = True
    MouseDefault
    objPtInfo.PrgBarClear
    DoEvents
   '
End Sub

Private Sub ssRst_Click(ByVal Col As Long, ByVal Row As Long)
   '
    Dim i As Long
    Dim strTestCd As String
    Dim strSpcCd As String
    Dim strCalType As String
    Dim strTmpVal As String
    
    Dim dblTotVolume As Double
    Dim dblSerumCrea As Double
    Dim dblUrineCrea As Double
    Dim strTmp As String

    Dim dblCal1     As Double
    Dim dblCal2     As Double
    Dim dblCal3     As Double
    Dim dblCal4     As Double

    '## ����ǥ�� Clear
    If Row = 0 And Col = 4 Then
        With ssRst
            .Col = 4
            blnExpect = IIf(blnExpect, False, True)
            For i = 1 To .MaxRows
                .Row = i
                If .CellType = CellTypeCheckBox Then
                    .Value = IIf(blnExpect, 0, 1)
                End If
'                If .CellType = CellTypeCheckBox Then .Value = 0
            Next
        End With
    End If
    
    '## ����� Clear
    If Row = 0 And Col = 2 Then
        With ssRst
            For i = 1 To .MaxRows
                .Row = i
                
                objPtInfo.Result.Item(i).RstCd = ""
                objPtInfo.Result.Item(i).RstVal = "0.0000"
                objPtInfo.Result.Item(i).DPDiv = ""
                objPtInfo.Result.Item(i).HLDiv = ""
                
                .Col = objPtInfo.SSCol("HLDIV"): .Value = ""
                .Col = objPtInfo.SSCol("DPDIV"): .Value = ""
                .Col = objPtInfo.SSCol("JUDGE"): .Value = ""
                .Col = objPtInfo.SSCol("RESULT"): .Value = ""
                .Col = .MaxCols: .Value = ""
                
            Next
        End With
    End If
    
    If Row <= 0 Then Exit Sub
    SpDispRtfText
   '
    
    If Col = 1 Then
    '�κд������
        If Row = 0 Then Exit Sub
        If Not P_RealTestMicSpecial Then Exit Sub
        ssRst.Row = Row:        ssRst.Col = Col
        If objPtInfo.Result.Item(Row).RstDiv = "*" Then
            If ssRst.ForeColor = vbWhite Then
                ssRst.ForeColor = DCM_LightRed
            Else
                ssRst.ForeColor = vbWhite
            End If
        Else
            If ssRst.ForeColor = DCM_MidBlue Then
                ssRst.ForeColor = DCM_LightRed
            Else
                ssRst.ForeColor = DCM_MidBlue
            End If
        End If
        chkCul.Value = 0
        For i = 1 To ssRst.DataRowCnt
            ssRst.Row = i: ssRst.Col = 1
            If ssRst.ForeColor = DCM_LightRed Then
                chkCul.Value = 1
            End If
        Next
    
    ElseIf Col = 3 Then
        ssRst.Row = Row: ssRst.Col = 3
        If P_ApplyCalculation Then
            strTestCd = objPtInfo.Result.Item(Row).TestCd
            strSpcCd = objPtInfo.Result.Item(Row).SpcCd
            strCalType = objPtInfo.GetCalType(strTestCd, strSpcCd)
            
            If strCalType <> "" Then
                Select Case strCalType
                    Case "1", "2", "3"
                        '## 1: Creatinine, MTP, Ca, UA, BUN (24H Urine)
                        '## 2: Na, K, Cl, Amylase (24H Urine)
                        '## 3: Amylase (2H Urine)
                        '## Total Volume
                        strTmpVal = InputBox("Total Volume", "���", , 8000, 8000)
                        If Trim(strTmpVal) = "" Then
                            Exit Sub
                        Else
                            dblTotVolume = Val(strTmpVal)
                            If CheckComment = False Then
                                rtfComment.Text = rtfComment.Text & "Total Volume: " & strTmpVal & vbCRLF
                            End If
                        End If
                        
                        Call objPtInfo.CalculateResult(Row, strCalType, dblTotVolume, dblSerumCrea, dblUrineCrea)
                    Case "4"    '## CCR (24H Urine)
                        '## 1.Total Volume
                        strTmpVal = InputBox("Total Volume", "���", , 8000, 8000)
                        If Trim(strTmpVal) = "" Then
                            Exit Sub
                        Else
                            dblTotVolume = Val(strTmpVal)
                            If CheckComment = False Then
                                rtfComment.Text = rtfComment.Text & "Total Volume: " & strTmpVal & vbCRLF
                            End If
                        End If
                        
                        '## 2.Urine Creatinine
                        strTmpVal = InputBox("Urine Creatinine", "���", , 8000, 8000)
                        If Trim$(strTmpVal) = "" Then
                            Exit Sub
                        Else
                            dblUrineCrea = Val(strTmpVal)
                            rtfComment.Text = rtfComment.Text & "Urine Creatinine: " & strTmpVal & vbCRLF
                        End If
                        
                        '## 3.Serum Creatinine
                        strTmpVal = InputBox("Serum Creatinine", "���", , 8000, 8000)
                        If Trim$(strTmpVal) = "" Then
                            Exit Sub
                        Else
                            dblSerumCrea = Val(strTmpVal)
                            rtfComment.Text = rtfComment.Text & "Serum Creatinine: " & strTmpVal & vbCRLF
                        End If
                        
                        '## 4.Ű,������ Factor
                        Dim dblHuman As Double
                        
                        strTmpVal = InputBox("üǥ����", "���", , 8000, 8000)
                        If Trim$(strTmpVal) = "" Then
                            Exit Sub
                        Else
                            dblHuman = Val(strTmpVal)
                            rtfComment.Text = rtfComment.Text & "üǥ����: " & strTmpVal & vbCRLF
                        End If
                        Call objPtInfo.CalculateResult(Row, strCalType, dblTotVolume, dblSerumCrea, dblUrineCrea, dblHuman)
                    Case "5"    '## LDL-Cholesterol (Serum)
                        '## 1.Cholesterol
                        strTmpVal = InputBox("Cholesterol", "���", , 8000, 8000)
                        If Trim(strTmpVal) = "" Then
                            Exit Sub
                        Else
                            dblSerumCrea = Val(strTmpVal)
                            rtfComment.Text = rtfComment.Text & "Cholesterol: " & strTmpVal & vbCRLF
                        End If
                        
                        '## 2.HDL-Cholesterol
                        strTmpVal = InputBox("HDL-Cholesterol", "���", , 8000, 8000)
                        If Trim(strTmpVal) = "" Then
                            Exit Sub
                        Else
                            dblUrineCrea = Val(strTmpVal)
                            rtfComment.Text = rtfComment.Text & "HDL-Cholesterol: " & strTmpVal & vbCRLF
                        End If
                        
                        '## 3.TG
                        Dim dblTG As Double
                        
                        strTmpVal = InputBox("TG", "���", , 8000, 8000)
                        If Trim(strTmpVal) = "" Then
                            Exit Sub
                        Else
                            dblTG = Val(strTmpVal)
                            rtfComment.Text = rtfComment.Text & "TG: " & strTmpVal & vbCRLF
                        End If
                        Call objPtInfo.CalculateResult(Row, strCalType, dblTotVolume, dblSerumCrea, dblUrineCrea, dblTG)
                    Case "6"
                        '## 1.MPV
                        strTmpVal = InputBox("MPV", "���", , 8000, 8000)
                        If Trim(strTmpVal) = "" Then
                            Exit Sub
                        Else
                            dblTotVolume = Val(strTmpVal)
                        End If
                        
                        '## 2.PLT
                        strTmpVal = InputBox("PLT", "���", , 8000, 8000)
                        If Trim(strTmpVal) = "" Then
                            Exit Sub
                        Else
                            dblSerumCrea = Val(strTmpVal)
                        End If
                        Call objPtInfo.CalculateResult(Row, strCalType, dblTotVolume, dblSerumCrea, dblUrineCrea)
                    Case "7"    '## ACCR ������
                        '## 5.1.12: �̻��(2005-06-03)
                        '   - ACCR ������ �߰�
                        '## 1.Amylase(Serum)
                        strTmpVal = InputBox("Amylase(Serum)", "���", , 8000, 8000)
                        If Trim(strTmpVal) = "" Then
                            Exit Sub
                        Else
                            dblCal1 = Val(strTmpVal)
                        End If
                        
                        '## 2.Creatinine(Serum)
                        strTmpVal = InputBox("Creatinine(Serum)", "���", , 8000, 8000)
                        If Trim(strTmpVal) = "" Then
                            Exit Sub
                        Else
                            dblCal2 = Val(strTmpVal)
                        End If
                        
                        '## 3.Amylase(24Urine)
                        strTmpVal = InputBox("Amylase(24Urine)", "���", , 8000, 8000)
                        If Trim(strTmpVal) = "" Then
                            Exit Sub
                        Else
                            dblCal3 = Val(strTmpVal)
                        End If
                        
                        '## 4.Creatinine(24Urine)
                        strTmpVal = InputBox("Creatinine(24Urine)", "���", , 8000, 8000)
                        If Trim(strTmpVal) = "" Then
                            Exit Sub
                        Else
                            dblCal4 = Val(strTmpVal)
                        End If
                        
                        '## 5.Total Volumn
                        strTmpVal = InputBox("Total Volumn", "���", , 8000, 8000)
                        If Trim(strTmpVal) = "" Then
                            Exit Sub
                        Else
                            dblTotVolume = Val(strTmpVal)
                        End If
                        Call objPtInfo.CalculateResult(Row, strCalType, dblTotVolume, dblCal1, dblCal2, dblCal3, dblCal4)
                    Case "8"
                        '## 2007.10.09 ������ �߰� : Result = �˻����� / Ư�� Creatnine �����
                        strTmpVal = InputBox("Creatnine", "���", , 8000, 8000)
                        If Trim(strTmpVal) = "" Then
                            Exit Sub
                        Else
                            dblTotVolume = Val(strTmpVal)
                            rtfComment.Text = rtfComment.Text & "Creatinine: " & strTmpVal & vbCRLF
                        End If
                        Call objPtInfo.CalculateResult(Row, strCalType, dblTotVolume, dblSerumCrea, dblUrineCrea)
                    Case "10"
                        '## 1: Creatinine, MTP, Ca, UA, BUN (24H Urine)
                        '## 2: Na, K, Cl, Amylase (24H Urine)
                        '## 3: Amylase (2H Urine)
                        '## Total Volume
                        strTmpVal = InputBox("Total Volume", "���", , 8000, 8000)
                        If Trim(strTmpVal) = "" Then
                            Exit Sub
                        Else
                            dblTotVolume = Val(strTmpVal)
                            If CheckComment = False Then
                                rtfComment.Text = rtfComment.Text & "Total Volume: " & strTmpVal & vbCRLF
                            End If
                        End If
                        
                        Call objPtInfo.CalculateResult(Row, strCalType, dblTotVolume, dblSerumCrea, dblUrineCrea)
                End Select
            End If
            ssRst.Row = Row: ssRst.Col = 3
            ssRst.CellType = CellTypeStaticText
            ssRst.Text = "��"
            ssRst.ForeColor = DCM_Blue
        End If
    End If
End Sub

Private Sub ssRst_GotFocus()
    If MsgFg Then Exit Sub
    If LeaveCellFg Then Exit Sub

    With ssRst
        If .MaxRows = 0 Then Exit Sub
        .Row = 1
        .Col = objPtInfo.SSCol("RESULT")
        .Action = ActionActiveCell
        .EditEnterAction = EditEnterActionDown
    End With
End Sub

Private Sub ssRst_KeyUp(KeyCode As Integer, Shift As Integer)
   '
    If KeyCode = 38 Or KeyCode = 40 Then
        SpDispRtfText
    ElseIf KeyCode = vbKeyF2 Then
        Call ssRst_RightClick(1, ssRst.ActiveCol, ssRst.ActiveRow, 100, 100)
    End If
  '
End Sub

Private Sub ssRst_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
Dim strWorkArea     As String
Dim strAccDt        As String
Dim strAccSeq       As String
Dim strSQL          As String

Dim M2LAB302        As New ADODB.Recordset

    If ClickType <> 1 Then Exit Sub

    If MsgFg Then Exit Sub
    If Row <= 0 Then Exit Sub
    objPtInfo.SsTop = picRst.Top
    objPtInfo.SsLeft = picRst.Left
    ssRst.Row = Row
    ssRst.Col = Col
    ssRst.Action = ActionActiveCell
    objPtInfo.MfyFg = False
    MsgFg = True
    Call objPtInfo.PopUp(, Col)
    MsgFg = False
  '
    '2014-08-13 PSK �������� �����Ȯ���Ͽ� FootNote �ۼ��Ѵ�.
    Call objPtInfo.ResultCheck(Row)
    Select Case objPtInfo.Result.Item(Row).TestCd
     Case "B004102"
          strWorkArea = objPtInfo.Result.Item(Row).WorkArea
          strAccDt = objPtInfo.Result.Item(Row).AccDt
          strAccSeq = objPtInfo.Result.Item(Row).AccSeq
          
          strSQL = "SELECT * FROM S2LAB302 where testcd = 'B004118' AND workarea = '" & strWorkArea & "' " & vbCRLF
          strSQL = strSQL & "AND accdt = '" & strAccDt & "' AND accseq = " & strAccSeq & " "
          M2LAB302.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
          If Not M2LAB302.EOF() And Not M2LAB302.BOF() Then
             ssRst.Row = Row: ssRst.Col = Col
             Select Case Trim(ssRst.Text & "")
              Case "1-4��", "1���̸�"
                   If Len(Trim(rtfComment.Text & "")) = 0 Then
                     rtfComment.Text = "inadequate RBC number :" & vbCRLF & "RBC 1~4/H.P.F���Ϸ� ���� �ʹ� ���� ���" & vbCRLF & "����� ��Ȯ���� ������ �����Ƿ� RBC ���� ������ �������� �ʽ��ϴ�." & vbCRLF
                  End If
             End Select
          End If
          M2LAB302.Close: Set M2LAB302 = Nothing
    End Select
End Sub
'Private Sub ssRst_LostFocus()
'    Dim strTmp          As String
'    Dim strTmp1         As String
'    Dim strUTmp         As String
'    Dim strRstVal       As String
'
'    Dim strResultVal    As String
'    Dim strResultChk    As String
'    Dim strTestCd       As String
'
'    If ssRst.ActiveRow < 1 Then Exit Sub
'
'    ssRst.Row = ssRst.ActiveRow
'    ssRst.Col = objPtInfo.SSCol("RESULT")
'    strTestCd = objPtInfo.Result.Item(ssRst.ActiveRow).TestCd
'
'    strTmp = UCase(ssRst.Value)
'    strUTmp = ssRst.Value
'
'    ssRst.Col = objPtInfo.SSCol("MAXCOL"): strTmp1 = ssRst.Value
'    strRstVal = Trim(medGetP(objPtInfo.GetRstCdValString(strTestCd, strTmp1), 1, COL_DIV))
'
'    If strTmp = strRstVal Or strUTmp = strRstVal Then
'        blnRstChange = True
'        Exit Sub
'    End If
'
'    strResultVal = objPtInfo.GetRstCdValString(strTestCd, strTmp)
'    strResultChk = Trim(medGetP(strResultVal, 2, COL_DIV))
'    strResultVal = Trim(medGetP(strResultVal, 1, COL_DIV))
'
'    If strTmp <> strResultVal Then
'    '����ڵ尪�� �ִ�.
'        ssRst.Col = objPtInfo.SSCol("RESULT"): ssRst.Value = strResultVal
'        ssRst.Col = objPtInfo.SSCol("MAXCOL"): ssRst.Value = strTmp
'        If strResultChk <> "" Then
'            objPtInfo.Result.Item(ssRst.ActiveRow).DPDiv = ""
'            objPtInfo.Result.Item(ssRst.ActiveRow).HLDiv = ""
'            ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = ""
'            ssRst.Col = objPtInfo.SSCol("DPDIV"):   ssRst.Value = ""
'            ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.Value = ""
'        End If
'
'        Select Case strResultChk
'            Case "*"
'                    objPtInfo.Result.Item(ssRst.ActiveRow).HLDiv = "N"
'                    ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.Value = "N"
'                                                            ssRst.FontBold = True
'                                                            ssRst.ForeColor = DCM_LightBlue
''                                                            ssRst.ForeColor = DCM_LightRed
'                    ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "Abnormal"
'                                                            ssRst.FontBold = True
'                                                            ssRst.ForeColor = DCM_LightRed
''                    objPtInfo.Result.Item(ssRst.ActiveRow).DPDiv = "N"
''                    ssRst.Col = objPtInfo.SSCol("DPDIV"):   ssRst.Value = "N"
''                                                            ssRst.FontBold = True
''                                                            ssRst.ForeColor = DCM_LightBlue
''                    ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "N"
''                                                            ssRst.FontBold = True
''                                                            ssRst.ForeColor = DCM_LightBlue
'            Case "L"
'                    objPtInfo.Result.Item(ssRst.ActiveRow).HLDiv = strResultChk
'                    ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.Value = "��Low"
'                                                            ssRst.FontBold = True
'                                                            ssRst.ForeColor = DCM_LightBlue
'                    ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "��Low"
'                                                            ssRst.FontBold = True
'                                                            ssRst.ForeColor = DCM_LightBlue
'            Case "H"
'                    objPtInfo.Result.Item(ssRst.ActiveRow).HLDiv = strResultChk
'                    ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.Value = "High��"
'                                                            ssRst.FontBold = True
'                                                            ssRst.ForeColor = DCM_LightRed
'                    ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "High��"
'                                                            ssRst.FontBold = True
'                                                            ssRst.ForeColor = DCM_LightRed
'        End Select
'        blnRstChange = True
'    Else
'    '����ڵ尪�� ����
'        ssRst.Col = objPtInfo.SSCol("MAXCOL"): ssRst.Value = strTmp
'        ssRst.Col = objPtInfo.SSCol("DPDIV"):   ssRst.Value = ""
'        ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.Value = ""
'        ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = ""
'        objPtInfo.Result.Item(ssRst.ActiveRow).DPDiv = ""
'        objPtInfo.Result.Item(ssRst.ActiveRow).HLDiv = ""
'    End If
'End Sub

Private Sub ssRst_Advance(ByVal AdvanceNext As Boolean)
    Dim strCodeValue    As String
    Dim strRstType      As String
    Dim strErr          As String
    Dim strTestCd       As String
    Dim strResultVal    As String
    Dim strResultChk    As String
    Dim lngMaxCol       As String
    Dim lngResultCol    As String
    
    Dim Col             As Long
    Dim Row             As Long
   '
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    Row = ssRst.ActiveRow
    If Row < 0 Then Exit Sub
    lngResultCol = objPtInfo.SSCol("RESULT")
    lngMaxCol = objPtInfo.SSCol("MAXCOL")
    
    On Error GoTo ErrLevaeCell:

    Col = ssRst.ActiveCol
    If Col = lngResultCol Then
        Call objPtInfo.ResultCheck
        strRstType = objPtInfo.Result.Item(Row).RstType
        If strRstType = "N" Then
            strErr = objPtInfo.Result.Item(Row).AvalVal
            If objPtInfo.IsAvalVal = False Then
                If strErr <> "0" Then
                    strErr = "��ȿ���� �Է� ����. (" & objPtInfo.Result.Item(Row).AvalVal & "�ڸ�)"
                Else
                    strErr = "��ȿ���� �Է� ����. (�������� �Է�)"
                End If
                GoTo ErrLevaeCell
            Else
                lblErr.Caption = ""
                Call objPtInfo.NumValCheck
            End If
        ElseIf strRstType = "A" Then
            If objPtInfo.IsAlphaCd = False Then
                strErr = "ALPHA ����ڵ� �Է� ����!"
                GoTo ErrLevaeCell
            Else
               lblErr.Caption = ""
            End If
        ElseIf strRstType = "R" Then
            If objPtInfo.IsRateCd = False Then
                strErr = "������� �Է� ����!"
                GoTo ErrLevaeCell
            Else
                lblErr.Caption = ""
            End If
        ElseIf strRstType = "F" Then
            If objPtInfo.IsFreeResult = False Then
                strErr = "FREE��� �Է� ����! (10�ڸ��̳�)"
                GoTo ErrLevaeCell
            Else
                Call objPtInfo.NumValCheck
                lblErr.Caption = ""
            End If
        End If
    End If
    
    strTestCd = objPtInfo.Result.Item(Row).TestCd

    If Col = lngResultCol Then
'        ssRst.Row = Row: ssRst.Col = lngResultCol: strCodeValue = UCase(Trim(ssRst.Value))
        ssRst.Row = Row: ssRst.Col = lngMaxCol: strCodeValue = UCase(Trim(ssRst.Value))
        If strCodeValue = "" Then
            ssRst.Row = Row: ssRst.Col = lngResultCol: strCodeValue = UCase(Trim(ssRst.Value))
        End If
        If strCodeValue <> "" Then
            strResultVal = objPtInfo.GetRstCdValString(strTestCd, strCodeValue)
            strResultChk = Trim(medGetP(strResultVal, 2, COL_DIV))
            strResultVal = Trim(medGetP(strResultVal, 1, COL_DIV))
        
            If strResultVal <> ssRst.Value Then
                ssRst.Row = Row: ssRst.Col = lngResultCol:  ssRst.Value = strResultVal
                ssRst.Row = Row: ssRst.Col = lngMaxCol:     ssRst.Value = strCodeValue
'                If strResultChk <> "" Then
'                    objPtInfo.Result.Item(Row).DPDiv = ""
'                    objPtInfo.Result.Item(Row).HLDiv = ""
'                End If
                Select Case strResultChk
                    Case "*"
                            objPtInfo.Result.Item(Row).HLDiv = "N"
                            ssRst.Col = objPtInfo.SSCol("HLDiv"):   ssRst.Value = "N"
                                                                    ssRst.FontBold = True
                                                                    ssRst.ForeColor = DCM_LightBlue
                            ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "Abnormal"
                                                                    ssRst.FontBold = True
                                                                    ssRst.ForeColor = DCM_LightRed
'                            objPtInfo.Result.Item(Row).DPDiv = "N"
'                            ssRst.Col = objPtInfo.SSCol("DPDIV"):   ssRst.Value = "N"
'                                                                    ssRst.FontBold = True
'                                                                    ssRst.ForeColor = DCM_LightBlue
'                            ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "N"
'                                                                    ssRst.FontBold = True
'                                                                    ssRst.ForeColor = DCM_LightBlue
                    Case "L"
                            objPtInfo.Result.Item(Row).HLDiv = strResultChk
                            ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.Value = "��Low"
                                                                    ssRst.FontBold = True
                                                                    ssRst.ForeColor = DCM_LightBlue
                            ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "��Low"
                                                                    ssRst.FontBold = True
                                                                    ssRst.ForeColor = DCM_LightBlue
                    Case "H"
                            objPtInfo.Result.Item(Row).HLDiv = strResultChk
                            ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.Value = "High��"
                                                                    ssRst.FontBold = True
                                                                    ssRst.ForeColor = DCM_LightRed
                            ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "High��"
                                                                    ssRst.FontBold = True
                                                                    ssRst.ForeColor = DCM_LightRed
                End Select
            Else
                ssRst.Row = Row: ssRst.Col = lngMaxCol:     ssRst.Value = strCodeValue
            End If
            
        Else
            strResultVal = objPtInfo.GetRstCdValString(strTestCd, strCodeValue)
            strResultChk = Trim(medGetP(strResultVal, 2, COL_DIV))
            strResultVal = Trim(medGetP(strResultVal, 1, COL_DIV))
            
            If strResultVal <> strCodeValue Then
                ssRst.Col = lngResultCol:   ssRst.Value = strResultVal
                ssRst.Col = lngMaxCol:      ssRst.Value = strCodeValue
                Select Case strResultChk
                    Case "*"
                            objPtInfo.Result.Item(Row).HLDiv = "N"
                            ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.Value = "N"
                                                                    ssRst.FontBold = True
                                                                    ssRst.ForeColor = DCM_LightBlue
                            ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "Abnormal"
                                                                    ssRst.FontBold = True
                                                                    ssRst.ForeColor = DCM_LightRed
'                            objPtInfo.Result.Item(Row).DPDiv = "N"
'                            ssRst.Col = objPtInfo.SSCol("DPDIV"):   ssRst.Value = "N"
'                                                                    ssRst.FontBold = True
'                                                                    ssRst.ForeColor = DCM_LightBlue
'                            ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "N"
'                                                                    ssRst.FontBold = True
'                                                                    ssRst.ForeColor = DCM_LightBlue
                    Case "L"
                            objPtInfo.Result.Item(Row).HLDiv = strResultChk
                            ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.Value = "��Low"
                                                                    ssRst.FontBold = True
                                                                    ssRst.ForeColor = DCM_LightBlue
                            ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "��Low"
                                                                    ssRst.FontBold = True
                                                                    ssRst.ForeColor = DCM_LightBlue
                    Case "H"
                            objPtInfo.Result.Item(Row).HLDiv = strResultChk
                            ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.Value = "High��"
                                                                    ssRst.FontBold = True
                                                                    ssRst.ForeColor = DCM_LightRed
                            ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "High��"
                                                                    ssRst.FontBold = True
                                                                    ssRst.ForeColor = DCM_LightRed
                End Select
            Else
                If strRstType = "F" Then
                    ssRst.Col = lngResultCol:   ssRst.Value = strCodeValue
                    ssRst.Col = lngMaxCol:      ssRst.Value = strCodeValue
                ElseIf strRstType = "N" Then
                    If IsNumeric(strCodeValue) Then
                        ssRst.Col = lngResultCol:   ssRst.Value = strCodeValue
                        ssRst.Col = lngMaxCol:      ssRst.Value = strCodeValue
                    Else
                        ssRst.Col = lngResultCol:   ssRst.Value = ""
                        ssRst.Col = lngMaxCol:      ssRst.Value = ""
                    End If
                Else
                    ssRst.Col = lngResultCol:   ssRst.Value = strCodeValue
                    ssRst.Col = lngMaxCol:      ssRst.Value = strCodeValue
                End If
            End If
        End If
    End If
    
    ssRst.Row = Row
    ssRst.Col = 2
    If Trim(ssRst.Value) = "" Then
        ssRst.Col = 14: ssRst.Value = ""
    End If
    
    LeaveCellFg = False
    Exit Sub
   '
ErrLevaeCell:
        
    With ssRst
        .Row = Row: .Col = objPtInfo.SSCol("RESULT"): .Value = ""
    End With
    objPtInfo.ResultCheck
    
    MsgFg = True
    MsgBox strErr, vbCritical + vbOKOnly, "����Է� Ȯ��"
    MsgFg = False
    LeaveCellFg = True
    
    On Error Resume Next
    ssRst.SetFocus
End Sub

Private Sub ssRst_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    
    Dim strCodeValue    As String       '�Է°�
    Dim strRstType      As String       '���Ÿ��
    Dim strErr          As String       '�����޼���
    Dim strTestCd       As String       '������ �˻��ڵ�
    Dim strResultVal    As String       '�����
    Dim strResultChk    As String       '����ڵ��Է°� üũ
    Dim lngResultCol    As Long         '����Է� Col
    Dim lngMaxCol       As Long         '������� Col
    
    strResultVal = "": strResultChk = ""
    lngMaxCol = objPtInfo.SSCol("MAXCOL")
    lngResultCol = objPtInfo.SSCol("RESULT")
    
    If Row < 1 Then Exit Sub
    If MsgFg Then Exit Sub
    
    On Error GoTo ErrLevaeCell
    
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    If NewRow > 0 Then Call frmRealTestShow.ComboDisplay(objPtInfo.Result.Item(NewRow).TestCd, strCombo, cboRelTest, cmdSpecial, cmdMicro)
    If Row = ssRst.MaxRows Then
        'Advance �̺�Ʈ���� ��Ŀ���� �������忡�� �ٸ���Ʈ�ѷ� �Ѿ��
        'LeaveCell�̺�Ʈ�� �����̸� �����ϱ� ���ؼ� exit sub�� ��
        '�㳪, ESR�� �ƴ� �ٸ� �����ۿ� ���ؼ��� �׸��� �ϳ��϶� EXIT SUb�� ����
        '����ġ üũ�� �ȵȴ�.
        blnRstChange = False
        If lngResultCol <> Col Then blnRstChange = True
        If blnRstChange = True Then Exit Sub
'        If lngResultCol = Col Then Call ssRst_LostFocus
'
'        If UCase(Me.ActiveControl.Name) = "SSRST" Then Exit Sub
        If blnRstChange = True Then Exit Sub
'    Else
'        If NewRow > 0 Then Call frmRealTestShow.ComboDisplay(objPtInfo.Result.Item(NewRow).TestCd, strCombo, cboRelTest, cmdSpecial, cmdMicro)
'        If lngResultCol <> Col Then blnRstChange = True
'        If lngResultCol = Col Then Call ssRst_LostFocus
'        If blnRstChange = True Then Exit Sub
    End If

    lblErr.Caption = ""
    If Col = objPtInfo.SSCol("RESULT") Then
        Call objPtInfo.ResultCheck(Row)
        strRstType = objPtInfo.Result.Item(Row).RstType
        If strRstType = "N" Then
            strErr = objPtInfo.Result.Item(Row).AvalVal
            If objPtInfo.IsAvalVal = False Then
                If strErr <> "0" Then
                    strErr = "��ȿ���� �Է� ����. (" & objPtInfo.Result.Item(Row).AvalVal & "�ڸ�)"
                Else
                    strErr = "��ȿ���� �Է� ����. (�������� �Է�)"
                End If
                GoTo ErrLevaeCell
            Else
            Call objPtInfo.NumValCheck
         End If
      ElseIf strRstType = "A" Then
         If objPtInfo.IsAlphaCd(Row) = False Then
            strErr = "ALPHA ����ڵ� �Է� ����!"
            GoTo ErrLevaeCell
         End If
      ElseIf strRstType = "R" Then
         If objPtInfo.IsRateCd = False Then
            strErr = "������� �Է� ����!"
            GoTo ErrLevaeCell
         End If
      ElseIf strRstType = "F" Then
         If objPtInfo.IsFreeResult = False Then
            strErr = "FREE��� �Է� ����! (10�ڸ��̳�)"
            GoTo ErrLevaeCell
         End If
         Call objPtInfo.NumValCheck
      End If
      ssRst.EditEnterAction = EditEnterActionDown
   End If
   '
   Call SpDispRtfText(NewRow)
    
    strTestCd = objPtInfo.Result.Item(Row).TestCd
    If Col = lngResultCol Then
'        ssRst.Row = Row: ssRst.Col = lngResultCol: strCodeValue = UCase(Trim(ssRst.Value))
        ssRst.Row = Row: ssRst.Col = lngMaxCol: strCodeValue = UCase(Trim(ssRst.Value))
        If strCodeValue = "" Then
            ssRst.Row = Row: ssRst.Col = lngResultCol: strCodeValue = UCase(Trim(ssRst.Value))
        End If
        If strCodeValue <> "" Then
            '���� Col�� ���� �������(popup�̿�)
'            ssRst.Col = lngMaxCol:          ssRst.Value = strCodeValue
            strResultVal = objPtInfo.GetRstCdValString(strTestCd, strCodeValue)       '�����
            strResultChk = Trim(medGetP(strResultVal, 2, COL_DIV))          '���üũ��
            strResultVal = Trim(medGetP(strResultVal, 1, COL_DIV))          '�����
            
            ssRst.Col = lngResultCol:   ssRst.Value = strResultVal
            ssRst.Col = lngMaxCol:      ssRst.Value = strCodeValue
'            If strResultChk <> "" Then
'                objPtInfo.Result.Item(Row).DPDiv = ""
'                objPtInfo.Result.Item(Row).HLDiv = ""
'            End If
            Select Case strResultChk
                Case "*"
                        objPtInfo.Result.Item(Row).HLDiv = "N"
                        ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.Value = "N"
                                                                ssRst.FontBold = True
                                                                ssRst.ForeColor = DCM_LightBlue
'                                                                ssRst.ForeColor = DCM_LightRed
                                                                
                        ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "Abnormal"
                                                                ssRst.FontBold = True
                                                                ssRst.ForeColor = DCM_LightRed
                Case "L"
                        objPtInfo.Result.Item(Row).HLDiv = strResultChk
                        ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.Value = "��Low"
                                                                ssRst.FontBold = True
                                                                ssRst.ForeColor = DCM_LightBlue
                        ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "��Low"
                                                                ssRst.FontBold = True
                                                                ssRst.ForeColor = DCM_LightBlue
                Case "H"
                        objPtInfo.Result.Item(Row).HLDiv = strResultChk
                        ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.Value = "High��"
                                                                ssRst.FontBold = True
                                                                ssRst.ForeColor = DCM_LightRed
                        ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "High��"
                                                                ssRst.FontBold = True
                                                                ssRst.ForeColor = DCM_LightRed
            End Select
            
'            If strResultVal <> ssRst.Value Then
'                ssRst.Col = lngResultCol:   ssRst.Value = strResultVal
'                ssRst.Col = lngMaxCol:      ssRst.Value = strCodeValue
'                If strResultChk <> "" Then
'                    objPtInfo.Result.Item(Row).DPDiv = ""
'                    objPtInfo.Result.Item(Row).HLDiv = ""
'                End If
'                Select Case strResultChk
'                    Case "*"
'                            objPtInfo.Result.Item(Row).DPDiv = "N"
'                            ssRst.Col = objPtInfo.SSCol("DPDIV"):   ssRst.Value = "N"
'                                                                    ssRst.FontBold = True
'                                                                    ssRst.ForeColor = DCM_LightBlue
'                            ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "N"
'                                                                    ssRst.FontBold = True
'                                                                    ssRst.ForeColor = DCM_LightBlue
'                    Case "L"
'                            objPtInfo.Result.Item(Row).HLDiv = strResultChk
'                            ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.Value = "��Low"
'                                                                    ssRst.FontBold = True
'                                                                    ssRst.ForeColor = DCM_LightBlue
'                            ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "��Low"
'                                                                    ssRst.FontBold = True
'                                                                    ssRst.ForeColor = DCM_LightBlue
'                    Case "H"
'                            objPtInfo.Result.Item(Row).HLDiv = strResultChk
'                            ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.Value = "High��"
'                                                                    ssRst.FontBold = True
'                                                                    ssRst.ForeColor = DCM_LightRed
'                            ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "High��"
'                                                                    ssRst.FontBold = True
'                                                                    ssRst.ForeColor = DCM_LightRed
'                End Select
'            Else
'                ssRst.Row = Row: ssRst.Col = lngMaxCol:     ssRst.Value = strCodeValue
'            End If
        Else
            '����Col�� ���� �������(�����Է�)
            ssRst.Col = lngResultCol: strCodeValue = UCase(Trim(ssRst.Value))
            strResultVal = objPtInfo.GetRstCdValString(strTestCd, strCodeValue)       '�����
            strResultChk = Trim(medGetP(strResultVal, 2, COL_DIV))          '���üũ��
            strResultVal = Trim(medGetP(strResultVal, 1, COL_DIV))          '�����
            If strResultVal <> strCodeValue Then
                ssRst.Col = lngResultCol:   ssRst.Value = strResultVal
                ssRst.Col = lngMaxCol:      ssRst.Value = strCodeValue
'                If strResultChk <> "" Then
'                    objPtInfo.Result.Item(Row).DPDiv = ""
'                    objPtInfo.Result.Item(Row).HLDiv = ""
'                End If
                Select Case strResultChk
                    Case "*"
                            objPtInfo.Result.Item(Row).HLDiv = "N"
                            ssRst.Col = objPtInfo.SSCol("HLDiv"):   ssRst.Value = "N"
                                                                    ssRst.FontBold = True
                                                                    ssRst.ForeColor = DCM_LightBlue
'                            ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = strResultChk
                            ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "Abnormal"
                                                                    ssRst.FontBold = True
                                                                    ssRst.ForeColor = DCM_LightRed
'                            objPtInfo.Result.Item(Row).DPDiv = strResultChk
'                            ssRst.Col = objPtInfo.SSCol("DPDIV"):   ssRst.Value = strResultChk
'                                                                    ssRst.FontBold = True
'                                                                    ssRst.ForeColor = DCM_LightBlue
'                            ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = strResultChk
'                                                                    ssRst.FontBold = True
'                                                                    ssRst.ForeColor = DCM_LightBlue
                    Case "L"
                            objPtInfo.Result.Item(Row).HLDiv = strResultChk
                            ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.Value = "��Low"
                                                                    ssRst.FontBold = True
                                                                    ssRst.ForeColor = DCM_LightBlue
                            ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "��Low"
                                                                    ssRst.FontBold = True
                                                                    ssRst.ForeColor = DCM_LightBlue
                    Case "H"
                            objPtInfo.Result.Item(Row).HLDiv = strResultChk
                            ssRst.Col = objPtInfo.SSCol("HLDIV"):   ssRst.Value = "High��"
                                                                    ssRst.FontBold = True
                                                                    ssRst.ForeColor = DCM_LightRed
                            ssRst.Col = objPtInfo.SSCol("JUDGE"):   ssRst.Value = "High��"
                                                                    ssRst.FontBold = True
                                                                    ssRst.ForeColor = DCM_LightRed
                End Select
            Else
                If strRstType = "F" Then
                    ssRst.Col = lngResultCol:   ssRst.Value = strCodeValue
                    ssRst.Col = lngMaxCol:      ssRst.Value = strCodeValue
                ElseIf strRstType = "N" Then
                    If IsNumeric(strCodeValue) Then
                        ssRst.Col = lngResultCol:   ssRst.Value = strCodeValue
                        ssRst.Col = lngMaxCol:      ssRst.Value = strCodeValue
                    Else
                        ssRst.Col = lngResultCol:   ssRst.Value = ""
                        ssRst.Col = lngMaxCol:      ssRst.Value = ""
                    End If
                Else
                    ssRst.Col = lngResultCol:   ssRst.Value = strCodeValue
                    ssRst.Col = lngMaxCol:      ssRst.Value = strCodeValue
                End If
            End If
        End If
    End If
    
    LeaveCellFg = False
    Exit Sub
   '
ErrLevaeCell:
   If ssRst.Visible = True Then
       With ssRst
          .Row = Row: .Col = objPtInfo.SSCol("RESULT"): .Value = ""
          .Action = ActionActiveCell
       End With
       objPtInfo.ResultCheck
        
       MsgFg = True
       MsgBox strErr, vbCritical, "����Է� Ȯ��"
       MsgFg = False
       
       LeaveCellFg = True
    
       Cancel = True
       On Error Resume Next
       ssRst.SetFocus
    End If
End Sub

Private Sub ssRst_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
   '
   If Row < 1 Then Exit Sub
   objPtInfo.SpToolTip Row, Col, MultiLine, ShowTip, TipText, TipWidth
   ssRst.TextTip = TextTipFloatingFocusOnly
   '
End Sub

Private Sub SpDispRtfText(Optional Row As Long = 0)
   '
   If Row < 0 Then Exit Sub
   If Row = 0 Then
      ssRst.Row = ssRst.ActiveRow
   Else
      ssRst.Row = Row
   End If
   ssRst.Col = objPtInfo.SSCol("TXT")
   With objPtInfo.Result.Item(ssRst.Row)
      If ssRst.CellType = CellTypePicture Or ssRst.Text = "T" Then
         If .TxtType <> "0" Then
            rtfText.Text = .TextRst
            rtfText.Enabled = True
            cmdTextTemplete.Enabled = True
            rtfText.BackColor = &HEEFFFE    'vbWhite
         Else
            rtfText.Text = ""
            rtfText.Enabled = False
            cmdTextTemplete.Enabled = False
            rtfText.BackColor = DCM_LightGray
         End If
      Else
         rtfText.Text = ""
         rtfText.Enabled = False
         cmdTextTemplete.Enabled = False
         rtfText.BackColor = DCM_LightGray
      End If
   End With
   '
End Sub

Private Sub CmdTemplete(ByVal blnVisible As Boolean)
   '
   cmdTextTemplete.Enabled = blnVisible
   cmdRemarkTemplete.Enabled = blnVisible
   cmdCommentTemplete.Enabled = blnVisible
   '
End Sub

Private Sub TrasferListPop(ByVal EqpCd As String)
   If EqpCd = "" Then Exit Sub
   Set objCodeList = New clsPopUpList
   With objCodeList
'      Set .MyOraSE = OraSE
'      Set .MyDb = dbconn
        .Connection = DBConn
        .FormCaption = "Instrument List"
        .Tag = "Transfer"
        .FormHeight = 2895
        .FormWidth = 4995
        .ColumnHeaderWidth = "1214.929;1110.047;1110.047;1110.047"
        .ColumnHeaderText = "��ü��ġ;��ü��ȣ;��������;���۽ð�"
'        .SqlStmt = "SELECT a.transno,a.spcyy||'-'|| to_char(a.spcno)  AS SpcNo, a.transdt, a.transtm ,a.eqpcd  FROM s2lab306 a  WHERE  a.transdt >= '20030824'  AND   exists(SELECT c.accseq FROM s2lab201 c, s2lab302 d             WHERE c.spcyy = a.spcyy             AND c.spcno = a.spcno               AND   d.workarea = c.workarea           AND d.accdt = c.accdt           AND d.accseq = c.accseq             AND  (d.vfydt = ''  or  d.vfydt is null) AND d.eqpcd=a.eqpcd) ORDER BY transdt, transtm, transno "
        .HideSearchTool = True
        .SortColumn = 3
        
        .LoadPopUp objPtInfo.GetSqlTransferPop(EqpCd)
        
'      .ListCaption = "Instrument List"
'      .ListColHeader = "Name" & vbTab & "Code"
'      .Top = Me.cmdTrans.Top + 2000
'      .Left = Me.cmdTrans.Left - 1100
'      .Width = 3450
'      .Height = 3000
'      .Tag = "Transfer"
'      .CaptionOn = False
'      .MultiSel = False
'      .PopupList objPtInfo.GetSqlTransferPop(EqpCd), 2
   End With
End Sub

Private Sub cmdCul_Click()
    Dim objTestCd   As New clsDictionary
    Dim sPtid       As String
    Dim ii          As Integer
    
    Me.MousePointer = vbHourglass
    
    Set objCuM = New frmTmpCumulative

    objTestCd.Clear
    objTestCd.FieldInialize "testcd", "spccd"

    objTestCd.Sort = False
    
    For ii = 1 To ssRst.MaxRows
        ssRst.Row = ii
        ssRst.Col = 1
        With objPtInfo.Result.Item(ii)
            If chkCul.Value = 0 Then
                If objTestCd.Exists("testcd") = False Then
                    objTestCd.AddNew .TestCd, .SpcCd
                End If
            Else
                If ssRst.ForeColor = DCM_LightRed Then
                    If objTestCd.Exists("testcd") = False Then
                        objTestCd.AddNew .TestCd, .SpcCd
                    End If
                End If
            End If
            sPtid = objPtInfo.PtId
        End With
    Next ii
    objTestCd.Sort = True

    With objCuM
        .Top = Me.Top + 2000
        .Left = Me.Left + 200
        .MousePointer = vbDefault
        .Caption = "ȯ��ID: " & sPtid & " �������"
        Call .DisplayItem(objTestCd, sPtid)
        DoEvents
        
        .WindowState = 0
        .Show vbModal
        DoEvents
    End With

    Set objTestCd = Nothing
    
    Me.MousePointer = vbDefault
End Sub
Private Sub cmdRmk_Click()
    Dim objSQL   As clsLISResultReview
    Dim RS       As Recordset
    Dim aryTmp() As String
    Dim strTmp   As String
    Dim SSQL     As String
    Dim ii       As Integer
    
    txtMesg.Text = ""
    Set objSQL = New clsLISResultReview
    SSQL = objSQL.GetOrderRemark(objPtInfo.Result.Item(1).WorkArea, objPtInfo.Result.Item(1).AccDt, objPtInfo.Result.Item(1).AccSeq)
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    If Not RS.EOF Then
        strTmp = "ó������ : " & Format(RS.Fields("orddt").Value & "", "####-##-##") & vbCRLF
        strTmp = strTmp & "ó���ȣ : " & RS.Fields("ordno").Value & "" & vbCRLF
        strTmp = strTmp & "ó����  " & vbCRLF
        txtMesg.Text = strTmp
        strTmp = ""
        aryTmp = Split(RS.Fields("mesg").Value & "", vbCRLF)
        For ii = LBound(aryTmp) To UBound(aryTmp)
            strTmp = strTmp & " " & aryTmp(ii) & vbCRLF
        Next
        txtMesg.Text = txtMesg.Text & strTmp
        fraMesg.Visible = True
        fraMesg.ZOrder 0
    End If
    
    Set RS = Nothing
    Set objSQL = Nothing
End Sub

Private Sub cmdOk_Click()
    fraMesg.Visible = False
    ssRst.SetFocus
End Sub

'-----------------------------------------------------------------------------'
'   ��� : Comment���� "Total Volume:" ���ڿ� ��ȸ
'   ��ȯ : ����(True), ������(False)
'-----------------------------------------------------------------------------'
Private Function CheckComment() As Boolean
    Dim strTemp As String
    
    strTemp = rtfComment.Text
    If InStr(strTemp, "Total Volume:") > 0 Then
        CheckComment = True
    End If
End Function
