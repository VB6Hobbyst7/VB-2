VERSION 5.00
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frmEntryNo 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   0  '����
   Caption         =   "ABO��� �������"
   ClientHeight    =   8640
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13725
   Icon            =   "frmEntryNo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8640
   ScaleWidth      =   13725
   ShowInTaskbar   =   0   'False
   Begin MedControls1.LisLabel LisLabel3 
      Height          =   315
      Left            =   1897
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   2655
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   -2147483634
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
      Caption         =   "  �˻� ���"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel13 
      Height          =   315
      Left            =   1897
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   6600
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   -2147483634
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
      Caption         =   "Remark"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel12 
      Height          =   315
      Left            =   1897
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   4470
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   -2147483634
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
      Caption         =   "Comment by Accession No"
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
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
      Height          =   705
      Left            =   1890
      TabIndex        =   51
      Top             =   6825
      Width           =   9945
      Begin VB.ComboBox cboRemark 
         Height          =   300
         Left            =   150
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   6
         Top             =   240
         Width           =   9690
      End
   End
   Begin VB.Frame fraComment 
      BackColor       =   &H00DBE6E6&
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1845
      Left            =   1890
      TabIndex        =   48
      Tag             =   "20003"
      Top             =   4695
      Width           =   9945
      Begin VB.CommandButton cmdCommentTemplete 
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
         Left            =   9495
         Picture         =   "frmEntryNo.frx":076A
         Style           =   1  '�׷���
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   1380
         Width           =   315
      End
      Begin VB.TextBox txtComment 
         Height          =   1485
         Left            =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  '����
         TabIndex        =   5
         Top             =   225
         Width           =   9270
      End
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   315
      Left            =   1897
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1155
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   -2147483634
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
      Caption         =   "  ȯ�� ����"
      Appearance      =   0
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DBE6E6&
      Height          =   1215
      Left            =   1890
      TabIndex        =   17
      Top             =   1380
      Width           =   9945
      Begin MedControls1.LisLabel lblPtId 
         Height          =   315
         Left            =   1290
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   300
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         BackColor       =   14411494
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
      End
      Begin MedControls1.LisLabel lblPtNm 
         Height          =   330
         Left            =   4545
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   315
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         BackColor       =   14411494
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
      End
      Begin MedControls1.LisLabel lblSexAge 
         Height          =   330
         Left            =   7920
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   300
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         BackColor       =   14411494
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
      End
      Begin MedControls1.LisLabel lblDept 
         Height          =   330
         Left            =   1290
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   660
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         BackColor       =   14411494
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
      End
      Begin MedControls1.LisLabel lblWard 
         Height          =   330
         Left            =   4545
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   660
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         BackColor       =   14411494
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
      End
      Begin MedControls1.LisLabel lblDoct 
         Height          =   330
         Left            =   7920
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   660
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         BackColor       =   14411494
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
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   3
         Left            =   6900
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   300
         Width           =   975
         _ExtentX        =   1720
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
         Caption         =   "��/����"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   4
         Left            =   6900
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   660
         Width           =   975
         _ExtentX        =   1720
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
         Caption         =   "��ġ��"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   315
         Index           =   0
         Left            =   270
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   300
         Width           =   975
         _ExtentX        =   1720
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
         Caption         =   "ȯ��ID"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   1
         Left            =   3525
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   300
         Width           =   975
         _ExtentX        =   1720
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
         Caption         =   "ȯ�ڸ�"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   2
         Left            =   3525
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   660
         Width           =   975
         _ExtentX        =   1720
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
         Caption         =   "����"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lbldt 
         Height          =   330
         Index           =   5
         Left            =   270
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   660
         Width           =   975
         _ExtentX        =   1720
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
         Caption         =   "�����"
         Appearance      =   0
      End
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "����(&S)"
      Height          =   510
      Left            =   7830
      Style           =   1  '�׷���
      TabIndex        =   7
      Tag             =   "15101"
      Top             =   7785
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "ȭ������(&C)"
      CausesValidation=   0   'False
      Height          =   510
      Left            =   9180
      Style           =   1  '�׷���
      TabIndex        =   8
      Tag             =   "124"
      Top             =   7785
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "����(&X)"
      CausesValidation=   0   'False
      Height          =   510
      Left            =   10500
      Style           =   1  '�׷���
      TabIndex        =   9
      Tag             =   "128"
      Top             =   7785
      Width           =   1320
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DBE6E6&
      Height          =   1530
      Left            =   1890
      TabIndex        =   31
      Top             =   2880
      Width           =   9945
      Begin VB.CommandButton cmdPop 
         BackColor       =   &H00F4F0F2&
         Height          =   315
         Index           =   3
         Left            =   7815
         Picture         =   "frmEntryNo.frx":0C9C
         Style           =   1  '�׷���
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   1110
         Width           =   330
      End
      Begin VB.TextBox txtRhSub 
         Alignment       =   2  '��� ����
         Height          =   315
         Left            =   6705
         TabIndex        =   4
         Text            =   "123456789012"
         Top             =   1110
         Width           =   1095
      End
      Begin VB.CommandButton cmdPop 
         BackColor       =   &H00F4F0F2&
         Height          =   315
         Index           =   1
         Left            =   7815
         Picture         =   "frmEntryNo.frx":0D4E
         Style           =   1  '�׷���
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   675
         Width           =   330
      End
      Begin VB.TextBox txtRh 
         Alignment       =   2  '��� ����
         Height          =   315
         Left            =   6705
         TabIndex        =   2
         Text            =   "123456789012"
         Top             =   675
         Width           =   1095
      End
      Begin VB.CommandButton cmdPop 
         BackColor       =   &H00F4F0F2&
         Height          =   315
         Index           =   2
         Left            =   2910
         Picture         =   "frmEntryNo.frx":0E00
         Style           =   1  '�׷���
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   1110
         Width           =   330
      End
      Begin VB.TextBox txtABOSub 
         Alignment       =   2  '��� ����
         Height          =   315
         Left            =   1800
         TabIndex        =   3
         Text            =   "123456789012"
         Top             =   1110
         Width           =   1095
      End
      Begin VB.CommandButton cmdPop 
         BackColor       =   &H00F4F0F2&
         Height          =   315
         Index           =   0
         Left            =   2910
         Picture         =   "frmEntryNo.frx":0EB2
         Style           =   1  '�׷���
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   675
         Width           =   330
      End
      Begin VB.CheckBox chkVerify 
         BackColor       =   &H00C0E0FF&
         Caption         =   "������ ����"
         Height          =   225
         Left            =   8280
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   270
         Value           =   1  'Ȯ��
         Width           =   1455
      End
      Begin VB.TextBox txtABO 
         Alignment       =   2  '��� ����
         Height          =   315
         Left            =   1800
         TabIndex        =   1
         Text            =   "123456789012"
         Top             =   675
         Width           =   1095
      End
      Begin VB.OptionButton optType 
         BackColor       =   &H00FCEFE9&
         Caption         =   "Front Typing"
         Height          =   420
         Index           =   0
         Left            =   180
         Style           =   1  '�׷���
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   180
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton optType 
         BackColor       =   &H00FCEFE9&
         Caption         =   "Back Typing"
         Height          =   420
         Index           =   1
         Left            =   1875
         Style           =   1  '�׷���
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   165
         Width           =   1695
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Left            =   180
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   675
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
         BackColor       =   10392451
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
         Alignment       =   1
         Caption         =   "ABO ���"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel5 
         Height          =   315
         Left            =   180
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   1110
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
         BackColor       =   10392451
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
         Alignment       =   1
         Caption         =   "ABO SubType"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel6 
         Height          =   315
         Left            =   5100
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   1110
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
         BackColor       =   10392451
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
         Alignment       =   1
         Caption         =   "Rh SubType"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel7 
         Height          =   315
         Left            =   5115
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   675
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
         BackColor       =   10392451
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
         Alignment       =   1
         Caption         =   "Rh ���"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblABO 
         Height          =   315
         Left            =   3255
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   675
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   556
         BackColor       =   14411494
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
      End
      Begin MedControls1.LisLabel lblABOSub 
         Height          =   315
         Left            =   3255
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   1110
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   556
         BackColor       =   14411494
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
      End
      Begin MedControls1.LisLabel lblRh 
         Height          =   315
         Left            =   8160
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   675
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   556
         BackColor       =   14411494
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
      End
      Begin MedControls1.LisLabel lblRhSub 
         Height          =   315
         Left            =   8160
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   1110
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   556
         BackColor       =   14411494
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
      End
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   315
      Left            =   1897
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   225
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   -2147483634
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
      Caption         =   "  ���� ��ȣ(��ü ��ȣ)"
      Appearance      =   0
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00DBE6E6&
      Height          =   645
      Left            =   1890
      TabIndex        =   11
      Top             =   450
      Width           =   9945
      Begin VB.TextBox txtAccNo 
         Alignment       =   2  '��� ����
         Height          =   375
         Left            =   1515
         TabIndex        =   0
         Text            =   "123456789012"
         Top             =   180
         Width           =   1590
      End
      Begin VB.CheckBox chkBarcode 
         BackColor       =   &H00FFC0C0&
         Caption         =   "���ڵ帮���� �б�"
         Height          =   225
         Left            =   8040
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   255
         Width           =   1800
      End
      Begin MedControls1.LisLabel lblBarNo 
         Height          =   360
         Left            =   4485
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   180
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   635
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
         Caption         =   "01-031028-1234"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblAccNo 
         Height          =   360
         Left            =   315
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   180
         Width           =   1185
         _ExtentX        =   2090
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
         Caption         =   "���� ��ȣ"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel lblBarcode 
         Height          =   360
         Left            =   3285
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   180
         Width           =   1185
         _ExtentX        =   2090
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
         Caption         =   "���ڵ� ��ȣ"
         Appearance      =   0
      End
   End
End
Attribute VB_Name = "frmEntryNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'������ �ۼ��Ǿ��� �ҽ��� ���� ����� ���Ӱ� �ۼ���. �ֳĸ� ����� �۵��� �ȵǴ� ����ŵ�..
'Modify By legends 2004/12/09 ���� ������������

'������ȣ�� �Է��ϸ� bbs303���� ������ �ε�
'�̹� �����Ͱ� �ִ� ��쿡�� vfyid�� �α� �Ƶ� ������ ���θ� �Ǵ��Ͽ� �ٸ� ��쿡�� ���� ���� �Ѵ�.
'

Public Event FormClose()

Private Enum KindOfTyping
    NoUse = -1
    AnyUse = 0
    FrontUse = 1
    BackUse = 2
End Enum

Private WithEvents objPopList As clsPopUpList
Attribute objPopList.VB_VarHelpID = -1
Private WithEvents objComment As frmTempSearch
Attribute objComment.VB_VarHelpID = -1

Private Sub chkBarcode_Click()
    If chkBarcode.Value = 1 Then
        lblAccNo.Caption = "���ڵ� ��ȣ"
        lblBarcode.Caption = "���� ��ȣ"
    Else
        lblAccNo.Caption = "���� ��ȣ" '
        lblBarcode.Caption = "���ڵ� ��ȣ"
    End If
    
    txtAccNo.Text = ""
    lblBarNo.Caption = ""
    
    Call ClearPtInfo
    Call ClearOthers
    
    On Error Resume Next
    txtAccNo.SetFocus
End Sub

Private Sub cmdClear_Click()
    txtAccNo.Text = ""
    lblBarNo.Caption = ""
    Call ClearPtInfo
    Call ClearOthers
    On Error Resume Next
    txtAccNo.SetFocus
End Sub

Private Sub cmdCommentTemplete_Click()
    Set objComment = New frmTempSearch
    
    With objComment
        .CurrentComment = txtComment.Text
        .Show vbModal
    End With
            
    Set objComment = Nothing
End Sub

Private Sub cmdExit_Click()
    Unload Me
    RaiseEvent FormClose
End Sub

Private Sub cmdPop_Click(Index As Integer)
    Dim Rs As Recordset
    
    If txtAccNo.Text = "" Then Exit Sub
    
    Set objPopList = New clsPopUpList
    Set Rs = New Recordset

    Set Rs = GetABOResult(Index)

    If Rs.EOF = False Then
        With objPopList
            .Tag = Index

            .Recordset = Rs
            .HideSearchTool = True
            .SelectByClick = True
            .FormWidth = 4635
            .FormHeight = 2880
            .FormCaption = "����ڵ� ã��"
            .ColumnHeaderText = "����ڵ�;����ڵ��"
            .ColumnHeaderWidth = "1110.047;3075.024"
            .HideColumnHeaders = True
            .LoadPopUp
        End With
    End If
    Set objPopList = Nothing
End Sub

Private Sub cmdSave_Click()
    Dim objABOSql As clsABOSql
    Dim strRemark As String
    Dim strComment As String
'    Dim strVfyDt As String
'    Dim strVfyTm As String
    Dim blnSave As Boolean
    Dim strWorkarea As String
    Dim strAccdt As String
    Dim strAccseq As String
    
    If CheckValidation = False Then Exit Sub

    If cboRemark.ListIndex = 0 Or cboRemark.ListIndex = -1 Then
        strRemark = ""
    Else
        strRemark = medGetP(cboRemark.Text, 1, vbTab)
    End If
    strComment = txtComment.Text
    
    If chkBarcode.Value = 0 Then
        strWorkarea = medGetP(Trim(txtAccNo.Text), 1, "-")
        strAccdt = IIf(medGetP(Trim(txtAccNo.Text), 2, "-") Like "99*", _
                      "19" & medGetP(Trim(txtAccNo.Text), 2, "-"), _
                      "20" & medGetP(Trim(txtAccNo.Text), 2, "-"))
        strAccseq = medGetP(Trim(txtAccNo.Text), 3, "-")
    Else
        strWorkarea = medGetP(Trim(lblBarNo.Caption), 1, "-")
        strAccdt = IIf(medGetP(Trim(lblBarNo.Caption), 2, "-") Like "99*", _
                      "19" & medGetP(Trim(lblBarNo.Caption), 2, "-"), _
                      "20" & medGetP(Trim(lblBarNo.Caption), 2, "-"))
        strAccseq = medGetP(Trim(lblBarNo.Caption), 3, "-")
    End If

    Set objABOSql = New clsABOSql
    blnSave = objABOSql.SaveABOResult(strWorkarea, strAccdt, strAccseq, _
           IIf(chkVerify.Value = 1, False, True), IIf(optType(0).Value, 0, 1), _
           txtABO.Text, txtRh.Text, txtABOSub.Text, txtRhSub.Text, strRemark, strComment) ', ObjMyUser.EmpID, strVfyDt, strVfyTm)
    Set objABOSql = Nothing
    
    If blnSave Then
        MsgBox "���������� ó���Ǿ����ϴ�.", vbInformation
        Call cmdClear_Click
    Else
        MsgBox "ó������ ������ �߻��Ͽ����ϴ�.", vbExclamation
        Call cmdClear_Click
    End If
End Sub

Private Function CheckValidation() As Boolean
    CheckValidation = False
    
    If txtAccNo.Text = "" Then
        MsgBox "������ȣ(��ü��ȣ)�� �Է��Ͻʽÿ�.", vbExclamation
        txtAccNo.SetFocus
        Exit Function
    End If
    
    If optType(0).Value = False And optType(1).Value = False Then
        MsgBox "����Է� Ÿ���� �����Ͻʽÿ�.", vbExclamation
        Exit Function
    End If
    
    If txtABO.Text = "" And txtABOSub.Text = "" Then
        MsgBox "ABO����� �Է��Ͻʽÿ�.", vbExclamation
        Exit Function
    End If

    If txtRh.Text = "" And txtRhSub.Text = "" Then
        MsgBox "RH����� �Է��Ͻʽÿ�.", vbExclamation
        Exit Function
    End If

    If txtABO.Text <> "" And txtABOSub.Text <> "" Then
        MsgBox "ABO�� Subtype�� �ϳ��� �Է��Ͻʽÿ�.", vbExclamation
        Exit Function
    End If

    If txtRh.Text <> "" And txtRhSub.Text <> "" Then
        MsgBox "RH�� Subgroup�� �ϳ��� �Է��Ͻʽÿ�.", vbExclamation
        Exit Function
    End If
    
    CheckValidation = True
End Function

'Private Function Save() As Boolean
'    '------------------------------------------------------
'    '1. bbs303�� �����Ų��.
'    '2. ����, double check�Ϸ�����̸� lab302���� �����Ѵ�.
'    '
'    '* lab302�� ����� �������(rsttype)�� F�� �����Ų��.
'    '------------------------------------------------------
'    Dim comment As String
'    Dim remark As String
'
'    Dim isVerify As Boolean
'    Dim typing As Long
'
'    Dim vfydt As String
'    Dim vfytm As String
'
'End Function

Private Sub Form_Activate()
    On Error Resume Next
    txtAccNo.SetFocus
End Sub

'Private Sub cmdSave_Click()
'    Dim vWorkarea As String
'    Dim vAccdt As String
'    Dim vAccseq As String
'
'    If txtAccNo = "" Then Exit Sub
'
'    vWorkarea = medGetP(txtAccNo, 1, "-")
'    vAccdt = medGetP(txtAccNo, 2, "-")
'    vAccseq = medGetP(txtAccNo, 3, "-")
'
'    If Mid(vAccdt, 1, 1) = "9" Then
'        vAccdt = "19" & vAccdt
'    Else
'        vAccdt = "20" & vAccdt
'    End If
'
'    If Save(vWorkarea, vAccdt, vAccseq) = True Then
'        ClearAll
'    End If
'End Sub

Private Sub Form_Load()
'    IsFirst = True
    txtAccNo.Text = ""
    lblBarNo.Caption = ""
    Call ClearPtInfo
    Call ClearOthers
    
'    Call LoadTestCd
    Call LoadRemark(cboRemark)
End Sub

Private Sub ClearPtInfo()
    lblPtId.Caption = ""
    lblPtNm.Caption = ""
    lblSexAge.Caption = ""
    lblDept.Caption = ""
    lblWard.Caption = ""
    lblDoct.Caption = ""
End Sub

Private Sub ClearOthers()
    optType(0).Value = True
    optType(1).Value = False
    optType(0).Enabled = True
    optType(1).Enabled = True
    chkVerify.Value = 1
    txtABO.Text = "": lblABO.Caption = ""
    txtRh.Text = "": lblRh.Caption = ""
    txtABOSub.Text = "": lblABOSub.Caption = ""
    txtRhSub.Text = "": lblRHSub.Caption = ""
    txtComment.Text = ""
    cboRemark.ListIndex = -1
End Sub

Private Sub objComment_Selected(ByVal vSelectedComment As String)
    If vSelectedComment <> "" Then txtComment.Text = vSelectedComment
End Sub

Private Sub objPopList_SelectedItem(ByVal pSelectedItem As String)
'����Է�
    Select Case objPopList.Tag
        Case 0
            txtABO.Text = objPopList.SelectedItems(0)
            lblABO.Caption = objPopList.SelectedItems(1)
        Case 1
            txtRh.Text = objPopList.SelectedItems(0)
            lblRh.Caption = objPopList.SelectedItems(1)
        Case 2
            txtABOSub.Text = objPopList.SelectedItems(0)
            lblABOSub.Caption = objPopList.SelectedItems(1)
        Case 3
            txtRhSub.Text = objPopList.SelectedItems(0)
            lblRHSub.Caption = objPopList.SelectedItems(1)
    End Select
End Sub

Private Function GetABOResult(ByVal vIndex As String) As Recordset
    Dim strSQL As String

    'abo, rh, sub, rh
    Select Case vIndex
        Case 0 'ABO
            If optType(0).Value Then
                strSQL = GetSqlABOBackList
            ElseIf optType(1).Value Then
                strSQL = GetSqlABOFrontList
            End If
        Case 1 ' Rh
            strSQL = GetSqlRhList
        Case 2 'ABO Sub
            strSQL = GetSqlABOSubList
        Case 3 'Rh Sub
            strSQL = GetSqlRhSubList
    End Select

    If strSQL <> "" Then
        Set GetABOResult = New Recordset

        GetABOResult.Open strSQL, DBConn
    End If
End Function

Private Sub txtABO_Change()
    If lblABO.Caption <> "" Then lblABO.Caption = ""
End Sub

Private Sub txtABO_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtABO_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtABO.Text = "" Then Exit Sub
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtABO_Validate(Cancel As Boolean)
    Dim strABO As String
    
    If txtAccNo.Text = "" Then Exit Sub
    If txtABO.Text = "" Then Exit Sub
    
    If optType(0).Value Then
        strABO = GetABOFrontNm(txtABO.Text)
    ElseIf optType(1).Value Then
        strABO = GetABOBackNm(txtABO.Text)
    End If
    
    If strABO = "" Then
        Cancel = True
        MsgBox "��ϵ��� ���� �ڵ��Դϴ�.", vbExclamation
    Else
        lblABO.Caption = strABO
    End If
    
    If Cancel Then SendKeys "{Home}+{End}"
End Sub

Private Sub txtABOSub_Change()
    If lblABOSub.Caption <> "" Then lblABOSub.Caption = ""
End Sub

Private Sub txtABOSub_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtABOSub_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtABOSub.Text = "" Then Exit Sub
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtABOSub_Validate(Cancel As Boolean)
    Dim strABOSub As String
    
    If txtAccNo.Text = "" Then Exit Sub
    If txtABOSub.Text = "" Then Exit Sub
    
    strABOSub = GetABOSUBNM(txtABOSub.Text)
    
    If strABOSub = "" Then
        Cancel = True
        MsgBox "��ϵ��� ���� �ڵ��Դϴ�.", vbExclamation
    Else
        lblABOSub.Caption = strABOSub
    End If
    
    If Cancel Then SendKeys "{Home}+{End}"
End Sub

Private Sub txtAccNo_Change()
'ȭ�� ����
    Dim lngLen As Long
    Static lngAccDt As Long
    
    On Error Resume Next
    If Screen.ActiveControl.Name <> txtAccNo.Name Then Exit Sub
    
    If chkBarcode.Value = 0 Then    '������ȣ�� �Է��Ҷ��� ��ȿ
        With txtAccNo
            lngLen = Len(Trim(.Text))
            
            If lngLen < 2 Then
                lngAccDt = 0
            End If
            
            If lngLen = 2 Then
                .Text = .Text & "-"
                .SelStart = Len(.Text)
                                
                lngAccDt = GetLenOfAccDt(Mid(txtAccNo.Text, 1, 2))
            End If
            
            If lngLen > 2 And lngLen = lngAccDt + 3 And lngAccDt <> 0 Then
                .Text = .Text & "-"
                .SelStart = Len(.Text)
            End If
        End With
    End If
    
    If lblPtId.Caption <> "" Then
        lblBarNo.Caption = ""
        Call ClearPtInfo
        Call ClearOthers
    End If
End Sub

Private Sub txtAccNo_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtAccNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(txtAccNo.Text) = "" Then Exit Sub
    
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtAccNo_KeyPress(KeyAscii As Integer)
    If chkBarcode.Value = 0 Then '������ȣ�� �Է��Ҷ��� ��ȿ
        If KeyAscii = vbKeyBack Then
            With txtAccNo
                If .Text = "" Then Exit Sub
                If Mid(.Text, Len(.Text)) = "-" Then
                    If Len(.Text) > 2 Then
                        .Text = Mid(.Text, 1, Len(.Text) - 2)
                        .SelStart = Len(.Text)
                        KeyAscii = 0
                    End If
                End If
            End With
        End If
    End If
End Sub

Private Sub txtAccNo_Validate(Cancel As Boolean)
    Dim objABOSql As clsABOSql
    Dim Rs As Recordset
    Dim strAccNo As String
    
    If txtAccNo.Text = "" Then Exit Sub
    
    If chkBarcode.Value = 1 Then
        strAccNo = GetAccNo
        lblBarNo.Caption = strAccNo
    Else
        strAccNo = Trim(txtAccNo.Text)
        lblBarNo.Caption = GetSpcNo
    End If

    If strAccNo = "" Then
        MsgBox "��ȿ�� ���ڵ� ������ �ƴմϴ�.", vbExclamation
        Cancel = True
        GoTo SetMe
    End If
    
    If lblBarNo.Caption = "" Then
        MsgBox "��ȿ�� ������ȣ�� �ƴմϴ�.", vbExclamation
        Cancel = True
        GoTo SetMe
    End If

    Dim strWorkarea As String
    Dim strAccdt As String
    Dim strAccseq As String
    
    strWorkarea = medGetP(strAccNo, 1, "-")
    strAccdt = medGetP(strAccNo, 2, "-")
    strAccdt = IIf(strAccdt Like "99*", "19" & strAccdt, "20" & strAccdt)
    strAccseq = medGetP(strAccNo, 3, "-")
    
    
    Set objABOSql = New clsABOSql
    
    '�������� �б�
    Set Rs = New Recordset
    Set Rs = objABOSql.GetAccessInfo(strWorkarea, strAccdt, strAccseq)
    
    '���� Display-------------------------------------------------------
    lblPtId.Caption = Rs.Fields("ptid").Value & ""
    lblPtNm.Caption = GetPtNm(Rs.Fields("ptid").Value & "")
    lblSexAge.Caption = Rs.Fields("sex").Value & "" & "/" & (Val(Rs.Fields("ageday").Value & "") \ 365)
    lblDept.Caption = GetDeptNm(Rs.Fields("deptcd").Value & "")
    lblWard.Caption = GetWardNm(Rs.Fields("wardid").Value & "")
    lblDoct.Caption = GetDoctNm(Rs.Fields("majdoct").Value & "")
    
    '���� �б�
    If Val(Rs.Fields("stscd").Value & "") < 2 Then
        MsgBox "�������� ���� ������ȣ(��ü��ȣ)�Դϴ�.", vbExclamation
        Cancel = True
        GoTo SetMe
    End If

    If Val(Rs.Fields("stscd").Value & "") > 4 Then
        MsgBox "�̹� ����� ��ϵǾ����ϴ�.", vbExclamation
        Cancel = True
        GoTo SetMe
    End If
    
    '����ũ-------------------------------------------------------------
    cboRemark.ListIndex = -1
    If Rs.Fields("rmkcd").Value & "" & "" <> "" Then
        cboRemark.ListIndex = medComboFind(cboRemark, Rs.Fields("rmkcd").Value & "" & "")
    End If
    Set Rs = Nothing
    
    '�ڸ�Ʈ �б�
    Set Rs = New Recordset
    Set Rs = objABOSql.GetAccComment(strWorkarea, strAccdt, strAccseq)
    
    If Rs.EOF = False Then
        txtComment.Text = Rs.Fields("rsttxt").Value & ""
    End If
    Set Rs = Nothing
    
    '������ �˻簡 �ִ� �� ���� üũ
    If objABOSql.IsExistABO(strWorkarea, strAccdt, strAccseq) = False Then '�˻簡 ���� ���..
        MsgBox "������ �˻簡 �����ϴ�.", vbExclamation
        Cancel = True
        GoTo SetMe
    End If
    
    '�̹� �˻簡 �Ϸ�Ǿ����� �˻��Ѵ�.--------------------------------
    Select Case CanEditABOResult(strWorkarea, strAccdt, strAccseq)
        Case KindOfTyping.NoUse 'BBS303 ��������� ���� ���
            Cancel = True
            GoTo SetMe
        Case KindOfTyping.AnyUse
            optType(0).Enabled = True
            optType(1).Enabled = True
        Case KindOfTyping.FrontUse
            optType(0).Value = True
            optType(0).Enabled = False
            optType(1).Enabled = False
        Case KindOfTyping.BackUse
            optType(1).Value = True
            optType(0).Enabled = False
            optType(1).Enabled = False
    End Select
    
    Set objABOSql = Nothing
    
    txtABO.SetFocus
    Exit Sub
    
SetMe:
    Set Rs = Nothing
    Set objABOSql = Nothing
    If Cancel Then SendKeys "{Home}+{End}"
End Sub

Private Function GetAccNo() As String
    Dim Rs As Recordset
    Dim strSQL As String
    Dim strSpcYY As String
    Dim strSpcNo As Long
    
    strSpcYY = Mid(Trim(txtAccNo.Text), 1, P_SpcYyLength)
    If Mid(Trim(txtAccNo.Text), P_SpcYyLength + 1, P_SpcNoLength) <> "" Then
        strSpcNo = Format(Mid(Trim(txtAccNo.Text), P_SpcYyLength + 1, P_SpcNoLength), "#0")
    End If
        
    strSQL = " select vWorkarea,vAccdt,vAccseq from " & T_LAB201 & _
             " where " & DBW("spcyy=", strSpcYY) & _
             " and " & DBW("spcno=", strSpcNo)
    
    Set Rs = New Recordset
    Rs.Open strSQL, DBConn
    
    If Rs.EOF Then
        GetAccNo = ""
    Else
        GetAccNo = Rs.Fields("workarea").Value & "" & "-" & _
                   Mid(Rs.Fields("accdt").Value, 3) & "" & "-" & _
                   Rs.Fields("accseq").Value & ""
    End If
    
    Set Rs = Nothing
End Function

Private Function GetSpcNo() As String
    Dim Rs As Recordset
    Dim strSQL As String
    Dim strWorkarea As String
    Dim strAccdt As String
    Dim strAccseq As String
    
    strWorkarea = medGetP(Trim(txtAccNo.Text), 1, "-")
    strAccdt = IIf(medGetP(Trim(txtAccNo.Text), 2, "-") Like "99*", _
                  "19" & medGetP(Trim(txtAccNo.Text), 2, "-"), _
                  "20" & medGetP(Trim(txtAccNo.Text), 2, "-"))
    strAccseq = medGetP(Trim(txtAccNo.Text), 3, "-")
    
    strSQL = " select spcyy,spcno from " & T_LAB201 & _
             " where " & DBW("workarea=", strWorkarea) & _
             " and " & DBW("accdt=", strAccdt) & _
             " and " & DBW("accseq=", strAccseq)
    
    Set Rs = New Recordset
    Rs.Open strSQL, DBConn
    
    If Rs.EOF Then
        GetSpcNo = ""
    Else
        GetSpcNo = Rs.Fields("spcyy").Value & "" & Format(Rs.Fields("spcno").Value & "", LIS_BarFormat)
    End If
        
    Set Rs = Nothing
End Function

Private Function CanEditABOResult(ByVal vWorkarea As String, ByVal vAccdt As String, ByVal vAccseq As String) As Long
    '-----------------------------------------------
    '��ȯ��  < 0 : ����� �� ����.
    '        = 0 : �����ؼ� ����� �� �ִ�.
    '        = 1 : Front Typing�� �� �ִ�.
    '        = 2 : Back  Typing�� �� �ִ�.
    '-----------------------------------------------
'��� ���� : ���-> vfydt�� ��� �ִ�, �Է�->vfyid�� ��� �ִ�

    Dim objABOSql As clsABOSql
    Dim Rs As Recordset

    Dim vfydt  As String    '���� Verify�Ǹ� set�ȴ�.
    Dim strVfyid1 As String    'Front Typing �����
    Dim strVfydt1 As String    'Front Typing ����
    Dim strVfytm1 As String    'Front Typing �ð�
    Dim strVfyid2 As String    'Back  Typing �����
    Dim strVfydt2 As String    'Back  Typing ����
    Dim strVfytm2 As String    'Back  Typing �ð�

    Set objABOSql = New clsABOSql

    Set Rs = objABOSql.GetABOResultInfo(vWorkarea, vAccdt, vAccseq)
    
    If Rs.EOF Then
        Set Rs = Nothing
        Set objABOSql = Nothing
        CanEditABOResult = KindOfTyping.NoUse '��������� ���� ���
        MsgBox "�Է��� ������ �˻簡 �����ϴ�.", vbExclamation
        Exit Function
    End If
    
    With Rs
        vfydt = .Fields("vfydt").Value & ""
        
        strVfyid1 = .Fields("vfyid1").Value & ""
        strVfydt1 = .Fields("vfydt1").Value & ""
        strVfytm1 = .Fields("vfytm1").Value & ""

        strVfyid2 = .Fields("vfyid2").Value & ""
        strVfydt2 = .Fields("vfydt2").Value & ""
        strVfytm2 = .Fields("vfytm2").Value & ""
        
        '�Ѵ� ���̸� �ƹ��ų� ����Է� ����
        '���� �ϳ��� �ԷµǾ� ������ �α� ����ڿ� ���ؼ� ������ ���ǥ��, �ٸ��� ���� ���ϵ���
        '�Ѵ� ���� �ִ� ��쿡�� �̹� ���Ȯ���� �����ϱ� �������� �ؾ���.
        
        If strVfydt1 <> "" And strVfydt2 <> "" Then '�̹� ����� ��� ��ϵ� �ִ� ���
            Set Rs = Nothing
            Set objABOSql = Nothing
            CanEditABOResult = KindOfTyping.NoUse
            MsgBox "�̹� ��� ����� ��ϵǾ����ϴ�.", vbExclamation
            Exit Function
        ElseIf strVfydt1 <> "" And strVfydt2 = "" Then 'Front����� ����� ���
            If strVfyid1 = ObjMyUser.EmpID Then 'Front�� ����� ����ڿ� empid�� ���� ���
            '�̹� ����� ��ϵǾ����� �˸��� ����
                CanEditABOResult = KindOfTyping.NoUse
                MsgBox "�̹� Front Type����� ��ϵǾ����ϴ�.", vbExclamation
            ElseIf strVfyid2 = ObjMyUser.EmpID Then 'Back�� ����� ����ڿ� empid�� ���� ���
            '�̹� �Էµ� ����� ȭ�鿡 ǥ��
                CanEditABOResult = KindOfTyping.BackUse
                
                txtABO.Text = .Fields("abo2").Value & ""
                lblABO.Caption = GetABOBackNm(.Fields("abo2").Value & "")
                
                txtRh.Text = .Fields("rh2").Value & ""
                lblRh.Caption = GetRHNM(.Fields("rh2").Value & "")
                
                txtABOSub.Text = .Fields("abosub").Value & ""
                lblABOSub.Caption = GetABOSUBNM(.Fields("abosub").Value & "")
                
                txtRhSub.Text = .Fields("rhsub").Value & ""
                lblRHSub.Caption = GetRHSUBNM(.Fields("rhsub").Value & "")
            Else 'vfyid�� �ԷµǾ� �ִ��� ���� üũ
                If strVfyid2 = "" Then '����Է��� �ȵǾ� �ִ� ���
                    '����� �Է��� �� �ִ� ���·� ����� ��, �ƹ����� ���ص� ��
                    CanEditABOResult = KindOfTyping.BackUse
                Else '����Է��� �Ǿ� �ִ� ���
                    '�̹� �ٸ� ����ڿ� ���� ����� �ԷµǾ����� �˸��� ����
                    CanEditABOResult = KindOfTyping.NoUse
                    MsgBox "����� �̹� �ԷµǾ� �ֽ��ϴ�. ��� �Է� : " & GetEmpNm(strVfyid2), vbExclamation
                End If
            End If
        ElseIf strVfydt1 = "" And strVfydt2 <> "" Then 'Back����� �Է��� ���
            If strVfyid2 = ObjMyUser.EmpID Then 'Back�� ����� ����ڿ� empid�� ���� ���
            '�̹� ����� ��ϵǾ����� �˸��� ����
                CanEditABOResult = KindOfTyping.NoUse
                MsgBox "�̹� Back Type����� ��ϵǾ����ϴ�.", vbExclamation
            ElseIf strVfyid1 = ObjMyUser.EmpID Then 'Front�� ����� ����ڿ� empid�� ���� ���
            '�̹� �Էµ� ����� ȭ�鿡 ǥ��
                CanEditABOResult = KindOfTyping.FrontUse
                
                txtABO.Text = .Fields("abo1").Value & ""
                lblABO.Caption = GetABOFrontNm(.Fields("abo1").Value & "")
                
                txtRh.Text = .Fields("rh1").Value & ""
                lblRh.Caption = GetRHNM(.Fields("rh1").Value & "")
                
                txtABOSub.Text = .Fields("abosub").Value & ""
                lblABOSub.Caption = GetABOSUBNM(.Fields("abosub").Value & "")
                
                txtRhSub.Text = .Fields("rhsub").Value & ""
                lblRHSub.Caption = GetRHSUBNM(.Fields("rhsub").Value & "")
            Else 'vfyid�� �ԷµǾ� �ִ��� ���� üũ
                If strVfyid1 = "" Then '����Է��� �ȵǾ� �ִ� ���
                    '����� �Է��� �� �ִ� ���·� ����� ��, �ƹ����� ���ص� ��
                    CanEditABOResult = KindOfTyping.FrontUse
                Else '����Է��� �Ǿ� �ִ� ���
                    '�̹� �ٸ� ����ڿ� ���� ����� �ԷµǾ����� �˸��� ����
                    CanEditABOResult = KindOfTyping.NoUse
                    MsgBox "����� �̹� �ԷµǾ� �ֽ��ϴ�. ��� �Է� : " & GetEmpNm(strVfyid1), vbExclamation
                End If
            End If
        ElseIf strVfydt1 = "" And strVfydt2 = "" Then '�Ѵ� �������� �ȵ� ���..
            '����Է��� �Ǿ� ���� �� �����Ƿ� üũ�ؾ� �Ѵ�.
            If strVfyid1 <> "" And strVfyid2 <> "" Then
                If strVfyid1 = ObjMyUser.EmpID Then 'Front�Է��ڿ� empid�� ���� ��� ���ǥ��
                    CanEditABOResult = KindOfTyping.FrontUse
                    
                    txtABO.Text = .Fields("abo1").Value & ""
                    lblABO.Caption = GetABOFrontNm(.Fields("abo1").Value & "")
                    
                    txtRh.Text = .Fields("rh1").Value & ""
                    lblRh.Caption = GetRHNM(.Fields("rh1").Value & "")
                    
                    txtABOSub.Text = .Fields("abosub").Value & ""
                    lblABOSub.Caption = GetABOSUBNM(.Fields("abosub").Value & "")
                    
                    txtRhSub.Text = .Fields("rhsub").Value & ""
                    lblRHSub.Caption = GetRHSUBNM(.Fields("rhsub").Value & "")
                ElseIf strVfyid2 = ObjMyUser.EmpID Then 'Back �Է��ڿ� empid�� ���� ��� ��� ǥ��
                    CanEditABOResult = KindOfTyping.BackUse
                    
                    txtABO.Text = .Fields("abo2").Value & ""
                    lblABO.Caption = GetABOBackNm(.Fields("abo2").Value & "")
                    
                    txtRh.Text = .Fields("rh2").Value & ""
                    lblRh.Caption = GetRHNM(.Fields("rh2").Value & "")
                    
                    txtABOSub.Text = .Fields("abosub").Value & ""
                    lblABOSub.Caption = GetABOSUBNM(.Fields("abosub").Value & "")
                    
                    txtRhSub.Text = .Fields("rhsub").Value & ""
                    lblRHSub.Caption = GetRHSUBNM(.Fields("rhsub").Value & "")
                Else
                    '�̹� �ٸ� ����ڿ� ���� ����� �ԷµǾ����� �˸��� ����
                    CanEditABOResult = KindOfTyping.NoUse
                    MsgBox "�ٸ� ����ڿ� ���� ����� �ԷµǾ����ϴ�.", vbExclamation
                End If
            ElseIf strVfyid1 <> "" And strVfyid2 = "" Then
                If strVfyid1 = ObjMyUser.EmpID Then 'Front�Է��ڿ� empid�� ���� ��� ���ǥ��
                    CanEditABOResult = KindOfTyping.FrontUse
                    
                    txtABO.Text = .Fields("abo1").Value & ""
                    lblABO.Caption = GetABOFrontNm(.Fields("abo1").Value & "")
                    
                    txtRh.Text = .Fields("rh1").Value & ""
                    lblRh.Caption = GetRHNM(.Fields("rh1").Value & "")
                    
                    txtABOSub.Text = .Fields("abosub").Value & ""
                    lblABOSub.Caption = GetABOSUBNM(.Fields("abosub").Value & "")
                    
                    txtRhSub.Text = .Fields("rhsub").Value & ""
                    lblRHSub.Caption = GetRHSUBNM(.Fields("rhsub").Value & "")
                Else
                    '�̹� �ٸ� ����ڿ� ���� ����� �ԷµǾ����� �˸��� ����
                    CanEditABOResult = KindOfTyping.NoUse
                    MsgBox "����� �̹� �ԷµǾ� �ֽ��ϴ�. ��� �Է� : " & GetEmpNm(strVfyid1), vbExclamation
                End If
            ElseIf strVfyid1 = "" And strVfyid2 <> "" Then
                If strVfyid2 = ObjMyUser.EmpID Then 'Back�Է��ڿ� empid�� ���� ��� ���ǥ��
                    CanEditABOResult = KindOfTyping.BackUse
                    
                    txtABO.Text = .Fields("abo2").Value & ""
                    lblABO.Caption = GetABOFrontNm(.Fields("abo2").Value & "")
                    
                    txtRh.Text = .Fields("rh2").Value & ""
                    lblRh.Caption = GetRHNM(.Fields("rh2").Value & "")
                    
                    txtABOSub.Text = .Fields("abosub").Value & ""
                    lblABOSub.Caption = GetABOSUBNM(.Fields("abosub").Value & "")
                    
                    txtRhSub.Text = .Fields("rhsub").Value & ""
                    lblRHSub.Caption = GetRHSUBNM(.Fields("rhsub").Value & "")
                Else
                    '�̹� �ٸ� ����ڿ� ���� ����� �ԷµǾ����� �˸��� ����
                    CanEditABOResult = KindOfTyping.NoUse
                    MsgBox "����� �̹� �ԷµǾ� �ֽ��ϴ�. ��� �Է� : " & GetEmpNm(strVfyid2), vbExclamation
                End If
            ElseIf strVfyid1 = "" And strVfyid2 = "" Then
                CanEditABOResult = KindOfTyping.AnyUse
            End If
        End If
   End With
    
    Set Rs = Nothing
    Set objABOSql = Nothing
End Function

Private Sub txtRh_Change()
    If lblRh.Caption <> "" Then lblRh.Caption = ""
End Sub

Private Sub txtRh_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtRh_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtRh.Text = "" Then Exit Sub
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtRh_Validate(Cancel As Boolean)
    Dim strRh As String
    
    If txtAccNo.Text = "" Then Exit Sub
    If txtRh.Text = "" Then Exit Sub
    
    strRh = GetRHNM(txtRh.Text)
    
    If strRh = "" Then
        Cancel = True
        MsgBox "��ϵ��� ���� �ڵ��Դϴ�.", vbExclamation
    Else
        lblRh.Caption = strRh
    End If
    
    If Cancel Then SendKeys "{Home}+{End}"
End Sub

Private Sub txtRhSub_Change()
    If lblRHSub.Caption <> "" Then lblRHSub.Caption = ""
End Sub

Private Sub txtRhSub_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtRhSub_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtRhSub.Text = "" Then Exit Sub
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtRhSub_Validate(Cancel As Boolean)
    Dim strRhSub As String
    
    If txtAccNo.Text = "" Then Exit Sub
    If txtRhSub.Text = "" Then Exit Sub
    
    strRhSub = GetRHSUBNM(txtRhSub.Text)
    
    If strRhSub = "" Then
        Cancel = True
        MsgBox "��ϵ��� ���� �ڵ��Դϴ�.", vbExclamation
    Else
        lblRHSub.Caption = strRhSub
    End If
    
    If Cancel Then SendKeys "{Home}+{End}"
End Sub
