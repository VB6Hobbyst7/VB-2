VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmPatSearch 
   BorderStyle     =   1  '���� ����
   Caption         =   "�˻��� ��ȸ"
   ClientHeight    =   8925
   ClientLeft      =   7440
   ClientTop       =   2250
   ClientWidth     =   8355
   Icon            =   "frmPatSearch.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8925
   ScaleWidth      =   8355
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.Frame fraWork 
      Height          =   765
      Left            =   90
      TabIndex        =   29
      Top             =   90
      Width           =   8145
      Begin VB.ComboBox cboChk 
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
         ItemData        =   "frmPatSearch.frx":014A
         Left            =   4020
         List            =   "frmPatSearch.frx":0157
         TabIndex        =   42
         Top             =   240
         Width           =   825
      End
      Begin VB.TextBox txtSNo 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BeginProperty Font 
            Name            =   "����"
            Size            =   18
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7410
         TabIndex        =   40
         Text            =   "1"
         Top             =   180
         Width           =   495
      End
      Begin VB.CommandButton cmdDown 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   6420
         Style           =   1  '�׷���
         TabIndex        =   39
         Top             =   180
         Width           =   375
      End
      Begin VB.CommandButton cmdUp 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   6030
         Style           =   1  '�׷���
         TabIndex        =   38
         Top             =   180
         Width           =   375
      End
      Begin VB.OptionButton optState 
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   1260
         TabIndex        =   37
         Top             =   0
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.OptionButton optState 
         Caption         =   "���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   2010
         TabIndex        =   36
         Top             =   0
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.OptionButton optState 
         Caption         =   "���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   2790
         TabIndex        =   35
         Top             =   0
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "��ũ��ȸ"
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
         Left            =   4860
         TabIndex        =   30
         Top             =   180
         Width           =   1125
      End
      Begin MSComCtl2.DTPicker dtpEDate 
         Height          =   345
         Left            =   2580
         TabIndex        =   31
         Top             =   240
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   121569281
         CurrentDate     =   40248
      End
      Begin MSComCtl2.DTPicker dtpSDate 
         Height          =   345
         Left            =   990
         TabIndex        =   32
         Top             =   240
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   121241601
         CurrentDate     =   40248
      End
      Begin VB.Label Label1 
         Caption         =   "���۹�ȣ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   405
         Left            =   6900
         TabIndex        =   41
         Top             =   210
         Width           =   585
      End
      Begin VB.Label Label13 
         Caption         =   "ó������"
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
         Left            =   90
         TabIndex        =   34
         Top             =   300
         Width           =   915
      End
      Begin VB.Label Label9 
         Appearance      =   0  '���
         BackColor       =   &H80000005&
         BackStyle       =   0  '����
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2430
         TabIndex        =   33
         Top             =   330
         Width           =   105
      End
   End
   Begin VB.Frame sspOrder 
      Caption         =   "Frame1"
      Height          =   3855
      Left            =   10260
      TabIndex        =   9
      Top             =   4020
      Visible         =   0   'False
      Width           =   7755
      Begin VB.CheckBox chkAllOrder 
         Caption         =   "Check1"
         Height          =   345
         Left            =   3240
         TabIndex        =   24
         Top             =   180
         Width           =   225
      End
      Begin VB.TextBox txtNo 
         Appearance      =   0  '���
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1050
         TabIndex        =   17
         Top             =   90
         Width           =   1395
      End
      Begin VB.TextBox txtPID 
         Appearance      =   0  '���
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1050
         TabIndex        =   16
         Top             =   510
         Width           =   1395
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  '���
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1050
         TabIndex        =   15
         Top             =   930
         Width           =   1395
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "Ȯ��"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   90
         TabIndex        =   14
         Top             =   3060
         Width           =   1215
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "�ݱ�"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1380
         TabIndex        =   13
         Top             =   3060
         Width           =   1215
      End
      Begin VB.TextBox txtSex 
         Appearance      =   0  '���
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1050
         TabIndex        =   12
         Top             =   1350
         Width           =   915
      End
      Begin VB.TextBox txtAge 
         Appearance      =   0  '���
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1050
         TabIndex        =   11
         Top             =   1770
         Width           =   915
      End
      Begin VB.TextBox txtDate 
         Appearance      =   0  '���
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   60
         TabIndex        =   10
         Top             =   2700
         Visible         =   0   'False
         Width           =   1965
      End
      Begin FPSpread.vaSpread vasOrder 
         Height          =   3555
         Left            =   2760
         TabIndex        =   18
         Top             =   0
         Width           =   4455
         _Version        =   393216
         _ExtentX        =   7858
         _ExtentY        =   6271
         _StockProps     =   64
         ColHeaderDisplay=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   10
         MaxRows         =   100
         ScrollBars      =   2
         SpreadDesigner  =   "frmPatSearch.frx":016D
      End
      Begin VB.Label Label7 
         BackStyle       =   0  '����
         Caption         =   "��ü��ȣ"
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
         Left            =   0
         TabIndex        =   23
         Top             =   150
         Width           =   1005
      End
      Begin VB.Label Label6 
         BackStyle       =   0  '����
         Caption         =   "ȯ�ڹ�ȣ"
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
         Left            =   0
         TabIndex        =   22
         Top             =   570
         Width           =   1005
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '����
         Caption         =   "ȯ���̸�"
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
         Left            =   0
         TabIndex        =   21
         Top             =   990
         Width           =   1005
      End
      Begin VB.Label Label4 
         BackStyle       =   0  '����
         Caption         =   "����"
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
         Left            =   0
         TabIndex        =   20
         Top             =   1410
         Width           =   1005
      End
      Begin VB.Label Label5 
         BackStyle       =   0  '����
         Caption         =   "����"
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
         Left            =   0
         TabIndex        =   19
         Top             =   1830
         Width           =   1005
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "�ݱ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   6945
      Style           =   1  '�׷���
      TabIndex        =   2
      Top             =   8310
      Width           =   1320
   End
   Begin VB.CheckBox chkAll 
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   990
      TabIndex        =   6
      Top             =   1005
      Width           =   165
   End
   Begin VB.CommandButton cmdOrder 
      Caption         =   "Order ����"
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   11310
      Style           =   1  '�׷���
      TabIndex        =   5
      Top             =   8790
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.CommandButton cmdWorkList 
      Caption         =   "���� ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   5580
      Style           =   1  '�׷���
      TabIndex        =   0
      Top             =   8310
      Width           =   1335
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   4200
      Style           =   1  '�׷���
      TabIndex        =   1
      Top             =   8310
      Visible         =   0   'False
      Width           =   1320
   End
   Begin FPSpread.vaSpread vasPrint 
      Height          =   2610
      Left            =   9570
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   5280
      _Version        =   393216
      _ExtentX        =   9313
      _ExtentY        =   4604
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
      GrayAreaBackColor=   16777215
      MaxCols         =   8
      MaxRows         =   100
      ScrollBars      =   2
      ShadowColor     =   15526606
      ShadowDark      =   13815180
      SpreadDesigner  =   "frmPatSearch.frx":1218
   End
   Begin FPSpread.vaSpread vasCode 
      Height          =   3645
      Left            =   15120
      TabIndex        =   4
      Top             =   180
      Visible         =   0   'False
      Width           =   2745
      _Version        =   393216
      _ExtentX        =   4842
      _ExtentY        =   6429
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
      SpreadDesigner  =   "frmPatSearch.frx":2441
   End
   Begin FPSpread.vaSpread vasList 
      Height          =   7275
      Left            =   60
      TabIndex        =   7
      Top             =   930
      Width           =   8205
      _Version        =   393216
      _ExtentX        =   14473
      _ExtentY        =   12832
      _StockProps     =   64
      ColHeaderDisplay=   0
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      MaxCols         =   12
      MaxRows         =   100
      ScrollBars      =   2
      ShadowColor     =   15987699
      ShadowDark      =   13815180
      SpreadDesigner  =   "frmPatSearch.frx":26B4
   End
   Begin MSComCtl2.DTPicker dtpStopDt 
      Height          =   345
      Left            =   14280
      TabIndex        =   25
      Top             =   9750
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   121962497
      CurrentDate     =   40248
   End
   Begin MSComCtl2.DTPicker dtpStartDt 
      Height          =   345
      Left            =   12750
      TabIndex        =   26
      Top             =   9750
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   121962497
      CurrentDate     =   40248
   End
   Begin VB.Label Label20 
      Caption         =   "��ȸ����"
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
      Left            =   11850
      TabIndex        =   28
      Top             =   9810
      Width           =   915
   End
   Begin VB.Label Label12 
      Appearance      =   0  '���
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   14160
      TabIndex        =   27
      Top             =   9840
      Width           =   105
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '����
      Caption         =   "����Ϸ� : ������, �̿Ϸ� : ������"
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   8580
      TabIndex        =   8
      Top             =   6180
      Visible         =   0   'False
      Width           =   3675
   End
End
Attribute VB_Name = "frmPatSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iIndex As Integer

Public glRow As Long
Public gOCnt As Integer
Public gCount As String

Private Sub btnClear_Click()
    ClearSpread vasList
    
End Sub

'Private Sub btnSch_Click()
'    Dim sSch1, sSch2 As String
'    Dim iRow As Integer
'    Dim i As Integer
'    Dim sCnt As String
'    Dim sExamCode As String
'    Dim sExamName As String
'
'    'vasList.MaxRows = 100
'
'    'üũ, Rack, Pos, SampleNo, ȯ�ڹ�ȣ, ȯ���̸�, ����, ����, �ֹι�ȣ, ��������
'    '�˻����
'    sSch1 = Format(dtpSDate.Text, "yymmdd") & "0001"
'    sSch2 = Format(dtpEDate.Text, "yymmdd") & "9999"
'
'    SQL = "SELECT a.PTNO, " & vbCrLf
'    SQL = SQL & " a.SNAME, a.SEX, a.AGE, '', " & vbCrLf
'    SQL = SQL & " '20' || substr(a.SPECNO, 1, 6), substr(a.SPECNO, 7, 4), a.SPECNO, count(SUBCODE) " & vbCrLf
'    SQL = SQL & "From TWEXAM_SPECMST a, TWEXAM_RESULTC b " & vbCrLf
'    SQL = SQL & "WHERE a.SPECNO = '" & Trim(txtBarCode) & "' " & vbCrLf
'    SQL = SQL & "  AND b.SPECNO = a.SPECNO " & vbCrLf
'    SQL = SQL & "  AND b.SUBCODE In (" & gAllExam & ") " & vbCrLf
'    SQL = SQL & "  AND b.STATUS in ('2','3') " & vbCrLf
'    SQL = SQL & "Group by a.PTNO, " & vbCrLf
'    SQL = SQL & " a.SNAME, a.SEX, a.AGE, '', a.BDATE, " & vbCrLf
'    SQL = SQL & " '20' || substr(a.SPECNO, 1, 6), substr(a.SPECNO, 7, 4), a.SPECNO "
'    Res = db_select_Vas(gServer, SQL, vasList, vasList.DataRowCnt + 1, 4)
'    If Res = -1 Then
'        SaveQuery SQL
'        Exit Sub
'    End If
'
'    'vasSort vasList, 11
'
'    For iRow = 1 To vasList.DataRowCnt
'        sExamCode = ""
'        sExamName = ""
'        ClearSpread vasOrder
'
'        SQL = "SELECT SUBCODE " & vbCrLf
'        SQL = SQL & "From TWEXAM_RESULTC  " & vbCrLf
'        SQL = SQL & "WHERE SPECNO = '" & Trim(GetText(vasList, iRow, 11)) & "' " & vbCrLf
'        SQL = SQL & "  AND SUBCODE In (" & gAllExam & ") " & vbCrLf
'        SQL = SQL & "  AND STATUS in ('2','3') "
'        Res = db_select_Vas(gServer, SQL, vasOrder)
'        vasSort vasOrder, 1
'
'        For i = 1 To vasOrder.DataRowCnt
'            sExamCode = sExamCode & "'" & Trim(GetText(vasOrder, i, 1)) & "',"
'        Next i
'        If Len(sExamCode) > 0 Then
'            sExamCode = Left(sExamCode, Len(sExamCode) - 1)
'        End If
'        ClearSpread vasOrder
'        SQL = "Select examname From equipexam" & vbCrLf & _
'              " Where Equipno = '" & gEquip & "' " & vbCrLf & _
'              "  and examcode in (" & sExamCode & ") "
'        Res = db_select_Vas(gLocal, SQL, vasOrder)
'        For i = 1 To vasOrder.DataRowCnt
'            sExamName = sExamName & Trim(GetText(vasOrder, i, 1)) & "/"
'        Next i
'        If Len(sExamName) > 0 Then
'            sExamName = Left(sExamName, Len(sExamName) - 1)
'
'            vasList.Row = iRow
'            vasList.Col = 1
'            vasList.Value = 1
'        End If
'        vasList.SetText 12, iRow, sExamName
'
'        vasList.Row = iRow
'        vasList.Col = 2
'        vasList.TypeComboBoxCurSel = 0
'
'        SQL = "select state, SEQNO from Worklist " & vbCrLf & _
'              "WHERE examdate = '" & Format(CDate(frmInterface.txtToday.Text), "yyyymmdd") & "' " & vbCrLf & _
'              "  AND SampleID = '" & Trim(GetText(vasList, iRow, 11)) & "' "
'        Res = db_select_Col(gLocal, SQL)
'        vasList.SetText 3, iRow, Trim(gReadBuf(1))
'        Select Case Trim(gReadBuf(0))
'        Case "A"
'            SetBackColor vasList, iRow, iRow, 5, 5, 255, 255, 112
'        Case "B", "C"
'            SetBackColor vasList, iRow, iRow, 5, 5, 202, 255, 112
'        Case Else
'            SetBackColor vasList, iRow, iRow, 5, 5, 255, 255, 255
'        End Select
'    Next iRow
'
'    vasList.MaxRows = vasList.DataRowCnt
'    vasList.RowHeight(-1) = 13.3
'
'End Sub

Private Sub chkAll_Click()
    Dim iRow As Integer
    
    If chkAll.Value = 1 Then
        For iRow = 1 To vasList.DataRowCnt
            vasList.Row = iRow
            vasList.Col = 1
            
            vasList.Value = 1
        Next iRow
    ElseIf chkAll.Value = 0 Then
        For iRow = 1 To vasList.DataRowCnt
            vasList.Row = iRow
            vasList.Col = 1
            
            vasList.Value = 0
        Next iRow
    End If
End Sub

Private Sub chkAllOrder_Click()
    If chkAllOrder.Value = 1 Then
        vasOrder.Row = -1
        vasOrder.Col = 1
        vasOrder.Value = 1
    Else
        vasOrder.Row = -1
        vasOrder.Col = 1
        vasOrder.Value = 0
    End If
End Sub

'Private Sub cmdCalendar_Click(Index As Integer)
'    iIndex = Index
'    If Index = 0 Then
'        monvCal.Left = dtpSDate.Left
'        monvCal.Top = 570
'        monvCal.Visible = True
'
'        monvCal.Value = dtpSDate.Text
'    ElseIf Index = 1 Then
'        monvCal.Left = dtpEDate.Left
'        monvCal.Top = 570
'        monvCal.Visible = True
'
'        monvCal.Value = dtpEDate.Text
'    End If
'    'monvCal.Visible = True
'End Sub

Private Sub cmdClose_Click()
'    txtDate.Text = ""
'    txtPID.Text = ""
'    txtName.Text = ""
'    txtSex.Text = ""
'    txtAge.Text = ""
'
'    ClearSpread vasOrder
'
    sspOrder.Visible = False
End Sub

Private Sub cmdDown_Click()
    Dim lRow As Long
    
    lRow = vasList.ActiveRow
    
    vasList.SwapRange 1, lRow, vasList.MaxCols, lRow, 1, lRow + 1
    vasActiveCell vasList, lRow + 1, 2
    vasList_Click 2, lRow + 1
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

'Private Sub cmdOK_Click()
''Local�� ȯ�ڿ� ���� �˻��׸� �����ϱ�
'Dim sCnt As String
'Dim iRow As Integer
'Dim sExamCode As String
'Dim sEquipCode As String
'Dim sAge As String
'Dim i As Integer
'
'    sCnt = ""
'
'    SQL = " Select count(*) From pat_res " & vbCrLf & _
'          " Where examdate = '" & Trim(txtDate) & "' " & vbCrLf & _
'          " And equipno = '" & gEquip & "' " & vbCrLf & _
'          " And barcode = '" & Trim(txtNo) & "' " & vbCrLf & _
'          " And sendflag = 'O' "
'    Res = db_select_Var(gLocal, SQL, sCnt)
'
'    If sCnt = "" Then
'        sCnt = "0"
'    End If
'
'    If txtAge.Text = "" Then
'        txtAge.Text = "0"
'    Else
'        sAge = Trim(txtAge.Text)
'    End If
'
'    If sCnt > 0 Then
'            SQL = " Delete From pat_res " & vbCrLf & _
'                  " Where examdate = '" & Trim(txtDate) & "' " & vbCrLf & _
'                  " And equipno = '" & gEquip & "' " & vbCrLf & _
'                  " And barcode = '" & Trim(txtNo.Text) & "' " & vbCrLf & _
'                  " And sendflag = 'O' "
'            Res = SendQuery(gLocal, SQL)
'
'            If Res = -1 Then
'                SaveQuery SQL
'            End If
'    End If
'
'    For iRow = 1 To vasOrder.DataRowCnt
'        vasOrder.Row = iRow
'        vasOrder.Col = 1
'
'        If vasOrder.Value = 1 Then
'            sExamCode = Trim(GetText(vasOrder, iRow, 2))
'            sEquipCode = GetEquip_ExamCode(sExamCode)
'
'            SQL = " Insert Into pat_res(examdate, equipno, barcode, equipcode,  " & vbCrLf & _
'                  " examcode, pid, pname, psex, page, resdate, sendflag)  " & vbCrLf & _
'                  " Values ( '" & Trim(txtDate) & "', '" & gEquip & "', '" & Trim(txtNo.Text) & "' , '" & Trim(sEquipCode) & "', " & vbCrLf & _
'                  " '" & sExamCode & "', '" & Trim(txtPID.Text) & "', " & vbCrLf & _
'                  " '" & Trim(txtName.Text) & "', '" & Trim(txtSex.Text) & "', " & sAge & ", " & vbCrLf & _
'                  " '" & Trim(GetDateFull) & "', 'O') "
'            Res = SendQuery(gLocal, SQL)
'
'            If Res = -1 Then
'                SaveQuery SQL
'            End If
'        ElseIf vasOrder.Value = 0 Then
'            If sCnt = 0 Then
'
'            ElseIf sCnt > 0 Then
'                sExamCode = Trim(GetText(vasOrder, iRow, 2))
'
'                SQL = " Delete From pat_res " & vbCrLf & _
'                      " Where examdate = '" & Trim(txtDate) & "' " & vbCrLf & _
'                      " And equipno = '" & gEquip & "' " & vbCrLf & _
'                      " And barcode = '" & Trim(txtNo.Text) & "' " & vbCrLf & _
'                      " And examcode = '" & sExamCode & "' "
'                Res = SendQuery(gLocal, SQL)
'
'                If Res = -1 Then
'                    SaveQuery SQL
'                End If
'            End If
'        End If
'    Next iRow
'
'    sspOrder.Visible = False
'End Sub


'
'Private Sub cmdOrder_Click()
'    Dim llRow_Order As Long
'    Dim iRow As Integer
'    Dim jRow As Integer
'    Dim I As Integer
'    Dim iCnt As Integer
'
'    Dim sEquipCode As String
'    Dim sOrderCode As String
'    Dim sOrder As String
'
'    Dim sID As String
'
'    Dim lsCurDate As String
'    Dim lsSampleNo As String
'    Dim lsType As String
'    Dim lsTypeSelect As Integer
'
'    If IsNumeric(txtRack) = False Or IsNumeric(txtPos) = False Then
'        MsgBox "Rack, Pos�� Ȯ���ϼ���!", vbCritical, "�˸�"
'        Exit Sub
'    End If
'
''    If IsNumeric(txtStart) Then
''        lsSampleNo = Trim(txtStart)
''    Else
''        lsSampleNo = "1"
''    End If
'
'    lsCurDate = Format(Date, "yyyymmdd") & Format(Time, "hhnnss")
'
''    ClearSpread frmInterface.vasOrder
'
'    llRow_Order = 1
'
'    For iRow = 1 To vasList.DataRowCnt
'        If Trim(GetText(vasList, iRow, 3)) <> "" Then
'            SetText vasList, Format(Trim(GetText(vasList, iRow, 3)), "0#"), iRow, 3
'        End If
'    Next iRow
'
'    vasSort vasList, 3
'
'    For iRow = 1 To vasList.DataRowCnt
'        vasList.Row = iRow
'        vasList.Col = 1
'
'        If vasList.Value = 1 Then
'            'ó�氡������
'            sOrderCode = ""
'
'            vasList.SetText 3, iRow, txtPos
'
'            txtPos = CStr(CInt(txtPos) + 1)
'
'            ClearSpread vasCode
'
'            sID = Trim(GetText(vasList, iRow, 10))     '��ü��ȣ
'
''            If Trim(GetText(vasList, iRow, 3)) = "" Then
''                SetText vasList, txtPos, iRow, 3
''            End If
''
''            lsSampleNo = CLng(lsSampleNo) + 1
''            txtStart = lsSampleNo
'
'            frmInterface.vasOrder.SetText 1, llRow_Order, sID
'            frmInterface.vasOrder.SetText 2, llRow_Order, Trim(txtRack)
'            'frmInterface.vasOrder.SetText 3, llRow_Order, Trim(txtPos)
'            frmInterface.vasOrder.SetText 3, llRow_Order, Trim(GetText(vasList, iRow, 3))
'            frmInterface.vasOrder.SetText 4, llRow_Order, ""
'
'            llRow_Order = llRow_Order + 1
'            If llRow_Order > frmInterface.vasOrder.MaxRows Then
'                frmInterface.vasOrder.MaxRows = llRow_Order
'            End If
'
''            If IsNumeric(txtPos) Then
''                txtPos = CInt(txtPos) + 1
''            End If
'        End If
'    Next iRow
'
'    'WorkList ����
'    cmdWorkList_Click
'
''    gRecodeType = "Q"
''
''    comSend = "stENQ"
'
'    If frmInterface.vasOrder.DataRowCnt > 0 Then
'        gOrderMessage = Trim(GetText(frmInterface.vasOrder, 1, 1))
'        gRack = Trim(GetText(frmInterface.vasOrder, 1, 2))
'        gPos = Trim(GetText(frmInterface.vasOrder, 1, 3))
'        gSampleNo = ""
'
'        gOrderCnt = 0
'
'        gPreMsg = chrENQ
'        Save_Raw_Data "[Tx]" & gPreMsg
'        frmInterface.MSComm1.Output = gPreMsg
'    End If
'
'    Unload Me
'End Sub

Private Sub cmdPrint_Click()
Dim iRow As Integer
Dim j As Integer

Dim sCurDate As String
Dim sSerDate As String
Dim sHead As String
Dim sFoot As String
    
    ClearSpread vasPrint

    j = 1

    'If optGubun(1).Value = True Then
    '    vasPrint.RowHeight(-1) = 39.2
    'Else
        vasPrint.RowHeight(-1) = 25.9
    'End If
    
    For iRow = 1 To vasList.DataRowCnt
        vasList.Row = iRow
        vasList.Col = 1

        If vasList.Value = 1 Then
            SetText vasPrint, Trim(GetText(vasList, iRow, 11)), j, 1     '��ü��ȣ
            SetText vasPrint, Trim(GetText(vasList, iRow, 4)), j, 2     'ȯ�ڹ�ȣ
            SetText vasPrint, Trim(GetText(vasList, iRow, 5)), j, 3     'ȯ���̸�

            SetText vasPrint, Trim(GetText(vasList, iRow, 6)), j, 4     '����
            SetText vasPrint, Trim(GetText(vasList, iRow, 7)), j, 5     '����
            'SetText vasPrint, Trim(GetText(vasList, iRow, 7)), j, 6     '�ֹι�ȣ
            SetText vasPrint, Trim(GetText(vasList, iRow, 9)), j, 7     'ó������
            SetText vasPrint, Trim(GetText(vasList, iRow, 12)), j, 8     'ó������
            
            j = j + 1
        End If
    Next iRow
    
    If vasPrint.DataRowCnt < 1 Then
        MsgBox "����� �ڷᰡ �����ϴ�.", , "�� ��"
        Exit Sub
    End If
    
    sCurDate = GetDateFull
    
    sSerDate = Trim(dtpSDate.Value) & " - " & Trim(dtpEDate.Value)
    
    '2004/08/11 �̻��� - ���ι��⿡�� ���ι������� ����
    vasPrint.PrintOrientation = 1   ' SS_PRINTORIENT_PORTRAIT
    vasPrint.PrintAbortMsg = "�μ��� �Դϴ� ..."
    vasPrint.PrintJobName = "WorkList ���"
    

    sHead = "/fn""�ü�ü"" /fz""12"" /fb1 /fi0 /fu0 " & "/c" & "�� WorkList ��" & "/n/n " & _
            "/fn""����ü"" /fz""10"" /fb0 /fi0 /fu0 " & "/c" & "ó������ : " & dtpSDate & " ~ " & dtpEDate
    'If optGubun(0).Value = True Then
    '    sHead = sHead & " (����)" & "/n/n"
    'ElseIf optGubun(1).Value = True Then
    '    sHead = sHead & " (����)" & "/n/n"
    'End If

    sFoot = "/fn""����ü"" /fz""10"" /fb1 /fi0 /fu0 " & "/l" & sCurDate & "/fn""�ü�ü"" /fz""11"" /fb1 /fi0 /fu0 /r" & " �˻��"
    
    vasPrint.PrintHeader = sHead
    vasPrint.PrintFooter = sFoot

    vasPrint.PrintMarginTop = 680
    vasPrint.PrintMarginBottom = 680
'���� SS�� ���Ī���� �����
'    vaslist.PrintMarginLeft = 720
    vasPrint.PrintMarginLeft = 0
    vasPrint.PrintMarginRight = 0
    
    vasPrint.PrintColor = True
    vasPrint.PrintGrid = True
    
'Set printing range
    vasPrint.PrintType = 0  'SS_PRINT_ALL(default)

    vasPrint.PrintShadows = True

    vasPrint.Action = 13 'SS_ACTION_PRINT
End Sub

'Private Sub cmdSearch_1_Click()
'    Dim sSch1, sSch2 As String
'    Dim iRow As Integer
'    Dim sCnt As String
'
'    ClearSpread vasList
'
'    vasList.MaxRows = 100
'
'
'    'üũ, Rack, Pos, SampleNo, ȯ�ڹ�ȣ, ȯ���̸�, ����, ����, �ֹι�ȣ, ��������
'    '�˻����
'    sSch1 = Format(dtpSDate.Text, "yyyy-mm-dd")
'    sSch2 = Format(dtpEDate.Text, "yyyy-mm-dd")
'
'    SQL = " Select max(a.DR_CHART), b.PE_SUJINJA, '', '', b.PE_JUMIN, a.DR_DATE, '', '' " & vbCrLf & _
'          " From DEPARTDAT a, PERSON b " & vbCrLf & _
'          " Where a.DR_DATE between '" & sSch1 & "' and '" & sSch2 & "' " & vbCrLf & _
'          " And a.DR_CODE in (" & gAllExam & ") " & vbCrLf & _
'          " And a.DR_CHART = b.PE_CHART "
'
''    If optState(0).Value = True Then        '����
''        SQL = SQL & vbCrLf & _
''              " And c.GD_RESULT = ''  "
''    ElseIf optState(1).Value = True Then    '���
''        SQL = SQL & vbCrLf & _
''              " And c.GD_RESULT <> '' "
''    ElseIf optState(2).Value = True Then
''    End If
'
'        SQL = SQL & vbCrLf & _
'              " Group by b.PE_SUJINJA, b.PE_JUMIN, a.DR_DATE " & vbCrLf & _
'              " Order by 1 "
'
'    Res = db_select_Vas(gServer, SQL, vasList, 1, 5)
'    If Res = -1 Then
'        SaveQuery SQL
'        Exit Sub
'    End If
'
'    vasList.MaxRows = vasList.DataRowCnt
'
'    For iRow = 1 To vasList.DataRowCnt
'        CalAgeSex Trim(GetText(vasList, iRow, 9)), Format(dtpSDate.Text, "yyyy/mm/dd")
'        If gPatGen.Age = "" Then
'            gPatGen.Age = 0
'        End If
'        SetText vasList, gPatGen.Sex, iRow, 7
'        SetText vasList, gPatGen.Age, iRow, 8
'
'        sCnt = ""
'
'        SQL = " Select count(GD_CODE) From GUMSADAT " & vbCrLf & _
'              " Where GD_DATE = '" & Trim(GetText(vasList, iRow, 10)) & "' " & vbCrLf & _
'              " And GD_CHART = '" & Trim(GetText(vasList, iRow, 5)) & "' " & vbCrLf & _
'              " And GD_CODE in (" & gAllExam & ") "
'        Res = db_select_Var(gServer, SQL, sCnt)
'
'        If sCnt = "" Then
'            sCnt = "0"
'        End If
'
'        If sCnt = "0" Then
'            SetForeColor vasList, iRow, iRow, 0, 0, 0
'        ElseIf CInt(sCnt) > 0 Then
'            SetForeColor vasList, iRow, iRow, 250, 0, 0
'        End If
'    Next iRow
'
'End Sub

Private Sub cmdSearch_Click()
    Dim sSch1, sSch2 As String
    Dim iRow As Integer
    Dim i, X As Long
    Dim sCnt As String
    Dim sExamCode As String
    Dim sExamName As String
    Dim FilNum
    Dim TxtString As String
    Dim TxtRece As String
    Dim PChartNum As String
    Dim PName As String
    Dim PJumin As String
    Dim PID As String
    Dim PExamCode As String
    Dim PReceDate As String
    Dim PAge As String
    Dim pSex As String
    Dim STxt, NumTxt As Long
    Dim SQL As String
    Dim PEquipno As String
    Dim PExamname As String
    Dim PEquipCode As String
    Dim j As Long
    Dim BarFlag As Integer
    Dim TxtPat As String
    Dim TestNum, IOGubun As String
    Dim FindFile As String
    Dim StartDate As String
    Dim EndDate As String
    
    Dim strSaveSeq  As String
    Dim strRsltDate As String
    Dim strHospDate As String
    
    Dim strYM   As String
    Dim strDay  As String
    Dim strDay1  As String
    Dim strDay2  As String
    
    Dim strRcpDate As String


    sSch1 = Format(dtpSDate.Value, "yyyy.mm.dd")
    sSch2 = Format(dtpEDate.Value, "yyyy.mm.dd")
    
    strDay1 = Format(dtpSDate.Value, "dd")
    strDay2 = Format(dtpEDate.Value, "dd")
    
    'strYM = Format(Now, "yyyymm")
    strYM = Format(dtpSDate.Value, "yyyymm")
    strDay = Format(Now, "dd")
    
    ClearSpread vasList
    vasList.MaxRows = 0
    
    '-- ����
    If cboChk.ListIndex = 1 Then
              SQL = "SELECT distinct b.cdate as ��������, a.pno as ��Ʈ��ȣ, a.pname as �̸�" & vbCr
        SQL = SQL & " FROM mt_patinet a " & vbCr
        SQL = SQL & "    , his_chart02 b " & vbCr
        SQL = SQL & "    , his_chart03 c " & vbCr
        SQL = SQL & " WHERE b.cdate  BETWEEN '" & sSch1 & "' AND '" & sSch2 & "'" & vbCr
        SQL = SQL & "   AND a.pno = b.cpno  " & vbCr
        SQL = SQL & "   AND a.pno = c.cpno  " & vbCr
        SQL = SQL & " ORDER BY b.cdate, a.pno  "
        
        SetRawData "[����]" & SQL
        
        '-- Record Count ������
        cn_Ser_Bit.CursorLocation = adUseClient
        Set RS = cn_Ser_Bit.Execute(SQL, , 1)
        If Not RS.EOF = True And Not RS.BOF = True Then
            Do Until RS.EOF
                With vasList
                    .ReDraw = False
                    .MaxRows = .MaxRows + 1
                    SetText vasList, "1", .MaxRows, colCHECKBOX
                    SetText vasList, "����", .MaxRows, 3
                    strRcpDate = Mid(RS.Fields("��������"), 1, 4) & "-" & Mid(RS.Fields("��������"), 6, 2) & "-" & Mid(RS.Fields("��������"), 9, 2)
                    SetText vasList, strRcpDate, .MaxRows, 10
                    SetText vasList, Trim(RS.Fields("��Ʈ��ȣ")) & "", .MaxRows, 4
                    SetText vasList, Trim(RS.Fields("�̸�")) & "", .MaxRows, 6
                End With
                
                RS.MoveNext
            Loop
            chkAll.Value = "1"
            
            Call txtSNo_KeyPress(13)
            
        Else
            chkAll.Value = "0"
        End If
    
        RS.Close
    '-- ����
    ElseIf cboChk.ListIndex = 2 Then
              SQL = "SELECT distinct A.SPRT_DATE as ��������, A.SPRT_CHAM_ID as ��Ʈ��ȣ, A.SPRT_TIME, P.CHAM_WHANJA as �̸�" & vbCr
        SQL = SQL & " FROM Month..SPRT" & strYM & " A" & vbCr
        SQL = SQL & "    , Month..OENT" & strYM & " B" & vbCr
        SQL = SQL & "    , Neosoft..CC_CHAM P " & vbCr
    '    SQL = SQL & " WHERE A.SPRT_DATE = '" & strDay & "'" & vbCr
        SQL = SQL & " WHERE A.SPRT_DATE BETWEEN '" & strDay1 & "' AND '" & strDay2 & "'" & vbCr
'        SQL = SQL & "   AND A.SPRT_FLAG = 0 " & vbCr
    '    SQL = SQL & "   AND A.SPRT_DEPT_ID = 'A1' " & vbCr
        SQL = SQL & "   AND B.OENT_MOMR_ID IN (" & gAllExam & ")" & vbCr
        SQL = SQL & "   AND A.SPRT_DATE    = B.OENT_DATE      " & vbCr
        SQL = SQL & "   AND A.SPRT_CHAM_ID = B.OENT_CHAM_ID   " & vbCr
        SQL = SQL & "   AND A.SPRT_CHAM_ID = P.CHAM_ID        " & vbCr
        SQL = SQL & " ORDER BY A.SPRT_DATE, A.SPRT_TIME  "
        
        SetRawData "[����]" & SQL
        
        '-- Record Count ������
        cn_Ser.CursorLocation = adUseClient
        Set RS = cn_Ser.Execute(SQL, , 1)
        If Not RS.EOF = True And Not RS.BOF = True Then
            Do Until RS.EOF
                With vasList
                    .ReDraw = False
                    .MaxRows = .MaxRows + 1
                    SetText vasList, "1", .MaxRows, colCHECKBOX
                    SetText vasList, "����", .MaxRows, 3
                    strRcpDate = Mid(RS.Fields("��������"), 1, 4) & "-" & Mid(RS.Fields("��������"), 6, 2) & "-" & Mid(RS.Fields("��������"), 9, 2)
                    SetText vasList, strRcpDate, .MaxRows, 10
                    SetText vasList, Trim(RS.Fields("��Ʈ��ȣ")) & "", .MaxRows, 4
                    SetText vasList, Trim(RS.Fields("�̸�")) & "", .MaxRows, 6
                End With
                
                RS.MoveNext
            Loop
            chkAll.Value = "1"
            
            Call txtSNo_KeyPress(13)
            
        Else
            chkAll.Value = "0"
        End If
    
        RS.Close
    Else
        '-- ���� + ����
        '-- ����
              SQL = "SELECT distinct b.cdate as ��������, a.pno as ��Ʈ��ȣ, a.pname as �̸�" & vbCr
        SQL = SQL & " FROM mt_patinet a " & vbCr
        SQL = SQL & "    , his_chart02 b " & vbCr
        SQL = SQL & "    , his_chart03 c " & vbCr
        SQL = SQL & " WHERE b.cdate  BETWEEN '" & sSch1 & "' AND '" & sSch2 & "'" & vbCr
        SQL = SQL & "   AND a.pno = b.cpno  " & vbCr
        SQL = SQL & "   AND a.pno = c.cpno  " & vbCr
        SQL = SQL & " ORDER BY b.cdate, a.pno  "
        
        SetRawData "[����0]" & SQL
        
        '-- Record Count ������
        cn_Ser_Bit.CursorLocation = adUseClient
        Set RS = cn_Ser_Bit.Execute(SQL, , 1)
        If Not RS.EOF = True And Not RS.BOF = True Then
            Do Until RS.EOF
                With vasList
                    .ReDraw = False
                    .MaxRows = .MaxRows + 1
                    SetText vasList, "1", .MaxRows, colCHECKBOX
                    SetText vasList, "����", .MaxRows, 3
                    strRcpDate = Mid(RS.Fields("��������"), 1, 4) & "-" & Mid(RS.Fields("��������"), 6, 2) & "-" & Mid(RS.Fields("��������"), 9, 2)
                    SetText vasList, strRcpDate, .MaxRows, 10
                    SetText vasList, Trim(RS.Fields("��Ʈ��ȣ")) & "", .MaxRows, 4
                    SetText vasList, Trim(RS.Fields("�̸�")) & "", .MaxRows, 6
                End With
                
                RS.MoveNext
            Loop
            'chkAll.Value = "1"
            
            Call txtSNo_KeyPress(13)
            
        Else
            chkAll.Value = "0"
        End If
    
        RS.Close
        
        '-- ����
              SQL = "SELECT distinct A.SPRT_DATE as ��������, A.SPRT_CHAM_ID as ��Ʈ��ȣ, A.SPRT_TIME, P.CHAM_WHANJA as �̸�" & vbCr
        SQL = SQL & " FROM Month..SPRT" & strYM & " A" & vbCr
        SQL = SQL & "    , Month..OENT" & strYM & " B" & vbCr
        SQL = SQL & "    , Neosoft..CC_CHAM P " & vbCr
    '    SQL = SQL & " WHERE A.SPRT_DATE = '" & strDay & "'" & vbCr
        SQL = SQL & " WHERE A.SPRT_DATE BETWEEN '" & strDay1 & "' AND '" & strDay2 & "'" & vbCr
'        SQL = SQL & "   AND A.SPRT_FLAG = 0 " & vbCr
    '    SQL = SQL & "   AND A.SPRT_DEPT_ID = 'A1' " & vbCr
        SQL = SQL & "   AND B.OENT_MOMR_ID IN (" & gAllExam & ")" & vbCr
        SQL = SQL & "   AND A.SPRT_DATE    = B.OENT_DATE      " & vbCr
        SQL = SQL & "   AND A.SPRT_CHAM_ID = B.OENT_CHAM_ID   " & vbCr
        SQL = SQL & "   AND A.SPRT_CHAM_ID = P.CHAM_ID        " & vbCr
        SQL = SQL & " ORDER BY A.SPRT_DATE, A.SPRT_TIME  "
        
        SetRawData "[����0]" & SQL
        
        '-- Record Count ������
        cn_Ser.CursorLocation = adUseClient
        Set RS = cn_Ser.Execute(SQL, , 1)
        If Not RS.EOF = True And Not RS.BOF = True Then
            Do Until RS.EOF
                With vasList
                    .ReDraw = False
                    .MaxRows = .MaxRows + 1
                    SetText vasList, "1", .MaxRows, colCHECKBOX
                    SetText vasList, "����", .MaxRows, 3
                    SetText vasList, Format(strYM & Trim(RS.Fields("��������")) & "", "####-##-##"), .MaxRows, 10
                    SetText vasList, Trim(RS.Fields("��Ʈ��ȣ")) & "", .MaxRows, 4
                    SetText vasList, Trim(RS.Fields("�̸�")) & "", .MaxRows, 6
                End With
                
                RS.MoveNext
            Loop
            'chkAll.Value = "1"
            
            Call txtSNo_KeyPress(13)
            
        Else
            chkAll.Value = "0"
        End If
    
        RS.Close
        
    End If
    
    vasList.MaxRows = vasList.DataRowCnt
    vasList.RowHeight(-1) = 13.3
    
    
End Sub

Private Sub cmdUp_Click()
    Dim lRow As Long
    
    lRow = vasList.ActiveRow
    
    vasList.SwapRange 1, lRow, vasList.MaxCols, lRow, 1, lRow - 1
    vasActiveCell vasList, lRow - 1, 2
    vasList_Click 2, lRow - 1
End Sub

Private Sub cmdWorkList_Click()
    Dim lRow As Long
    Dim lCol As Long
    Dim lDestRow As Long
    
    frmInterface.vasID.MaxRows = 0
    
'    lDestRow = frmInterface.vasID.DataRowCnt + 1
'
'    If frmInterface.vasID.MaxRows < lDestRow Then
'        frmInterface.vasID.MaxRows = lDestRow
'    End If
    
    For lRow = 1 To vasList.DataRowCnt
        vasList.Row = lRow
        vasList.Col = 1
        
        lDestRow = frmInterface.vasID.DataRowCnt + 1
    
        If frmInterface.vasID.MaxRows < lDestRow Then
            frmInterface.vasID.MaxRows = lDestRow
        End If
        
        If vasList.Value = 1 And Trim(GetText(vasList, lRow, 4)) <> "" Then
            SetText frmInterface.vasID, Trim(GetText(vasList, lRow, 2)), lDestRow, colDISKNO    '1
            SetText frmInterface.vasID, Trim(GetText(vasList, lRow, 4)), lDestRow, colBARCODE   '107554
            SetText frmInterface.vasID, Trim(GetText(vasList, lRow, 5)), lDestRow, colCHARTNO   'o26826
            SetText frmInterface.vasID, Trim(GetText(vasList, lRow, 6)), lDestRow, colPNAME     '���ο�
            SetText frmInterface.vasID, Trim(GetText(vasList, lRow, 7)), lDestRow, colPSEX      'M
            SetText frmInterface.vasID, Trim(GetText(vasList, lRow, 8)), lDestRow, colPAGE      '58
            SetText frmInterface.vasID, Trim(GetText(vasList, lRow, 10)), lDestRow, colHOSPDATE '20160410
            
            SetText frmInterface.vasID, Trim(GetText(vasList, lRow, 10)), lDestRow, colHOSPDATE '20160410
            
            SetText frmInterface.vasID, Trim(GetText(vasList, lRow, 9)), lDestRow, colPOSNO     '5911271006538 �ֹι�ȣ ��� colPOSNO
            
            SetText frmInterface.vasID, Trim(GetText(vasList, lRow, 3)), lDestRow, colINOUT '�Կ�
            
            lDestRow = lDestRow + 1
        End If
    Next lRow
    
    frmInterface.vasID.RowHeight(-1) = 12
    Unload Me
    
End Sub

'Private Sub Command1_Click()
'    Dim lRow As Long
'
'    lRow = vasList.ActiveRow
'
'    If lRow = 1 Then Exit Sub
'
'    lRow = lRow - 1
'
'    vasActiveCell vasList, lRow, 5
'
'    vasList_DblClick 5, lRow
'
'End Sub

'Private Sub Command2_Click()
'    Dim lRow As Long
'
'    lRow = vasList.ActiveRow
'
'    If lRow = vasList.DataRowCnt Then Exit Sub
'
'    lRow = lRow + 1
'
'    vasActiveCell vasList, lRow, 5
'
'    vasList_DblClick 5, lRow
'End Sub

Private Sub Form_Activate()
    'dtpSDate.SetFocus
   ' vasActiveCell vasList, 1, 2
End Sub

Private Sub Form_Load()

    dtpSDate.Value = Date
    dtpEDate.Value = Date
    
    ClearSpread vasList
    
    chkAll.Value = 0
    
    cboChk.ListIndex = 0
    
    cmdSearch_Click
        
End Sub

'Private Sub monvCal_DateClick(ByVal DateClicked As Date)
'    If iIndex = 0 Then
'        dtpSDate.Text = Trim(Format(DateClicked, "yyyy-mm-dd"))
'    Else
'        dtpEDate.Text = Trim(Format(DateClicked, "yyyy-mm-dd"))
'    End If
'    monvCal.Visible = False
'End Sub

'Private Sub Text1_Change()
'
'End Sub
'
'Private Sub txtBarCode_GotFocus()
'    SelectFocus txtBarCode
'End Sub
'
'Private Sub txtBarCode_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'        If Len(txtBarCode) <> 10 Then
'            txtBarCode.SetFocus
'            Exit Sub
'        End If
'        btnSch_Click
'        txtBarCode = ""
'    End If
'End Sub



Private Sub txtSNo_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    
    If KeyAscii = 13 Then
        With vasList
            For i = .ActiveRow To .MaxRows
                .Row = i
                .Col = colSAVESEQ
                .Text = txtSNo.Text
                txtSNo.Text = txtSNo.Text + 1
'                If txtSNo.Text = "31" Then
'                    txtSNo.Text = "1"
'                End If
            Next
        End With
    End If
End Sub

Private Sub vasList_Click(ByVal Col As Long, ByVal Row As Long)
    If sspOrder.Visible = True Then sspOrder.Visible = False

    If Row = 0 Then
        vasSort vasList, Col
    End If

    If Row < 0 Or Row > vasList.DataRowCnt Then
        cmdUp.Enabled = False
        cmdDown.Enabled = False
    End If

    If Row = 1 Then
        cmdUp.Enabled = False
        cmdDown.Enabled = True
    ElseIf Row = vasList.DataRowCnt Then
        cmdUp.Enabled = True
        cmdDown.Enabled = False
    Else
        cmdUp.Enabled = True
        cmdDown.Enabled = True
    End If
End Sub

'Private Sub vasList_DblClick(ByVal Col As Long, ByVal Row As Long)
''    Dim lRow, lCol As Long
''    Dim lDestRow As Long
''
''    lDestRow = Form_Main.vasExam.DataRowCnt + 1
''
''    lRow = vasList.ActiveRow
''
''    For lCol = 2 To 8
''        If lCol = 8 Then        'ó������
''            SetText Form_Main.vasExam, Trim(GetText(vasList, lRow, 8)), lDestRow, 12
''        ElseIf lCol = 2 Then    '��ü��ȣ
''            SetText Form_Main.vasExam, Trim(GetText(vasList, lRow, 2)), lDestRow, 2
''        Else
''            SetText Form_Main.vasExam, Trim(GetText(vasList, lRow, lCol)), lDestRow, lCol + 3
''        End If
''    Next lCol
'
''    Unload Me
'
''===================================================================
''2004/08/03 �̻��� - ȯ�� ����Ŭ���� �� �˻��׸� ���÷��� �ǵ���
'Dim sCnt As String
'Dim sExamCode As String
'Dim sEquipCode As String
'
'Dim iRow As Integer
'Dim jRow As Integer
'
'    txtDate = GetText(vasList, Row, 9)
'
'    txtNo = Trim(GetText(vasList, Row, 10))
'    txtPID = Trim(GetText(vasList, Row, 4))
'    txtName = Trim(GetText(vasList, Row, 5))
'
'    txtSex = Trim(GetText(vasList, Row, 6))
'    txtAge = Trim(GetText(vasList, Row, 7))
'
'    chkAllOrder.Value = 0
'
'    ClearSpread vasOrder
'
'    '�˻��ڵ� ��������
'
'    SQL = "Select '',RstOdrCod,'' "
'    SQL = SQL & vbCrLf & " from Rstinf "
'    SQL = SQL & vbCrLf & " where RstLabNum = '" & txtNo & "' "
'    SQL = SQL & vbCrLf & "   and RstOdrCod In (" & gAllExam & ") "
'
'    Res = db_select_Vas(gServer, SQL, vasOrder)
''    vasSort vasOrder, 2
'    If Res = -1 Then
'        SaveQuery SQL
'        Exit Sub
'    End If
'
'    vasOrder.MaxRows = vasOrder.DataRowCnt
'
'    For jRow = 1 To vasOrder.DataRowCnt
'        SQL = " select ExamName from EquipExam " & vbCrLf & _
'              " where equipno = '" & gEquip & "' and ExamCode = '" & Trim(GetText(vasOrder, jRow, 2)) & "' "
'        Res = db_select_Col(gLocal, SQL)
'
'        If Res = 1 Then
'            SetText vasOrder, Trim(gReadBuf(0)), jRow, 3
'        End If
'    Next jRow
'
'    sspOrder.Visible = True
'
'End Sub

Private Sub vasList_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim iRow As Integer
    
    iRow = vasList.ActiveRow
    
    If KeyCode = vbKeyDelete Then
        If iRow < 1 Or iRow > vasList.DataRowCnt Then Exit Sub
        DeleteRow vasList, iRow, iRow
    End If
End Sub

