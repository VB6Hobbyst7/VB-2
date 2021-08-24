VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmPatSearch 
   BorderStyle     =   1  '단일 고정
   Caption         =   "검사자 조회"
   ClientHeight    =   8790
   ClientLeft      =   7440
   ClientTop       =   2250
   ClientWidth     =   10170
   Icon            =   "frmPatSearch.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   10170
   StartUpPosition =   2  '화면 가운데
   Begin VB.Frame fraWork 
      Height          =   1215
      Left            =   90
      TabIndex        =   26
      Top             =   90
      Width           =   10005
      Begin VB.TextBox txtQCBar 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1290
         TabIndex        =   42
         Top             =   720
         Width           =   2925
      End
      Begin InetCtlsObjects.Inet Inet1 
         Left            =   9420
         Top             =   660
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "출력"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   11340
         Style           =   1  '그래픽
         TabIndex        =   40
         Top             =   150
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.CommandButton cmdWorkList 
         Caption         =   "오더 전송"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   7530
         Style           =   1  '그래픽
         TabIndex        =   39
         Top             =   180
         Width           =   1140
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "닫기"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   8685
         Style           =   1  '그래픽
         TabIndex        =   38
         Top             =   180
         Width           =   1140
      End
      Begin VB.TextBox txtSNo 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   18
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   13110
         TabIndex        =   36
         Text            =   "1"
         Top             =   180
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.CommandButton cmdDown 
         BackColor       =   &H00FFFFFF&
         Caption         =   "▼"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   6930
         Style           =   1  '그래픽
         TabIndex        =   35
         Top             =   210
         Width           =   555
      End
      Begin VB.CommandButton cmdUp 
         BackColor       =   &H00FFFFFF&
         Caption         =   "▲"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   6360
         Style           =   1  '그래픽
         TabIndex        =   34
         Top             =   210
         Width           =   555
      End
      Begin VB.OptionButton optState 
         Caption         =   "일반"
         Height          =   180
         Index           =   0
         Left            =   4290
         TabIndex        =   33
         Top             =   240
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.OptionButton optState 
         Caption         =   "QC"
         Height          =   180
         Index           =   1
         Left            =   4290
         TabIndex        =   32
         Top             =   480
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "워크조회"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   5100
         TabIndex        =   27
         Top             =   180
         Width           =   1200
      End
      Begin MSComCtl2.DTPicker dtpEDate 
         Height          =   345
         Left            =   2820
         TabIndex        =   28
         Top             =   240
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   63307777
         CurrentDate     =   40248
      End
      Begin MSComCtl2.DTPicker dtpSDate 
         Height          =   345
         Left            =   1230
         TabIndex        =   29
         Top             =   240
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   63307777
         CurrentDate     =   40248
      End
      Begin VB.Label Label2 
         Caption         =   "QC바코드"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   210
         TabIndex        =   41
         Top             =   810
         Width           =   1005
      End
      Begin VB.Label Label1 
         Caption         =   "시작번호"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   405
         Left            =   12600
         TabIndex        =   37
         Top             =   240
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label Label13 
         Caption         =   "처방일자"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   31
         Top             =   300
         Width           =   915
      End
      Begin VB.Label Label9 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2670
         TabIndex        =   30
         Top             =   330
         Width           =   105
      End
   End
   Begin VB.Frame sspOrder 
      Caption         =   "Frame1"
      Height          =   3855
      Left            =   15390
      TabIndex        =   6
      Top             =   4020
      Visible         =   0   'False
      Width           =   7755
      Begin VB.CheckBox chkAllOrder 
         Caption         =   "Check1"
         Height          =   345
         Left            =   3240
         TabIndex        =   21
         Top             =   180
         Width           =   225
      End
      Begin VB.TextBox txtNo 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1050
         TabIndex        =   14
         Top             =   90
         Width           =   1395
      End
      Begin VB.TextBox txtPID 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1050
         TabIndex        =   13
         Top             =   510
         Width           =   1395
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
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
         Top             =   930
         Width           =   1395
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "확인"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   90
         TabIndex        =   11
         Top             =   3060
         Width           =   1215
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "닫기"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1380
         TabIndex        =   10
         Top             =   3060
         Width           =   1215
      End
      Begin VB.TextBox txtSex 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1050
         TabIndex        =   9
         Top             =   1350
         Width           =   915
      End
      Begin VB.TextBox txtAge 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1050
         TabIndex        =   8
         Top             =   1770
         Width           =   915
      End
      Begin VB.TextBox txtDate 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   60
         TabIndex        =   7
         Top             =   2700
         Visible         =   0   'False
         Width           =   1965
      End
      Begin FPSpread.vaSpread vasOrder 
         Height          =   3555
         Left            =   2760
         TabIndex        =   15
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
         SpreadDesigner  =   "frmPatSearch.frx":014A
      End
      Begin VB.Label Label7 
         BackStyle       =   0  '투명
         Caption         =   "검체번호"
         BeginProperty Font 
            Name            =   "굴림체"
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
         Top             =   150
         Width           =   1005
      End
      Begin VB.Label Label6 
         BackStyle       =   0  '투명
         Caption         =   "환자번호"
         BeginProperty Font 
            Name            =   "굴림체"
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
         Top             =   570
         Width           =   1005
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '투명
         Caption         =   "환자이름"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   18
         Top             =   990
         Width           =   1005
      End
      Begin VB.Label Label4 
         BackStyle       =   0  '투명
         Caption         =   "성별"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   17
         Top             =   1410
         Width           =   1005
      End
      Begin VB.Label Label5 
         BackStyle       =   0  '투명
         Caption         =   "나이"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   16
         Top             =   1830
         Width           =   1005
      End
   End
   Begin VB.CheckBox chkAll 
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   630
      TabIndex        =   3
      Top             =   1485
      Width           =   165
   End
   Begin VB.CommandButton cmdOrder 
      Caption         =   "Order 전송"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   11310
      Style           =   1  '그래픽
      TabIndex        =   2
      Top             =   8790
      Visible         =   0   'False
      Width           =   1320
   End
   Begin FPSpread.vaSpread vasPrint 
      Height          =   2610
      Left            =   14700
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   5280
      _Version        =   393216
      _ExtentX        =   9313
      _ExtentY        =   4604
      _StockProps     =   64
      ColHeaderDisplay=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
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
      SpreadDesigner  =   "frmPatSearch.frx":120B
   End
   Begin FPSpread.vaSpread vasCode 
      Height          =   3645
      Left            =   15120
      TabIndex        =   1
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
      SpreadDesigner  =   "frmPatSearch.frx":2449
   End
   Begin FPSpread.vaSpread vasList 
      Height          =   7275
      Left            =   90
      TabIndex        =   4
      Top             =   1410
      Width           =   10005
      _Version        =   393216
      _ExtentX        =   17648
      _ExtentY        =   12832
      _StockProps     =   64
      ColHeaderDisplay=   0
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      MaxCols         =   16
      MaxRows         =   100
      ScrollBars      =   2
      ShadowColor     =   15987699
      ShadowDark      =   13815180
      SpreadDesigner  =   "frmPatSearch.frx":26D1
   End
   Begin MSComCtl2.DTPicker dtpStopDt 
      Height          =   345
      Left            =   14280
      TabIndex        =   22
      Top             =   9750
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   63307777
      CurrentDate     =   40248
   End
   Begin MSComCtl2.DTPicker dtpStartDt 
      Height          =   345
      Left            =   12750
      TabIndex        =   23
      Top             =   9750
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   63307777
      CurrentDate     =   40248
   End
   Begin VB.Label Label20 
      Caption         =   "조회일자"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   11850
      TabIndex        =   25
      Top             =   9810
      Width           =   915
   End
   Begin VB.Label Label12 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "굴림"
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
      TabIndex        =   24
      Top             =   9840
      Width           =   105
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '투명
      Caption         =   "결과완료 : 빨간색, 미완료 : 검정색"
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   6240
      TabIndex        =   5
      Top             =   8820
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
'    '체크, Rack, Pos, SampleNo, 환자번호, 환자이름, 성별, 나이, 주민번호, 접수일자
'    '검사상태
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
''Local에 환자에 대한 검사항목 저장하기
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
'        MsgBox "Rack, Pos을 확인하세요!", vbCritical, "알림"
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
'            '처방가져오기
'            sOrderCode = ""
'
'            vasList.SetText 3, iRow, txtPos
'
'            txtPos = CStr(CInt(txtPos) + 1)
'
'            ClearSpread vasCode
'
'            sID = Trim(GetText(vasList, iRow, 10))     '검체번호
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
'    'WorkList 전송
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
            SetText vasPrint, Trim(GetText(vasList, iRow, 11)), j, 1     '검체번호
            SetText vasPrint, Trim(GetText(vasList, iRow, 4)), j, 2     '환자번호
            SetText vasPrint, Trim(GetText(vasList, iRow, 5)), j, 3     '환자이름

            SetText vasPrint, Trim(GetText(vasList, iRow, 6)), j, 4     '성별
            SetText vasPrint, Trim(GetText(vasList, iRow, 7)), j, 5     '나이
            'SetText vasPrint, Trim(GetText(vasList, iRow, 7)), j, 6     '주민번호
            SetText vasPrint, Trim(GetText(vasList, iRow, 9)), j, 7     '처방일자
            SetText vasPrint, Trim(GetText(vasList, iRow, 12)), j, 8     '처방일자
            
            j = j + 1
        End If
    Next iRow
    
    If vasPrint.DataRowCnt < 1 Then
        MsgBox "출력할 자료가 없습니다.", , "알 림"
        Exit Sub
    End If
    
    sCurDate = GetDateFull
    
    sSerDate = Trim(dtpSDate.Value) & " - " & Trim(dtpEDate.Value)
    
    '2004/08/11 이상은 - 세로방향에서 가로방향으로 수정
    vasPrint.PrintOrientation = 1   ' SS_PRINTORIENT_PORTRAIT
    vasPrint.PrintAbortMsg = "인쇄중 입니다 ..."
    vasPrint.PrintJobName = "WorkList 출력"
    

    sHead = "/fn""궁서체"" /fz""12"" /fb1 /fi0 /fu0 " & "/c" & "▣ WorkList ▣" & "/n/n " & _
            "/fn""굴림체"" /fz""10"" /fb0 /fi0 /fu0 " & "/c" & "처방일자 : " & dtpSDate & " ~ " & dtpEDate
    'If optGubun(0).Value = True Then
    '    sHead = sHead & " (진료)" & "/n/n"
    'ElseIf optGubun(1).Value = True Then
    '    sHead = sHead & " (검진)" & "/n/n"
    'End If

    sFoot = "/fn""굴림체"" /fz""10"" /fb1 /fi0 /fu0 " & "/l" & sCurDate & "/fn""궁서체"" /fz""11"" /fb1 /fi0 /fu0 /r" & " 검사실"
    
    vasPrint.PrintHeader = sHead
    vasPrint.PrintFooter = sFoot

    vasPrint.PrintMarginTop = 680
    vasPrint.PrintMarginBottom = 680
'현재 SS가 비대칭으로 출력함
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
'    '체크, Rack, Pos, SampleNo, 환자번호, 환자이름, 성별, 나이, 주민번호, 접수일자
'    '검사상태
'    sSch1 = Format(dtpSDate.Text, "yyyy-mm-dd")
'    sSch2 = Format(dtpEDate.Text, "yyyy-mm-dd")
'
'    SQL = " Select max(a.DR_CHART), b.PE_SUJINJA, '', '', b.PE_JUMIN, a.DR_DATE, '', '' " & vbCrLf & _
'          " From DEPARTDAT a, PERSON b " & vbCrLf & _
'          " Where a.DR_DATE between '" & sSch1 & "' and '" & sSch2 & "' " & vbCrLf & _
'          " And a.DR_CODE in (" & gAllExam & ") " & vbCrLf & _
'          " And a.DR_CHART = b.PE_CHART "
'
''    If optState(0).Value = True Then        '접수
''        SQL = SQL & vbCrLf & _
''              " And c.GD_RESULT = ''  "
''    ElseIf optState(1).Value = True Then    '결과
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
    Dim sParam As String
    Dim sRcvData, sData As String
    Dim varRcvData As Variant
    Dim varTstCode As Variant
    Dim i As Integer
    Dim strTstCD As String
    Dim strItems As String
    Dim intRow As Integer
    Dim strTestCds As String
    
On Error GoTo ErrorTrap
    
    sSch1 = Format(dtpSDate.Value, "yyyymmdd")
    sSch2 = Format(dtpEDate.Value, "yyyymmdd")
    
    ClearSpread vasList
    vasList.MaxRows = 0
    
    'strTestCds = "LIM305▦LIM306▦"
    'strTestCds = "LIM305"
    
    
    If optState(0).Value = True Then
        'sParam = "submit_id=TRLII00119&"                                           'submit ID
        sParam = "submit_id=TRLII00101&"                                            'submit ID
        sParam = sParam & "business_id=lis&"                                        'business_id
        sParam = sParam & "ex_interface=" & NUAPI.UID & "|" & NUAPI.HOSPCD & "&"    '사용자ID|기관코드
        sParam = sParam & "instcd=" & NUAPI.HOSPCD & "&"                            '기관코드
        sParam = sParam & "eqmtcd=" & NUAPI.INSTCD & "&"                            '장비코드
        sParam = sParam & "startdd=" & sSch1 & "&"                                  '시작작업일자
        sParam = sParam & "enddd=" & sSch2 & "&"                                    '종료작업일자
    Else
        sParam = "submit_id=TRLQI00101&"                                            'submit ID
        sParam = sParam & "business_id=lis&"                                        'business_id
        sParam = sParam & "ex_interface=" & NUAPI.UID & "|" & NUAPI.HOSPCD & "&"    '사용자ID|기관코드
        sParam = sParam & "instcd=" & NUAPI.HOSPCD & "&"                            '기관코드
        sParam = sParam & "eqmtcd=" & NUAPI.INSTCD & "&"                            '장비코드
        sParam = sParam & "startdd=" & sSch1 & "&"                                  '시작작업일자
        sParam = sParam & "enddd=" & sSch2 & "&"                                    '종료작업일자
    End If
    
    '==> 서버로 오더조회
    'SetRawData "[WL_IN]" & sParam
        'spcacptdt 접수일자
        'acptflag 입원외래구분
        'bcno 검체번호
        'PID 등록번호
        'patnm 환자명
        'sexage 나이성별
        'erprcpflag 응급구분
        'workno 작업번호
        'tsectnm 검사계명
        'ifreqcdlist 장비요청코드
        'tclscdlist 검사리스트
        'urinextrvol 유린값
        'retestyn 재검여부
        'rsltstat 결과상태
    sRcvData = OpenURLWithIE2(NUAPI.APIURL & sParam, Inet1)
    
    Call SetSQLData("워크조회", NUAPI.APIURL & sParam & vbNewLine & sRcvData)
    
    If InStr(1, sRcvData, "<?xml version") > 0 Then
        varRcvData = Split(sRcvData, "CDATA[")
    End If
    
    If UBound(varRcvData) >= 0 Then
        For i = 1 To UBound(varRcvData)
            varRcvData(i) = Mid(varRcvData(i), 1, InStr(varRcvData(i), "]") - 1)
        Next
        
        For i = 1 To UBound(varRcvData) Step 14
            With vasList
                .MaxRows = .MaxRows + 1
                
            
                intRow = .MaxRows
                .Row = intRow
                '.Col = 7
                '.BackColor = vbGreen '&HC6FEFF '&H80C0FF
                                                
                .SetText 1, intRow, "1"
                .SetText 2, intRow, Format(Mid(varRcvData(i), 1, 8), "####-##-##")
                .SetText 3, intRow, varRcvData(i + 1) & ""
                .SetText 4, intRow, varRcvData(i + 2) & ""
                .SetText 5, intRow, varRcvData(i + 3) & ""
                .SetText 6, intRow, varRcvData(i + 4) & ""
                .SetText 7, intRow, mGetP(varRcvData(i + 5) & "", 1, "/")
                .SetText 8, intRow, mGetP(varRcvData(i + 5) & "", 2, "/")
                .SetText 9, intRow, varRcvData(i + 6) & ""
                .SetText 10, intRow, varRcvData(i + 7) & ""
                .SetText 11, intRow, varRcvData(i + 8) & ""
                
                strTestCds = varRcvData(i + 9) & ""
                strTestCds = Replace(strTestCds, "▦", "")
                
                If InStr(varRcvData(i + 10) & "", "LIM305") > 0 Then
                    .SetText 14, intRow, "Inhalant"
                ElseIf InStr(varRcvData(i + 10) & "", "LIM306") > 0 Then
                    .SetText 14, intRow, "Food"
                End If
                .RowHeight(-1) = 12
            End With
        Next
    End If
    
    chkAll.Value = "1"
                 
    'vasList.MaxRows = vasList.DataRowCnt
    vasList.RowHeight(-1) = 13.3

    Exit Sub
    
ErrorTrap:
    
    MsgBox "조회 오류", vbOKOnly + vbCritical, Me.Caption


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
            SetText frmInterface.vasID, "1", lDestRow, colChECKBOX
            SetText frmInterface.vasID, Trim(GetText(vasList, lRow, 2)), lDestRow, colHOSPDATE
            SetText frmInterface.vasID, Trim(GetText(vasList, lRow, 3)), lDestRow, colIO
            SetText frmInterface.vasID, Trim(GetText(vasList, lRow, 4)), lDestRow, colBARCODE
            SetText frmInterface.vasID, Trim(GetText(vasList, lRow, 5)), lDestRow, colPID
            SetText frmInterface.vasID, Trim(GetText(vasList, lRow, 6)), lDestRow, colPNAME
            SetText frmInterface.vasID, Trim(GetText(vasList, lRow, 7)), lDestRow, colPSEX
            SetText frmInterface.vasID, Trim(GetText(vasList, lRow, 8)), lDestRow, colPAGE
            SetText frmInterface.vasID, Trim(GetText(vasList, lRow, 9)), lDestRow, colER
            SetText frmInterface.vasID, Trim(GetText(vasList, lRow, 10)), lDestRow, colWORKNO
            SetText frmInterface.vasID, Trim(GetText(vasList, lRow, 11)), lDestRow, colPARTNM
            'SetText frmInterface.vasID, Trim(GetText(vasList, lRow, 13)), lDestRow, colASSAYNM
            SetText frmInterface.vasID, Trim(GetText(vasList, lRow, 14)), lDestRow, colASSAYNM
            
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
    vasActiveCell vasList, 1, 2
End Sub

Private Sub Form_Load()

    dtpSDate.Value = Date
    dtpEDate.Value = Date
    
    ClearSpread vasList
    
    chkAll.Value = 0
    
    txtQCBar.Text = ""
    
    Call cmdSearch_Click
        
End Sub

Private Sub txtQCBar_KeyPress(KeyAscii As Integer)
    Dim sParam      As String
    Dim sRcvData    As String
    Dim varRcvData  As Variant
    Dim i           As Integer
    Dim intRow      As Integer
    Dim strTestCds  As String
    
    
    
    If KeyAscii = vbKeyReturn Then
                 sParam = "submit_id=TRLQI00101&"                                   'submit ID
        sParam = sParam & "business_id=lis&"                                        'business_id
        sParam = sParam & "ex_interface=" & NUAPI.UID & "|" & NUAPI.HOSPCD & "&"    '사용자ID|기관코드
        sParam = sParam & "bcno=" & Trim(txtQCBar.Text) & "&"                       'QC바코드
        sParam = sParam & "eqmtcd=" & NUAPI.INSTCD & "&"                            '장비코드
        sParam = sParam & "instcd=" & NUAPI.HOSPCD & "&"                            '기관코드
        sParam = sParam & "userid=" & NUAPI.UID                                     '사용자ID
        
        sRcvData = OpenURLWithIE2(NUAPI.APIURL & sParam, Inet1)
        
        Call SetSQLData("QC조회", NUAPI.APIURL & sParam & vbNewLine & sRcvData)
        
        If InStr(1, sRcvData, "<?xml version") > 0 Then
            varRcvData = Split(sRcvData, "CDATA[")
        End If
        
'        <acptdt><![CDATA[20170223092042]]></acptdt>
'        <bcno><![CDATA[Q560801F0]]></bcno>
'        <testcd><![CDATA[LIM305▦LIM30572▦LIM30571▦LIM30570▦LIM30569▦LIM30568▦LIM30567▦LIM30566▦LIM30565▦LIM30564▦LIM30563▦LIM30562▦LIM30561▦LIM30560▦LIM30558▦LIM30557▦LIM30555▦LIM30554▦LIM30553▦LIM30552▦LIM30551▦LIM30550▦LIM30549▦LIM30548▦LIM30547▦LIM30545▦LIM30544▦LIM30543▦LIM30542▦LIM30541▦LIM30540▦LIM30539▦LIM30538▦LIM30535▦LIM30531▦LIM30530▦LIM30529▦LIM30528▦LIM30526▦LIM30525▦LIM30524▦LIM30523▦LIM30522▦LIM30521▦LIM30520▦LIM30519▦LIM30518▦LIM30517▦LIM30516▦LIM30514▦LIM30513▦LIM30512▦LIM30511▦LIM30508▦LIM30507▦LIM30506▦LIM30505▦LIM30504▦LIM30503▦LIM30502▦LIM30501▦]]></testcd>
'        <testnm><![CDATA[MAST (Inhalant)▦Mango (망고)▦Mussel (홍합)▦Banana (바나나)▦Shellfish (조개)▦Raw chestnut (생밤)▦Pupa (번데기)▦Anchovy (멸치)▦Kiwi (키위)▦Cucumber (오이)▦Lilac pollen (라일락)▦Pigweed (털비름)▦Goldenrod (미역취 국화)▦Russian thistle (명아주과풀)▦Dandelion (민들레)▦Oxeye daisy (불란서 국화)▦Acacia (아카시아)▦Japanese cedar (삼나무)▦Pine (소나무)▦Ash mix (물푸레)▦Poplar mix (포플라)▦Sallow willow (수양버들)▦Sycamore mix (플라타너스)▦Penicillium notatum (곰팡이류)▦Latex (라텍스)▦Honey bee (꿀벌)▦Redtop, bent grass (외겨이삭)▦Reed (갈대)▦Timothy grass (큰조아재비)▦Orchard grass (오리새)▦Bermuda grass (우산잔디)▦Sweet vernal grass (향기풀)▦Hazel (개암나무)▦Sheep (양)▦Japanese hop (환삼덩굴)▦Mugwort (쑥)▦Ragweed, short (돼지풀)▦Oak white (참나무)▦Alder (오리나무)▦Alternaria alternata (곰팡이류)▦Aspergillus fumigatus (곰팡이류)▦Cladosporium herbarum (곰팡이류)▦Cockroach (바퀴벌레)▦House dust (집먼지)▦Rye pollens (호밀풀)▦CCD (Bromelain)▦Mackerel
'                    (고등어)▦Peach (복숭아)▦Cacao (카카오)▦Potato (감자)▦Shrimp (새우)▦Crab (게)▦Soy bean (콩)▦Milk (우유)▦Egg white (계란흰자)▦Dog (개)▦Cat (고양이)▦Storage mite (저장진드기)▦D. farinae (진드기)▦D. pteronyssinus (진드기)▦Total IgE (총 IgE)▦]]></testnm>
'        <matrcd><![CDATA[MAST 60종]]></matrcd>
'        <matrnm><![CDATA[Inhalant 60종]]></matrnm>
'        <levlcd><![CDATA[42]]></levlcd>
        
        
        If UBound(varRcvData) >= 0 Then
            For i = 1 To UBound(varRcvData)
                varRcvData(i) = Mid(varRcvData(i), 1, InStr(varRcvData(i), "]") - 1)
            Next
            
            For i = 1 To UBound(varRcvData) Step 14
                With vasList
                    .MaxRows = .MaxRows + 1
                    intRow = .MaxRows
                    .Row = intRow
                    '.Col = 7
                    '.BackColor = vbGreen '&HC6FEFF '&H80C0FF
                                                    
                    .SetText 1, intRow, "1"
                    .SetText 2, intRow, Format(Mid(varRcvData(i), 1, 8), "####-##-##")
                    '.SetText 3, intRow, varRcvData(i + 1) & ""
                    .SetText 4, intRow, varRcvData(i + 1) & ""
                    '.SetText 5, intRow, varRcvData(i + 2) & ""
                    
                    '.SetText 6, intRow, varRcvData(i + 4) & ""
                    
                    If varRcvData(i + 6) & "" = "41" Then
                        .SetText 6, intRow, "QCINNEG"
                    ElseIf varRcvData(i + 6) & "" = "42" Then
                        .SetText 6, intRow, "QCINPOS"
                    ElseIf varRcvData(i + 6) & "" = "SJ" Then
                        .SetText 6, intRow, "QCFDNEG"
                    ElseIf varRcvData(i + 6) & "" = "SK" Then
                        .SetText 6, intRow, "QCFDPOS"
                    End If

                    '.SetText 7, intRow, mGetP(varRcvData(i + 5) & "", 1, "/")
                    '.SetText 8, intRow, mGetP(varRcvData(i + 5) & "", 2, "/")
                    '.SetText 9, intRow, varRcvData(i + 6) & ""
                    '.SetText 10, intRow, varRcvData(i + 7) & ""
                    '.SetText 11, intRow, varRcvData(i + 8) & ""
                    
                    strTestCds = varRcvData(i + 2) & ""
                    strTestCds = Replace(strTestCds, "▦", "")
                    
                    If InStr(varRcvData(i + 2) & "", "LIM305") > 0 Then
                        .SetText 14, intRow, "Inhalant"
                    ElseIf InStr(varRcvData(i + 2) & "", "LIM306") > 0 Then
                        .SetText 14, intRow, "Food"
                    End If
                    .RowHeight(-1) = 12
                End With
            Next
        End If
    
    End If
    
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
''        If lCol = 8 Then        '처방일자
''            SetText Form_Main.vasExam, Trim(GetText(vasList, lRow, 8)), lDestRow, 12
''        ElseIf lCol = 2 Then    '검체번호
''            SetText Form_Main.vasExam, Trim(GetText(vasList, lRow, 2)), lDestRow, 2
''        Else
''            SetText Form_Main.vasExam, Trim(GetText(vasList, lRow, lCol)), lDestRow, lCol + 3
''        End If
''    Next lCol
'
''    Unload Me
'
''===================================================================
''2004/08/03 이상은 - 환자 더블클릭시 상세 검사항목 디스플레이 되도록
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
'    '검사코드 가져오기
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

