VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmSCLTrans 
   Caption         =   "SCL 결과 연계"
   ClientHeight    =   10560
   ClientLeft      =   -270
   ClientTop       =   345
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSCLTrans.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10560
   ScaleWidth      =   15240
   Begin VB.CommandButton cmdSave1 
      BackColor       =   &H8000000E&
      Caption         =   "실행"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   7590
      Picture         =   "frmSCLTrans.frx":030A
      Style           =   1  '그래픽
      TabIndex        =   41
      Top             =   360
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox txtSheet1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1800
      TabIndex        =   40
      Text            =   "Remark"
      Top             =   2070
      Visible         =   0   'False
      Width           =   5475
   End
   Begin FPSpread.vaSpread vasRemark 
      Height          =   2265
      Left            =   2040
      TabIndex        =   39
      Top             =   3210
      Visible         =   0   'False
      Width           =   5805
      _Version        =   393216
      _ExtentX        =   10239
      _ExtentY        =   3995
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
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "frmSCLTrans.frx":100F
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   10770
      Top             =   270
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "출 력"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   10560
      TabIndex        =   38
      Top             =   900
      Width           =   1245
   End
   Begin FPSpread.vaSpread vasTemp 
      Height          =   2265
      Left            =   2040
      TabIndex        =   7
      Top             =   5520
      Visible         =   0   'False
      Width           =   4425
      _Version        =   393216
      _ExtentX        =   7805
      _ExtentY        =   3995
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   50
      MaxRows         =   3000
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "frmSCLTrans.frx":5503
   End
   Begin VB.TextBox txtTemp 
      Height          =   585
      Left            =   2010
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   36
      Top             =   4800
      Visible         =   0   'False
      Width           =   3885
   End
   Begin VB.TextBox txtPBResult 
      Height          =   765
      Left            =   8070
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   35
      Top             =   5730
      Visible         =   0   'False
      Width           =   5685
   End
   Begin VB.TextBox txtImpResult 
      Height          =   795
      Left            =   8070
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   34
      Top             =   6600
      Visible         =   0   'False
      Width           =   5685
   End
   Begin VB.TextBox txtRec 
      Height          =   795
      Left            =   8040
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   33
      Top             =   7500
      Visible         =   0   'False
      Width           =   5685
   End
   Begin VB.TextBox txtEtc1 
      Height          =   705
      Left            =   2010
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   32
      Top             =   3990
      Visible         =   0   'False
      Width           =   4905
   End
   Begin VB.TextBox txtRemark 
      Height          =   705
      Left            =   2010
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   31
      Top             =   3180
      Visible         =   0   'False
      Width           =   4905
   End
   Begin VB.TextBox txtDiagnosis 
      Height          =   795
      Left            =   8070
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   30
      Top             =   3930
      Visible         =   0   'False
      Width           =   5685
   End
   Begin VB.TextBox txtComment 
      Height          =   795
      Left            =   8070
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   29
      Top             =   4800
      Visible         =   0   'False
      Width           =   5685
   End
   Begin VB.TextBox txtMicro 
      Height          =   795
      Left            =   8070
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   28
      Top             =   3030
      Visible         =   0   'False
      Width           =   5685
   End
   Begin VB.TextBox txtGross 
      Height          =   765
      Left            =   8100
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   27
      Top             =   2130
      Visible         =   0   'False
      Width           =   5685
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   315
      Left            =   120
      TabIndex        =   25
      Top             =   10170
      Width           =   14985
      _ExtentX        =   26432
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
      Max             =   250
   End
   Begin VB.CheckBox chkAll 
      Height          =   255
      Left            =   810
      TabIndex        =   22
      Top             =   1680
      Width           =   165
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H8000000E&
      Caption         =   "취소"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   11910
      Picture         =   "frmSCLTrans.frx":E4F5
      Style           =   1  '그래픽
      TabIndex        =   20
      Top             =   900
      Width           =   1485
   End
   Begin FPSpread.vaSpread vasList 
      Height          =   8265
      Left            =   120
      TabIndex        =   19
      Top             =   1560
      Width           =   15045
      _Version        =   393216
      _ExtentX        =   26538
      _ExtentY        =   14579
      _StockProps     =   64
      ColHeaderDisplay=   0
      EditModePermanent=   -1  'True
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   33
      MaxRows         =   3000
      RetainSelBlock  =   0   'False
      SelectBlockOptions=   0
      SpreadDesigner  =   "frmSCLTrans.frx":F1FA
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H8000000E&
      Caption         =   "실행"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   13470
      Picture         =   "frmSCLTrans.frx":25520
      Style           =   1  '그래픽
      TabIndex        =   18
      Top             =   135
      Width           =   1395
   End
   Begin FPSpread.vaSpread vasRt 
      Height          =   4005
      Left            =   120
      TabIndex        =   17
      Top             =   5280
      Visible         =   0   'False
      Width           =   10215
      _Version        =   393216
      _ExtentX        =   18018
      _ExtentY        =   7064
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
      SpreadDesigner  =   "frmSCLTrans.frx":26225
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   9570
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   5040
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.CommandButton cmdExcute 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   8130
      Picture         =   "frmSCLTrans.frx":2644D
      Style           =   1  '그래픽
      TabIndex        =   15
      Top             =   4500
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.TextBox txtSQL 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   14
      Top             =   4470
      Visible         =   0   'False
      Width           =   8025
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H8000000E&
      Caption         =   "나가기"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   13470
      Picture         =   "frmSCLTrans.frx":2654F
      Style           =   1  '그래픽
      TabIndex        =   13
      Top             =   900
      Width           =   1395
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   1395
      Left            =   120
      TabIndex        =   8
      Top             =   90
      Width           =   10335
      _Version        =   65536
      _ExtentX        =   18230
      _ExtentY        =   2461
      _StockProps     =   15
      BackColor       =   13160660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtUID 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   8670
         TabIndex        =   24
         Top             =   1005
         Width           =   1305
      End
      Begin VB.CommandButton cmdDisplay 
         Caption         =   "불러오기"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   8880
         TabIndex        =   21
         Top             =   60
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.TextBox txtCol2 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5340
         TabIndex        =   4
         Text            =   "50"
         Top             =   1020
         Width           =   675
      End
      Begin VB.TextBox txtRow2 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6030
         TabIndex        =   5
         Text            =   "200"
         Top             =   1020
         Width           =   675
      End
      Begin VB.TextBox txtCol 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1650
         TabIndex        =   2
         Text            =   "1"
         Top             =   1020
         Width           =   675
      End
      Begin VB.TextBox txtRow 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2340
         TabIndex        =   3
         Text            =   "2"
         Top             =   1020
         Width           =   675
      End
      Begin VB.TextBox txtFile 
         BackColor       =   &H00FEE7F3&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1650
         TabIndex        =   0
         Top             =   270
         Width           =   5565
      End
      Begin VB.TextBox txtSheet 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         TabIndex        =   1
         Text            =   "검사결과"
         Top             =   660
         Width           =   5565
      End
      Begin VB.Label Label5 
         BackStyle       =   0  '투명
         Caption         =   "입력자"
         Height          =   255
         Left            =   7860
         TabIndex        =   23
         Top             =   1050
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "마지막 Col && Row"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3720
         TabIndex        =   12
         Top             =   1080
         Width           =   1590
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "시작 Col && Row"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   210
         TabIndex        =   11
         Top             =   1080
         Width           =   1395
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "선택 Sheet  명"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   1305
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "ExcelFileName"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   330
         Width           =   1365
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8790
      Top             =   4860
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000E&
      Caption         =   "Excel 파일 불러오기"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   11910
      MaskColor       =   &H8000000F&
      Picture         =   "frmSCLTrans.frx":27254
      Style           =   1  '그래픽
      TabIndex        =   6
      Top             =   135
      Width           =   1485
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00FFFFFF&
      Caption         =   "조회"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   7530
      Picture         =   "frmSCLTrans.frx":282E7
      Style           =   1  '그래픽
      TabIndex        =   26
      Top             =   240
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '투명
      Caption         =   "※ 만약 's를 입력하고자 하면 ''s로 입력하세요!!!"
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   150
      TabIndex        =   37
      Top             =   9900
      Width           =   7185
   End
End
Attribute VB_Name = "frmSCLTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'================================================
'2005/10/05 이상은 - 출력 버튼 추가
'2007/02/07 이상은 - 메모결과 컬럼 Edit -> Lable로 변경
'                  - 메모결과가 잘림
'================================================

Const colCheck = 1
Const colSlipCode = 2
Const colSlipName = 3
Const colReqDate = 4        '접수일자
Const colBarCode = 5
Const colPID = 6
Const colPName = 7
Const colPJumin = 8
Const colPAge = 9
Const colPSex = 10
Const colDept = 11
Const colWard = 12
Const colOutCode = 13       '의뢰검사코드
Const colOutName = 14       '의뢰검사명
Const colExamCode = 15      '병원검사코드
Const colExamName = 16      '병원검사명
Const colResult = 17        '검사결과
Const colMemoResult = 18    '문장결과
Const colSpecimenCode = 19
Const colSpecimenName = 20
Const colDecision = 21
Const colRemark = 22
Const colBarCode1 = 23      '기타기록
Const colComment = 24
Const colRefValue = 25
Const colSeqNo = 26

Dim gConnect As Integer
Dim saveflag As Integer
Dim gExamCode As String
Dim gExamAlias As String

Dim bBlockSelected As Boolean
Dim lRow1 As Long
Dim lRow2 As Long
Dim lCol1 As Long
Dim lCol2 As Long

Private Sub chkAll_Click()
    vasList.Row = -1
    vasList.col = 1
    
    If chkAll.Value = 1 Then
        vasList.Value = 1
    ElseIf chkAll.Value = 0 Then
        vasList.Value = 0
    End If
End Sub

Private Sub cmdCancel_Click()
    Dim i As Integer

    ClearSpread vasTemp
    ClearSpread vasList
    
    vasList.Row = -1
    vasList.col = 1
    
    chkAll.Value = 0
    chkAll_Click
        
    For i = 1 To vasList.MaxRows
        SetBackColor vasList, i, i, 1, 1, 255, 255, 255
    Next i
End Sub

Private Sub cmdDisplay_Click()
    Dim iRow As Integer
    Dim jRow As Integer
    Dim i As Integer
    
    Dim iiRow As Integer
    Dim jjRow As Integer
    
    Dim sOutExamCode As String      '의뢰검사코드
    Dim sExamCode As String         '병원검사코드
    
    Dim sBarcode As String
    Dim sPID As String
    Dim sRemark As String
    
    jRow = 1
    
    If iRow > vasTemp.DataRowCnt Then
        Exit Sub
    End If
    
    vasList.MaxRows = 800
    
    For iRow = 1 To vasTemp.DataRowCnt
        sPID = ""
        sOutExamCode = ""
        sExamCode = ""
        
        SetText vasList, Trim(GetText(vasTemp, iRow, 2)), jRow, colReqDate          '접수일자
        SetText vasList, Trim(GetText(vasTemp, iRow, 31)), jRow, colBarCode         '검체번호
        
        sPID = Trim(GetText(vasTemp, iRow, 1))

        SetText vasList, sPID, jRow, colPID             '환자번호
        SetText vasList, Trim(GetText(vasTemp, iRow, 4)), jRow, colPName            '환자이름
        
        SetText vasList, Trim(GetText(vasTemp, iRow, 6)), jRow, colPAge             '나이
        SetText vasList, Trim(GetText(vasTemp, iRow, 7)), jRow, colPSex             '성별
        
        sOutExamCode = Trim(GetText(vasTemp, iRow, 39))                             '의뢰 검사코드
        SetText vasList, sOutExamCode, jRow, colOutCode
        
        SetText vasList, Trim(GetText(vasTemp, iRow, 25)), jRow, colOutName         'SRL 검사명칭
        SetText vasList, Trim(GetText(vasTemp, iRow, 25)), jRow, colExamName
        
        '2005/09/05 이상은
        '검사코드 체크하는 부분 수정
        sExamCode = Trim(GetText(vasTemp, iRow, 41))
        
        If sExamCode <> "" Then
            SetText vasList, sExamCode, jRow, colExamCode
        Else
'            '바코드번호, 환자번호, 의뢰코드
'            If GetExamCode(Trim(GetText(vasTemp, iRow, 2)), Trim(GetText(vasTemp, iRow, 6)), sOutExamCode) = 1 Then
'                SetText vasList, gExamCode, jRow, colExamCode           '검사코드
'                'SetText vasList, gExamAlias, jRow, 14                  '검사명칭
'            End If
        End If
        
        SetText vasList, Trim(GetText(vasTemp, iRow, 20)), jRow, colResult          '결과1
        SetText vasList, Trim(GetText(vasTemp, iRow, 38)), jRow, colMemoResult      '결과2
        
'        SetText vasList, "", jRow, colSpecimenCode    '검체코드
'        SetText vasList, "", jRow, colSpecimenName    '검체명
'
'        If Trim(GetText(vasTemp, iRow, 14)) = "." Then
'            SetText vasList, "", jRow, colDecision
'        Else
'            SetText vasList, Trim(GetText(vasTemp, iRow, 14)), jRow, colDecision    '판정
'        End If

'        SetText vasList, Trim(GetText(vasTemp, iRow, 15)), jRow, colRemark          '결과리마크
'
'        SetText vasList, Trim(GetText(vasTemp, iRow, 16)), jRow, colRefValue        '참고치
'
'        SetText vasList, Trim(GetText(vasTemp, iRow, 8)), jRow, colSeqNo            '접수번호
        
        '2004/08/12 이상은 - 결과완료 여부 체크하기
'        gReadBuf(0) = ""
'        SQL = " Select ExamState From ExamRes " & vbCrLf & _
'              " Where HID = '117' " & vbCrLf & _
'              " And PID = '" & Trim(GetText(vasList, jRow, colPID)) & "' " & vbCrLf & _
'              " And SpecimenID = '" & Trim(GetText(vasList, jRow, colBarCode)) & "' " & vbCrLf & _
'              " And ExamCode = '" & Trim(GetText(vasList, jRow, colExamCode)) & "' "
'        res = db_select_Col(SQL)
'
'        If Trim(gReadBuf(0)) = "D" Then
'            SetBackColor vasList, jRow, jRow, 1, 1, 0, 0, 225
'        End If

        jRow = jRow + 1
    Next iRow
    
    vasList.MaxRows = vasList.DataRowCnt
    
    '비고사항 디스플레이
    For iiRow = 1 To vasList.DataRowCnt
        sBarcode = Trim(GetText(vasList, iiRow, colBarCode))
        sRemark = Trim(GetText(vasList, iiRow, colRemark))
        
        If sRemark <> "" Then
            For jjRow = 1 To vasRemark.DataRowCnt
                'If sBarcode = Trim(GetText(vasRemark, jjRow, 1)) Then
                If sBarcode = Trim(GetText(vasRemark, jjRow, 1)) And sRemark = Trim(GetText(vasRemark, jjRow, 2)) Then
                    vasList.SetText colRemark, iiRow, Trim(GetText(vasRemark, jjRow, 3))
                    
                    Exit For
                End If
            Next jjRow
        End If
    Next iiRow
End Sub

'Function GetExamCode_원본(argSCLCode As String)
'
'    GetExamCode_원본 = -1
'
'    If argSCLCode = "" Then
'        Exit Function
'    End If
'
'    gReadBuf(0) = ""
'    gReadBuf(1) = ""
'
'    SQL = " Select b.ExamCode, a.ExamAlias From ExamMaster a, SCLExamMaster b " & vbCrLf & _
'          " Where a.HID = '117' " & vbCrLf & _
'          " And a.HID = b.HID " & vbCrLf & _
'          " And a.ExamCode = b.ExamCode " & vbCrLf & _
'          " And b.SCLCode = '" & Trim(argSCLCode) & "' "
'    res = db_select_Col(SQL)
'
'    If res = 0 Then
'        GetExamCode_원본 = -1
'        Exit Function
'    End If
'
'    If gReadBuf(0) <> "" Then
'        gExamCode = Trim(gReadBuf(0))
'        gExamAlias = Trim(gReadBuf(1))
'    End If
'
'    GetExamCode_원본 = 1
'
'End Function

Function GetExamCode(argBarCode As String, argPID As String, argOutCode As String)
    
    GetExamCode = -1
    
    If argOutCode = "" Then
        Exit Function
    End If
    
    gReadBuf(0) = ""
    gReadBuf(1) = ""
   
    '검사코드, 검사명
'    SQL = " Select b.ExamCode, a.ExamAlias " & vbCrLf & _
'          " From ExamMaster a, NEOExamMaster b, ExamRes c " & vbCrLf & _
'          " Where a.HID = '117' " & vbCrLf & _
'          " And c.PID = '" & Trim(argPID) & "' " & vbCrLf & _
'          " And a.HID = b.HID " & vbCrLf & _
'          " And a.ExamCode = b.ExamCode " & vbCrLf & _
'          " And a.HID = c.HID " & vbCrLf & _
'          " And a.ExamCode = c.ExamCode " & vbCrLf & _
'          " And b.NEOCode = '" & Trim(argOutCode) & "' "
'
'    If argBarCode <> "" Then
'        SQL = SQL & vbCrLf & _
'            " And c.SpecimenID = '" & Trim(argBarCode) & "' "
'    End If

    SQL = " Select a.ExamCode, b.ExamAlias " & CR & _
          " From NEOExamMaster a, ExamMaster b " & CR & _
          " Where a.HID = '117' " & CR & _
          " And a.NEOCODE = '" & Trim(argOutCode) & "' " & CR & _
          " And a.HID = b.HID And a.ExamCode = b.ExamCode "
    res = db_select_Col(SQL)
    
    If res = 0 Then
        GetExamCode = -1
        Exit Function
    End If
    
    If gReadBuf(0) <> "" Then
        gExamCode = Trim(gReadBuf(0))
        gExamAlias = Trim(gReadBuf(1))
    End If
    
    GetExamCode = 1

End Function

Private Sub cmdEquipExam_Click()
    Dim iRow As Integer
    Dim sCnt As String
    
    db_BeginTran
    
    SQL = "Delete From EquipExam "
    SendQuery SQL
    
    For iRow = 2 To vasList.DataRowCnt
        SQL = "Insert Into EquipExam (HID, EquipCode, EquipExamCode, ExamCode, UseFlag, Input_UID, Input_DateTime) " & CR & _
              "Values ('" & Trim(GetText(vasList, iRow, 1)) & "', " & _
                      "'" & Trim(GetText(vasList, iRow, 2)) & "', " & _
                      "'" & Trim(GetText(vasList, iRow, 3)) & "', " & _
                      "'" & Trim(GetText(vasList, iRow, 4)) & "', " & _
                      "'" & Trim(GetText(vasList, iRow, 5)) & "', " & _
                      "'" & Trim(GetText(vasList, iRow, 6)) & "', " & _
                      "'" & Trim(GetText(vasList, iRow, 7)) & "' ) "
        res = SendQuery(SQL)
        If res = -1 Then
            db_RollBack
            SaveQuery SQL
            Exit Sub
        End If
    
    Next iRow
    
    db_Commit
    
    txtSQL.Text = "Select HID, EquipCode, EquipExamCode, ExamCode, UseFlag, Input_UID, Input_DateTime From EquipExam "
    cmdExcute_Click

End Sub

Private Sub cmdExcute_Click()
    Dim i, j As Integer
    Dim argSQL As String
    
On Error GoTo ErrHandle
           
    ClearSpread vasRt
    argSQL = Trim(txtSQL.Text)
    
    Set cmdSQL.ActiveConnection = cn
    cmdSQL.CommandText = argSQL
    Set rs = cmdSQL.Execute
    
    If vasRt.MaxCols < rs.Fields.Count Then
        vasRt.MaxCols = rs.Fields.Count
    End If
    
    If rs.EOF = True Or rs.BOF = True Then
        Exit Sub
    End If
    
    'rs.MoveFirst
    i = 1
    While Not rs.EOF
        If vasRt.MaxRows < i Then
            vasRt.MaxRows = i
        End If
        For j = 0 To rs.Fields.Count - 1
            vasRt.Row = i
            vasRt.col = j + 1
            If IsNull(rs.Fields.Item(j).Value) Then
                vasRt.Text = ""
            Else
                vasRt.Text = rs.Fields.Item(j).Value
            End If
        Next j
        rs.MoveNext
        i = i + 1
    Wend
    
    vasRt.MaxRows = i - 1
    
    rs.Close
    
    Exit Sub
ErrHandle:
    MsgBox err.Number & " : " & Error(err.Number), vbCritical
End Sub

Private Sub cmdPrint_Click()
Dim sCurDate As String
Dim sTitle As String
Dim sHead As String
Dim sFoot As String

On Error GoTo ErrGoto
    
    If vasList.DataRowCnt < 1 Then
        MsgBox "출력할 자료가 없습니다.", , "알 림"
        Exit Sub
    End If
    
    CommonDialog1.CancelError = True
    
    sCurDate = GetDateFull
    
    sTitle = "의뢰결과연계"
    
    sHead = "/fn""궁서체"" /fz""12"" /fb0 /fi0 /fu0 " & "/c" & "▣ " & sTitle & " ▣" & "/n/n "

    sFoot = "/fn""굴림체"" /fz""10"" /fb1 /fi0 /fu0 " & "/l" & sCurDate & "/fn""궁서체"" /fz""11"" /fb1 /fi0 /fu0 /r" & gHosInfo.HName
    
    vasList.PrintOrientation = 1   ' SS_PRINTORIENT_PORTRAIT
    vasList.PrintAbortMsg = "인쇄중 입니다 ..."
    vasList.PrintJobName = "Auto LIS - 의뢰결과연계"
    vasList.PrintHeader = sHead
    vasList.PrintFooter = sFoot
    
    vasList.PrintMarginTop = 2200
    vasList.PrintMarginBottom = 500
    
    vasList.PrintMarginLeft = 100
    vasList.PrintMarginRight = 0
    
    vasList.PrintColor = True
    vasList.PrintGrid = True
    
    'Set printing range
    'vasList.PrintType = 0   'SS_PRINT_ALL(default)

    'Set printing range
    vasList.Row = 1
    vasList.Row2 = vasList.DataRowCnt
    vasList.col = 1
    vasList.Col2 = 15

    vasList.PrintType = PrintTypeCellRange
    
    vasList.PrintShadows = True

    vasList.Action = 13     'SS_ACTION_PRINT
    
    
ErrGoto:
    '사용자가 취소버튼을 눌렀습니다.
    Exit Sub
    
End Sub

Private Sub cmdSave_Click()
    Dim key1, key2, key3, key4 As String
    Dim sOcmNum     As String
    Dim sOdrNum     As String
    Dim sOdrSeq     As String
    
    Dim ResInf      As ResInfRec
    Dim Found       As Integer
    Dim sCurKey     As String
    Dim sResCurKey  As String
    Dim sResCmpKey  As String
    Dim sCmpKey     As String
    Dim sResRetVal  As String
    Dim sRetVal     As String
    Dim sValue      As String
    Dim sNow        As String
    
    Dim iRow        As Integer
    Dim iiRow       As Integer
    Dim jRow        As Integer
    
    Dim i           As Integer
    Dim j           As Integer
    Dim k           As Integer
    Dim ii          As Integer
    Dim sLen        As String
    Dim iPos        As Integer
    
    Dim sExamState  As String
    Dim sSpecID     As String
    Dim sReceNo     As String
    Dim sPID        As String
    Dim sPName      As String
    
    Dim sResClassCode   As String
    Dim sOutCode        As String          '의뢰 검사코드
    Dim sExamCode       As String         '병원 검사코드
    Dim sExamName       As String
    Dim sRefValue       As String         '참고치
    Dim sResult         As String
    Dim sCnt As String
    Dim sTmp As String
    Dim sTmp1 As String
    Dim sTmp2 As String
    Dim sTmp3 As String
    Dim sTmp4 As String
    Dim sSeqNo As String
    
    Dim sRemark As String
    Dim sResMemo
    Dim sResMemo1

    If MsgBox("저장하시겠습니까?", vbOKCancel) = vbCancel Then
        Exit Sub
    End If
       
    sSpecID = ""
    sReceNo = ""
    
    sOutCode = ""
    sExamCode = ""
    
'    db_BeginTran
    
    For iRow = 1 To vasList.DataRowCnt
        vasList.Row = iRow
        vasList.col = 1
        
        If vasList.Value = 1 Then
            If Trim(GetText(vasList, iRow, colBarCode)) = "" Then
                MsgBox "검체번호를 확인하세요", vbExclamation
                Exit Sub
            End If
            
'            If Trim(txtUID) = "" Then
'                MsgBox "입력자를 확인하세요"
'                txtUID.SetFocus
'                Exit Sub
'            End If
            
            sSpecID = Trim(GetText(vasList, iRow, colBarCode))
            sPID = Trim(GetText(vasList, iRow, colPID))
        
            If Trim(GetText(vasList, iRow, colExamCode)) = "" Then
                MsgBox "검사코드를 확인하세요!", vbExclamation
                Exit Sub
            End If
            
            key1 = ""
            
            If Len(sSpecID) = 10 Then
                SQL = "Select barcode, ocmnum, odrnum, odrseq, acpdte, acpcod, acpnum, spmcode "
                SQL = SQL & " from barcodeinfo "
                SQL = SQL & "where barcode = '" & sSpecID & "' "
                SQL = SQL & "and acpcod = 'REF' "
                res = db_select_Col(SQL)
                If Trim(gReadBuf(0)) = sSpecID Then
                    If Len(Trim(gReadBuf(4))) = 10 Then
                        key1 = Format(Trim(gReadBuf(4)), "YYYYMMDD")
                    Else
                        key1 = Trim(gReadBuf(4))
                    End If
                    key2 = Trim(gReadBuf(5))
                    key3 = SetSpace(Trim(gReadBuf(6)), 10)
                    key4 = Trim(gReadBuf(7))
                Else
                    MsgBox Trim(GetText(vasList, iRow, colPName)) & " 접수가 안 되었습니다", vbInformation
                    Exit Sub
                End If
                If key1 = "" Then
                    sOcmNum = Trim(gReadBuf(1))
                    sOdrNum = Trim(gReadBuf(2))
                    sOdrSeq = Trim(gReadBuf(3))
                    
                    sResCurKey = sOcmNum & Chr(5) & sOdrNum & Chr(5) & sOdrSeq & Chr(5)
                    
                    i = 0
                    
                    sResCmpKey = ""
                        
                    sResCurKey = mSetNext("ResInfOcmOdrOdr", sResCurKey)
                    Do
                        sResCurKey = mReadNext("ResInfOcmOdrOdr", sResCurKey, sResCmpKey, sResRetVal)
                        Debug.Print sResRetVal
                        'Save_Raw_Data sResRetVal
                        
                        If sResCurKey = "" Then Exit Do
                        
                        If piece(sResRetVal, Chr(5), 6) <> sOcmNum Then Exit Do
                        If piece(sResRetVal, Chr(5), 39) <> sOdrNum Then Exit Do
                        
                        key1 = piece(sResRetVal, Chr(5), 1)
                        key2 = piece(sResRetVal, Chr(5), 2)
                        key3 = piece(sResRetVal, Chr(5), 3)
                        key4 = piece(sResRetVal, Chr(5), 4)
                        
                        Exit Do
                    Loop
                        
                    If key1 = "" Then Exit Sub
                
                End If
                
            Else
'                key1 = sDate
'                key2 = gEquipSlip
'                key3 = SetSpace(GetText(vasID, asRow, colReceNo), 10)
                'key4 = ResInf.ResSpmCod
            End If

        
            
            '검사항목 형식 메모가 아닌경우***********************************************
            If Trim(GetText(vasList, iRow, colResult)) <> "" Then
                sExamCode = Trim(GetText(vasList, iRow, colExamCode))
                sResult = Trim(GetText(vasList, iRow, colResult))
                
                sCurKey = key1 & Chr(5) & key2 & Chr(5) & key3 & Chr(5) & key4 & Chr(5) & sExamCode & Chr(5)
                
                'Save_Raw_Data sCurKey & " : " & sResult
                
                'Debug.Print sCurKey & " : " & sResult
                sCmpKey = sCurKey
                sCurKey = mSetReadEqual("ResInf", sCurKey, sValue)
                If Trim(sCurKey) <> "" Then
                    Call ResInfLoad(sValue, ResInf)
                            
                    ResInf.ResMzhMnt = sResult       '검사결과
                    
                    Call ResInfStore(sCurKey, sValue, ResInf)
                    If Not mUpdate("ResInf", sCurKey, sValue) Then
                        'Save_Raw_Data ResInf.ResShtNam & "(" & ResInf.ResLabCod & ") 의 결과 Update Error"
                    End If
                    
'                    If UCase(Trim(sResult)) = "Positive" Then
'                        sResult = Trim(GetText(vasTemp, i, 4)) & "/R"
'                    Else
'                        sResult = Trim(GetText(vasTemp, i, 4)) & "/N"
'                    End If
                Else
                    sResult = Trim(GetText(vasList, iRow, colResult))
                End If

            
                sCurKey = key1 & Chr(5) & key2 & Chr(5) & key3 & Chr(5) & key4 & Chr(5) & sExamCode & Chr(5)
                
                'Save_Raw_Data sCurKey & " : " & sResult
                Debug.Print sCurKey & " : " & sResult
                
                sCmpKey = sCurKey
                sCurKey = mSetReadEqual("ResInf", sCurKey, sValue)
                If Trim(sCurKey) <> "" Then
                    Call ResInfLoad(sValue, ResInf)
    
                    ResInf.ResMzhMnt = sResult       '검사결과
                    
                    Call ResInfStore(sCurKey, sValue, ResInf)
                    If Not mUpdate("ResInf", sCurKey, sValue) Then
                        SetBackColor vasList, iRow, iRow, 1, 1, 255, 0, 0
                        'Save_Raw_Data ResInf.ResShtNam & "(" & ResInf.ResLabCod & ") 의 결과 Update Error"
                        'Exit Function
                    End If
                    
                    SetBackColor vasList, iRow, iRow, 1, 1, 202, 255, 112
                End If
            End If
            

            '검사항목 형식 메모인 경우***************************************************
            If Trim(GetText(vasList, iRow, colMemoResult)) <> "" Then
                sExamCode = Trim(GetText(vasList, iRow, colExamCode))
                sResult = Trim(GetText(vasList, iRow, colMemoResult))
                
                sCurKey = key1 & Chr(5) & key2 & Chr(5) & key3 & Chr(5) & key4 & Chr(5) & sExamCode & Chr(5)
                
                'Save_Raw_Data sCurKey & " : " & sResult
                
                'Debug.Print sCurKey & " : " & sResult
                sCmpKey = sCurKey
                sCurKey = mSetReadEqual("ResInf", sCurKey, sValue)
                If Trim(sCurKey) <> "" Then
                    Call ResInfLoad(sValue, ResInf)
                            
                    ResInf.ResMzhMnt = sResult       '검사결과
                    
                    Call ResInfStore(sCurKey, sValue, ResInf)
                    If Not mUpdate("ResInf", sCurKey, sValue) Then
                        'Save_Raw_Data ResInf.ResShtNam & "(" & ResInf.ResLabCod & ") 의 결과 Update Error"
                    End If
                    
'                    If UCase(Trim(sResult)) = "Positive" Then
'                        sResult = Trim(GetText(vasTemp, i, 4)) & "/R"
'                    Else
'                        sResult = Trim(GetText(vasTemp, i, 4)) & "/N"
'                    End If
                Else
                    sResult = Trim(GetText(vasList, i, colResult))
                End If

            
                sCurKey = key1 & Chr(5) & key2 & Chr(5) & key3 & Chr(5) & key4 & Chr(5) & sExamCode & Chr(5)
                
                'Save_Raw_Data sCurKey & " : " & sResult
                Debug.Print sCurKey & " : " & sResult
                
                sCmpKey = sCurKey
                sCurKey = mSetReadEqual("ResInf", sCurKey, sValue)
                If Trim(sCurKey) <> "" Then
                    Call ResInfLoad(sValue, ResInf)
    
                    ResInf.ResMzhMnt = sResult       '검사결과
                    
                    Call ResInfStore(sCurKey, sValue, ResInf)
                    If Not mUpdate("ResInf", sCurKey, sValue) Then
                        SetBackColor vasList, iRow, iRow, 1, 1, 255, 0, 0
                        'Save_Raw_Data ResInf.ResShtNam & "(" & ResInf.ResLabCod & ") 의 결과 Update Error"
                        'Exit Function
                    End If
                    
                    SetBackColor vasList, iRow, iRow, 1, 1, 202, 255, 112
                End If
            End If
        End If
    Next iRow
    
'    db_Commit
    
    MsgBox "작업 완료 되었습니다!", vbExclamation
End Sub

Private Sub cmdSave1_Click()
    Dim iRow As Integer
    Dim iiRow As Integer
    Dim jRow As Integer
    
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim ii As Integer
    Dim sLen As String
    Dim iPos As Integer
    
    Dim sExamState As String
    Dim sSpecID As String
    Dim sReceNo As String
    Dim sPID As String
    Dim sPName As String
    
    Dim sResClassCode As String
    Dim sOutCode As String          '의뢰 검사코드
    Dim sExamCode As String         '병원 검사코드
    Dim sExamName As String
    Dim sRefValue As String         '참고치
    
    Dim sCnt As String
    Dim sTmp As String
    Dim sTmp1 As String
    Dim sTmp2 As String
    Dim sTmp3 As String
    Dim sTmp4 As String
    Dim sSeqNo As String
    
    Dim sRemark As String
    Dim sResMemo
    Dim sResMemo1
    
    'Triple Test 결과 ======
    Dim sEtc As String
    Dim sAFPMOM As String
    Dim sHCGMOM As String
    Dim sHCGMOM1 As String
    Dim sE3MOM As String
    Dim sE3MOM1 As String
    Dim sRDNRes As String
    Dim sRNTRes As String
    Dim sScreenRes As String
    
    Dim sTripleRes As String
    Dim sAFPRes As String
    Dim sHCGRes As String
    Dim sE3Res As String
    
    'AFB Sensitivity 결과 ==
    Dim sRes As String
    Dim sRes1 As String
    Dim sRes2 As String
    Dim sRes3 As String
    Dim sRes4 As String
    Dim sRes5 As String
    Dim sRes6 As String
    Dim sRes7 As String
    Dim sRes8 As String
    Dim sRes9 As String
    Dim sRes10 As String
    Dim sRes11 As String
    Dim sRes12 As String
    '=======================

    If MsgBox("저장하시겠습니까?", vbOKCancel) = vbCancel Then
        Exit Sub
    End If
       
    sSpecID = ""
    sReceNo = ""
    
    sOutCode = ""
    sExamCode = ""
    
'    db_BeginTran
    
    For iRow = 1 To vasList.DataRowCnt
        vasList.Row = iRow
        vasList.col = 1
        
        If vasList.Value = 1 Then
            If Trim(GetText(vasList, iRow, colBarCode)) = "" Then
                MsgBox "검체번호를 확인하세요", vbExclamation
                Exit Sub
            End If
            
            If Trim(txtUID) = "" Then
                MsgBox "입력자를 확인하세요"
                txtUID.SetFocus
                Exit Sub
            End If
            
            sSpecID = Trim(GetText(vasList, iRow, colBarCode))
            sPID = Trim(GetText(vasList, iRow, colPID))
            
            '접수번호
            SQL = " Select max(ReceNo) From ExamRes " & vbCrLf & _
                  " Where HID = '117' " & vbCrLf & _
                  " And PID = '" & sPID & "' " & vbCrLf & _
                  " And SpecimenID = '" & Trim(sSpecID) & "' "
            res = db_select_Var(SQL, sReceNo)
            
            '2004/08/10 이상은 - 검사코드가 없을 경우 처리하기
            If Trim(GetText(vasList, iRow, colExamCode)) = "" Then
                MsgBox "검사코드를 확인하세요!", vbExclamation
                Exit Sub
            End If
            
            '검사항목 결과종류
            gReadBuf(0) = ""
            SQL = " Select ResClassCode " & CR & _
                  "From ExamMaster " & CR & _
                  " Where HID = '117' " & CR & _
                  " And ExamCode = '" & Trim(GetText(vasList, iRow, colExamCode)) & "' " & vbCrLf & _
                  " And UseFlag = 'Y' "
            res = db_select_Col(SQL)

            If res = 1 Then
                sResClassCode = Trim(gReadBuf(0))
            End If
            
            If sResClassCode = "" Then
                MsgBox "검사결과 종류가 정해져있지 않습니다." & CR & _
                       "관리자에서 확인하세요!", vbExclamation
                Exit Sub
            ElseIf sResClassCode = "4" Then     '미생물 검사 결과
                MsgBox "미생물 결과 종류는 입력할 수 없습니다" & CR & _
                       "관리자에서 확인하세요!", vbExclamation
                Exit Sub
            End If
            
            '검사항목 형식 메모가 아닌경우***********************************************
            If sResClassCode = "1" Or sResClassCode = "2" Then
                'db_BeginTran
                
                '만약에 결과에 재검이라는 말이 포함되면 결과완료 하지 말 것
                sExamState = ""
                
                SQL = " Select ExamState From ExamRes " & CR & _
                      " Where HID = '117' " & vbCrLf & _
                      " And PID = '" & Trim(GetText(vasList, iRow, colPID)) & "' " & vbCrLf & _
                      " And ReceNo = '" & Trim(sReceNo) & "' " & vbCrLf & _
                      " And SpecimenID = '" & Trim(sSpecID) & "' " & vbCrLf & _
                      " And ExamCode = '" & Trim(GetText(vasList, iRow, colExamCode)) & "' "
                res = db_select_Col(SQL)
                
                If Trim(GetText(vasList, iRow, colResult)) <> "" Then
                    iPos = InStr(1, Trim(GetText(vasList, iRow, colResult)), "재검")
                    
                    If iPos > 0 Then
                        sExamState = Trim(gReadBuf(0))
                    Else
                        sExamState = "D"
                    End If
                End If
            
                '2009.09.21 이상은
                sRemark = ""
                If Trim(GetText(vasList, iRow, colRemark)) <> "" Then
                    iPos = InStr(1, Trim(GetText(vasList, iRow, colRemark)), "'s")
                    If iPos > 0 Then
                        sRemark = Mid(Trim(GetText(vasList, iRow, colRemark)), 1, iPos - 1) & "'" & Mid(Trim(GetText(vasList, iRow, colRemark)), iPos)
                    Else
                        sRemark = Trim(GetText(vasList, iRow, colRemark))
                    End If
                 End If
                 
'                SQL = " Update ExamRes Set " & vbCrLf & _
'                      " Result = '" & Trim(GetText(vasList, iRow, colResult)) & "', " & vbCrLf & _
'                      " RefValue = '" & Trim(GetText(vasList, iRow, colRefValue)) & "', " & vbCrLf & _
'                      " Decision = '" & Trim(GetText(vasList, iRow, colDecision)) & "', " & vbCrLf & _
'                      " PanicFlag = '', " & vbCrLf & _
'                      " DeltaFlag = '', " & vbCrLf & _
'                      " ExamState = '" & sExamState & "', " & vbCrLf & _
'                      " Remark = '" & Trim(GetText(vasList, iRow, colRemark)) & "', " & vbCrLf & _
'                      " ExamUID = '" & Trim(txtUID.Text) & "', " & vbCrLf & _
'                      " ExamDate = '" & GetDateFull & "',  " & vbCrLf & _
'                      " Input_UID = '" & Trim(txtUID.Text) & "', " & vbCrLf & _
'                      " Input_DateTime ='" & GetDateFull & "' "
                                            
'2008.07.07 이상은 - 검사상태 업데이트 하지 말 것
'                SQL = " Update ExamRes Set " & vbCrLf & _
'                      " Result = '" & Trim(GetText(vasList, iRow, colResult)) & "', " & vbCrLf & _
'                      " RefValue = '상세참조', " & vbCrLf & _
'                      " Decision = '" & Trim(GetText(vasList, iRow, colDecision)) & "', " & vbCrLf & _
'                      " PanicFlag = '', " & vbCrLf & _
'                      " DeltaFlag = '', " & vbCrLf & _
'                      " ExamState = '" & sExamState & "', " & vbCrLf & _
'                      " Remark = '" & Trim(GetText(vasList, iRow, colRemark)) & "', " & vbCrLf & _
'                      " ExamUID = '" & Trim(txtUID.Text) & "', " & vbCrLf & _
'                      " ExamDate = TO_DATE('" & GetDateFull & "', 'mm/dd/yyyy hh24:mi:ss'),  " & vbCrLf & _
'                      " Input_UID = '" & Trim(txtUID.Text) & "', " & vbCrLf & _
'                      " Input_DateTime = TO_DATE('" & GetDateFull & "', 'mm/dd/yyyy hh24:mi:ss') "
                      
                '2009.03.31 이상은 상세참조 -> 상세결과조회
                SQL = " Update ExamRes Set " & vbCrLf & _
                      " Result = '" & Trim(GetText(vasList, iRow, colResult)) & "', " & vbCrLf & _
                      " RefValue = '상세결과조회', " & vbCrLf & _
                      " Decision = '" & Trim(GetText(vasList, iRow, colDecision)) & "', " & vbCrLf & _
                      " PanicFlag = '', " & vbCrLf & _
                      " DeltaFlag = '', " & vbCrLf & _
                      " Remark = '" & sRemark & "', " & vbCrLf & _
                      " ExamUID = '" & Trim(txtUID.Text) & "', " & vbCrLf & _
                      " ExamDate = TO_DATE('" & GetDateFull & "', 'mm/dd/yyyy hh24:mi:ss'),  " & vbCrLf & _
                      " Input_UID = '" & Trim(txtUID.Text) & "', " & vbCrLf & _
                      " Input_DateTime = TO_DATE('" & GetDateFull & "', 'mm/dd/yyyy hh24:mi:ss') "
                      
                SQL = SQL & vbCrLf & _
                      " Where HID = '117' " & vbCrLf & _
                      " And PID = '" & Trim(GetText(vasList, iRow, colPID)) & "' " & vbCrLf & _
                      " And ReceNo = '" & Trim(sReceNo) & "' " & vbCrLf & _
                      " And SpecimenID = '" & Trim(sSpecID) & "' " & vbCrLf & _
                      " And ExamCode = '" & Trim(GetText(vasList, iRow, colExamCode)) & "' "
                res = SendQuery(SQL)
                
                If res = -1 Then
                    db_RollBack
                    
                    SaveQuery SQL
                    SetBackColor vasList, iRow, iRow, 1, 1, 255, 0, 0
                    
                    Exit Sub
                ElseIf res = 0 Then
                    SaveQuery SQL
                    Exit Sub
                End If
                
                '비고사항 저장하기******************************
'                SQL = " Update ExamRes Set " & vbCrLf & _
'                      " Remark = '" & Trim(GetText(vasList, iRow, colRemark)) & "' " & vbCrLf & _
'                      " Where HID = '117' " & vbCrLf & _
'                      " And PID = '" & Trim(GetText(vasList, iRow, colPID)) & "' " & vbCrLf & _
'                      " And ReceNo = '" & Trim(sReceNo) & "' " & vbCrLf & _
'                      " And SpecimenID = '" & Trim(sSpecID) & "' " & vbCrLf & _
'                      " And ExamCode = '" & Trim(GetText(vasList, iRow, colExamCode)) & "' "
'
'                res = SendQuery(SQL)
'                If res = -1 Then
'                    db_RollBack
'
'                    SaveQuery SQL
'                    SetBackColor vasList, iRow, iRow, 1, 1, 255, 0, 0
'
'                    Exit Sub
'                ElseIf res = 0 Then
'                    SaveQuery SQL
'                    Exit Sub
'                End If
'
                db_Commit
'
                If res = 1 Then
                    SetBackColor vasList, iRow, iRow, 1, 1, 202, 255, 112
                ElseIf res = -1 Then
                    SetBackColor vasList, iRow, iRow, 1, 1, 255, 0, 0
                End If
                
                '의뢰결과 테이블에 저장하기*********************
                sCnt = ""
                SQL = " Select count(*) From OutExamRes " & vbCrLf & _
                      " Where HID = '117' " & vbCrLf & _
                      " And PID =  '" & Trim(sPID) & "' " & vbCrLf & _
                      " And ReceNo = '" & Trim(sReceNo) & "' " & vbCrLf & _
                      " And SpecimenID = '" & Trim(sSpecID) & "' " & vbCrLf & _
                      " And ExamCode = '" & Trim(GetText(vasList, iRow, colExamCode)) & "' "
                res = db_select_Col(SQL)
                sCnt = Trim(gReadBuf(0))
                If sCnt = "" Then sCnt = "0"
                
                If sCnt = "0" Then
                    SQL = " Insert Into OutExamRes(HID, PID, ReceNo, SpecimenID, ExamCode, " & vbCrLf & _
                          " Result, RefValue, Decision, Remark, Input_UID, Input_DateTime) " & vbCrLf & _
                          " values('117', '" & Trim(sPID) & "', '" & Trim(sReceNo) & "', " & vbCrLf & _
                          "       '" & Trim(sSpecID) & "', '" & Trim(GetText(vasList, iRow, colExamCode)) & "', " & vbCrLf & _
                          "       '" & Trim(GetText(vasList, iRow, colResult)) & "', '" & Trim(GetText(vasList, iRow, colRefValue)) & "'," & vbCrLf & _
                          "       '" & Trim(GetText(vasList, iRow, colDecision)) & "', '" & sRemark & "', " & vbCrLf & _
                          "       '" & Trim(txtUID.Text) & "', TO_DATE('" & GetDateFull & "', 'mm/dd/yyyy hh24:mi:ss')) "
                Else
                    SQL = " Update OutExamRes Set " & vbCrLf & _
                          " Result = '" & Trim(GetText(vasList, iRow, colResult)) & "', " & vbCrLf & _
                          " RefValue = '" & Trim(GetText(vasList, iRow, colRefValue)) & "'," & vbCrLf & _
                          " Decision = '" & Trim(GetText(vasList, iRow, colDecision)) & "', " & vbCrLf & _
                          " Remark = '" & sRemark & "', " & vbCrLf & _
                          " Input_UID = '" & Trim(txtUID.Text) & "', " & vbCrLf & _
                          " Input_DateTime = TO_DATE('" & GetDateFull & "', 'mm/dd/yyyy hh24:mi:ss') " & vbCrLf & _
                          " Where HID = '117' " & vbCrLf & _
                          " And PID =  '" & Trim(sPID) & "' " & vbCrLf & _
                          " And ReceNo = '" & Trim(sReceNo) & "' " & vbCrLf & _
                          " And SpecimenID = '" & Trim(sSpecID) & "' " & vbCrLf & _
                          " And ExamCode = '" & Trim(GetText(vasList, iRow, colExamCode)) & "' "
                End If
                res = SendQuery(SQL)
                If res = -1 Then
                    db_RollBack
                    
                    SaveQuery SQL
                    SetBackColor vasList, iRow, iRow, 1, 1, 255, 0, 0
                    
                    Exit Sub
                ElseIf res = 0 Then
                    SaveQuery SQL
                    Exit Sub
                End If
                
                db_Commit
                
                If res = 1 Then
                    SetBackColor vasList, iRow, iRow, 1, 1, 202, 255, 112
                ElseIf res = -1 Then
                    SetBackColor vasList, iRow, iRow, 1, 1, 255, 0, 0
                End If
            End If
            
            '결과2 를 ExamResMemo에 Insert
            '검사항목 형식 메모인 경우***************************************************
            If sResClassCode = "3" Then
                sOutCode = Trim(GetText(vasList, iRow, colOutCode))
                sExamCode = Trim(GetText(vasList, iRow, colExamCode))
                sExamName = Trim(GetText(vasList, iRow, colOutName))
                
                sRefValue = Trim(GetText(vasList, iRow, colRefValue))
                
'=================================================================================
                Select Case sOutCode
'2008.11.04 이상은 - AbA1c 메모결과로 변경
'                Case "0009100"    'HbA1c
'                    sTmp = ""
'
'                    For jRow = iRow + 1 To iRow + 3
'                        vasList.Row = jRow
'                        vasList.col = 1
'
'                        If vasList.Value = 1 Then
'                            If sSpecID = Trim(GetText(vasList, jRow, colBarCode)) Then
'                                If sTmp = "" Then
'                                    sTmp = NLeftString("검사명", 15) & NMidString("결과", 10) & NMidString("판정", 10) & "참조치"
'
'                                    sTmp = sTmp & CR & NLeftString(Trim(GetText(vasList, jRow, colOutName)), 15) & NMidString(Trim(GetText(vasList, jRow, colResult)), 10) & NMidString(Trim(GetText(vasList, jRow, colDecision)), 10) & Trim(GetText(vasList, jRow, colRefValue))
'
'                                    SetBackColor vasList, iRow, iRow, 1, 1, 202, 255, 112
'                                Else
''                                    sTmp1 = ""
''                                    sTmp2 = ""
''                                    sTmp1 = Trim(GetText(vasList, jRow, colRefValue))
''                                    If InStr(Trim(GetText(vasList, jRow, colRefValue)), Chr(13)) > 0 Then
''                                        iPos = InStr(Trim(GetText(vasList, jRow, colRefValue)), Chr(13))
''                                        sTmp2 = Trim(Mid(sTmp1, 1, iPos - 1))
''
''                                        sTmp1 = Mid(sTmp1, iPos + 2)
''                                    End If
''
''                                    If InStr(sTmp1, Chr(13)) > 0 Then
''                                        iPos = InStr(sTmp1, Chr(13))
''                                        sTmp3 = Trim(Mid(sTmp1, 1, iPos - 1))
''
''                                        sTmp1 = Mid(sTmp1, iPos + 2)
''                                    Else
''                                        sTmp2 = sTmp2 & " " & Trim(sTmp1)
''                                    End If
'
'                                    Select Case Trim(GetText(vasList, jRow, colOutCode))
'                                    Case "0009102"
'                                        sRefValue = ""
'                                        sRefValue = "mg/dL 2 개월 평균혈당치의미"
'                                    Case "0009103"
'                                        sRefValue = ""
'                                        sRefValue = "Good: ≤ 7.0 Fair: 7.1-10.0 Poor: > 10.0"
'                                    End Select
'                                    sTmp = sTmp & CR & NLeftString(Trim(GetText(vasList, jRow, colOutName)), 15) & NMidString(Trim(GetText(vasList, jRow, colResult)), 10) & NMidString(Trim(GetText(vasList, jRow, colDecision)), 10) & sRefValue
'                                End If
'                                SetBackColor vasList, jRow, jRow, 1, 1, 202, 255, 112
'
'                                iRow = jRow
'                            Else
'                                'iRow = jRow
'                                iRow = jRow - 1
'                                Exit For
'                            End If
'                        Else
'                            iRow = jRow
'                            Exit For
'                        End If
'                    Next jRow
'
'                    sResMemo = ""
'                    sResMemo = "< " & sExamName & " 결과 > " & CR & CR & sTmp
                '********************************************************************
                
'                Case "1000100"    'CCR         '2008.12.08 이상은 - 결과형식으로 메모로 변경함
'                    sTmp = ""
'
'                    For jRow = iRow To iRow + 7
'                        vasList.Row = jRow
'                        vasList.col = 1
'
'                        If vasList.Value = 1 Then
'                            If sSpecID = Trim(GetText(vasList, jRow, colBarCode)) Then
'                                If sTmp = "" Then
'                                    sTmp = NLeftString("검사명", 22) & NMidString("결과", 10) & NMidString("판정", 10) & "참조치"
'
'                                    sTmp = sTmp & CR & NLeftString(Trim(GetText(vasList, jRow, colOutName)), 22) & NMidString(Trim(GetText(vasList, jRow, colResult)), 10) & NMidString(Trim(GetText(vasList, jRow, colDecision)), 10) & Trim(GetText(vasList, jRow, colRefValue))
'
'                                    SetBackColor vasList, iRow, iRow, 1, 1, 202, 255, 112
'                                Else
'                                    sTmp = sTmp & CR & NLeftString(Trim(GetText(vasList, jRow, colOutName)), 22) & NMidString(Trim(GetText(vasList, jRow, colResult)), 10) & NMidString(Trim(GetText(vasList, jRow, colDecision)), 10) & sRefValue
'                                End If
'                                SetBackColor vasList, jRow, jRow, 1, 1, 202, 255, 112
'
'                                iRow = jRow
'                            Else
'                                'iRow = jRow
'                                iRow = jRow - 1
'                                Exit For
'                            End If
'                        Else
'                            iRow = jRow
'                            Exit For
'                        End If
'                    Next jRow
'
'                    sResMemo = ""
'                    sResMemo = "< " & sExamName & " 결과 > " & CR & CR & sTmp
                '********************************************************************
                
'                Case "70120"    'Culture & Sensitivity
'                    sTmp = ""
'
'                    For jRow = iRow To vasList.DataRowCnt
'                        vasList.Row = jRow
'                        vasList.col = 1
'
'                        If vasList.Value = 1 Then
'                            If sSpecID = Trim(GetText(vasList, jRow, colBarCode)) Then
'                                If sTmp = "" Then
'                                    sTmp = Trim(GetText(vasList, jRow, colMemoResult))
'                                Else
'                                    sTmp = sTmp & CR & Trim(GetText(vasList, jRow, colMemoResult))
'                                End If
'                                SetBackColor vasList, jRow, jRow, 1, 1, 202, 255, 112
'                            Else
'                                'iRow = jRow
'                                iRow = jRow - 1
'                                Exit For
'                            End If
'                        Else
'                            iRow = jRow
'                            Exit For
'                        End If
'                    Next jRow
'
'                    sResMemo = ""
'                    sResMemo = "< " & sExamName & " 결과 > " & CR & CR & sTmp
                '********************************************************************
                
'                Case "70230"    'Blood Culture
'                    '2005/10/07 이상은
'                    '몇번째 Blood Culture인지 추가함
'                    sTmp1 = ""
'                    sPName = Trim(GetText(vasList, iRow, colPName))
'
'                    If IsNumeric(Right(sPName, 1)) = True Then
'                        sTmp1 = "Blood Culture" & " " & Right(sPName, 1)
'                    End If
'
'                    sTmp = ""
'
'                    For jRow = iRow To vasList.DataRowCnt
'                        vasList.Row = jRow
'                        vasList.col = 1
'
'                        If vasList.Value = 1 Then
'                            If sSpecID = Trim(GetText(vasList, jRow, colBarCode)) Then
'                                If sTmp = "" Then
'                                    sTmp = Trim(GetText(vasList, jRow, colMemoResult))
'                                Else
'                                    sTmp = sTmp & CR & Trim(GetText(vasList, jRow, colMemoResult))
'                                End If
'                                SetBackColor vasList, jRow, jRow, 1, 1, 202, 255, 112
'                            Else
'                                'iRow = jRow
'                                iRow = jRow - 1
'                                Exit For
'                            End If
'                        Else
'                            iRow = jRow
'                            Exit For
'                        End If
'                    Next jRow
'
'                    sResMemo = ""
'                    sResMemo = "< " & sExamName & " 결과 > " & CR & CR & sTmp & CR & CR & sTmp1
                '********************************************************************
                
'                '2005/10/05 이상은 - ADA검사 Body Fluid인 경우만 결과 + 참고치 => 메모로 함
'                Case "10212"    'ADA검사
'                    If sExamCode = "B2721CC" Then
'                        '2005/10/24 이상은********************************
'                        '참고치 부분 수정
''                        sTmp = ""
''                        sTmp = Trim(GetText(vasList, jRow, colResult))
''
''                        sResMemo = ""
''                        sResMemo = "< " & sExamName & " 결과 > " & CR & CR & sTmp & CR & CR & "참고치 = " & sRefValue
'
'                        sTmp = ""
'
'                        For jRow = iRow To vasList.DataRowCnt
'                            vasList.Row = jRow
'                            vasList.col = 1
'
'                            If vasList.Value = 1 Then
'                                If sSpecID = Trim(GetText(vasList, jRow, colBarCode)) Then
'                                    If sTmp = "" Then
'                                        sTmp = "결과 = " & Trim(GetText(vasList, jRow, colResult)) & " 참고치 = " & Trim(GetText(vasList, jRow, colRefValue))
'                                    Else
'                                        sTmp = sTmp & CR & "결과 = " & Trim(GetText(vasList, jRow, colResult)) & " 참고치 = " & Trim(GetText(vasList, jRow, colRefValue))
'                                    End If
'                                    SetBackColor vasList, jRow, jRow, 1, 1, 202, 255, 112
'                                Else
'                                    'iRow = jRow
'                                    iRow = jRow - 1
'                                    Exit For
'                                End If
'                            Else
'                                iRow = jRow
'                                Exit For
'                            End If
'                        Next jRow
'
'                        sResMemo = ""
'                        sResMemo = "< " & sExamName & " 결과 > " & CR & CR & sTmp
'                    End If
                '********************************************************************
                
'                '2005/11/11 이상은
'                Case "41690"    'Measle Virus IgM
'                    sTmp = ""
'
'                    For jRow = iRow To vasList.DataRowCnt
'                        vasList.Row = jRow
'                        vasList.col = 1
'
'                        If vasList.Value = 1 Then
'                            If sSpecID = Trim(GetText(vasList, jRow, colBarCode)) And sOutCode = Trim(GetText(vasList, jRow, colOutCode)) Then
'                                If Trim(GetText(vasList, iRow, colMemoResult)) <> "" Then
'                                    If sTmp = "" Then
'                                        sTmp = "결과 = " & Trim(GetText(vasList, jRow, colMemoResult)) & " 참고치 = " & Trim(GetText(vasList, jRow, colRefValue))
'                                    Else
'                                        sTmp = sTmp & CR & "결과 = " & Trim(GetText(vasList, jRow, colMemoResult)) & " 참고치 = " & Trim(GetText(vasList, jRow, colRefValue))
'                                    End If
'                                ElseIf Trim(GetText(vasList, iRow, colResult)) <> "" Then
'                                    If sTmp = "" Then
'                                        sTmp = "결과 = " & Trim(GetText(vasList, jRow, colResult)) & " 참고치 = " & Trim(GetText(vasList, jRow, colRefValue))
'                                    Else
'                                        sTmp = sTmp & CR & "결과 = " & Trim(GetText(vasList, jRow, colResult)) & " 참고치 = " & Trim(GetText(vasList, jRow, colRefValue))
'                                    End If
'                                End If
'
'                                SetBackColor vasList, jRow, jRow, 1, 1, 202, 255, 112
'                            Else
'                                'iRow = jRow
'                                iRow = jRow - 1
'                                Exit For
'                            End If
'                        Else
'                            iRow = jRow
'                            Exit For
'                        End If
'                    Next jRow
'
'                    sResMemo = ""
'
'                    sResMemo = "< " & sExamName & " 결과 > " & CR & CR & sTmp
                '********************************************************************
'
'                Case "41740"    'Mumps Virus IgM
'                    sTmp = ""
'
'                    For jRow = iRow To vasList.DataRowCnt
'                        vasList.Row = jRow
'                        vasList.col = 1
'
'                        If vasList.Value = 1 Then
'                            If sSpecID = Trim(GetText(vasList, jRow, colBarCode)) And sOutCode = Trim(GetText(vasList, jRow, colOutCode)) Then
'                                If Trim(GetText(vasList, iRow, colMemoResult)) <> "" Then
'                                    If sTmp = "" Then
'                                        sTmp = "결과 = " & Trim(GetText(vasList, jRow, colMemoResult)) & " 참고치 = " & Trim(GetText(vasList, jRow, colRefValue))
'                                    Else
'                                        sTmp = sTmp & CR & "결과 = " & Trim(GetText(vasList, jRow, colMemoResult)) & " 참고치 = " & Trim(GetText(vasList, jRow, colRefValue))
'                                    End If
'                                ElseIf Trim(GetText(vasList, iRow, colResult)) <> "" Then
'                                    If sTmp = "" Then
'                                        sTmp = "결과 = " & Trim(GetText(vasList, jRow, colResult)) & " 참고치 = " & Trim(GetText(vasList, jRow, colRefValue))
'                                    Else
'                                        sTmp = sTmp & CR & "결과 = " & Trim(GetText(vasList, jRow, colResult)) & " 참고치 = " & Trim(GetText(vasList, jRow, colRefValue))
'                                    End If
'                                End If
'
'                                SetBackColor vasList, jRow, jRow, 1, 1, 202, 255, 112
'                            Else
'                                'iRow = jRow
'                                iRow = jRow - 1
'                                Exit For
'                            End If
'                        Else
'                            iRow = jRow
'                            Exit For
'                        End If
'                    Next jRow
'
'                    sResMemo = ""
'
'                    sResMemo = "< " & sExamName & " 결과 > " & CR & CR & sTmp
                '********************************************************************
                
'                Case "50770"        'H.Pylori lgG
'                    sTmp = ""
'
'                    For jRow = iRow To vasList.DataRowCnt
'                        vasList.Row = jRow
'                        vasList.col = 1
'
'                        If vasList.Value = 1 Then
'                            If sSpecID = Trim(GetText(vasList, jRow, colBarCode)) And sOutCode = Trim(GetText(vasList, jRow, colOutCode)) Then
'                                If Trim(GetText(vasList, iRow, colMemoResult)) <> "" Then
'                                    If sTmp = "" Then
'                                        sTmp = "결과 = " & Trim(GetText(vasList, jRow, colMemoResult)) & " 참고치 = " & Trim(GetText(vasList, jRow, colRefValue))
'                                    Else
'                                        sTmp = sTmp & CR & "결과 = " & Trim(GetText(vasList, jRow, colMemoResult)) & " 참고치 = " & Trim(GetText(vasList, jRow, colRefValue))
'                                    End If
'                                ElseIf Trim(GetText(vasList, iRow, colResult)) <> "" Then
'                                    If sTmp = "" Then
'                                        sTmp = "결과 = " & Trim(GetText(vasList, jRow, colResult)) & " 참고치 = " & Trim(GetText(vasList, jRow, colRefValue))
'                                    Else
'                                        sTmp = sTmp & CR & "결과 = " & Trim(GetText(vasList, jRow, colResult)) & " 참고치 = " & Trim(GetText(vasList, jRow, colRefValue))
'                                    End If
'                                End If
'
'                                SetBackColor vasList, jRow, jRow, 1, 1, 202, 255, 112
'                            Else
'                                'iRow = jRow
'                                iRow = jRow - 1
'                                Exit For
'                            End If
'                        Else
'                            iRow = jRow
'                            Exit For
'                        End If
'                    Next jRow
'
'                    sResMemo = ""
'
'                    sResMemo = "< " & sExamName & " 결과 > " & CR & CR & sTmp
                '********************************************************************
'
'                Case "00960"        '2005/11/15 이상은 추가 - C-Peptide
'                    sTmp = ""
'
'                    For jRow = iRow To vasList.DataRowCnt
'                        vasList.Row = jRow
'                        vasList.col = 1
'
'                        If vasList.Value = 1 Then
'                            If sSpecID = Trim(GetText(vasList, jRow, colBarCode)) And sOutCode = Trim(GetText(vasList, jRow, colOutCode)) Then
'                                If Trim(GetText(vasList, iRow, colMemoResult)) <> "" Then
'                                    If sTmp = "" Then
'                                        sTmp = "결과 = " & Trim(GetText(vasList, jRow, colMemoResult)) & " 참고치 = " & Trim(GetText(vasList, jRow, colRefValue))
'                                    Else
'                                        sTmp = sTmp & CR & "결과 = " & Trim(GetText(vasList, jRow, colMemoResult)) & " 참고치 = " & Trim(GetText(vasList, jRow, colRefValue))
'                                    End If
'                                ElseIf Trim(GetText(vasList, iRow, colResult)) <> "" Then
'                                    If sTmp = "" Then
'                                        sTmp = "결과 = " & Trim(GetText(vasList, jRow, colResult)) & " 참고치 = " & Trim(GetText(vasList, jRow, colRefValue))
'                                    Else
'                                        sTmp = sTmp & CR & "결과 = " & Trim(GetText(vasList, jRow, colResult)) & " 참고치 = " & Trim(GetText(vasList, jRow, colRefValue))
'                                    End If
'                                End If
'
'                                SetBackColor vasList, jRow, jRow, 1, 1, 202, 255, 112
'                            Else
'                                'iRow = jRow
'                                iRow = jRow - 1
'                                Exit For
'                            End If
'                        Else
'                            iRow = jRow
'                            Exit For
'                        End If
'                    Next jRow
'
'                    sResMemo = ""
'
'                    sResMemo = "< " & sExamName & " 결과 > " & CR & CR & sTmp
                '********************************************************************
'
                Case Else
                    sResMemo = ""
                    
                    If Trim(GetText(vasList, iRow, colMemoResult)) <> "" Then
                        iPos = InStr(1, Trim(GetText(vasList, iRow, colMemoResult)), "'s")
                        If iPos > 0 Then
                            sResMemo = Mid(Trim(GetText(vasList, iRow, colMemoResult)), 1, iPos - 1) & "'" & Mid(Trim(GetText(vasList, iRow, colMemoResult)), iPos)
                        Else
                            sResMemo = "< " & sExamName & " 결과 > " & CR & CR & Trim(GetText(vasList, iRow, colMemoResult))
                        End If
                    ElseIf Trim(GetText(vasList, iRow, colResult)) <> "" Then
                        iPos = InStr(1, Trim(GetText(vasList, iRow, colResult)), "'s")
                        If iPos > 0 Then
                            sResMemo = Mid(Trim(GetText(vasList, iRow, colResult)), 1, iPos - 1) & "'" & Mid(Trim(GetText(vasList, iRow, colResult)), iPos)
                        Else
                            sResMemo = "< " & sExamName & " 결과 > " & CR & CR & Trim(GetText(vasList, iRow, colResult))
                        End If
                    End If
                        
                End Select
'=================================================================================
                
                '메모 결과 저장하기
                sCnt = ""
                SQL = " Select count(*) From ExamResMemo " & vbCrLf & _
                      " Where HID = '117' " & vbCrLf & _
                      " And PID =  '" & Trim(sPID) & "' " & vbCrLf & _
                      " And ReceNo = '" & Trim(sReceNo) & "' " & vbCrLf & _
                      " And SpecimenID = '" & Trim(sSpecID) & "' " & vbCrLf & _
                      " And ExamCode = '" & Trim(sExamCode) & "' "
                res = db_select_Var(SQL, sCnt)
            
                If sCnt = "" Then
                    sCnt = "0"
                End If
                
                If sCnt = "0" Then
                    SQL = " Insert Into ExamResMemo (HID, PID, ReceNo, ExamCode, ResMemo, SpecimenID, Input_UID, Input_DateTime) " & vbCrLf & _
                          " Values('117', '" & Trim(sPID) & "', " & vbCrLf & _
                          "        '" & Trim(sReceNo) & "', '" & Trim(sExamCode) & "', " & vbCrLf & _
                          "        '" & Trim(sResMemo) & "', '" & Trim(sSpecID) & "', " & vbCrLf & _
                          "        '" & Trim(txtUID.Text) & "',TO_DATE('" & GetDateFull & "', 'mm/dd/yyyy hh24:mi:ss')) "
                    res = SendQuery(SQL)
                    
                    If res = -1 Then
                        SaveQuery SQL
                        SetBackColor vasList, iRow, iRow, 1, 1, 255, 0, 0
                        'db_RollBack
                        Exit Sub
                    ElseIf res = 0 Then
                        SaveQuery SQL
                        Exit Sub
                    End If
                Else
                    SQL = " Update ExamResMemo Set " & vbCrLf & _
                          " ResMemo = '" & Trim(sResMemo) & "', " & vbCrLf & _
                          " Input_UID = '" & Trim(txtUID.Text) & "', " & vbCrLf & _
                          " Input_DateTime = TO_DATE('" & GetDateFull & "', 'mm/dd/yyyy hh24:mi:ss') " & vbCrLf & _
                          " Where HID = '117' " & vbCrLf & _
                          " And PID =  '" & Trim(sPID) & "' " & vbCrLf & _
                          " And ReceNo = '" & Trim(sReceNo) & "' " & vbCrLf & _
                          " And SpecimenID = '" & Trim(sSpecID) & "' " & vbCrLf & _
                          " And ExamCode = '" & Trim(sExamCode) & "' "
                    res = SendQuery(SQL)
                    
                    If res = -1 Then
                        SaveQuery SQL
                        SetBackColor vasList, iRow, iRow, 1, 1, 255, 0, 0
'                        db_RollBack
                        Exit Sub
                    ElseIf res = 0 Then
                        SaveQuery SQL
                        Exit Sub
                    End If
                End If
                
                '의뢰결과 테이블에 저장하기*********************
                sCnt = ""
                SQL = " Select count(*) From OutExamRes " & vbCrLf & _
                      " Where HID = '117' " & vbCrLf & _
                      " And PID =  '" & Trim(sPID) & "' " & vbCrLf & _
                      " And ReceNo = '" & Trim(sReceNo) & "' " & vbCrLf & _
                      " And SpecimenID = '" & Trim(sSpecID) & "' " & vbCrLf & _
                      " And ExamCode = '" & Trim(GetText(vasList, iRow, colExamCode)) & "' "
                res = db_select_Col(SQL)
                sCnt = Trim(gReadBuf(0))
                If sCnt = "" Then sCnt = "0"
                
                '2009.02.10 이상은 - 메모결과는 비고사항 연계 안하게 할 것
                If sCnt = "0" Then
'                    SQL = " Insert Into OutExamRes(HID, PID, ReceNo, SpecimenID, ExamCode, " & vbCrLf & _
'                          " ResMemo, RefValue, Decision, Remark, Input_UID, Input_DateTime) " & vbCrLf & _
'                          " values('117', '" & Trim(sPID) & "', '" & Trim(sReceNo) & "', " & vbCrLf & _
'                          "       '" & Trim(sSpecID) & "', '" & Trim(GetText(vasList, iRow, colExamCode)) & "', " & vbCrLf & _
'                          "       '" & Trim(sResMemo) & "', '" & Trim(GetText(vasList, iRow, colRefValue)) & "'," & vbCrLf & _
'                          "       '" & Trim(GetText(vasList, iRow, colDecision)) & "', '" & Trim(GetText(vasList, iRow, colRemark)) & "', " & vbCrLf & _
'                          "       '" & Trim(txtUID.Text) & "', TO_DATE('" & GetDateFull & "', 'mm/dd/yyyy hh24:mi:ss')) "

                    SQL = " Insert Into OutExamRes(HID, PID, ReceNo, SpecimenID, ExamCode, " & vbCrLf & _
                          " ResMemo, RefValue, Decision, Remark, Input_UID, Input_DateTime) " & vbCrLf & _
                          " values('117', '" & Trim(sPID) & "', '" & Trim(sReceNo) & "', " & vbCrLf & _
                          "       '" & Trim(sSpecID) & "', '" & Trim(GetText(vasList, iRow, colExamCode)) & "', " & vbCrLf & _
                          "       '" & Trim(sResMemo) & "', '" & Trim(GetText(vasList, iRow, colRefValue)) & "'," & vbCrLf & _
                          "       '" & Trim(GetText(vasList, iRow, colDecision)) & "', '', " & vbCrLf & _
                          "       '" & Trim(txtUID.Text) & "', TO_DATE('" & GetDateFull & "', 'mm/dd/yyyy hh24:mi:ss')) "
                Else
'                    SQL = " Update OutExamRes Set " & vbCrLf & _
'                          " ResMemo = '" & Trim(sResMemo) & "', " & vbCrLf & _
'                          " RefValue = '" & Trim(GetText(vasList, iRow, colRefValue)) & "'," & vbCrLf & _
'                          " Decision = '" & Trim(GetText(vasList, iRow, colDecision)) & "', " & vbCrLf & _
'                          " Remark = '" & Trim(GetText(vasList, iRow, colRemark)) & "', " & vbCrLf & _
'                          " Input_UID = '" & Trim(txtUID.Text) & "', " & vbCrLf & _
'                          " Input_DateTime = TO_DATE('" & GetDateFull & "', 'mm/dd/yyyy hh24:mi:ss') " & vbCrLf & _
'                          " Where HID = '117' " & vbCrLf & _
'                          " And PID =  '" & Trim(sPID) & "' " & vbCrLf & _
'                          " And ReceNo = '" & Trim(sReceNo) & "' " & vbCrLf & _
'                          " And SpecimenID = '" & Trim(sSpecID) & "' " & vbCrLf & _
'                          " And ExamCode = '" & Trim(GetText(vasList, iRow, colExamCode)) & "' "

                    SQL = " Update OutExamRes Set " & vbCrLf & _
                          " ResMemo = '" & Trim(sResMemo) & "', " & vbCrLf & _
                          " RefValue = '" & Trim(GetText(vasList, iRow, colRefValue)) & "'," & vbCrLf & _
                          " Decision = '" & Trim(GetText(vasList, iRow, colDecision)) & "', " & vbCrLf & _
                          " Input_UID = '" & Trim(txtUID.Text) & "', " & vbCrLf & _
                          " Input_DateTime = TO_DATE('" & GetDateFull & "', 'mm/dd/yyyy hh24:mi:ss') " & vbCrLf & _
                          " Where HID = '117' " & vbCrLf & _
                          " And PID =  '" & Trim(sPID) & "' " & vbCrLf & _
                          " And ReceNo = '" & Trim(sReceNo) & "' " & vbCrLf & _
                          " And SpecimenID = '" & Trim(sSpecID) & "' " & vbCrLf & _
                          " And ExamCode = '" & Trim(GetText(vasList, iRow, colExamCode)) & "' "
                End If
                res = SendQuery(SQL)
                If res = -1 Then
                    db_RollBack
                    
                    SaveQuery SQL
                    SetBackColor vasList, iRow, iRow, 1, 1, 255, 0, 0
                    
                    Exit Sub
                ElseIf res = 0 Then
                    SaveQuery SQL
                    Exit Sub
                End If
                
                db_Commit
                
                If res = 1 Then
                    SetBackColor vasList, iRow, iRow, 1, 1, 202, 255, 112
                ElseIf res = -1 Then
                    SetBackColor vasList, iRow, iRow, 1, 1, 255, 0, 0
                End If
                
                If res = 1 Then
                    '2008.07.07 이상은 - 검사상태 업데이트 하지 말 것
'                    SQL = " Update ExamRes Set " & vbCrLf & _
'                          " Result = '*', " & vbCrLf & _
'                          " ExamState = 'D', " & vbCrLf & _
'                          " Input_UID = '" & Trim(txtUID.Text) & "', " & vbCrLf & _
'                          " Input_DateTime = TO_DATE('" & GetDateFull & "', 'mm/dd/yyyy hh24:mi:ss') " & vbCrLf & _
'                          " Where HID = '117' " & vbCrLf & _
'                          " And PID = '" & Trim(sPID) & "' " & vbCrLf & _
'                          " And ReceNo = '" & Trim(sReceNo) & "' " & vbCrLf & _
'                          " And SpecimenID = '" & Trim(sSpecID) & "' " & vbCrLf & _
'                          " And ExamCode = '" & Trim(sExamCode) & "' "
                    
                    '2009.03.31 이상은 결과지참조 -> 상세결과조회
                    SQL = " Update ExamRes Set " & vbCrLf & _
                          " Result = '상세결과조회', " & vbCrLf & _
                          " Input_UID = '" & Trim(txtUID.Text) & "', " & vbCrLf & _
                          " Input_DateTime = TO_DATE('" & GetDateFull & "', 'mm/dd/yyyy hh24:mi:ss') " & vbCrLf & _
                          " Where HID = '117' " & vbCrLf & _
                          " And PID = '" & Trim(sPID) & "' " & vbCrLf & _
                          " And ReceNo = '" & Trim(sReceNo) & "' " & vbCrLf & _
                          " And SpecimenID = '" & Trim(sSpecID) & "' " & vbCrLf & _
                          " And ExamCode = '" & Trim(sExamCode) & "' "
                    res = SendQuery(SQL)
                    
                    If res = -1 Then
                        SaveQuery SQL
                        SetBackColor vasList, iRow, iRow, 1, 1, 255, 0, 0
                        Exit Sub
                    ElseIf res = 0 Then
                        SaveQuery SQL
                        Exit Sub
                    End If
                    
                    If res = 1 Then
                        SetBackColor vasList, iRow, iRow, 1, 1, 202, 255, 112
                    ElseIf res = -1 Then
                        SetBackColor vasList, iRow, iRow, 1, 1, 255, 0, 0
                    End If
                End If
            End If
        End If
    Next iRow
    
'    db_Commit
    
    MsgBox "작업 완료 되었습니다!", vbExclamation
End Sub


Private Sub cmdSearch_Click()
'    frmWorkList.Show
End Sub

Private Sub Command1_Click()
    'Excel Object Library 와 연결합니다.
    Dim xl As New Excel.Application
    Dim xlw As Excel.Workbook
    
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
    Dim sRemark As String
    
    If Trim(txtFile.Text) = "" Then
        txtFile.SetFocus
        Exit Sub
    End If
    
    ClearSpread vasTemp
    
    '해당 Excel 파일을 연다.****************
    Set xlw = xl.Workbooks.Open(Trim(txtFile.Text))
    
    '가져올 데이터를 포함하고있는 Excel Sheet 를 선택한다.
    xlw.Sheets(Trim(txtSheet.Text)).Select

    i = 1
    j = 1
    k = 1
    
    For i = CLng(txtRow) To CLng(txtRow2)
        For j = CLng(txtCol) To CLng(txtCol2)

'2006/03/14 이상은
'            'If xlw.Application.Cells(i, 1).Value = "" Then
'            '2005/03/22 이상은
'            'If Trim(xlw.Application.Cells(i, 1).Value) = "" Then
'            If Trim(xlw.Application.Cells(i, 2).Value) = "" Then
'                Exit For
'            Else
                SetText vasTemp, xlw.Application.Cells(i, j).Value, k, j
'            End If
        Next j

        k = k + 1
        
        If ProgressBar1.Value >= ProgressBar1.Max Then
            ProgressBar1.Value = 0
            ProgressBar1.Enabled = False
        Else
            ProgressBar1.Value = ProgressBar1.Value + 1
        End If
    Next i
          
    ' Close worksheet without save changes.
'    xlw.Close False
'
'    Set xlw = Nothing
'    Set xl = Nothing
    
    '비고사항*******************************
'    sRemark = "Remark"
'    xlw.Sheets(sRemark).Select
'
'    i = 1
'    j = 1
'    k = 1
'
'    For i = CLng(txtRow) To CLng(txtRow2)
'        For j = CLng(txtCol) To CLng(txtCol2)
'            SetText vasRemark, xlw.Application.Cells(i, j).Value, k, j
'        Next j
'
'        k = k + 1
'
'        If ProgressBar1.Value >= ProgressBar1.Max Then
'            ProgressBar1.Value = 0
'            ProgressBar1.Enabled = False
'        Else
'            ProgressBar1.Value = ProgressBar1.Value + 1
'        End If
'    Next i
          
    ' Close worksheet without save changes.
    xlw.Close False
    
    Set xlw = Nothing
    Set xl = Nothing

    ProgressBar1.Value = 0
    
    '데이터 불러오기
    cmdDisplay_Click

End Sub

Private Sub cmdExit_Click()
    If gConnect = True Then
        DisConnect
        gConnect = False
    End If
    
    Call KillProcess("SCL결과연계.exe")
    
    End
End Sub

Private Sub Form_Load()
    '서버접속
'    gConnect = False
'
'    If gConnect = False Then
'        If Connect = False Then
'            Exit Sub
'        Else
'            gConnect = True
'        End If
'    End If
'
'    mvbFrm.Mvb1.MServer = "CN_IPTCP:211.57.171.3[6001]"
    
    ProgressBar1.Enabled = True
    
    'txtSheet.Text = ""
End Sub

Private Sub Form_Terminate()
    If gConnect = True Then
        DisConnect
        gConnect = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If gConnect = True Then
        DisConnect
        gConnect = False
    End If
    
    Call KillProcess("SCL결과연계.exe")
    
    End
End Sub


Private Sub txtCol_GotFocus()
    SelectFocus txtCol
End Sub

Private Sub txtCol_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtRow.SetFocus
    End If
End Sub

Private Sub txtCol2_GotFocus()
    SelectFocus txtCol2
End Sub

Private Sub txtCol2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtRow2.SetFocus
    End If
End Sub

Private Sub txtFile_DblClick()
Dim sTmp As String
Dim iPos As Integer

    CommonDialog1.Filter = "Excel Files (*.xls)|*.xls|All (*.*)|*.*"
    CommonDialog1.ShowOpen
    txtFile.Text = CommonDialog1.Filename
    
    If gHosInfo.HID = "117" Then    '문경제일병원
        '2004/09/22 이상은 수정=================
        sTmp = Dir(txtFile.Text, vbDirectory)
        iPos = InStr(1, sTmp, ".")
        txtSheet.Text = Mid(sTmp, 1, iPos - 1)
        '=======================================
    ElseIf gHosInfo.HID = "117" Then    '해동병원
        txtSheet.Text = "Result"
    End If
    
    txtSheet.SetFocus
End Sub

Private Sub txtFile_GotFocus()
    SelectFocus txtFile
End Sub

Private Sub txtFile_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtSheet.SetFocus
    End If
End Sub

Private Sub txtRow_GotFocus()
    SelectFocus txtRow
End Sub

Private Sub txtRow_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtCol2.SetFocus
    End If
End Sub

Private Sub txtRow2_GotFocus()
    SelectFocus txtRow2
End Sub

Private Sub txtRow2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        Command1.SetFocus
    End If
End Sub

Private Sub txtSheet_GotFocus()
    SelectFocus txtSheet
End Sub

Private Sub txtSheet_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtCol.SetFocus
    End If
End Sub

Private Sub txtUID_GotFocus()
    SelectFocus txtUID
End Sub

Private Sub vasList_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    Dim iRow As Integer
    Dim i As Integer
    Dim lRow, lCol As Long
    Dim lr, lc As Long
    
    Dim sTmp As String
    
'    bBlockSelected = True
'    lRow1 = BlockRow
'    lCol1 = BlockCol
'    lRow2 = BlockRow2
'    lCol2 = BlockCol2
'
'    lRow = vasList.ActiveRow
'    lCol = vasList.ActiveCol
'
'    If bBlockSelected = True Then
'        sTmp = GetText(vasList, lRow, lCol)
'
'        If lRow1 = -1 Or lRow2 = -1 Then
'            lRow1 = 1
'            lRow2 = vasList.DataRowCnt
'        End If
'        If lCol1 = -1 Or lCol2 = -1 Then
'            lCol1 = 1
'            lCol2 = vasList.DataColCnt
'        End If
'
'        For lr = lRow1 To lRow2
'            For lc = lCol1 To lCol2
'                SetText vasList, sTmp, lr, lc
'            Next lc
'        Next lr
'
'        vasActiveCell vasList, lRow2, lCol2
'        bBlockSelected = False
'        vasList.BlockMode = False
'    End If
End Sub

Private Sub vasList_DblClick(ByVal col As Long, ByVal Row As Long)
    'Sorting
    '체크버튼, 접수일자, 검체번호, 챠트번호, 환자이름, 주민번호
    
    If Row = 0 Then
        Select Case col
        Case 2  '접수일자
            vasSort vasList, 2, 3, 4, 5, 6
        Case 3  '검체번호
            vasSort vasList, 3, 4, 5, 6, 2
        Case 4  '챠트번호
            vasSort vasList, 4, 5, 6, 3, 2
        Case 5  '환자이름
            vasSort vasList, 5, 6, 4, 3, 2
        End Select
    End If
End Sub

Private Sub vasList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim iRow As Integer
Dim iCol As Integer

    iRow = vasList.ActiveRow
    iCol = 3
    
    If KeyCode = vbKeyReturn Then
        If iCol = 3 Then
            vasActiveCell vasList, iRow + 1, 3
            vasList.SetFocus
        End If
    End If
    
End Sub
