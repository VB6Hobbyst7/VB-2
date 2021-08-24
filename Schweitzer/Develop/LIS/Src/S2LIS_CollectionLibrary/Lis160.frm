VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frm160WardBarReprint 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   ClientHeight    =   9210
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   14535
   ControlBox      =   0   'False
   Icon            =   "Lis160.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9210
   ScaleWidth      =   14535
   ShowInTaskbar   =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   300
      Index           =   1
      Left            =   6240
      TabIndex        =   22
      Top             =   45
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   529
      BackColor       =   8388608
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "검색조건"
      LeftGab         =   100
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00E0E0E0&
      Caption         =   "검색"
      Height          =   555
      Left            =   75
      Style           =   1  '그래픽
      TabIndex        =   37
      Top             =   375
      Width           =   3915
   End
   Begin MSComCtl2.DTPicker dtpColDt 
      Height          =   270
      Left            =   2550
      TabIndex        =   36
      Top             =   60
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   476
      _Version        =   393216
      Format          =   61407233
      CurrentDate     =   37505
   End
   Begin MedControls1.LisLabel lblWardId 
      Height          =   240
      Left            =   1335
      TabIndex        =   34
      Top             =   90
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   423
      BackColor       =   8388608
      ForeColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      AutoSize        =   -1  'True
      Caption         =   "병동없음"
   End
   Begin VB.CommandButton cmdWardHelp 
      BackColor       =   &H00F7FDFD&
      Caption         =   "▼"
      Height          =   225
      Left            =   1080
      Style           =   1  '그래픽
      TabIndex        =   33
      Top             =   90
      Width           =   255
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   300
      Index           =   2
      Left            =   12705
      TabIndex        =   23
      Top             =   45
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   529
      BackColor       =   8388608
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "출력장수"
      LeftGab         =   100
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   300
      Index           =   0
      Left            =   4005
      TabIndex        =   21
      Top             =   45
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   529
      BackColor       =   8388608
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "처방구분"
      LeftGab         =   100
   End
   Begin VB.CommandButton cmdReprint 
      BackColor       =   &H00E0E0E0&
      Caption         =   "재출력(&S)"
      Height          =   510
      Left            =   10500
      Style           =   1  '그래픽
      TabIndex        =   4
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00E0E0E0&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   13140
      Style           =   1  '그래픽
      TabIndex        =   6
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00E0E0E0&
      Caption         =   "화면지움(&C)"
      Height          =   510
      Left            =   11820
      Style           =   1  '그래픽
      TabIndex        =   5
      Top             =   8535
      Width           =   1320
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DBE6E6&
      Height          =   1050
      Left            =   12705
      TabIndex        =   8
      Top             =   285
      Width           =   1740
      Begin VB.TextBox txtLabelCnt 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   285
         TabIndex        =   2
         Text            =   "1"
         Top             =   450
         Width           =   690
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   300
         Left            =   975
         TabIndex        =   9
         Top             =   435
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         BuddyControl    =   "txtLabelCnt"
         BuddyDispid     =   196615
         OrigLeft        =   3840
         OrigTop         =   330
         OrigRight       =   4080
         OrigBottom      =   645
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label Label3 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "장"
         Height          =   180
         Left            =   1275
         TabIndex        =   10
         Tag             =   "151"
         Top             =   510
         Width           =   195
      End
   End
   Begin VB.CheckBox chkSelAll 
      BackColor       =   &H00E0E0E0&
      Caption         =   "전체선택(&A)"
      ForeColor       =   &H00553755&
      Height          =   255
      Left            =   4125
      TabIndex        =   7
      Top             =   1395
      Width           =   1350
   End
   Begin FPSpread.vaSpread tblOrdSheet 
      Height          =   6780
      Left            =   4020
      TabIndex        =   3
      Tag             =   "10114"
      Top             =   1695
      Width           =   10440
      _Version        =   196608
      _ExtentX        =   18415
      _ExtentY        =   11959
      _StockProps     =   64
      AutoCalc        =   0   'False
      AutoClipboard   =   0   'False
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      EditEnterAction =   5
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GridColor       =   14737632
      MaxCols         =   28
      ProcessTab      =   -1  'True
      Protect         =   0   'False
      ScrollBars      =   2
      ShadowColor     =   14737632
      ShadowDark      =   12632256
      ShadowText      =   0
      SpreadDesigner  =   "Lis160.frx":08CA
      StartingColNumber=   2
      VirtualRows     =   24
      VisibleCols     =   5
      VisibleRows     =   500
   End
   Begin MSComctlLib.ListView lvwPtList 
      Height          =   7545
      Left            =   75
      TabIndex        =   27
      Top             =   930
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   13309
      View            =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16775406
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움체"
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
      BackColor       =   &H00DBE6E6&
      Height          =   555
      Left            =   30
      TabIndex        =   28
      Tag             =   "136"
      Top             =   1065
      Visible         =   0   'False
      Width           =   3795
      Begin VB.TextBox txtSearchKey 
         Height          =   300
         Left            =   120
         MaxLength       =   10
         TabIndex        =   31
         Top             =   165
         Width           =   1350
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
         Left            =   2115
         TabIndex        =   30
         Tag             =   "15305"
         Top             =   210
         Value           =   -1  'True
         Width           =   825
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
         Left            =   1605
         TabIndex        =   29
         Tag             =   "15304"
         Top             =   210
         Width           =   510
      End
      Begin VB.Label lblReset 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         BackStyle       =   0  '투명
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
         Left            =   3045
         MouseIcon       =   "Lis160.frx":4C40
         MousePointer    =   99  '사용자 정의
         TabIndex        =   32
         Top             =   210
         Width           =   495
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H00808080&
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  '단색
         Height          =   285
         Index           =   1
         Left            =   2940
         Shape           =   4  '둥근 사각형
         Top             =   180
         Width           =   675
      End
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   315
      Index           =   3
      Left            =   75
      TabIndex        =   35
      Top             =   45
      Width           =   3900
      _ExtentX        =   6879
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "환자검색"
      LeftGab         =   100
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DBE6E6&
      Height          =   420
      Left            =   4005
      TabIndex        =   38
      Top             =   270
      Width           =   2220
      Begin VB.OptionButton optDiv 
         BackColor       =   &H00DBE6E6&
         Caption         =   "검체별"
         Height          =   255
         Index           =   1
         Left            =   1140
         TabIndex        =   40
         Top             =   120
         Width           =   930
      End
      Begin VB.OptionButton optDiv 
         BackColor       =   &H00DBE6E6&
         Caption         =   "환자별"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   39
         Top             =   120
         Value           =   -1  'True
         Width           =   930
      End
   End
   Begin VB.Frame fraPt 
      BackColor       =   &H00DBE6E6&
      Height          =   720
      Left            =   4005
      TabIndex        =   11
      Top             =   600
      Width           =   2220
      Begin VB.OptionButton optSearchKey 
         BackColor       =   &H00E0E0E0&
         Caption         =   "전체"
         ForeColor       =   &H004A4189&
         Height          =   465
         Index           =   0
         Left            =   105
         Style           =   1  '그래픽
         TabIndex        =   19
         Top             =   180
         Width           =   630
      End
      Begin VB.OptionButton optSearchKey 
         BackColor       =   &H00DBE6E6&
         Caption         =   "혈액은행"
         Height          =   180
         Index           =   2
         Left            =   855
         TabIndex        =   24
         Tag             =   "B"
         Top             =   435
         Width           =   1170
      End
      Begin VB.OptionButton optSearchKey 
         BackColor       =   &H00DBE6E6&
         Caption         =   "임상병리"
         Height          =   270
         Index           =   1
         Left            =   855
         TabIndex        =   12
         Tag             =   "L"
         Top             =   165
         Width           =   1170
      End
   End
   Begin VB.Frame fraSearchKey 
      BackColor       =   &H00DBE6E6&
      Height          =   1035
      Index           =   1
      Left            =   6240
      TabIndex        =   41
      Top             =   285
      Visible         =   0   'False
      Width           =   6450
      Begin VB.TextBox txtAccNo 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   5685
         TabIndex        =   45
         Top             =   225
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.TextBox txtAccDt 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   4410
         TabIndex        =   44
         Top             =   225
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.TextBox txtWorkArea 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   3585
         TabIndex        =   43
         Top             =   225
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.TextBox txtSpcNo 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1275
         TabIndex        =   42
         Top             =   225
         Width           =   2025
      End
      Begin MedControls1.LisLabel lblPtNm1 
         Height          =   330
         Left            =   1275
         TabIndex        =   46
         Top             =   600
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   582
         BackColor       =   15597309
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
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
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   345
         Index           =   3
         Left            =   195
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   210
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   609
         BackColor       =   14411494
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   "검체번호"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   330
         Index           =   4
         Left            =   195
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   600
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   582
         BackColor       =   14411494
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   "성    명"
         Appearance      =   0
      End
      Begin VB.Line Line1 
         Visible         =   0   'False
         X1              =   4215
         X2              =   4365
         Y1              =   390
         Y2              =   390
      End
      Begin VB.Line Line2 
         Visible         =   0   'False
         X1              =   5505
         X2              =   5655
         Y1              =   390
         Y2              =   390
      End
   End
   Begin VB.Frame fraSearchKey 
      BackColor       =   &H00DBE6E6&
      Height          =   1050
      Index           =   0
      Left            =   6240
      TabIndex        =   13
      Top             =   285
      Width           =   6450
      Begin VB.OptionButton optDuration 
         BackColor       =   &H00DBE6E6&
         Caption         =   "기간제한없슴"
         Height          =   285
         Index           =   1
         Left            =   4575
         TabIndex        =   26
         Top             =   240
         Width           =   1410
      End
      Begin VB.OptionButton optDuration 
         BackColor       =   &H00DBE6E6&
         Caption         =   "최근 1개월"
         Height          =   285
         Index           =   0
         Left            =   3300
         TabIndex        =   25
         Top             =   240
         Width           =   1170
      End
      Begin VB.ComboBox cboOrdDate 
         Height          =   300
         ItemData        =   "Lis160.frx":550A
         Left            =   4080
         List            =   "Lis160.frx":5514
         Style           =   2  '드롭다운 목록
         TabIndex        =   1
         Top             =   630
         Width           =   1860
      End
      Begin VB.TextBox txtPtId 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1095
         TabIndex        =   0
         Top             =   210
         Width           =   1935
      End
      Begin MedControls1.LisLabel lblPtNm 
         Height          =   300
         Left            =   1095
         TabIndex        =   14
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   529
         BackColor       =   15597309
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
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
      Begin VB.Label lblPtId 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "환자 ID"
         Height          =   180
         Left            =   375
         TabIndex        =   18
         Tag             =   "105"
         Top             =   285
         Width           =   585
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "성    명"
         Height          =   180
         Left            =   360
         TabIndex        =   17
         Tag             =   "103"
         Top             =   675
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "처 방 일"
         Height          =   180
         Left            =   3345
         TabIndex        =   16
         Tag             =   "105"
         Top             =   690
         Width           =   660
      End
      Begin VB.Label lblOrdDtCnt 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "1"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   6015
         TabIndex        =   15
         Top             =   705
         Width           =   90
      End
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00808080&
      FillColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   4005
      Shape           =   4  '둥근 사각형
      Top             =   1365
      Width           =   2205
   End
   Begin VB.Label lblMessage 
      BackColor       =   &H00DBE6E6&
      BackStyle       =   0  '투명
      Caption         =   "☞ 처방내역을 검색중입니다..."
      ForeColor       =   &H00553755&
      Height          =   270
      Left            =   6315
      TabIndex        =   20
      Top             =   1425
      Width           =   8145
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00CCFFFF&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00808080&
      Height          =   300
      Index           =   0
      Left            =   6240
      Shape           =   4  '둥근 사각형
      Top             =   1365
      Width           =   8205
   End
End
Attribute VB_Name = "frm160WardBarReprint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private MyPatient   As clsPatient
Private MySql       As clsLISSqlStatement
Private objSQL      As clsLISSqlReview

Private mvarWardId  As String
Private OrdFg       As Boolean
Private ClearFg     As Boolean
Private SelFg       As Boolean
Private MsgFg       As Boolean
Private IsFirst     As Boolean
Private PtFg        As Boolean

Public Event LastFormUnload()
Public Event ThisFormUnload()

'WardId
Public Property Let WardId(ByVal vData As String)
    mvarWardId = vData
End Property

Public Property Get WardId() As String
    WardId = mvarWardId
End Property

Private Sub cboOrdDate_Click()

    If txtPtId.Text = "" Then
       txtPtId.SetFocus
       Exit Sub
    End If
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    If Screen.ActiveControl.Name = cmdExit.Name Then Exit Sub
    If Screen.ActiveControl.Name = cmdClear.Name Then Exit Sub
    If Screen.ActiveControl.Name = optSearchKey(0).Name Then Exit Sub
      
    MouseRunning
    
    lblMessage.Caption = "▶  " & lblPtNm.Caption & " 님의 채취내역을 조회중입니다.."
    Call DisplayOrder
    lblMessage.Caption = ""
    
    MouseDefault
    
    cmdReprint.Enabled = True
    If OrdFg Then
        tblOrdSheet.SetFocus
    Else
        cmdReprint.Enabled = False
        txtPtId.SetFocus
        Call txtPtId_GotFocus
    End If

End Sub

Private Sub chkSelAll_Click()
    Dim i As Integer
    
    SelFg = True
        With tblOrdSheet
        For i = 1 To .DataRowCnt
            .Row = i
            .Col = 1
            .Value = chkSelAll.Value
        Next
    End With
    SelFg = False
 
End Sub

Private Sub cmdClear_Click()
    Call ClearRtn
    
    txtPtId.Text = ""
    txtSpcNo.Text = ""
    If optDiv(0).Value Then
        txtPtId.SetFocus
    Else
        txtSpcNo.SetFocus
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
    
    Set MySql = Nothing
    Set objSQL = Nothing
    Set MyPatient = Nothing
    If IsLastForm Then RaiseEvent LastFormUnload
    RaiseEvent ThisFormUnload
End Sub
Private Function BarPrint_Check() As Boolean
    Dim ii As Integer
    
    With tblOrdSheet
        For ii = 1 To .DataRowCnt
            .Row = ii: .Col = 1
            If .Value = 1 Then
                BarPrint_Check = True
                Exit For
            End If
        Next
    End With
End Function
Private Sub cmdReprint_Click()

    Dim objBAR              As clsBarcode
    Dim tmpLabNo           As Variant
    Dim TestNames          As String
    Dim BarBuffer(0 To 15) As String
    Dim strStatFg          As String
    Dim strSpcNo           As String
    Dim AccFg              As Boolean
    Dim FzFg               As Boolean
    Dim i                  As Long
    
    
    If BarPrint_Check = False Then
        MsgBox "출력대상을 선택한후 재출력 버튼을 클릭하세요.", vbInformation + vbOKOnly, "출력대상선택"
        Exit Sub
    End If
    
    Set objBAR = New clsBarcode
'    Set objBAR.MyDB = dbconn
    Set objBAR.TableInfo = New clsTables
    Set objBAR.FieldInfo = New clsFields
    
    TestNames = ""
    
    Screen.MousePointer = vbArrowHourglass
    lblMessage.Caption = " Barcode Label을 출력중입니다."
    
    With tblOrdSheet
        For i = 1 To .DataRowCnt
            .Row = i
            .Col = 1
            If .Value = 1 Then
                Call .GetText(7, i + 1, tmpLabNo)
                .Col = 7
                If .Value <> tmpLabNo Then
                    Erase BarBuffer
                    
                    .Col = 18:  TestNames = TestNames & .Value & ","
                    
                    .Col = 25:  BarBuffer(0) = .Value                   '처방구분
               
                    .Col = 20:
                        If P_ApplyBuildingInfo Then
                            If BarBuffer(0) = APS_ORDDIV Then
                                BarBuffer(1) = APSName
                            Else
                                BarBuffer(1) = Mid(.Value, 1, 2)        '건물명
                            End If
                        Else
                            Select Case BarBuffer(0)
                                Case LIS_ORDDIV: BarBuffer(1) = LABName
                                Case BBS_ORDDIV: BarBuffer(1) = BBSName
                                Case APS_ORDDIV: BarBuffer(1) = APSName
                            End Select
                        End If
                        
                    .Col = 13:
                        Select Case BarBuffer(0)
                            Case LIS_ORDDIV: BarBuffer(2) = .Value                   'WorkArea
                            Case BBS_ORDDIV: BarBuffer(2) = BBSBarNm
                            Case APS_ORDDIV: BarBuffer(2) = APSBarNm
                        End Select
                            
                    .Col = 16:  BarBuffer(3) = Mid(.Value, 3)           'AccDt
                    .Col = 14:  Select Case BarBuffer(0)
                                Case BBS_ORDDIV:
                                    .Col = 7
                                    BarBuffer(4) = Format(.Value, String(11, "@"))
                                Case Else:
                                    .Col = 14
                                    BarBuffer(4) = IIf(.Value = "0", "", Format(.Value, String(4, "@")))    'SpcNo
                                End Select
                    .Col = 19:  BarBuffer(5) = .Value                   'SpcNo
                                BarBuffer(6) = MyPatient.Ptid           '환자ID
                                BarBuffer(7) = MyPatient.PtNm 'Mid(MyPatient.PtNm, 1, 3)
                                
                    .Col = 12:  BarBuffer(8) = .Value                   '검체명
                    .Col = 15:  BarBuffer(9) = .Value                   '보관코드
                    .Col = 17:
                                If BarBuffer(5) = strSpcNo Then
                                    BarBuffer(10) = IIf(strStatFg = "1", strStatFg, .Value) 'StatFg 구분
                                Else
                                    BarBuffer(10) = .Value
                                End If
                    .Col = 27:
                                If .Value = "" Then
                                    .Col = 22: BarBuffer(11) = .Value   '진료과코드
                                    If BarBuffer(11) <> "" Then
                                        .Col = 21
                                        If .Value <> "" Then
                                            BarBuffer(11) = BarBuffer(11) & "/" & .Value
                                        End If
                                    Else
                                        .Col = 21
                                        If .Value <> "" Then
                                            BarBuffer(11) = .Value
                                        End If
                                    End If
                                Else
                                    BarBuffer(11) = .Value              '병동ID
                                    .Col = 21
                                    If .Value <> "" Then
                                        BarBuffer(11) = BarBuffer(11) & "/" & .Value
                                    End If
                                End If
                    .Col = 8:   BarBuffer(12) = Mid(.Value, 5, 2) & "/" & _
                                                Mid(.Value, 7, 2)       '처방일
                    .Col = 24:  BarBuffer(13) = .Value                  '희망채혈일시
                                BarBuffer(14) = TestNames               '검사명
                                BarBuffer(15) = txtLabelCnt.Text        '라벨출력장수
                    .Col = 23:
                                AccFg = IIf(Val(.Value) >= 2, True, False)
                    .Col = 26:
                                FzFg = IIf(.Value = "1", True, False)
                    
                    Call objBAR.Label_PrintOut( _
                                                BarBuffer(1), BarBuffer(2), BarBuffer(3), _
                                                BarBuffer(4), BarBuffer(5), BarBuffer(6), BarBuffer(7), _
                                                BarBuffer(8), BarBuffer(9), BarBuffer(10), BarBuffer(11), _
                                                BarBuffer(12), BarBuffer(13), BarBuffer(14), BarBuffer(15), _
                                                AccFg, FzFg)

'                    Call medSleep(2000)
                    TestNames = ""
                Else
                    .Col = 17
                        If .Value = "1" Then
                            .Col = 19: strSpcNo = .Value
                            strStatFg = "1"
                        End If
                    .Col = 18
                    TestNames = TestNames & .Value & ","
                End If
            End If
        Next
    End With

    Call ClearRtn
    
    Screen.MousePointer = vbDefault
    lblMessage.Caption = ""
    If optDiv(0).Value Then
        txtPtId.Text = ""
        txtPtId.SetFocus
    Else
        txtSpcNo.Text = ""
        lblPtNm1.Caption = ""
        
    End If
    Set objBAR = Nothing
End Sub

Private Sub cmdSearch_Click()
    Call txtSearchKey_KeyPress(vbKeyReturn)
    dtpColDt.SetFocus
End Sub

Private Sub cmdWardHelp_Click()

    Dim objDeptHelp As New clsPopUpList
'    Dim objWard As New clsBasisData
    
    lvwPtList.ListItems.Clear
    mvarWardId = "": lblWardId.Caption = ""
    With objDeptHelp
        .Connection = DBConn
        .FormCaption = "병동리스트"
        .ColumnHeaderText = "병동;병동명"
        .LoadPopUp GetSQLWard ', 2000, 1500  ', ObjLISComCode.WardId
        If .SelectedString <> "" Then
            mvarWardId = medGetP(.SelectedString, 1, ";")
            lblWardId.Caption = mvarWardId
        End If
    End With
    Set objDeptHelp = Nothing
'    Set objWard = Nothing

End Sub

Private Sub dtpColDt_CloseUp()
    Call txtSearchKey_KeyPress(vbKeyReturn)
    cmdSearch.SetFocus
End Sub

Private Sub dtpColDt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdSearch.SetFocus
    End If
End Sub

Private Sub Form_Activate()
    If Not IsFirst Then Exit Sub
    IsFirst = False
    
    PtFg = False
    SelFg = False
    cboOrdDate.Clear
    optSearchKey(0).Value = True
    optDuration(0).Value = True
    lblOrdDtCnt.Caption = ""
    ClearFg = True
    txtSearchKey.Text = ""
    If Trim(gWardId) <> "" Then
        lblWardId.Caption = Trim(gWardId)
    Else
        lblWardId.Caption = "병동없음"
    End If
    medInitLvwHead lvwPtList, "환자ID,환자성명,주민등록번호,생년월일,성별/나이", _
                       "150,150,1000,300,100"
    '** Default 병동기준 조회
    Call txtSearchKey_KeyPress(vbKeyReturn)
End Sub

Private Sub Form_Load()
    IsFirst = True
    
    dtpColDt.Value = Format(GetSystemDate, "yyyy/mm/dd")
    Set MyPatient = New clsPatient
    Set MySql = New clsLISSqlStatement
    Set objSQL = New clsLISSqlReview
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call ICSPatientMark
    Set MySql = Nothing
    Set objSQL = Nothing
    Set MyPatient = Nothing
End Sub

Private Sub lvwPtList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static lngOrder As Long
    optDiv(0).Value = True
    With lvwPtList
        lngOrder = (lngOrder + 1) Mod 2
        .SortKey = ColumnHeader.Index - 1
        .SortOrder = Choose(lngOrder + 1, lvwAscending, lvwDescending)
        .Sorted = True
    End With
End Sub

Private Sub lvwPtList_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    '환자정보 Display
    If Item = "" Then Exit Sub
    DoEvents
    With Item
        txtPtId.Text = .Text                '환자ID
        Call txtPtId_LostFocus
    End With
    
End Sub

Private Sub optDiv_Click(Index As Integer)
    fraSearchKey(0).Visible = False
    fraSearchKey(1).Visible = False
    If Index = 0 Then
        fraPt.Enabled = True
        fraSearchKey(0).Visible = True
        txtPtId.Text = ""
        txtPtId.SetFocus
    Else
        fraPt.Enabled = False
        fraSearchKey(1).Visible = True
        txtWorkArea.Text = "": txtAccDt.Text = "": txtAccNo.Text = ""
        txtSpcNo.Text = "": txtSpcNo.SetFocus
    End If
    Call medClearTable(tblOrdSheet)
    LisLabel2(1).ZOrder 0
End Sub

Private Sub txtSpcNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If txtSpcNo.Text = "" Then Exit Sub
        Call GetSpcDataQuery
    End If
End Sub

Private Sub GetSpcDataQuery()
    Dim strSpcYY As String
    Dim strSpcNo As String
    Dim strWA    As String
    Dim strAccdt As String
    Dim strAccno As String
    Dim ii       As Integer
    
    strSpcYY = Mid(txtSpcNo.Text, 1, 2)
    strSpcNo = Mid(txtSpcNo.Text, 3)
    txtWorkArea.Text = "": txtAccDt.Text = "": txtAccNo.Text = ""
    lblPtNm1.Caption = ""
    Call MySql.GetLabNo(strSpcYY, strSpcNo, strWA, strAccdt, strAccno)
    If strWA = "" Then
        MsgBox "해당검체에 대한 정보가 없거나 임상병리 처방이 아닙니다." & _
               "확인후 출력하십시요.", vbInformation + vbOKOnly, "Info"
        txtSpcNo.Text = ""
        If txtSpcNo.Enabled Then txtSpcNo.SetFocus
        Exit Sub
    End If
    
    txtWorkArea.Text = strWA
    txtAccDt.Text = strAccdt
    txtAccNo.Text = strAccno
    
    
    Call MouseRunning
    lblMessage.Caption = "접수번호 " & txtWorkArea.Text & "-" & txtAccDt.Text & "-" & txtAccNo.Text & " 를 조회중입니다.."
    Call DisplayOrder
    lblMessage.Caption = ""
    With tblOrdSheet
        For ii = 1 To .DataRowCnt
            .Row = ii: .Col = 1: .Value = 1
        Next
    End With
    cmdReprint.Enabled = True
    Call cmdReprint_Click
    txtSpcNo.Text = "": txtSpcNo.SetFocus
    cmdReprint.Enabled = False
    Call MouseDefault

End Sub


Private Sub optSearchKey_Click(Index As Integer)
   
    On Error GoTo Err_Trap
    
    optSearchKey(0).ForeColor = vbBlue
    optSearchKey(1).ForeColor = vbBlack
    fraSearchKey(0).Visible = True
    fraSearchKey(1).Visible = False

Err_Trap:
    Call ClearRtn
    
    If txtPtId.Text = "" Then txtPtId.SetFocus
   
End Sub

Private Sub tblOrdSheet_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    Dim SvLabNo     As String
    Dim i           As Long
    Dim SvButtonVal As Integer
    
    If Col <> 1 Then Exit Sub
    If SelFg Then Exit Sub
    
    With tblOrdSheet
        .Row = Row
        .Col = 1:  SvButtonVal = .Value
        .Col = 7:  SvLabNo = Trim(.Value)
        For i = 1 To .DataRowCnt
            If i <> Row Then
                .Row = i
                .Col = 7
                If Trim(.Value) = SvLabNo Then
                    .Col = 1
                    If .Value <> SvButtonVal Then .Value = SvButtonVal
                End If
            End If
        Next
    End With
   
End Sub

Private Sub cboOrdDate_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
       tblOrdSheet.SetFocus
    End If

End Sub


'% 환자ID가 변경되면 화면Clear
Private Sub txtPtId_Change()

    If Not ClearFg Then Call ClearRtn

End Sub

'% 환자 ID
Private Sub txtPtId_GotFocus()
    With txtPtId
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

'% 환자정보 검색
Private Sub txtPtId_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        Call ICSPatientMark(txtPtId.Text, enICSNum.LIS_ALL)
        cboOrdDate.SetFocus
    End If
End Sub

Private Sub txtPtId_LostFocus()
      
    If txtPtId.Text = "" Then Exit Sub
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    If Screen.ActiveControl.Name = optSearchKey(0).Name Then Exit Sub
    If Screen.ActiveControl.Name = cmdExit.Name Then Exit Sub
    If Screen.ActiveControl.Name = cmdClear.Name Then Exit Sub

    If IsNumeric(txtPtId.Text) Then
        txtPtId.Text = Format(txtPtId.Text, P_PatientIdFormat)
    End If
    
    Set MyPatient = Nothing
    Set MyPatient = New clsPatient
    
    With MyPatient
'        Call .ClearData   '클래스 내 변수 초기화
        If .GETPatient(txtPtId.Text) Then
            lblPtNm.Caption = .PtNm         '성명
            PtFg = True
            ClearFg = False
            If Not LoadOrderDate Then
                MsgFg = True
                MsgBox MyPatient.PtNm & " 님의 처방내역이 없습니다"
                txtPtId.Text = ""
                txtPtId.SetFocus
                MsgFg = False
                Call txtPtId_GotFocus
                Exit Sub
            End If
        Else
            MsgBox "등록되지 않은 환자ID입니다.. 다시 입력하세요.."
            txtPtId.Text = ""
            ClearFg = True
            PtFg = False
            txtPtId.SetFocus
            Call txtPtId_GotFocus
            Exit Sub
        End If
    End With

End Sub


'% 검색한 처방을 테이블에 디스플레이 한다.
Private Sub DisplayOrder()
    
    Dim Rs          As Recordset
    Dim SqlStmt     As String

    Dim SvOrdNo     As String
    Dim SvSpcNm     As String
    Dim strOrdDiv   As String
    Dim strAccdt    As String
    Dim strColDt    As String
    Dim strStsCd    As String
    Dim strTestDiv  As String
    Dim strStsNm    As String
    Dim strWorkArea As String
    Dim strAccSeq   As String
    Dim strUnit     As String
    Dim lngColor    As Long
    Dim iBtnFg      As Long
    
    Dim i           As Integer
    
    DoEvents
     
    strOrdDiv = GetOrdDiv
         

    If optDiv(0).Value Then
        
        SqlStmt = MySql.SqlWardBarReprint(1, txtPtId.Text, _
                                          Format(cboOrdDate.Text, CS_DateDbFormat), _
                                          strOrdDiv)
    Else
        SqlStmt = MySql.SqlBarReprint(2, txtWorkArea.Text, txtAccDt.Text, txtAccNo.Text)
    
    End If
    Set Rs = New Recordset
    Rs.Open SqlStmt, DBConn
    
    If Rs.EOF Then
        MsgBox MyPatient.PtNm & " 님의 채취내역이 없습니다", vbInformation, "간호사채취"
        txtPtId.Text = ""
        Call ClearRtn
        GoTo NoData
    End If
   
    With tblOrdSheet
      
        .ReDraw = False
        .MaxRows = 0
        If Rs.RecordCount < 20 Then
            .MaxRows = 20
            .Row = Rs.RecordCount + 1
            .Row2 = 20
            .Col = 1: .Col2 = .MaxCols
            .BlockMode = True
            .Lock = True
            .Protect = True
            .BlockMode = False
        Else
            .MaxRows = Rs.RecordCount + 1  '데이타 건수
        End If
        .RowHeight(-1) = 13
      
        'Locking Cells
        .Row = -1
        .Col = 2: .Col2 = .MaxCols
        .BlockMode = True
        .Lock = True
        .Protect = True
        .BlockMode = False
            
'        MyPatient.Ptid = Trim("" & Rs.Fields("PtId").Value)
'        MyPatient.PtNm = getpatientname(Trim("" & Rs.Fields("ptid").Value))
        If optDiv(1).Value Then lblPtNm1.Caption = MyPatient.PtNm
'        MyPatient.WardId = Trim("" & Rs.Fields("HosilId").Value)
      
        For i = 1 To Rs.RecordCount
         
            lblMessage.Caption = lblMessage.Caption & "."
            DoEvents
         
            .Row = i
            .Col = 1: .Value = 0
            
            .Row = i    '**ButtonClicked 이벤트가 발생하여 Row값이 바뀌므로 다시 한번 셋팅.
            
            '** 변경 처방일 => 채혈일 By M.G.Choi
            'RS.Fields("new_coldt").Value
            If strColDt <> Trim("" & Rs.Fields("new_coldt").Value) Then
                .Col = 2: .Value = Format("" & Rs.Fields("new_coldt").Value, CS_DateMask)   '채혈일
                .Col = 3: .Value = Trim("" & Rs.Fields("OrdNo").Value)   '처방번호
                .Col = 5: .Value = Trim("" & Rs.Fields("SpcNm").Value)   '검체
                strColDt = Trim("" & Rs.Fields("new_coldt").Value)
                SvOrdNo = Trim("" & Rs.Fields("OrdNo").Value)            '처방번호
                SvSpcNm = Trim("" & Rs.Fields("SpcNm").Value)            '검체
            End If
            If SvOrdNo <> Trim("" & Rs.Fields("OrdNo").Value) Then
                .Col = 3: .Value = Trim("" & Rs.Fields("OrdNo").Value)   '처방번호
                .Col = 5: .Value = Trim("" & Rs.Fields("SpcNm").Value)   '검체
                SvOrdNo = Trim("" & Rs.Fields("OrdNo").Value)            '처방번호
                SvSpcNm = Trim("" & Rs.Fields("SpcNm").Value)            '검체
            End If
            If SvSpcNm <> Trim("" & Rs.Fields("SpcNm").Value) Then
                .Col = 5: .Value = Trim("" & Rs.Fields("SpcNm").Value)   '검체
                SvSpcNm = Trim("" & Rs.Fields("SpcNm").Value)
            End If
         
            .Col = 4: .Value = Trim("" & Rs.Fields("TestNm").Value)      '처방명
            
             '.ForeColor = &HDF6A3E        '약간 파란색
            Select Case Rs.Fields("orddiv")
                Case "A": .ForeColor = &H5E3F00     '&HDF6A3E     '약간 파란색
                Case "B": .ForeColor = &H496835     '&H6C6181     '&H81815A     '약간녹색
                Case "L": .ForeColor = &H553755
            End Select
            
            .Col = 6: .Value = Choose(Val("" & Rs.Fields("StatFg").Value) + 1, "", "Y")     '응급여부
                         .ForeColor = &HFF&       '빨간색
            Select Case Trim("" & Rs.Fields("OrdDiv").Value)
            Case "A", "L"
                .Col = 7: .Value = Trim("" & Rs.Fields("LabNo").Value)       'LabNo
            Case "B"
                .Col = 7: .Value = Trim("" & Rs.Fields("SpcYy").Value) & "-" & _
                                             Rs.Fields("SpcNo").Value
            End Select
            .Col = 8: .Value = Trim("" & Rs.Fields("OrdDt").Value)       '처방일
            .Col = 9: .Value = Trim("" & Rs.Fields("OrdNo").Value)       '처방번호
            .Col = 10: .Value = Trim("" & Rs.Fields("OrdSeq").Value)     '처방Seq
            .Col = 11: .Value = Trim("" & Rs.Fields("OrdCd").Value)      '검사코드
            .Col = 12: .Value = Trim("" & Rs.Fields("SpcNm").Value)      '검체명
            .Col = 13: .Value = Trim("" & Rs.Fields("WorkArea").Value)   'WorkArea
            strWorkArea = Trim("" & Rs.Fields("WorkArea").Value)
            .Col = 14: .Value = Trim("" & Rs.Fields("AccSeq").Value)     'AccSeq
            strAccSeq = Trim("" & Rs.Fields("AccSeq").Value)
            .Col = 15: .Value = Trim("" & Rs.Fields("StoreCd").Value)    '보관코드
            .Col = 16: .Value = Trim("" & Rs.Fields("AccDt").Value)      'AccDt  채혈일
            strAccdt = Trim("" & Rs.Fields("AccDt").Value)
            .Col = 17: .Value = Trim("" & Rs.Fields("StatFg").Value)     '응급여부
            .Col = 18: .Value = Trim("" & Rs.Fields("AbbrNm5").Value)    '약어명
            .Col = 19: .Value = Trim("" & Rs.Fields("SpcYy").Value) & _
                                Format(Val(Rs.Fields("SpcNo").Value), CS_BarFormat)     '검체번호
            .Col = 20: .Value = IIf(P_ApplyBuildingInfo, Trim("" & Rs.Fields("BuildNm").Value), "") '건물명
            .Col = 21: .Value = Trim("" & Rs.Fields("HosilId").Value)    '호실코드
            .Col = 22: .Value = Trim("" & Rs.Fields("DeptCd").Value)     '진료과코드
            .Col = 23: .Value = Trim("" & Rs.Fields("StsCd").Value)      'status
            strStsCd = Trim("" & Rs.Fields("StsCd").Value)
            .Col = 24: .Value = Mid(Trim("" & Rs.Fields("ReqTm").Value), 1, 2) & ":" & _
                                Mid(Trim("" & Rs.Fields("ReqTm").Value), 3, 2)  '희망채혈일시
            .Col = 25: .Value = Trim("" & Rs.Fields("OrdDiv").Value)     'OrdDiv
            strOrdDiv = Trim("" & Rs.Fields("OrdDiv").Value)
            .Col = 26: .Value = Trim("" & Rs.Fields("FzFg").Value)       '냉동절편구분
            .Col = 27: .Value = Trim("" & Rs.Fields("wardid").Value)     '병동코드
            strTestDiv = Trim("" & Rs.Fields("testdiv").Value)
            strUnit = "0"
            
            '** 추가 [진행상태] Display By M.G.Choi 2002.09.05
            '======================================================================
            .Col = 28
            If Trim("" & Rs.Fields("orddiv").Value) = "B" Then
                Select Case Trim("" & Rs.Fields("StsCd").Value)
                    Case "0":
                        If Trim(.Value) = "*" Then
                            .Value = "처방" & "*"
                        Else
                            .Value = "처방"
                        End If
                    Case "1":
                        If Trim(.Value) = "*" Then
                            .Value = "채취" & "*"
                        Else
                            .Value = "채취"
                        End If
                    Case "2":
                        If Trim(.Value) = "*" Then
                            .Value = "접수" & "*"
                        Else
                            .Value = "접수"
                        End If
                    Case Else
                            If strWorkArea <> "" Then
                                If Trim(.Value) = "*" Then
                                  .Value = BBS_STATUS(strWorkArea, strAccdt, strAccSeq, strUnit) & "*": .ForeColor = DCM_Blue
                                Else
                                    .Value = BBS_STATUS(strWorkArea, strAccdt, strAccSeq, strUnit): .ForeColor = DCM_Blue
                                End If
                            End If
                End Select
            Else
            
                Call GetOrderStatus(strOrdDiv, strStsCd, strTestDiv, _
                                   strStsNm, lngColor, iBtnFg, strWorkArea, strAccdt, strAccSeq, "0")
                If Trim(.Value) = "*" Then
                    .Value = strStsNm & "*": .ForeColor = lngColor
                Else
                    .Value = strStsNm: .ForeColor = lngColor
                End If
            End If
            '======================================================================
            
            Rs.MoveNext
        Next
        .ReDraw = True
      
    End With
    cmdReprint.Enabled = True
    OrdFg = True
    ClearFg = False
   
NoData:
    Set Rs = Nothing
   
End Sub

Private Function BBS_STATUS(ByVal WorkArea As String, ByVal AccDt As String, ByVal AccSeq As String, ByVal unitqty As String) As String
    Dim strtmp As String
    Dim lngA   As Long
    Dim lngAC  As Long
    Dim lngD   As Long
    Dim lngR   As Long
    
    strtmp = objSQL.GetDeliveryCnt(WorkArea, AccDt, AccSeq)
    If strtmp <> "" Then
        lngA = medGetP(strtmp, 1, COL_DIV)
        lngAC = medGetP(strtmp, 2, COL_DIV)
        lngD = medGetP(strtmp, 3, COL_DIV)
        lngR = medGetP(strtmp, 4, COL_DIV)
        
        If unitqty = lngA - lngAC And unitqty = lngD - lngR Then
            BBS_STATUS = "완결"
        ElseIf lngA - lngAC = lngD - lngR Then
            BBS_STATUS = "대기"
        ElseIf unitqty >= lngA - lngAC And lngA - lngAC > lngD - lngR Then
            BBS_STATUS = "준비"
        End If
        
    Else
        BBS_STATUS = "검사중"
    End If
    
End Function

Private Sub GetOrderStatus(ByVal pOrdDiv As String, ByVal pStsCd As String, _
                               ByVal pTestDiv As String, ByRef pStsNm As String, _
                               ByRef pStsColor As Long, ByRef pBttnFg As Long, _
                               ByVal WorkArea As String, ByVal AccDt As String, _
                               ByVal AccSeq As String, ByVal unitqty As String)

    Select Case Trim(pStsCd)
        Case enStsCd.StsCd_LIS_Order:
             pStsNm = STS_LIS_Order:     pStsColor = DCM_Gray: pBttnFg = 1 '회색
        
        Case enStsCd.StsCd_LIS_Collection:
             pStsNm = STS_LIS_HaveSpc:   pStsColor = DCM_Gray: pBttnFg = 1 '회색
        
        Case enStsCd.StsCd_LIS_Accession:
            pBttnFg = 1
            pStsNm = STS_LIS_Access:    pStsColor = DCM_Gray: pBttnFg = 1 '회색
        Case enStsCd.StsCd_LIS_InProcess:
             pStsNm = STS_LIS_Worksheet: pStsColor = DCM_Gray: pBttnFg = 1 '회색
        
        Case enStsCd.StsCd_LIS_MidRst:
            
            pBttnFg = 1
            If pOrdDiv = APS_ORDDIV Then
                pStsNm = STS_LIS_Reading:   pStsColor = DCM_Gray           '회색
            Else
                If pTestDiv = TST_MicTest Then            '미생물검사
                    pStsNm = STS_LIS_MidRst:    pStsColor = DCM_Black      '검정색
                Else: pStsNm = STS_LIS_Partial: pStsColor = DCM_Black      '검정색
                End If
            End If
        Case enStsCd.StsCd_LIS_FinRst:
             pBttnFg = 0: pStsColor = DCM_Black  '검정색
             pStsNm = IIf(pOrdDiv = APS_ORDDIV, STS_LIS_MidRst, _
                      IIf(pTestDiv = TST_MicTest, STS_LIS_FinRst, STS_LIS_Verify))
        
        Case enStsCd.StsCd_LIS_Modify:
             pBttnFg = 0: pStsColor = DCM_Black  '검정색
             pStsNm = IIf(pOrdDiv = APS_ORDDIV, STS_LIS_Verify, STS_LIS_Modify)
        
        Case "7":
             pBttnFg = 0: pStsColor = DCM_Black  '검정색
             pStsNm = STS_LIS_Modify
    End Select

End Sub

Private Function BBS_STATUS_Result(ByVal WorkArea As String, ByVal AccDt As String, ByVal AccSeq As String, ByVal unitqty As String) As String
    BBS_STATUS_Result = "접수"
End Function

Private Function GetOrdDiv() As String
    
    Dim i As Long
    On Error Resume Next
    GetOrdDiv = ""
    For i = 0 To optSearchKey.Count - 1
        If optSearchKey(i).Value Then
            GetOrdDiv = optSearchKey(i).Tag
            Exit For
        End If
    Next
    
End Function


Private Sub ClearRtn()
   
    'optSearchKey(0).Value = True
    With tblOrdSheet
        .Row = -1
        .Col = -1
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
    End With
    txtLabelCnt.Text = "1"
    cboOrdDate.Clear
    lblPtNm.Caption = ""
'    lblPtNm1.Caption = ""
    lblOrdDtCnt.Caption = ""
   
    cmdReprint.Enabled = False
    OrdFg = False
    Set MyPatient = Nothing
    Set MyPatient = New clsPatient
'    Set MyPatient.objDB = dbconn
    
    SelFg = False
    ClearFg = True
    lblMessage.Caption = ""
    chkSelAll.Value = 0
End Sub



Public Function LoadOrderDate() As Boolean
    Dim Rs          As Recordset
    
    MySql.OrderDate = Format(dtpColDt.Value, "yyyymmdd")
    
    Set Rs = New Recordset
    Rs.Open MySql.SqlGetOrdDateForBarprint(txtPtId.Text, GetOrdDiv, optDuration(0).Value), DBConn
    
    If Rs.EOF Then
        LoadOrderDate = False
    Else
        LoadOrderDate = True
        cboOrdDate.Clear
        While (Not Rs.EOF)
            cboOrdDate.AddItem Format(Rs.Fields("orddt").Value, CS_DateMask)
            Rs.MoveNext
        Wend
        If cboOrdDate.ListCount > 1 Then
            lblOrdDtCnt.Caption = CStr(cboOrdDate.ListCount)
        Else
            lblOrdDtCnt.Caption = ""
        End If
        cboOrdDate.ListIndex = 0
    End If
    Set Rs = Nothing

End Function

Public Sub Call_PtId_KeyPress()

    Call txtPtId_KeyPress(vbKeyReturn)

End Sub

Public Sub Call_ToDate_LostFocus()

    Call cboOrdDate_KeyDown(vbKeyReturn, 0)
   
End Sub

Private Sub txtSearchKey_GotFocus()

    With txtSearchKey
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub

'% 환자ID 또는 성명으로 검색 리스트 작성.
Private Sub txtSearchKey_KeyPress(KeyAscii As Integer)
    
    Dim Rs          As Recordset
    Dim itmX        As ListItem
    Dim strPtid     As String
    Dim strWardId   As String
    Dim strColDt    As String
    Dim strSQL      As String
    Dim strJumin1   As String
    Dim strJumin2   As String
    
    Dim lngSearch   As Long
    
    If KeyAscii = vbKeyReturn Then
        lngSearch = IIf(optSort(0).Value, 1, 2) + 4 'True:환자ID, False:환자명
        
        If lngSearch = 1 And Not IsNumeric(txtSearchKey.Text) Then Exit Sub
        
        strWardId = lblWardId.Caption
        If strWardId = "병동없음" Then
            Exit Sub
        End If
        
        strColDt = Format(dtpColDt, "yyyymmdd")
        
        
        strSQL = "SELECT distinct a." & F_PTID & " as ptid, a." & F_PTNM & " as ptnm, " & F_SSN2("a") & " as SSN, " & _
                   F_DOB2("a") & " as DOB , d.coldt, d.coltm " & _
                 "  FROM " & T_HIS001 & " a ," & T_LAB101 & " b ," & T_LAB102 & " c," & T_LAB201 & " d " & _
                 " WHERE " & DBW("d.coldt =", strColDt) & _
                 "   AND " & DBW("d.stscd <", enStsCd.StsCd_LIS_MidRst) & _
                 "   AND " & DBW("b.donefg >", enStsCd.StsCd_LIS_Order) & _
                 "   AND " & DBW("b.orddiv =", LIS_ORDDIV) & _
                 "   AND " & DBW("b.bussdiv =", enBussDiv.BussDiv_InPatient) & _
                 "   AND b.wardid = '" & strWardId & "'" & _
                 "   AND (c.dcfg = '' or c.dcfg is null) " & _
                 "   AND d.workarea= c.workarea " & _
                 "   AND d.accdt = c.accdt " & _
                 "   AND d.accseq = c.accseq " & _
                 "   AND c.ptid = b.ptid " & _
                 "   AND c.orddt = b.orddt " & _
                 "   AND c.ordno = b.ordno " & _
                 "   AND d.ptid = a." & F_PTID

        strSQL = strSQL & " ORDER BY ptid, coldt ,coltm desc " '"   AND a.NAME like '%' " & _

        Set Rs = New Recordset
        Rs.Open strSQL, DBConn
        
        lvwPtList.ListItems.Clear
        If Rs.EOF = False Then
            With lvwPtList
                Do Until Rs.EOF
                    If strPtid <> Rs.Fields("ptid").Value & "" Then
                        Set itmX = .ListItems.Add(, , Rs.Fields("ptid").Value & "")
                        itmX.SubItems(1) = Rs.Fields("ptnm").Value & ""
                        
                        If Rs.Fields("SSN").Value <> "" Then
                            strJumin1 = Mid$(Rs.Fields("SSN").Value & "", 1, 6)
                            strJumin2 = Mid$(Rs.Fields("SSN").Value & "", 7, 7)
                            itmX.SubItems(2) = strJumin1 & "-" & strJumin2
                        End If
                        
                        itmX.SubItems(3) = Format(Rs.Fields("DOB").Value & "", CS_DateLongMask)
                        itmX.SubItems(4) = IIf((Mid(Rs.Fields("ssn").Value & "", 7, 1) Mod 2) = 1, "남", "여")
                        If IsDate(itmX.SubItems(3)) Then
                            itmX.SubItems(4) = itmX.SubItems(4) & " / " & DateDiff("yyyy", itmX.SubItems(3), GetSystemDate)
                        End If
                        
                        strPtid = Rs.Fields("ptid").Value & ""
                    End If
                        
                    If .ListItems.Count >= 1000 Then Rs.MoveLast
                    Rs.MoveNext
                Loop
            End With
        Else
            MsgBox "조건에 맞는 자료가 없습니다. 확인후 검색하세요", vbInformation + vbOKOnly, Me.Caption
            dtpColDt.SetFocus
        End If
    End If

    Set Rs = Nothing
End Sub
