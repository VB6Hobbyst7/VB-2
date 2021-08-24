VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.MDIForm frmMDI 
   BackColor       =   &H00FFFFFF&
   Caption         =   "한국해양조사협회 자료수집현황 모니터링"
   ClientHeight    =   9810
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   20100
   Icon            =   "frmMDI.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows 기본값
   WindowState     =   2  '최대화
   Begin VB.PictureBox picMenu 
      Align           =   1  '위 맞춤
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   20040
      TabIndex        =   0
      Top             =   0
      Width           =   20100
      Begin VB.Timer tmrNow 
         Left            =   13650
         Top             =   90
      End
      Begin VB.CommandButton cmdView 
         BackColor       =   &H00C0FFFF&
         Caption         =   "조위관측소"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   0
         Left            =   150
         Style           =   1  '그래픽
         TabIndex        =   3
         Top             =   60
         Width           =   1575
      End
      Begin VB.CommandButton cmdView 
         BackColor       =   &H00FFFFFF&
         Caption         =   "해양관측부이"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   1
         Left            =   1740
         Style           =   1  '그래픽
         TabIndex        =   2
         Top             =   60
         Width           =   1575
      End
      Begin VB.CommandButton cmdView 
         BackColor       =   &H00FFFFFF&
         Caption         =   "조위관측소자료"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   2
         Left            =   3330
         MaskColor       =   &H00FFFFFF&
         Style           =   1  '그래픽
         TabIndex        =   1
         Top             =   60
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker dtpToday 
         Height          =   345
         Left            =   7530
         TabIndex        =   4
         Top             =   120
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   135987200
         CurrentDate     =   43884
      End
      Begin MSComCtl2.DTPicker dtpTotime 
         Height          =   345
         Left            =   10140
         TabIndex        =   6
         Top             =   120
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   135987202
         CurrentDate     =   43884
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '투명
         Caption         =   "현재시간"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6510
         TabIndex        =   5
         Top             =   180
         Width           =   1005
      End
   End
   Begin VB.Menu mnuMenu01 
      Caption         =   "  파일 "
      Begin VB.Menu mnuHelp01 
         Caption         =   "원격지원(TeamViewer)"
         Visible         =   0   'False
      End
      Begin VB.Menu menuSep01 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu menuMDBSync 
         Caption         =   "MDB Sync"
         Visible         =   0   'False
      End
      Begin VB.Menu menuSep02 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuExit 
         Caption         =   "종료"
      End
   End
   Begin VB.Menu mnuMenu02 
      Caption         =   " 설정 "
      Visible         =   0   'False
      Begin VB.Menu menuUser 
         Caption         =   " 사용자 설정 "
      End
      Begin VB.Menu menuSep001 
         Caption         =   "-"
      End
      Begin VB.Menu menuComp 
         Caption         =   " 고객사 설정 "
      End
      Begin VB.Menu menuSep002 
         Caption         =   "-"
      End
      Begin VB.Menu menuPack 
         Caption         =   " 포장 설정 "
      End
      Begin VB.Menu mnuSep12 
         Caption         =   "-"
      End
      Begin VB.Menu menuRegProd 
         Caption         =   " 제품 마스터 "
      End
      Begin VB.Menu mnuSep26 
         Caption         =   "-"
      End
      Begin VB.Menu menuMastr 
         Caption         =   " 공통코드 설정 "
      End
      Begin VB.Menu mnuSep27 
         Caption         =   "-"
      End
      Begin VB.Menu menuRegComm 
         Caption         =   " 바코드 통신설정"
      End
   End
   Begin VB.Menu menuReg 
      Caption         =   " 등록 "
      Visible         =   0   'False
      Begin VB.Menu menuRegLabel 
         Caption         =   " 라벨 등록 "
      End
      Begin VB.Menu menuSep201 
         Caption         =   "-"
      End
      Begin VB.Menu menuRegBar 
         Caption         =   " 바코드등록 "
      End
      Begin VB.Menu menuSep202 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHosp 
         Caption         =   "▣ 기관정보 설정"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu menuWork 
      Caption         =   " 작업 "
      Visible         =   0   'False
      Begin VB.Menu menuOrder 
         Caption         =   " 작업지시서 등록 "
      End
   End
   Begin VB.Menu menuPrint 
      Caption         =   " 출력 "
      Visible         =   0   'False
      Begin VB.Menu menuReelPrint 
         Caption         =   " Reel 라벨출력 "
      End
      Begin VB.Menu menuSep501 
         Caption         =   "-"
      End
      Begin VB.Menu menuPPBoxPrint 
         Caption         =   " PP Box 라벨출력 "
      End
      Begin VB.Menu menuSep502 
         Caption         =   "-"
      End
      Begin VB.Menu menuICEBoxPrint 
         Caption         =   " ICE Box 라벨출력 "
      End
      Begin VB.Menu menuSep503 
         Caption         =   "-"
      End
      Begin VB.Menu menuRePrint 
         Caption         =   " 라벨 재출력 "
      End
      Begin VB.Menu menuSep504 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu menuTestPrint 
         Caption         =   " 테스트 출력 "
         Visible         =   0   'False
      End
      Begin VB.Menu menuSep506 
         Caption         =   "-"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMenu05 
      Caption         =   " 옵션 "
      Visible         =   0   'False
      Begin VB.Menu mnuOpt 
         Caption         =   "▣ 옵션 설정"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuBarcodeUse 
         Caption         =   "▷ 바코드 사용"
         WindowList      =   -1  'True
         Begin VB.Menu mnuBarcode 
            Caption         =   "바코드사용"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuSeqno 
            Caption         =   "순번사용"
         End
         Begin VB.Menu mnuRackPos 
            Caption         =   "Rack/Pos"
         End
         Begin VB.Menu mnuCheckBox 
            Caption         =   "체크순"
         End
      End
      Begin VB.Menu mnuSep13 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveResult 
         Caption         =   "▷ 적용 결과"
         Begin VB.Menu mnuEqpResult 
            Caption         =   "장비결과"
         End
         Begin VB.Menu mnuLisResult 
            Caption         =   "LIS결과"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuSep14 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "▷ 결과 전송"
         Begin VB.Menu mnuSaveAuto 
            Caption         =   "자동"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuSaveManual 
            Caption         =   "수동"
         End
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEMR 
         Caption         =   "▷ EMR 설정"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMenu03 
      Caption         =   " 기타 "
      Visible         =   0   'False
      Begin VB.Menu mnuHelp02 
         Caption         =   "원격지원(LG Uplus)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHelp03 
         Caption         =   "원격지원(ez Help)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep21 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCommTest 
         Caption         =   "통신테스트"
      End
   End
End
Attribute VB_Name = "frmMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdView_Click(Index As Integer)
    Dim i   As Integer

    For i = 0 To 2
        cmdView(i).BackColor = vbWhite
    Next

    cmdView(Index).BackColor = &HC0FFFF
    
    If Index = 0 Then
        Call frmShow(frmJoui)
    ElseIf Index = 1 Then
        Call frmShow(frmBuwi)
    ElseIf Index = 2 Then
        Call frmShow(frmJouiDetail)
    End If
    
End Sub

Private Sub MDIForm_Load()
    
    tmrNow.Interval = 1000
    tmrNow.Enabled = True
    
    dtpToday = Now
    dtpTotime = Now
    
    Call frmShow(frmJoui)
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    
    End

End Sub


Private Sub mnuExit_Click()
    
    End

End Sub

Private Sub tmrNow_Timer()

    dtpTotime = Now
    
End Sub
