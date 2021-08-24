VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.MDIForm frmMDI 
   BackColor       =   &H00FFFFFF&
   Caption         =   "국도화학 바코드 발행 시스템"
   ClientHeight    =   8355
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   19515
   Icon            =   "frmMDI.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows 기본값
   WindowState     =   2  '최대화
   Begin VB.PictureBox Picture1 
      Align           =   1  '위 맞춤
      BackColor       =   &H00FFFFFF&
      Height          =   765
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   28500
      TabIndex        =   0
      Top             =   0
      Width           =   28560
      Begin VB.TextBox Text1 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   20790
         TabIndex        =   5
         Text            =   "℃  ￠  "
         Top             =   210
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   645
         Left            =   90
         TabIndex        =   1
         Top             =   0
         Width           =   6975
         Begin VB.TextBox txtTestID 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3870
            Locked          =   -1  'True
            TabIndex        =   7
            Text            =   "1234567890"
            Top             =   180
            Width           =   1185
         End
         Begin VB.TextBox txtTestNm 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5040
            Locked          =   -1  'True
            TabIndex        =   6
            Text            =   "1234567890"
            Top             =   180
            Width           =   1185
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   1140
            TabIndex        =   3
            Top             =   180
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   136118273
            CurrentDate     =   43865
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '투명
            Caption         =   "작업자ID : "
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   2880
            TabIndex        =   8
            Top             =   240
            Width           =   855
         End
         Begin VB.Label lblWorkDate 
            BackStyle       =   0  '투명
            Caption         =   "작업일자 :"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   240
            TabIndex        =   2
            Top             =   240
            Width           =   825
         End
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00808080&
         FillColor       =   &H00E0E0E0&
         Height          =   540
         Left            =   7080
         Top             =   90
         Width           =   5895
      End
      Begin VB.Image Image1 
         Height          =   615
         Left            =   15960
         Picture         =   "frmMDI.frx":144A
         Top             =   30
         Visible         =   0   'False
         Width           =   2490
      End
      Begin VB.Label lblFrmInfo 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "HITACHI 7020"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   7260
         TabIndex        =   4
         Top             =   240
         Width           =   5115
      End
   End
   Begin VB.Menu mnuMenu01 
      Caption         =   "  파일 "
      Begin VB.Menu mnuExit 
         Caption         =   "종료"
      End
   End
   Begin VB.Menu mnuMenu02 
      Caption         =   " 설정 "
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
         Visible         =   0   'False
      End
   End
   Begin VB.Menu menuReg 
      Caption         =   " 등록 "
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
   Begin VB.Menu menuPrint 
      Caption         =   " 출력 "
      Begin VB.Menu menuTestPrint 
         Caption         =   "테스트 출력"
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
      Begin VB.Menu mnuHelp01 
         Caption         =   "원격지원(TeamViewer)"
      End
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

Private Sub MDIForm_Load()

    '권한별 메뉴보이기
    If gKUKDO.USERGRD = "1" Then
        mnuMenu02.Visible = True
    Else
        mnuMenu02.Visible = False
    End If
        
    '화면이름 표시
    lblFrmInfo.Caption = ""
    
    txtTestID.Text = gKUKDO.USERID
    txtTestNm.Text = gKUKDO.USERNM
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    
    End

End Sub

Private Sub menuComp_Click()

    Call frmShow(frmMstComp)
    
End Sub

Private Sub menuPack_Click()

    Call frmShow(frmMstPack)

End Sub

Private Sub menuRegLabel_Click()

    Call frmShow(frmRegLabel)

End Sub

Private Sub menuRegBar_Click()
    
    Call frmShow(frmRegBar)

End Sub

Private Sub menuRegProd_Click()
    
    Call frmShow(frmRegProd)

End Sub

Private Sub menuTestPrint_Click()

    Call frmShow(frmTestPrint)

End Sub

Private Sub menuUser_Click()
    
    Call frmShow(frmMstUser)
    
End Sub

