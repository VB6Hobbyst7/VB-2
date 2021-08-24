VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmRegLabel 
   BackColor       =   &H00FFFFFF&
   Caption         =   "라벨정보등록"
   ClientHeight    =   13110
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20850
   BeginProperty Font 
      Name            =   "맑은 고딕"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   13110
   ScaleWidth      =   20850
   Tag             =   "LBL_LABEL_MASTER,DETAIL"
   WindowState     =   2  '최대화
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   90
      TabIndex        =   21
      Top             =   60
      Width           =   20000
      Begin VB.ComboBox cboLabel 
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   10920
         Style           =   2  '드롭다운 목록
         TabIndex        =   3
         Top             =   360
         Width           =   2085
      End
      Begin VB.ComboBox cboProd 
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6360
         Style           =   2  '드롭다운 목록
         TabIndex        =   2
         Top             =   360
         Width           =   3105
      End
      Begin VB.ComboBox cboComp 
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1350
         Style           =   2  '드롭다운 목록
         TabIndex        =   1
         Top             =   360
         Width           =   3795
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00E0E0E0&
         Caption         =   "조회"
         Height          =   375
         Left            =   13290
         Style           =   1  '그래픽
         TabIndex        =   4
         Top             =   300
         Width           =   1095
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00E0E0E0&
         Caption         =   "화면정리"
         Height          =   375
         Left            =   14430
         Style           =   1  '그래픽
         TabIndex        =   5
         ToolTipText     =   "현재화면을 모두 지웁니다"
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '투명
         Caption         =   "▶ 구분"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   9900
         TabIndex        =   40
         Top             =   390
         Width           =   1065
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "▶ 제품"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   5340
         TabIndex        =   39
         Top             =   390
         Width           =   1065
      End
      Begin VB.Label Label9 
         BackStyle       =   0  '투명
         Caption         =   "▶ 고객사 "
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   330
         TabIndex        =   22
         Top             =   390
         Width           =   1065
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   11595
      Left            =   90
      TabIndex        =   0
      Top             =   1050
      Width           =   20000
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   10095
         Left            =   6000
         TabIndex        =   23
         Top             =   300
         Width           =   13245
         Begin VB.CommandButton cmdSetDefault 
            Caption         =   "기본항목 불러오기"
            Height          =   405
            Left            =   10980
            TabIndex        =   41
            Top             =   4590
            Width           =   2085
         End
         Begin VB.ComboBox cboCompCd 
            BeginProperty Font 
               Name            =   "돋움체"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2490
            Style           =   2  '드롭다운 목록
            TabIndex        =   6
            Top             =   360
            Width           =   3735
         End
         Begin VB.TextBox txtProdLabelCd 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00C0FFFF&
            Height          =   375
            Left            =   2490
            MaxLength       =   30
            TabIndex        =   9
            Text            =   "R-0388_0"
            Top             =   1860
            Width           =   3720
         End
         Begin VB.ComboBox cboProdCd 
            BackColor       =   &H00C0FFFF&
            Height          =   375
            ItemData        =   "frmRegLabel.frx":0000
            Left            =   2490
            List            =   "frmRegLabel.frx":0002
            Style           =   2  '드롭다운 목록
            TabIndex        =   7
            Top             =   810
            Width           =   2085
         End
         Begin VB.CommandButton cmdRemove 
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            Caption         =   "(-) 항목제거"
            Height          =   405
            Left            =   1650
            Style           =   1  '그래픽
            TabIndex        =   16
            Top             =   4590
            Width           =   1395
         End
         Begin VB.CommandButton cmdAdd 
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            Caption         =   "(+) 항목추가"
            Height          =   405
            Left            =   240
            Style           =   1  '그래픽
            TabIndex        =   15
            Top             =   4590
            Width           =   1395
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  '없음
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
            Left            =   9630
            TabIndex        =   36
            Top             =   9240
            Width           =   3525
            Begin VB.CommandButton cmdClose 
               BackColor       =   &H00E0E0E0&
               Caption         =   "닫기"
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   465
               Left            =   2340
               Style           =   1  '그래픽
               TabIndex        =   20
               Top             =   150
               Width           =   1095
            End
            Begin VB.CommandButton cmdOK 
               Appearance      =   0  '평면
               BackColor       =   &H00E0E0E0&
               Caption         =   "저장"
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   465
               Left            =   60
               Style           =   1  '그래픽
               TabIndex        =   18
               Top             =   150
               Width           =   1095
            End
            Begin VB.CommandButton cmdDelete 
               BackColor       =   &H00E0E0E0&
               Caption         =   "삭제"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "맑은 고딕"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   465
               Left            =   1200
               Style           =   1  '그래픽
               TabIndex        =   19
               Top             =   150
               Width           =   1095
            End
         End
         Begin VB.ComboBox cboPrtCode 
            Height          =   375
            ItemData        =   "frmRegLabel.frx":0004
            Left            =   2490
            List            =   "frmRegLabel.frx":0006
            Style           =   2  '드롭다운 목록
            TabIndex        =   12
            Top             =   3120
            Width           =   3735
         End
         Begin VB.ComboBox cboPrtSide 
            Height          =   375
            ItemData        =   "frmRegLabel.frx":0008
            Left            =   2490
            List            =   "frmRegLabel.frx":000A
            Style           =   2  '드롭다운 목록
            TabIndex        =   11
            Top             =   2700
            Width           =   3735
         End
         Begin VB.TextBox txtMaxTot 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   2460
            MaxLength       =   2
            TabIndex        =   13
            Text            =   "0"
            Top             =   3570
            Width           =   3720
         End
         Begin VB.TextBox txtLabelPrtNo 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   2490
            MaxLength       =   2
            TabIndex        =   10
            Text            =   "0"
            Top             =   2280
            Width           =   3720
         End
         Begin VB.ComboBox cboLabelType 
            Height          =   375
            ItemData        =   "frmRegLabel.frx":000C
            Left            =   2490
            List            =   "frmRegLabel.frx":000E
            Style           =   2  '드롭다운 목록
            TabIndex        =   8
            Top             =   1260
            Width           =   3735
         End
         Begin VB.TextBox txtProdCd 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00D0E0E0&
            Height          =   375
            Left            =   4590
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   25
            Top             =   810
            Width           =   1605
         End
         Begin VB.TextBox txtProdLen 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00D0E0E0&
            Height          =   375
            Left            =   7350
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   24
            Top             =   810
            Width           =   1140
         End
         Begin VB.CheckBox chkUsedYN 
            BackColor       =   &H00FFFFFF&
            Caption         =   "사용"
            Height          =   255
            Left            =   2580
            TabIndex        =   14
            Top             =   4050
            Width           =   795
         End
         Begin FPSpread.vaSpread spdRegLabelDetail 
            Height          =   4125
            Left            =   240
            TabIndex        =   17
            Top             =   5070
            Width           =   12825
            _Version        =   393216
            _ExtentX        =   22622
            _ExtentY        =   7276
            _StockProps     =   64
            ColsFrozen      =   8
            DisplayRowHeaders=   0   'False
            EditEnterAction =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GrayAreaBackColor=   16777215
            GridColor       =   15921919
            GridShowVert    =   0   'False
            MaxCols         =   15
            MaxRows         =   20
            RetainSelBlock  =   0   'False
            ScrollBarMaxAlign=   0   'False
            ScrollBars      =   2
            ScrollBarShowMax=   0   'False
            ShadowColor     =   16775150
            SpreadDesigner  =   "frmRegLabel.frx":0010
            ScrollBarTrack  =   3
            ShowScrollTips  =   3
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00808080&
            BorderWidth     =   2
            Height          =   30
            Left            =   240
            Top             =   4470
            Width           =   12735
         End
         Begin VB.Shape Shape2 
            BorderColor     =   &H00808080&
            BorderWidth     =   2
            Height          =   30
            Left            =   270
            Top             =   1710
            Width           =   12735
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "라벨코드"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   38
            Top             =   1860
            Width           =   2205
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "PPBox Reel 내부코드"
            ForeColor       =   &H80000008&
            Height          =   405
            Index           =   16
            Left            =   240
            TabIndex        =   35
            Top             =   3120
            Width           =   2190
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "PP Box Reel출력"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   4
            Left            =   240
            TabIndex        =   34
            Top             =   2700
            Width           =   2205
         End
         Begin VB.Label lblMaxQty 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "박스당 Reel 최대수량"
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   240
            TabIndex        =   33
            Top             =   3570
            Width           =   2175
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "라벨출력수량"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   15
            Left            =   240
            TabIndex        =   32
            Top             =   2280
            Width           =   2205
         End
         Begin VB.Label lblUser 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "제품구분"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   31
            Top             =   1260
            Width           =   2205
         End
         Begin VB.Label lblUser 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "제품길이"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   0
            Left            =   6300
            TabIndex        =   30
            Top             =   810
            Width           =   1020
         End
         Begin VB.Label lblUser 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "사용여부"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   5
            Left            =   240
            TabIndex        =   29
            Top             =   3990
            Width           =   2175
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "고객사"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   2
            Left            =   240
            TabIndex        =   28
            Top             =   360
            Width           =   2205
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "제품명"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   27
            Top             =   810
            Width           =   2205
         End
         Begin VB.Label lblWorkDate 
            Alignment       =   2  '가운데 맞춤
            BackStyle       =   0  '투명
            Caption         =   "미터(M)"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   8550
            TabIndex        =   26
            Top             =   870
            Width           =   975
         End
      End
      Begin FPSpread.vaSpread spdRegLabel 
         Height          =   9975
         Left            =   210
         TabIndex        =   37
         Top             =   390
         Width           =   5745
         _Version        =   393216
         _ExtentX        =   10134
         _ExtentY        =   17595
         _StockProps     =   64
         ColsFrozen      =   8
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         GridColor       =   15921919
         GridShowVert    =   0   'False
         MaxCols         =   19
         MaxRows         =   20
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         ScrollBarExtMode=   -1  'True
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   16774120
         SpreadDesigner  =   "frmRegLabel.frx":0BED
         ScrollBarTrack  =   3
         ShowScrollTips  =   3
      End
   End
End
Attribute VB_Name = "frmRegLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-----------------------------------------------------------------------------'
'   파일명  : frmRegLabel.frm
'   작성자  : 오세원
'   내  용  : 라벨 등록
'   작성일  : 2020-02-07
'   버  전  : 1.0.0
'   고  객  : 국도화학
'-----------------------------------------------------------------------------'


Private Sub cboComp_Click()
    Dim strCompCd   As String
    
    strCompCd = Trim(mGetP(cboComp.Text, 2, "|"))

    Call GetProdList_CodeName("", strCompCd)

End Sub



Private Sub cboCompCd_Click()
    Dim strCompCd   As String

    strCompCd = Trim(mGetP(cboCompCd.Text, 2, "|"))

    Call GetProdList_CodeName_Reg("", strCompCd)
    
End Sub

Private Sub cboLabelType_Click()
    
    txtProdLabelCd.Text = ""
    
    spdRegLabelDetail.MaxRows = 0
    
    If Len(txtProdLabelCd.Text) = 2 Then
        txtProdLabelCd.Text = Mid(cboLabelType, 1, 1) & "-"
    End If
    
    If Mid(cboLabelType, 1, 1) = "R" Then
        lblMaxQty.Caption = ""
    ElseIf Mid(cboLabelType, 1, 1) = "P" Then
        lblMaxQty.Caption = "박스당 Reel 최대수량"
    ElseIf Mid(cboLabelType, 1, 1) = "I" Then
        lblMaxQty.Caption = "박스당 PPBox 최대수량"
    End If

End Sub


Private Sub cboProdCd_Click()
    
    txtProdCd.Text = Trim(mGetP(cboProdCd.Text, 2, "|"))
    
    Call GetComp_CodeName(txtProdCd.Text)
    
    txtProdLabelCd.Text = ""
    
    spdRegLabelDetail.MaxRows = 0
    
End Sub

'-- 항목추가
Private Sub cmdAdd_Click()
    Dim pAdoRS      As ADODB.Recordset
    Dim intRow      As Integer
    Dim intNum      As Integer
    Dim intMaxNum   As Integer
    
    intMaxNum = 0
    With spdRegLabelDetail
        For intRow = 1 To .MaxRows
            intNum = GetText(spdRegLabelDetail, intRow, 1)
            If intMaxNum < intNum Then
                intMaxNum = intNum
            End If
        Next
    
        intMaxNum = intMaxNum + 1
        
        .MaxRows = .MaxRows + 1
            
        'ITEM_NO
        Call SetText(spdRegLabelDetail, intMaxNum, spdRegLabelDetail.MaxRows, 1)
        
        .Row = .MaxRows
        .Col = 2
        .CellType = CellTypeComboBox
        
        Set pAdoRS = New ADODB.Recordset
        Set pAdoRS = Get_TempMaster("T01", "", "", "", "Desc")
        If pAdoRS Is Nothing Then
            '등록된 정보 없음
        Else
            Do Until pAdoRS.EOF
                .TypeComboBoxString = pAdoRS.Fields("NAME1").Value & ""
                pAdoRS.MoveNext
            Loop
        End If
    
        pAdoRS.Close
        
        
        '순서
        Call SetText(spdRegLabelDetail, CStr(spdRegLabelDetail.MaxRows), spdRegLabelDetail.MaxRows, 3)
        '항목출력
        Call SetText(spdRegLabelDetail, "예", spdRegLabelDetail.MaxRows, 5)
        Call SetText(spdRegLabelDetail, "0", spdRegLabelDetail.MaxRows, 6)
        Call SetText(spdRegLabelDetail, "0", spdRegLabelDetail.MaxRows, 7)
        Call SetText(spdRegLabelDetail, "1", spdRegLabelDetail.MaxRows, 9)
        '회전
        Call SetText(spdRegLabelDetail, "0", spdRegLabelDetail.MaxRows, 10)
        '사용구분
        Call SetText(spdRegLabelDetail, "1", spdRegLabelDetail.MaxRows, 11)
    End With

End Sub

Private Sub cmdClear_Click()
        
    spdRegLabel.MaxRows = 0
    spdRegLabelDetail.MaxRows = 0
    
    txtProdCd.Text = ""
    txtProdLen.Text = ""
    txtProdLabelCd.Text = ""
    
    cboPrtSide.Clear
    cboPrtSide.AddItem "아니오"
    cboPrtSide.AddItem "예"
    cboPrtSide.ListIndex = 0
    
    cboPrtCode.Clear
    cboPrtCode.AddItem "아니오"
    cboPrtCode.AddItem "예"
    cboPrtCode.ListIndex = 0
    
    txtLabelPrtNo.Text = "0"
    txtMaxTot.Text = "0"
    
End Sub

Private Sub cmdClose_Click()
    
    Unload Me
    
End Sub

Private Sub cmdDelete_Click()

    Dim intRow      As Integer
    Dim intCol      As Integer
    Dim intItemNo   As Integer
    
    If MsgBox("선택항목을 삭제하시겠습니까?", vbYesNo + vbCritical + vbDefaultButton1, Me.Caption) = vbNo Then
        Exit Sub
    End If
    
    
    If cboCompCd.Text = "" Then
        MsgBox "고객사를 선택하세요", vbOKOnly + vbCritical, Me.Caption
        cboCompCd.SetFocus
        Exit Sub
    End If

    If txtProdCd.Text = "" Then
        MsgBox "제품명을 선택하세요", vbOKOnly + vbCritical, Me.Caption
        cboProdCd.SetFocus
        Exit Sub
    End If


    If txtProdLabelCd.Text = "" Then
        MsgBox "라벨코드를 입력하세요", vbOKOnly + vbCritical, Me.Caption
        txtProdLabelCd.SetFocus
        Exit Sub
    End If
        
    '-- 담기
    gLblMaster.LABELCD = txtProdLabelCd.Text                 'Key
    gLblDetail.LABELCD = txtProdLabelCd.Text                 'Key

    'INSERT
    If Set_Label_Master("DEL") Then
        If Set_Label_Detail("DEL", 1) Then
            Call cmdSearch_Click
            'Call CtlInitializing
            'Call GetProdList
        End If

    End If

End Sub

Private Sub cmdOK_Click()

    Call SetLabel

End Sub

Private Sub SetLabel()
    Dim intRow      As Integer
    Dim intCol      As Integer
    Dim intItemNo   As Integer
    
    '필수입력 체크
    If cboCompCd.Text = "" Then
        MsgBox "고객사를 선택하세요", vbOKOnly + vbCritical, Me.Caption
        cboCompCd.SetFocus
        Exit Sub
    End If

    If txtProdCd.Text = "" Then
        MsgBox "제품명을 선택하세요", vbOKOnly + vbCritical, Me.Caption
        cboProdCd.SetFocus
        Exit Sub
    End If


    If txtProdLabelCd.Text = "" Then
        MsgBox "라벨코드를 입력하세요", vbOKOnly + vbCritical, Me.Caption
        txtProdLabelCd.SetFocus
        Exit Sub
    End If
    
    '-- 담기
    gLblMaster.LABELCD = txtProdLabelCd.Text               'Key
    gLblMaster.PRODCD = txtProdCd.Text
    gLblMaster.COMPCD = Trim(mGetP(cboCompCd.Text, 2, "|")) 'txtCompCd.Text
    gLblMaster.LBLTYPE = Mid(cboLabelType.Text, 1, 1)
    
    gLblMaster.LBLPRTNO = txtLabelPrtNo.Text
    gLblMaster.LBLPRTSIDE = IIf(cboPrtSide.Text = "아니오", "N", "Y")
    gLblMaster.LBLBARSIDE1 = IIf(cboPrtCode.Text = "아니오", "N", "Y")      '내부코드 사용
    gLblMaster.LBLBARSIDE2 = ""                                             '삼성 PP BOX 우측
    gLblMaster.PRODMAXTOT = txtMaxTot.Text

    With spdRegLabelDetail
        gLblDetail.LABELCD = txtProdLabelCd.Text            'Key
        ReDim gLblDetail.LBLITEM_NO(.MaxRows) As String    'Key
        ReDim gLblDetail.LBLITEM_SEQ(.MaxRows) As String
        ReDim gLblDetail.LBLITEM_NAME(.MaxRows) As String
        ReDim gLblDetail.LBLITEM_NMPRT(.MaxRows) As String
        ReDim gLblDetail.LBLITEM_BARGU(.MaxRows) As String
        ReDim gLblDetail.LBLITEM_BARCD(.MaxRows) As String
        ReDim gLblDetail.LBLITEM_X(.MaxRows) As String
        ReDim gLblDetail.LBLITEM_Y(.MaxRows) As String
        ReDim gLblDetail.LBLITEM_FONT(.MaxRows) As String
        ReDim gLblDetail.LBLITEM_ROT(.MaxRows) As String
        ReDim gLblDetail.YN(.MaxRows) As String
        
        For intRow = 1 To .DataRowCnt
            gLblDetail.LBLITEM_NO(intRow) = GetText(spdRegLabelDetail, intRow, 1)
            gLblDetail.LBLITEM_SEQ(intRow) = GetText(spdRegLabelDetail, intRow, 3)
            gLblDetail.LBLITEM_NAME(intRow) = GetText(spdRegLabelDetail, intRow, 4)
            gLblDetail.LBLITEM_NMPRT(intRow) = IIf(GetText(spdRegLabelDetail, intRow, 5) = "예", "Y", "N")
            gLblDetail.LBLITEM_X(intRow) = GetText(spdRegLabelDetail, intRow, 6)
            gLblDetail.LBLITEM_Y(intRow) = GetText(spdRegLabelDetail, intRow, 7)
            '바코드구분 :
            ' >> 바코드일경우 바코드 종류   : "1" : 1D , "2" : 2D
            ' >> 아닐    경우 구분자 : . /
            gLblDetail.LBLITEM_BARGU(intRow) = GetText(spdRegLabelDetail, intRow, 8)
            'BC : code128, BQ : QR코드
            If Mid(GetText(spdRegLabelDetail, intRow, 8), 1, 1) = "1" Then '1D 바코드
                gLblDetail.LBLITEM_BARCD(intRow) = "BC"
            Else
                gLblDetail.LBLITEM_BARCD(intRow) = "BQ"
            End If
            
            gLblDetail.LBLITEM_FONT(intRow) = GetText(spdRegLabelDetail, intRow, 9)
            gLblDetail.LBLITEM_ROT(intRow) = GetText(spdRegLabelDetail, intRow, 10)
            gLblDetail.YN(intRow) = IIf(GetText(spdRegLabelDetail, intRow, 11) = "1", "Y", "N")
        Next
    End With
    
    If chkUsedYN.Value = "1" Then
        gLblMaster.YN = "Y"
    Else
        gLblMaster.YN = "N"
    End If
                
    '-- Insert / Update 찾아오기
    Set AdoRs = Get_LabelMaster(gLblMaster.LABELCD)
        
    If AdoRs.RecordCount = 0 Then
        'INSERT
        If Set_Label_Master("IN") Then
            '상세내용 저장
            For intRow = 1 To spdRegLabelDetail.DataRowCnt
                If Set_Label_Detail("IN", intRow) Then
                End If
            Next
        End If
        Call cmdSearch_Click
    Else
        'UPDATE
        If Set_Label_Master("UP") Then
            If Set_Label_Detail("DEL", intRow) Then
                '상세내용 저장
                For intRow = 1 To spdRegLabelDetail.DataRowCnt
                    If Set_Label_Detail("IN", intRow) Then
                    End If
                Next
            End If
        End If
        Call cmdSearch_Click
    End If
    
    
End Sub


      
Private Sub GetProdList_CodeName(Optional ByVal pProdCd As String, Optional ByVal pCompCd As String)
    Dim pAdoRS      As ADODB.Recordset
    
    Set pAdoRS = New ADODB.Recordset
    
    Set pAdoRS = Get_ProdList_CodeName(pProdCd, pCompCd)
    
    'cboProdCd.Clear
    cboProd.Clear
    
    If pAdoRS Is Nothing Then
        '등록된 정보 없음
    Else
        'cboProdCd.AddItem "전체" & Space(50) & "|전체"
        cboProd.AddItem "전체" & Space(50) & "|전체"
        
        Do Until pAdoRS.EOF
     '       cboProdCd.AddItem pAdoRS.Fields("PROD_NAME").Value & "-" & pAdoRS.Fields("PROD_LENGTH").Value & Space(50) & "|" & pAdoRS.Fields("PROD_CD").Value
            cboProd.AddItem pAdoRS.Fields("PROD_NAME").Value & "-" & pAdoRS.Fields("PROD_LENGTH").Value & Space(50) & "|" & pAdoRS.Fields("PROD_CD").Value
            pAdoRS.MoveNext
        Loop
                            
        If pAdoRS.RecordCount > 0 Then
      '      cboProdCd.ListIndex = 0
            cboProd.ListIndex = 0
        End If
    End If
    
    pAdoRS.Close
    
End Sub
    
Private Sub GetProdList_CodeName_Reg(Optional ByVal pProdCd As String, Optional ByVal pCompCd As String)
    Dim pAdoRS      As ADODB.Recordset
    
    Set pAdoRS = New ADODB.Recordset
    
    Set pAdoRS = Get_ProdList_CodeName(pProdCd, pCompCd)
    
    cboProdCd.Clear
    'cboProd.Clear
    
    If pAdoRS Is Nothing Then
        '등록된 정보 없음
    Else
        'cboProdCd.AddItem "전체" & Space(50) & "|전체"
        cboProd.AddItem "전체" & Space(50) & "|전체"
        
        Do Until pAdoRS.EOF
            cboProdCd.AddItem pAdoRS.Fields("PROD_NAME").Value & "-" & pAdoRS.Fields("PROD_LENGTH").Value & Space(50) & "|" & pAdoRS.Fields("PROD_CD").Value
     '       cboProd.AddItem pAdoRS.Fields("PROD_NAME").Value & "-" & pAdoRS.Fields("PROD_LENGTH").Value & Space(50) & "|" & pAdoRS.Fields("PROD_CD").Value
            pAdoRS.MoveNext
        Loop
                            
        If pAdoRS.RecordCount > 0 Then
            cboProdCd.ListIndex = 0
     '       cboProd.ListIndex = 0
        End If
    End If
    
    pAdoRS.Close
    
End Sub
    
' 라벨 리스트 가져옴
Private Sub GetLabelList(Optional ByVal pProdCd As String, Optional ByVal pCompCd As String, Optional ByVal pLabelType As String)
    
    Dim strLabelType    As String
    
    Set AdoRs = Get_LabelList(pProdCd, pCompCd, pLabelType)
    
    If AdoRs Is Nothing Then
        '등록된 정보 없음
    Else
        Do Until AdoRs.EOF
            With spdRegLabel
                .MaxRows = .MaxRows + 1
                
                Call SetText(spdRegLabel, AdoRs.Fields("PROD_CD").Value & "", .MaxRows, 1)
                Call SetText(spdRegLabel, AdoRs.Fields("PROD_NAME").Value & "", .MaxRows, 2)
                Call SetText(spdRegLabel, AdoRs.Fields("PROD_LENGTH").Value & "", .MaxRows, 3)
                strLabelType = AdoRs.Fields("PROD_LABEL_TYPE").Value & ""
                Select Case UCase(strLabelType)
                    Case "R": Call SetText(spdRegLabel, "Reel", .MaxRows, 4)
                    Case "P": Call SetText(spdRegLabel, "PP Box", .MaxRows, 4)
                    Case "I": Call SetText(spdRegLabel, "ICE Box", .MaxRows, 4)
                End Select
                Call SetText(spdRegLabel, AdoRs.Fields("COMP_CD").Value & "", .MaxRows, 5)
                Call SetText(spdRegLabel, AdoRs.Fields("COMP_NAME").Value & "", .MaxRows, 6)
                Call SetText(spdRegLabel, AdoRs.Fields("PROD_LABEL_CD").Value & "", .MaxRows, 7)
                Call SetText(spdRegLabel, AdoRs.Fields("LABEL_PRT_NO").Value & "", .MaxRows, 8)
                Call SetText(spdRegLabel, AdoRs.Fields("LABEL_PRT_SIDE").Value & "", .MaxRows, 9)
                Call SetText(spdRegLabel, AdoRs.Fields("LABEL_BAR_SIDE01_TYPE").Value & "", .MaxRows, 10)
                Call SetText(spdRegLabel, AdoRs.Fields("LABEL_BAR_SIDE02_TYPE").Value & "", .MaxRows, 11)
                Call SetText(spdRegLabel, AdoRs.Fields("PROD_MAX_TOT").Value & "", .MaxRows, 12)
                If AdoRs.Fields("USED_YN").Value & "" = "Y" Then
                    Call SetText(spdRegLabel, "1", .MaxRows, 15)
                Else
                    Call SetText(spdRegLabel, "0", .MaxRows, 15)
                End If
                Call SetText(spdRegLabel, AdoRs.Fields("REGIST_ID").Value & "", .MaxRows, 16)
                Call SetText(spdRegLabel, AdoRs.Fields("REGIST_DT").Value & "", .MaxRows, 17)
                Call SetText(spdRegLabel, AdoRs.Fields("MODIFY_ID").Value & "", .MaxRows, 18)
                Call SetText(spdRegLabel, AdoRs.Fields("MODIFY_DT").Value & "", .MaxRows, 19)
            End With
            
            AdoRs.MoveNext
        Loop
        
    End If
    
    AdoRs.Close
    
End Sub


' 라벨 리스트 가져옴
Private Sub GetLabelDetail(ByVal pProdLabelCd As String)
    Dim strLabelType    As String
    Dim strItemName     As String
    Dim strBarGu        As String
    Dim strBarType      As String
    Dim strTempCode     As String
    
    If Mid(pProdLabelCd, 1, 1) = "R" Then
        strTempCode = "T01"
    ElseIf Mid(pProdLabelCd, 1, 1) = "P" Then
        strTempCode = "T02"
    ElseIf Mid(pProdLabelCd, 1, 1) = "I" Then
        strTempCode = "T03"
    End If
    
    Set AdoRs = Get_LabelDetail(pProdLabelCd)
            
    gLblDetail.LABELCD = pProdLabelCd
    
    If AdoRs Is Nothing Then
        '등록된 정보 없음
    Else
        Do Until AdoRs.EOF
            With spdRegLabelDetail
                .MaxRows = .MaxRows + 1
                
                Call SetText(spdRegLabelDetail, AdoRs.Fields("LABEL_ITEM_NO").Value & "", .MaxRows, 1)
                strItemName = Get_TempMaster_Name1(strTempCode, AdoRs.Fields("LABEL_ITEM_NO").Value & "", "", "", "ASC")
                
                Call SetText(spdRegLabelDetail, strItemName, .MaxRows, 2)
                Call SetText(spdRegLabelDetail, AdoRs.Fields("LABEL_ITEM_SEQ").Value & "", .MaxRows, 3)
                Call SetText(spdRegLabelDetail, AdoRs.Fields("LABEL_ITEM_NAME").Value & "", .MaxRows, 4)
                Call SetText(spdRegLabelDetail, IIf(AdoRs.Fields("LABEL_ITEM_NAME_PRT").Value & "" = "Y", "예", "아니오"), .MaxRows, 5)
                Call SetText(spdRegLabelDetail, AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "", .MaxRows, 6)
                Call SetText(spdRegLabelDetail, AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "", .MaxRows, 7)
                
                strBarGu = AdoRs.Fields("LABEL_ITEM_GU").Value & ""
                
                If strItemName = "바코드" Then
                    .Row = .MaxRows
                    .Col = 8
                    .CellType = CellTypeComboBox
                    .TypeComboBoxString = "2D 바코드"
                    .TypeComboBoxString = "1D 바코드"
                    .Text = strBarGu '"1D 바코드"
                Else
                    .Row = .MaxRows
                    .Col = 8
                    .CellType = CellTypeEdit
                    .TypeMaxEditLen = 1
                    .TypeHAlign = TypeHAlignCenter
                    .TypeVAlign = TypeVAlignCenter
                    
                    Call SetText(spdRegLabelDetail, strBarGu, .MaxRows, 8)
                End If
                
                Call SetText(spdRegLabelDetail, AdoRs.Fields("LABEL_ITEM_FONT").Value & "", .MaxRows, 9)
                Call SetText(spdRegLabelDetail, AdoRs.Fields("LABEL_ITEM_ROT").Value & "", .MaxRows, 10)
                If AdoRs.Fields("USED_YN").Value & "" = "Y" Then
                    Call SetText(spdRegLabelDetail, "1", .MaxRows, 11)
                Else
                    Call SetText(spdRegLabelDetail, "0", .MaxRows, 11)
                End If
'                Call SetText(spdRegLabelDetail, AdoRs.Fields("REGIST_ID").Value & "", .MaxRows, 16)
'                Call SetText(spdRegLabelDetail, AdoRs.Fields("REGIST_DT").Value & "", .MaxRows, 17)
'                Call SetText(spdRegLabelDetail, AdoRs.Fields("MODIFY_ID").Value & "", .MaxRows, 18)
'                Call SetText(spdRegLabelDetail, AdoRs.Fields("MODIFY_DT").Value & "", .MaxRows, 19)
            End With
            
            AdoRs.MoveNext
        Loop
        
    End If
    
    AdoRs.Close
    
End Sub

Private Sub cmdRemove_Click()
    Dim intItemNo   As Integer
    Dim intRow      As Integer
    
    If spdRegLabelDetail.ActiveRow = 0 Then
        Exit Sub
    End If
    
    intItemNo = GetText(spdRegLabelDetail, spdRegLabelDetail.ActiveRow, 1)
    
    If MsgBox(GetText(spdRegLabelDetail, spdRegLabelDetail.ActiveRow, 2) & " 항목을 삭제하시겠습니까?", vbYesNo + vbCritical + vbDefaultButton1, Me.Caption) = vbYes Then
        'Call SetLabel_Detail(intItemNo, "DEL")
    
        If spdRegLabelDetail.MaxRows > 0 Then
            Call DeleteRow(spdRegLabelDetail, spdRegLabelDetail.ActiveRow, spdRegLabelDetail.ActiveRow)
            spdRegLabelDetail.MaxRows = spdRegLabelDetail.MaxRows - 1
        End If
    End If
    
    DoEvents
    
    'For intRow = 1 To spdRegLabelDetail.MaxRows
    '    Call SetText(spdRegLabelDetail, intRow, intRow, 3)
    'Next
    
End Sub

Private Sub cmdSearch_Click()
    Dim strCompCd    As String
    Dim strProdCd    As String
    Dim strLabelCd   As String
    
    strCompCd = Trim(mGetP(cboComp.Text, 2, "|"))
    strProdCd = Trim(mGetP(cboProd.Text, 2, "|"))
    strLabelCd = UCase(Mid(cboLabel.Text, 1, 1))
    
    Call cmdClear_Click
    
    Call GetLabelList(strProdCd, strCompCd, strLabelCd)
    
End Sub

Private Sub cmdSetDefault_Click()
    Dim pAdoRS      As ADODB.Recordset
    Dim intRow      As Integer
    Dim strLablType As String
    Dim strGubunCd  As String
    
    intRow = 0
    
    If cboLabelType.Text = "" Then
        MsgBox "제품구분을 선택하세요", vbOKOnly, Me.Caption
        Exit Sub
    End If
    
    strLablType = UCase(Mid(cboLabelType.Text, 1, 1))
    
    If strLablType = "R" Then
        strGubunCd = "T01"
    ElseIf strLablType = "P" Then
        strGubunCd = "T02"
    ElseIf strLablType = "I" Then
        strGubunCd = "T03"
    End If
    
    Set pAdoRS = New ADODB.Recordset
    
    Set pAdoRS = Get_TempMaster(strGubunCd)

    If pAdoRS Is Nothing Then
        '등록된 정보 없음
    Else
        With spdRegLabelDetail
            .MaxRows = pAdoRS.RecordCount
        
            Do Until pAdoRS.EOF
                intRow = intRow + 1
                Call SetText(spdRegLabelDetail, pAdoRS.Fields("CODE1").Value & "", intRow, 1)
                Call SetText(spdRegLabelDetail, pAdoRS.Fields("NAME1").Value & "", intRow, 2)
                'Call SetText(spdRegLabelDetail, intRow, intRow, 3)
                Call SetText(spdRegLabelDetail, pAdoRS.Fields("SEQNO").Value & "", intRow, 3)
                
                '기본값
                Call SetText(spdRegLabelDetail, "예", intRow, 5)
                Call SetText(spdRegLabelDetail, "0", intRow, 10)
                Call SetText(spdRegLabelDetail, "1", intRow, 11)
                
                If pAdoRS.Fields("NAME1").Value & "" = "바코드" Then
                    .Row = intRow
                    .Col = 8
                    .CellType = CellTypeComboBox
                    
                    .TypeComboBoxString = "2D 바코드"
                    .TypeComboBoxString = "1D 바코드"
                    .Text = "1D 바코드"
                
                End If
                
                pAdoRS.MoveNext
            Loop
        End With
    End If

    pAdoRS.Close

End Sub

Private Function Get_TempMaster_Name1(ByVal pGubunCd As String, Optional pCode1 As String, Optional pCode2 As String, Optional pCode3 As String, Optional pSort As String) As String
    Dim pAdoRS      As ADODB.Recordset
    
    Get_TempMaster_Name1 = ""
    
    Set pAdoRS = New ADODB.Recordset
    Set pAdoRS = Get_TempMaster(pGubunCd, pCode1, pCode2, pCode3, pSort)

    If pAdoRS Is Nothing Then
        '등록된 정보 없음
    Else
        Do Until pAdoRS.EOF
            Get_TempMaster_Name1 = pAdoRS.Fields("NAME1").Value & ""
            pAdoRS.MoveNext
        Loop
    End If

    pAdoRS.Close

End Function


Private Sub Form_Load()

    Call CtlInitializing
    
    '고객사 리스트 가져오기
    Call GetCompList_CodeName
    
    '제품 리스트 가져오기
    Call GetProdList_CodeName("", "")
    
End Sub


Private Function GetCompList_Name(Optional ByVal pCompCd As String) As String
    Dim pAdoRS      As ADODB.Recordset
    
    Set pAdoRS = New ADODB.Recordset
    
    Set pAdoRS = Get_CompList_Name(pCompCd)

    If pAdoRS Is Nothing Then
        '등록된 정보 없음
    Else
        Do Until pAdoRS.EOF
            GetCompList_Name = pAdoRS.Fields("COMP_NAME").Value & ""

            pAdoRS.MoveNext
        Loop

    End If

    pAdoRS.Close

End Function

'-- 상단 고객사리스트 가져오기
Private Sub GetCompList_CodeName()
    Dim pAdoRS      As ADODB.Recordset
    
    Set pAdoRS = New ADODB.Recordset
    
    Set pAdoRS = Get_CompList_CodeName
    
    If pAdoRS Is Nothing Then
        '등록된 정보 없음
    Else
        cboComp.Clear
        cboCompCd.Clear
        
        cboComp.AddItem "전체" & Space(30) & "|" & "전체"
        
        Do Until pAdoRS.EOF
            cboComp.AddItem pAdoRS.Fields("COMP_NAME").Value & Space(15 - Len(pAdoRS.Fields("COMP_NAME").Value)) & pAdoRS.Fields("COMP_LINE").Value & Space(20) & "|" & pAdoRS.Fields("COMP_CD").Value & ""
            cboCompCd.AddItem pAdoRS.Fields("COMP_NAME").Value & Space(15 - Len(pAdoRS.Fields("COMP_NAME").Value)) & pAdoRS.Fields("COMP_LINE").Value & Space(20) & "|" & pAdoRS.Fields("COMP_CD").Value & ""
            
            pAdoRS.MoveNext
        Loop
        
        If pAdoRS.RecordCount > 0 Then
            cboComp.ListIndex = 0
            cboCompCd.ListIndex = 0
        End If
    End If
    
    pAdoRS.Close
    
End Sub

'-- 제품선택했을때 해당 고객사 가져오기
Private Sub GetComp_CodeName(ByVal pProdCd As String)
    Dim pAdoRS      As ADODB.Recordset
    
    Set pAdoRS = New ADODB.Recordset
    
    Set pAdoRS = Get_Comp_CodeName(pProdCd)
    
    If pAdoRS Is Nothing Then
        '등록된 정보 없음
    Else
        txtProdLen.Text = ""
        
        Do Until pAdoRS.EOF
            txtProdLen.Text = pAdoRS.Fields("PROD_LENGTH").Value & ""
            
            pAdoRS.MoveNext
        Loop
        
    End If
    
    pAdoRS.Close
    
End Sub


'-- 컨트롤초기화
Private Sub CtlInitializing()
    Dim pAdoRS      As ADODB.Recordset
    
    With spdRegLabel
        Call SetText(spdRegLabel, "제품코드", 0, 1):            .ColWidth(1) = 0
        Call SetText(spdRegLabel, "제품명", 0, 2):              .ColWidth(2) = 12
        Call SetText(spdRegLabel, "길이", 0, 3):                .ColWidth(3) = 6
        Call SetText(spdRegLabel, "구분", 0, 4):                .ColWidth(4) = 8
        Call SetText(spdRegLabel, "고객사코드", 0, 5):          .ColWidth(5) = 0
        Call SetText(spdRegLabel, "고객사", 0, 6):              .ColWidth(6) = 7
        Call SetText(spdRegLabel, "라벨코드", 0, 7):            .ColWidth(7) = 9
        Call SetText(spdRegLabel, "출력매수", 0, 8):            .ColWidth(8) = 0
        Call SetText(spdRegLabel, "측면출력구분", 0, 9):        .ColWidth(9) = 0
        Call SetText(spdRegLabel, "측면상1타입", 0, 10):        .ColWidth(10) = 0
        Call SetText(spdRegLabel, "측면상2타입", 0, 11):        .ColWidth(11) = 0
        Call SetText(spdRegLabel, "포장최대수량", 0, 12):       .ColWidth(12) = 0
        Call SetText(spdRegLabel, "미사용", 0, 13):             .ColWidth(13) = 0
        Call SetText(spdRegLabel, "미사용", 0, 14):             .ColWidth(14) = 0
        Call SetText(spdRegLabel, "사용여부", 0, 15):           .ColWidth(15) = 0
        Call SetText(spdRegLabel, "입력자", 0, 16):             .ColWidth(16) = 0
        Call SetText(spdRegLabel, "입력일시", 0, 17):           .ColWidth(17) = 0
        Call SetText(spdRegLabel, "수정자", 0, 18):             .ColWidth(18) = 0
        Call SetText(spdRegLabel, "수정일시", 0, 19):           .ColWidth(19) = 0
    
        .MaxRows = 0
    End With
    
    txtProdLabelCd.Text = ""
    txtProdCd.Text = ""
    txtProdLen.Text = ""
    
    cboLabel.Clear
    cboLabel.AddItem "전체"
    cboLabel.AddItem "Reel"
    cboLabel.AddItem "PP Box"
    cboLabel.AddItem "ICE Box"
    cboLabel.ListIndex = 0
    
    cboLabelType.Clear
    cboLabelType.AddItem "Reel"
    cboLabelType.AddItem "PP Box"
    cboLabelType.AddItem "ICE Box"
    cboLabelType.ListIndex = 0
    
    
    cboPrtSide.Clear
    cboPrtSide.AddItem "아니오"
    cboPrtSide.AddItem "예"
    cboPrtSide.ListIndex = 0
    
    cboPrtCode.Clear
    cboPrtCode.AddItem "아니오"
    cboPrtCode.AddItem "예"
    cboPrtCode.ListIndex = 0
    
    txtLabelPrtNo.Text = "0"
    txtMaxTot.Text = "0"
    
    With spdRegLabelDetail
        Call SetText(spdRegLabelDetail, "No", 0, 1):          .ColWidth(1) = 5
        Call SetText(spdRegLabelDetail, "항목", 0, 2):        .ColWidth(2) = 10
        Call SetText(spdRegLabelDetail, "순서", 0, 3):        .ColWidth(3) = 6
        Call SetText(spdRegLabelDetail, "내용", 0, 4):        .ColWidth(4) = 20
        Call SetText(spdRegLabelDetail, "항목출력", 0, 5):    .ColWidth(5) = 10
        Call SetText(spdRegLabelDetail, "X좌표", 0, 6):       .ColWidth(6) = 8
        Call SetText(spdRegLabelDetail, "Y좌표", 0, 7):       .ColWidth(7) = 8
        Call SetText(spdRegLabelDetail, "구분자", 0, 8):      .ColWidth(8) = 8
        Call SetText(spdRegLabelDetail, "폰트크기", 0, 9):    .ColWidth(9) = 8
        Call SetText(spdRegLabelDetail, "회전", 0, 10):       .ColWidth(10) = 8
        Call SetText(spdRegLabelDetail, "사용여부", 0, 11):   .ColWidth(11) = 12
        Call SetText(spdRegLabelDetail, "입력자", 0, 12):     .ColWidth(12) = 0
        Call SetText(spdRegLabelDetail, "입력일시", 0, 13):   .ColWidth(13) = 0
        Call SetText(spdRegLabelDetail, "수정자", 0, 14):     .ColWidth(14) = 0
        Call SetText(spdRegLabelDetail, "수정일시", 0, 15):   .ColWidth(15) = 0
    
        .MaxRows = 0
    End With
    
    chkUsedYN.Value = "1"
    If gKUKDO.USERGRD = "1" Then
        cmdDelete.Visible = True
    Else
        cmdDelete.Visible = False
    End If
    
    gSORT = 0

End Sub

Private Sub spdRegLabel_Click(ByVal Col As Long, ByVal Row As Long)
    Dim i               As Integer
    Dim strPrtSide      As String
    Dim strProdLabelCd  As String
    
    If Row = 0 Then
        Call SetSpreadSort(spdRegLabel)
        Exit Sub
    End If
        
    For i = 0 To cboCompCd.ListCount
        If Trim(mGetP(cboCompCd.List(i), 2, "|")) = GetText(spdRegLabel, Row, 5) Then
            cboCompCd.ListIndex = i
            Exit For
        End If
    Next
    
    
    For i = 0 To cboProdCd.ListCount
        If Trim(mGetP(cboProdCd.List(i), 2, "|")) = GetText(spdRegLabel, Row, 1) Then
            cboProdCd.ListIndex = i
            txtProdCd.Text = Trim(mGetP(cboProdCd.List(i), 2, "|"))
            Exit For
        End If
    Next
    
    txtProdLen.Text = GetText(spdRegLabel, Row, 3)
    For i = 0 To cboLabelType.ListCount
        If cboLabelType.List(i) = GetText(spdRegLabel, Row, 4) Then
            cboLabelType.ListIndex = i
            Exit For
        End If
    Next
    
    strProdLabelCd = GetText(spdRegLabel, Row, 7)
    txtProdLabelCd.Text = GetText(spdRegLabel, Row, 7)
    txtLabelPrtNo.Text = GetText(spdRegLabel, Row, 8)
    txtMaxTot.Text = GetText(spdRegLabel, Row, 12)
    
    strPrtSide = GetText(spdRegLabel, Row, 9)
    
    If strPrtSide = "Y" Then
        strPrtSide = "예"
    Else
        strPrtSide = "아니오"
    End If
    For i = 0 To cboPrtSide.ListCount
        If cboPrtSide.List(i) = strPrtSide Then
            cboPrtSide.ListIndex = i
            Exit For
        End If
    Next
    
    If GetText(spdRegLabel, Row, 15) = "1" Then
        chkUsedYN.Value = "1"
    Else
        chkUsedYN.Value = "0"
    End If

    spdRegLabelDetail.MaxRows = 0
    
    Call GetLabelDetail(strProdLabelCd)

    
End Sub

Private Sub spdRegLabelDetail_ComboSelChange(ByVal Col As Long, ByVal Row As Long)
    With spdRegLabelDetail
        If Col = 2 Then
            If GetText(spdRegLabelDetail, Row, Col) = "바코드" Then
                .Row = Row
                .Col = 8
                .CellType = CellTypeComboBox
                
                .TypeComboBoxString = "2D 바코드"
                .TypeComboBoxString = "1D 바코드"
                .Text = "1D 바코드"
            Else
            
            End If
        End If
    End With
End Sub


Private Sub txtProdLabelCd_GotFocus()
    If txtProdLabelCd.Text = "" Then
        txtProdLabelCd.Text = Mid(cboLabelType, 1, 1) & "-"
    End If
    
    txtProdLabelCd.SelStart = Len(txtProdLabelCd.Text)
'    txtID.SelLength = Len(txtID.Text)

End Sub
