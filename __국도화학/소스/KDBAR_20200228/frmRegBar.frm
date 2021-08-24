VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmRegBar 
   BackColor       =   &H00FFFFFF&
   Caption         =   "바코드 상세등록"
   ClientHeight    =   11460
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19950
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
   ScaleHeight     =   11460
   ScaleWidth      =   19950
   Tag             =   "LBL_BAR_MASTER,DETAIL"
   WindowState     =   2  '최대화
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
      TabIndex        =   15
      Top             =   1050
      Width           =   19395
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
         TabIndex        =   16
         Top             =   300
         Width           =   13245
         Begin VB.ComboBox cboGubun 
            Height          =   375
            Left            =   8190
            Style           =   2  '드롭다운 목록
            TabIndex        =   42
            Top             =   3870
            Width           =   1845
         End
         Begin VB.CommandButton cmdDate 
            Caption         =   "날짜코드범례"
            Height          =   405
            Left            =   11640
            TabIndex        =   41
            Top             =   3840
            Width           =   1395
         End
         Begin VB.TextBox Text2 
            Height          =   375
            Left            =   3300
            TabIndex        =   40
            Text            =   "Y"
            Top             =   3900
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Left            =   3840
            TabIndex        =   39
            Top             =   3900
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Command1"
            Height          =   315
            Left            =   5130
            TabIndex        =   38
            Top             =   3900
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.ComboBox cboProdLabelCd 
            Height          =   375
            ItemData        =   "frmRegBar.frx":0000
            Left            =   2490
            List            =   "frmRegBar.frx":0002
            Style           =   2  '드롭다운 목록
            TabIndex        =   10
            Top             =   2160
            Width           =   3735
         End
         Begin VB.CheckBox chkUsedYN 
            BackColor       =   &H00FFFFFF&
            Caption         =   "사용"
            Height          =   255
            Left            =   2580
            TabIndex        =   26
            Top             =   3240
            Width           =   795
         End
         Begin VB.TextBox txtProdLen 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00D0E0E0&
            Height          =   375
            Left            =   7350
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   25
            Top             =   810
            Width           =   1140
         End
         Begin VB.TextBox txtProdCd 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00D0E0E0&
            Height          =   375
            Left            =   4590
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   24
            Top             =   810
            Width           =   1605
         End
         Begin VB.ComboBox cboLabelType 
            Height          =   375
            ItemData        =   "frmRegBar.frx":0004
            Left            =   2490
            List            =   "frmRegBar.frx":0006
            Style           =   2  '드롭다운 목록
            TabIndex        =   8
            Top             =   1260
            Width           =   3735
         End
         Begin VB.TextBox txtBarCd 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   375
            Left            =   2490
            MaxLength       =   10
            TabIndex        =   11
            Top             =   2760
            Width           =   3720
         End
         Begin VB.ComboBox cboBarType 
            Height          =   375
            ItemData        =   "frmRegBar.frx":0008
            Left            =   2490
            List            =   "frmRegBar.frx":000A
            Style           =   2  '드롭다운 목록
            TabIndex        =   9
            Top             =   1710
            Width           =   3735
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
            TabIndex        =   20
            Top             =   8490
            Width           =   3525
            Begin VB.CommandButton cmdDelete 
               BackColor       =   &H00E0E0E0&
               Caption         =   "삭제"
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
               TabIndex        =   23
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
               TabIndex        =   22
               Top             =   150
               Width           =   1095
            End
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
               TabIndex        =   21
               Top             =   150
               Width           =   1095
            End
         End
         Begin VB.CommandButton cmdAdd 
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            Caption         =   "(+) 항목추가"
            Height          =   405
            Left            =   240
            Style           =   1  '그래픽
            TabIndex        =   19
            Top             =   3840
            Width           =   1395
         End
         Begin VB.CommandButton cmdRemove 
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            Caption         =   "(-) 항목제거"
            Height          =   405
            Left            =   1650
            Style           =   1  '그래픽
            TabIndex        =   18
            Top             =   3840
            Width           =   1395
         End
         Begin VB.ComboBox cboProdCd 
            Height          =   375
            ItemData        =   "frmRegBar.frx":000C
            Left            =   2490
            List            =   "frmRegBar.frx":000E
            Style           =   2  '드롭다운 목록
            TabIndex        =   7
            Top             =   810
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
         Begin VB.CommandButton cmdSetDefault 
            Caption         =   "항목 불러오기"
            Height          =   405
            Left            =   10050
            TabIndex        =   17
            Top             =   3840
            Width           =   1545
         End
         Begin FPSpread.vaSpread spdRegBarDetail 
            Height          =   4125
            Left            =   240
            TabIndex        =   27
            Top             =   4320
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
            SpreadDesigner  =   "frmRegBar.frx":0010
            ScrollBarTrack  =   3
            ShowScrollTips  =   3
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "바코드 구분"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   3
            Left            =   240
            TabIndex        =   37
            Top             =   1710
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
            TabIndex        =   35
            Top             =   870
            Width           =   975
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
            TabIndex        =   34
            Top             =   810
            Width           =   2205
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
            TabIndex        =   33
            Top             =   360
            Width           =   2205
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
            TabIndex        =   32
            Top             =   3210
            Width           =   2175
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
            TabIndex        =   31
            Top             =   810
            Width           =   1020
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
            TabIndex        =   30
            Top             =   1260
            Width           =   2205
         End
         Begin VB.Label lblComp 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "바코드 TYPE코드"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   15
            Left            =   240
            TabIndex        =   29
            Top             =   2760
            Width           =   2205
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
            TabIndex        =   28
            Top             =   2160
            Width           =   2205
         End
         Begin VB.Shape Shape2 
            BorderColor     =   &H00808080&
            BorderWidth     =   2
            Height          =   30
            Left            =   270
            Top             =   2640
            Width           =   12735
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00808080&
            BorderWidth     =   2
            Height          =   30
            Left            =   240
            Top             =   3660
            Width           =   12735
         End
      End
      Begin FPSpread.vaSpread spdRegBar 
         Height          =   9975
         Left            =   210
         TabIndex        =   36
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
         SpreadDesigner  =   "frmRegBar.frx":0BD6
         ScrollBarTrack  =   3
         ShowScrollTips  =   3
      End
   End
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
      TabIndex        =   0
      Top             =   60
      Width           =   19425
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
      Begin VB.ComboBox cboComp 
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1350
         Style           =   2  '드롭다운 목록
         TabIndex        =   1
         Top             =   360
         Width           =   3795
      End
      Begin VB.ComboBox cboProd 
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6360
         Style           =   2  '드롭다운 목록
         TabIndex        =   2
         Top             =   360
         Width           =   3105
      End
      Begin VB.ComboBox cboLabel 
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10920
         Style           =   2  '드롭다운 목록
         TabIndex        =   3
         Top             =   360
         Width           =   2085
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
         TabIndex        =   14
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
         TabIndex        =   13
         Top             =   390
         Width           =   1065
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
         TabIndex        =   12
         Top             =   390
         Width           =   1065
      End
   End
End
Attribute VB_Name = "frmRegBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-----------------------------------------------------------------------------'
'   파일명  : frmRegBar.frm
'   작성자  : 오세원
'   내  용  : 바코드 등록
'   작성일  : 2020-02-21
'   버  전  : 1.0.0
'   고  객  : 국도화학
'-----------------------------------------------------------------------------'



Private Sub cboBarType_Click()
'''    Dim strCompCd       As String
'''    Dim strLabelType    As String
'''    Dim strBarGu        As String
'''
'''    txtProdCd.Text = Trim(mGetP(cboProdCd.Text, 2, "|"))
'''    strCompCd = Trim(mGetP(cboCompCd.Text, 2, "|"))
'''    strLabelType = Mid(cboLabelType.Text, 1, 1)
'''    strBarGu = cboBarType.Text
'''
'''    '라벨 리스트 가져오기
'''    If strCompCd <> "" Then
'''        Call GetLabelList_CodeName(txtProdCd.Text, strCompCd, strLabelType, strBarGu)
'''    End If

End Sub

Private Sub cboComp_Click()
    Dim strCompCd   As String
    
    strCompCd = Trim(mGetP(cboComp.Text, 2, "|"))

    Call GetProdList_CodeName("", strCompCd)

End Sub



Private Sub cboCompCd_Click()
    Dim strCompCd   As String

    strCompCd = Trim(mGetP(cboCompCd.Text, 2, "|"))

    Call GetProdList_CodeName_Reg("", strCompCd)
    
    'Call GetLabelList_CodeName("", strCompCd)
    

End Sub


Private Sub cboLabelType_Click()
    Dim strCompCd       As String
    Dim strLabelType    As String
    
    spdRegBarDetail.MaxRows = 0
        
    txtProdCd.Text = Trim(mGetP(cboProdCd.Text, 2, "|"))
    strCompCd = Trim(mGetP(cboCompCd.Text, 2, "|"))
    strLabelType = Mid(cboLabelType.Text, 1, 1)
    
    
    '라벨 리스트 가져오기
    If strCompCd <> "" Then
        Call GetLabelList_CodeName(txtProdCd.Text, strCompCd, strLabelType, "")
    End If
    
End Sub

Private Sub cboProdCd_Click()
    Dim strCompCd    As String
    
    txtProdCd.Text = Trim(mGetP(cboProdCd.Text, 2, "|"))
    strCompCd = Trim(mGetP(cboCompCd.Text, 2, "|"))
    
    Call GetComp_CodeName(txtProdCd.Text)
    
    '라벨 리스트 가져오기
    Call GetLabelList_CodeName(txtProdCd.Text, strCompCd, "", "")
    
End Sub


Private Sub cboProdLabelCd_Click()
    
    If cboProdLabelCd.Text <> "" Then
        txtBarCd.Text = cboProdLabelCd.Text
        txtBarCd.Text = txtBarCd.Text & "_"
    End If

    txtBarCd.SelStart = Len(txtBarCd.Text)

End Sub

'-- 항목추가
Private Sub cmdAdd_Click()
    Dim pAdoRS      As ADODB.Recordset
    Dim intRow      As Integer
    Dim intNum      As Integer
    Dim intMaxNum   As Integer
    
    intMaxNum = 0
    With spdRegBarDetail
        For intRow = 1 To .MaxRows
            intNum = GetText(spdRegBarDetail, intRow, 1)
            If intMaxNum < intNum Then
                intMaxNum = intNum
            End If
        Next
    
        intMaxNum = intMaxNum + 1
        
        .MaxRows = .MaxRows + 1
            
        'ITEM_NO
        Call SetText(spdRegBarDetail, intMaxNum, spdRegBarDetail.MaxRows, 1)
        
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
        Call SetText(spdRegBarDetail, CStr(spdRegBarDetail.MaxRows), spdRegBarDetail.MaxRows, 3)
        '항목출력
        Call SetText(spdRegBarDetail, "예", spdRegBarDetail.MaxRows, 5)
        Call SetText(spdRegBarDetail, "0", spdRegBarDetail.MaxRows, 6)
        Call SetText(spdRegBarDetail, "0", spdRegBarDetail.MaxRows, 7)
        Call SetText(spdRegBarDetail, "1", spdRegBarDetail.MaxRows, 9)
        '회전
        Call SetText(spdRegBarDetail, "0", spdRegBarDetail.MaxRows, 10)
        '사용구분
        Call SetText(spdRegBarDetail, "1", spdRegBarDetail.MaxRows, 11)
    End With

End Sub

Private Sub cmdClear_Click()
        
    spdRegBar.MaxRows = 0
    spdRegBarDetail.MaxRows = 0
    
    txtProdCd.Text = ""
    txtProdLen.Text = ""
    txtBarCd.Text = ""
    
End Sub

Private Sub cmdClose_Click()
    
    Unload Me
    
End Sub

Private Sub cmdDate_Click()

    frmMstDateCode.Show
    
End Sub

Private Sub cmdDelete_Click()

    Dim intRow      As Integer
    Dim intCol      As Integer
    Dim intItemNo   As Integer
    
    
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


    If txtBarCd.Text = "" Then
        MsgBox "라벨코드를 입력하세요", vbOKOnly + vbCritical, Me.Caption
        txtBarCd.SetFocus
        Exit Sub
    End If
        
    '-- 담기
    gBarMaster.BARCD = txtBarCd.Text                  'Key
    gBarDetail.BARCD = txtBarCd.Text                  'Key

    'INSERT
    If Set_Bar_Master("DEL") Then
        If Set_Bar_Detail("DEL", 1) Then
            Call cmdSearch_Click
            'Call CtlInitializing
            'Call GetProdList
        End If

    End If

End Sub

Private Sub cmdOK_Click()

    Call SetBar
    
End Sub

Private Sub SetBar()
    Dim intRow      As Integer
    Dim intCol      As Integer
    Dim intItemNo   As Integer
    
    '필수입력 체크
    If txtBarCd.Text = "" Then
        MsgBox "바코드 TYPE코드를 선택하세요", vbOKOnly + vbCritical, Me.Caption
        txtBarCd.SetFocus
        Exit Sub
    End If

    '-- 담기
    gBarMaster.BARCD = txtBarCd.Text                'Key
    gBarMaster.PRODCD = txtProdCd.Text
    gBarMaster.COMPCD = Trim(mGetP(cboCompCd.Text, 2, "|"))
    gBarMaster.BARTYPE = Mid(cboBarType.Text, 1, 1)
    gBarMaster.BARGU = Mid(cboLabelType.Text, 1, 1)
    If chkUsedYN.Value = "1" Then
        gBarMaster.YN = "Y"
    Else
        gBarMaster.YN = "N"
    End If
    
    With spdRegBarDetail
        gBarDetail.BARCD = txtBarCd.Text                    'Key
        ReDim gBarDetail.BARITEM_NO(.MaxRows) As String     'Key
        ReDim gBarDetail.BARITEM_NAME(.MaxRows) As String
        ReDim gBarDetail.BARITEM_SEQ(.MaxRows) As String
        ReDim gBarDetail.BARCHRNUM(.MaxRows) As String
        ReDim gBarDetail.LBLITEMTYPE(.MaxRows) As String
        ReDim gBarDetail.YN(.MaxRows) As String
        
        For intRow = 1 To .DataRowCnt
            gBarDetail.BARITEM_NO(intRow) = GetText(spdRegBarDetail, intRow, 1)
            gBarDetail.BARITEM_NAME(intRow) = GetText(spdRegBarDetail, intRow, 2)
            gBarDetail.BARITEM_SEQ(intRow) = GetText(spdRegBarDetail, intRow, 3)
            gBarDetail.BARCHRNUM(intRow) = GetText(spdRegBarDetail, intRow, 5)
            gBarDetail.LBLITEMTYPE(intRow) = GetText(spdRegBarDetail, intRow, 4)
            gBarDetail.YN(intRow) = IIf(GetText(spdRegBarDetail, intRow, 11) = "1", "Y", "N")
        Next
    End With
                
    '-- Insert / Update 찾아오기
    Set AdoRs = Get_BarMaster(gBarMaster.BARCD)
        
    If AdoRs.RecordCount = 0 Then
        'INSERT
        If Set_Bar_Master("IN") Then
            '상세내용 저장
            For intRow = 1 To spdRegBarDetail.DataRowCnt
                If Set_Bar_Detail("IN", intRow) Then
                    Call cmdSearch_Click
                End If
            Next
        End If
    Else
        'UPDATE
        If Set_Bar_Master("UP") Then
            If Set_Bar_Detail("DEL", intRow) Then
                '상세내용 저장
                For intRow = 1 To spdRegBarDetail.DataRowCnt
                    If Set_Bar_Detail("IN", intRow) Then
                        Call cmdSearch_Click
                    End If
                Next
            End If
        End If
    End If
    
    
End Sub


'라벨 리스트 가져오기
Private Sub GetLabelList_CodeName(Optional ByVal pProdCd As String, Optional ByVal pCompCd As String, Optional ByVal pLabelType As String, Optional ByVal pBarGu As String)
    Dim pAdoRS      As ADODB.Recordset
    
    Set pAdoRS = New ADODB.Recordset
    
    Set pAdoRS = Get_LabelList_CodeName(pProdCd, pCompCd, pLabelType, pBarGu)
    
    cboProdLabelCd.Clear
    txtBarCd.Enabled = False
    
    If pAdoRS Is Nothing Then
        '등록된 정보 없음
    Else
        Do Until pAdoRS.EOF
            cboProdLabelCd.AddItem pAdoRS.Fields("PROD_LABEL_CD").Value & ""
            pAdoRS.MoveNext
        Loop
                            
        If pAdoRS.RecordCount > 0 Then
            cboProdLabelCd.ListIndex = 0
            txtBarCd.Enabled = True
        End If
    End If
    
    pAdoRS.Close
    
End Sub
    
'제품 리스트 가져오기(조회용)
Private Sub GetProdList_CodeName(Optional ByVal pProdCd As String, Optional ByVal pCompCd As String)
    Dim pAdoRS      As ADODB.Recordset
    
    Set pAdoRS = New ADODB.Recordset
    
    Set pAdoRS = Get_ProdList_CodeName(pProdCd, pCompCd)
    
    cboProd.Clear
    
    If pAdoRS Is Nothing Then
        '등록된 정보 없음
    Else
        cboProd.AddItem "전체" & Space(50) & "|전체"
        
        Do Until pAdoRS.EOF
            cboProd.AddItem pAdoRS.Fields("PROD_NAME").Value & "-" & pAdoRS.Fields("PROD_LENGTH").Value & Space(50) & "|" & pAdoRS.Fields("PROD_CD").Value
            pAdoRS.MoveNext
        Loop
                            
        If pAdoRS.RecordCount > 0 Then
            cboProd.ListIndex = 0
        End If
    End If
    
    pAdoRS.Close
    
End Sub
    
    
'제품 리스트 가져오기(등록용)
Private Sub GetProdList_CodeName_Reg(Optional ByVal pProdCd As String, Optional ByVal pCompCd As String)
    Dim pAdoRS      As ADODB.Recordset
    
    Set pAdoRS = New ADODB.Recordset
    
    Set pAdoRS = Get_ProdList_CodeName(pProdCd, pCompCd)
    
    cboProdCd.Clear
    
    If pAdoRS Is Nothing Then
        '등록된 정보 없음
    Else
        cboProd.AddItem "전체" & Space(50) & "|전체"
        
        Do Until pAdoRS.EOF
            cboProdCd.AddItem pAdoRS.Fields("PROD_NAME").Value & "-" & pAdoRS.Fields("PROD_LENGTH").Value & Space(50) & "|" & pAdoRS.Fields("PROD_CD").Value
            pAdoRS.MoveNext
        Loop
                            
        If pAdoRS.RecordCount > 0 Then
            cboProdCd.ListIndex = 0
        End If
    End If
    
    pAdoRS.Close
    
End Sub
    
' 바코드 리스트 가져옴
Private Sub GetBarList(Optional ByVal pProdCd As String, Optional ByVal pCompCd As String, Optional ByVal pLabelType As String)
    
    Dim strLabelType    As String
    
    Set AdoRs = Get_BarList(pProdCd, pCompCd, pLabelType)
    
    If AdoRs Is Nothing Then
        '등록된 정보 없음
    Else
        Do Until AdoRs.EOF
            With spdRegBar
                .MaxRows = .MaxRows + 1
                
                Call SetText(spdRegBar, AdoRs.Fields("PROD_CD").Value & "", .MaxRows, 1)
                Call SetText(spdRegBar, AdoRs.Fields("PROD_NAME").Value & "", .MaxRows, 2)
                Call SetText(spdRegBar, AdoRs.Fields("PROD_LENGTH").Value & "", .MaxRows, 3)
                strLabelType = AdoRs.Fields("BAR_GU").Value & ""
                Select Case UCase(strLabelType)
                    Case "R": Call SetText(spdRegBar, "Reel", .MaxRows, 4)
                    Case "P": Call SetText(spdRegBar, "PP Box", .MaxRows, 4)
                    Case "I": Call SetText(spdRegBar, "ICE Box", .MaxRows, 4)
                End Select
                Call SetText(spdRegBar, AdoRs.Fields("COMP_CD").Value & "", .MaxRows, 5)
                Call SetText(spdRegBar, AdoRs.Fields("COMP_NAME").Value & "", .MaxRows, 6)
                Call SetText(spdRegBar, AdoRs.Fields("BAR_CD").Value & "", .MaxRows, 7)
                If AdoRs.Fields("USED_YN").Value & "" = "Y" Then
                    Call SetText(spdRegBar, "1", .MaxRows, 15)
                Else
                    Call SetText(spdRegBar, "0", .MaxRows, 15)
                End If
                Call SetText(spdRegBar, AdoRs.Fields("REGIST_ID").Value & "", .MaxRows, 16)
                Call SetText(spdRegBar, AdoRs.Fields("REGIST_DT").Value & "", .MaxRows, 17)
                Call SetText(spdRegBar, AdoRs.Fields("MODIFY_ID").Value & "", .MaxRows, 18)
                Call SetText(spdRegBar, AdoRs.Fields("MODIFY_DT").Value & "", .MaxRows, 19)
            End With
            
            AdoRs.MoveNext
        Loop
        
    End If
    
    AdoRs.Close
    
End Sub


' 바코드 리스트 가져옴
Private Sub GetBarDetail(ByVal pProdBarCd As String)
    Dim strLabelType    As String
    Dim strItemName     As String
    Dim strItemGu       As String
    Dim strBarGu        As String
    Dim strBarType      As String
    
    Set AdoRs = Get_BarDetail(pProdBarCd)
            
    gLblDetail.LABELCD = pProdBarCd
    
    If AdoRs Is Nothing Then
        '등록된 정보 없음
    Else
        Do Until AdoRs.EOF
            With spdRegBarDetail
                .MaxRows = .MaxRows + 1
                
                Call SetText(spdRegBarDetail, AdoRs.Fields("BAR_ITEM_NO").Value & "", .MaxRows, 1)
                strItemName = Get_TempMaster_Name1("T11", AdoRs.Fields("BAR_ITEM_NO").Value & "")
                strItemGu = mGetP(strItemName, 2, "_")
                
                Call SetText(spdRegBarDetail, strItemName, .MaxRows, 2)
                Call SetText(spdRegBarDetail, AdoRs.Fields("BAR_ITEM_SEQ").Value & "", .MaxRows, 3)
                
                Select Case strItemGu
                Case "년"
                    .Row = .MaxRows
                    .Col = 4
                    .CellType = CellTypeComboBox
                    .TypeComboBoxString = "Y6"
                    .TypeComboBoxString = "Y5"
                    .TypeComboBoxString = "Y4"
                    .TypeComboBoxString = "Y3"
                    .TypeComboBoxString = "Y2"
                    .TypeComboBoxString = "Y1"
                
                Case "월"
                    .Row = .MaxRows
                    .Col = 4
                    .CellType = CellTypeComboBox
                    .TypeComboBoxString = "M3"
                    .TypeComboBoxString = "M2"
                    .TypeComboBoxString = "M1"
                
                Case "일"
                    .Row = .MaxRows
                    .Col = 4
                    .CellType = CellTypeComboBox
                    .TypeComboBoxString = "D4"
                    .TypeComboBoxString = "D3"
                    .TypeComboBoxString = "D2"
                    .TypeComboBoxString = "D1"
                Case Else
                    .Row = .MaxRows
                    .Col = 4
                    .CellType = CellTypeEdit
                    .TypeMaxEditLen = 1
                    .TypeHAlign = TypeHAlignCenter
                    .TypeVAlign = TypeVAlignCenter
                End Select
                
                Call SetText(spdRegBarDetail, AdoRs.Fields("LABEL_ITEM_TYPE").Value & "", .MaxRows, 4)
                Call SetText(spdRegBarDetail, AdoRs.Fields("BAR_CHR_NUM").Value & "", .MaxRows, 5)
                If AdoRs.Fields("USED_YN").Value & "" = "Y" Then
                    Call SetText(spdRegBarDetail, "1", .MaxRows, 11)
                Else
                    Call SetText(spdRegBarDetail, "0", .MaxRows, 11)
                End If
'                Call SetText(spdRegBarDetail, AdoRs.Fields("REGIST_ID").Value & "", .MaxRows, 16)
'                Call SetText(spdRegBarDetail, AdoRs.Fields("REGIST_DT").Value & "", .MaxRows, 17)
'                Call SetText(spdRegBarDetail, AdoRs.Fields("MODIFY_ID").Value & "", .MaxRows, 18)
'                Call SetText(spdRegBarDetail, AdoRs.Fields("MODIFY_DT").Value & "", .MaxRows, 19)
            End With
            
            AdoRs.MoveNext
        Loop
        
    End If
    
    AdoRs.Close
    
End Sub

Private Sub cmdRemove_Click()
    Dim intItemNo   As Integer
    Dim intRow      As Integer
    
    intItemNo = GetText(spdRegBarDetail, spdRegBarDetail.ActiveRow, 1)
    
    If MsgBox(GetText(spdRegBarDetail, spdRegBarDetail.ActiveRow, 2) & " 항목을 삭제하시겠습니까?", vbYesNo + vbCritical + vbDefaultButton1, Me.Caption) = vbYes Then
        'Call SetLabel_Detail(intItemNo, "DEL")
    
        If spdRegBarDetail.MaxRows > 0 Then
            Call DeleteRow(spdRegBarDetail, spdRegBarDetail.ActiveRow, spdRegBarDetail.ActiveRow)
            spdRegBarDetail.MaxRows = spdRegBarDetail.MaxRows - 1
        End If
    End If
    
    DoEvents
    
'    For intRow = 1 To spdRegBarDetail.MaxRows
'        Call SetText(spdRegBarDetail, intRow, intRow, 3)
'    Next
    
End Sub

Private Sub cmdSearch_Click()
    Dim strCompCd    As String
    Dim strProdCd    As String
    Dim strLabelCd   As String
    
    strCompCd = Trim(mGetP(cboComp.Text, 2, "|"))
    strProdCd = Trim(mGetP(cboProd.Text, 2, "|"))
    strLabelCd = UCase(Mid(cboLabel.Text, 1, 1))
    
    Call cmdClear_Click
    
    Call GetBarList(strProdCd, strCompCd, strLabelCd)
    
End Sub

Private Sub cmdSetDefault_Click()
    Dim pAdoRS      As ADODB.Recordset
    Dim intRow      As Integer
    Dim strItemName As String
    Dim strItemGu   As String
        
    intRow = 0
    strItemName = ""
    
    Set pAdoRS = New ADODB.Recordset
    
    Set pAdoRS = Get_TempMaster(mGetP(cboGubun.Text, 2, "|"))

    If pAdoRS Is Nothing Then
        '등록된 정보 없음
    Else
        With spdRegBarDetail
            .MaxRows = pAdoRS.RecordCount
        
            Do Until pAdoRS.EOF
                intRow = intRow + 1
                strItemName = pAdoRS.Fields("NAME1").Value & ""
                strItemGu = mGetP(strItemName, 2, "_")
                
                Call SetText(spdRegBarDetail, pAdoRS.Fields("CODE1").Value & "", intRow, 1)
                Call SetText(spdRegBarDetail, strItemName, intRow, 2)
                'Call SetText(spdRegBarDetail, intRow, intRow, 3)
                'Call SetText(spdRegBarDetail, Format(pAdoRS.Fields("SEQNO").Value, "00"), intRow, 3)
                Call SetText(spdRegBarDetail, pAdoRS.Fields("SEQNO").Value & "", intRow, 3)
                
                Select Case strItemGu
                Case "년"
                    .Row = intRow
                    .Col = 4
                    .CellType = CellTypeComboBox
                    .TypeComboBoxString = "Y6"
                    .TypeComboBoxString = "Y5"
                    .TypeComboBoxString = "Y4"
                    .TypeComboBoxString = "Y3"
                    .TypeComboBoxString = "Y2"
                    .TypeComboBoxString = "Y1"
                
                Case "월"
                    .Row = intRow
                    .Col = 4
                    .CellType = CellTypeComboBox
                    .TypeComboBoxString = "M3"
                    .TypeComboBoxString = "M2"
                    .TypeComboBoxString = "M1"
                
                Case "일"
                    .Row = intRow
                    .Col = 4
                    .CellType = CellTypeComboBox
                    .TypeComboBoxString = "D4"
                    .TypeComboBoxString = "D3"
                    .TypeComboBoxString = "D2"
                    .TypeComboBoxString = "D1"
                Case Else
                    .Row = intRow
                    .Col = 4
                    .CellType = CellTypeEdit
                    .TypeMaxEditLen = 1
                    .TypeHAlign = TypeHAlignCenter
                    .TypeVAlign = TypeVAlignCenter
                End Select

'                '기본값
                Call SetText(spdRegBarDetail, "1", intRow, 11)
                
                pAdoRS.MoveNext
            Loop
        End With
    End If

    pAdoRS.Close

End Sub

Private Function Get_TempMaster_Name1(ByVal pGubunCd As String, Optional pCode1 As String, Optional pCode2 As String, Optional pCode3 As String) As String
    Dim pAdoRS      As ADODB.Recordset
    
    Get_TempMaster_Name1 = ""
    
    Set pAdoRS = New ADODB.Recordset
    Set pAdoRS = Get_TempMaster(pGubunCd, pCode1)

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


Private Sub Command1_Click()
    Dim strTmp
    
    'strTmp = Get_YMD(Text2.Text, Text1.Text)

    strTmp = Get_Len(Text2.Text, Text1.Text)
    
    MsgBox strTmp
    
End Sub

Private Sub Form_Load()

    Call CtlInitializing
    
    '고객사 리스트 가져오기
    Call GetCompList_CodeName
    
    '제품 리스트 가져오기
    Call GetProdList_CodeName("", "")
    
    Call GetTempList
    
End Sub

Private Function GetTempList(Optional ByVal pCompCd As String) As String
    Dim pAdoRS      As ADODB.Recordset
    
    cboGubun.Clear
    
    Set pAdoRS = New ADODB.Recordset
    Set pAdoRS = Get_TempMaster_Gubun
    If pAdoRS Is Nothing Then
        '등록된 정보 없음
    Else
        Do Until pAdoRS.EOF
            If Mid(pAdoRS.Fields("GUBUN_CD").Value, 1, 2) = "T1" Then
                cboGubun.AddItem pAdoRS.Fields("GUBUN_MEMO").Value & Space(10) & "|" & pAdoRS.Fields("GUBUN_CD").Value
            End If
            pAdoRS.MoveNext
        Loop
    End If
    
    If pAdoRS.RecordCount > 0 Then
        cboGubun.ListIndex = 0
    End If
    
    pAdoRS.Close
    
End Function

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
    
    With spdRegBar
        Call SetText(spdRegBar, "제품코드", 0, 1):              .ColWidth(1) = 0
        Call SetText(spdRegBar, "제품명", 0, 2):                .ColWidth(2) = 12
        Call SetText(spdRegBar, "길이", 0, 3):                  .ColWidth(3) = 6
        Call SetText(spdRegBar, "구분", 0, 4):                  .ColWidth(4) = 8
        Call SetText(spdRegBar, "고객사코드", 0, 5):            .ColWidth(5) = 0
        Call SetText(spdRegBar, "고객사", 0, 6):                .ColWidth(6) = 7
        Call SetText(spdRegBar, "TYPE코드", 0, 7):            .ColWidth(7) = 9
        Call SetText(spdRegBar, "미사용", 0, 8):                .ColWidth(8) = 0
        Call SetText(spdRegBar, "미사용", 0, 9):                .ColWidth(9) = 0
        Call SetText(spdRegBar, "미사용", 0, 10):               .ColWidth(10) = 0
        Call SetText(spdRegBar, "미사용", 0, 11):               .ColWidth(11) = 0
        Call SetText(spdRegBar, "미사용", 0, 12):               .ColWidth(12) = 0
        Call SetText(spdRegBar, "미사용", 0, 13):             .ColWidth(13) = 0
        Call SetText(spdRegBar, "미사용", 0, 14):             .ColWidth(14) = 0
        Call SetText(spdRegBar, "사용여부", 0, 15):           .ColWidth(15) = 0
        Call SetText(spdRegBar, "입력자", 0, 16):             .ColWidth(16) = 0
        Call SetText(spdRegBar, "입력일시", 0, 17):           .ColWidth(17) = 0
        Call SetText(spdRegBar, "수정자", 0, 18):             .ColWidth(18) = 0
        Call SetText(spdRegBar, "수정일시", 0, 19):           .ColWidth(19) = 0
    
        .MaxRows = 0
    End With
    
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
    
    cboBarType.Clear
    cboBarType.AddItem "1D 바코드"
    cboBarType.AddItem "2D 바코드"
    cboBarType.ListIndex = 0

    
    With spdRegBarDetail
        Call SetText(spdRegBarDetail, "No", 0, 1):            .ColWidth(1) = 5
        Call SetText(spdRegBarDetail, "항목", 0, 2):          .ColWidth(2) = 20
        Call SetText(spdRegBarDetail, "순서", 0, 3):          .ColWidth(3) = 6
        Call SetText(spdRegBarDetail, "구분코드", 0, 4):      .ColWidth(4) = 20
        Call SetText(spdRegBarDetail, "자리수", 0, 5):        .ColWidth(5) = 10
        Call SetText(spdRegBarDetail, "미사용", 0, 6):        .ColWidth(6) = 0
        Call SetText(spdRegBarDetail, "미사용", 0, 7):        .ColWidth(7) = 0
        Call SetText(spdRegBarDetail, "미사용", 0, 8):        .ColWidth(8) = 0
        Call SetText(spdRegBarDetail, "미사용", 0, 9):        .ColWidth(9) = 0
        Call SetText(spdRegBarDetail, "미사용", 0, 10):       .ColWidth(10) = 0
        Call SetText(spdRegBarDetail, "사용여부", 0, 11):     .ColWidth(11) = 12
        Call SetText(spdRegBarDetail, "입력자", 0, 12):       .ColWidth(12) = 0
        Call SetText(spdRegBarDetail, "입력일시", 0, 13):     .ColWidth(13) = 0
        Call SetText(spdRegBarDetail, "수정자", 0, 14):       .ColWidth(14) = 0
        Call SetText(spdRegBarDetail, "수정일시", 0, 15):     .ColWidth(15) = 0
    
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

Private Sub spdRegBar_Click(ByVal Col As Long, ByVal Row As Long)
    Dim i               As Integer
    Dim strPrtSide      As String
    Dim strBarCd        As String
    
    If Row = 0 Then
        Call SetSpreadSort(spdRegBar)
        Exit Sub
    End If
        
    For i = 0 To cboCompCd.ListCount
        If Trim(mGetP(cboCompCd.List(i), 2, "|")) = GetText(spdRegBar, Row, 5) Then
            cboCompCd.ListIndex = i
            Exit For
        End If
    Next
    
    
    For i = 0 To cboProdCd.ListCount
        If Trim(mGetP(cboProdCd.List(i), 2, "|")) = GetText(spdRegBar, Row, 1) Then
            cboProdCd.ListIndex = i
            txtProdCd.Text = Trim(mGetP(cboProdCd.List(i), 2, "|"))
            Exit For
        End If
    Next
    
    txtProdLen.Text = GetText(spdRegBar, Row, 3)
    For i = 0 To cboLabelType.ListCount
        If cboLabelType.List(i) = GetText(spdRegBar, Row, 4) Then
            cboLabelType.ListIndex = i
            Exit For
        End If
    Next
    
    strBarCd = GetText(spdRegBar, Row, 7)
    txtBarCd.Text = strBarCd
    
    If GetText(spdRegBar, Row, 15) = "1" Then
        chkUsedYN.Value = "1"
    Else
        chkUsedYN.Value = "0"
    End If

    spdRegBarDetail.MaxRows = 0
    
    Call GetBarDetail(strBarCd)

    
End Sub

Private Sub spdRegBarDetail_Click(ByVal Col As Long, ByVal Row As Long)
    
    If Row = 0 Then
        Call SetSpreadSort(spdRegBarDetail)
        Exit Sub
    End If

End Sub

Private Sub spdRegBarDetail_ComboSelChange(ByVal Col As Long, ByVal Row As Long)
    With spdRegBarDetail
        If Col = 2 Then
            If GetText(spdRegBarDetail, Row, Col) = "바코드" Then
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


'Private Sub txtProdLabelCd_GotFocus()
'    If txtProdLabelCd.Text = "" Then
'        txtProdLabelCd.Text = Mid(cboLabelType, 1, 1) & "-"
'    End If
'
'    txtProdLabelCd.SelStart = Len(txtProdLabelCd.Text)
''    txtID.SelLength = Len(txtID.Text)
'
'End Sub

'Private Sub spdRegBarDetail_KeyPress(KeyAscii As Integer)
'
'    With spdRegBarDetail
'        If .ActiveCol = 3 Then
'            If KeyAscii = vbKeyReturn Then
'                .Row = .ActiveRow
'                .Col = 3
'                .Text = Format(.Text, "00")
'            End If
'        End If
'    End With
'End Sub
