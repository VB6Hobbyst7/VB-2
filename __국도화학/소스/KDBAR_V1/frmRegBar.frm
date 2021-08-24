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
      Height          =   10455
      Left            =   90
      TabIndex        =   6
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
         Height          =   10005
         Left            =   5220
         TabIndex        =   7
         Top             =   300
         Width           =   12645
         Begin VB.ComboBox cboLabelType 
            Height          =   375
            ItemData        =   "frmRegBar.frx":0000
            Left            =   2670
            List            =   "frmRegBar.frx":0002
            Style           =   2  '드롭다운 목록
            TabIndex        =   31
            Top             =   1620
            Width           =   3735
         End
         Begin VB.CheckBox chkUsedYN 
            BackColor       =   &H00FFFFFF&
            Caption         =   "사용"
            Height          =   255
            Left            =   2730
            TabIndex        =   21
            Top             =   2520
            Width           =   795
         End
         Begin VB.TextBox txtProdLen 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   10230
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   20
            Text            =   "화성사업소"
            Top             =   780
            Width           =   1140
         End
         Begin VB.TextBox txtCompCd 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   2670
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   19
            Text            =   "화성사업소"
            Top             =   1200
            Width           =   3705
         End
         Begin VB.TextBox txtProdCd 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   6420
            MaxLength       =   5
            TabIndex        =   18
            Text            =   "화성사업소"
            Top             =   780
            Width           =   2055
         End
         Begin VB.ComboBox cboBarType 
            Height          =   375
            ItemData        =   "frmRegBar.frx":0004
            Left            =   2670
            List            =   "frmRegBar.frx":0006
            Style           =   2  '드롭다운 목록
            TabIndex        =   17
            Top             =   2040
            Width           =   3735
         End
         Begin VB.TextBox txtBarCd 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   2670
            MaxLength       =   10
            TabIndex        =   16
            Text            =   "0"
            Top             =   360
            Width           =   3720
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
            Left            =   8070
            TabIndex        =   12
            Top             =   7800
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
               TabIndex        =   15
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
               TabIndex        =   14
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
               TabIndex        =   13
               Top             =   150
               Width           =   1095
            End
         End
         Begin VB.CommandButton cmdAdd 
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            Caption         =   "(+) 항목추가"
            Height          =   405
            Left            =   420
            Style           =   1  '그래픽
            TabIndex        =   11
            Top             =   3090
            Width           =   1395
         End
         Begin VB.CommandButton cmdRemove 
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            Caption         =   "(-) 항목제거"
            Height          =   405
            Left            =   1830
            Style           =   1  '그래픽
            TabIndex        =   10
            Top             =   3090
            Width           =   1395
         End
         Begin VB.ComboBox cboProdCd 
            Height          =   375
            ItemData        =   "frmRegBar.frx":0008
            Left            =   2670
            List            =   "frmRegBar.frx":000A
            Style           =   2  '드롭다운 목록
            TabIndex        =   9
            Top             =   780
            Width           =   3735
         End
         Begin VB.TextBox txtCompNm 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   6420
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   8
            Text            =   "화성사업소"
            Top             =   1200
            Width           =   2055
         End
         Begin FPSpread.vaSpread spdRegBarDetail 
            Height          =   4125
            Left            =   420
            TabIndex        =   30
            Top             =   3540
            Width           =   11055
            _Version        =   393216
            _ExtentX        =   19500
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
            MaxCols         =   10
            MaxRows         =   20
            RetainSelBlock  =   0   'False
            ScrollBarMaxAlign=   0   'False
            ScrollBars      =   2
            ScrollBarShowMax=   0   'False
            ShadowColor     =   16775150
            SpreadDesigner  =   "frmRegBar.frx":000C
            ScrollBarTrack  =   3
            ShowScrollTips  =   3
         End
         Begin VB.Label lblUser 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "제품구분"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   2
            Left            =   420
            TabIndex        =   32
            Top             =   1620
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
            Left            =   11400
            TabIndex        =   28
            Top             =   840
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
            Left            =   420
            TabIndex        =   27
            Top             =   780
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
            Left            =   420
            TabIndex        =   26
            Top             =   1200
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
            Left            =   420
            TabIndex        =   25
            Top             =   2460
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
            Left            =   8700
            TabIndex        =   24
            Top             =   780
            Width           =   1500
         End
         Begin VB.Label lblUser 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
            Caption         =   "바코드타입"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   1
            Left            =   420
            TabIndex        =   23
            Top             =   2040
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
            Left            =   420
            TabIndex        =   22
            Top             =   360
            Width           =   2205
         End
      End
      Begin FPSpread.vaSpread spdRegBar 
         Height          =   8235
         Left            =   210
         TabIndex        =   29
         Top             =   390
         Width           =   4935
         _Version        =   393216
         _ExtentX        =   8705
         _ExtentY        =   14526
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
         SpreadDesigner  =   "frmRegBar.frx":091C
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
      Begin VB.TextBox txtComp 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   7530
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   4
         Top             =   300
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00E0E0E0&
         Caption         =   "화면정리"
         Height          =   375
         Left            =   6330
         Style           =   1  '그래픽
         TabIndex        =   3
         ToolTipText     =   "현재화면을 모두 지웁니다"
         Top             =   300
         Width           =   1095
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00E0E0E0&
         Caption         =   "조회"
         Height          =   375
         Left            =   5220
         Style           =   1  '그래픽
         TabIndex        =   2
         Top             =   300
         Width           =   1095
      End
      Begin VB.ComboBox cboComp 
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
         Left            =   1350
         Style           =   2  '드롭다운 목록
         TabIndex        =   1
         Top             =   360
         Width           =   3795
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
         TabIndex        =   5
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
'Option Explicit
''-----------------------------------------------------------------------------'
''   파일명  : frmRegBar.frm
''   작성자  : 오세원
''   내  용  : 바코드 등록
''   작성일  : 2020-02-14
''   버  전  : 1.0.0
''   고  객  : 국도화학
''-----------------------------------------------------------------------------'
'
'
'Private Sub cboComp_Click()
'
'    txtProdCd.Text = ""
'    txtComp.Text = mGetP(cboComp.Text, 2, "|")
'
'    Call GetProdList("", cboComp.Text)
'
''    Call GetProdList_CodeName("", txtComp.Text)
'
'End Sub
'
'
'Private Sub cboProdCd_Click()
'
'    txtProdCd.Text = Trim(mGetP(cboProdCd.Text, 2, "|"))
'
'    Call GetComp_CodeName(txtProdCd.Text)
'
'End Sub
'
''-- 항목추가
'Private Sub cmdAdd_Click()
'    Dim intRow      As Integer
'    Dim intNum      As Integer
'    Dim intMaxNum   As Integer
'
'    intMaxNum = 0
'    With spdRegBarDetail
'        For intRow = 1 To .MaxRows
'            intNum = GetText(spdRegBarDetail, intRow, 1)
'            If intMaxNum < intNum Then
'                intMaxNum = intNum
'            End If
'        Next
'    End With
'    intMaxNum = intMaxNum + 1
'
'    spdRegBarDetail.MaxRows = spdRegBarDetail.MaxRows + 1
'
'    'NO
'    Call SetText(spdRegBarDetail, intMaxNum, spdRegBarDetail.MaxRows, 1)
'    '순서
'    Call SetText(spdRegBarDetail, CStr(spdRegBarDetail.MaxRows), spdRegBarDetail.MaxRows, 2)
'    '항목명
'    Call SetText(spdRegBarDetail, "", spdRegBarDetail.MaxRows, 3)
'    '자리수
'    Call SetText(spdRegBarDetail, "0", spdRegBarDetail.MaxRows, 4)
'    '항목구분
'    Call SetText(spdRegBarDetail, "", spdRegBarDetail.MaxRows, 5)
'    '사용구분
'    Call SetText(spdRegBarDetail, "1", spdRegBarDetail.MaxRows, 6)
'
'End Sub
'
'Private Sub cmdClear_Click()
'
'    spdRegBar.MaxRows = 0
'    spdRegBarDetail.MaxRows = 0
'
'    '-- 1 Line
'    txtProdCd.Text = ""
'    'txtProdNm.Text = ""
'    txtCompCd.Text = ""
'    txtProdLen.Text = ""
'
'    cboLabelType.Clear
'    cboLabelType.AddItem "Reel"
'    cboLabelType.AddItem "PP Box"
'    cboLabelType.AddItem "ICE Box"
'    cboLabelType.ListIndex = 0
'
'    cboBarType.Clear
'    cboBarType.AddItem "1D 바코드"
'    cboBarType.AddItem "QR 코드"
'    cboBarType.ListIndex = 0
'
'End Sub
'
'Private Sub cmdClose_Click()
'
'    Unload Me
'
'End Sub
'
'Private Sub cmdDelete_Click()
'
''    Call SetLabel
'
'End Sub
'
'Private Sub cmdOK_Click()
'
'    Call SetLabel
'
'End Sub
'
'Private Sub SetLabel()
'    Dim intRow      As Integer
'    Dim intCol      As Integer
'    Dim intItemNo   As Integer
'
'    '필수입력 체크
'    If txtBarCd.Text = "" Then
'        MsgBox "바코드타입코드를 입력하세요", vbOKOnly + vbCritical, Me.Caption
'        txtBarCd.SetFocus
'        Exit Sub
'    End If
'
'    If txtCompCd.Text = "" Then
'        MsgBox "고객사를 선택하세요", vbOKOnly + vbCritical, Me.Caption
'        txtCompCd.SetFocus
'        Exit Sub
'    End If
'
'    If txtProdLen.Text = "" Then
'        MsgBox "제품길이를 입력하세요", vbOKOnly + vbCritical, Me.Caption
'        txtProdLen.SetFocus
'        Exit Sub
'    End If
'
''    If txtLabelPrtNo.Text = "" Then
''        MsgBox "라벨출력수량을 입력하세요", vbOKOnly + vbCritical, Me.Caption
''        txtLabelPrtNo.SetFocus
''        Exit Sub
''    End If
'
'
'    '-- 담기
'    gBarInfo.BARCD = txtBarCd.Text
'    gBarInfo.PRODCD = txtProdCd.Text
'    gBarInfo.COMPCD = txtCompCd.Text
'    gBarInfo.LBLTYPE = Mid(cboLabelType.Text, 1, 1)
'    gBarInfo.BARTYPE = Mid(cboBarType.Text, 1, 1)
'
'    With spdRegBarDetail
'        gBarMst.BARCD = txtProdCd.Text                 'Key
'
'        ReDim gBarMst.BARITEMNO(.MaxRows) As String    'Key
'        ReDim gBarMst.BARITEMSEQ(.MaxRows) As String
'        ReDim gBarMst.BARITEMNAME(.MaxRows) As String
'        ReDim gBarMst.BARCHRNUM(.MaxRows) As String
'        ReDim gBarMst.LBLITEMTYPE(.MaxRows) As String
'        ReDim gBarMst.YN(.MaxRows) As String
'
'        For intRow = 1 To .DataRowCnt
'            gBarMst.BARITEMNO(intRow) = GetText(spdRegBarDetail, intRow, 1)
'            gBarMst.BARITEMSEQ(intRow) = GetText(spdRegBarDetail, intRow, 2)
'            gBarMst.BARITEMNAME(intRow) = GetText(spdRegBarDetail, intRow, 3)
'            gBarMst.BARCHRNUM(intRow) = GetText(spdRegBarDetail, intRow, 4)
'            gBarMst.LBLITEMTYPE(intRow) = GetText(spdRegBarDetail, intRow, 5)
'            gBarMst.YN(intRow) = IIf(GetText(spdRegBarDetail, intRow, 6) = "1", "Y", "N")
'        Next
'    End With
'
'    If chkUsedYN.Value = "1" Then
'        gBarInfo.YN = "Y"
'    Else
'        gBarInfo.YN = "N"
'    End If
'
'    '-- Insert / Update 찾아오기
'    Set AdoRs = Get_BarList(txtProdCd.Text, txtCompCd.Text, Mid(cboLabelType.Text, 1, 1))
'
'    '-- 저장
'    If AdoRs.RecordCount = 0 Then
'        'INSERT
'        If Set_Bar("IN") Then
'            '상세내용 저장
'            For intRow = 1 To spdRegBarDetail.DataRowCnt
'                intItemNo = GetText(spdRegBarDetail, intRow, 1)
'                Call SetBar_Master(intItemNo)
'            Next
'
'            'Call CtlInitializing
'            'Call GetProdList
'        End If
'    Else
'        'UPDATE
'        If Set_Bar("UP") Then
'            '상세내용 저장
'            For intRow = 1 To spdRegBarDetail.DataRowCnt
'                intItemNo = GetText(spdRegBarDetail, intRow, 1)
'                Call SetBar_Master(intItemNo)
'            Next
'
'            'Call CtlInitializing
'            'Call GetProdList
'        End If
'    End If
'
'
'End Sub
'
'Private Sub SetLabel_Master(ByVal pItemNo As Integer, Optional pDelFlag As String)
''    Dim intRow      As Integer
''    Dim intCol      As Integer
'
'    '-- Insert / Update 찾아오기
'    Set AdoRs = Get_LabelMasterList(txtProdCd.Text, txtCompCd.Text, gLblInfo.LBLTYPE, pItemNo)
'
'    '-- 저장
'    If AdoRs.RecordCount = 0 Then
'        'INSERT
'        If Set_Label_Master("IN", pItemNo) Then
''            Call CtlInitializing
''            Call GetProdList
'        End If
'    Else
'        If pDelFlag = "DEL" Then
'            'DELETE
'            If Set_Label_Master("DEL", pItemNo) Then
'    '            Call CtlInitializing
'    '            Call GetProdList
'            End If
'        Else
'            'UPDATE
'            If Set_Label_Master("UP", pItemNo) Then
'    '            Call CtlInitializing
'    '            Call GetProdList
'            End If
'        End If
'    End If
'
'End Sub
'
'Private Sub SetBar_Master(ByVal pItemNo As Integer, Optional pDelFlag As String)
''    Dim intRow      As Integer
''    Dim intCol      As Integer
'
'    '-- Insert / Update 찾아오기
'    Set AdoRs = Get_LabelMasterList(txtProdCd.Text, txtCompCd.Text, gLblInfo.LBLTYPE, pItemNo)
'
'    '-- 저장
'    If AdoRs.RecordCount = 0 Then
'        'INSERT
'        If Set_Label_Master("IN", pItemNo) Then
''            Call CtlInitializing
''            Call GetProdList
'        End If
'    Else
'        If pDelFlag = "DEL" Then
'            'DELETE
'            If Set_Label_Master("DEL", pItemNo) Then
'    '            Call CtlInitializing
'    '            Call GetProdList
'            End If
'        Else
'            'UPDATE
'            If Set_Label_Master("UP", pItemNo) Then
'    '            Call CtlInitializing
'    '            Call GetProdList
'            End If
'        End If
'    End If
'
'End Sub
'
'Private Sub GetProdList_CodeName(Optional ByVal pProdCd As String, Optional ByVal pCompCd As String)
'    Dim pAdoRS      As ADODB.Recordset
'
'    Set pAdoRS = New ADODB.Recordset
'
'    Set pAdoRS = Get_ProdList_CodeName(pProdCd, pCompCd)
'
'    If pAdoRS Is Nothing Then
'        '등록된 정보 없음
'    Else
'        cboProdCd.Clear
'        Do Until pAdoRS.EOF
'            cboProdCd.AddItem pAdoRS.Fields("PROD_NAME").Value & Space(50) & "|" & pAdoRS.Fields("PROD_CD").Value
'            pAdoRS.MoveNext
'        Loop
'
'        If pAdoRS.RecordCount > 0 Then
'            cboProdCd.ListIndex = 0
'        End If
'    End If
'
'    pAdoRS.Close
'
'End Sub
'
''화면에는 안보임(Hidden처리)
'Private Sub GetProdList(Optional ByVal pProdCd As String, Optional ByVal pCompCd As String)
'
'    Dim strCompNm       As String
'
'    Set AdoRs = Get_ProdList(pProdCd, pCompCd)
'
'    If AdoRs Is Nothing Then
'        '등록된 정보 없음
'    Else
'        Do Until AdoRs.EOF
'            With spdRegProd
'                .MaxRows = .MaxRows + 1
'
'                Call SetText(spdRegProd, AdoRs.Fields("PROD_CD").Value & "", .MaxRows, 1)
'                Call SetText(spdRegProd, AdoRs.Fields("COMP_CD").Value & "", .MaxRows, 2)
'                strCompNm = GetCompList_Name(AdoRs.Fields("COMP_CD").Value & "")
'                Call SetText(spdRegProd, strCompNm, .MaxRows, 3)
'                Call SetText(spdRegProd, AdoRs.Fields("PROD_NAME").Value & "", .MaxRows, 4)
'                Call SetText(spdRegProd, AdoRs.Fields("PROD_LENGTH").Value & "", .MaxRows, 5)
'                Call SetText(spdRegProd, AdoRs.Fields("PROD_MATERIAL_CD").Value & "", .MaxRows, 6)
'                Call SetText(spdRegProd, AdoRs.Fields("EXPIR_MONTH").Value & "", .MaxRows, 7)
'                Call SetText(spdRegProd, AdoRs.Fields("PROD_STOR_TEMP").Value & "", .MaxRows, 8)
'                Call SetText(spdRegProd, AdoRs.Fields("PROD_SIZE").Value & "", .MaxRows, 9)
'                Call SetText(spdRegProd, AdoRs.Fields("PROD_CHIMEI_PN").Value & "", .MaxRows, 10)
'                Call SetText(spdRegProd, AdoRs.Fields("VENDER_CD").Value & "", .MaxRows, 11)
'                Call SetText(spdRegProd, AdoRs.Fields("PROD_LINE_FA").Value & "", .MaxRows, 12)
'                Call SetText(spdRegProd, AdoRs.Fields("PROD_SLIT_FA").Value & "", .MaxRows, 13)
'                Call SetText(spdRegProd, AdoRs.Fields("PROD_CONTROL_YN").Value & "", .MaxRows, 14)
'                Call SetText(spdRegProd, AdoRs.Fields("PROD_PCN_NO").Value & "", .MaxRows, 15)
'                Call SetText(spdRegProd, AdoRs.Fields("ITEM_BARCODE").Value & "", .MaxRows, 16)
'                If AdoRs.Fields("USED_YN").Value & "" = "Y" Then
'                    Call SetText(spdRegLabel, "1", .MaxRows, 17)
'                Else
'                    Call SetText(spdRegLabel, "0", .MaxRows, 17)
'                End If
'                Call SetText(spdRegProd, AdoRs.Fields("REGIST_ID").Value & "", .MaxRows, 18)
'                Call SetText(spdRegProd, AdoRs.Fields("REGIST_DT").Value & "", .MaxRows, 19)
'                Call SetText(spdRegProd, AdoRs.Fields("MODIFY_ID").Value & "", .MaxRows, 20)
'                Call SetText(spdRegProd, AdoRs.Fields("MODIFY_DT").Value & "", .MaxRows, 21)
'            End With
'
'            AdoRs.MoveNext
'        Loop
'
'    End If
'
'    AdoRs.Close
'
'    spdRegProd.RowHeight(0) = 12
'    spdRegProd.RowHeight(-1) = 12
'
'End Sub
'
'' 라벨 리스트 가져옴
'Private Sub GetLabelList(Optional ByVal pProdCd As String, Optional ByVal pCompCd As String, Optional ByVal pLabelType As String)
'
'    Dim strLabelType    As String
'
'    Set AdoRs = Get_LabelList(pProdCd, pCompCd, pLabelType)
'
''        Call SetText(spdRegLabel, "제품코드", 0, 1):            .ColWidth(1) = 0
''        Call SetText(spdRegLabel, "제품명", 0, 2):              .ColWidth(1) = 10
''        Call SetText(spdRegLabel, "제품길이", 0, 3):            .ColWidth(4) = 8
''        Call SetText(spdRegLabel, "제품타입", 0, 4):            .ColWidth(5) = 8
''        Call SetText(spdRegLabel, "고객사코드", 0, 5):          .ColWidth(2) = 0
''        Call SetText(spdRegLabel, "고객사명", 0, 6):            .ColWidth(3) = 8
''        Call SetText(spdRegLabel, "출력매수", 0, 7):            .ColWidth(6) = 8
''        Call SetText(spdRegLabel, "박스당릴기본수량", 0, 8):    .ColWidth(7) = 10
''        Call SetText(spdRegLabel, "측면출력구분", 0, 9):        .ColWidth(8) = 10
''        Call SetText(spdRegLabel, "측면상1타입", 0, 10):         .ColWidth(9) = 10
''        Call SetText(spdRegLabel, "측면상2타입", 0, 11):        .ColWidth(10) = 8
''        Call SetText(spdRegLabel, "측면3타입", 0, 12):          .ColWidth(11) = 8
''        Call SetText(spdRegLabel, "측면4타입", 0, 13):          .ColWidth(12) = 8
''        Call SetText(spdRegLabel, "포장최대수량", 0, 14):       .ColWidth(13) = 10
''        Call SetText(spdRegLabel, "사용여부", 0, 15):           .ColWidth(14) = 10
''        Call SetText(spdRegLabel, "입력자", 0, 16):             .ColWidth(15) = 10
''        Call SetText(spdRegLabel, "입력일시", 0, 17):           .ColWidth(16) = 10
''        Call SetText(spdRegLabel, "수정자", 0, 18):             .ColWidth(17) = 10
''        Call SetText(spdRegLabel, "수정일시", 0, 19):           .ColWidth(18) = 10
'
'    If AdoRs Is Nothing Then
'        '등록된 정보 없음
'    Else
'        Do Until AdoRs.EOF
'            With spdRegLabel
'                .MaxRows = .MaxRows + 1
'
'                Call SetText(spdRegLabel, AdoRs.Fields("PROD_CD").Value & "", .MaxRows, 1)
'                Call SetText(spdRegLabel, AdoRs.Fields("PROD_NAME").Value & "", .MaxRows, 2)
'                Call SetText(spdRegLabel, AdoRs.Fields("PROD_LENGTH").Value & "", .MaxRows, 3)
'                strLabelType = AdoRs.Fields("PROD_LABEL_TYPE").Value & ""
'                Select Case UCase(strLabelType)
'                    Case "R": Call SetText(spdRegLabel, "Reel", .MaxRows, 4)
'                    Case "P": Call SetText(spdRegLabel, "PP Box", .MaxRows, 4)
'                    Case "I": Call SetText(spdRegLabel, "ICE Box", .MaxRows, 4)
'                End Select
'                Call SetText(spdRegLabel, AdoRs.Fields("COMP_CD").Value & "", .MaxRows, 5)
'                'strCompNm = GetCompList_Name(AdoRs.Fields("COMP_CD").Value & "")
'                'Call SetText(spdRegLabel, strCompNm, .MaxRows, 3)
'                Call SetText(spdRegLabel, AdoRs.Fields("COMP_NAME").Value & "", .MaxRows, 6)
'                Call SetText(spdRegLabel, AdoRs.Fields("LABEL_PRT_NO").Value & "", .MaxRows, 7)
'                Call SetText(spdRegLabel, AdoRs.Fields("LABEL_PRT_DEFAULT_NO").Value & "", .MaxRows, 8)
'                Call SetText(spdRegLabel, AdoRs.Fields("LABEL_PRT_SIDE").Value & "", .MaxRows, 9)
'                Call SetText(spdRegLabel, AdoRs.Fields("LABEL_BAR_SIDE01_TYPE").Value & "", .MaxRows, 10)
'                Call SetText(spdRegLabel, AdoRs.Fields("LABEL_BAR_SIDE02_TYPE").Value & "", .MaxRows, 11)
'                Call SetText(spdRegLabel, AdoRs.Fields("LABEL_BAR_SIDE03_TYPE").Value & "", .MaxRows, 12)
'                Call SetText(spdRegLabel, AdoRs.Fields("LABEL_BAR_SIDE04_TYPE").Value & "", .MaxRows, 13)
'                Call SetText(spdRegLabel, AdoRs.Fields("PROD_MAX_TOT").Value & "", .MaxRows, 14)
'                If AdoRs.Fields("USED_YN").Value & "" = "Y" Then
'                    Call SetText(spdRegLabel, "1", .MaxRows, 15)
'                Else
'                    Call SetText(spdRegLabel, "0", .MaxRows, 15)
'                End If
'                Call SetText(spdRegLabel, AdoRs.Fields("REGIST_ID").Value & "", .MaxRows, 16)
'                Call SetText(spdRegLabel, AdoRs.Fields("REGIST_DT").Value & "", .MaxRows, 17)
'                Call SetText(spdRegLabel, AdoRs.Fields("MODIFY_ID").Value & "", .MaxRows, 18)
'                Call SetText(spdRegLabel, AdoRs.Fields("MODIFY_DT").Value & "", .MaxRows, 19)
'            End With
'
'            AdoRs.MoveNext
'        Loop
'
'    End If
'
'    AdoRs.Close
'
'End Sub
'
'' 라벨 리스트 가져옴
'Private Sub GetLabelMaster(Optional ByVal pProdCd As String, Optional ByVal pCompCd As String, Optional ByVal pLabelType As String)
'
'    Dim strLabelType    As String
'    Dim strItemName     As String
'    Dim strBarGu        As String
'    Dim strBarType      As String
'
'    Set AdoRs = Get_LabelMaster(pProdCd, pCompCd, pLabelType)
'
'    gLblMst.PRODCD = pProdCd
'    gLblMst.COMPCD = pCompCd
'    gLblMst.LBLTYPE = pLabelType
'
'    If AdoRs Is Nothing Then
'        '등록된 정보 없음
'    Else
'        Do Until AdoRs.EOF
'            With spdRegLabelDetail
'                .MaxRows = .MaxRows + 1
'
'                Call SetText(spdRegLabelDetail, AdoRs.Fields("LABEL_ITEM_NO").Value & "", .MaxRows, 1)
'                strItemName = AdoRs.Fields("LABEL_ITEM_NAME").Value & ""
'                Call SetText(spdRegLabelDetail, strItemName, .MaxRows, 2)
'                Call SetText(spdRegLabelDetail, AdoRs.Fields("LABEL_ITEM_SEQ").Value & "", .MaxRows, 3)
'                Call SetText(spdRegLabelDetail, AdoRs.Fields("LABEL_ITEM_MEMO").Value & "", .MaxRows, 4)
'                Call SetText(spdRegLabelDetail, IIf(AdoRs.Fields("LABEL_ITEM_NAME_PRT").Value & "" = "Y", "예", "아니오"), .MaxRows, 5)
'                Call SetText(spdRegLabelDetail, AdoRs.Fields("LABEL_ITEM_X_COORD").Value & "", .MaxRows, 6)
'                Call SetText(spdRegLabelDetail, AdoRs.Fields("LABEL_ITEM_Y_COORD").Value & "", .MaxRows, 7)
'
'                strBarGu = AdoRs.Fields("LABEL_ITEM_BAR_GU").Value & ""
'                If strItemName = "바코드" Then
'                    If strBarGu = "1" Then
'                        strBarType = "1D 바코드"
'                    ElseIf strBarGu = "2" Then
'                        strBarType = "2D 바코드"
'                    Else
'                        strBarType = strBarGu
'                    End If
'                Else
'                    strBarType = strBarGu
'                End If
'
'                If strItemName = "바코드" Then
'                    .Row = .MaxRows
'                    .Col = 8
'                    .CellType = CellTypeComboBox
'                    .TypeComboBoxString = "2D 바코드"
'                    .TypeComboBoxString = "1D 바코드"
'                    .Text = "1D 바코드"
'                Else
'                    .Row = .MaxRows
'                    .Col = 8
'                    .CellType = CellTypeEdit
'                    .TypeMaxEditLen = 1
'                    .TypeHAlign = TypeHAlignCenter
'                    .TypeVAlign = TypeVAlignCenter
'
'                    Call SetText(spdRegLabelDetail, strBarType, .MaxRows, 8)
'                End If
'
'
'
'                Call SetText(spdRegLabelDetail, AdoRs.Fields("LABEL_ITEM_FONTSIZE").Value & "", .MaxRows, 9)
'                Call SetText(spdRegLabelDetail, AdoRs.Fields("LABEL_ITEM_ROT").Value & "", .MaxRows, 10)
'                If AdoRs.Fields("USED_YN").Value & "" = "Y" Then
'                    Call SetText(spdRegLabelDetail, "1", .MaxRows, 11)
'                Else
'                    Call SetText(spdRegLabelDetail, "0", .MaxRows, 11)
'                End If
''                Call SetText(spdRegLabelDetail, AdoRs.Fields("REGIST_ID").Value & "", .MaxRows, 16)
''                Call SetText(spdRegLabelDetail, AdoRs.Fields("REGIST_DT").Value & "", .MaxRows, 17)
''                Call SetText(spdRegLabelDetail, AdoRs.Fields("MODIFY_ID").Value & "", .MaxRows, 18)
''                Call SetText(spdRegLabelDetail, AdoRs.Fields("MODIFY_DT").Value & "", .MaxRows, 19)
'            End With
'
'            AdoRs.MoveNext
'        Loop
'
'    End If
'
'    AdoRs.Close
'
'End Sub
'
'Private Sub cmdRemove_Click()
'    Dim intItemNo   As Integer
'
'    intItemNo = GetText(spdRegLabelDetail, spdRegLabelDetail.ActiveRow, 1)
'
'    If MsgBox(GetText(spdRegLabelDetail, spdRegLabelDetail.ActiveRow, 2) & " 항목을 삭제하시겠습니까?", vbYesNo + vbCritical + vbDefaultButton1, Me.Caption) = vbYes Then
'        Call SetLabel_Master(intItemNo, "DEL")
'
'        If spdRegLabelDetail.MaxRows > 0 Then
'            Call DeleteRow(spdRegLabelDetail, spdRegLabelDetail.ActiveRow, spdRegLabelDetail.ActiveRow)
'            spdRegLabelDetail.MaxRows = spdRegLabelDetail.MaxRows - 1
'        End If
'    End If
'
'End Sub
'
'Private Sub cmdSearch_Click()
'
'    If txtComp.Text = "" Then
'        Exit Sub
'    End If
'
'    Call cmdClear_Click
'
'    Call GetProdList("", txtComp.Text)
'
'    Call GetProdList_CodeName("", txtComp.Text)
'
'    Call GetLabelList("", txtComp.Text)
'
'
''''    txtProdCd.Text = ""
''''    txtComp.Text = mGetP(cboComp.Text, 2, "|")
''''
''''    Call GetProdList("", cboComp.Text)
''''
''''    Call GetProdList_CodeName("", txtComp.Text)
'
'End Sub
'
'Private Sub Form_Load()
'
'    Call CtlInitializing
'
'    '고객사 리스트 가져오기
'    Call GetCompList_CodeName
'
'    '제품 리스트 가져오기
''    Call GetProdList_CodeName("", txtComp.Text)
'
'End Sub
'
'
'Private Function GetCompList_Name(Optional ByVal pCompCd As String) As String
'    Dim pAdoRS      As ADODB.Recordset
'
'    Set pAdoRS = New ADODB.Recordset
'
'    Set pAdoRS = Get_CompList_Name(pCompCd)
'
'    If pAdoRS Is Nothing Then
'        '등록된 정보 없음
'    Else
'        Do Until pAdoRS.EOF
'            GetCompList_Name = pAdoRS.Fields("COMP_NAME").Value & ""
'
'            pAdoRS.MoveNext
'        Loop
'
'    End If
'
'    pAdoRS.Close
'
'End Function
'
''-- 상단 고객사리스트 가져오기
'Private Sub GetCompList_CodeName()
'    Dim pAdoRS      As ADODB.Recordset
'
'    Set pAdoRS = New ADODB.Recordset
'
'    Set pAdoRS = Get_CompList_CodeName
'
'    If pAdoRS Is Nothing Then
'        '등록된 정보 없음
'    Else
'        cboComp.Clear
'
'        cboComp.AddItem "전체" & Space(30) & "|" & "전체"
'
'        Do Until pAdoRS.EOF
'            'cboComp.AddItem pAdoRS.Fields("COMP_NAME").Value & Space(20) & "|" & pAdoRS.Fields("COMP_CD").Value & ""
'            cboComp.AddItem pAdoRS.Fields("COMP_NAME").Value & Space(15 - Len(pAdoRS.Fields("COMP_NAME").Value)) & pAdoRS.Fields("COMP_LINE").Value & Space(20) & "|" & pAdoRS.Fields("COMP_CD").Value & ""
'
'            pAdoRS.MoveNext
'        Loop
'
'        If pAdoRS.RecordCount > 0 Then
'            cboComp.ListIndex = 0
'        End If
'    End If
'
'    pAdoRS.Close
'
'End Sub
'
''-- 제품선택했을때 해당 고객사 가져오기
'Private Sub GetComp_CodeName(ByVal pProdCd As String)
'    Dim pAdoRS      As ADODB.Recordset
'
'    Set pAdoRS = New ADODB.Recordset
'
'    Set pAdoRS = Get_Comp_CodeName(pProdCd)
'
'    If pAdoRS Is Nothing Then
'        '등록된 정보 없음
'    Else
'        txtCompCd.Text = ""
'        txtCompNm.Text = ""
'        txtProdLen.Text = ""
'
'        Do Until pAdoRS.EOF
'            txtCompNm.Text = pAdoRS.Fields("COMP_NAME").Value & ""
'            txtCompCd.Text = pAdoRS.Fields("COMP_CD").Value & ""
'            txtProdLen.Text = pAdoRS.Fields("PROD_LENGTH").Value & ""
'
'            pAdoRS.MoveNext
'        Loop
'
'    End If
'
'    pAdoRS.Close
'
'End Sub
'
'
''-- 컨트롤초기화
'Private Sub CtlInitializing()
'
'    With spdRegProd
'        Call SetText(spdRegProd, "제품코드", 0, 1):         .ColWidth(1) = 0
'        Call SetText(spdRegProd, "고객사코드", 0, 2):       .ColWidth(2) = 0
'        Call SetText(spdRegProd, "고객사명", 0, 3):         .ColWidth(3) = 8
'        Call SetText(spdRegProd, "제품명", 0, 4):           .ColWidth(4) = 11
'        Call SetText(spdRegProd, "제품길이", 0, 5):         .ColWidth(5) = 10
'        Call SetText(spdRegProd, "자재코드", 0, 6):         .ColWidth(6) = 8
'        Call SetText(spdRegProd, "유효기간", 0, 7):         .ColWidth(7) = 10
'        Call SetText(spdRegProd, "보관온도", 0, 8):         .ColWidth(8) = 10
'        Call SetText(spdRegProd, "사이즈", 0, 9):           .ColWidth(9) = 8
'        Call SetText(spdRegProd, "CHIMEI코드", 0, 10):      .ColWidth(10) = 8
'        Call SetText(spdRegProd, "VENDOR코드", 0, 11):      .ColWidth(11) = 8
'        Call SetText(spdRegProd, "제조라인공장", 0, 12):    .ColWidth(12) = 10
'        Call SetText(spdRegProd, "SLITTING공장", 0, 13):    .ColWidth(13) = 10
'        Call SetText(spdRegProd, "관리선이탈여부", 0, 14):  .ColWidth(14) = 8
'        Call SetText(spdRegProd, "PCN차수", 0, 15):         .ColWidth(15) = 8
'        Call SetText(spdRegProd, "바코드", 0, 16):          .ColWidth(16) = 8
'        Call SetText(spdRegProd, "사용여부", 0, 17):        .ColWidth(17) = 10
'        Call SetText(spdRegProd, "입력자", 0, 18):          .ColWidth(18) = 10
'        Call SetText(spdRegProd, "입력일시", 0, 19):        .ColWidth(19) = 10
'        Call SetText(spdRegProd, "수정자", 0, 20):          .ColWidth(20) = 10
'        Call SetText(spdRegProd, "수정일시", 0, 21):        .ColWidth(21) = 10
'
'        .MaxRows = 0
'    End With
'
'    With spdRegLabel
'        Call SetText(spdRegLabel, "제품코드", 0, 1):            .ColWidth(1) = 0
'        Call SetText(spdRegLabel, "제품명", 0, 2):              .ColWidth(2) = 12
'        Call SetText(spdRegLabel, "제품길이", 0, 3):            .ColWidth(3) = 8
'        Call SetText(spdRegLabel, "제품타입", 0, 4):            .ColWidth(4) = 8
'        Call SetText(spdRegLabel, "고객사코드", 0, 5):          .ColWidth(5) = 0
'        Call SetText(spdRegLabel, "고객사명", 0, 6):            .ColWidth(6) = 9
'        Call SetText(spdRegLabel, "출력매수", 0, 7):            .ColWidth(7) = 8
'        Call SetText(spdRegLabel, "박스당릴기본수량", 0, 8):    .ColWidth(8) = 10
'        Call SetText(spdRegLabel, "측면출력구분", 0, 9):        .ColWidth(9) = 10
'        Call SetText(spdRegLabel, "측면상1타입", 0, 10):        .ColWidth(10) = 10
'        Call SetText(spdRegLabel, "측면상2타입", 0, 11):        .ColWidth(11) = 8
'        Call SetText(spdRegLabel, "측면3타입", 0, 12):          .ColWidth(12) = 8
'        Call SetText(spdRegLabel, "측면4타입", 0, 13):          .ColWidth(13) = 8
'        Call SetText(spdRegLabel, "포장최대수량", 0, 14):       .ColWidth(14) = 10
'        Call SetText(spdRegLabel, "사용여부", 0, 15):           .ColWidth(15) = 10
'        Call SetText(spdRegLabel, "입력자", 0, 16):             .ColWidth(16) = 10
'        Call SetText(spdRegLabel, "입력일시", 0, 17):           .ColWidth(17) = 10
'        Call SetText(spdRegLabel, "수정자", 0, 18):             .ColWidth(18) = 10
'        Call SetText(spdRegLabel, "수정일시", 0, 19):           .ColWidth(19) = 10
'
'        .MaxRows = 0
'    End With
'
'    '-- 1 Line
'    txtProdCd.Text = ""
''    txtProdNm.Text = ""
'    txtCompCd.Text = ""
'    txtCompNm.Text = ""
'    txtProdLen.Text = ""
'    cboLabelType.Clear
'    cboLabelType.AddItem "Reel"
'    cboLabelType.AddItem "PP Box"
'    cboLabelType.AddItem "ICE Box"
'    cboLabelType.ListIndex = 0
'
'    cboPrtSide.Clear
'    cboPrtSide.AddItem "아니오"
'    cboPrtSide.AddItem "예"
'    cboPrtSide.ListIndex = 0
'
'    '-- 2 Line
'    txtLabelPrtNo.Text = "0"
'
'    '-- 3 Line
'    With spdRegLabelDetail
'        Call SetText(spdRegLabelDetail, "No", 0, 1):          .ColWidth(1) = 5
'        Call SetText(spdRegLabelDetail, "항목", 0, 2):        .ColWidth(2) = 10
'        Call SetText(spdRegLabelDetail, "순서", 0, 3):        .ColWidth(3) = 6
'        Call SetText(spdRegLabelDetail, "내용", 0, 4):        .ColWidth(4) = 20
'        Call SetText(spdRegLabelDetail, "항목출력", 0, 5):    .ColWidth(5) = 10
'        Call SetText(spdRegLabelDetail, "X좌표", 0, 6):       .ColWidth(6) = 8
'        Call SetText(spdRegLabelDetail, "Y좌표", 0, 7):       .ColWidth(7) = 8
'        Call SetText(spdRegLabelDetail, "구분자", 0, 8):      .ColWidth(8) = 8
'        Call SetText(spdRegLabelDetail, "폰트크기", 0, 9):    .ColWidth(9) = 8
'        Call SetText(spdRegLabelDetail, "회전", 0, 10):        .ColWidth(10) = 8
'        Call SetText(spdRegLabelDetail, "사용여부", 0, 11):   .ColWidth(11) = 12
'        Call SetText(spdRegLabelDetail, "입력자", 0, 12):     .ColWidth(12) = 0
'        Call SetText(spdRegLabelDetail, "입력일시", 0, 13):   .ColWidth(13) = 0
'        Call SetText(spdRegLabelDetail, "수정자", 0, 14):     .ColWidth(14) = 0
'        Call SetText(spdRegLabelDetail, "수정일시", 0, 15):   .ColWidth(15) = 0
'
'        .MaxRows = 0
'    End With
'
'    chkUsedYN.Value = "1"
'    If gKUKDO.USERGRD = "1" Then
'        cmdDelete.Visible = True
'    Else
'        cmdDelete.Visible = False
'    End If
'
'    gSORT = 0
'
'End Sub
'
'
'Private Sub Label9_DblClick()
'
'    If spdRegProd.Visible = True Then
'        spdRegProd.Visible = False
'    Else
'        spdRegProd.Visible = True
'    End If
'
'End Sub
'
'Private Sub spdRegLabel_Click(ByVal Col As Long, ByVal Row As Long)
'    Dim i           As Integer
'    Dim strPrtSide  As String
'
'    If Row = 0 Then
'        Call SetSpreadSort(spdRegLabel)
'        Exit Sub
'    End If
'
'    For i = 0 To cboProdCd.ListCount
'        If Trim(mGetP(cboProdCd.List(i), 2, "|")) = GetText(spdRegLabel, Row, 1) Then
'            cboProdCd.ListIndex = i
'            Exit For
'        End If
'    Next
'
'    '콤보박스에서 처리되므로 불필요
'    'txtProdCd.Text = GetText(spdRegLabel, Row, 1)
'
'    txtProdLen.Text = GetText(spdRegLabel, Row, 3)
'    For i = 0 To cboLabelType.ListCount
'        If cboLabelType.List(i) = GetText(spdRegLabel, Row, 4) Then
'            cboLabelType.ListIndex = i
'            Exit For
'        End If
'    Next
'
'    txtLabelPrtNo.Text = GetText(spdRegLabel, Row, 7)
'    txtLabelPrtDefaultNo.Text = GetText(spdRegLabel, Row, 8)
'
'    strPrtSide = GetText(spdRegLabel, Row, 9)
'    If strPrtSide = "Y" Then
'        strPrtSide = "예"
'    Else
'        strPrtSide = "아니오"
'    End If
'    For i = 0 To cboPrtSide.ListCount
'        If cboPrtSide.List(i) = strPrtSide Then
'            cboPrtSide.ListIndex = i
'            Exit For
'        End If
'    Next
'
'    If GetText(spdRegLabel, Row, 15) = "1" Then
'        chkUsedYN.Value = "1"
'    Else
'        chkUsedYN.Value = "0"
'    End If
'
'    spdRegLabelDetail.MaxRows = 0
'
'    Call GetLabelMaster(txtProdCd.Text, txtCompCd.Text, Mid(cboLabelType.Text, 1, 1))
'
'
'End Sub
'
'
'
'Private Sub spdRegLabelDetail_ComboSelChange(ByVal Col As Long, ByVal Row As Long)
'    With spdRegLabelDetail
'        If Col = 2 Then
'            If GetText(spdRegLabelDetail, Row, Col) = "바코드" Then
'                .Row = Row
'                .Col = 8
'                .CellType = CellTypeComboBox
'
'                .TypeComboBoxString = "2D 바코드"
'                .TypeComboBoxString = "1D 바코드"
'                .Text = "1D 바코드"
'            Else
'
'            End If
'        End If
'    End With
'End Sub
'
