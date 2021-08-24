VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmMaster 
   BackColor       =   &H00FFFFFF&
   Caption         =   "공통코드설정"
   ClientHeight    =   9000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17325
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9000
   ScaleWidth      =   17325
   Tag             =   "TEMP_MASTER"
   WindowState     =   2  '최대화
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '없음
      Height          =   1755
      Left            =   90
      TabIndex        =   18
      Top             =   6660
      Width           =   15915
      Begin VB.ComboBox cboProdCd 
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
         ItemData        =   "frmMaster.frx":0000
         Left            =   3060
         List            =   "frmMaster.frx":0002
         Style           =   2  '드롭다운 목록
         TabIndex        =   30
         Top             =   1290
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.TextBox txtDesc 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
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
         Left            =   9540
         MaxLength       =   100
         TabIndex        =   13
         Text            =   "화성사업소"
         Top             =   780
         Width           =   1485
      End
      Begin VB.TextBox txtGubun 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00EBFBFF&
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
         Left            =   450
         MaxLength       =   3
         TabIndex        =   5
         Text            =   "123456"
         Top             =   780
         Width           =   975
      End
      Begin VB.TextBox txtSeq 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
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
         Left            =   8640
         MaxLength       =   5
         TabIndex        =   12
         Text            =   "마이클조단"
         Top             =   780
         Width           =   885
      End
      Begin VB.TextBox txtCode1 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00EBFBFF&
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
         Left            =   1440
         MaxLength       =   5
         TabIndex        =   6
         Text            =   "0001"
         Top             =   780
         Width           =   885
      End
      Begin VB.TextBox txtCode2 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
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
         Left            =   2340
         MaxLength       =   5
         TabIndex        =   7
         Text            =   "화성사업소"
         Top             =   780
         Width           =   885
      End
      Begin VB.TextBox txtCode3 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
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
         Left            =   3240
         MaxLength       =   5
         TabIndex        =   8
         Text            =   "123456"
         Top             =   780
         Width           =   885
      End
      Begin VB.TextBox txtValue1 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
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
         Left            =   4140
         MaxLength       =   100
         TabIndex        =   9
         Text            =   "123456"
         Top             =   780
         Width           =   1485
      End
      Begin VB.TextBox txtValue2 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
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
         Left            =   5640
         MaxLength       =   100
         TabIndex        =   10
         Text            =   "123456"
         Top             =   780
         Width           =   1485
      End
      Begin VB.TextBox txtValue3 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
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
         Left            =   7140
         MaxLength       =   100
         TabIndex        =   11
         Text            =   "123456"
         Top             =   780
         Width           =   1485
      End
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
         Left            =   14610
         Style           =   1  '그래픽
         TabIndex        =   17
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdClear 
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         Caption         =   "지움"
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
         Left            =   12330
         Style           =   1  '그래픽
         TabIndex        =   15
         Top             =   720
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
         Left            =   11190
         Style           =   1  '그래픽
         TabIndex        =   14
         Top             =   720
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
         Left            =   13470
         Style           =   1  '그래픽
         TabIndex        =   16
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblComp 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '단일 고정
         Caption         =   "코드값3"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   10
         Left            =   7140
         TabIndex        =   27
         Top             =   390
         Width           =   1485
      End
      Begin VB.Label lblComp 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '단일 고정
         Caption         =   "코드값2"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   9
         Left            =   5640
         TabIndex        =   26
         Top             =   390
         Width           =   1485
      End
      Begin VB.Label lblComp 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '단일 고정
         Caption         =   "코드값1"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   8
         Left            =   4140
         TabIndex        =   25
         Top             =   390
         Width           =   1485
      End
      Begin VB.Label lblComp 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '단일 고정
         Caption         =   "코드3"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   7
         Left            =   3240
         TabIndex        =   24
         Top             =   390
         Width           =   885
      End
      Begin VB.Label lblComp 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '단일 고정
         Caption         =   "비고"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   4
         Left            =   9540
         TabIndex        =   23
         Top             =   390
         Width           =   1485
      End
      Begin VB.Label lblComp 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '단일 고정
         Caption         =   "코드2"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   3
         Left            =   2340
         TabIndex        =   22
         Top             =   390
         Width           =   885
      End
      Begin VB.Label lblComp 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '단일 고정
         Caption         =   "코드1"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   2
         Left            =   1440
         TabIndex        =   21
         Top             =   390
         Width           =   885
      End
      Begin VB.Label lblComp 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '단일 고정
         Caption         =   "순서"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   1
         Left            =   8640
         TabIndex        =   20
         Top             =   390
         Width           =   885
      End
      Begin VB.Label lblComp 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '단일 고정
         Caption         =   "구분"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   0
         Left            =   450
         TabIndex        =   19
         Top             =   390
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '없음
      Height          =   6555
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   15915
      Begin VB.CommandButton cmdDate 
         BackColor       =   &H00E0E0E0&
         Caption         =   "날짜코드범례"
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
         Left            =   14310
         Style           =   1  '그래픽
         TabIndex        =   29
         ToolTipText     =   "현재화면을 모두 지웁니다"
         Top             =   120
         Width           =   1395
      End
      Begin VB.ComboBox cboGubun 
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
         Top             =   180
         Width           =   3795
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00E0E0E0&
         Caption         =   "조회"
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
         Left            =   12030
         Style           =   1  '그래픽
         TabIndex        =   2
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdFrmClear 
         BackColor       =   &H00E0E0E0&
         Caption         =   "화면정리"
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
         Left            =   13170
         Style           =   1  '그래픽
         TabIndex        =   3
         ToolTipText     =   "현재화면을 모두 지웁니다"
         Top             =   120
         Width           =   1095
      End
      Begin FPSpread.vaSpread spdMaster 
         Height          =   5595
         Left            =   270
         TabIndex        =   4
         Top             =   690
         Width           =   15435
         _Version        =   393216
         _ExtentX        =   27226
         _ExtentY        =   9869
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
         MaxCols         =   9
         MaxRows         =   20
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         ScrollBarExtMode=   -1  'True
         SelectBlockOptions=   0
         ShadowColor     =   16774636
         SpreadDesigner  =   "frmMaster.frx":0004
         ScrollBarTrack  =   3
         ShowScrollTips  =   3
      End
      Begin VB.Label Label9 
         BackStyle       =   0  '투명
         Caption         =   "▶ 구분 "
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
         TabIndex        =   28
         Top             =   210
         Width           =   1065
      End
   End
End
Attribute VB_Name = "frmMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cboGubun_Click()
    
    Call cmdSearch_Click
    
End Sub

'-----------------------------------------------------------------------------'
'   파일명  : frmMaster.frm
'   작성자  : 오세원
'   내  용  : 공통코드 설정
'   작성일  : 2020-02-21
'   버  전  : 1.0.0
'   고  객  : 국도화학
'-----------------------------------------------------------------------------'

Private Sub cmdClear_Click()
    
    spdMaster.MaxRows = 0
    
    txtGubun.Text = ""
    txtSeq.Text = ""
    txtCode1.Text = ""
    txtCode2.Text = ""
    txtCode3.Text = ""
    txtValue1.Text = ""
    txtValue2.Text = ""
    txtValue3.Text = ""
    txtDesc.Text = ""

End Sub

Private Sub cmdClose_Click()

    Unload Me

End Sub

Private Sub cmdDate_Click()

    frmMstDateCode.Show

End Sub

Private Sub cmdDelete_Click()
    
    '-- 담기
    gTemp.GUBUN = txtGubun.Text
    gTemp.Seq = txtSeq.Text
    gTemp.CODE1 = txtCode1.Text
    gTemp.CODE2 = txtCode2.Text
    gTemp.CODE3 = txtCode3.Text
    gTemp.CDVAL1 = txtValue1.Text
    gTemp.CDVAL2 = txtValue2.Text
    gTemp.CDVAL3 = txtValue3.Text
    gTemp.DESC = txtDesc.Text
    
    If gTemp.GUBUN <> "" And gTemp.CODE1 <> "" Then
        If Set_Temp("DEL") Then
            Call CtlInitializing
            Call GetTempList(gTemp.GUBUN)
        End If
    End If
    
End Sub

Private Sub cmdFrmClear_Click()

    Call cmdClear_Click
    
End Sub

Private Sub cmdOK_Click()

    Call SetMaster

End Sub

Private Sub SetMaster()
    
    '필수입력 체크
    If txtGubun.Text = "" Then
        MsgBox "구분을 입력하세요", vbOKOnly + vbCritical, Me.Caption
        txtGubun.SetFocus
        Exit Sub
    End If
        
    If txtSeq.Text = "" Then
        MsgBox "순서를 입력하세요", vbOKOnly + vbCritical, Me.Caption
        txtSeq.SetFocus
        Exit Sub
    End If
        
    If txtCode1.Text = "" Then
        MsgBox "코드1을 입력하세요", vbOKOnly + vbCritical, Me.Caption
        txtCode1.SetFocus
        Exit Sub
    End If
        
    If txtValue1.Text = "" Then
        MsgBox "코드명1을 입력하세요", vbOKOnly + vbCritical, Me.Caption
        txtValue1.SetFocus
        Exit Sub
    End If
        
        
    '-- 담기
    gTemp.GUBUN = txtGubun.Text
    gTemp.Seq = txtSeq.Text
    gTemp.CODE1 = txtCode1.Text
    gTemp.CODE2 = txtCode2.Text
    gTemp.CODE3 = txtCode3.Text
    gTemp.CDVAL1 = txtValue1.Text
    gTemp.CDVAL2 = txtValue2.Text
    gTemp.CDVAL3 = txtValue3.Text
    gTemp.DESC = txtDesc.Text
    
    '-- Insert / Update 찾아오기
    Set AdoRs = Get_TempList(gTemp.GUBUN, gTemp.CODE1, gTemp.CODE2, gTemp.CODE3)
        
    '-- 저장
    If AdoRs.RecordCount = 0 Then
        'INSERT
        If Set_Temp("IN") Then
            Call CtlInitializing
            Call GetTempList(gTemp.GUBUN)
        End If
    Else
        'UPDATE
        If Set_Temp("UP") Then
            Call CtlInitializing
            Call GetTempList(gTemp.GUBUN)

        End If
    End If
    
End Sub

Private Sub cmdSearch_Click()
    
    Call cmdClear_Click
    
    Call GetTempList(Trim(mGetP(cboGubun.Text, 1, "|")))
    
End Sub

' 공통코드 리스트 가져옴
Private Sub GetTempList(Optional ByVal pGubunCd As String)
    Dim pAdoRS      As ADODB.Recordset
    Dim intRow      As Integer
    
    intRow = 0
    
    Set pAdoRS = New ADODB.Recordset
    
    Set pAdoRS = Get_TempMaster(pGubunCd)

    If pAdoRS Is Nothing Then
        '등록된 정보 없음
    Else
        With spdMaster
            .MaxRows = pAdoRS.RecordCount
        
            Do Until pAdoRS.EOF
                intRow = intRow + 1
                Call SetText(spdMaster, pAdoRS.Fields("GUBUN_CD").Value & "", intRow, 1)
                Call SetText(spdMaster, pAdoRS.Fields("SEQNO").Value & "", intRow, 2)
                Call SetText(spdMaster, pAdoRS.Fields("CODE1").Value & "", intRow, 3)
                Call SetText(spdMaster, pAdoRS.Fields("CODE2").Value & "", intRow, 4)
                Call SetText(spdMaster, pAdoRS.Fields("CODE3").Value & "", intRow, 5)
                Call SetText(spdMaster, pAdoRS.Fields("NAME1").Value & "", intRow, 6)
                Call SetText(spdMaster, pAdoRS.Fields("NAME2").Value & "", intRow, 7)
                Call SetText(spdMaster, pAdoRS.Fields("NAME3").Value & "", intRow, 8)
                Call SetText(spdMaster, pAdoRS.Fields("GUBUN_MEMO").Value & "", intRow, 9)
                pAdoRS.MoveNext
            Loop
        End With
    End If

    pAdoRS.Close
    
    
End Sub



Private Sub Form_Load()

    Call CtlInitializing
    
    'Call GetProdList_CodeName
    
End Sub

'Private Sub GetProdList_CodeName(Optional ByVal pProdCd As String, Optional ByVal pCompCd As String)
'    Dim pAdoRS      As ADODB.Recordset
'
'    Set pAdoRS = New ADODB.Recordset
'
'    Set pAdoRS = Get_ProdList_CodeName(pProdCd, pCompCd)
'
'    cboProdCd.Clear
'
'    If pAdoRS Is Nothing Then
'        '등록된 정보 없음
'    Else
'        'cboProdCd.AddItem "전체" & Space(50) & "|전체"
'
'        Do Until pAdoRS.EOF
'            cboProdCd.AddItem pAdoRS.Fields("PROD_NAME").Value & "-" & pAdoRS.Fields("PROD_LENGTH").Value & Space(50) & "|" & pAdoRS.Fields("PROD_CD").Value
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


'-- 컨트롤초기화
Private Sub CtlInitializing()
    Dim pAdoRS      As ADODB.Recordset
    Dim i           As Integer
    Dim intIdx      As Integer
    
    With spdMaster
'        Call SetText(spdComp, "고객사코드", 0, 1):  .ColWidth(1) = 10
'        Call SetText(spdComp, "대상회사", 0, 2):    .ColWidth(2) = 10
'        Call SetText(spdComp, "라인", 0, 3):        .ColWidth(3) = 16
'        Call SetText(spdComp, "User Code", 0, 4):        .ColWidth(4) = 8
'        Call SetText(spdComp, "순서", 0, 5):        .ColWidth(5) = 4
'        Call SetText(spdComp, "사용여부", 0, 6):    .ColWidth(6) = 10
'        Call SetText(spdComp, "입력자", 0, 7):      .ColWidth(7) = 10
'        Call SetText(spdComp, "입력일시", 0, 8):    .ColWidth(8) = 20
'        Call SetText(spdComp, "수정자", 0, 9):      .ColWidth(9) = 10
'        Call SetText(spdComp, "수정일시", 0, 10):   .ColWidth(10) = 20
    
        .MaxRows = 0
    End With
    
    txtGubun.Text = ""
    txtSeq.Text = ""
    txtCode1.Text = ""
    txtCode2.Text = ""
    txtCode3.Text = ""
    txtValue1.Text = ""
    txtValue2.Text = ""
    txtValue3.Text = ""
    txtDesc.Text = ""
    intIdx = 0
    i = 0
    
    cboGubun.Clear
    
    Set pAdoRS = New ADODB.Recordset
    Set pAdoRS = Get_TempMaster_Gubun
    If pAdoRS Is Nothing Then
        '등록된 정보 없음
    Else
        Do Until pAdoRS.EOF
            cboGubun.AddItem pAdoRS.Fields("GUBUN_CD").Value & Space(10) & "|" & pAdoRS.Fields("GUBUN_MEMO").Value
            If gTemp.GUBUN <> "" Then
                If gTemp.GUBUN = pAdoRS.Fields("GUBUN_CD").Value Then
                    intIdx = i
                End If
            End If
            pAdoRS.MoveNext
            i = i + 1
        Loop
    End If
    
    If pAdoRS.RecordCount > 0 Then
        cboGubun.ListIndex = intIdx
    End If
    
    pAdoRS.Close
    
    
End Sub

Private Sub spdMaster_Click(ByVal Col As Long, ByVal Row As Long)

    If Row = 0 Then
        Call SetSpreadSort(spdMaster)
        Exit Sub
    End If
    
    txtGubun.Text = GetText(spdMaster, Row, 1)
    txtSeq.Text = GetText(spdMaster, Row, 2)
    txtCode1.Text = GetText(spdMaster, Row, 3)
    txtCode2.Text = GetText(spdMaster, Row, 4)
    txtCode3.Text = GetText(spdMaster, Row, 5)
    txtValue1.Text = GetText(spdMaster, Row, 6)
    txtValue2.Text = GetText(spdMaster, Row, 7)
    txtValue3.Text = GetText(spdMaster, Row, 8)
    txtDesc.Text = GetText(spdMaster, Row, 9)
    

End Sub
