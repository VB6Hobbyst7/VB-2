VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmMstComp 
   BackColor       =   &H00FFFFFF&
   Caption         =   "고객사설정"
   ClientHeight    =   10230
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15450
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
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10230
   ScaleWidth      =   15450
   Tag             =   "LBL_M_COMP"
   WindowState     =   2  '최대화
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '없음
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   90
      TabIndex        =   9
      Top             =   8010
      Width           =   20000
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
         Left            =   12450
         Style           =   1  '그래픽
         TabIndex        =   20
         Top             =   690
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
         Left            =   10170
         Style           =   1  '그래픽
         TabIndex        =   19
         Top             =   690
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
         Left            =   11310
         Style           =   1  '그래픽
         TabIndex        =   18
         Top             =   690
         Width           =   1095
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
         Left            =   13590
         Style           =   1  '그래픽
         TabIndex        =   17
         Top             =   690
         Width           =   1095
      End
      Begin VB.TextBox txtCompDisNo 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6570
         MaxLength       =   3
         TabIndex        =   5
         Text            =   "화성사업소"
         Top             =   780
         Width           =   620
      End
      Begin VB.TextBox txtCompCD 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00EBFBFF&
         Height          =   375
         Left            =   450
         MaxLength       =   3
         TabIndex        =   1
         Text            =   "123456"
         Top             =   780
         Width           =   1245
      End
      Begin VB.TextBox txtCompNm 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00EBFBFF&
         Height          =   375
         Left            =   1710
         MaxLength       =   30
         TabIndex        =   2
         Text            =   "마이클조단"
         Top             =   780
         Width           =   1245
      End
      Begin VB.TextBox txtCompLine 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2970
         MaxLength       =   100
         TabIndex        =   3
         Text            =   "0001"
         Top             =   780
         Width           =   1875
      End
      Begin VB.TextBox txtCompView 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4860
         MaxLength       =   10
         TabIndex        =   4
         Text            =   "화성사업소"
         Top             =   780
         Width           =   1695
      End
      Begin VB.TextBox txtCompRegID 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   8460
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   7
         Text            =   "123456"
         Top             =   780
         Width           =   1245
      End
      Begin VB.CheckBox chkCompYN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "사용"
         Height          =   255
         Left            =   7440
         TabIndex        =   6
         Top             =   840
         Width           =   795
      End
      Begin VB.Label lblComp 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '단일 고정
         Caption         =   "사용여부"
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   5
         Left            =   7200
         TabIndex        =   16
         Top             =   390
         Width           =   1245
      End
      Begin VB.Label lblComp 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '단일 고정
         Caption         =   "입력자"
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   6
         Left            =   8460
         TabIndex        =   10
         Top             =   390
         Width           =   1245
      End
      Begin VB.Label lblComp 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '단일 고정
         Caption         =   "순서"
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   4
         Left            =   6570
         TabIndex        =   11
         Top             =   390
         Width           =   615
      End
      Begin VB.Label lblComp 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '단일 고정
         Caption         =   "User Code(표기)"
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   3
         Left            =   4860
         TabIndex        =   12
         Top             =   390
         Width           =   1665
      End
      Begin VB.Label lblComp 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '단일 고정
         Caption         =   "라인"
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   2
         Left            =   2970
         TabIndex        =   13
         Top             =   390
         Width           =   1875
      End
      Begin VB.Label lblComp 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '단일 고정
         Caption         =   "대상회사"
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   1
         Left            =   1710
         TabIndex        =   14
         Top             =   390
         Width           =   1245
      End
      Begin VB.Label lblComp 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '단일 고정
         Caption         =   "고객사 코드"
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   0
         Left            =   450
         TabIndex        =   15
         Top             =   390
         Width           =   1245
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '없음
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7905
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   20000
      Begin FPSpread.vaSpread spdComp 
         Height          =   7545
         Left            =   90
         TabIndex        =   8
         Top             =   240
         Width           =   14925
         _Version        =   393216
         _ExtentX        =   26326
         _ExtentY        =   13309
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
         MaxCols         =   10
         MaxRows         =   20
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         ScrollBarExtMode=   -1  'True
         SelectBlockOptions=   0
         ShadowColor     =   16773345
         SpreadDesigner  =   "frmComp.frx":0000
         ScrollBarTrack  =   3
         ShowScrollTips  =   3
      End
   End
End
Attribute VB_Name = "frmMstComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-----------------------------------------------------------------------------'
'   파일명  : frmMstComp.frm
'   작성자  : 오세원
'   내  용  : 고객사 설정
'   작성일  : 2020-02-05
'   버  전  : 1.0.0
'   고  객  : 국도화학
'-----------------------------------------------------------------------------'

Private Sub cmdClear_Click()
    
    txtCompCd.Text = ""
    txtCompNm.Text = ""
    txtCompLine.Text = ""
    txtCompView.Text = ""
    txtCompDisNo.Text = ""
    chkCompYN.Value = "1"
    txtCompRegID.Text = gKUKDO.USERID

End Sub

Private Sub cmdClose_Click()

    Unload Me

End Sub


'-- 관리자용
Private Sub cmdDelete_Click()
    
    gComp.CD = txtCompCd.Text
    gComp.NAME = txtCompNm.Text
    gComp.LINE = txtCompLine.Text
    gComp.VIEW = txtCompView.Text
    gComp.DISNO = txtCompDisNo.Text
    
    If chkCompYN.Value = "1" Then
        gComp.YN = "Y"
    Else
        gComp.YN = "N"
    End If
    
    If Set_Comp("DEL") Then
        Call CtlInitializing
        Call GetCompList
    End If
    
End Sub

Private Sub cmdOK_Click()

    Call SetComp
    
End Sub

Private Sub Form_Load()

    Call CtlInitializing
    
    Call GetCompList
    
End Sub

Private Sub GetCompList()

    Set AdoRs = Get_CompList
    
    If AdoRs Is Nothing Then
        '등록된 정보 없음
    Else
        Do Until AdoRs.EOF
            With spdComp
                .MaxRows = .MaxRows + 1
                
                Call SetText(spdComp, AdoRs.Fields("COMP_CD").Value & "", .MaxRows, 1)
                Call SetText(spdComp, AdoRs.Fields("COMP_NAME").Value & "", .MaxRows, 2)
                Call SetText(spdComp, AdoRs.Fields("COMP_LINE").Value & "", .MaxRows, 3)
                Call SetText(spdComp, AdoRs.Fields("COMP_VIEW").Value & "", .MaxRows, 4)
                Call SetText(spdComp, AdoRs.Fields("COMP_DIS_NO").Value & "", .MaxRows, 5)
                If AdoRs.Fields("USED_YN").Value & "" = "Y" Then
                    Call SetText(spdComp, "1", .MaxRows, 6)
                Else
                    Call SetText(spdComp, "0", .MaxRows, 6)
                End If
                
                Call SetText(spdComp, AdoRs.Fields("REGIST_ID").Value & "", .MaxRows, 7)
                Call SetText(spdComp, AdoRs.Fields("REGIST_DT").Value & "", .MaxRows, 8)
                Call SetText(spdComp, AdoRs.Fields("MODIFY_ID").Value & "", .MaxRows, 9)
                Call SetText(spdComp, AdoRs.Fields("MODIFY_DT").Value & "", .MaxRows, 10)
            End With
            
            AdoRs.MoveNext
        Loop
    
    End If
    
    AdoRs.Close
    
End Sub

Private Sub SetComp()
    
    '필수입력 체크
    If txtCompCd.Text = "" Then
        MsgBox "고객사코드를 입력하세요", vbOKOnly + vbCritical, Me.Caption
        txtCompCd.SetFocus
        Exit Sub
    End If
        
    If txtCompNm.Text = "" Then
        MsgBox "고객사명을 입력하세요", vbOKOnly + vbCritical, Me.Caption
        txtCompNm.SetFocus
        Exit Sub
    End If
        
    '-- 담기
    gComp.CD = txtCompCd.Text
    gComp.NAME = txtCompNm.Text
    gComp.LINE = txtCompLine.Text
    gComp.VIEW = txtCompView.Text
    gComp.DISNO = txtCompDisNo.Text
    If chkCompYN.Value = "1" Then
        gComp.YN = "Y"
    Else
        gComp.YN = "N"
    End If
    
    '-- Insert / Update 찾아오기
    Set AdoRs = Get_CompList(txtCompCd.Text)
        
    '-- 저장
    If AdoRs.RecordCount = 0 Then
        'INSERT
        If Set_Comp("IN") Then
            Call CtlInitializing
            Call GetCompList
        End If
    Else
        'UPDATE
        If Set_Comp("UP") Then
            Call CtlInitializing
            Call GetCompList
        End If
    End If
    
End Sub

'-- 컨트롤초기화
Private Sub CtlInitializing()
    
    With spdComp
        Call SetText(spdComp, "고객사코드", 0, 1):  .ColWidth(1) = 10
        Call SetText(spdComp, "대상회사", 0, 2):    .ColWidth(2) = 10
        Call SetText(spdComp, "라인", 0, 3):        .ColWidth(3) = 16
        Call SetText(spdComp, "User Code", 0, 4):        .ColWidth(4) = 8
        Call SetText(spdComp, "순서", 0, 5):        .ColWidth(5) = 4
        Call SetText(spdComp, "사용여부", 0, 6):    .ColWidth(6) = 10
        Call SetText(spdComp, "입력자", 0, 7):      .ColWidth(7) = 10
        Call SetText(spdComp, "입력일시", 0, 8):    .ColWidth(8) = 20
        Call SetText(spdComp, "수정자", 0, 9):      .ColWidth(9) = 10
        Call SetText(spdComp, "수정일시", 0, 10):   .ColWidth(10) = 20
    
        .MaxRows = 0
    End With
    
    txtCompCd.Text = ""
    txtCompNm.Text = ""
    txtCompLine.Text = ""
    txtCompView.Text = ""
    txtCompDisNo.Text = ""
    chkCompYN.Value = "1"
    txtCompRegID.Text = gKUKDO.USERID
    
    If gKUKDO.USERGRD = "1" Then
        cmdDelete.Visible = True
    Else
        cmdDelete.Visible = False
    End If
    
    gSORT = 0
    
End Sub

'-- 사용자 선택
Private Sub spdComp_Click(ByVal Col As Long, ByVal Row As Long)

    If Row = 0 Then
        Call SetSpreadSort(spdComp)
        Exit Sub
    End If
    
    txtCompCd.Text = GetText(spdComp, Row, 1)
    txtCompNm.Text = GetText(spdComp, Row, 2)
    txtCompLine.Text = GetText(spdComp, Row, 3)
    txtCompView.Text = GetText(spdComp, Row, 4)
    txtCompDisNo.Text = GetText(spdComp, Row, 5)
    If GetText(spdComp, Row, 6) = "1" Then
        chkCompYN.Value = "1"
    Else
        chkCompYN.Value = "0"
    End If
    txtCompRegID.Text = GetText(spdComp, Row, 7)
    
End Sub

