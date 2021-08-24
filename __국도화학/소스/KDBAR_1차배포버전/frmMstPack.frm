VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmMstPack 
   BackColor       =   &H00FFFFFF&
   Caption         =   "포장설정"
   ClientHeight    =   9990
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19335
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
   ScaleHeight     =   9990
   ScaleWidth      =   19335
   Tag             =   "LBL_M_PACK"
   WindowState     =   2  '최대화
   Begin VB.Frame Frame2 
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
      Height          =   7905
      Left            =   90
      TabIndex        =   19
      Top             =   60
      Width           =   20000
      Begin FPSpread.vaSpread spdPack 
         Height          =   7545
         Left            =   90
         TabIndex        =   20
         Top             =   240
         Width           =   18195
         _Version        =   393216
         _ExtentX        =   32094
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
         MaxCols         =   14
         MaxRows         =   20
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         ScrollBarExtMode=   -1  'True
         SelectBlockOptions=   0
         ShadowColor     =   16774636
         SpreadDesigner  =   "frmMstPack.frx":0000
         ScrollBarTrack  =   3
         ShowScrollTips  =   3
      End
   End
   Begin VB.Frame Frame1 
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
      Height          =   1755
      Left            =   90
      TabIndex        =   0
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
         Left            =   16020
         Style           =   1  '그래픽
         TabIndex        =   28
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
         Left            =   13740
         Style           =   1  '그래픽
         TabIndex        =   27
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
         Left            =   14880
         Style           =   1  '그래픽
         TabIndex        =   26
         Top             =   720
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
         Left            =   17160
         Style           =   1  '그래픽
         TabIndex        =   25
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtPackCatGbn 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   8100
         MaxLength       =   50
         TabIndex        =   8
         Text            =   "123456"
         Top             =   780
         Width           =   1485
      End
      Begin VB.TextBox txtPackProLen 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6600
         MaxLength       =   20
         TabIndex        =   7
         Text            =   "123456"
         Top             =   780
         Width           =   1485
      End
      Begin VB.TextBox txtPackProWidth 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   5100
         MaxLength       =   20
         TabIndex        =   6
         Text            =   "123456"
         Top             =   780
         Width           =   1485
      End
      Begin VB.TextBox txtPackCatWidth 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3600
         MaxLength       =   20
         TabIndex        =   5
         Text            =   "123456"
         Top             =   780
         Width           =   1485
      End
      Begin VB.CheckBox chkPackYN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "사용"
         Height          =   255
         Left            =   10440
         TabIndex        =   10
         Top             =   840
         Width           =   795
      End
      Begin VB.TextBox txtPackRegID 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   11490
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   11
         Text            =   "123456"
         Top             =   780
         Width           =   1245
      End
      Begin VB.TextBox txtPackDia 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2970
         MaxLength       =   10
         TabIndex        =   4
         Text            =   "화성사업소"
         Top             =   780
         Width           =   620
      End
      Begin VB.TextBox txtPackCore 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2340
         MaxLength       =   10
         TabIndex        =   3
         Text            =   "0001"
         Top             =   780
         Width           =   620
      End
      Begin VB.TextBox txtPackNm 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00EBFBFF&
         Height          =   375
         Left            =   1440
         MaxLength       =   5
         TabIndex        =   2
         Text            =   "마이클조단"
         Top             =   780
         Width           =   885
      End
      Begin VB.TextBox txtPackCD 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00EBFBFF&
         Height          =   375
         Left            =   450
         MaxLength       =   3
         TabIndex        =   1
         Text            =   "123456"
         Top             =   780
         Width           =   975
      End
      Begin VB.TextBox txtPackDisNo 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   9600
         MaxLength       =   3
         TabIndex        =   9
         Text            =   "화성사업소"
         Top             =   780
         Width           =   620
      End
      Begin VB.Label lblComp 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '단일 고정
         Caption         =   "포장코드"
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   0
         Left            =   450
         TabIndex        =   18
         Top             =   390
         Width           =   975
      End
      Begin VB.Label lblComp 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '단일 고정
         Caption         =   "포장명"
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   1
         Left            =   1440
         TabIndex        =   17
         Top             =   390
         Width           =   885
      End
      Begin VB.Label lblComp 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '단일 고정
         Caption         =   "Core"
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   2
         Left            =   2340
         TabIndex        =   16
         Top             =   390
         Width           =   615
      End
      Begin VB.Label lblComp 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '단일 고정
         Caption         =   "Dia"
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   3
         Left            =   2970
         TabIndex        =   15
         Top             =   390
         Width           =   615
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
         Left            =   9600
         TabIndex        =   14
         Top             =   390
         Width           =   615
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
         Left            =   10230
         TabIndex        =   13
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
         Left            =   11490
         TabIndex        =   12
         Top             =   390
         Width           =   1245
      End
      Begin VB.Label lblComp 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '단일 고정
         Caption         =   "Width(카트리지)"
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   7
         Left            =   3600
         TabIndex        =   21
         Top             =   390
         Width           =   1485
      End
      Begin VB.Label lblComp 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '단일 고정
         Caption         =   $"frmMstPack.frx":0BDD
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   8
         Left            =   5100
         TabIndex        =   22
         Top             =   390
         Width           =   1485
      End
      Begin VB.Label lblComp 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '단일 고정
         Caption         =   "Length"
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   9
         Left            =   6600
         TabIndex        =   23
         Top             =   390
         Width           =   1485
      End
      Begin VB.Label lblComp 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '단일 고정
         Caption         =   "카트리지 구분"
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   10
         Left            =   8100
         TabIndex        =   24
         Top             =   390
         Width           =   1485
      End
   End
End
Attribute VB_Name = "frmMstPack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-----------------------------------------------------------------------------'
'   파일명  : frmMstPack.frm
'   작성자  : 오세원
'   내  용  : 포장코드 설정
'   작성일  : 2020-02-06
'   버  전  : 1.0.0
'   고  객  : 국도화학
'-----------------------------------------------------------------------------'

Private Sub cmdClear_Click()
    
    txtPackCD.Text = ""
    txtPackNm.Text = ""
    txtPackCore.Text = ""
    txtPackDia.Text = ""
    txtPackCatWidth.Text = ""
    txtPackProWidth.Text = ""
    txtPackProLen.Text = ""
    txtPackCatGbn.Text = ""
    txtPackDisNo.Text = ""
    chkPackYN.Value = "1"
    txtPackRegID.Text = gKUKDO.USERID

End Sub

Private Sub cmdClose_Click()

    Unload Me

End Sub


'-- 관리자용
Private Sub cmdDelete_Click()
    
    gPack.CD = txtPackCD.Text
    gPack.NAME = txtPackNm.Text
    gPack.CORE = txtPackCore.Text
    gPack.DIA = txtPackDia.Text
    gPack.CWID = txtPackCatWidth.Text
    gPack.PWID = txtPackProWidth.Text
    gPack.pLen = txtPackProLen.Text
    gPack.CGBN = txtPackCatGbn.Text
    gPack.DISNO = txtPackDisNo.Text
    
    If chkPackYN.Value = "1" Then
        gPack.YN = "Y"
    Else
        gPack.YN = "N"
    End If
    
    If AdoRs.RecordCount <> 0 Then
        If Set_Pack("DEL") Then
            Call CtlInitializing
            Call GetPackList
        End If
    End If
    
End Sub

Private Sub cmdOK_Click()

    Call SetPack
    
End Sub

Private Sub Form_Load()

    Call CtlInitializing
    
    Call GetPackList
    
End Sub

Private Sub GetPackList()

    Set AdoRs = Get_PackList
    
    If AdoRs Is Nothing Then
        '등록된 정보 없음
    Else
        Do Until AdoRs.EOF
            With spdPack
                .MaxRows = .MaxRows + 1
                
                Call SetText(spdPack, AdoRs.Fields("PACK_CD").Value & "", .MaxRows, 1)
                Call SetText(spdPack, AdoRs.Fields("PACK_NAME").Value & "", .MaxRows, 2)
                Call SetText(spdPack, AdoRs.Fields("PACK_CORE").Value & "", .MaxRows, 3)
                Call SetText(spdPack, AdoRs.Fields("PACK_DIA").Value & "", .MaxRows, 4)
                Call SetText(spdPack, AdoRs.Fields("PACK_CAT_WIDTH").Value & "", .MaxRows, 5)
                Call SetText(spdPack, AdoRs.Fields("PACK_PRO_WIDTH").Value & "", .MaxRows, 6)
                Call SetText(spdPack, AdoRs.Fields("PACK_PRO_LENGTH").Value & "", .MaxRows, 7)
                Call SetText(spdPack, AdoRs.Fields("PACK_CAT_GU").Value & "", .MaxRows, 8)
                Call SetText(spdPack, AdoRs.Fields("PACK_DIS_NO").Value & "", .MaxRows, 9)
                If AdoRs.Fields("USED_YN").Value & "" = "Y" Then
                    Call SetText(spdPack, "1", .MaxRows, 10)
                Else
                    Call SetText(spdPack, "0", .MaxRows, 10)
                End If
                
                Call SetText(spdPack, AdoRs.Fields("REGIST_ID").Value & "", .MaxRows, 11)
                Call SetText(spdPack, AdoRs.Fields("REGIST_DT").Value & "", .MaxRows, 12)
                Call SetText(spdPack, AdoRs.Fields("MODIFY_ID").Value & "", .MaxRows, 13)
                Call SetText(spdPack, AdoRs.Fields("MODIFY_DT").Value & "", .MaxRows, 14)
            End With
            
            AdoRs.MoveNext
        Loop
        
    End If
    
    AdoRs.Close

End Sub

Private Sub SetPack()
    
    '필수입력 체크
    If txtPackCD.Text = "" Then
        MsgBox "제품코드를 입력하세요", vbOKOnly + vbCritical, Me.Caption
        txtPackCD.SetFocus
        Exit Sub
    End If
        
    If txtPackNm.Text = "" Then
        MsgBox "No(제품명)를 입력하세요", vbOKOnly + vbCritical, Me.Caption
        txtPackNm.SetFocus
        Exit Sub
    End If
        
    '-- 담기
    gPack.CD = txtPackCD.Text
    gPack.NAME = txtPackNm.Text
    gPack.CORE = txtPackCore.Text
    gPack.DIA = txtPackDia.Text
    gPack.CWID = txtPackCatWidth.Text
    gPack.PWID = txtPackProWidth.Text
    gPack.pLen = txtPackProLen.Text
    gPack.CGBN = txtPackCatGbn.Text
    gPack.DISNO = txtPackDisNo.Text
    If chkPackYN.Value = "1" Then
        gPack.YN = "Y"
    Else
        gPack.YN = "N"
    End If
                
    '-- Insert / Update 찾아오기
    Set AdoRs = Get_PackList(txtPackCD.Text)
        
    '-- 저장
    If AdoRs.RecordCount = 0 Then
        'INSERT
        If Set_Pack("IN") Then
            Call CtlInitializing
            Call GetPackList
        End If
    Else
        'UPDATE
        If Set_Pack("UP") Then
            Call CtlInitializing
            Call GetPackList
        End If
    End If
    
End Sub

'-- 컨트롤초기화
Private Sub CtlInitializing()
    
    With spdPack
        .MaxCols = 14
        Call SetText(spdPack, "포장코드", 0, 1):  .ColWidth(1) = 8
        Call SetText(spdPack, "포장명", 0, 2):    .ColWidth(2) = 8
        Call SetText(spdPack, "Core", 0, 3):        .ColWidth(3) = 6
        Call SetText(spdPack, "Dia", 0, 4):        .ColWidth(4) = 6
        Call SetText(spdPack, "Width(카트리지)", 0, 5):        .ColWidth(5) = 12
        Call SetText(spdPack, "Width(제품)", 0, 6):    .ColWidth(6) = 12
        Call SetText(spdPack, "Length", 0, 7):      .ColWidth(7) = 12
        Call SetText(spdPack, "카트리지 구분", 0, 8):    .ColWidth(8) = 14
        Call SetText(spdPack, "순서", 0, 9):      .ColWidth(9) = 4
        Call SetText(spdPack, "사용여부", 0, 10):      .ColWidth(10) = 10
        Call SetText(spdPack, "입력자", 0, 11):      .ColWidth(11) = 10
        Call SetText(spdPack, "입력일시", 0, 12):   .ColWidth(12) = 10
        Call SetText(spdPack, "수정자", 0, 13):      .ColWidth(13) = 10
        Call SetText(spdPack, "수정일시", 0, 14):   .ColWidth(14) = 10
    
        .MaxRows = 0
    End With
    
    txtPackCD.Text = ""
    txtPackNm.Text = ""
    txtPackCore.Text = ""
    txtPackDia.Text = ""
    txtPackCatWidth.Text = ""
    txtPackProWidth.Text = ""
    txtPackProLen.Text = ""
    txtPackCatGbn.Text = ""
    txtPackDisNo.Text = ""
    chkPackYN.Value = "1"
    txtPackRegID.Text = gKUKDO.USERID
    
    If gKUKDO.USERGRD = "1" Then
        cmdDelete.Visible = True
    Else
        cmdDelete.Visible = False
    End If
    
    gSORT = 0
    
End Sub

'-- 사용자 선택
Private Sub spdPack_Click(ByVal Col As Long, ByVal Row As Long)

    If Row = 0 Then
        Call SetSpreadSort(spdPack)
        Exit Sub
    End If
    
    txtPackCD.Text = GetText(spdPack, Row, 1)
    txtPackNm.Text = GetText(spdPack, Row, 2)
    txtPackCore.Text = GetText(spdPack, Row, 3)
    txtPackDia.Text = GetText(spdPack, Row, 4)
    txtPackCatWidth.Text = GetText(spdPack, Row, 5)
    txtPackProWidth.Text = GetText(spdPack, Row, 6)
    txtPackProLen.Text = GetText(spdPack, Row, 7)
    txtPackCatGbn.Text = GetText(spdPack, Row, 8)
    
    txtPackDisNo.Text = GetText(spdPack, Row, 9)
    
    If GetText(spdPack, Row, 10) = "1" Then
        chkPackYN.Value = "1"
    Else
        chkPackYN.Value = "0"
    End If
    txtPackRegID.Text = GetText(spdPack, Row, 11)
    
End Sub


