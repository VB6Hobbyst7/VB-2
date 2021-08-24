VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmMstUser 
   BackColor       =   &H00FFFFFF&
   Caption         =   "사용자설정"
   ClientHeight    =   11310
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16590
   BeginProperty Font 
      Name            =   "맑은 고딕"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUser.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11310
   ScaleWidth      =   16590
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
      TabIndex        =   18
      Top             =   60
      Width           =   15225
      Begin FPSpread.vaSpread spdUser 
         Height          =   7545
         Left            =   90
         TabIndex        =   19
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
         ShadowColor     =   16774120
         SpreadDesigner  =   "frmUser.frx":27A2
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
      Height          =   1395
      Left            =   90
      TabIndex        =   0
      Top             =   8010
      Width           =   15225
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
         Left            =   13560
         Style           =   1  '그래픽
         TabIndex        =   20
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
         Left            =   11280
         Style           =   1  '그래픽
         TabIndex        =   9
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
         Left            =   10140
         Style           =   1  '그래픽
         TabIndex        =   8
         Top             =   690
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
         Left            =   12420
         Style           =   1  '그래픽
         TabIndex        =   10
         Top             =   690
         Width           =   1095
      End
      Begin VB.CheckBox chkUsedYN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "사용"
         Height          =   255
         Left            =   7110
         TabIndex        =   6
         Top             =   810
         Width           =   765
      End
      Begin VB.ComboBox cboUserComp 
         Height          =   375
         ItemData        =   "frmUser.frx":31C4
         Left            =   5580
         List            =   "frmUser.frx":31CE
         TabIndex        =   5
         Text            =   "사용자"
         Top             =   780
         Width           =   1215
      End
      Begin VB.TextBox txtUserRegID 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   8160
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   7
         Text            =   "123456"
         Top             =   780
         Width           =   1245
      End
      Begin VB.TextBox txtUserDepart 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4290
         MaxLength       =   30
         TabIndex        =   4
         Text            =   "화성사업소"
         Top             =   780
         Width           =   1245
      End
      Begin VB.TextBox txtUserPW 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00EBFBFF&
         Height          =   375
         Left            =   3000
         MaxLength       =   20
         TabIndex        =   3
         Text            =   "0001"
         Top             =   780
         Width           =   1245
      End
      Begin VB.TextBox txtUserNm 
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
      Begin VB.TextBox txtUserID 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00EBFBFF&
         Height          =   375
         Left            =   420
         MaxLength       =   6
         TabIndex        =   1
         Text            =   "123456"
         Top             =   780
         Width           =   1245
      End
      Begin VB.Label lblUser 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '단일 고정
         Caption         =   "입력자"
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   6
         Left            =   8160
         TabIndex        =   17
         Top             =   390
         Width           =   1245
      End
      Begin VB.Label lblUser 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '단일 고정
         Caption         =   "사용여부"
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   5
         Left            =   6870
         TabIndex        =   16
         Top             =   390
         Width           =   1245
      End
      Begin VB.Label lblUser 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '단일 고정
         Caption         =   "사용자권한"
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   4
         Left            =   5580
         TabIndex        =   15
         Top             =   390
         Width           =   1245
      End
      Begin VB.Label lblUser 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '단일 고정
         Caption         =   "부서"
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   3
         Left            =   4290
         TabIndex        =   14
         Top             =   390
         Width           =   1245
      End
      Begin VB.Label lblUser 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '단일 고정
         Caption         =   "비밀번호"
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   2
         Left            =   3000
         TabIndex        =   13
         Top             =   390
         Width           =   1245
      End
      Begin VB.Label lblUser 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '단일 고정
         Caption         =   "사용자명"
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   1
         Left            =   1710
         TabIndex        =   12
         Top             =   390
         Width           =   1245
      End
      Begin VB.Label lblUser 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '단일 고정
         Caption         =   "사용자ID"
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   0
         Left            =   420
         TabIndex        =   11
         Top             =   390
         Width           =   1245
      End
   End
End
Attribute VB_Name = "frmMstUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-----------------------------------------------------------------------------'
'   파일명  : frmMstUser.frm
'   작성자  : 오세원
'   내  용  : 사용자 설정
'   작성일  : 2020-02-04
'   버  전  : 1.0.0
'   고  객  : 국도화학
'-----------------------------------------------------------------------------'

Private Sub cmdClear_Click()
    
    txtUserID.Text = ""
    txtUserNm.Text = ""
    txtUserPW.Text = ""
    txtUserDepart.Text = ""
    cboUserComp.ListIndex = 0
    chkUsedYN.Value = "1"
    txtUserRegID.Text = gKUKDO.USERID

End Sub

Private Sub cmdClose_Click()

    Unload Me

End Sub


'-- 관리자용
Private Sub cmdDelete_Click()
    
    gUSER.ID = txtUserID.Text
    gUSER.NAME = txtUserNm.Text
    gUSER.PW = txtUserPW.Text
    gUSER.DEPT = txtUserDepart.Text
    
    If cboUserComp.Text = "사용자" Then
        gUSER.COMP = "2"
    Else
        gUSER.COMP = "1"
    End If
    
    If chkUsedYN.Value = "1" Then
        gUSER.YN = "Y"
    Else
        gUSER.YN = "N"
    End If
    
    If Set_User("DEL") Then
        Call CtlInitializing
        Call GetUserList
    End If
    
End Sub

Private Sub cmdOK_Click()

    Call SetUser
    
End Sub

Private Sub Form_Load()

    Call CtlInitializing
    
    Call GetUserList
    
End Sub

Private Sub GetUserList()

    Set AdoRs = Get_UserList
    
    If AdoRs Is Nothing Then
        '등록된 정보 없음
    Else
        Do Until AdoRs.EOF
            With spdUser
                .MaxRows = .MaxRows + 1
                
                Call SetText(spdUser, AdoRs.Fields("USER_CD").Value & "", .MaxRows, 1)
                Call SetText(spdUser, AdoRs.Fields("USER_NAME").Value & "", .MaxRows, 2)
                Call SetText(spdUser, AdoRs.Fields("USER_PW").Value & "", .MaxRows, 3)
                Call SetText(spdUser, AdoRs.Fields("USER_DEPART").Value & "", .MaxRows, 4)
                
                If AdoRs.Fields("USER_COMP").Value & "" = "1" Then
                    Call SetText(spdUser, "관리자", .MaxRows, 5)
                Else
                    Call SetText(spdUser, "사용자", .MaxRows, 5)
                End If
                
                If AdoRs.Fields("USED_YN").Value & "" = "Y" Then
                    Call SetText(spdUser, "1", .MaxRows, 6)
                Else
                    Call SetText(spdUser, "0", .MaxRows, 6)
                End If
                
                Call SetText(spdUser, AdoRs.Fields("REGIST_ID").Value & "", .MaxRows, 7)
                Call SetText(spdUser, AdoRs.Fields("REGIST_DT").Value & "", .MaxRows, 8)
                Call SetText(spdUser, AdoRs.Fields("MODIFY_ID").Value & "", .MaxRows, 9)
                Call SetText(spdUser, AdoRs.Fields("MODIFY_DT").Value & "", .MaxRows, 10)
            End With
            
            AdoRs.MoveNext
        Loop
    
    End If
    
    AdoRs.Close
    
End Sub

Private Sub SetUser()
    
    '필수입력 체크
    If txtUserID.Text = "" Then
        MsgBox "사용자ID를 입력하세요", vbOKOnly + vbCritical, Me.Caption
        txtUserID.SetFocus
        Exit Sub
    End If
        
    If txtUserNm.Text = "" Then
        MsgBox "사용자명을 입력하세요", vbOKOnly + vbCritical, Me.Caption
        txtUserNm.SetFocus
        Exit Sub
    End If
        
    If txtUserPW.Text = "" Then
        MsgBox "사용자 비밀번호를 입력하세요", vbOKOnly + vbCritical, Me.Caption
        txtUserPW.SetFocus
        Exit Sub
    End If
        
    '-- 담기
    gUSER.ID = txtUserID.Text
    gUSER.NAME = txtUserNm.Text
    gUSER.PW = txtUserPW.Text
    gUSER.DEPT = txtUserDepart.Text
    If cboUserComp.Text = "사용자" Then
        gUSER.COMP = "2"
    Else
        gUSER.COMP = "1"
    End If
    If chkUsedYN.Value = "1" Then
        gUSER.YN = "Y"
    Else
        gUSER.YN = "N"
    End If
    
    '-- Insert / Update 찾아오기
    Set AdoRs = Get_UserList(txtUserID.Text)
        
    '-- 저장
    If AdoRs.RecordCount = 0 Then
        'INSERT
        If Set_User("IN") Then
            Call CtlInitializing
            Call GetUserList
        End If
    Else
        'UPDATE
        If Set_User("UP") Then
            Call CtlInitializing
            Call GetUserList
        End If
    End If
    
End Sub

'-- 컨트롤초기화
Private Sub CtlInitializing()
    
    With spdUser
        Call SetText(spdUser, "사용자ID", 0, 1):    .ColWidth(1) = 10
        Call SetText(spdUser, "사용자명", 0, 2):    .ColWidth(2) = 10
        Call SetText(spdUser, "비번", 0, 3):        .ColWidth(3) = 8
        Call SetText(spdUser, "부서", 0, 4):        .ColWidth(4) = 8
        Call SetText(spdUser, "권한", 0, 5):        .ColWidth(5) = 8
        Call SetText(spdUser, "사용여부", 0, 6):    .ColWidth(6) = 10
        Call SetText(spdUser, "입력자", 0, 7):      .ColWidth(7) = 10
        Call SetText(spdUser, "입력일시", 0, 8):    .ColWidth(8) = 20
        Call SetText(spdUser, "수정자", 0, 9):      .ColWidth(9) = 10
        Call SetText(spdUser, "수정일시", 0, 10):   .ColWidth(10) = 20
    
        .MaxRows = 0
    End With
    
    txtUserID.Text = ""
    txtUserNm.Text = ""
    txtUserPW.Text = ""
    txtUserDepart.Text = ""
    cboUserComp.ListIndex = 0
    chkUsedYN.Value = "1"
    txtUserRegID.Text = gKUKDO.USERID
    
    If gKUKDO.USERGRD = "1" Then
        cmdDelete.Visible = True
    Else
        cmdDelete.Visible = False
    End If
    
    gSORT = 0

End Sub

'-- 사용자 선택
Private Sub spdUser_Click(ByVal Col As Long, ByVal Row As Long)

    If Row = 0 Then
        Call SetSpreadSort(spdUser)
        Exit Sub
    End If
    
    txtUserID.Text = GetText(spdUser, Row, 1)
    txtUserNm.Text = GetText(spdUser, Row, 2)
    txtUserPW.Text = GetText(spdUser, Row, 3)
    txtUserDepart.Text = GetText(spdUser, Row, 4)
    If GetText(spdUser, Row, 5) = "사용자" Then
        cboUserComp.ListIndex = 0
    Else
        cboUserComp.ListIndex = 1
    End If
    If GetText(spdUser, Row, 6) = "1" Then
        chkUsedYN.Value = "1"
    Else
        chkUsedYN.Value = "0"
    End If
    txtUserRegID.Text = GetText(spdUser, Row, 7)
    
End Sub
