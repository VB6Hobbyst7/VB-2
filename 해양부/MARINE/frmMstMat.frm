VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmMstMat 
   Caption         =   "자재코드 등록"
   ClientHeight    =   11220
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16800
   LinkTopic       =   "Form1"
   ScaleHeight     =   11220
   ScaleWidth      =   16800
   StartUpPosition =   3  'Windows 기본값
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   " 자재코드 정보 입력 "
      Height          =   1275
      Left            =   90
      TabIndex        =   2
      Top             =   8010
      Width           =   15225
      Begin VB.TextBox txtMatCd 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   450
         MaxLength       =   6
         TabIndex        =   11
         Text            =   "123456"
         Top             =   690
         Width           =   1245
      End
      Begin VB.TextBox txtMatNm 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1710
         MaxLength       =   30
         TabIndex        =   10
         Text            =   "마이클조단"
         Top             =   690
         Width           =   3765
      End
      Begin VB.TextBox txtUserPW 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5490
         MaxLength       =   20
         TabIndex        =   9
         Text            =   "0001"
         Top             =   690
         Width           =   1245
      End
      Begin VB.TextBox txtUserRegID 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   8010
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   8
         Text            =   "123456"
         Top             =   690
         Width           =   1245
      End
      Begin VB.CheckBox chkUsedYN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "사용"
         Height          =   255
         Left            =   6990
         TabIndex        =   7
         Top             =   720
         Width           =   795
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00E0E0E0&
         Caption         =   "닫기"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   12390
         Style           =   1  '그래픽
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFFF&
         Caption         =   "저장"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   10110
         Style           =   1  '그래픽
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdClear 
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFC0&
         Caption         =   "지움"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   11250
         Style           =   1  '그래픽
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00E0E0E0&
         Caption         =   "삭제"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   13530
         Style           =   1  '그래픽
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00808080&
         Height          =   285
         Index           =   0
         Left            =   450
         Top             =   390
         Width           =   1245
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00808080&
         Height          =   285
         Index           =   1
         Left            =   1710
         Top             =   390
         Width           =   3765
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00808080&
         Height          =   285
         Index           =   2
         Left            =   5490
         Top             =   390
         Width           =   1245
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00808080&
         Height          =   285
         Index           =   5
         Left            =   6750
         Top             =   390
         Width           =   1245
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FF0000&
         BorderColor     =   &H00808080&
         Height          =   285
         Index           =   6
         Left            =   8010
         Top             =   390
         Width           =   1245
      End
      Begin VB.Label lblUser 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFC0&
         Caption         =   "자재코드"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   450
         TabIndex        =   16
         Top             =   420
         Width           =   1245
      End
      Begin VB.Label lblUser 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFC0&
         Caption         =   "자재명"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   1
         Left            =   1710
         TabIndex        =   15
         Top             =   420
         Width           =   3765
      End
      Begin VB.Label lblUser 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFC0&
         Caption         =   "순서"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   2
         Left            =   5490
         TabIndex        =   14
         Top             =   420
         Width           =   1245
      End
      Begin VB.Label lblUser 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFC0&
         Caption         =   "사용여부"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   5
         Left            =   6750
         TabIndex        =   13
         Top             =   420
         Width           =   1245
      End
      Begin VB.Label lblUser 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFC0&
         Caption         =   "입력자"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   6
         Left            =   8010
         TabIndex        =   12
         Top             =   420
         Width           =   1245
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   " 자재코드 리스트 "
      Height          =   7905
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   15225
      Begin FPSpread.vaSpread spdMat 
         Height          =   7545
         Left            =   90
         TabIndex        =   1
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
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         ScrollBarExtMode=   -1  'True
         SelectBlockOptions=   0
         ShadowColor     =   15400934
         SpreadDesigner  =   "frmMstMat.frx":0000
         ScrollBarTrack  =   3
         ShowScrollTips  =   3
      End
   End
End
Attribute VB_Name = "frmMstMat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-----------------------------------------------------------------------------'
'   파일명  : frmMstMat.frm
'   작성자  : 오세원
'   내  용  : 자제코드등록
'   작성일  : 2020-02-10
'   버  전  : 1.0.0
'   고  객  : 국도화학
'-----------------------------------------------------------------------------'

Private Sub cmdClear_Click()
    
    txtMatCd.Text = ""
    txtMatNm.Text = ""
    chkUsedYN.Value = "1"
    txtUserRegID.Text = gKUKDO.USERID

End Sub

Private Sub cmdClose_Click()

    Unload Me

End Sub


'-- 관리자용
Private Sub cmdDelete_Click()
    
    gMAT.CD = txtUserID.Text
    gMAT.NAME = txtUserNm.Text
    
    If chkUsedYN.Value = "1" Then
        gMAT.YN = "Y"
    Else
        gMAT.YN = "N"
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
    gMAT.ID = txtUserID.Text
    gMAT.NAME = txtUserNm.Text
    gMAT.PW = txtUserPW.Text
    gMAT.DEPT = txtUserDepart.Text
    If cboUserComp.Text = "사용자" Then
        gMAT.COMP = "2"
    Else
        gMAT.COMP = "1"
    End If
    If chkUsedYN.Value = "1" Then
        gMAT.YN = "Y"
    Else
        gMAT.YN = "N"
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

