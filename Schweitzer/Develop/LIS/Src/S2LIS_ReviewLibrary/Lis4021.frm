VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frm4021CumCdSet 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "누적코드 등록"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   12930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      Height          =   510
      Left            =   10095
      Style           =   1  '그래픽
      TabIndex        =   12
      Tag             =   "124"
      Top             =   105
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   11430
      Style           =   1  '그래픽
      TabIndex        =   13
      Tag             =   "128"
      Top             =   105
      Width           =   1320
   End
   Begin VB.CommandButton cmdNew 
      BackColor       =   &H00F4F0F2&
      Caption         =   "&New"
      Height          =   405
      Left            =   150
      Style           =   1  '그래픽
      TabIndex        =   11
      Top             =   6345
      Width           =   1305
   End
   Begin VB.CommandButton cmdRefresh 
      BackColor       =   &H00F4F0F2&
      Caption         =   "Re&fresh"
      Height          =   405
      Left            =   1800
      Style           =   1  '그래픽
      TabIndex        =   10
      Top             =   6360
      Width           =   1305
   End
   Begin VB.TextBox txtPassWd 
      BackColor       =   &H00F1F5F4&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  '사용 못함
      Left            =   7275
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   195
      Width           =   1245
   End
   Begin VB.ListBox lstCumList 
      BackColor       =   &H00FFF9F4&
      BeginProperty Font 
         Name            =   "돋움체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6105
      Left            =   135
      TabIndex        =   14
      Top             =   195
      Width           =   2970
   End
   Begin VB.TextBox txtCumCd 
      BackColor       =   &H00F1F5F4&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4140
      TabIndex        =   0
      Top             =   195
      Width           =   2100
   End
   Begin VB.Frame fraDetail 
      BackColor       =   &H00DBE6E6&
      Height          =   6240
      Left            =   3150
      TabIndex        =   15
      Top             =   540
      Width           =   9615
      Begin VB.ListBox lstTestListIndex 
         Height          =   3660
         Left            =   30
         TabIndex        =   16
         Top             =   2565
         Visible         =   0   'False
         Width           =   2580
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00F4F0F2&
         Caption         =   "삭제(&D)"
         Height          =   510
         Left            =   8265
         Style           =   1  '그래픽
         TabIndex        =   26
         Tag             =   "132"
         Top             =   180
         Width           =   1320
      End
      Begin VB.CommandButton cmdDeptList 
         BackColor       =   &H00D1DCD7&
         Caption         =   "▼"
         Height          =   345
         Left            =   7725
         MousePointer    =   14  '화살표와 물음표
         Style           =   1  '그래픽
         TabIndex        =   23
         Top             =   750
         Width           =   285
      End
      Begin MedControls1.LisLabel lblDeptNm 
         Height          =   345
         Left            =   8025
         TabIndex        =   22
         Top             =   750
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   609
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
      End
      Begin VB.OptionButton optDeptFg 
         BackColor       =   &H00DBE6E6&
         Caption         =   "과별"
         Height          =   285
         Index           =   1
         Left            =   6225
         TabIndex        =   21
         Top             =   795
         Width           =   675
      End
      Begin VB.OptionButton optDeptFg 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Default"
         Height          =   285
         Index           =   0
         Left            =   5055
         TabIndex        =   20
         Top             =   795
         Width           =   945
      End
      Begin VB.CommandButton cmdSpcList 
         BackColor       =   &H00DEDBDD&
         Caption         =   "▼"
         Height          =   345
         Left            =   1935
         MousePointer    =   14  '화살표와 물음표
         Style           =   1  '그래픽
         TabIndex        =   18
         Top             =   735
         Width           =   285
      End
      Begin MedControls1.LisLabel lblSpcNm 
         Height          =   360
         Left            =   2235
         TabIndex        =   17
         Top             =   735
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   635
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
      End
      Begin VB.TextBox txtSpcCd 
         BackColor       =   &H00F1F5F4&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1020
         TabIndex        =   3
         Top             =   735
         Width           =   915
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00F4F0F2&
         Caption         =   "저장(&S)"
         Height          =   510
         Left            =   6945
         Style           =   1  '그래픽
         TabIndex        =   9
         Tag             =   "132"
         Top             =   180
         Width           =   1320
      End
      Begin VB.ListBox lstTestList 
         BackColor       =   &H00F7FFFF&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4260
         Left            =   225
         Style           =   1  '확인란
         TabIndex        =   6
         Top             =   1635
         Width           =   4245
      End
      Begin VB.CommandButton cmdMove 
         BackColor       =   &H00CDE7FA&
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   0
         Left            =   4560
         Style           =   1  '그래픽
         TabIndex        =   7
         Top             =   2895
         Width           =   435
      End
      Begin VB.CommandButton cmdMove 
         BackColor       =   &H00CDE7FA&
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   1
         Left            =   4560
         Style           =   1  '그래픽
         TabIndex        =   8
         Top             =   3615
         Width           =   435
      End
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00F1F5F4&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1020
         TabIndex        =   4
         Top             =   1215
         Width           =   2700
      End
      Begin VB.CommandButton cmdReset 
         BackColor       =   &H00F4F0F2&
         Caption         =   "&Reset"
         Height          =   360
         Left            =   3735
         Style           =   1  '그래픽
         TabIndex        =   5
         Top             =   1215
         Width           =   735
      End
      Begin VB.TextBox txtCodeNm 
         BackColor       =   &H00F1F5F4&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1020
         TabIndex        =   2
         Top             =   330
         Width           =   5820
      End
      Begin VB.ListBox lstDeptList 
         Appearance      =   0  '평면
         BackColor       =   &H00F7FFF7&
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   6900
         TabIndex        =   25
         Top             =   1125
         Visible         =   0   'False
         Width           =   2595
      End
      Begin FPSpread.vaSpread tblSelList 
         Height          =   4815
         Left            =   5070
         TabIndex        =   24
         Top             =   1185
         Width           =   4470
         _Version        =   196608
         _ExtentX        =   7885
         _ExtentY        =   8493
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayColHeaders=   0   'False
         DisplayRowHeaders=   0   'False
         EditModePermanent=   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GridColor       =   14737632
         GridShowHoriz   =   0   'False
         GridShowVert    =   0   'False
         MaxCols         =   4
         ScrollBars      =   2
         SelectBlockOptions=   0
         SpreadDesigner  =   "Lis4021.frx":0000
         Appearance      =   1
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   345
         Index           =   2
         Left            =   45
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   330
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   609
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "코 드 명"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   345
         Index           =   3
         Left            =   45
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   735
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   609
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "검체코드"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   345
         Index           =   4
         Left            =   45
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   1215
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   609
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "검색"
         Appearance      =   0
      End
      Begin VB.Label lblDeptCd 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00D1D8D3&
         BorderStyle     =   1  '단일 고정
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6915
         TabIndex        =   27
         Top             =   750
         Width           =   780
      End
   End
   Begin VB.ListBox lstSpcList 
      BackColor       =   &H00FCF8FB&
      Height          =   3840
      Left            =   3960
      TabIndex        =   19
      Top             =   1650
      Visible         =   0   'False
      Width           =   4170
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   345
      Index           =   1
      Left            =   3150
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   195
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   609
      BackColor       =   10392451
      ForeColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "누적코드"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   330
      Index           =   0
      Left            =   6300
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   210
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   582
      BackColor       =   10392451
      ForeColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "비밀번호"
      Appearance      =   0
   End
End
Attribute VB_Name = "frm4021CumCdSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private MySql As New clsLISSqlStatement

Private strPassWd As String
Private blnChanged As Boolean
Private blnNewFg As Boolean
Private MsgFg As Boolean

Private mIsManager As Boolean
Private mDeptCd As String

Public Property Get IsManager() As Boolean
    IsManager = mIsManager
End Property

Public Property Let IsManager(ByVal vNewValue As Boolean)
    mIsManager = vNewValue
End Property

Public Property Get DeptCd() As String
    DeptCd = mDeptCd
End Property

Public Property Let DeptCd(ByVal vNewValue As String)
    mDeptCd = vNewValue
End Property


Private Sub cmdClear_Click()
    Call ClearRtn
    Call LockCumCd(True)
    Call LockPWD(True)
    fraDetail.Enabled = False
    cmdSave.Enabled = False
    cmdDelete.Enabled = False
    txtCumCd.SetFocus
End Sub

Private Sub cmdDelete_Click()

    Dim SqlStmt As String
    Dim Resp As VbMsgBoxResult
    
    Resp = MsgBox("해당 누적코드를 정말 삭제하시겠습니까?", vbQuestion + vbOKCancel, "메세지")
    If Resp = vbCancel Then Exit Sub
    
    On Error GoTo Err_Trap
    SqlStmt = MySql.SqlDeleteLAB031(LC2_CumItem, Trim(txtCumCd.Text))
    DBConn.BeginTrans
    DBConn.Execute SqlStmt
    DBConn.CommitTrans
    Call cmdClear_Click
    Call cmdRefresh_Click
    Exit Sub
    
Err_Trap:
    MsgBox Err.Description, vbExclamation

End Sub

Private Sub cmdDeptList_Click()
   
   With lstDeptList
      .Visible = True
      .ZOrder 0
   End With
    
End Sub

Private Sub cmdExit_Click()
    Unload Me
    Set MySql = Nothing
'    Set frm4021CumCdSet = Nothing
End Sub

Private Sub cmdNew_Click()
    Call ClearRtn
    Call LockCumCd(False)
    Call LockPWD(False)
    fraDetail.Enabled = True
    cmdSave.Enabled = True
    cmdDelete.Enabled = True
    blnNewFg = True
    blnChanged = True
    txtCumCd.SetFocus
End Sub

Private Sub cmdRefresh_Click()
    Call LoadCumList(lstCumList)
End Sub

Private Sub cmdSave_Click()
    Dim SqlStmt As String
    Dim strTmp As String
    Dim strSpc As String
    Dim I As Long
    
    If Trim(txtCumCd.Text) = "" Then
        txtCumCd.SetFocus
        Exit Sub
    End If
    If Trim(txtPassWd.Text) = "" Then
        txtPassWd.SetFocus
        Exit Sub
    End If
    If Trim(txtCodeNm.Text) = "" Then
        txtCodeNm.SetFocus
        Exit Sub
    End If
    If tblSelList.MaxRows = 0 Then
        lstTestList.SetFocus
        Exit Sub
    End If
    
    SqlStmt = MySql.SqlDeleteLAB031(LC2_CumItem, Trim(txtCumCd.Text))
    
    On Error GoTo Err_Trap
    
    DBConn.BeginTrans
    DBConn.Execute SqlStmt
    With tblSelList
        For I = 1 To tblSelList.MaxRows
            .Row = I
            .Col = 1: strSpc = .Value
            .Col = 2: strTmp = medGetP(.Value, 1, " ")
            SqlStmt = MySql.SqlsaveLAB031(LC2_CumItem, Trim(txtCumCd.Text), strTmp, txtCodeNm.Text, _
                                        lblDeptCd.Caption, txtPassWd.Text, Trim(CStr(I)), _
                                        strSpc, gEmpId, "", 1)
            DBConn.Execute SqlStmt
        Next
    End With
    DBConn.CommitTrans
    Call cmdClear_Click
    Call cmdRefresh_Click
    Exit Sub
Err_Trap:
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub cmdSpcList_Click()
    lstSpcList.Visible = True
    lstSpcList.ZOrder 0
End Sub

Private Sub Form_Load()
    
    'Me.Show
    DoEvents
    
    MouseRunning
    Call LoadCumList(lstCumList)
    Call LoadSpcList(lstSpcList)
    Call LoadDeptList
    MouseDefault
    
    Call medHorScrol(lstTestList)
    Call ClearRtn
    Call LockCumCd(True)
    Call LockPWD(True)
    fraDetail.Enabled = False
    cmdSave.Enabled = False
    cmdDelete.Enabled = False
    
End Sub

Private Sub ClearRtn()
    txtCumCd.Text = ""
    txtPassWd.Text = ""
    txtCodeNm.Text = ""
    txtSpcCd.Text = ""
    txtSearch.Text = ""
    lblSpcNm.Caption = ""
    'lstSelList.Clear
    tblSelList.MaxRows = 0
    lstTestList.Clear
    blnNewFg = False
    blnChanged = False
    fraDetail.Enabled = False
    cmdSave.Enabled = False
    cmdDelete.Enabled = False
    MsgFg = False
    optDeptFg(0).Enabled = False    'mIsManager
    optDeptFg(0).Value = False  'mIsManager
    optDeptFg(1).Value = True   'Not mIsManager
'    cmdDeptList.Enabled = mIsManager
    'If mDeptCd <> "" Then
    If gDeptCd <> "" Then
        Dim I As Integer
        I = medListFind(lstDeptList, gDeptCd)   'mDeptCd)
        cmdDeptList.Enabled = False
        lblDeptCd.Caption = gDeptCd 'mDeptCd
        lblDeptNm.Caption = medGetP(lstDeptList.List(I), 2, vbTab)
    Else
        lblDeptCd.Caption = "0"
        lblDeptNm.Caption = ""
    End If
End Sub


Private Sub LockCumCd(ByVal blnLock As Boolean)

    txtCumCd.Locked = blnLock
    If blnLock Then
        txtCumCd.BackColor = DCM_LightGray
    Else
        txtCumCd.Text = ""
        txtCumCd.BackColor = vbWhite
    End If
    
End Sub

Private Sub LockPWD(ByVal blnLock As Boolean)

    txtPassWd.Locked = blnLock
    txtPassWd.Enabled = Not blnLock
    If blnLock Then
        txtPassWd.BackColor = DCM_LightGray
    Else
        txtPassWd.Text = ""
        txtPassWd.BackColor = vbWhite
    End If
    
End Sub


Private Sub lstCumList_Click()
    
    Call ClearRtn
    txtCumCd.Text = medGetP(lstCumList.Text, 1, vbTab)
    txtCodeNm.Text = medGetP(lstCumList.Text, 2, vbTab)
    Call LockCumCd(True)
    Call LockPWD(False)
    Call DisplayItem(txtCumCd.Text)
    fraDetail.Enabled = False
    'If (Not IsManager) And (lblDeptCd.Caption <> mDeptCd) Then
    If (lblDeptCd.Caption <> gDeptCd) Then
        txtPassWd.Enabled = False
        txtPassWd.BackColor = DCM_LightGray
    Else
        txtPassWd.Enabled = True
        txtPassWd.BackColor = vbWhite
        txtPassWd.SetFocus
    End If
    
End Sub

Private Sub lstDeptList_Click()
    lblDeptCd.Caption = medGetP(lstDeptList.Text, 1, vbTab)
    lblDeptNm.Caption = medGetP(lstDeptList.Text, 2, vbTab)
    lstDeptList.Visible = False
End Sub

Private Sub lstSpcList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call lstSpcList_MouseDown(1, 0, 0, 0)
        txtSearch.SetFocus
    End If
End Sub

Private Sub lstSpcList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
        lstSpcList.Visible = False
        txtSpcCd.Text = medGetP(lstSpcList.Text, 1, vbTab)
        lblSpcNm.Caption = medGetP(lstSpcList.Text, 2, vbTab)
        DoEvents
        Call LoadSpcItem(lstTestList, lstTestListIndex, txtSpcCd.Text)
    End If
    
End Sub

Private Sub optDeptFg_Click(Index As Integer)
    If Index = 0 Then
        cmdDeptList.Enabled = False
        lblDeptCd.Caption = "0"
        lblDeptNm.Caption = ""
    Else
        cmdDeptList.Enabled = True
        lblDeptCd.Caption = ""
        lblDeptNm.Caption = ""
    End If
    lstDeptList.Visible = False
End Sub

Private Sub tblSelList_Change(ByVal Col As Long, ByVal Row As Long)
    Dim I As Integer
    Dim iSeq As Integer
    If MsgFg Then Exit Sub
    If Col = 3 Then
        With tblSelList
            .Row = Row: .Col = 3
            iSeq = .Value
            
            MsgFg = True
            If iSeq < Row Then
                For I = iSeq To Row - 1
                    .Row = I: .Col = 3
                    .Value = .Value + 1
                Next
            ElseIf iSeq > Row Then
                For I = Row + 1 To .MaxRows
                    .Row = I: .Col = 3
                    .Value = .Value - 1
                Next
            End If
            MsgFg = False
            
            .SortBy = SortByRow
            .SortKey(1) = 3
            .SortKeyOrder(1) = SortKeyOrderAscending
            .Row = 1: .Row2 = .MaxRows
            .Col = 1: .Col2 = .MaxCols
            .BlockMode = True
            .Action = ActionSort
            .BlockMode = False
        End With
    End If
End Sub

Private Sub tblSelList_Click(ByVal Col As Long, ByVal Row As Long)
    Dim iFlag As Integer
    If Col = 1 Or Col = 2 Then
        With tblSelList
            .Row = Row: .Col = 4
            If .Value = "1" Then
                .Value = "0"
            Else
                .Value = "1"
            End If
            iFlag = Val(.Value)
            .Col = -1
            .Row = Row: .Row2 = Row
            .BlockMode = True
            .BackColor = Choose(iFlag + 1, &HF5F8F8, &H800000)
            .ForeColor = Choose(iFlag + 1, vbBlack, vbWhite)
            .BlockMode = False
        End With
    End If
End Sub

Private Sub txtCodeNm_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtCumCd_GotFocus()
    With txtCumCd
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub txtCumCd_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtPassWd_GotFocus()
    With txtPassWd
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtPassWd_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If fraDetail.Enabled Then
            txtCodeNm.SetFocus
        Else
            txtCumCd.SetFocus
        End If
    End If
End Sub

Private Sub txtPassWd_LostFocus()
    
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    If ActiveControl.Name = cmdExit.Name Then Exit Sub
    If ActiveControl.Name = cmdClear.Name Then Exit Sub
    If ActiveControl.Name = cmdNew.Name Then Exit Sub
    If ActiveControl.Name = cmdRefresh.Name Then Exit Sub
    If ActiveControl.Name = lstCumList.Name Then Exit Sub
    
    If Not blnNewFg And txtCumCd.Text <> "" Then
        If strPassWd <> txtPassWd.Text Then
            MsgBox "비밀번호가 일치하지 않습니다. 다시 입력하세요.", vbExclamation, "메세지"
            txtPassWd.SetFocus
            Call txtPassWd_GotFocus
            Exit Sub
        Else
            Call LockCumCd(True)
            fraDetail.Enabled = True
            cmdSave.Enabled = True
            cmdDelete.Enabled = True
            txtCodeNm.SetFocus
        End If
    End If
End Sub

Private Sub txtSearch_Change()
    
    Dim I As Integer
    
    If txtSearch.Text = "" Then Exit Sub
    
    I = medListFind(lstTestList, txtSearch.Text)
    If I < 0 Then I = medListFind(lstTestListIndex, txtSearch.Text)
    lstTestList.ListIndex = I
    
End Sub

Private Sub txtSearch_GotFocus()
    With txtSearch
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub cmdMove_Click(Index As Integer)

    Dim I As Integer
    Dim J As Integer
    
    Select Case Index
    Case 0:
        With tblSelList
            For I = 0 To lstTestList.ListCount - 1
                If lstTestList.Selected(I) Then
                    For J = 1 To .MaxRows
                        .Row = J
                        .Col = 1
                        If .Value = txtSpcCd.Text Then
                            .Col = 2
                            If .Value = lstTestList.List(I) Then GoTo Skip
                        End If
                    Next
                    .MaxRows = .MaxRows + 1
                    .Row = .MaxRows
                    .Col = 1: .Value = txtSpcCd.Text
                    .Col = 2: .Value = lstTestList.List(I)
                    .Col = 3: .Value = .Row
                    'lstSelList.AddItem lstTestList.List(i)
                End If
Skip:
            Next
        End With
    Case 1:
        With tblSelList
            For I = .MaxRows To 1 Step -1
                .Row = I
                .Col = 4
                If .Value = "1" Then
                    .Action = ActionDeleteRow
                    .MaxRows = .MaxRows - 1
                End If
            Next
            For I = 1 To .MaxRows
                .Row = I
                .Col = 3
                .Value = I
            Next
        End With
        'For i = lstSelList.ListCount - 1 To 0 Step -1
            'If lstSelList.Selected(i) Then _
                Call lstSelList.RemoveItem(i)
        'Next
    End Select
    blnChanged = True
    
End Sub

Private Sub cmdReset_Click()
    
    Dim I As Integer
    
    blnChanged = True
    For I = 0 To lstTestList.ListCount - 1
        lstTestList.Selected(I) = False
    Next
    txtSearch.Text = ""
    If fraDetail.Enabled Then txtSearch.SetFocus
    
End Sub

Private Sub DisplayItem(ByVal pCumCd As String)
    Dim SqlStmt As String
    Dim RS As Recordset
    Dim I As Integer
    Dim objRstSQL As New clsLISSqlReview
    Dim tmpStr As String
    
    SqlStmt = objRstSQL.SqlGetCumItem(pCumCd)
    Set RS = New Recordset
    RS.Open SqlStmt, DBConn
    
    'lstSelList.Clear
    tblSelList.MaxRows = 0
    If Not RS.EOF Then
        strPassWd = Trim("" & RS.Fields("PWD").Value)
        I = medListFind(lstDeptList, Trim(RS.Fields("DeptCd").Value))
        If I >= 0 Then
            optDeptFg(1).Value = True
            lblDeptCd.Caption = Trim("" & RS.Fields("DeptCd").Value)
            lblDeptNm.Caption = medGetP(lstDeptList.List(I), 2, vbTab)
        Else
            optDeptFg(0).Value = True
            lblDeptCd.Caption = Trim("" & RS.Fields("DeptCd").Value)
            lblDeptNm.Caption = ""
        End If
        txtSpcCd.Text = Trim("" & RS.Fields("Field5").Value)
        lstSpcList.ListIndex = medListFind(lstSpcList, txtSpcCd.Text)
        Call lstSpcList_MouseDown(1, 0, 0, 0)
    End If
    While (Not RS.EOF)
        If tblSelList.MaxRows < Val("" & RS.Fields("RptSeq").Value) Then _
            tblSelList.MaxRows = Val("" & RS.Fields("RptSeq").Value)
        tmpStr = Trim("" & RS.Fields("TestCd").Value) & Space(9)
        tblSelList.Row = Val("" & RS.Fields("RptSeq").Value)
        tblSelList.Col = 1
        tblSelList.Value = Trim("" & RS.Fields("Field5").Value)
        tblSelList.Col = 2
        tblSelList.Value = Mid(tmpStr, 1, 10) & Trim("" & RS.Fields("TestNm").Value)
        tblSelList.Col = 3
        tblSelList.Value = Val("" & RS.Fields("RptSeq").Value)
        'lstSelList.AddItem Mid(tmpStr, 1, 10) & _
                        Trim(rs.Fields("TestNm").Value), Val(rs.Fields("RptSeq").Value) - 1
        RS.MoveNext
    Wend
    Set RS = Nothing
    Set objRstSQL = Nothing
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyDown And lstTestList.ListCount > 0 Then
        lstTestList.Visible = True
        lstTestList.ListIndex = 0
        lstTestList.ZOrder 0
        lstTestList.SetFocus
    End If

End Sub

Private Sub txtSpcCd_Change()
    lstSpcList.ListIndex = medListFind(lstSpcList, txtSpcCd.Text)
End Sub

Private Sub txtSpcCd_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyDown And lstSpcList.ListCount > 0 Then
        lstSpcList.Visible = True
        lstSpcList.ZOrder 0
        lstSpcList.SetFocus
    End If

End Sub

Private Sub txtSpcCd_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call lstSpcList_MouseDown(1, 0, 0, 0)
        If mIsManager Then
            optDeptFg(0).SetFocus
        Else
            lstTestList.SetFocus
        End If
        Exit Sub
    End If
    If lstSpcList.ListCount > 0 Then
        lstSpcList.Visible = True
        lstSpcList.ZOrder 0
        Call medCodeHelp(KeyAscii, lstSpcList, txtSpcCd.Text, txtSpcCd, optDeptFg(0))
    End If
End Sub

Private Sub LoadDeptList()

'   Dim tmpSQL As String
'
'   With lstDeptList
'      'Set .MyOraSE = OraSe
'      .Clear
'      'tmpSql = MySql.SqlHIS003CodeList
'      'Set rs = OpenRecordSet(tmpSql)
'      ObjLISComCode.DeptCd.MoveFirst
'      While (Not ObjLISComCode.DeptCd.EOF)
'        .AddItem "" & ObjLISComCode.DeptCd.Fields("DeptCd") & vbTab & ObjLISComCode.DeptCd.Fields("DeptNm")
'        ObjLISComCode.DeptCd.MoveNext
'      Wend
'      'rs.Close
'      'Set rs = Nothing
'   End With
'    Dim objDept As clsBasisData
    Dim RS As Recordset
    
'    Set objDept = New clsBasisData
    Set RS = New Recordset
    
    RS.Open GetSQLDeptList, DBConn
    
    Do Until RS.EOF
        lstDeptList.AddItem RS.Fields("deptcd").Value & "" & vbTab & _
                            RS.Fields("deptnm").Value & ""
        RS.MoveNext
    Loop
    
    Set RS = Nothing
'    Set objDept = Nothing
    
End Sub

Public Sub LoadCumList(ByRef lstList As ListBox, Optional ByVal pDeptCd As String = "ALL")
    Dim SqlStmt As String
    Dim RS As Recordset
    
    SqlStmt = MySql.SqlGetCumList(pDeptCd)
    Set RS = New Recordset
    RS.Open SqlStmt, DBConn
    
    lstList.Clear
    While (Not RS.EOF)
        lstList.AddItem "" & Trim(RS.Fields("CumCd").Value) & vbTab & _
                        "" & Trim(RS.Fields("CumNm").Value)
        RS.MoveNext
    Wend
    Set RS = Nothing
    Set MySql = Nothing
    DoEvents
End Sub

Public Sub LoadSpcList(ByRef lstList As ListBox)

'    lstList.Clear
'
'    ObjLISComCode.LisSpc.MoveFirst
'    While (Not ObjLISComCode.LisSpc.EOF)
'        lstList.AddItem ObjLISComCode.LisSpc.Fields("spccd") & vbTab & _
'                        ObjLISComCode.LisSpc.Fields("spcnm")
'        ObjLISComCode.LisSpc.MoveNext
'    Wend
    Dim RS As Recordset
    Dim strSQL As String
    
    strSQL = "SELECT a.cdval1 spccd, a.field4 spcnm, a.field3 spcabbr, a.field5 spcbarnm, " & _
             "       a.field1 multifg, a.field2 spcgrp, b.field2 labrange  " & _
             "FROM " & T_LAB032 & " b, " & T_LAB032 & " a " & _
             "WHERE  a.cdindex = 'C215' " & _
             "AND    " & DBJ("b.cdindex = 'C217'") & _
             "AND    " & DBJ("b.cdval1  =* a.field2")
    
    Set RS = New Recordset
    RS.Open strSQL, DBConn
    
    lstList.Clear
    Do Until RS.EOF
        lstList.AddItem RS.Fields("spccd").Value & "" & vbTab & _
                        RS.Fields("spcnm").Value & ""
        RS.MoveNext
    Loop
    
    Set RS = Nothing
End Sub

Public Sub LoadSpcItem(ByRef lstList As ListBox, ByRef lstList1 As ListBox, ByVal pSpcCd As String)

    Dim SqlStmt As String
    Dim RS As Recordset
    Dim tmpStr As String
    Dim I%
    
    '상세항목 제외...
    SqlStmt = MySql.SqlLoadSpcItem(pSpcCd)
    Set RS = New Recordset
    RS.Open SqlStmt, DBConn
    
    lstList.Clear
    lstList1.Clear
    If RS.EOF Then GoTo NoData
    
    For I = 1 To RS.RecordCount
        tmpStr = "" & RS.Fields("TestCd").Value & Space(9)
        lstList.AddItem Mid(tmpStr, 1, 10) & _
                        "" & RS.Fields("TestNm").Value
        lstList1.AddItem "" & RS.Fields("TestNm").Value & vbTab & "" & RS.Fields("TestCd").Value
        RS.MoveNext
    Next I
    
NoData:
    Set RS = Nothing
    
End Sub
