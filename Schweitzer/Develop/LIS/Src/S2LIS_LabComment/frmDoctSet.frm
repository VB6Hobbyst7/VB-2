VERSION 5.00
Object = "{9167B9A7-D5FA-11D2-86CA-00104BD5476F}#5.0#0"; "DRCTL1.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frmDoctSet 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "의사별 설정화면"
   ClientHeight    =   9045
   ClientLeft      =   1710
   ClientTop       =   1005
   ClientWidth     =   9465
   Icon            =   "frmDoctSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9045
   ScaleWidth      =   9465
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00EEEEEE&
      Height          =   8235
      Left            =   150
      TabIndex        =   23
      Top             =   705
      Width           =   9135
      Begin VB.TextBox txtCertNo 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BorderStyle     =   0  '없음
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4875
         TabIndex        =   1
         Top             =   1215
         Width           =   2565
      End
      Begin VB.CommandButton cmdNewTemp 
         Height          =   285
         Index           =   1
         Left            =   4230
         Picture         =   "frmDoctSet.frx":0442
         Style           =   1  '그래픽
         TabIndex        =   36
         Tag             =   "R"
         Top             =   6150
         Width           =   450
      End
      Begin VB.CommandButton cmdNewTemp 
         Height          =   285
         Index           =   0
         Left            =   4230
         Picture         =   "frmDoctSet.frx":0949
         Style           =   1  '그래픽
         TabIndex        =   35
         Tag             =   "C"
         Top             =   4155
         Width           =   450
      End
      Begin MedControls1.LisLabel lblTemp 
         Height          =   255
         Index           =   0
         Left            =   675
         TabIndex        =   32
         Top             =   4170
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   450
         BackColor       =   14212826
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         AutoSize        =   -1  'True
         Caption         =   "6. 검증/판독소견(Comments) Template"
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00D8DEDA&
         Caption         =   "종료(&X)"
         Height          =   405
         Left            =   7530
         Style           =   1  '그래픽
         TabIndex        =   16
         Top             =   210
         Width           =   1050
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00D8DEDA&
         Caption         =   "저장(&S)"
         Height          =   405
         Left            =   6420
         Style           =   1  '그래픽
         TabIndex        =   15
         Top             =   210
         Width           =   1050
      End
      Begin VB.TextBox txtRcmd 
         Height          =   1470
         Left            =   735
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   13
         Top             =   6465
         Width           =   7605
      End
      Begin VB.ComboBox cboRcmd 
         Height          =   300
         Left            =   4695
         Style           =   2  '드롭다운 목록
         TabIndex        =   14
         Top             =   6150
         Width           =   3660
      End
      Begin VB.TextBox txtCmt 
         Height          =   1395
         Left            =   720
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   11
         Top             =   4470
         Width           =   7605
      End
      Begin VB.ComboBox cboCmt 
         Height          =   300
         Left            =   4695
         Style           =   2  '드롭다운 목록
         TabIndex        =   12
         Top             =   4155
         Width           =   3630
      End
      Begin VB.TextBox txtOthers 
         Appearance      =   0  '평면
         BorderStyle     =   0  '없음
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5415
         TabIndex        =   10
         Top             =   3570
         Width           =   2565
      End
      Begin VB.CheckBox chkMethod 
         BackColor       =   &H00D8DEDA&
         Caption         =   "Others;"
         ForeColor       =   &H00DD6131&
         Height          =   255
         Index           =   5
         Left            =   4440
         TabIndex        =   9
         Top             =   3585
         Width           =   975
      End
      Begin VB.CheckBox chkMethod 
         BackColor       =   &H00D8DEDA&
         Caption         =   "Repeat / Recheck"
         ForeColor       =   &H00DD6131&
         Height          =   255
         Index           =   4
         Left            =   900
         TabIndex        =   8
         Top             =   3600
         Width           =   2685
      End
      Begin VB.CheckBox chkMethod 
         BackColor       =   &H00D8DEDA&
         Caption         =   "Panic/Alert Value Verification"
         ForeColor       =   &H00DD6131&
         Height          =   255
         Index           =   3
         Left            =   4440
         TabIndex        =   7
         Top             =   3270
         Width           =   3645
      End
      Begin VB.CheckBox chkMethod 
         BackColor       =   &H00D8DEDA&
         Caption         =   "Delta Check Verification"
         ForeColor       =   &H00DD6131&
         Height          =   255
         Index           =   2
         Left            =   900
         TabIndex        =   6
         Top             =   3285
         Width           =   2685
      End
      Begin VB.CheckBox chkMethod 
         BackColor       =   &H00D8DEDA&
         Caption         =   "Internal Quality Control"
         ForeColor       =   &H00DD6131&
         Height          =   255
         Index           =   1
         Left            =   4440
         TabIndex        =   5
         Top             =   2955
         Width           =   2685
      End
      Begin VB.CheckBox chkMethod 
         BackColor       =   &H00D8DEDA&
         Caption         =   "Calibration Verification"
         ForeColor       =   &H00DD6131&
         Height          =   255
         Index           =   0
         Left            =   900
         TabIndex        =   4
         Top             =   2970
         Width           =   2685
      End
      Begin MedControls1.LisLabel LisLabel5 
         Height          =   255
         Left            =   645
         TabIndex        =   30
         Top             =   2655
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   450
         BackColor       =   14212826
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   "5. 검증방법 Setting"
      End
      Begin VB.TextBox txtPtCnt 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BorderStyle     =   0  '없음
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   3
         Top             =   2130
         Width           =   630
      End
      Begin VB.TextBox txtDoctNo 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BorderStyle     =   0  '없음
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2130
         TabIndex        =   0
         Top             =   765
         Width           =   2565
      End
      Begin VB.TextBox txtDayCnt 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BorderStyle     =   0  '없음
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1560
         TabIndex        =   2
         Top             =   1680
         Width           =   630
      End
      Begin MedControls1.LisLabel LisLabel1 
         Height          =   390
         Left            =   480
         TabIndex        =   26
         Top             =   1620
         Width           =   8085
         _ExtentX        =   14261
         _ExtentY        =   688
         BackColor       =   14212826
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   "3. 입원 후              일 경과된 환자를 대상으로 한다."
         LeftGab         =   200
      End
      Begin MedControls1.LisLabel LisLabel2 
         Height          =   390
         Left            =   480
         TabIndex        =   27
         Top             =   705
         Width           =   8085
         _ExtentX        =   14261
         _ExtentY        =   688
         BackColor       =   14212826
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   "1. 전문의 번호  :"
         LeftGab         =   200
      End
      Begin MedControls1.LisLabel LisLabel3 
         Height          =   390
         Left            =   495
         TabIndex        =   28
         Top             =   2070
         Width           =   8085
         _ExtentX        =   14261
         _ExtentY        =   688
         BackColor       =   14212826
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   "4. 하루에               명의 환자에 대해 보고서를 작성한다."
         LeftGab         =   200
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   1425
         Left            =   495
         TabIndex        =   29
         Top             =   2535
         Width           =   8085
         _ExtentX        =   14261
         _ExtentY        =   2514
         BackColor       =   14212826
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   ""
         LeftGab         =   200
      End
      Begin MedControls1.LisLabel lblTemp 
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   34
         Top             =   6180
         Width           =   3900
         _ExtentX        =   6879
         _ExtentY        =   450
         BackColor       =   14212826
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         AutoSize        =   -1  'True
         Caption         =   "7. 추천(Recommendation) Template"
      End
      Begin MedControls1.LisLabel LisLabel8 
         Height          =   2025
         Left            =   525
         TabIndex        =   33
         Top             =   6045
         Width           =   8085
         _ExtentX        =   14261
         _ExtentY        =   3572
         BackColor       =   14212826
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   ""
         LeftGab         =   200
      End
      Begin MedControls1.LisLabel LisLabel6 
         Height          =   1965
         Left            =   510
         TabIndex        =   31
         Top             =   4020
         Width           =   8085
         _ExtentX        =   14261
         _ExtentY        =   3466
         BackColor       =   14212826
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   ""
         LeftGab         =   200
      End
      Begin MedControls1.LisLabel LisLabel7 
         Height          =   390
         Left            =   480
         TabIndex        =   39
         Top             =   1155
         Width           =   8085
         _ExtentX        =   14261
         _ExtentY        =   688
         BackColor       =   14212826
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   "2. 대한임상병리 학회 검사실 신임제도 인증번호  :"
         LeftGab         =   200
      End
   End
   Begin DRcontrol1.DrFrame fraNewTemp 
      Height          =   4125
      Left            =   660
      TabIndex        =   17
      Top             =   4665
      Visible         =   0   'False
      Width           =   8100
      _ExtentX        =   14288
      _ExtentY        =   7276
      Title           =   "Title"
      TitlePos        =   1
      BackColor       =   14411494
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmdTempExit 
         BackColor       =   &H00C7D8D8&
         Caption         =   "취소"
         Height          =   285
         Left            =   7305
         Style           =   1  '그래픽
         TabIndex        =   22
         Top             =   60
         Width           =   675
      End
      Begin VB.CommandButton cmdTempSave 
         BackColor       =   &H00C7D8D8&
         Caption         =   "저장"
         Height          =   285
         Left            =   6555
         Style           =   1  '그래픽
         TabIndex        =   21
         Top             =   60
         Width           =   720
      End
      Begin VB.TextBox txtTempNm 
         Height          =   285
         Left            =   3825
         TabIndex        =   19
         Top             =   540
         Width           =   4185
      End
      Begin VB.TextBox txtTempCd 
         Height          =   285
         Left            =   720
         TabIndex        =   18
         Top             =   540
         Width           =   1890
      End
      Begin VB.TextBox txtTemplate 
         Height          =   3165
         Left            =   75
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   20
         Top             =   855
         Width           =   7935
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "코드명 : "
         Height          =   180
         Left            =   3090
         TabIndex        =   38
         Top             =   585
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "코드 :"
         Height          =   180
         Left            =   150
         TabIndex        =   37
         Top             =   570
         Width           =   480
      End
   End
   Begin VB.Label lblDoctNm 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1080
      TabIndex        =   25
      Top             =   345
      Width           =   1140
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Caption         =   "* 다음은                 선생님의 '종합검증/판독' 관련 설정내용 입니다."
      Height          =   180
      Left            =   420
      TabIndex        =   24
      Top             =   345
      Width           =   6075
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00DBF2FD&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Height          =   480
      Left            =   165
      Shape           =   4  '둥근 사각형
      Top             =   180
      Width           =   9105
   End
End
Attribute VB_Name = "frmDoctSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_SaveFg As Boolean

Private Sub cboCmt_Click()
    Dim strKey As String
    strKey = objDoctor.DoctId & "C" & Trim(Mid(cboCmt.Text, 1, 5))
    txtCmt.Text = objDoctor.txtCmt(strKey).Txtrst
    txtCmt.Enabled = True
End Sub

Private Sub cboRcmd_Click()
    Dim strKey As String
    strKey = objDoctor.DoctId & "R" & Trim(Mid(cboRcmd.Text, 1, 5))
    txtRcmd.Text = objDoctor.txtCmt(strKey).Txtrst
    txtRcmd.Enabled = True
End Sub

Private Sub chkMethod_Click(Index As Integer)
    If Index = 5 Then
        If chkMethod(Index).Value = 1 Then
            txtOthers.Enabled = True
        Else
            txtOthers.Text = ""
            txtOthers.Enabled = False
        End If
    End If
End Sub

Private Sub chkMethod_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"

End Sub

Private Sub cmdExit_Click()
    Unload Me
    Set frmDoctSet = Nothing
End Sub

Private Sub cmdNewTemp_Click(Index As Integer)
    fraNewTemp.Title = lblTemp(Index).Caption
    fraNewTemp.Tag = cmdNewTemp(Index).Tag
    txtTempCd.Text = ""
    txtTempNm.Text = ""
    txtTemplate.Text = ""
    fraNewTemp.Visible = True
    fraNewTemp.ZOrder 0
    txtTempCd.SetFocus
End Sub

Private Sub cmdSave_Click()
    Dim i As Integer
    
    If txtDoctNo.Text = "" Then
        MsgBox "전문의입력이 누락되었습니다.", vbInformation + vbOKOnly, Me.Caption
        Exit Sub
    End If
    
    With objDoctor
        .Doctno = txtDoctNo.Text
        .Certno = txtCertNo.Text
        .Daycnt = txtDayCnt.Text
        .Ptcnt = txtPtCnt.Text
        .Method = ""
        For i = 0 To chkMethod.Count - 1
            .Method = .Method & chkMethod(i).Value
        Next
        .Others = txtOthers.Text
        .EntDt = Format(Now, CS_DateDbFormat)
        .SaveDoctInfo
        For i = 1 To .CmtCount
            Call .txtCmt(i).SaveTemplate
        Next
    End With
    MsgBox "정상적으로 저장되었습니다.", vbInformation, "메세지"
    m_SaveFg = True
End Sub

Private Sub cmdTempExit_Click()
    txtTempCd.Text = ""
    txtTempNm.Text = ""
    txtTemplate.Text = ""
    fraNewTemp.Visible = False
End Sub

Private Sub cmdTempSave_Click()
        
    Dim NewData As New clsDoctCmt
    With NewData
        .DoctId = objDoctor.DoctId
        .Txtdiv = fraNewTemp.Tag
        .Txtcd = txtTempCd.Text
        Call .GetTmpText
        If Not .NewFg Then
            MsgBox "이미 등록된 코드입니다.", vbExclamation, "메세지"
            txtTempCd.Text = ""
            txtTempCd.SetFocus
            Set NewData = Nothing
            Exit Sub
        End If
        .Txtnm = txtTempNm.Text
        .Txtrst = txtTemplate.Text
        Call .SaveTemplate
        Call objDoctor.AddCmt(NewData, .DoctId & .Txtdiv & .Txtcd)
        If .Txtdiv = "C" Then
            cboCmt.AddItem Format(.Txtcd, "!@@@@@@") & .Txtnm
        Else
            cboRcmd.AddItem Format(.Txtcd, "!@@@@@@") & .Txtnm
        End If
    End With
    Set NewData = Nothing
    fraNewTemp.Visible = False
    
End Sub

Private Sub Form_Load()

    Me.Left = 0
    Me.Top = 0
    
    KeyPreview = True
    Call LoadData
    
End Sub


Private Sub LoadData()
    Dim i As Integer
    
    With objDoctor
        lblDoctNm.Caption = .DoctNm
        txtDoctNo.Text = .Doctno
        txtCertNo.Text = .Certno
        txtDayCnt.Text = .Daycnt
        txtPtCnt.Text = .Ptcnt
        For i = 0 To chkMethod.Count - 1
            chkMethod(i).Value = Val(Mid(.Method, i + 1, 1))
        Next
        txtOthers.Text = .Others
        cboCmt.Clear
        cboRcmd.Clear
        For i = 1 To .CmtCount
            If .txtCmt(i).Txtdiv = "C" Then
                cboCmt.AddItem Format(.txtCmt(i).Txtcd, "!@@@@@@") & .txtCmt(i).Txtnm
            Else
                cboRcmd.AddItem Format(.txtCmt(i).Txtcd, "!@@@@@@") & .txtCmt(i).Txtnm
            End If
        Next
        txtCmt.Enabled = False
        txtRcmd.Enabled = False
    End With
    m_SaveFg = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Dim Resp As VbMsgBoxResult
    
    If Not m_SaveFg Then
        Resp = MsgBox("변경된 데이타를 저장하지 않고 종료하시겠습니까?", vbQuestion + vbYesNo, "메세지")
        If Resp = vbNo Then
            Cancel = True
        End If
    End If
    
End Sub

Private Sub txtCmt_LostFocus()
    If cboCmt.ListIndex >= 0 Then
        Dim strKey As String
        strKey = objDoctor.DoctId & "C" & Trim(Mid(cboCmt.Text, 1, 5))
        objDoctor.txtCmt(strKey).Txtrst = txtCmt.Text
    End If
End Sub

Private Sub txtDayCnt_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtDoctNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtOthers_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtPtCnt_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtRcmd_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtRcmd_LostFocus()
    If cboRcmd.ListIndex >= 0 Then
        Dim strKey As String
        strKey = objDoctor.DoctId & "R" & medGetP(cboRcmd.Text, 1, vbTab)
        objDoctor.txtCmt(strKey).Txtrst = txtRcmd.Text
    End If
End Sub

Private Sub txtTempCd_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtTempNm_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub
