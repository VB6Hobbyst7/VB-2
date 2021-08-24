VERSION 5.00
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frmIIS612 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   4  '고정 도구 창
   Caption         =   "사용장비 선택"
   ClientHeight    =   8925
   ClientLeft      =   4080
   ClientTop       =   285
   ClientWidth     =   11175
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8925
   ScaleWidth      =   11175
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H00DBE6E6&
      Height          =   3585
      Left            =   60
      TabIndex        =   11
      Top             =   4530
      Width           =   11070
      Begin VB.CheckBox chkAutoRerun2 
         BackColor       =   &H00DBE6E6&
         Caption         =   "▶ 자동재검 사용유무"
         Height          =   180
         Left            =   255
         TabIndex        =   28
         Top             =   2490
         Value           =   1  '확인
         Width           =   2175
      End
      Begin VB.OptionButton optBarPos2 
         BackColor       =   &H00DBE6E6&
         Caption         =   "장비"
         Height          =   180
         Index           =   1
         Left            =   5490
         TabIndex        =   23
         Top             =   2055
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CheckBox chkHLCheck2 
         BackColor       =   &H00DBE6E6&
         Caption         =   "▶ 자동 결과등록시 H/L결과는 제외"
         Enabled         =   0   'False
         Height          =   180
         Left            =   2850
         TabIndex        =   25
         Top             =   1620
         Width           =   3210
      End
      Begin VB.OptionButton optBarPos2 
         BackColor       =   &H00DBE6E6&
         Caption         =   "PC"
         Height          =   180
         Index           =   0
         Left            =   4875
         TabIndex        =   24
         Top             =   2055
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CheckBox chkBarcode2 
         BackColor       =   &H00DBE6E6&
         Caption         =   "▶ 바코드 사용유무"
         Height          =   180
         Left            =   255
         TabIndex        =   19
         Top             =   2055
         Value           =   1  '확인
         Width           =   2175
      End
      Begin VB.TextBox txtEqpCd2 
         BackColor       =   &H00F7FFF7&
         Height          =   330
         Left            =   255
         MaxLength       =   8
         TabIndex        =   14
         Top             =   690
         Width           =   2160
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00DBE6E6&
         Height          =   330
         Index           =   1
         Left            =   2415
         Picture         =   "frmIIS612.frx":0000
         Style           =   1  '그래픽
         TabIndex        =   13
         Top             =   675
         Width           =   405
      End
      Begin VB.CheckBox chkAutoVerify2 
         BackColor       =   &H00DBE6E6&
         Caption         =   "▶ 자동 결과등록 사용"
         Height          =   180
         Left            =   255
         TabIndex        =   12
         Top             =   1620
         Width           =   2175
      End
      Begin MedControls1.LisLabel lblEqpNm2 
         Height          =   345
         Left            =   2925
         TabIndex        =   15
         Top             =   675
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   609
         BackColor       =   16252919
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   ""
         LeftGab         =   100
      End
      Begin VB.Label lblBarPos2 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "▶ 바코드 리더기 위치 :"
         Height          =   180
         Left            =   2850
         TabIndex        =   26
         Top             =   2055
         Visible         =   0   'False
         Width           =   1920
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "▶ 장비코드"
         Height          =   180
         Left            =   255
         TabIndex        =   16
         Top             =   375
         Width           =   960
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000011&
         X1              =   135
         X2              =   10930
         Y1              =   1230
         Y2              =   1230
      End
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00DBE6E6&
      Caption         =   "종 료(&X)"
      Height          =   495
      Left            =   9900
      Style           =   1  '그래픽
      TabIndex        =   4
      Top             =   8205
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00DBE6E6&
      Caption         =   "화면지움(&C)"
      Height          =   495
      Left            =   8685
      Style           =   1  '그래픽
      TabIndex        =   3
      Top             =   8205
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00DBE6E6&
      Caption         =   "저 장(&S)"
      Height          =   495
      Left            =   7470
      Style           =   1  '그래픽
      TabIndex        =   2
      Top             =   8205
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   3585
      Left            =   60
      TabIndex        =   1
      Top             =   435
      Width           =   11070
      Begin VB.CheckBox chkAutoRerun1 
         BackColor       =   &H00DBE6E6&
         Caption         =   "▶ 자동재검 사용유무"
         Height          =   180
         Left            =   255
         TabIndex        =   27
         Top             =   2490
         Value           =   1  '확인
         Width           =   2175
      End
      Begin VB.OptionButton optBarPos1 
         BackColor       =   &H00DBE6E6&
         Caption         =   "장비"
         Height          =   180
         Index           =   1
         Left            =   5490
         TabIndex        =   22
         Top             =   2055
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.OptionButton optBarPos1 
         BackColor       =   &H00DBE6E6&
         Caption         =   "PC"
         Height          =   180
         Index           =   0
         Left            =   4875
         TabIndex        =   21
         Top             =   2055
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CheckBox chkBarcode1 
         BackColor       =   &H00DBE6E6&
         Caption         =   "▶ 바코드 사용유무"
         Height          =   180
         Left            =   255
         TabIndex        =   18
         Top             =   2055
         Value           =   1  '확인
         Width           =   2175
      End
      Begin VB.CheckBox chkHLCheck1 
         BackColor       =   &H00DBE6E6&
         Caption         =   "▶ 자동 결과등록시 H/L결과는 제외"
         Enabled         =   0   'False
         Height          =   180
         Left            =   2850
         TabIndex        =   17
         Top             =   1620
         Width           =   3210
      End
      Begin VB.CheckBox chkAutoVerify1 
         BackColor       =   &H00DBE6E6&
         Caption         =   "▶ 자동 결과등록 사용"
         Height          =   180
         Left            =   255
         TabIndex        =   10
         Top             =   1620
         Width           =   2175
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00DBE6E6&
         Height          =   330
         Index           =   0
         Left            =   2415
         Picture         =   "frmIIS612.frx":0E42
         Style           =   1  '그래픽
         TabIndex        =   5
         Top             =   675
         Width           =   405
      End
      Begin VB.TextBox txtEqpCd1 
         BackColor       =   &H00F7FFF7&
         Height          =   330
         Left            =   255
         MaxLength       =   8
         TabIndex        =   0
         Top             =   690
         Width           =   2160
      End
      Begin MedControls1.LisLabel lblEqpNm1 
         Height          =   345
         Left            =   2925
         TabIndex        =   9
         Top             =   675
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   609
         BackColor       =   16252919
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   ""
         LeftGab         =   100
      End
      Begin VB.Label lblBarPos1 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "▶ 바코드 리더기 위치 :"
         Height          =   180
         Left            =   2850
         TabIndex        =   20
         Top             =   2055
         Visible         =   0   'False
         Width           =   1920
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000011&
         X1              =   135
         X2              =   10930
         Y1              =   1230
         Y2              =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "▶ 장비코드"
         Height          =   180
         Left            =   255
         TabIndex        =   6
         Top             =   375
         Width           =   960
      End
   End
   Begin VB.Label lblName 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "검사장비 1."
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   1305
      TabIndex        =   7
      Top             =   165
      Width           =   1065
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00808080&
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  '단색
      Height          =   375
      Left            =   60
      Top             =   60
      Width           =   3495
   End
   Begin VB.Label Label5 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "검사장비 2."
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   1305
      TabIndex        =   8
      Top             =   4260
      Width           =   1065
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00808080&
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  '단색
      Height          =   375
      Left            =   60
      Top             =   4155
      Width           =   3495
   End
End
Attribute VB_Name = "frmIIS612"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------'
'   파일명  : frmIIS612.frm
'   작성자  : 오세원
'   내  용  : 현재 PC에서 사용할 검사장비 선택폼
'   작성일  : 2015-10-30
'   버  전  : 1.0.0
'-----------------------------------------------------------------------------'

Option Explicit

Private mEqpChoice        As clsIISEqpChoice    '사용장비 선택 클래스
Private WithEvents mCode1 As clsIISCodeList     '코드리스트 클래스
Attribute mCode1.VB_VarHelpID = -1
Private WithEvents mCode2 As clsIISCodeList     '코드리스트 클래스
Attribute mCode2.VB_VarHelpID = -1

Private Sub Form_Load()
    With Me
        .Top = 0: .Left = 4030
        .Height = mdiIISMain.ScaleHeight
        
        '   - 모니터의 해상도가 변해도 항상 폼의 ScaleHeight에 맞도록 수정
        .Width = mdiIISMain.ScaleWidth - 4030
    End With

    Set mEqpChoice = New clsIISEqpChoice
    Call CtlClear
    Me.Show
    DoEvents
    
    '## 현재 PC에 설정된 장비표시
    Call GetEqpList
End Sub

Private Sub Form_Activate()
    mdiIISMain.lblMenuNm = Me.Caption
    frmIIS600.tvwMenu.Nodes("IIS612").Selected = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mEqpChoice = Nothing
    Set frmIIS612 = Nothing
End Sub

Private Sub cmdSave_Click()
    Dim objMenu As clsIISHopMenu        '병원별 메뉴설정 클래스
    Dim objHop  As clsIISMenuInfo       '툴바구성 클래스
    
    '## 레지에 저장
    With mEqpChoice
        '## 검사장비1
        .EqpCd1 = Trim(txtEqpCd1.Text)
        .Barcode1 = chkBarcode1.Value
        .AutoVfy1 = chkAutoVerify1.Value
        .HLCheck1 = chkHLCheck1.Value
        .BarPos1 = IIf(optBarPos1(0).Value = True, 0, 1)
        
        '## 검사장비2
        .EqpCd2 = Trim(txtEqpCd2.Text)
        .Barcode2 = chkBarcode2.Value
        .AutoVfy2 = chkAutoVerify2.Value
        .HLCheck2 = chkHLCheck2.Value
        .BarPos2 = IIf(optBarPos2(0).Value = True, 0, 1)
        
        '   - 자동재검 사용유무 옵션추가
        .AutoRerun1 = chkAutoRerun1.Value
        .AutoRerun2 = chkAutoRerun2.Value
        
        If .SetEqp Then
            mdiIISMain.sbrStatus.Panels(2).Text = "정상적으로 저장되었습니다."
            
            '## 메인폼의 풀다운메뉴, 툴바 다시표시
            Set objMenu = New clsIISHopMenu
            Set objHop = New clsIISMenuInfo
            
            Call objMenu.GetFullMenu
            Call objHop.GetToolbar
            Set objMenu = Nothing
            Set objHop = Nothing
        Else
            mdiIISMain.sbrStatus.Panels(2).Text = "저장중에 에러가 발생했습니다."
        End If
    End With
End Sub

Private Sub cmdClear_Click()
    Call CtlClear
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSearch_Click(Index As Integer)
    Select Case Index
        Case 0
            Set mCode1 = New clsIISCodeList
            With mCode1
                .Caption = "검사장비 리스트"
                .HeaderCd = "장비코드"
                .HeaderCdNm = "장비명"
                .CodeListByRs mEqpChoice.GetUsingEqp
            End With
            Set mCode1 = Nothing
        Case 1
            Set mCode2 = New clsIISCodeList
            With mCode2
                .Caption = "검사장비 리스트"
                .HeaderCd = "장비코드"
                .HeaderCdNm = "장비명"
                .CodeListByRs mEqpChoice.GetUsingEqp
            End With
            Set mCode2 = Nothing
    End Select
End Sub

Private Sub chkAutoVerify1_Click()
    If chkAutoVerify1.Value = 0 Then
        chkHLCheck1.Value = 0
        chkHLCheck1.Enabled = False
    Else
        chkHLCheck1.Enabled = True
    End If
End Sub

Private Sub chkAutoVerify2_Click()
    If chkAutoVerify2.Value = 0 Then
        chkHLCheck1.Value = 0
        chkHLCheck2.Enabled = False
    Else
        chkHLCheck2.Enabled = True
    End If
End Sub

Private Sub chkBarcode1_Click()
    If chkBarcode1.Value = BarcodeUseEnum.ccUseBarcode Then
        lblBarPos1.Visible = True
        optBarPos1(0).Visible = True
        optBarPos1(1).Visible = True
    Else
        lblBarPos1.Visible = False
        optBarPos1(0).Visible = False
        optBarPos1(1).Visible = False
    End If
End Sub

Private Sub chkBarcode2_Click()
    If chkBarcode2.Value = BarcodeUseEnum.ccUseBarcode Then
        lblBarPos2.Visible = True
        optBarPos2(0).Visible = True
        optBarPos2(1).Visible = True
    Else
        lblBarPos2.Visible = False
        optBarPos2(0).Visible = False
        optBarPos2(1).Visible = False
    End If
End Sub

Private Sub txtEqpCd1_GotFocus()
    With txtEqpCd1
        .SelStart = 0
        .SelLength = Len(.Text)
    End With '
End Sub

Private Sub txtEqpCd1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtEqpCd1_KeyPress(KeyAscii As Integer)
    '## 소문자가 입력되면 대문자로 변경
    If KeyAscii >= 96 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    End If
End Sub

Private Sub txtEqpCd1_LostFocus()
    Dim strEqpCd As String      '장비코드
    Dim strEqpNm As String      '장비명
    
    '## 입력된 장비코드가 현재 사용중인 장비인지 검사
    strEqpCd = Trim(txtEqpCd1.Text)
    If strEqpCd = "" Then Exit Sub
    lblEqpNm1.Caption = "": chkAutoVerify1.Value = 0
    
    If strEqpCd = Trim(txtEqpCd2.Text) Then
        MsgBox "해당 장비코드는 이미 선택되어 있습니다.", vbInformation, "정보"
        With txtEqpCd1
            .SetFocus
            .Text = ""
        End With
        Exit Sub
    End If
    
    strEqpNm = mEqpChoice.GetEqpNm(strEqpCd)
    If strEqpNm = "" Then
        MsgBox "등록된 장비코드가 아닙니다.", vbInformation, "정보"
        With txtEqpCd1
            .SetFocus
            .Text = ""
        End With
    Else
        lblEqpNm1.Caption = strEqpNm
    End If
End Sub

Private Sub txtEqpCd2_GotFocus()
    With txtEqpCd2
        .SelStart = 0
        .SelLength = Len(.Text)
    End With '
End Sub

Private Sub txtEqpCd2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtEqpCd2_KeyPress(KeyAscii As Integer)
    '## 소문자가 입력되면 대문자로 변경
    If KeyAscii >= 96 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    End If
End Sub

Private Sub txtEqpCd2_LostFocus()
    Dim strEqpCd As String      '장비코드
    Dim strEqpNm As String      '장비명
    
    '## 입력된 장비코드가 현재 사용중인 장비인지 검사
    strEqpCd = Trim(txtEqpCd2.Text)
    If strEqpCd = "" Then Exit Sub
    lblEqpNm2.Caption = "": chkAutoVerify2.Value = 0
    
    If strEqpCd = Trim(txtEqpCd1.Text) Then
        MsgBox "해당 장비코드는 이미 선택되어 있습니다.", vbInformation, "정보"
        With txtEqpCd2
            .SetFocus
            .Text = ""
        End With
        Exit Sub
    End If
    
    strEqpNm = mEqpChoice.GetEqpNm(strEqpCd)
    If strEqpNm = "" Then
        MsgBox "등록된 장비코드가 아닙니다.", vbInformation, "정보"
        With txtEqpCd2
            .SetFocus
            .Text = ""
        End With
    Else
        lblEqpNm2.Caption = strEqpNm
    End If
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 현재 PC에 설정된 장비를 표시
'-----------------------------------------------------------------------------'
Private Sub GetEqpList()
    With mEqpChoice
        If .GetEqp Then
            '## 검사장비1
            txtEqpCd1.Text = .EqpCd1
            lblEqpNm1.Caption = .EqpNm1
            chkAutoVerify1.Value = .AutoVfy1
            If chkAutoVerify1.Value = AutoVfyEnum.ccYes Then
                chkHLCheck1.Enabled = True
                chkHLCheck1.Value = .HLCheck1
            Else
                chkHLCheck1.Enabled = False
            End If
            
            chkBarcode1.Value = .Barcode1
            If chkBarcode1.Value = BarcodeUseEnum.ccUseBarcode Then
                lblBarPos1.Visible = True
                optBarPos1(0).Visible = True
                optBarPos1(1).Visible = True
                If .BarPos1 = ccPC Then
                    optBarPos1(0).Value = True
                Else
                    optBarPos1(1).Value = True
                End If
            Else
                lblBarPos1.Visible = False
                optBarPos1(0).Visible = False
                optBarPos1(1).Visible = False
            End If
                        
            '## 검사장비2
            txtEqpCd2.Text = .EqpCd2
            lblEqpNm2.Caption = .EqpNm2
            chkAutoVerify2.Value = .AutoVfy2
            If chkAutoVerify2.Value = 1 Then
                chkHLCheck2.Enabled = True
                chkHLCheck2.Value = .HLCheck2
            Else
                chkHLCheck2.Enabled = False
            End If
            
            chkBarcode2.Value = .Barcode2
            If chkBarcode2.Value = BarcodeUseEnum.ccUseBarcode Then
                lblBarPos2.Visible = True
                optBarPos2(0).Visible = True
                optBarPos2(1).Visible = True
                If .BarPos2 = ccPC Then
                    optBarPos2(0).Value = True
                Else
                    optBarPos2(1).Value = True
                End If
            Else
                lblBarPos2.Visible = False
                optBarPos2(0).Visible = False
                optBarPos2(1).Visible = False
            End If
            
            '   - 자동재검 사용유무 옵션 추가
            chkAutoRerun1.Value = .AutoRerun1
            chkAutoRerun2.Value = .AutoRerun2
        Else
            MsgBox "현재 PC에 설정된 장비를 표시중 에러가 발생했습니다.", vbCritical, "오류"
        End If
    End With
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 컨트롤 초기화
'-----------------------------------------------------------------------------'
Private Sub CtlClear()
    txtEqpCd1.Text = "":        lblEqpNm1.Caption = ""
    chkAutoVerify1.Value = 0:   txtEqpCd2.Text = ""
    lblEqpNm2.Caption = "":     chkAutoVerify2.Value = 0
End Sub

'-----------------------------------------------------------------------------'
'   기능 : CodeList폼의 이벤트 처리1
'-----------------------------------------------------------------------------'
Private Sub mCode1_SelectedItem(ByRef pSelItem As String)
    Dim strEqpCd As String      '장비코드
    
    strEqpCd = mGetP(pSelItem, 1, DIV)
    If strEqpCd = Trim(txtEqpCd2.Text) Then
        MsgBox "해당 장비코드는 이미 선택되어 있습니다.", vbInformation, "정보"
        pSelItem = ""
    Else
        txtEqpCd1.Text = strEqpCd
        lblEqpNm1.Caption = mGetP(pSelItem, 2, DIV)
    End If
End Sub

'-----------------------------------------------------------------------------'
'   기능 : CodeList폼의 이벤트 처리2
'-----------------------------------------------------------------------------'
Private Sub mCode2_SelectedItem(ByRef pSelItem As String)
    Dim strEqpCd As String      '장비코드
    
    strEqpCd = mGetP(pSelItem, 1, DIV)
    If strEqpCd = Trim(txtEqpCd1.Text) Then
        MsgBox "해당 장비코드는 이미 선택되어 있습니다.", vbInformation, "정보"
        pSelItem = ""
    Else
        txtEqpCd2.Text = strEqpCd
        lblEqpNm2.Caption = mGetP(pSelItem, 2, DIV)
    End If
End Sub

