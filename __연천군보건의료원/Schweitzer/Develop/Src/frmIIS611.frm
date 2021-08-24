VERSION 5.00
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmIIS611 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   4  '고정 도구 창
   Caption         =   "장비별 검사항목 설정"
   ClientHeight    =   8925
   ClientLeft      =   4080
   ClientTop       =   285
   ClientWidth     =   11175
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8925
   ScaleWidth      =   11175
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00DBE6E6&
      Caption         =   "종 료(&X)"
      Height          =   495
      Left            =   9900
      Style           =   1  '그래픽
      TabIndex        =   11
      Top             =   8205
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00DBE6E6&
      Caption         =   "화면지움(&C)"
      Height          =   495
      Left            =   8685
      Style           =   1  '그래픽
      TabIndex        =   10
      Top             =   8205
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   8145
      Left            =   5340
      TabIndex        =   16
      Top             =   -30
      Width           =   5790
      Begin VB.TextBox txtIntNm 
         BackColor       =   &H00F7FFF7&
         Height          =   330
         Left            =   255
         MaxLength       =   15
         TabIndex        =   1
         Top             =   1740
         Width           =   3120
      End
      Begin VB.CommandButton cmdTestCdDelete 
         BackColor       =   &H00DBE6E6&
         Caption         =   "검사코드 삭제(&D)"
         Height          =   495
         Left            =   4155
         Style           =   1  '그래픽
         TabIndex        =   9
         Top             =   6150
         Width           =   1545
      End
      Begin VB.CommandButton cmdTestCdSave 
         BackColor       =   &H00DBE6E6&
         Caption         =   "검사코드 추가(&S)"
         Height          =   495
         Left            =   2610
         Style           =   1  '그래픽
         TabIndex        =   8
         Top             =   6150
         Width           =   1545
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00DBE6E6&
         Height          =   330
         Index           =   1
         Left            =   2415
         Picture         =   "frmIIS611.frx":0000
         Style           =   1  '그래픽
         TabIndex        =   13
         Top             =   5460
         Width           =   405
      End
      Begin VB.CommandButton cmdIntNmDelete 
         BackColor       =   &H00DBE6E6&
         Caption         =   "검사명 삭제"
         Height          =   495
         Left            =   4155
         Style           =   1  '그래픽
         TabIndex        =   6
         Top             =   4140
         Width           =   1545
      End
      Begin VB.CommandButton cmdIntNmSave 
         BackColor       =   &H00DBE6E6&
         Caption         =   "검사명 추가"
         Height          =   495
         Left            =   2610
         Style           =   1  '그래픽
         TabIndex        =   5
         Top             =   4140
         Width           =   1545
      End
      Begin VB.TextBox txtTestCd 
         BackColor       =   &H00F7FFF7&
         Height          =   330
         Left            =   255
         MaxLength       =   10
         TabIndex        =   7
         Top             =   5475
         Width           =   2160
      End
      Begin VB.TextBox txtToVal 
         BackColor       =   &H00F7FFF7&
         Height          =   330
         Left            =   255
         MaxLength       =   10
         TabIndex        =   4
         Top             =   4290
         Visible         =   0   'False
         Width           =   2160
      End
      Begin VB.TextBox txtFrVal 
         BackColor       =   &H00F7FFF7&
         Height          =   330
         Left            =   255
         MaxLength       =   10
         TabIndex        =   3
         Top             =   3435
         Visible         =   0   'False
         Width           =   2160
      End
      Begin VB.TextBox txtEqpCd 
         BackColor       =   &H00F7FFF7&
         Height          =   330
         Left            =   255
         MaxLength       =   8
         TabIndex        =   0
         Top             =   645
         Width           =   2160
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00DBE6E6&
         Height          =   330
         Index           =   0
         Left            =   2415
         Picture         =   "frmIIS611.frx":0E42
         Style           =   1  '그래픽
         TabIndex        =   12
         Top             =   630
         Width           =   405
      End
      Begin VB.TextBox txtIntBase 
         BackColor       =   &H00F7FFF7&
         Height          =   330
         Left            =   255
         MaxLength       =   10
         TabIndex        =   2
         Top             =   2595
         Width           =   3120
      End
      Begin MedControls1.LisLabel lblEqpNm 
         Height          =   345
         Left            =   2925
         TabIndex        =   21
         Top             =   630
         Width           =   2625
         _ExtentX        =   4630
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
      Begin MedControls1.LisLabel lblTestNm 
         Height          =   345
         Left            =   2925
         TabIndex        =   25
         Top             =   5460
         Width           =   2625
         _ExtentX        =   4630
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "▶ 검사명"
         Height          =   180
         Left            =   255
         TabIndex        =   26
         Top             =   1425
         Width           =   780
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000011&
         X1              =   75
         X2              =   5700
         Y1              =   4755
         Y2              =   4755
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "▶ 검사코드"
         Height          =   180
         Left            =   255
         TabIndex        =   24
         Top             =   5190
         Width           =   960
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "▶ BorderLine To Value"
         Height          =   180
         Left            =   255
         TabIndex        =   23
         Top             =   4005
         Visible         =   0   'False
         Width           =   1980
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "▶ BorderLine From Value"
         Height          =   180
         Left            =   255
         TabIndex        =   22
         Top             =   3150
         Visible         =   0   'False
         Width           =   2190
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "▶ 인터페이스 기준"
         Height          =   180
         Left            =   255
         TabIndex        =   20
         Top             =   2280
         Width           =   1560
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000011&
         X1              =   75
         X2              =   5700
         Y1              =   1170
         Y2              =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "▶ 장비코드"
         Height          =   180
         Left            =   255
         TabIndex        =   17
         Top             =   330
         Width           =   960
      End
   End
   Begin MSComctlLib.ListView lvwTestCds 
      Height          =   3405
      Left            =   45
      TabIndex        =   15
      Top             =   4710
      Width           =   5280
      _ExtentX        =   9313
      _ExtentY        =   6006
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "검사코드"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "검사명"
         Object.Width           =   6826
      EndProperty
   End
   Begin MSComctlLib.ListView lvwIntNms 
      Height          =   3855
      Left            =   45
      TabIndex        =   14
      Top             =   450
      Width           =   5280
      _ExtentX        =   9313
      _ExtentY        =   6800
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "검사명"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "인터페이스 기준"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "From"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "To"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "EqpCd"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00DBE6E6&
      Caption         =   "※ 검사항목 설정을 임의로 변경하면 인터페이스 장애가 발생할 수 있습니다."
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   90
      TabIndex        =   27
      Top             =   8340
      Width           =   8325
   End
   Begin VB.Label Label5 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "검사코드 리스트"
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
      Left            =   1110
      TabIndex        =   19
      Top             =   4425
      Width           =   1455
   End
   Begin VB.Label lblName 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "장비 검사항목 리스트"
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
      Left            =   870
      TabIndex        =   18
      Top             =   165
      Width           =   1935
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
   Begin VB.Shape Shape2 
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00808080&
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  '단색
      Height          =   375
      Left            =   60
      Top             =   4320
      Width           =   3495
   End
End
Attribute VB_Name = "frmIIS611"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------'
'   파일명  : frmIIS611.frm
'   작성자  :
'   내  용  : 장비별 검사항목 설정폼
'   버  전  :
'-----------------------------------------------------------------------------'

Option Explicit

'## Clear Enum
Private Enum ClearEnum
    ccAll           '모든 컨트롤 Clear
    ccIntNm         '검사명 입력,수정,삭제시에만 Clear
End Enum

'## Check Enum
Private Enum CheckEnum
    ccIntNmSave       '장비기준 검사명 입력,수정시 유효성 Check
    ccIntNmDelete     '장비기준 검사명 삭제시 유효성 Check
    ccIntTestCd       '검사코드 입력,삭제 유효성 Check
End Enum

Private mIntNm            As clsIISIntNm        '장비별 검사항목 설정 클래스
Private WithEvents mCode1 As clsIISCodeList     '코드리스트 클래스1
Attribute mCode1.VB_VarHelpID = -1
Private WithEvents mCode2 As clsIISCodeList     '코드리스트 클래스2
Attribute mCode2.VB_VarHelpID = -1

Private Sub Form_Load()
    With Me
        .Top = 0: .Left = 4030
        .Height = mdiIISMain.ScaleHeight
        
        '   - 모니터의 해상도가 변해도 항상 폼의 ScaleHeight에 맞도록 수정
        .Width = mdiIISMain.ScaleWidth - 4030
    End With

    Set mIntNm = New clsIISIntNm
    Call CtlClear(ccAll)
    Me.Show
    DoEvents
End Sub

Private Sub Form_Activate()
    mdiIISMain.lblMenuNm = Me.Caption
    frmIIS600.tvwMenu.Nodes("IIS611").Selected = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mIntNm = Nothing
    Set frmIIS611 = Nothing
End Sub

Private Sub cmdIntNmSave_Click()
    Dim itmX        As ListItem
    Dim strEqpCd    As String       '장비코드
    Dim strIntNm    As String       '검사명

    If CheckCode(ccIntNmSave) = False Then Exit Sub

    strEqpCd = Trim(txtEqpCd.Text)
    strIntNm = Trim(txtIntNm.Text)

    Me.MousePointer = vbHourglass

    Set itmX = lvwIntNms.FindItem(strIntNm, lvwText)
    With mIntNm
        .EqpCd = strEqpCd
        .IntNm = strIntNm
        .IntBase = Trim(txtIntBase.Text)
        .FrVal = Trim(txtFrVal.Text)
        .ToVal = Trim(txtToVal.Text)

        '## 존재하는 검사명이면 Update, 없으면 Insert
        If itmX Is Nothing Then
            If .AddIntNm Then
                mdiIISMain.sbrStatus.Panels(2).Text = "정상적으로 저장되었습니다."
            Else
                mdiIISMain.sbrStatus.Panels(2).Text = "저장중에 에러가 발생했습니다."
            End If
        Else
            If .ModifyIntNm Then
                mdiIISMain.sbrStatus.Panels(2).Text = "정상적으로 수정되었습니다."
            Else
                mdiIISMain.sbrStatus.Panels(2).Text = "수정중에 에러가 발생했습니다."
            End If
        End If
    End With

    Call CtlClear(ccIntNm)
    Call GetIntNms(strEqpCd)
    txtEqpCd.Text = strEqpCd

    Set itmX = lvwIntNms.FindItem(strIntNm, lvwText)
    If Not (itmX Is Nothing) Then
        lvwIntNms.ListItems(itmX.Index).Selected = True
        lvwIntNms.ListItems(itmX.Index).EnsureVisible
    End If
    Set itmX = Nothing
    txtIntNm.SetFocus

    Me.MousePointer = vbDefault
End Sub

Private Sub cmdIntNmDelete_Click()
    Dim strEqpCd    As String   '장비코드
    Dim strIntNm    As String   '장비기준 검사명
    Dim intTemp     As Integer

    If CheckCode(ccIntNmDelete) = False Then Exit Sub

    intTemp = MsgBox("정말 삭제할까요?", vbYesNo + vbQuestion, "확인")
    If intTemp = vbNo Then Exit Sub

    Me.MousePointer = vbHourglass

    strEqpCd = Trim(txtEqpCd.Text)
    strIntNm = Trim(txtIntNm.Text)

    If mIntNm.DelIntNm(strEqpCd, strIntNm) Then
        mdiIISMain.sbrStatus.Panels(2).Text = "정상적으로 삭제되었습니다."
    Else
        mdiIISMain.sbrStatus.Panels(2).Text = "삭제중에 에러가 발생했습니다."
    End If

    Call CtlClear(ccIntNm)
    txtEqpCd.Text = strEqpCd
    Call GetIntNms(strEqpCd)
    txtIntNm.SetFocus

    Me.MousePointer = vbDefault
End Sub

Private Sub cmdTestCdSave_Click()
    Dim itmX        As ListItem
    Dim strEqpCd    As String       '장비코드
    Dim strIntNm    As String       '검사명
    Dim strTestCd   As String       '검사코드

    If CheckCode(ccIntTestCd) = False Then Exit Sub

    Me.MousePointer = vbHourglass

    strEqpCd = Trim(txtEqpCd.Text)
    strIntNm = Trim(txtIntNm.Text)
    strTestCd = Trim(txtTestCd.Text)

    With mIntNm
        .EqpCd = strEqpCd
        .IntNm = strIntNm
        .TestCd = strTestCd

        '## 무조건 삭제후 저장
        If .AddTestCd Then
            mdiIISMain.sbrStatus.Panels(2).Text = "정상적으로 저장되었습니다."
        Else
            mdiIISMain.sbrStatus.Panels(2).Text = "저장중에 에러가 발생했습니다."
        End If
    End With

    txtTestCd.Text = "": lblTestNm.Caption = ""
    Call GetTestCds(strEqpCd, strIntNm)

    Set itmX = lvwTestCds.FindItem(strTestCd, lvwText)
    If Not (itmX Is Nothing) Then
        lvwTestCds.ListItems(itmX.Index).Selected = True
        lvwTestCds.ListItems(itmX.Index).EnsureVisible
    End If
    Set itmX = Nothing
    txtTestCd.SetFocus

    Me.MousePointer = vbDefault
End Sub

Private Sub cmdTestCdDelete_Click()
    Dim itmX        As ListItem
    Dim strEqpCd    As String       '장비코드
    Dim strIntNm    As String       '검사명
    Dim strTestCd   As String       '검사코드

    If CheckCode(ccIntTestCd) = False Then Exit Sub

    Me.MousePointer = vbHourglass

    strEqpCd = Trim(txtEqpCd.Text)
    strIntNm = Trim(txtIntNm.Text)
    strTestCd = Trim(txtTestCd.Text)

    With mIntNm
        .EqpCd = strEqpCd
        .IntNm = strIntNm
        .TestCd = strTestCd

        If .DelTestCd Then
            mdiIISMain.sbrStatus.Panels(2).Text = "정상적으로 삭제되었습니다."
        Else
            mdiIISMain.sbrStatus.Panels(2).Text = "삭제중에 에러가 발생했습니다."
        End If
    End With

    txtTestCd.Text = "": lblTestNm.Caption = ""
    Call GetTestCds(strEqpCd, strIntNm)
    txtTestCd.SetFocus

    Me.MousePointer = vbDefault
End Sub

Private Sub cmdClear_Click()
    lvwIntNms.ListItems.Clear
    lvwTestCds.ListItems.Clear
    Call CtlClear(ccAll)
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
                .CodeListByRs mIntNm.GetUsingEqp
            End With
            Set mCode1 = Nothing
        Case 1
            Set mCode2 = New clsIISCodeList
            With mCode2
                .Caption = "검사항목 리스트"
                .HeaderCd = "검사코드"
                .HeaderCdNm = "검사명"
                .CodeListByRs mIntNm.GetTestCd
            End With
            Set mCode2 = Nothing
    End Select
End Sub

Private Sub lvwIntNms_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static intOrder1 As Integer

    With lvwIntNms
        .SortKey = ColumnHeader.Index - 1
        .SortOrder = IIf(intOrder1 = 0, lvwAscending, lvwDescending)
        .Sorted = True
        intOrder1 = (intOrder1 + 1) Mod 2
    End With
End Sub

Private Sub lvwIntNms_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Me.MousePointer = vbHourglass

    '## 검사명에 대한 정보표시
    txtIntNm.Text = Item.Text
    txtIntBase.Text = Item.SubItems(1)
    txtFrVal.Text = Item.SubItems(2)
    txtToVal.Text = Item.SubItems(3)
    txtEqpCd.Text = Item.SubItems(4)

    '## 장비기준 검사명에 등록되 검사코드 표시
    Call GetTestCds(Item.SubItems(4), Item.Text)

    Me.MousePointer = vbDefault
End Sub

Private Sub lvwTestCds_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static intOrder2 As Integer

    With lvwTestCds
        .SortKey = ColumnHeader.Index - 1
        .SortOrder = IIf(intOrder2 = 0, lvwAscending, lvwDescending)
        .Sorted = True
        intOrder2 = (intOrder2 + 1) Mod 2
    End With
End Sub
'
Private Sub lvwTestCds_ItemClick(ByVal Item As MSComctlLib.ListItem)
    '## 검사코드, 검사명 표시
    txtTestCd.Text = Item.Text
    lblTestNm.Caption = Item.SubItems(1)
End Sub

Private Sub txtEqpCd_GotFocus()
    With txtEqpCd
        .SelStart = 0
        .SelLength = Len(.Text)
    End With '
End Sub

Private Sub txtEqpCd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtEqpCd_KeyPress(KeyAscii As Integer)
    '## 소문자가 입력되면 대문자로 변경
    If KeyAscii >= 96 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    End If
End Sub

Private Sub txtEqpCd_LostFocus()
    Dim strEqpCd As String      '장비코드
    Dim strEqpNm As String      '장비명

    '## 등록된 장비인지 검사
    '   - 등록된 장비이면 등록된 장비기준 검사명 정보를 표시
    '   - 등록된 장비가 아니면 메시지 표시
    strEqpCd = Trim(txtEqpCd.Text)
    If strEqpCd = "" Then Exit Sub
    Call CtlClear(ccAll)

    Me.MousePointer = vbHourglass

    txtEqpCd.Text = strEqpCd
    strEqpNm = mIntNm.GetEqpNm(strEqpCd)
    If strEqpNm = "" Then
        MsgBox "등록된 장비코드가 아닙니다.", vbInformation, "정보"
        With txtEqpCd
            .SetFocus
            .Text = ""
        End With
    Else
        lblEqpNm.Caption = strEqpNm
        Call GetIntNms(strEqpCd)
    End If

    Me.MousePointer = vbDefault
End Sub

Private Sub txtIntNm_GotFocus()
    With txtIntNm
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtIntNm_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtIntNm_LostFocus()
    Dim itmX        As ListItem
    Dim strIntNm    As String       '장비기준 검사명

    strIntNm = Trim(txtIntNm.Text)
    If strIntNm = "" Then Exit Sub

    With lvwIntNms
        Set itmX = .FindItem(strIntNm, lvwText)
        If Not (itmX Is Nothing) Then
            .ListItems(itmX.Index).Selected = True
            .ListItems(itmX.Index).EnsureVisible
            Call lvwIntNms_ItemClick(itmX)
        Else
            '## 1.2.2:  (2005-05-11)
            '   - 해당 검사명이 없으면 장비기준 검사명을 새로 입력할수 있도록 초기화
            txtIntBase.Text = ""
        End If
        Set itmX = Nothing
    End With
End Sub

Private Sub txtIntBase_GotFocus()
    With txtIntBase
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtIntBase_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtFrVal_GotFocus()
    With txtFrVal
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtFrVal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtFrVal_KeyPress(KeyAscii As Integer)
    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack _
        And KeyAscii <> vbKeyDecimal And KeyAscii <> vbKeyDelete Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtToVal_GotFocus()
    With txtToVal
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtToVal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtToVal_KeyPress(KeyAscii As Integer)
    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack _
        And KeyAscii <> vbKeyDecimal And KeyAscii <> vbKeyDelete Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtTestCd_GotFocus()
    With txtTestCd
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtTestCd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtTestCd_KeyPress(KeyAscii As Integer)
    '## 소문자가 입력되면 대문자로 변경
    If KeyAscii >= 96 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    End If
End Sub

Private Sub txtTestCd_LostFocus()
    Dim itmX      As ListItem
    Dim strTestCd As String     '검사코드
    Dim strTestNm As String     '검사명

    '## 상세자코드, 단일항목인지 체크하여 맞으면 검사명 표시
    '## 틀리면 메시지 표시
    strTestCd = Trim(txtTestCd.Text)
    If strTestCd = "" Then Exit Sub
    lblTestNm.Caption = ""

    strTestNm = mIntNm.GetTestNm(strTestCd)
    If strTestNm = "" Then
        MsgBox "장비별 검사코드는 단일항목, 상세 자코드만 등록할수 있습니다.", vbInformation, "정보"
        With txtTestCd
            .SetFocus
            .Text = ""
        End With
    Else
        lblTestNm.Caption = strTestNm
        With lvwTestCds
            Set itmX = .FindItem(strTestCd, lvwText)
            If Not (itmX Is Nothing) Then
                .ListItems(itmX.Index).Selected = True
                .ListItems(itmX.Index).EnsureVisible
            End If
            Set itmX = Nothing
        End With
    End If
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 해당 장비코드의 장비기준 검사명을 lvwIntNms에 표시
'   인수
'       1. pEqpCd : 장비코드
'-----------------------------------------------------------------------------'
Private Sub GetIntNms(ByVal pEqpCd As String)
    Dim Rs      As ADODB.Recordset
    Dim itmX    As ListItem

On Error GoTo Errors
    With lvwIntNms
        lvwTestCds.ListItems.Clear
        .ListItems.Clear

        Set Rs = mIntNm.GetIntNms(pEqpCd)
        If Not (Rs.BOF Or Rs.EOF) Then
            Do Until Rs.EOF
                Set itmX = .ListItems.Add(, , Rs.Fields("INTNM").Value)
                itmX.SubItems(1) = Rs.Fields("INTBASE").Value & ""
                itmX.SubItems(2) = Rs.Fields("FRVAL").Value & ""
                itmX.SubItems(3) = Rs.Fields("TOVAL").Value & ""
                itmX.SubItems(4) = Rs.Fields("EQPCD").Value
    
                Rs.MoveNext
            Loop
            Set itmX = Nothing
    
            If .ListItems.Count > 18 Then
                .ColumnHeaders(4).Width = 850
            Else
                .ColumnHeaders(4).Width = 1100
            End If
        End If
    End With
    Rs.Close
    Set Rs = Nothing
    Exit Sub

Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "frmIIS611", "GetIntNms", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 해당 장비코드, 장비기준 검사명에 등록된 검사코드 리스트를
'          lvwTestCds에 표시
'   인수 :
'       - pEqpCd : 장비코드
'       - pIntNm : 검사명
'-----------------------------------------------------------------------------'
Private Sub GetTestCds(ByVal pEqpCd As String, ByVal pIntNm As String)
    Dim Rs      As ADODB.Recordset
    Dim itmX    As ListItem

On Error GoTo Errors
    txtTestCd.Text = "": lblTestNm.Caption = ""
    With lvwTestCds
        .ListItems.Clear

        Set Rs = mIntNm.GetTestCds(pEqpCd, pIntNm)
        If Not (Rs.BOF Or Rs.EOF) Then
            Do Until Rs.EOF
                Set itmX = .ListItems.Add(, , Rs.Fields("TESTCD").Value)
                itmX.SubItems(1) = Rs.Fields("TESTNM").Value
                Rs.MoveNext
            Loop
            Set itmX = Nothing
    
            If .ListItems.Count > 15 Then
                .ColumnHeaders(2).Width = 3640
            Else
                .ColumnHeaders(2).Width = 3870
            End If
        End If
    End With
    Rs.Close
    Set Rs = Nothing
    Exit Sub

Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "frmIIS611", "GetTestCds", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 입력, 수정시 필요한 정보의 유효성 Check
'   인수 :
'       1.pFlag : CheckEnum 상수
'   반환 : True(유효), False(무효)
'-----------------------------------------------------------------------------'
Private Function CheckCode(ByVal pFlag As CheckEnum) As Boolean
    '## 장비코드
    If txtEqpCd.Text = "" Then
        MsgBox "장비코드를 입력하세요.", vbInformation, "정보"
        txtEqpCd.SetFocus
        Exit Function
    End If

    '## 검사명
    If txtIntNm.Text = "" Then
        MsgBox "검사명을 입력하세요.", vbInformation, "정보"
        txtIntNm.SetFocus
        Exit Function
    End If

    If pFlag = ccIntTestCd Then
        '## 검사코드
        If txtTestCd.Text = "" Then
            MsgBox "검사코드를 입력하세요.", vbInformation, "정보"
            txtTestCd.SetFocus
            Exit Function
        End If
    End If
    
    CheckCode = True
End Function

'-----------------------------------------------------------------------------'
'   기능 : 컨트롤 초기화
'-----------------------------------------------------------------------------'
Private Sub CtlClear(ByVal pFlag As ClearEnum)
    If pFlag = ccAll Then
        txtEqpCd.Text = "":      lblEqpNm.Caption = ""
    End If
    
    txtIntNm.Text = "":          txtIntBase.Text = ""
    txtFrVal.Text = "":          txtToVal.Text = ""
    txtTestCd.Text = "":         lblTestNm.Caption = ""
End Sub

'-----------------------------------------------------------------------------'
'   기능 : CodeList폼의 이벤트 처리1
'   인수 :
'       - pSelItem : 장비코드|장비명
'-----------------------------------------------------------------------------'
Private Sub mCode1_SelectedItem(ByRef pSelItem As String)
    Me.MousePointer = vbHourglass

    txtEqpCd.Text = mGetP(pSelItem, 1, DIV)
    lblEqpNm.Caption = mGetP(pSelItem, 2, DIV)
    Call CtlClear(ccIntNm)
    Call GetIntNms(txtEqpCd.Text)

    Me.MousePointer = vbDefault
End Sub

'-----------------------------------------------------------------------------'
'   기능 : CodeList폼의 이벤트 처리2
'   인수 :
'       - pSelItem : 검사코드|검사명
'-----------------------------------------------------------------------------'
Private Sub mCode2_SelectedItem(ByRef pSelItem As String)
    Dim itmX As ListItem

    Me.MousePointer = vbHourglass

    txtTestCd.Text = mGetP(pSelItem, 1, DIV)
    lblTestNm.Caption = mGetP(pSelItem, 2, DIV)

    With lvwTestCds
        Set itmX = .FindItem(txtTestCd.Text, lvwText)
        If Not (itmX Is Nothing) Then
            .ListItems(itmX.Index).Selected = True
            .ListItems(itmX.Index).EnsureVisible
        End If
        Set itmX = Nothing
    End With
    Me.MousePointer = vbDefault
End Sub

