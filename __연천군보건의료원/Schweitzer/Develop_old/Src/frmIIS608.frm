VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frmIIS608 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   4  '고정 도구 창
   Caption         =   "다빈도처방 관리"
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
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00DBE6E6&
      Caption         =   "삭 제(&D)"
      Height          =   495
      Left            =   7470
      Style           =   1  '그래픽
      TabIndex        =   3
      Top             =   8205
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00DBE6E6&
      Caption         =   "저 장(&S)"
      Height          =   495
      Left            =   6255
      Style           =   1  '그래픽
      TabIndex        =   2
      Top             =   8205
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00DBE6E6&
      Caption         =   "화면지움(&C)"
      Height          =   495
      Left            =   8685
      Style           =   1  '그래픽
      TabIndex        =   4
      Top             =   8205
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   8145
      Left            =   3585
      TabIndex        =   9
      Top             =   -30
      Width           =   7545
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00DBE6E6&
         Height          =   330
         Index           =   1
         Left            =   2760
         Picture         =   "frmIIS608.frx":0000
         Style           =   1  '그래픽
         TabIndex        =   7
         Top             =   2100
         Width           =   405
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00DBE6E6&
         Height          =   330
         Index           =   0
         Left            =   2760
         Picture         =   "frmIIS608.frx":0E42
         Style           =   1  '그래픽
         TabIndex        =   6
         Top             =   765
         Width           =   405
      End
      Begin VB.TextBox txtRepeatCd 
         BackColor       =   &H00F7FFF7&
         Height          =   330
         Left            =   255
         MaxLength       =   10
         TabIndex        =   0
         Top             =   780
         Width           =   2505
      End
      Begin VB.TextBox txtSpcCd 
         BackColor       =   &H00F7FFF7&
         Height          =   330
         Left            =   255
         MaxLength       =   2
         TabIndex        =   1
         Top             =   2115
         Width           =   2505
      End
      Begin MedControls1.LisLabel lblRepeatNm 
         Height          =   345
         Left            =   3270
         TabIndex        =   13
         Top             =   765
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
      Begin MedControls1.LisLabel lblSpcNm 
         Height          =   345
         Left            =   3270
         TabIndex        =   14
         Top             =   2100
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "▶ 검체코드"
         Height          =   180
         Left            =   255
         TabIndex        =   11
         Top             =   1800
         Width           =   960
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000011&
         X1              =   75
         X2              =   7500
         Y1              =   1380
         Y2              =   1380
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "▶ 다빈도처방 코드"
         Height          =   180
         Left            =   255
         TabIndex        =   10
         Top             =   480
         Width           =   1560
      End
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00DBE6E6&
      Caption         =   "종 료(&X)"
      Height          =   495
      Left            =   9900
      Style           =   1  '그래픽
      TabIndex        =   5
      Top             =   8205
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvwList 
      Height          =   7665
      Left            =   45
      TabIndex        =   8
      Top             =   450
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   13520
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "검사코드"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "검사명"
         Object.Width           =   4251
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "검체코드"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "검체코드명"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label lblName 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "다빈도 처방 리스트"
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
      Left            =   975
      TabIndex        =   12
      Top             =   165
      Width           =   1725
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
End
Attribute VB_Name = "frmIIS608"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------'
'   파일명  : frmIIS608.frm (우리LIS랑 조인할때 사용)
'   작성자  : 이상대
'   내  용  : 다빈도 처방 설정폼
'   작성일  : 2004-02-23
'   버  전  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mRepeat As clsIISRepeat     '다빈도처방 클래스
Private WithEvents mCode1 As clsIISCodeList     '코드리스트 클래스
Attribute mCode1.VB_VarHelpID = -1
Private WithEvents mCode2 As clsIISCodeList     '코드리스트 클래스
Attribute mCode2.VB_VarHelpID = -1

Private Sub Form_Load()
    With Me
        .Top = 0: .Left = 4030
        .Height = mdiIISMain.ScaleHeight: .Width = 11270
    End With
    
    Set mRepeat = New clsIISRepeat
    Call CtlClear
    Me.Show
    DoEvents

    Me.MousePointer = vbHourglass
    '## 다빈도 처방리스트 표시
    Call GetRepeatList
    Me.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
    mdiIISMain.lblMenuNm = Me.Caption
    frmIIS600.tvwMenu.Nodes("IIS608").Selected = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mRepeat = Nothing
    Set frmIIS608 = Nothing
End Sub

Private Sub cmdSave_Click()
    Dim itmX        As ListItem
    Dim strRepeatCd As String   '다빈도처방 코드
    Dim strSpcCd    As String   '검체코드

    '## 입력된 코드의 유효성 Check
    If CheckCode = False Then Exit Sub
    
    strRepeatCd = Trim(txtRepeatCd.Text)
    strSpcCd = Trim(txtSpcCd.Text)
    
    '## 다빈도처방 항목이 존재하면 Update, 없으면 Insert
    Me.MousePointer = vbHourglass
    
    Set itmX = lvwList.FindItem(strRepeatCd, lvwText)
    With mRepeat
        .RepeatCd = strRepeatCd
        .SpcCd = strSpcCd
        If itmX Is Nothing Then
            If .AddRepeatCd Then
                mdiIISMain.sbrStatus.Panels(2).Text = "정상적으로 저장되었습니다."
            Else
                mdiIISMain.sbrStatus.Panels(2).Text = "저장중에 에러가 발생했습니다."
            End If
        Else
            If .ModifyRepeatCd Then
                mdiIISMain.sbrStatus.Panels(2).Text = "정상적으로 수정되었습니다."
            Else
                mdiIISMain.sbrStatus.Panels(2).Text = "수정중에 에러가 발생했습니다."
            End If
        End If
    End With
    Call CtlClear
    Call GetRepeatList
    
    With lvwList
        Set itmX = .FindItem(strRepeatCd, lvwText)
        If Not (itmX Is Nothing) Then
            .ListItems(itmX.Index).Selected = True
            .ListItems(itmX.Index).EnsureVisible
        End If
        Set itmX = Nothing
    End With
    txtRepeatCd.SetFocus
    
    Me.MousePointer = vbDefault
End Sub

Private Sub cmdDelete_Click()
    Dim itmX        As ListItem
    Dim strRepeatCd As String       '다빈도처방 코드
    Dim intTemp     As Integer

    If txtRepeatCd.Text = "" Then
        MsgBox "다빈도처방 코드를 입력하세요.", vbInformation, "정보"
        txtRepeatCd.SetFocus
        Exit Sub
    End If
    
    strRepeatCd = Trim(txtRepeatCd.Text)
    intTemp = MsgBox("정말 삭제할까요?", vbYesNo + vbQuestion, "확인")
    If intTemp = vbNo Then Exit Sub

    Me.MousePointer = vbHourglass
    
    Set itmX = lvwList.FindItem(strRepeatCd, lvwText)
    With mRepeat
        .RepeatCd = strRepeatCd
        If .DelRepeatCd Then
            mdiIISMain.sbrStatus.Panels(2).Text = "정상적으로 삭제되었습니다."
        Else
            mdiIISMain.sbrStatus.Panels(2).Text = "삭제중에 에러가 발생했습니다."
        End If
    End With
    Set itmX = Nothing
    Call CtlClear
    Call GetRepeatList
    txtRepeatCd.SetFocus
    
    Me.MousePointer = vbDefault
End Sub

Private Sub cmdClear_Click()
    Call CtlClear
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSearch_Click(Index As Integer)
    Dim strRepeatCd As String       '다빈도처방 코드
    
    Select Case Index
        Case 0
            Set mCode1 = New clsIISCodeList
            With mCode1
                .Caption = "그룹, 상세모코드 리스트"
                .HeaderCd = "검사코드"
                .HeaderCdNm = "검사명"
                .CodeListBySql mRepeat.GetRepeatCdBySql
            End With
            Set mCode1 = Nothing
        Case 1
            strRepeatCd = Trim(txtRepeatCd.Text)
            If strRepeatCd = "" Then Exit Sub
            
            Set mCode2 = New clsIISCodeList
            With mCode2
                .Caption = "검체리스트"
                .HeaderCd = "검체코드"
                .HeaderCdNm = "검체명"
                .CodeListBySql mRepeat.GetSpcCdBySql(strRepeatCd)
            End With
            Set mCode2 = Nothing
    End Select
End Sub

Private Sub lvwList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static intOrder As Integer

    With lvwList
        .SortKey = ColumnHeader.Index - 1
        .SortOrder = IIf(intOrder = 0, lvwAscending, lvwDescending)
        .Sorted = True
        intOrder = (intOrder + 1) Mod 2
    End With
End Sub

Private Sub lvwList_ItemClick(ByVal Item As MSComctlLib.ListItem)
    '## 해당 다빈도처방 대한 정보 표시
    Call CtlClear
    With Item
        txtRepeatCd.Text = .Text
        lblRepeatNm.Caption = .SubItems(1)
        txtSpcCd.Text = .SubItems(2)
        lblSpcNm.Caption = .SubItems(3)
    End With
End Sub

Private Sub txtRepeatCd_GotFocus()
    With txtRepeatCd
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtRepeatCd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtRepeatCd_KeyPress(KeyAscii As Integer)
    '## 소문자가 입력되면 대문자로 변경
    If KeyAscii >= 96 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    End If
End Sub

Private Sub txtRepeatCd_LostFocus()
    Dim itmX        As ListItem
    Dim strRepeatCd As String       '다빈도처방 코드
    Dim strRepeatNm As String       '다빈도처방 코드명
    
    '## 1.입력된 다빈도처방 코드가 존재하면 해당코드 정보표시
    '## 2.없으면 코드가 유효한지 판단하여 무효하면 메시지출력, 유효하면 새로 입력
    strRepeatCd = Trim(txtRepeatCd.Text)
    If strRepeatCd = "" Then Exit Sub
    lblRepeatNm.Caption = ""
    txtSpcCd.Text = "":         lblSpcNm.Caption = ""

    Set itmX = lvwList.FindItem(strRepeatCd, lvwText)
    If itmX Is Nothing Then
        strRepeatNm = mRepeat.GetRepeatNm(strRepeatCd)
        If strRepeatNm = "" Then
            MsgBox "다빈도처방 코드는 그룹코드, 상세모코드 이어야 합니다.", vbInformation, "정보"
            With txtRepeatCd
                .SetFocus
                .Text = ""
            End With
        Else
            lblRepeatNm.Caption = strRepeatNm
        End If
    Else
        With lvwList
            .ListItems(itmX.Index).Selected = True
            .ListItems(itmX.Index).EnsureVisible
            Call lvwList_ItemClick(itmX)
        End With
    End If
    Set itmX = Nothing
End Sub

Private Sub txtSpcCd_GotFocus()
    With txtSpcCd
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtSpcCd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtSpcCd_KeyPress(KeyAscii As Integer)
    '## 소문자가 입력되면 대문자로 변경
    If KeyAscii >= 96 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    End If
End Sub

Private Sub txtSpcCd_LostFocus()
    Dim strRepeatCd As String       '다빈도처방 코드
    Dim strSpcCd    As String       '검체코드
    Dim strSpcNm    As String       '검체명
    
    strRepeatCd = Trim(txtRepeatCd.Text)
    strSpcCd = Trim(txtSpcCd.Text)
    If strRepeatCd = "" Or strSpcCd = "" Then Exit Sub
    lblSpcNm.Caption = ""
    
    strSpcNm = mRepeat.GetSpcNm(strRepeatCd, strSpcCd)
    If strSpcNm = "" Then
        MsgBox strRepeatCd & "에 등록된 검체코드가 아닙니다.", vbInformation, "정보"
        With txtSpcCd
            .SetFocus
            .Text = ""
        End With
    Else
        lblSpcNm.Caption = strSpcNm
    End If
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 다빈도처방 리스트를 lvwList에 표시
'-----------------------------------------------------------------------------'
Private Sub GetRepeatList()
    Dim Rs      As ADODB.Recordset
    Dim itmX    As ListItem

On Error GoTo Errors
    With lvwList
        .ListItems.Clear
        
        Set Rs = mRepeat.GetRepeatCd
        If Rs.BOF Or Rs.EOF Then GoTo EndLine

        Do Until Rs.EOF
            Set itmX = .ListItems.Add(, , Rs.Fields("TESTCD").Value)
            itmX.SubItems(1) = Rs.Fields("TESTNM").Value
            itmX.SubItems(2) = Rs.Fields("SPCCD").Value
            itmX.SubItems(3) = Rs.Fields("SPCNM").Value
            Rs.MoveNext
        Loop

        If .ListItems.Count > 37 Then
            .ColumnHeaders(2).Width = 2210
        Else
            .ColumnHeaders(2).Width = 2410
        End If
    End With

EndLine:
    Rs.Close
    Set Rs = Nothing
    Set itmX = Nothing
    Exit Sub

Errors:
    Set Rs = Nothing
    Set itmX = Nothing
    Error.SetLog App.EXEName, "frmIIS608", "GetRepeatList", Err.Description, GetSysDate
    MsgBox Error.Description, vbCritical, "오류"
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 입력된 다빈도처방 코드, 검체코드의 유효성 검사
'   반환 : True(유효), False(무효)
'-----------------------------------------------------------------------------'
Private Function CheckCode() As Boolean
    If txtRepeatCd.Text = "" Then
        MsgBox "다빈도처방 코드를 입력하세요.", vbInformation, "정보"
        txtRepeatCd.SetFocus
        Exit Function
    End If

    If txtSpcCd.Text = "" Then
        MsgBox "검체코드를 입력하세요.", vbInformation, "정보"
        txtSpcCd.SetFocus
        Exit Function
    End If
    
    CheckCode = True
End Function

'-----------------------------------------------------------------------------'
'   기능 : CodeList폼의 이벤트 처리1
'-----------------------------------------------------------------------------'
Private Sub mCode1_SelectedItem(ByRef pSelItem As String)
    Dim itmX As ListItem
    
    txtRepeatCd.Text = mGetP(pSelItem, 1, DIV)
    lblRepeatNm.Caption = mGetP(pSelItem, 2, DIV)
    
    With lvwList
        Set itmX = .FindItem(txtRepeatCd.Text, lvwText)
        If Not (itmX Is Nothing) Then
            .ListItems(itmX.Index).Selected = True
            .ListItems(itmX.Index).EnsureVisible
            Call lvwList_ItemClick(itmX)
        End If
        Set itmX = Nothing
    End With
End Sub

'-----------------------------------------------------------------------------'
'   기능 : CodeList폼의 이벤트 처리2
'-----------------------------------------------------------------------------'
Private Sub mCode2_SelectedItem(ByRef pSelItem As String)
    txtSpcCd.Text = mGetP(pSelItem, 1, DIV)
    lblSpcNm.Caption = mGetP(pSelItem, 2, DIV)
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 컨트롤 초기화
'-----------------------------------------------------------------------------'
Private Sub CtlClear()
    txtRepeatCd.Text = "":      txtSpcCd.Text = ""
    lblRepeatNm.Caption = "":   lblSpcNm.Caption = ""
End Sub
