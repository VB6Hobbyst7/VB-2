VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frmIIS606 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   4  '고정 도구 창
   Caption         =   "상세항목 관리"
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
   Begin VB.CommandButton cmdAllDelete 
      BackColor       =   &H00DBE6E6&
      Caption         =   "모두삭제(&A)"
      Height          =   495
      Left            =   6255
      Style           =   1  '그래픽
      TabIndex        =   4
      ToolTipText     =   "상세 모코드에 해당하는 상세 자코드를 모두삭제합니다."
      Top             =   8205
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00DBE6E6&
      Caption         =   "종 료(&X)"
      Height          =   495
      Left            =   9900
      Style           =   1  '그래픽
      TabIndex        =   7
      Top             =   8205
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00DBE6E6&
      Caption         =   "화면지움(&C)"
      Height          =   495
      Left            =   8685
      Style           =   1  '그래픽
      TabIndex        =   6
      Top             =   8205
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00DBE6E6&
      Caption         =   "저 장(&S)"
      Height          =   495
      Left            =   5040
      Style           =   1  '그래픽
      TabIndex        =   3
      Top             =   8205
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00DBE6E6&
      Caption         =   "삭 제(&D)"
      Height          =   495
      Left            =   7470
      Style           =   1  '그래픽
      TabIndex        =   5
      Top             =   8205
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   8145
      Left            =   3585
      TabIndex        =   2
      Top             =   -30
      Width           =   7545
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00DBE6E6&
         Height          =   330
         Index           =   1
         Left            =   2415
         Picture         =   "frmIIS606.frx":0000
         Style           =   1  '그래픽
         TabIndex        =   9
         Top             =   1560
         Width           =   405
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00DBE6E6&
         Height          =   330
         Index           =   0
         Left            =   2415
         Picture         =   "frmIIS606.frx":0E42
         Style           =   1  '그래픽
         TabIndex        =   8
         Top             =   780
         Width           =   405
      End
      Begin VB.TextBox txtChild 
         BackColor       =   &H00F7FFF7&
         Height          =   330
         Left            =   255
         MaxLength       =   10
         TabIndex        =   1
         Top             =   1575
         Width           =   2160
      End
      Begin VB.TextBox txtParent 
         BackColor       =   &H00F7FFF7&
         Height          =   330
         Left            =   255
         MaxLength       =   10
         TabIndex        =   0
         Top             =   795
         Width           =   2160
      End
      Begin MedControls1.LisLabel lblParentNm 
         Height          =   345
         Left            =   2925
         TabIndex        =   16
         Top             =   780
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
      Begin MedControls1.LisLabel lblChildNm 
         Height          =   345
         Left            =   2925
         TabIndex        =   17
         Top             =   1560
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
         Caption         =   "▶ 상세 자코드"
         Height          =   180
         Left            =   255
         TabIndex        =   13
         Top             =   1260
         Width           =   1200
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000011&
         X1              =   75
         X2              =   7500
         Y1              =   2115
         Y2              =   2115
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "▶ 상세 모코드"
         Height          =   180
         Left            =   255
         TabIndex        =   12
         Top             =   480
         Width           =   1200
      End
   End
   Begin MSComctlLib.ListView lvwParent 
      Height          =   2820
      Left            =   45
      TabIndex        =   10
      Top             =   450
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   4974
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
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "검사명"
         Object.Width           =   4251
      EndProperty
   End
   Begin MSComctlLib.ListView lvwChild 
      Height          =   4440
      Left            =   45
      TabIndex        =   11
      Top             =   3690
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   7832
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
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "검사명"
         Object.Width           =   4251
      EndProperty
   End
   Begin VB.Label Label5 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "상세 자코드 리스트"
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
      TabIndex        =   15
      Top             =   3405
      Width           =   1725
   End
   Begin VB.Label lblName 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "상세 모코드 리스트"
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
      TabIndex        =   14
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
   Begin VB.Shape Shape2 
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00808080&
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  '단색
      Height          =   375
      Left            =   60
      Top             =   3300
      Width           =   3495
   End
End
Attribute VB_Name = "frmIIS606"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------'
'   파일명  : frmIIS606.frm
'   작성자  : 이상대
'   내  용  : 상세항목 설정폼
'   작성일  : 2004-02-20
'   버  전  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mDetail As clsIISDetail     '상세항목 클래스
Private WithEvents mCode1 As clsIISCodeList     '코드리스트 클래스
Attribute mCode1.VB_VarHelpID = -1
Private WithEvents mCode2 As clsIISCodeList     '코드리스트 클래스
Attribute mCode2.VB_VarHelpID = -1

Private mTestCd As String           '검사코드

Public Property Let TestCd(ByVal vData As String)
    mTestCd = vData
End Property

Private Sub Form_Load()
    With Me
        .Top = 0: .Left = 4030
        .Height = mdiIISMain.ScaleHeight: .Width = 11270
    End With

    Set mDetail = New clsIISDetail
    Call CtlClear
    Me.Show
    DoEvents

    Me.MousePointer = vbHourglass
    
    '## 모코드 리스트 조회
    Call GetParentList
    
    '## 검사코드 마스터에서 폼을 표시하는 경우 해당 검사코드의 상세 자코드 표시
    If mTestCd <> "" Then
        txtParent.Text = mTestCd
        Call txtParent_LostFocus
        txtChild.SetFocus
    End If
    
    Me.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
    mdiIISMain.lblMenuNm = Me.Caption
    frmIIS600.tvwMenu.Nodes("IIS606").Selected = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mDetail = Nothing
    Set frmIIS606 = Nothing
End Sub

Private Sub cmdSave_Click()
    Dim itmX        As ListItem
    Dim strParentCd As String   '모코드
    Dim strChildCd  As String   '자코드
    
    If CheckCode() = False Then Exit Sub
    
    strParentCd = Trim(txtParent.Text)
    strChildCd = Trim(txtChild.Text)
    
    Set itmX = lvwChild.FindItem(strChildCd, lvwText)
    If Not (itmX Is Nothing) Then
        MsgBox "이미 존재하는 상세 자코드 입니다.", vbInformation, "정보"
        Set itmX = Nothing
        Exit Sub
    End If
    
    '## 현재 존재하는 모코드+자코드인지 파악해서 존재하면 메세지 표시
    '## 존재하지 않으면 입력
    Me.MousePointer = vbHourglass
    
    With mDetail
        .ParentCd = strParentCd
        .ChildCd = strChildCd
        If .AddDetail Then
            mdiIISMain.sbrStatus.Panels(2).Text = "정상적으로 저장되었습니다."
        Else
            mdiIISMain.sbrStatus.Panels(2).Text = "저장중에 에러가 발생했습니다."
        End If
    End With
    Call CtlClear
    Call GetParentList
    
    With lvwParent
        Set itmX = .FindItem(strParentCd, lvwText)
        If Not (itmX Is Nothing) Then
            .ListItems(itmX.Index).Selected = True
            .ListItems(itmX.Index).EnsureVisible
            Call lvwParent_ItemClick(itmX)
        End If
        Set itmX = Nothing
    End With
    txtChild.SetFocus
    
    Me.MousePointer = vbDefault
End Sub

Private Sub cmdAllDelete_Click()
    Dim strParentCd As String   '모코드
    Dim intTemp     As Integer
    
    If strParentCd = "" Then
        MsgBox "대표항목 코드를 입력하세요.", vbInformation, "정보"
        txtParent.SetFocus
        Exit Sub
    End If
    
    strParentCd = Trim(txtParent.Text)
    
    intTemp = MsgBox("대표항목에 포함된 모든 상세항목이 삭제됩니다. 정말 삭제할까요?", vbYesNo + vbQuestion, "확인")
    If intTemp = vbNo Then Exit Sub
    
    Me.MousePointer = vbHourglass
    
    '## 상세모코드에 포함된 모든 자코드 삭제
    With mDetail
        .ParentCd = strParentCd
        If .DelDetailAll Then
            mdiIISMain.sbrStatus.Panels(2).Text = "정상적으로 삭제되었습니다."
        Else
            mdiIISMain.sbrStatus.Panels(2).Text = "삭제중에 에러가 발생했습니다."
        End If
    End With
    Call CtlClear
    Call GetParentList
    txtParent.SetFocus
    
    Me.MousePointer = vbDefault
End Sub

Private Sub cmdDelete_Click()
    Dim itmX        As ListItem
    Dim strParentCd As String   '모코드
    Dim strChildCd  As String   '자코드
    Dim intTemp     As Integer
    
    If CheckCode = False Then Exit Sub
    
    strParentCd = Trim(txtParent.Text)
    strChildCd = Trim(txtChild.Text)
    
    intTemp = MsgBox("정말 삭제할까요?", vbYesNo + vbQuestion, "확인")
    If intTemp = vbNo Then Exit Sub
    
    Me.MousePointer = vbHourglass
    
    With mDetail
        .ParentCd = strParentCd
        .ChildCd = strChildCd
        If .DelDetail Then
            mdiIISMain.sbrStatus.Panels(2).Text = "정상적으로 삭제되었습니다."
        Else
            mdiIISMain.sbrStatus.Panels(2).Text = "삭제중에 에러가 발생했습니다."
        End If
    End With
    Call CtlClear
    Call GetParentList
    
    With lvwParent
        Set itmX = .FindItem(strParentCd, lvwText)
        If Not (itmX Is Nothing) Then
            .ListItems(itmX.Index).Selected = True
            .ListItems(itmX.Index).EnsureVisible
            Call lvwParent_ItemClick(itmX)
        End If
        Set itmX = Nothing
    End With
    txtChild.SetFocus
    
    Me.MousePointer = vbDefault
End Sub

Private Sub cmdClear_Click()
    Call CtlClear
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub txtParent_GotFocus()
    With txtParent
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtParent_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtParent_KeyPress(KeyAscii As Integer)
    '## 소문자가 입력되면 대문자로 변경
    If KeyAscii >= 96 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    End If
End Sub

Private Sub txtParent_LostFocus()
    Dim itmX        As ListItem
    Dim strParentCd As String   '상세 모코드
    Dim strParentNm As String   '상세 자코드
    
    '## 1.입력된 검사코드가 lvwParent에 존재하는 경우 해당코드의 정보를 표시하고
    '## 2.존재하지 않으면 코드의 모코드여부를 파악하여 모코드가 아니면 경고메시지,
    '   모코드이면 새로입력할수 있도록 한다.
    strParentCd = Trim(txtParent.Text)
    If strParentCd = "" Then Exit Sub
    lblParentNm.Caption = ""
    txtChild.Text = "":         lblChildNm.Caption = ""
    
    Set itmX = lvwParent.FindItem(strParentCd, lvwText)
    If itmX Is Nothing Then
        '## 입력된 코드가 존재하지 않는 경우
        strParentNm = mDetail.GetParentNm(strParentCd)
        If strParentNm = "" Then
            MsgBox "입력한 코드는 상세 모코드가 아닙니다.", vbInformation, "정보"
            With txtParent
                .SetFocus
                .Text = ""
            End With
        Else
            lblParentNm.Caption = strParentNm
        End If
    Else
        '## 입력된 코드가 존재하는 경우
        With lvwParent
            .ListItems(itmX.Index).Selected = True
            .ListItems(itmX.Index).EnsureVisible
            Call lvwParent_ItemClick(itmX)
        End With
    End If
    Set itmX = Nothing
End Sub

Private Sub txtChild_GotFocus()
    With txtChild
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtChild_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtChild_KeyPress(KeyAscii As Integer)
    '## 소문자가 입력되면 대문자로 변경
    If KeyAscii >= 96 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    End If
End Sub

Private Sub txtChild_LostFocus()
    Dim itmX        As ListItem
    Dim strChildCd  As String       '자코드
    Dim strChildNm  As String       '자코드명
    
    '## 1.입력된 검사코드가 lvwChild 존재하는 경우 해당코드의 정보를 표시하고
    '## 2.존재하지 않으면 코드의 자코드여부를 파악하여 자코드가 아니면 경고메시지,
    '   자코드이면 새로입력할수 있도록 한다.
    strChildCd = Trim(txtChild.Text)
    If strChildCd = "" Then Exit Sub
    lblChildNm.Caption = ""
    
    Set itmX = lvwChild.FindItem(strChildCd, lvwText)
    If itmX Is Nothing Then
        '## 입력된 코드가 존재하지 않는 경우
        strChildNm = mDetail.GetChildNm(strChildCd)
        If strChildNm = "" Then
            MsgBox "입력한 코드는 상세 자코드가 아닙니다.", vbInformation, "정보"
            With txtChild
                .SetFocus
                .Text = ""
            End With
        Else
            lblChildNm.Caption = strChildNm
        End If
    Else
        '## 입력된 코드가 존재하는 경우
        With lvwChild
            .ListItems(itmX.Index).Selected = True
            .ListItems(itmX.Index).EnsureVisible
            Call lvwChild_ItemClick(itmX)
        End With
    End If
    Set itmX = Nothing
End Sub

Private Sub cmdSearch_Click(Index As Integer)
    Select Case Index
        Case 0
            Set mCode1 = New clsIISCodeList
            With mCode1
                .Caption = "상세 모코드 리스트"
                .HeaderCd = "검사코드"
                .HeaderCdNm = "검사명"
                .CodeListBySql mDetail.GetParentListSql
            End With
            Set mCode1 = Nothing
        Case 1
            Set mCode2 = New clsIISCodeList
            With mCode2
                .Caption = "상세 자코드 리스트"
                .HeaderCd = "검사코드"
                .HeaderCdNm = "검사명"
                .CodeListBySql mDetail.GetChildListSql
            End With
            Set mCode2 = Nothing
    End Select
End Sub

Private Sub lvwParent_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static intOrder As Integer

    With lvwParent
        .SortKey = ColumnHeader.Index - 1
        .SortOrder = IIf(intOrder = 0, lvwAscending, lvwDescending)
        .Sorted = True
        intOrder = (intOrder + 1) Mod 2
    End With
End Sub

Private Sub lvwParent_ItemClick(ByVal Item As MSComctlLib.ListItem)
    '## 모코드의 Textbox에 코드, 코드명 표시
    Call CtlClear
    txtParent.Text = Item.Text
    lblParentNm.Caption = Item.SubItems(1)
    
    '## 선택된 모코드의 자코드 리스트 표시
    Call GetChildList
End Sub

Private Sub lvwChild_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static intOrder As Integer

    With lvwChild
        .SortKey = ColumnHeader.Index - 1
        .SortOrder = IIf(intOrder = 0, lvwAscending, lvwDescending)
        .Sorted = True
        intOrder = (intOrder + 1) Mod 2
    End With
End Sub

Private Sub lvwChild_ItemClick(ByVal Item As MSComctlLib.ListItem)
    '## 선택된 자코드의 코드, 코드명을 Textbox에 표시
    txtChild.Text = Item.Text
    lblChildNm.Caption = Item.SubItems(1)
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 모코드 리스트를 lvwParent에 표시
'-----------------------------------------------------------------------------'
Private Sub GetParentList()
    Dim Rs          As ADODB.Recordset
    Dim itmX        As ListItem
    
On Error GoTo Errors
    Set Rs = mDetail.GetParentList
    If Not (Rs.BOF Or Rs.EOF) Then
        With lvwParent
            .ListItems.Clear
            lvwChild.ListItems.Clear
            
            Do Until Rs.EOF
                Set itmX = .ListItems.Add(, , Rs.Fields("TESTCD").Value)
                itmX.SubItems(1) = Rs.Fields("TESTNM").Value
                Rs.MoveNext
            Loop
            
            If .ListItems.Count > 12 Then
                .ColumnHeaders(2).Width = 2210
            Else
                .ColumnHeaders(2).Width = 2410
            End If
        End With
    End If
    Rs.Close
    Set Rs = Nothing
    Set itmX = Nothing
    Exit Sub
    
Errors:
    Set Rs = Nothing
    Set itmX = Nothing
    Error.SetLog App.EXEName, "frmIIS606", "GetParentList", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 자코드 리스트를 lvwParent에 표시
'-----------------------------------------------------------------------------'
Private Sub GetChildList()
    Dim Rs          As ADODB.Recordset
    Dim itmX        As ListItem
    
On Error GoTo Errors
    With lvwChild
        .ListItems.Clear
        
        Set Rs = mDetail.GetChildList(Trim(txtParent.Text))
        If Not (Rs.BOF Or Rs.EOF) Then
            Do Until Rs.EOF
                Set itmX = .ListItems.Add(, , Rs.Fields("TESTCD").Value)
                itmX.SubItems(1) = Rs.Fields("TESTNM").Value
                Rs.MoveNext
            Loop
            
            If .ListItems.Count > 21 Then
                .ColumnHeaders(2).Width = 2210
            Else
                .ColumnHeaders(2).Width = 2410
            End If
        End If
    End With
    Rs.Close
    Set Rs = Nothing
    Set itmX = Nothing
    Exit Sub

Errors:
    Set Rs = Nothing
    Set itmX = Nothing
    Error.SetLog App.EXEName, "frmIIS606", "GetChildList", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 입력된 상세모, 자코드의 유효성 검사
'   반환 : True(유효), False(무효)
'-----------------------------------------------------------------------------'
Private Function CheckCode() As Boolean
    Dim strTemp As String
    
    If txtParent.Text = "" Then
        MsgBox "대표항목 코드를 입력하세요.", vbInformation, "정보"
        txtParent.SetFocus
        Exit Function
    End If
    
    If txtChild.Text = "" Then
        MsgBox "상세항목 코드를 입력하세요.", vbInformation, "정보"
        txtChild.SetFocus
        Exit Function
    End If
    
    CheckCode = True
End Function

'-----------------------------------------------------------------------------'
'   기능 : 컨트롤 초기화
'-----------------------------------------------------------------------------'
Private Sub CtlClear()
    txtParent.Text = "":        txtChild.Text = ""
    lblParentNm.Caption = "":   lblChildNm.Caption = ""
End Sub

'-----------------------------------------------------------------------------'
'   기능 : CodeList폼의 이벤트 처리1
'-----------------------------------------------------------------------------'
Private Sub mCode1_SelectedItem(ByRef pSelItem As String)
    Dim itmX As ListItem
    
    txtParent.Text = mGetP(pSelItem, 1, DIV)
    lblParentNm.Caption = mGetP(pSelItem, 2, DIV)
    
    With lvwParent
        Set itmX = .FindItem(txtParent.Text, lvwText)
        If Not (itmX Is Nothing) Then
            .ListItems(itmX.Index).Selected = True
            .ListItems(itmX.Index).EnsureVisible
            Call lvwParent_ItemClick(itmX)
        End If
        Set itmX = Nothing
    End With
End Sub

'-----------------------------------------------------------------------------'
'   기능 : CodeList폼의 이벤트 처리2
'-----------------------------------------------------------------------------'
Private Sub mCode2_SelectedItem(ByRef pSelItem As String)
    Dim itmX As ListItem
    
    txtChild.Text = mGetP(pSelItem, 1, DIV)
    lblChildNm.Caption = mGetP(pSelItem, 2, DIV)
    
    With lvwChild
        Set itmX = .FindItem(txtChild.Text, lvwText)
        If Not (itmX Is Nothing) Then
            .ListItems(itmX.Index).Selected = True
            .ListItems(itmX.Index).EnsureVisible
            Call lvwChild_ItemClick(itmX)
        End If
        Set itmX = Nothing
    End With
End Sub
