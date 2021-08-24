VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmIIS601 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   4  '고정 도구 창
   Caption         =   "검체코드 관리"
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
      TabIndex        =   5
      Top             =   8205
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00DBE6E6&
      Caption         =   "저 장(&S)"
      Height          =   495
      Left            =   6255
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
      TabIndex        =   6
      Top             =   8205
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   8145
      Left            =   3585
      TabIndex        =   8
      Top             =   -30
      Width           =   7545
      Begin VB.TextBox txtSpcFullNm 
         BackColor       =   &H00F7FFF7&
         Height          =   330
         Left            =   255
         MaxLength       =   40
         TabIndex        =   2
         Top             =   2910
         Width           =   3750
      End
      Begin VB.TextBox txtSpcBarNm 
         BackColor       =   &H00F7FFF7&
         Height          =   330
         Left            =   255
         MaxLength       =   10
         TabIndex        =   3
         Top             =   3690
         Width           =   3750
      End
      Begin VB.TextBox txtSpcNm 
         BackColor       =   &H00F7FFF7&
         Height          =   330
         Left            =   255
         MaxLength       =   20
         TabIndex        =   1
         Top             =   2115
         Width           =   3750
      End
      Begin VB.TextBox txtSpcCd 
         BackColor       =   &H00F7FFF7&
         Height          =   330
         Left            =   255
         MaxLength       =   2
         TabIndex        =   0
         Top             =   780
         Width           =   2505
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "▶ 검체명(약어)"
         Height          =   180
         Left            =   255
         TabIndex        =   13
         Top             =   1800
         Width           =   1290
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "▶ 검체명(전체)"
         Height          =   180
         Left            =   255
         TabIndex        =   12
         Top             =   2595
         Width           =   1290
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "▶ 검체명(Barcode)"
         Height          =   180
         Left            =   255
         TabIndex        =   11
         Top             =   3375
         Width           =   1635
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
         Caption         =   "▶ 검체코드"
         Height          =   180
         Left            =   255
         TabIndex        =   10
         Top             =   480
         Width           =   960
      End
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
   Begin MSComctlLib.ListView lvwSpcList 
      Height          =   7665
      Left            =   45
      TabIndex        =   9
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
         Text            =   "코드"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "코드명"
         Object.Width           =   4251
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "검체명(전체)"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "검체명(Barcode)"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label lblName 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "검체코드 리스트"
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
      TabIndex        =   14
      Top             =   165
      Width           =   1455
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
Attribute VB_Name = "frmIIS601"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------'
'   파일명  : frmIIS601.frm
'   작성자  : 이상대
'   내  용  : 검체코드 마스터
'   작성일  : 2004-01-02
'   버  전  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mSpc As clsIISSpc       '검체코드 클래스

Private Sub Form_Load()
    With Me
        .Top = 0: .Left = 4030
        .Height = mdiIISMain.ScaleHeight: .Width = 11270
    End With
    
    Set mSpc = New clsIISSpc
    Call CtlClear
    Me.Show
    DoEvents
    
    '## 검체코드 로드
    Me.MousePointer = vbHourglass
    Call GetSpcList
    Me.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
    mdiIISMain.lblMenuNm = Me.Caption
    frmIIS600.tvwMenu.Nodes("IIS601").Selected = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mSpc = Nothing
    Set frmIIS601 = Nothing
End Sub

Private Sub cmdSave_Click()
    Dim itmX        As ListItem
    Dim strSpcCd    As String       '검체코드
    Dim strSpcNm    As String       '검체명

    '## 검체코드
    If txtSpcCd.Text = "" Then
        MsgBox "검체코드를 입력하세요.", vbInformation, "오류"
        txtSpcCd.SetFocus
        Exit Sub
    End If
    
    '## 검체명(약어)
    If txtSpcNm.Text = "" Then
        MsgBox "검체명(약어)을 입력하세요.", vbInformation, "오류"
        txtSpcCd.SetFocus
        Exit Sub
    End If
    
    strSpcCd = Trim(txtSpcCd.Text)
    strSpcNm = Trim(txtSpcNm.Text)
    
    Me.MousePointer = vbHourglass
    
    Set itmX = lvwSpcList.FindItem(strSpcCd, lvwText)
    With mSpc
        .SpcCd = strSpcCd
        .SpcNm = strSpcNm
        .SpcFullNm = Trim(txtSpcFullNm.Text)
        .SpcBarNm = Trim(txtSpcBarNm.Text)

        '# 검체코드가 있으면 Update, 없으면 Insert
        If itmX Is Nothing Then
            If .AddSpcCd Then
                mdiIISMain.sbrStatus.Panels(2).Text = "정상적으로 저장되었습니다."
            Else
                mdiIISMain.sbrStatus.Panels(2).Text = "저장중에 에러가 발생했습니다."
            End If
        Else
            If .ModifySpcCd Then
                mdiIISMain.sbrStatus.Panels(2).Text = "정상적으로 수정되었습니다."
            Else
                mdiIISMain.sbrStatus.Panels(2).Text = "수정중에 에러가 발생했습니다."
            End If
        End If
    End With
    
    Call CtlClear
    Call GetSpcList
        
    With lvwSpcList
        Set itmX = .FindItem(strSpcCd)
        If Not (itmX Is Nothing) Then
            .ListItems(itmX.Index).Selected = True
            .ListItems(itmX.Index).EnsureVisible
        End If
        Set itmX = Nothing
    End With
    
    Me.MousePointer = vbDefault
End Sub

Private Sub cmdDelete_Click()
    Dim itmX        As ListItem
    Dim strSpcCd    As String       '검체코드
    Dim intTemp     As Integer

    If txtSpcCd.Text = "" Then
        MsgBox "검체코드를 입력하세요.", vbInformation, "오류"
        txtSpcCd.SetFocus
        Exit Sub
    End If
    
    strSpcCd = Trim(txtSpcCd.Text)
    
    intTemp = MsgBox("정말 삭제할까요?", vbYesNo + vbQuestion, "확인")
    If intTemp = vbNo Then Exit Sub

    Me.MousePointer = vbHourglass

    '## 검체삭제
    Set itmX = lvwSpcList.FindItem(strSpcCd, lvwText)
    If itmX Is Nothing Then Exit Sub
    Set itmX = Nothing

    With mSpc
        .SpcCd = strSpcCd
        If .DelSpcCd Then
            mdiIISMain.sbrStatus.Panels(2).Text = "정상적으로 삭제되었습니다."
        Else
            mdiIISMain.sbrStatus.Panels(2).Text = "삭제중에 에러가 발생했습니다."
        End If
    End With
    
    Call CtlClear
    Call GetSpcList
    
    Me.MousePointer = vbDefault
End Sub

Private Sub cmdClear_Click()
    Call CtlClear
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub lvwSpcList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static intOrder As Integer
    
    With lvwSpcList
        .SortKey = ColumnHeader.Index - 1
        .SortOrder = IIf(intOrder = 0, lvwAscending, lvwDescending)
        .Sorted = True
        intOrder = (intOrder + 1) Mod 2
    End With
End Sub

Private Sub lvwSpcList_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Call CtlClear
    txtSpcCd.Text = Item.Text
    txtSpcNm.Text = Item.SubItems(1)
    txtSpcFullNm.Text = Item.SubItems(2)
    txtSpcBarNm.Text = Item.SubItems(3)
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
    If KeyAscii >= 97 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    End If
End Sub

Private Sub txtSpcCd_LostFocus()
    Dim itmX     As ListItem
    Dim strSpcCd As String      '검체코드
    
    strSpcCd = Trim(txtSpcCd.Text)
    If strSpcCd = "" Then Exit Sub
    
    With lvwSpcList
        Set itmX = .FindItem(strSpcCd, lvwText)
        If itmX Is Nothing Then
            Call CtlClear
            txtSpcCd.Text = strSpcCd
        Else
            .ListItems(itmX.Index).Selected = True
            .ListItems(itmX.Index).EnsureVisible
            Call lvwSpcList_ItemClick(itmX)
        End If
        Set itmX = Nothing
    End With
End Sub

Private Sub txtSpcNm_GotFocus()
    With txtSpcNm
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtSpcNm_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtSpcFullNm_GotFocus()
    With txtSpcFullNm
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtSpcFullNm_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtSpcBarNm_GotFocus()
    With txtSpcBarNm
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtSpcBarNm_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 검체리스트 표시
'-----------------------------------------------------------------------------'
Private Sub GetSpcList()
    Dim Rs   As ADODB.Recordset
    Dim itmX As ListItem

On Error GoTo Errors:
    With lvwSpcList
        .ListItems.Clear

        Set Rs = mSpc.GetSpcCd
        If Rs.BOF Or Rs.EOF Then GoTo EndLine

        Do Until Rs.EOF
            Set itmX = .ListItems.Add(, , Rs.Fields("SPCCD").Value)
            itmX.SubItems(1) = Rs.Fields("SPCNM").Value
            itmX.SubItems(2) = Rs.Fields("SPCFULLNM").Value
            itmX.SubItems(3) = Rs.Fields("SPCBARNM").Value

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
    Error.SetLog App.EXEName, "frmIIS601", "GetSpcList", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 컨트롤 초기화
'-----------------------------------------------------------------------------'
Private Sub CtlClear()
    txtSpcCd.Text = ""
    txtSpcNm.Text = ""
    txtSpcFullNm.Text = ""
    txtSpcBarNm.Text = ""
End Sub
