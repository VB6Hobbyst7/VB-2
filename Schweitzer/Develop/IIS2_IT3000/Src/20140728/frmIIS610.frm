VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frmIIS610 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   4  '고정 도구 창
   Caption         =   "검사장비 통신설정"
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
   Begin MSComctlLib.ListView lvwEqpList 
      Height          =   4575
      Left            =   45
      TabIndex        =   11
      Top             =   450
      Width           =   11070
      _ExtentX        =   19526
      _ExtentY        =   8070
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
      NumItems        =   13
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "장비코드"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "장비명"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Port"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Baud Rate"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Data bit"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Stop bit"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Parity bit"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "보관일"
         Object.Width           =   1693
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "비고"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "온도구분"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Low"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "High"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "사용유무"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00DBE6E6&
      Caption         =   "삭 제(&D)"
      Height          =   495
      Left            =   7470
      Style           =   1  '그래픽
      TabIndex        =   8
      Top             =   8205
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00DBE6E6&
      Caption         =   "저 장(&S)"
      Height          =   495
      Left            =   6255
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
      TabIndex        =   9
      Top             =   8205
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   3165
      Left            =   60
      TabIndex        =   12
      Top             =   4950
      Width           =   11070
      Begin VB.ComboBox cboParity 
         BackColor       =   &H00F7FFF7&
         Height          =   300
         ItemData        =   "frmIIS610.frx":0000
         Left            =   2220
         List            =   "frmIIS610.frx":0002
         Style           =   2  '드롭다운 목록
         TabIndex        =   5
         Top             =   2430
         Width           =   2370
      End
      Begin VB.ComboBox cboStopbit 
         BackColor       =   &H00F7FFF7&
         Height          =   300
         Left            =   8070
         Style           =   2  '드롭다운 목록
         TabIndex        =   4
         Top             =   1875
         Width           =   2370
      End
      Begin VB.ComboBox cboDatabit 
         BackColor       =   &H00F7FFF7&
         Height          =   300
         Left            =   2220
         Style           =   2  '드롭다운 목록
         TabIndex        =   3
         Top             =   1875
         Width           =   2370
      End
      Begin VB.ComboBox cboBaud 
         BackColor       =   &H00F7FFF7&
         Height          =   300
         Left            =   8070
         Style           =   2  '드롭다운 목록
         TabIndex        =   2
         Top             =   1320
         Width           =   2370
      End
      Begin VB.ComboBox cboPort 
         BackColor       =   &H00F7FFF7&
         Height          =   300
         Left            =   2220
         Style           =   2  '드롭다운 목록
         TabIndex        =   1
         Top             =   1320
         Width           =   2370
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00DBE6E6&
         Height          =   330
         Left            =   3510
         Picture         =   "frmIIS610.frx":0004
         Style           =   1  '그래픽
         TabIndex        =   18
         Top             =   375
         Width           =   405
      End
      Begin VB.TextBox txtClient 
         BackColor       =   &H00F7FFF7&
         Height          =   330
         Left            =   8070
         MaxLength       =   3
         TabIndex        =   6
         Top             =   2430
         Width           =   2370
      End
      Begin VB.TextBox txtEqpCd 
         BackColor       =   &H00F7FFF7&
         Height          =   330
         Left            =   2220
         MaxLength       =   8
         TabIndex        =   0
         Top             =   390
         Width           =   1275
      End
      Begin MedControls1.LisLabel lblEqpNm 
         Height          =   345
         Left            =   4050
         TabIndex        =   19
         Top             =   375
         Width           =   2955
         _ExtentX        =   5212
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
         Caption         =   "Hitachi 7600"
         LeftGab         =   100
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "▶ ClientDb 보관일"
         Height          =   180
         Left            =   5745
         TabIndex        =   22
         Top             =   2445
         Width           =   1545
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "▶ Stop bit"
         Height          =   180
         Left            =   5745
         TabIndex        =   21
         Top             =   1920
         Width           =   870
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "▶ Baud Rate"
         Height          =   180
         Left            =   5745
         TabIndex        =   20
         Top             =   1380
         Width           =   1110
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "▶ Port"
         Height          =   180
         Left            =   255
         TabIndex        =   16
         Top             =   1380
         Width           =   570
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "▶ Data bit"
         Height          =   180
         Left            =   255
         TabIndex        =   15
         Top             =   1920
         Width           =   870
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "▶ Parity bit"
         Height          =   180
         Left            =   255
         TabIndex        =   14
         Top             =   2445
         Width           =   975
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000011&
         X1              =   75
         X2              =   10900
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "▶ 장비코드"
         Height          =   180
         Left            =   255
         TabIndex        =   13
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
      TabIndex        =   10
      Top             =   8205
      Width           =   1215
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00DBE6E6&
      Caption         =   "    인터페이스 장애가 발생할 수 있습니다."
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
      TabIndex        =   24
      Top             =   8460
      Width           =   6075
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00DBE6E6&
      Caption         =   "※ 장비 설정을 임의로 변경하면"
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
      TabIndex        =   23
      Top             =   8220
      Width           =   4455
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblName 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "검사장비 리스트"
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
      TabIndex        =   17
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
Attribute VB_Name = "frmIIS610"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------'
'   파일명  : frmIIS610.frm (우리LIS랑 조인할때 사용)
'   작성자  :
'   내  용  : 검사장비 통신설정 마스터
'   작성일  : 2004-03-03
'   버  전  :
'       1. 1.2.5:  (2005-07-20)
'-----------------------------------------------------------------------------'

Option Explicit

Private mEqpComm            As clsIISEqpComm        '장비 통신설정 클래스
Private WithEvents mCode    As clsIISCodeList       '코드리스트 클래스
Attribute mCode.VB_VarHelpID = -1

Private Sub Form_Load()
    With Me
        .Top = 0: .Left = 4030
        .Height = mdiIISMain.ScaleHeight
        
        '## 1.2.5:  (2005-07-20)
        '   - 모니터의 해상도가 변해도 항상 폼의 ScaleHeight에 맞도록 수정
        .Width = mdiIISMain.ScaleWidth - 4030
    End With

    Set mEqpComm = New clsIISEqpComm
    Call InitCombo
    Call CtlClear
    Me.Show
    DoEvents
    
    '## 장비 통신설정 정보표시
    Me.MousePointer = vbHourglass
    Call GetEqpComms
    Me.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
    mdiIISMain.lblMenuNm = Me.Caption
    frmIIS600.tvwMenu.Nodes("IIS610").Selected = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mEqpComm = Nothing
    Set frmIIS610 = Nothing
End Sub

Private Sub cmdSave_Click()
    Dim itmX        As ListItem
    Dim strEqpCd    As String       '장비코드

    '## 유효성 Check
    If CheckCode = False Then Exit Sub

    strEqpCd = Trim(txtEqpCd.Text)
    Set itmX = lvwEqpList.FindItem(strEqpCd, lvwText)

    Me.MousePointer = vbHourglass
    
    With mEqpComm
        .EqpCd = strEqpCd
        .Port = cboPort.Text
        .BaudRate = cboBaud.Text
        .Databit = cboDatabit.Text
        .Stopbit = cboStopbit.Text
        .Paritybit = cboParity.Text
        .StoredDt = Trim(txtClient.Text)
        
        '## 존재하는 장비코드이면 Update 존재하지 않으면 Insert
        If itmX Is Nothing Then
            If .AddEqpComm Then
                mdiIISMain.sbrStatus.Panels(2).Text = "정상적으로 저장되었습니다."
            Else
                mdiIISMain.sbrStatus.Panels(2).Text = "저장중에 에러가 발생했습니다."
            End If
        Else
            If .ModifyEqpComm Then
                mdiIISMain.sbrStatus.Panels(2).Text = "정상적으로 수정되었습니다."
            Else
                mdiIISMain.sbrStatus.Panels(2).Text = "수정중에 에러가 발생했습니다."
            End If
        End If
    End With
    Call CtlClear
    Call GetEqpComms
    
    Set itmX = lvwEqpList.FindItem(strEqpCd, lvwText)
    If Not (itmX Is Nothing) Then
        lvwEqpList.ListItems(itmX.Index).Selected = True
        lvwEqpList.ListItems(itmX.Index).EnsureVisible
    End If
    Set itmX = Nothing
    txtEqpCd.SetFocus

    Me.MousePointer = vbDefault
End Sub

Private Sub cmdDelete_Click()
    Dim itmX        As ListItem
    Dim strEqpCd    As String          '장비코드
    Dim intTemp     As Integer

    strEqpCd = Trim(txtEqpCd.Text)
    If strEqpCd = "" Then
        MsgBox "장비코드를 입력하세요.", vbInformation, "정보"
        Exit Sub
    End If

    intTemp = MsgBox("정말 삭제할까요?", vbYesNo + vbQuestion, "확인")
    If intTemp = vbNo Then Exit Sub

    Set itmX = lvwEqpList.FindItem(strEqpCd, lvwText)
    If itmX Is Nothing Then
        MsgBox "존재하지 않는 장비코드 입니다.", vbInformation, "정보"
        Exit Sub
    End If
    Set itmX = Nothing

    Me.MousePointer = vbHourglass

    With mEqpComm
        .EqpCd = strEqpCd
        If .DelEqpComm Then
            mdiIISMain.sbrStatus.Panels(2).Text = "정상적으로 삭제되었습니다."
        Else
            mdiIISMain.sbrStatus.Panels(2).Text = "삭제중에 에러가 발생했습니다."
        End If
    End With

    Call CtlClear
    Call GetEqpComms
    txtEqpCd.SetFocus

    Me.MousePointer = vbDefault
End Sub

Private Sub cmdClear_Click()
    Call CtlClear
    txtEqpCd.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSearch_Click()
    Set mCode = New clsIISCodeList
    With mCode
        .Caption = "장비리스트"
        .HeaderCd = "장비코드"
        .HeaderCdNm = "장비명"
        .CodeListByRs mEqpComm.GetUsingEqp
    End With
    Set mCode = Nothing
End Sub

Private Sub lvwEqpList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static intOrder As Integer

    With lvwEqpList
        .SortKey = ColumnHeader.Index - 1
        .SortOrder = IIf(intOrder = 0, lvwAscending, lvwDescending)
        .Sorted = True
        intOrder = (intOrder + 1) Mod 2
    End With
End Sub

Private Sub lvwEqpList_ItemClick(ByVal Item As MSComctlLib.ListItem)
    '## 장비코드에 대한 상세정보를 표시
    Call CtlClear
    
    With Item
        txtEqpCd.Text = .Text
        lblEqpNm.Caption = .SubItems(1)
        cboPort.Text = .SubItems(2)
        cboBaud.Text = .SubItems(3)
        cboDatabit.Text = .SubItems(4)
        cboStopbit.Text = .SubItems(5)
        cboParity.Text = .SubItems(6)
        txtClient.Text = .SubItems(7)
    End With
End Sub

Private Sub txtEqpCd_GotFocus()
    With txtEqpCd
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtEqpCd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txtEqpCd.Text = "" Then
            MsgBox "장비코드를 입력하세요.", vbInformation, "정보"
            Exit Sub
        End If
        
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
    Dim itmX     As ListItem
    Dim strEqpCd As String      '장비코드
    Dim strEqpNm As String      '장비명
    
    strEqpCd = Trim(txtEqpCd.Text)
    If strEqpCd = "" Then Exit Sub
    
    '## 존재하는 장비코드이면 포커스이동, 정보표시
    '## 존재하지 않는 장비코드이면 등록된 장비인지 검사
    Set itmX = lvwEqpList.FindItem(strEqpCd, lvwText)
    If Not (itmX Is Nothing) Then
        lvwEqpList.ListItems(itmX.Index).Selected = True
        lvwEqpList.ListItems(itmX.Index).EnsureVisible
        Call lvwEqpList_ItemClick(itmX)
    Else
        Call CtlClear
        txtEqpCd.Text = strEqpCd
        
        strEqpNm = mEqpComm.GetEqpNm(strEqpCd)
        If strEqpNm = "" Then
            MsgBox "등록된 장비코드가 아닙니다.", vbInformation, "정보"
            With txtEqpCd
                .Text = ""
                .SetFocus
            End With
        Else
            lblEqpNm.Caption = strEqpNm
        End If
    End If
    Set itmX = Nothing
End Sub

Private Sub txtClient_GotFocus()
    With txtClient
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtClient_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtClient_KeyPress(KeyAscii As Integer)
    '## 숫자, Backspace키만 입력할수 있도록함
    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 장비 통신설정 정보를 lvwEqpList에 표시
'-----------------------------------------------------------------------------'
Private Sub GetEqpComms()
    Dim Rs      As ADODB.Recordset
    Dim itmX    As ListItem

On Error GoTo Errors
    With lvwEqpList
        .ListItems.Clear

        Set Rs = mEqpComm.GetEqpComms
        If Rs.BOF Or Rs.EOF Then GoTo EndLine
    
        Do Until Rs.EOF
            Set itmX = .ListItems.Add(, , Rs.Fields("EQPCD").Value)
            itmX.SubItems(1) = Rs.Fields("EQPNM").Value
            itmX.SubItems(2) = Rs.Fields("PORT").Value & ""
            itmX.SubItems(3) = Rs.Fields("BAUDRATE").Value & ""
            itmX.SubItems(4) = Rs.Fields("DATABIT").Value & ""
            itmX.SubItems(5) = Rs.Fields("STOPBIT").Value & ""
            itmX.SubItems(6) = Rs.Fields("PARITYBIT").Value & ""
            itmX.SubItems(7) = Rs.Fields("STOREDDT").Value & ""
            Rs.MoveNext
        Loop
        Set itmX = Nothing

        If .ListItems.Count > 21 Then
            .ColumnHeaders(2).Width = 3250
        Else
            .ColumnHeaders(2).Width = 3500
        End If
    End With

EndLine:
    Rs.Close
    Set Rs = Nothing
    Exit Sub

Errors:
    Set Rs = Nothing
    Set itmX = Nothing
    Error.SetLog App.EXEName, "frmIIS610", "GetEqpComms", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 입력, 수정시 필요한 정보의 유효성 Check
'   반환 : True(유효), False(무효)
'-----------------------------------------------------------------------------'
Private Function CheckCode() As Boolean
    '## 장비코드
    If txtEqpCd.Text = "" Then
        MsgBox "장비코드를 입력하세요.", vbInformation, "정보"
        txtEqpCd.SetFocus
        Exit Function
    End If
    
    '## Port
    If cboPort.Text = "" Then
        MsgBox "Port를 선택하세요.", vbInformation, "정보"
        cboPort.SetFocus
        Exit Function
    End If
    
    '## Baud Rate
    If cboBaud.Text = "" Then
        MsgBox "Baud Rate를 선택하세요.", vbInformation, "정보"
        cboBaud.SetFocus
        Exit Function
    End If
    
    '## Data bit
    If cboDatabit.Text = "" Then
        MsgBox "Data bit를 선택하세요.", vbInformation, "정보"
        cboDatabit.SetFocus
        Exit Function
    End If
    
    '## Stop bit
    If cboStopbit.Text = "" Then
        MsgBox "Stop bit를 선택하세요.", vbInformation, "정보"
        cboStopbit.SetFocus
        Exit Function
    End If
    
    '## Parity bit
    If cboParity.Text = "" Then
        MsgBox "Parity bit를 선탁하세요.", vbInformation, "정보"
        cboParity.SetFocus
        Exit Function
    End If
    
    '## ClientDb보관일수
    If txtClient.Text = "" Then
        MsgBox "ClientDb 보관일수를 선택하세요.", vbInformation, "정보"
        txtClient.SetFocus
        Exit Function
    End If
    
    CheckCode = True
End Function

'-----------------------------------------------------------------------------'
'   기능 : Combobox 초기화
'-----------------------------------------------------------------------------'
Private Sub InitCombo()
    '## Port
    With cboPort
        .Clear
        .AddItem "1"
        .AddItem "2"
        .AddItem "3"
        .AddItem "4"
    End With
    
    '## Baud Rate
    With cboBaud
        .Clear
        .AddItem "300"
        .AddItem "600"
        .AddItem "1200"
        .AddItem "2400"
        .AddItem "4800"
        .AddItem "9600"
        .AddItem "14400"
        .AddItem "19200"
        .AddItem "28800"
    End With
    
    '## Data bit
    With cboDatabit
        .Clear
        .AddItem "7"
        .AddItem "8"
    End With
    
    '## Stop bit
    With cboStopbit
        .Clear
        .AddItem "1"
        .AddItem "2"
    End With
    
    '## Parity bit
    With cboParity
        .Clear
        .AddItem "None"
        .AddItem "Even"
        .AddItem "Odd"
    End With
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 컨트롤 Clear
'-----------------------------------------------------------------------------'
Private Sub CtlClear()
    txtEqpCd.Text = "":         lblEqpNm.Caption = ""
    cboPort.ListIndex = -1:     cboBaud.ListIndex = -1
    cboDatabit.ListIndex = -1:  cboStopbit.ListIndex = -1
    cboParity.ListIndex = -1:   txtClient.Text = ""
End Sub

'-----------------------------------------------------------------------------'
'   기능 : CodeList폼의 이벤트 처리
'-----------------------------------------------------------------------------'
Private Sub mCode_SelectedItem(ByRef pSelItem As String)
    Dim itmX As ListItem
    
    Call CtlClear
    txtEqpCd.Text = mGetP(pSelItem, 1, DIV)
    lblEqpNm.Caption = mGetP(pSelItem, 2, DIV)
    
    With lvwEqpList
        Set itmX = .FindItem(txtEqpCd.Text, lvwText)
        If Not (itmX Is Nothing) Then
            .ListItems(itmX.Index).Selected = True
            .ListItems(itmX.Index).EnsureVisible
            Call lvwEqpList_ItemClick(itmX)
        End If
        Set itmX = Nothing
    End With
End Sub
