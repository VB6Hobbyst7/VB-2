VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frmSearchPt 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "환자 조회"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6840
   ClipControls    =   0   'False
   Icon            =   "frmSearchPt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00EBF3ED&
      Caption         =   "확인(&O)"
      Height          =   510
      Left            =   2070
      Style           =   1  '그래픽
      TabIndex        =   10
      Tag             =   "0"
      Top             =   7335
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00EBF3ED&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   3420
      Style           =   1  '그래픽
      TabIndex        =   9
      Tag             =   "0"
      Top             =   7320
      Width           =   1320
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   375
      Left            =   300
      TabIndex        =   0
      Top             =   240
      Width           =   6300
      _ExtentX        =   11113
      _ExtentY        =   661
      BackColor       =   8388608
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "  검색 조건"
      Appearance      =   0
   End
   Begin VB.Frame fraSearch 
      BackColor       =   &H00DBE6E6&
      Height          =   810
      Left            =   300
      TabIndex        =   1
      Tag             =   "136"
      Top             =   540
      Width           =   6315
      Begin VB.TextBox txtMaxCnt 
         Height          =   315
         Left            =   3945
         MaxLength       =   4
         TabIndex        =   5
         Text            =   "1000"
         Top             =   300
         Width           =   510
      End
      Begin VB.OptionButton optSort 
         BackColor       =   &H00DBE6E6&
         Caption         =   "&ID"
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Tag             =   "15304"
         Top             =   345
         Width           =   495
      End
      Begin VB.OptionButton optSort 
         BackColor       =   &H00DBE6E6&
         Caption         =   "&Name"
         Height          =   255
         Index           =   1
         Left            =   705
         TabIndex        =   3
         Tag             =   "15305"
         Top             =   330
         Value           =   -1  'True
         Width           =   810
      End
      Begin VB.TextBox txtSearchKey 
         Height          =   300
         Left            =   1590
         MaxLength       =   10
         TabIndex        =   2
         Text            =   "테"
         Top             =   300
         Width           =   1470
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "최대             건 까지만 검색"
         Height          =   180
         Left            =   3420
         TabIndex        =   6
         Top             =   360
         Width           =   2340
      End
   End
   Begin MedControls1.LisLabel LisLabel3 
      Height          =   375
      Left            =   300
      TabIndex        =   7
      Top             =   1365
      Width           =   6300
      _ExtentX        =   11113
      _ExtentY        =   661
      BackColor       =   8388608
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "  검색 결과"
      Appearance      =   0
   End
   Begin MSComctlLib.ListView lvwPtList 
      Height          =   5475
      Left            =   300
      TabIndex        =   8
      Top             =   1740
      Width           =   6300
      _ExtentX        =   11113
      _ExtentY        =   9657
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "환자ID"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "환자명"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "주민번호"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "생년월일"
         Object.Width           =   2295
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "성별"
         Object.Width           =   1058
      EndProperty
   End
End
Attribute VB_Name = "frmSearchPt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Public Event Selected(ByVal vPtInfo As clsPatient)
Public Event SelectedId(ByVal vPtID As String)

'Private objPtInfo As clsPatient

Private Sub cmdExit_Click()
'    RaiseEvent Selected(Nothing)
    RaiseEvent SelectedId("")
End Sub

Private Sub cmdOk_Click()
    Call lvwPtList_DblClick
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    txtSearchKey.SetFocus
End Sub

Private Sub Form_Load()
    lvwPtList.ListItems.Clear
    optSort(1).Value = True
    txtSearchKey.Text = ""
End Sub

Private Sub lvwPtList_DblClick()
    If lvwPtList.SelectedItem Is Nothing Then Exit Sub
    
    RaiseEvent SelectedId(lvwPtList.SelectedItem.Text)
'    Set objPtInfo = New clsPatient
'
'    If objPtInfo.GETPatient(lvwPtList.SelectedItem.Text) Then
'        RaiseEvent Selected(objPtInfo)
'    Else
'        RaiseEvent Selected(Nothing)
'    End If
'
'    Set objPtInfo = Nothing
'    Unload Me
End Sub

Private Sub lvwPtList_KeyDown(KeyCode As Integer, Shift As Integer)
    Call lvwPtList_DblClick
End Sub

Private Sub txtMaxCnt_KeyPress(KeyAscii As Integer)
    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And (KeyAscii <> vbKeyBack) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub txtSearchKey_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtSearchKey_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtSearchKey.Text = "" Then Exit Sub
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtSearchKey_LostFocus()
    Dim objPtInfo As clsPatient
    Dim Rs As Recordset
    Dim itmX As ListItem
    Dim strSQL As String
    
    If txtSearchKey.Text = "" Then Exit Sub
    If (optSort(1).Value) And (Len(Trim(txtSearchKey.Text)) < 2) Then
        MsgBox "환자명으로 검색은 두자 이상을 입력해야 합니다.", vbExclamation
        Exit Sub
    End If
    
    Set objPtInfo = New clsPatient
    Set Rs = New Recordset
    
    strSQL = objPtInfo.GetSQLPtNt(IIf(optSort(0).Value, "1", "2"), txtSearchKey.Text)
    
    Screen.MousePointer = vbHourglass
    DoEvents
    
    Rs.Open strSQL, DBConn
    
    If Rs.EOF Then
        MsgBox "조건에 해당하는 환자가 없습니다.", vbExclamation
        GoTo NoData
    End If
    
    lvwPtList.ListItems.Clear
    Do Until Rs.EOF

        If lvwPtList.ListItems.Count > Val(txtMaxCnt.Text) Then Exit Do

        '환자아디,환자명,ssn,dob,sex
        Set itmX = lvwPtList.ListItems.Add()
            itmX.Text = Rs.Fields("ptid").Value & ""
            itmX.SubItems(1) = Rs.Fields("ptnm").Value & ""
            itmX.SubItems(2) = Rs.Fields("ssn").Value & ""
            itmX.SubItems(3) = Rs.Fields("dob").Value & ""
            itmX.SubItems(4) = IIf((Val(Mid(Rs.Fields("ssn").Value & "", 7, 1)) Mod 2) = 1, "남", "여")
        Rs.MoveNext
    Loop

NoData:
    Screen.MousePointer = vbDefault
    Set Rs = Nothing
    Set objPtInfo = Nothing
End Sub
