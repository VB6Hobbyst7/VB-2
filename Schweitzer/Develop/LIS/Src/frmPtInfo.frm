VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frmPtInfo 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "환자 조회"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6840
   ClipControls    =   0   'False
   Icon            =   "frmPtInfo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
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
      TabIndex        =   11
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
      TabIndex        =   10
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
         Left            =   3840
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
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   315
         Left            =   4365
         TabIndex        =   6
         Top             =   300
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtMaxCnt"
         BuddyDispid     =   196612
         OrigRight       =   240
         OrigBottom      =   735
         Max             =   999999
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "최대                  건 까지만 검색"
         Height          =   180
         Left            =   3360
         TabIndex        =   7
         Top             =   360
         Width           =   2640
      End
   End
   Begin MedControls1.LisLabel LisLabel3 
      Height          =   375
      Left            =   300
      TabIndex        =   8
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
      TabIndex        =   9
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
Attribute VB_Name = "frmPtInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private blnSearch As Boolean
Public Event Click(ByVal isSELECT As Boolean, ByVal ptInfo As clsPtInformation)

Private Sub cmdExit_Click()
    RaiseEvent Click(False, Nothing)
    Unload Me
End Sub

Private Sub cmdOk_Click()
    RaiseEventClick
    Unload Me
End Sub

Private Sub Form_Activate()
    txtSearchKey.SetFocus
End Sub

Private Sub Form_Load()

    lvwPtList.ListItems.Clear
    optSort(1).Value = True
    
    blnSearch = False '이름검색
    txtSearchKey = ""
    
End Sub

Private Sub lvwPtList_DblClick()
    If lvwPtList.SelectedItem Is Nothing Then Exit Sub
    cmdOk_Click
End Sub

Private Sub lvwPtList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If lvwPtList.SelectedItem Is Nothing Then Exit Sub
        cmdOk_Click
    End If
End Sub

Private Sub optSort_Click(Index As Integer)
    If Index = 0 Then
        blnSearch = True
    Else
        blnSearch = False
        Call medHanOn(txtSearchKey)
    End If
End Sub

Private Sub txtMaxCnt_KeyPress(KeyAscii As Integer)
    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And (KeyAscii <> vbKeyBack) Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub txtSearchKey_GotFocus()
    With txtSearchKey
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtSearchKey_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim objPtInfo As New clsPtInformation
    Dim DrRS As New Recordset
    Dim strOrdDt As String
    Dim iTmx As Object
    Dim objPrg As clsProgressBar
    Dim Cnt As Long, ColCnt As String
    Dim cc As String
    
    On Error GoTo ErrPtInfo
    strOrdDt = Format(GetSystemdate, "yyyymmdd")
'    objPtInfo.setDbConn DBConn
    
    If KeyCode = vbKeyReturn Then

        DrRS.Open objPtInfo.GetPtInfo(txtSearchKey, blnSearch, strOrdDt), dbconn
        
        If DrRS.EOF = False Then
        
            Set objPrg = New clsProgressBar
            Set objPrg.StatusBar = MainFrm.stsBar
            objPrg.Min = 1
            objPrg.Max = Val(txtMaxCnt.Text)
            objPrg.Value = 0
            
            With lvwPtList
                .ListItems.Clear
                
                Cnt = 0
                Do Until DrRS.EOF
                
                    Cnt = Cnt + 1
                    If Cnt > Val(txtMaxCnt.Text) Then Exit Do
                    
                    objPrg.Value = objPrg.Value + 1
                
                    If Mid(DrRS.Fields("ssn").Value & "", 7, 1) = "3" Or _
                       Mid(DrRS.Fields("ssn").Value & "", 7, 1) = "4" Then
                        cc = "20"
                    Else
                        cc = "19"
                    End If
                
                    Set iTmx = .ListItems.Add(, , "" & DrRS.Fields("ptid").Value & "")
                    iTmx.SubItems(1) = "" & DrRS.Fields("ptnm").Value & ""
                    iTmx.SubItems(2) = Mid("" & DrRS.Fields("SSN").Value & "", 1, 6) & "-" & _
                        Mid("" & DrRS.Fields("ssn").Value & "", 7)
                    iTmx.SubItems(3) = cc & Mid("" & DrRS.Fields("SSN").Value & "", 1, 2) & "-" & _
                        Mid("" & DrRS.Fields("SSN").Value & "", 3, 2) & "-" & _
                        Mid("" & DrRS.Fields("SSN").Value & "", 5, 2)
                    iTmx.SubItems(4) = IIf((Val(Mid("" & DrRS.Fields("ssn").Value & "", 7, 1)) Mod 2) = 1, "남", "여")
                    DrRS.MoveNext
                Loop
            End With

            Set objPrg = Nothing
        Else
            MsgBox "조건에 맞는 자료가 없습니다. 확인후 검색하세요", vbInformation + vbOKOnly, Me.Caption
        End If
        Set DrRS = Nothing
    
    End If
    
    Set objPtInfo = Nothing
    Exit Sub

ErrPtInfo:
    MsgBox "검색된 환자수가 너무 많습니다. 조건을 다시 입력하신 후 조회하세요.", vbExclamation, "환자검색"
    Set DrRS = Nothing
End Sub

Private Sub RaiseEventClick()
    '환자ID를 가지고, his002에서 병동환자인지 구별한다.
    '가장 최근의 입원일을 가지고 his002에서 조회한후 퇴원일이
    '있으면, 외래환자 없으면, 병동환자이다.
    Dim PtId As String
    Dim ptInfo As clsPtInformation
    
    If lvwPtList.ListItems.Count <= 0 Then
        RaiseEvent Click(False, Nothing)
    Else
        PtId = lvwPtList.SelectedItem.Text
        Set ptInfo = New clsPtInformation
'        ptInfo.setDbConn dbconn
    
        Call ptInfo.BedPt_Chk(PtId, Format(GetSystemdate, "yyyymmdd"))
        
        RaiseEvent Click(True, ptInfo)
    End If
    Set ptInfo = Nothing
End Sub

