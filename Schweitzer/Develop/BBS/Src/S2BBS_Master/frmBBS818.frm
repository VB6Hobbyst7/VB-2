VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBBS818 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "문진내역 마스터관리"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10845
   Icon            =   "frmBBS818.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   10845
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "저장(&S)"
      Height          =   420
      Left            =   4860
      Style           =   1  '그래픽
      TabIndex        =   4
      Top             =   7920
      Width           =   1260
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   420
      Left            =   6180
      Style           =   1  '그래픽
      TabIndex        =   3
      Top             =   7920
      Width           =   1260
   End
   Begin VB.CommandButton cmdRefresh 
      BackColor       =   &H00F4F0F2&
      Caption         =   "&Refresh"
      Height          =   420
      Left            =   3540
      Style           =   1  '그래픽
      TabIndex        =   2
      Top             =   7920
      Width           =   1260
   End
   Begin VB.CheckBox chkWide 
      BackColor       =   &H00DBE6E6&
      Caption         =   "한 줄로 보기"
      Height          =   195
      Left            =   9000
      TabIndex        =   1
      Top             =   60
      Width           =   1455
   End
   Begin FPSpread.vaSpread tblList 
      Height          =   4350
      Left            =   120
      TabIndex        =   0
      Top             =   300
      Width           =   10575
      _Version        =   196608
      _ExtentX        =   18653
      _ExtentY        =   7673
      _StockProps     =   64
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   14411494
      GridShowVert    =   0   'False
      MaxCols         =   5
      MaxRows         =   50
      OperationMode   =   2
      RestrictRows    =   -1  'True
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      ShadowText      =   0
      SpreadDesigner  =   "frmBBS818.frx":076A
      ScrollBarTrack  =   3
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   3195
      Left            =   120
      TabIndex        =   5
      Top             =   4560
      Width           =   10590
      Begin VB.TextBox txtPrtSeq 
         Height          =   315
         Left            =   8100
         MaxLength       =   3
         TabIndex        =   16
         Text            =   "AAAAAAAAAA"
         Top             =   480
         Width           =   750
      End
      Begin VB.CheckBox chkDelete 
         BackColor       =   &H00DBE6E6&
         Caption         =   "폐기여부"
         Height          =   225
         Left            =   480
         TabIndex        =   10
         Top             =   2655
         Width           =   1110
      End
      Begin VB.TextBox txtQuestCode 
         Height          =   315
         Left            =   1380
         MaxLength       =   3
         TabIndex        =   9
         Text            =   "AAAAAAAAAA"
         Top             =   480
         Width           =   750
      End
      Begin VB.TextBox txtQuest 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1710
         Left            =   1380
         MaxLength       =   500
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   8
         Text            =   "frmBBS818.frx":0D71
         Top             =   810
         Width           =   7770
      End
      Begin VB.OptionButton optQuest 
         BackColor       =   &H00DBE6E6&
         Caption         =   "예"
         Height          =   240
         Index           =   0
         Left            =   4620
         TabIndex        =   7
         Top             =   495
         Width           =   495
      End
      Begin VB.OptionButton optQuest 
         BackColor       =   &H00DBE6E6&
         Caption         =   "아니오"
         Height          =   255
         Index           =   1
         Left            =   5220
         TabIndex        =   6
         Top             =   495
         Width           =   855
      End
      Begin MedControls1.LisLabel lblDelete 
         Height          =   315
         Left            =   1830
         TabIndex        =   11
         Top             =   2595
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         BackColor       =   14411494
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
         Alignment       =   1
         Caption         =   "폐  기  일  자 :"
         Appearance      =   0
      End
      Begin MSComCtl2.DTPicker dtpDelete 
         Height          =   315
         Left            =   3195
         TabIndex        =   12
         Top             =   2595
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyy-MM-dd"
         Format          =   59637763
         CurrentDate     =   36847
      End
      Begin MedControls1.LisLabel LisLabel2 
         Height          =   195
         Left            =   3780
         TabIndex        =   13
         Top             =   495
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   344
         BackColor       =   14411494
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
         Caption         =   "정상값 :"
         Appearance      =   0
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   315
         Left            =   8850
         TabIndex        =   18
         Top             =   480
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         BuddyControl    =   "txtPrtSeq"
         BuddyDispid     =   196614
         OrigRight       =   240
         OrigBottom      =   735
         Max             =   99999
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         BackStyle       =   0  '투명
         Caption         =   "출력순서:"
         Height          =   180
         Left            =   7260
         TabIndex        =   17
         Top             =   540
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         BackStyle       =   0  '투명
         Caption         =   "문진코드 :"
         Height          =   180
         Left            =   480
         TabIndex        =   15
         Top             =   540
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         BackStyle       =   0  '투명
         Caption         =   "문진내역 :"
         Height          =   180
         Left            =   480
         TabIndex        =   14
         Top             =   900
         Width           =   840
      End
   End
End
Attribute VB_Name = "frmBBS818"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private objcom003 As New clsCom003

Private Sub chkWide_Click()
    TblWide
End Sub

Private Sub dtpDelete_CloseUp()
    If DeleteDate_Handle = False Then
       MsgBox "폐기일자는 현재 이후 일자만 선택 가능합니다!", vbCritical, "입력오류"
       dtpDelete.Value = Format(Date, "yyyy-mm-dd")
    End If
End Sub

Private Function DeleteDate_Handle() As Boolean
    Dim PDt As String
    Dim SDt As String
    
    PDt = Format(Date, PRESENTDATE_FORMAT)
    SDt = Format(dtpDelete, PRESENTDATE_FORMAT)
    
    If SDt < PDt Then
       DeleteDate_Handle = False
       Exit Function
    End If
    
    DeleteDate_Handle = True
End Function

Private Sub chkDelete_Click()
    If chkDelete.Value = 0 Then
       lblDelete.Visible = False
       dtpDelete.Visible = False
    Else
       lblDelete.Visible = True
       dtpDelete.Visible = True
       dtpDelete.Value = Format(Date, "yyyy-mm-dd")
       dtpDelete.Enabled = True
    End If
End Sub

Private Sub chkDelete_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdRefresh_Click()
    Clear
    txtQuestCode.SetFocus
End Sub

Private Sub cmdSave_Click()
    Dim ExpDt As String
    Dim yesno As String
    
    If txtQuestCode.Text = "" Then
       MsgBox "코드와 내역은 반드시 입력해야 합니다!", vbInformation, "알림"
       txtQuestCode.SetFocus
       Exit Sub
    End If
    
    If Trim(txtQuest.Text) = "" Then
       MsgBox "코드와 내역은 반드시 입력해야 합니다!", vbInformation, "알림"
       txtQuest.SetFocus
       Exit Sub
    End If
    
    If optQuest(0).Value = True Then
       yesno = "1"
    Else
       yesno = "0"
    End If
    If chkDelete.Value = 1 Then
        ExpDt = Format(dtpDelete, PRESENTDATE_FORMAT)
    Else
        ExpDt = ""
    End If
    
    With objcom003
        .CDINDEX = BC2_ASK
        .cdval1 = Trim(txtQuestCode)    ' 문진코드
        .field1 = yesno                 ' 정상값
        .field2 = Format(txtPrtSeq.Text, "00")            ' 출력순서
        .Field5 = ExpDt                 ' 폐기일자
        .Text1 = txtQuest               ' 문진내역
        
        If .Save() = True Then
            Clear
        End If
    End With
End Sub

Private Sub Form_Activate()
'    medMain.lblSubMenu.Caption = Me.Caption

End Sub

Private Sub Form_Load()
    Clear
End Sub

Private Sub tblList_Query()
    Dim strSql As String
    Dim RS As Recordset
    Dim MaxCnt As Integer
    Dim strText1 As String
    Dim i As Integer

    Set RS = objcom003.OpenRecordSet(BC2_ASK, , 1)

    With tblList
        .MaxRows = RS.RecordCount
        For i = 1 To RS.RecordCount
             .Row = i
             
             If i > .MaxRows Then .MaxRows = .MaxRows + 1
             
             .Col = 1:  .Value = RS.Fields("cdval1").Value & ""
             .Col = 2:  .Value = IIf((RS.Fields("field1").Value & "") = 1, "예", "아니오")
             .Col = 3:  .Value = RS.Fields("text1").Value & ""
             .Col = 4:  .Value = Val(RS.Fields("field2").Value & "")
             .Col = 5:
                        If RS.Fields("field5").Value & "" = "" Then
                            .Value = ""
                        Else
                            .Value = Format(RS.Fields("field5").Value & "", "####-##-##")
                        End If
             
             RS.MoveNext
        Next i
    End With
    
    Set RS = Nothing
    
'    '출력순서로 SORT
'    With tblList
'        .SortBy = SortByRow
'        .SortKey(1) = 4
'        .SortKey(2) = 1
'        .SortKeyOrder(1) = SortKeyOrderAscending
'        .SortKeyOrder(2) = SortKeyOrderAscending
'        .Col = 1: .Col2 = .MaxCols
'        .Row = 1: .Row2 = .DataRowCnt
'        .BlockMode = True
'        .Action = ActionSort
'        .BlockMode = False
'    End With
    
    Call TblWide
End Sub

Private Sub TblWide()
    Dim i As Long
    
    With tblList
        For i = 1 To .DataRowCnt
            .Row = i
            If chkWide.Value = 1 Then
                .RowHeight(i) = 12 '10.91
            Else
                .RowHeight(i) = .MaxTextRowHeight(i)
            End If
        Next i
    End With
End Sub

Private Sub Clear()
    'Lvw_Query
    tblList_Query
    
    txtQuestCode.Text = ""
    txtQuest.Text = ""
    txtPrtSeq = ""
    optQuest(0).Value = True
    chkDelete.Enabled = True
    chkDelete.Value = 0
    lblDelete.Visible = False
    dtpDelete.Visible = False
    cmdSave.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objcom003 = Nothing
    Unload Me
End Sub

Private Sub optQuest_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub tblList_Click(ByVal Col As Long, ByVal Row As Long)

    If Row < 1 Then Exit Sub
    If Row > tblList.MaxRows Then Exit Sub
    

    With tblList
        .Row = Row
        .Col = 1:   txtQuestCode = .Value
        .Col = 2:
                    If .Value = "예" Then
                        optQuest(0).Value = True
                    Else
                        optQuest(1).Value = True
                    End If
        .Col = 3:   txtQuest = .Value
        .Col = 4:   txtPrtSeq = .Value
        .Col = 5:
                    If .Value = "" Then
                        chkDelete.Value = 0
                        lblDelete.Visible = False
                        dtpDelete.Visible = False
                    Else
                        chkDelete.Value = 1
                        lblDelete.Visible = True
                        dtpDelete.Visible = True
                        dtpDelete = .Value
                    End If
                        
    End With
End Sub

Private Sub Move_CurRow()
    Dim i As Integer
    Dim strRow As String
    Dim CurRow As Integer
    
    With tblList
         For i = 1 To .MaxRows
             .Row = i: .Col = 1
             If .Value Like UCase(Trim(txtQuestCode)) & "*" Then
                .Row = i
                .Action = ActionActiveCell
                Exit Sub
             End If
         Next i
    End With
End Sub

Private Sub txtFindCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub tblList_KeyUp(KeyCode As Integer, Shift As Integer)
'    Dim CurRow As Integer
'    Dim strCol As String
'
'    SELECT Case KeyCode
'        Case 38
'             With tblList
'                  If .MaxRows > 0 Then
'                     .Action = ActionActiveCell
'                     If .Row = 1 Then Exit Sub
'                     CurRow = .Row - 1
'                     .Row = CurRow
'                     .Action = ActionActiveCell
'                     tblList_Click 1, .Row
'                  End If
'             End With
'        Case 40
'             With tblList
'                  If .MaxRows > 0 Then
'                     .Action = ActionActiveCell
'                     If .Row = .MaxRows Then Exit Sub
'                     CurRow = .Row + 1
'                     .Row = CurRow
'                     .Action = ActionActiveCell
'                     tblList_Click 1, .Row
'                  End If
'             End With
'
'    End SELECT
End Sub

Private Sub txtQuest_GotFocus()
    objcom003.SelFocus txtQuest
End Sub

Private Sub txtQuest_KeyPress(KeyAscii As Integer)
    'If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txtQuestCode_GotFocus()
    objcom003.SelFocus txtQuestCode
End Sub

Private Sub txtQuestCode_KeyPress(KeyAscii As Integer)
    Dim strHandle As Integer

    If KeyAscii = 13 Then
       If txtQuestCode.Text <> "" Then
          With objcom003
               '2(Update Handle):기존 Data가 존재 하므로 Dislpay
               strHandle = .DataHandleCOM003(BC2_ASK, txtQuestCode.Text)
               If strHandle = 0 Then
                  txtQuest.Text = ""
                  chkDelete.Enabled = True
                  chkDelete.Value = 0
                  lblDelete.Visible = False
                  dtpDelete.Visible = False
                  cmdSave.Enabled = True
                  SendKeys "{tab}"
                  Exit Sub
               End If
               
               Move_CurRow
               .DisplayCOM003 BC2_ASK, txtQuestCode.Text
               If .field1 = "예" Then
                  optQuest(0).Value = True
               Else
                  optQuest(1).Value = True
               End If
               txtQuest.Text = .Text1

               If .field2 <> "" Then
                  chkDelete.Value = 1
                  dtpDelete.Visible = True
                  dtpDelete.Value = Format(.field2, "####-##-##")
                  If DeleteDate_Handle = False Then
                     chkDelete.Enabled = False
                     dtpDelete.Enabled = False
                     cmdSave.Enabled = False
                     Exit Sub
                  End If
               Else
                  chkDelete.Value = 0
                  dtpDelete.Visible = False
               End If
          End With

          chkDelete.Enabled = True
          cmdSave.Enabled = True
       End If
       SendKeys "{tab}"
    End If
       
End Sub

Private Sub txtQuestCode_LostFocus()
    If txtQuestCode.Text <> "" Then txtQuestCode.Text = UCase(Trim(txtQuestCode))
End Sub




