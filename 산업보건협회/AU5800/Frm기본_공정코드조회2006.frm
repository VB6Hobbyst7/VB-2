VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm기본_공정코드조회2006 
   Caption         =   "공정코드 조회"
   ClientHeight    =   7815
   ClientLeft      =   3855
   ClientTop       =   3795
   ClientWidth     =   8235
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Frm기본_공정코드조회2006.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7815
   ScaleWidth      =   8235
   Begin VB.OptionButton OptSection 
      Caption         =   "코드명"
      Height          =   180
      Index           =   2
      Left            =   1740
      TabIndex        =   2
      Top             =   240
      Width           =   915
   End
   Begin VB.OptionButton OptSection 
      Caption         =   "코드"
      Height          =   180
      Index           =   1
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
   Begin VB.OptionButton OptSection 
      Caption         =   "전체"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Value           =   -1  'True
      Width           =   735
   End
   Begin VB.TextBox Txt찾을문자열 
      Height          =   315
      IMEMode         =   10  '한글 
      Left            =   60
      TabIndex        =   3
      Text            =   "123"
      Top             =   540
      Width           =   2655
   End
   Begin VB.CommandButton CmdUpdate 
      Caption         =   "수정(&U)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   4620
      Picture         =   "Frm기본_공정코드조회2006.frx":000C
      Style           =   1  '그래픽
      TabIndex        =   6
      Top             =   60
      Width           =   855
   End
   Begin VB.CommandButton CmdInput 
      Caption         =   "신규(&I)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   3720
      Picture         =   "Frm기본_공정코드조회2006.frx":044E
      Style           =   1  '그래픽
      TabIndex        =   5
      Top             =   60
      Width           =   855
   End
   Begin VB.CommandButton CmdPrint 
      Caption         =   "출력(&P)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   6420
      Picture         =   "Frm기본_공정코드조회2006.frx":0890
      Style           =   1  '그래픽
      TabIndex        =   8
      Top             =   60
      Width           =   855
   End
   Begin VB.CommandButton CmdView 
      Caption         =   "조회(&V)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   2820
      Picture         =   "Frm기본_공정코드조회2006.frx":0B9A
      Style           =   1  '그래픽
      TabIndex        =   4
      Top             =   60
      Width           =   855
   End
   Begin VB.CommandButton CmdDelete 
      Caption         =   "삭제(&D)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   5520
      Picture         =   "Frm기본_공정코드조회2006.frx":1464
      Style           =   1  '그래픽
      TabIndex        =   7
      Top             =   60
      Width           =   855
   End
   Begin VB.CommandButton CmdQuit 
      Caption         =   "닫기(&Q)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   7320
      Picture         =   "Frm기본_공정코드조회2006.frx":176E
      Style           =   1  '그래픽
      TabIndex        =   9
      Top             =   60
      Width           =   855
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   195
      Left            =   60
      TabIndex        =   10
      Top             =   900
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin FPSpread.vaSpread vaSpread공정코드 
      Height          =   6615
      Left            =   60
      TabIndex        =   11
      Top             =   1140
      Width           =   8115
      _Version        =   393216
      _ExtentX        =   14314
      _ExtentY        =   11668
      _StockProps     =   64
      AllowMultiBlocks=   -1  'True
      ArrowsExitEditMode=   -1  'True
      BackColorStyle  =   1
      ButtonDrawMode  =   4
      ColsFrozen      =   1
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   3
      MaxRows         =   20
      SpreadDesigner  =   "Frm기본_공정코드조회2006.frx":1BB0
      UserResize      =   1
   End
End
Attribute VB_Name = "Frm기본_공정코드조회2006"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDCK_CANCEL()
    If vaSpread공정코드.MaxRows > 0 Then vaSpread공정코드.MaxRows = 0
End Sub

Private Sub MDCK_DELETE()
    loAdoCnn.BeginTrans

    StrSQL = "         DELETE FROM BAG_GONGJUNG "
    StrSQL = StrSQL & " WHERE GONGJUNG_CODE  = '" & GET_CELL(vaSpread공정코드, 1, vaSpread공정코드.ActiveRow) & "' "
    If RunADO(StrSQL) = False Then Exit Sub
    
    loAdoCnn.CommitTrans
End Sub

Private Sub MDCK_INITIAL()
    Txt찾을문자열.Enabled = False
    Txt찾을문자열 = ""
    ProgressBar1.Visible = False
    Call MDCK_CANCEL
End Sub

Private Sub MDCK_KEY_CLEAR()

End Sub

Private Sub MDCK_PRINT()
    Dim strFont1  As String
    Dim strFont2  As String
    Dim strHead1  As String
    Dim strHead2  As String
    Dim strHead3  As String
    Dim strHead4  As String
    
    If vaSpread공정코드.MaxRows = 0 Then MsgBox "출력할 자료가 없습니다.", vbInformation, "확인": Exit Sub
    
    If MsgBox("출력하겠습니까?", vbQuestion + vbOKCancel, "출력여부") = vbCancel Then Exit Sub
    
    strFont1 = "/fn""굴림체""/fz""15""/fb1/fi0/fu1/fk0/fs1"
    strFont2 = "/fn""굴림체""/fz""10""/fb0/fi0/fu0/fk0/fs2"

    strHead1 = "/f1/c" & "공정코드" & "/n/n/n"
    
    With vaSpread공정코드
        .PrintAbortMsg = "공정코드 출력 중..."
        .PrintHeader = strFont1 + strHead1 + strFont2
        .PrintFooter = "/c" & "PAGE : " & "/P"
        .PrintBorder = True
        .PrintGrid = True
        .PrintColHeaders = True
        .PrintRowHeaders = True
        .PrintColor = False
        .PrintMarginTop = 500
        .PrintMarginBottom = 500
        .PrintMarginLeft = 500
        .PrintMarginRight = 0
        .PrintType = SS_PRINT_ALL
        .PrintShadows = False
        .PrintUseDataMax = False
        .Action = SS_ACTION_SMARTPRINT
    End With
End Sub

Private Sub MDCK_SAVE()

End Sub

Private Sub MDCK_VIEW()
    Dim ni%
    
    If Trim(Txt찾을문자열) = "" And OptSection(2).Value = True Then Exit Sub
    
    With vaSpread공정코드
        If .MaxRows > 0 Then .MaxRows = 0
        ProgressBar1.Value = 0
        ProgressBar1.Visible = True
        
        StrSQL = "         SELECT * "
        StrSQL = StrSQL & "  FROM BAG_GONGJUNG "
        Select Case True
            Case OptSection(0).Value '/전체
                StrSQL = StrSQL & " ORDER BY GONGJUNG_CODE "
            Case OptSection(1).Value '/코드
                StrSQL = StrSQL & " WHERE GONGJUNG_CODE = '" & Replace(Trim(Txt찾을문자열), "-", "") & "' "
                StrSQL = StrSQL & " ORDER BY GONGJUNG_CODE "
            Case OptSection(2).Value '/코드명
                StrSQL = StrSQL & " WHERE GONGJUNG_NAME LIKE '%" & Trim(Txt찾을문자열) & "%' "
                StrSQL = StrSQL & " ORDER BY GONGJUNG_NAME, GONGJUNG_CODE "
        End Select
        If ReadADO(StrSQL, 0) = True Then
            .MaxRows = GdRecordCount
            ProgressBar1.Max = GdRecordCount
            
            Do Until ARS(0).EOF
                ni% = ni% + 1: .Row = ni%: ProgressBar1.Value = ni%
                
                .Col = 1: .Text = Trim(ARS(0)!GONGJUNG_CODE & "") '/공정코드
                .Col = 2: .Text = Trim(ARS(0)!GONGJUNG_NAME & "") '/공정명
                .Col = 3                                      '/사용여부
                If Trim(ARS(0)!GONGJUNG_USE & "") = "1" Then
                    .Text = "○"
                Else
                    .Text = "Χ"
                End If
                
                If .MaxTextRowHeight(ni%) > 12.95 Then
                    .RowHeight(ni%) = .MaxTextRowHeight(ni%)
                End If
                
                ARS(0).MoveNext
            Loop
            Call CloseADO(ARS(0))
        End If
        ProgressBar1.Visible = False
    End With
End Sub

Private Sub CmdDelete_Click()
    If vaSpread공정코드.ActiveRow = 0 Then MsgBox "삭제할 내용을 선택하십시오", vbInformation, "확인": Exit Sub
    
    If MsgBox("공정코드 : " & GET_CELL(vaSpread공정코드, 1, vaSpread공정코드.ActiveRow) & vbCrLf & _
              "공정명   : " & GET_CELL(vaSpread공정코드, 2, vaSpread공정코드.ActiveRow) & vbCrLf & vbCrLf & _
              "위 자료를 삭제하겠습니까?", vbQuestion + vbOKCancel, "삭제질의") = vbCancel Then Exit Sub
    
    Call MDCK_DELETE
    
    Call MDCK_VIEW

    MsgBox "삭제되었습니다!", vbInformation, "확인"
End Sub

Private Sub CmdInput_Click()
    GstrInputUpdate = "1" '/1.Input, 2.Update
    GstrInputUpdateYN = "2" '/1.입력및수정함, 2.변화없음
    
    Frm기본_공정코드입력2006.Show vbModal
    
    If GstrInputUpdateYN = "1" Then Call MDCK_VIEW
End Sub

Private Sub CmdPrint_Click()
    Call MDCK_PRINT
End Sub

Private Sub CmdQuit_Click()
    Unload Me
End Sub

Private Sub CmdUpdate_Click()
    If vaSpread공정코드.ActiveRow = 0 Then MsgBox "수정할 대상을 선택하십시오!", vbInformation, "확인": Exit Sub
    
    GstrInputUpdate = "2" '/1.Input, 2.Update
    GstrInputUpdateYN = "2" '/1.입력및수정함, 2.변화없음
    GstrArgTemp1 = GET_CELL(vaSpread공정코드, 1, vaSpread공정코드.ActiveRow)
    
    Frm기본_공정코드입력2006.Show vbModal
    
    If GstrInputUpdateYN = "1" Then Call MDCK_VIEW
End Sub

Private Sub CmdView_Click()
    Call MDCK_VIEW
End Sub

Private Sub Form_Load()
    Me.Height = 8310
    Me.Width = 8355
    Me.Top = 0
    Me.Left = 0
    Me.Show
    
    Call MDCK_INITIAL
    DoEvents
    Call MDCK_VIEW
    
    vaSpread공정코드.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Frm기본_공정코드조회2006 = Nothing
End Sub

Private Sub OptSection_Click(Index As Integer)
    Select Case Index
        Case 0
            Txt찾을문자열.Enabled = False
        Case 1
            Txt찾을문자열.Enabled = True
            Txt찾을문자열.IMEMode = 8
        Case 2
            Txt찾을문자열.Enabled = True
            Txt찾을문자열.IMEMode = 10
    End Select
End Sub

Private Sub OptSection_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Index <> 0 Then
            Txt찾을문자열.SetFocus
        Else
            CmdView.SetFocus
        End If
    End If
End Sub

Private Sub Txt찾을문자열_GotFocus()
    Call TEXTSELECT
End Sub

Private Sub Txt찾을문자열_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub vaSpread공정코드_Click(ByVal Col As Long, ByVal Row As Long)
    With vaSpread공정코드
        If Row <> 0 Then Exit Sub
        If Col = 0 Then Exit Sub
        
        .Col = -1
        .Row = 1
        .Col2 = -1
        .Row2 = .MaxRows
        .BlockMode = True
        .SortBy = SS_SORT_BY_ROW
        
        .SortKey(1) = Col
        If Val(Mid(vaSpread공정코드.Tag, 2)) = Col Then
            If Left(vaSpread공정코드.Tag, 1) = "A" Then
                .SortKeyOrder(1) = SS_SORT_ORDER_DESCENDING
                vaSpread공정코드.Tag = "D" & CStr(Col)
            Else
                .SortKeyOrder(1) = SS_SORT_ORDER_ASCENDING
                vaSpread공정코드.Tag = "A" & CStr(Col)
            End If
        Else
            .SortKeyOrder(1) = SS_SORT_ORDER_ASCENDING
            vaSpread공정코드.Tag = "A" & CStr(Col)
        End If
        
        .Action = SS_ACTION_SORT
        .BlockMode = False
    End With
End Sub

Private Sub vaSpread공정코드_DblClick(ByVal Col As Long, ByVal Row As Long)
    If Row > 0 Then
        GstrInputUpdate = "2" '/1.Input, 2.Update
        GstrInputUpdateYN = "2" '/1.입력및수정함, 2.변화없음
        GstrArgTemp1 = GET_CELL(vaSpread공정코드, 1, vaSpread공정코드.ActiveRow)
        
        Frm기본_공정코드입력2006.Show vbModal
        
        If GstrInputUpdateYN = "1" Then Call MDCK_VIEW
    End If
End Sub

Private Sub vaSpread공정코드_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call vaSpread공정코드_DblClick(1, vaSpread공정코드.ActiveRow)
    End If
End Sub
