VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Begin VB.Form frmSpec12 
   Caption         =   "Slip구분-검사종류 관리"
   ClientHeight    =   6315
   ClientLeft      =   1140
   ClientTop       =   2325
   ClientWidth     =   11055
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   11055
   Begin Threed.SSPanel SSPanel1 
      Height          =   2355
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   10815
      _Version        =   65536
      _ExtentX        =   19076
      _ExtentY        =   4154
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtCodeName 
         Appearance      =   0  '평면
         BackColor       =   &H00C0E0FF&
         Height          =   315
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   240
         Width           =   2175
      End
      Begin Threed.SSCommand cmdHelp 
         Height          =   315
         Left            =   1860
         TabIndex        =   0
         Top             =   240
         Width           =   315
         _Version        =   65536
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   78
         Caption         =   "&H"
      End
      Begin VB.TextBox txtCodegu 
         Height          =   315
         Left            =   1200
         TabIndex        =   33
         Top             =   240
         Width           =   675
      End
      Begin VB.TextBox txtSugaCD 
         Height          =   315
         Left            =   1200
         TabIndex        =   5
         Top             =   1920
         Width           =   1455
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   855
         Left            =   9840
         TabIndex        =   24
         Top             =   720
         Width           =   795
         _Version        =   65536
         _ExtentX        =   1402
         _ExtentY        =   1508
         _StockProps     =   78
         Caption         =   "화면정리"
         BevelWidth      =   1
         Outline         =   0   'False
         MouseIcon       =   "frmSpec12.frx":0000
      End
      Begin Threed.SSCommand cmdDel 
         Height          =   855
         Left            =   9000
         TabIndex        =   23
         Top             =   720
         Width           =   795
         _Version        =   65536
         _ExtentX        =   1402
         _ExtentY        =   1508
         _StockProps     =   78
         Caption         =   "삭제확인"
         BevelWidth      =   1
         Outline         =   0   'False
         MouseIcon       =   "frmSpec12.frx":001C
      End
      Begin Threed.SSCommand cmdInsert 
         Height          =   855
         Left            =   8160
         TabIndex        =   11
         Top             =   720
         Width           =   795
         _Version        =   65536
         _ExtentX        =   1402
         _ExtentY        =   1508
         _StockProps     =   78
         Caption         =   "입력확인"
         BevelWidth      =   1
         Outline         =   0   'False
         MouseIcon       =   "frmSpec12.frx":0038
      End
      Begin VB.ComboBox cmbJangbi5 
         Height          =   300
         Left            =   4560
         TabIndex        =   10
         Top             =   1920
         Width           =   1335
      End
      Begin VB.ComboBox cmbJangbi4 
         Height          =   300
         Left            =   4560
         TabIndex        =   9
         Top             =   1620
         Width           =   1335
      End
      Begin VB.ComboBox cmbJangbi3 
         Height          =   300
         Left            =   4560
         TabIndex        =   8
         Top             =   1320
         Width           =   1335
      End
      Begin VB.ComboBox cmbJangbi2 
         Height          =   300
         Left            =   4560
         TabIndex        =   7
         Top             =   1020
         Width           =   1335
      End
      Begin VB.ComboBox cmbJangbi1 
         Height          =   300
         Left            =   4560
         TabIndex        =   6
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtYageo 
         Height          =   315
         Left            =   1200
         TabIndex        =   4
         Top             =   1620
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtCoDate 
         Height          =   315
         Left            =   1200
         TabIndex        =   3
         Top             =   1320
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   24576003
         CurrentDate     =   36297
      End
      Begin VB.TextBox txtCodenm 
         Height          =   315
         Left            =   1200
         TabIndex        =   2
         Top             =   1020
         Width           =   2175
      End
      Begin VB.TextBox txtCodeky 
         Height          =   315
         Left            =   1200
         TabIndex        =   1
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label14 
         BackColor       =   &H80000007&
         Caption         =   "Label9"
         Height          =   855
         Left            =   9900
         TabIndex        =   36
         Top             =   780
         Width           =   795
      End
      Begin VB.Label Label13 
         BackColor       =   &H80000007&
         Caption         =   "Label9"
         Height          =   855
         Left            =   9060
         TabIndex        =   35
         Top             =   780
         Width           =   795
      End
      Begin VB.Label Label12 
         BackColor       =   &H80000007&
         Caption         =   "Label9"
         Height          =   855
         Left            =   8220
         TabIndex        =   34
         Top             =   780
         Width           =   795
      End
      Begin VB.Label Label1 
         Caption         =   "Code 구분"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   300
         Width           =   915
      End
      Begin VB.Label labJangbi5 
         Appearance      =   0  '평면
         BackColor       =   &H80000000&
         BorderStyle     =   1  '단일 고정
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5940
         TabIndex        =   30
         Top             =   1920
         Width           =   1755
      End
      Begin VB.Label labJangbi4 
         Appearance      =   0  '평면
         BackColor       =   &H80000000&
         BorderStyle     =   1  '단일 고정
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5940
         TabIndex        =   29
         Top             =   1620
         Width           =   1755
      End
      Begin VB.Label labJangbi3 
         Appearance      =   0  '평면
         BackColor       =   &H80000000&
         BorderStyle     =   1  '단일 고정
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5940
         TabIndex        =   28
         Top             =   1320
         Width           =   1755
      End
      Begin VB.Label labJangbi2 
         Appearance      =   0  '평면
         BackColor       =   &H80000000&
         BorderStyle     =   1  '단일 고정
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5940
         TabIndex        =   27
         Top             =   1020
         Width           =   1755
      End
      Begin VB.Label labJangbi1 
         Appearance      =   0  '평면
         BackColor       =   &H80000000&
         BorderStyle     =   1  '단일 고정
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5940
         TabIndex        =   26
         Top             =   720
         Width           =   1755
      End
      Begin VB.Label Label11 
         Caption         =   "수가코드"
         Height          =   195
         Left            =   180
         TabIndex        =   25
         Top             =   1980
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "장비코드5"
         Height          =   255
         Left            =   3540
         TabIndex        =   22
         Top             =   1980
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "장비코드4"
         Height          =   255
         Left            =   3540
         TabIndex        =   21
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "장비코드3"
         Height          =   255
         Left            =   3540
         TabIndex        =   20
         Top             =   1380
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "장비코드2"
         Height          =   255
         Left            =   3540
         TabIndex        =   19
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "장비코드1"
         Height          =   255
         Left            =   3540
         TabIndex        =   18
         Top             =   780
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "약어"
         Height          =   195
         Left            =   180
         TabIndex        =   17
         Top             =   1680
         Width           =   555
      End
      Begin VB.Label Label4 
         Caption         =   "등록일자"
         Height          =   255
         Left            =   180
         TabIndex        =   16
         Top             =   1380
         Width           =   915
      End
      Begin VB.Label Label3 
         Caption         =   "검사이름"
         Height          =   195
         Left            =   180
         TabIndex        =   15
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "검사코드"
         Height          =   195
         Left            =   180
         TabIndex        =   14
         Top             =   780
         Width           =   855
      End
   End
   Begin FPSpreadADO.fpSpread ssSpec12 
      Height          =   3510
      Left            =   135
      TabIndex        =   12
      Top             =   2520
      Width           =   10815
      _Version        =   196608
      _ExtentX        =   19076
      _ExtentY        =   6191
      _StockProps     =   64
      BackColorStyle  =   1
      ColsFrozen      =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   10
      MaxRows         =   100
      ScrollBarExtMode=   -1  'True
      ScrollBars      =   2
      SpreadDesigner  =   "frmSpec12.frx":0054
      UserResize      =   1
      VisibleCols     =   10
      VisibleRows     =   100
      Appearance      =   1
   End
   Begin VB.Menu mnuQuit 
      Caption         =   "Quit"
   End
End
Attribute VB_Name = "frmSpec12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub cmdClear_Click()
    
    ssSpec12.Row = 1
    ssSpec12.Col = 1
    ssSpec12.Action = SS_ACTION_ACTIVE_CELL
    
    For i = 0 To Me.Count - 1
        If TypeOf Me.Controls(i) Is TextBox Then
            Me.Controls(i).Text = ""
        ElseIf TypeOf Me.Controls(i) Is ComboBox Then
            Me.Controls(i).ListIndex = -1
        ElseIf TypeOf Me.Controls(i) Is fpSpread Then
            Me.Controls(i).Row = 1
            Me.Controls(i).Row2 = Me.Controls(i).DataRowCnt
            Me.Controls(i).Col = 1
            Me.Controls(i).Col2 = Me.Controls(i).DataColCnt
            Me.Controls(i).BlockMode = True
            Me.Controls(i).Action = SS_ACTION_CLEAR_TEXT
            Me.Controls(i).BlockMode = False
        End If
    Next
    
    mdiMain.stbMain.Panels(1).Text = ""
        

End Sub

Private Sub cmdDel_Click()

    If Trim(txtCodeky.Text) = "" Then
        MsgBox "삭제할 Code 가 없습니다!.", vbInformation + vbOKOnly, "Please Data Check..."
        Exit Sub
    End If

    sMsg = Trim$(txtCodeky.Text) & " = " & Trim(txtCodenm.Text) & vbCrLf & "의 Data 를 삭제하시겠습니까?"
    If vbNo = MsgBox(sMsg, vbYesNo + vbQuestion, "삭제 확인 Box") Then
        Exit Sub
    End If
    
    strSql = ""
    strSql = strSql & " DELETE"
    strSql = strSql & " FROM   TWEXAM_SPECODE"
    strSql = strSql & " WHERE  Codegu = '" & Trim(txtCodegu.Text) & "'"
    strSql = strSql & " AND    Codeky = '" & Trim(txtCodeky.Text) & "'"
    adoConnect.BeginTrans
    If adoExec(strSql) = True Then
        adoConnect.CommitTrans
        mdiMain.stbMain.Panels(1).Text = "Data를 삭제하였습니다!.."
        GoSub Upper_Clear
    Else
        adoConnect.RollbackTrans
        mdiMain.stbMain.Panels(1).Text = "Data 삭제시 오류가 났습니다!..."
    End If
    
    Call cmdQry_Click
    Exit Sub
           
Upper_Clear:
    txtCodeky.Text = ""
    txtCodenm.Text = ""
    txtYageo.Text = ""
    dtCoDate.Value = Dual_Date_Get("yyyy-MM-dd")
    cmbJangbi1.Text = ""
    cmbJangbi2.Text = ""
    cmbJangbi3.Text = ""
    cmbJangbi4.Text = ""
    cmbJangbi5.Text = ""
    
    Return
    
End Sub

Public Sub cmdHelp_Click()
    
    frmHelpSpec.Show vbModal
    SendKeys "{TAB}"
    
    
End Sub

Private Sub cmdInsert_Click()
    Dim sdate       As String
    
    If Trim(txtCodeky.Text) = "" Then
        MsgBox "입력할 검사종류의 Code 가 없습니다!.."
        Exit Sub
    End If
    
    If Trim(txtCodegu.Text) = "" Then
        MsgBox "입력할 Slip 구분을 먼저 선택하세요!.."
        Exit Sub
    End If
    
    sdate = Format(dtCoDate.Value, "yyyy-MM-dd")
    
    strSql = ""
    strSql = strSql & " SELECT *"
    strSql = strSql & " FROM   TWEXAM_SPECODE"
    strSql = strSql & " WHERE  Codegu  =  '" & Trim(txtCodegu.Text) & "'"
    strSql = strSql & " AND    Codeky  =  '" & Trim(txtCodeky.Text) & "'"
    
    If False = adoSetOpen(strSql, adoSet) Then
        GoSub Specode12_Insert_Sub
    Else
        GoSub Specode12_Update_Sub
    End If
    Exit Sub
    
    
'/___________________________________________________________________________
Specode12_Insert_Sub:
    strSql = ""
    strSql = strSql & " INSERT INTO TWEXAM_SPECODE"
    strSql = strSql & "        (Codegu,  Codeky,  Codenm,  Codate,  Yageo,  Sugacd,"
    strSql = strSql & "         Jangbi1, Jangbi2, Jangbi3, Jangbi4, Jangbi5)"
    strSql = strSql & " VALUES( '" & Left(txtCodegu.Text, 2) & "',"
    strSql = strSql & "         '" & Trim(txtCodeky.Text) & "',"
    strSql = strSql & "         '" & Quot_Conv(Trim(txtCodenm.Text)) & "',"
    strSql = strSql & "              TO_DATE('" & sdate & "','YYYY-MM-DD'),"
    strSql = strSql & "         '" & Trim(txtYageo.Text) & "',"
    strSql = strSql & "         '" & Trim(txtSugaCD.Text) & "',"
    strSql = strSql & "         '" & Trim(cmbJangbi1.Text) & "',"
    strSql = strSql & "         '" & Trim(cmbJangbi2.Text) & "',"
    strSql = strSql & "         '" & Trim(cmbJangbi3.Text) & "',"
    strSql = strSql & "         '" & Trim(cmbJangbi4.Text) & "',"
    strSql = strSql & "         '" & Trim(cmbJangbi5.Text) & "')"
    
    If adoExec(strSql) Then
        mdiMain.stbMain.Panels(1).Text = "Data 가 신규 입력 되었습니다!.."
        GoSub Upper_Clear
        Call cmdQry_Click
    Else
        mdiMain.stbMain.Panels(1).Text = "신규 입력시 오류가 발생하였습니다!..."
    End If
    
    Return


Specode12_Update_Sub:
    strSql = ""
    strSql = strSql & " UPDATE TWEXAM_SPECODE"
    strSql = strSql & " SET    Codenm  =  '" & Quot_Conv(Trim(txtCodenm.Text)) & "',"
    strSql = strSql & "        Codate  =       TO_DATE('" & sdate & "','YYYY-MM-DD'),"
    strSql = strSql & "        Yageo   =  '" & Trim(txtYageo.Text) & "',"
    strSql = strSql & "        Sugacd  =  '" & Trim(txtSugaCD.Text) & "',"
    strSql = strSql & "        Jangbi1 =  '" & Trim(cmbJangbi1.Text) & "',"
    strSql = strSql & "        Jangbi2 =  '" & Trim(cmbJangbi2.Text) & "',"
    strSql = strSql & "        Jangbi3 =  '" & Trim(cmbJangbi3.Text) & "',"
    strSql = strSql & "        Jangbi4 =  '" & Trim(cmbJangbi4.Text) & "',"
    strSql = strSql & "        Jangbi5 =  '" & Trim(cmbJangbi5.Text) & "'"
    strSql = strSql & " WHERE  Codegu  =  '" & Left(txtCodegu.Text, 2) & "'"
    strSql = strSql & " AND    Codeky  =  '" & Trim(txtCodeky.Text) & "'"
    If adoExec(strSql) Then
        mdiMain.stbMain.Panels(1).Text = "Data 가 수정 되었습니다!.."
        GoSub Upper_Clear
        Call cmdQry_Click
    Else
        mdiMain.stbMain.Panels(1).Text = "Data 수정입력시 오류가 발생하였습니다!..."
    End If
    
    Return


Upper_Clear:
    txtCodeky.Text = ""
    txtCodenm.Text = ""
    txtYageo.Text = ""
    dtCoDate.Value = Dual_Date_Get("yyyy-MM-dd")
    cmbJangbi1.Text = ""
    cmbJangbi2.Text = ""
    cmbJangbi3.Text = ""
    cmbJangbi4.Text = ""
    cmbJangbi5.Text = ""
    
    Return

End Sub

Private Sub cmdPrint_Click()
    
    
    If ssSpec12.DataRowCnt = 0 Then Exit Sub
    If Trim(txtCodeName.Text) = "" Then Exit Sub
    If vbNo = MsgBox(Trim(txtCodeName.Text) & "Data 의 Print 작업을 하시겠습니까?", _
                     vbYesNo + vbQuestion, _
                     "출력 작업 확인MessageBox") Then Exit Sub
    
    Dim strFont(1)        As String
    Dim strHead(1)        As String
    
    strFont(0) = "/fn""굴림체"" /fz""16"" /fb1 /fi0 /fu0 /fk0 /fs1"
    strFont(1) = "/fn""굴림체"" /fz""10"" /fb0 /fi0 /fu0 /fk0 /fs2"
    strHead(0) = "/f1" & "/c" & Trim(txtCodeName.Text)
    'strHead(1) = "/f2" & "Page : " & "/p" & " of " & ssSpec12.PrintPageCount & "/r"
    
    ssSpec12.PrintHeader = strFont(0) + strHead(0) + "/n/n" + strFont(1) + strHead(1) + "/n" + strFont(1)
    ssSpec12.PrintFooter = "/f2" & "/c" & "Page : " & "/p" & " of " & ssSpec12.PrintPageCount
    ssSpec12.PrintMarginLeft = 300
    ssSpec12.PrintMarginRight = 0
    ssSpec12.PrintMarginTop = 150
    ssSpec12.PrintMarginBottom = 500
    ssSpec12.PrintColHeaders = True
    ssSpec12.PrintRowHeaders = True
    ssSpec12.PrintBorder = True
    ssSpec12.PrintColor = True
    ssSpec12.PrintGrid = True
    ssSpec12.PrintShadows = True
    ssSpec12.PrintUseDataMax = False
    ssSpec12.Row = 1
    ssSpec12.Row2 = ssSpec12.DataRowCnt
    ssSpec12.Col = 1
    ssSpec12.Col2 = ssSpec12.MaxCols
    ssSpec12.PrintOrientation = PrintOrientationLandscape
    ssSpec12.PrintType = PrintTypeCellRange
    ssSpec12.Action = ActionPrint
            
    
End Sub

Private Sub cmdQry_Click()

    If Trim(txtCodegu.Text) = "" Then
        Call cmdHelp_Click
    End If
    Screen.MousePointer = vbHourglass
    GoSub Query_Data
    Screen.MousePointer = vbDefault
    Exit Sub
    
    
Query_Data:
    strSql = ""
    strSql = strSql & " SELECT  a.*,"
    strSql = strSql & "         TO_CHAR(a.Codate, 'YYYY-MM-DD') Codate"
    strSql = strSql & " FROM    TWEXAM_SPECODE a"
    strSql = strSql & " WHERE   a.Codegu  =  '" & Trim(txtCodegu.Text) & "'"
    strSql = strSql & " ORDER   BY a.Codeky, a.Codate"
    
    ssSpec12.MaxRows = 0
    If False = adoSetOpen(strSql, adoSet) Then
        Return
    End If
    
    ssSpec12.MaxRows = adoSet.RecordCount
    
    Do Until adoSet.EOF
        ssSpec12.Row = ssSpec12.DataRowCnt + 1
        ssSpec12.Col = 1:  ssSpec12.Text = Trim(adoSet.Fields("Codeky").Value & "")
        ssSpec12.Col = 2:  ssSpec12.Text = Trim(adoSet.Fields("Codenm").Value & "")
        ssSpec12.Col = 3:  ssSpec12.Text = Trim(adoSet.Fields("Codate").Value & "")
        ssSpec12.Col = 4:  ssSpec12.Text = Trim(adoSet.Fields("Yageo").Value & "")
        ssSpec12.Col = 5:  ssSpec12.Text = Trim(adoSet.Fields("Sugacd").Value & "")
        ssSpec12.Col = 6:  ssSpec12.Text = Trim(adoSet.Fields("Jangbi1").Value & "")
        ssSpec12.Col = 7:  ssSpec12.Text = Trim(adoSet.Fields("Jangbi2").Value & "")
        ssSpec12.Col = 8:  ssSpec12.Text = Trim(adoSet.Fields("Jangbi3").Value & "")
        ssSpec12.Col = 9:  ssSpec12.Text = Trim(adoSet.Fields("Jangbi4").Value & "")
        ssSpec12.Col = 10: ssSpec12.Text = Trim(adoSet.Fields("Jangbi5").Value & "")
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    ssSpec12.SetFocus
    ssSpec12.Row = 1
    ssSpec12.Col = 1
    ssSpec12.Action = SS_ACTION_ACTIVE_CELL
    Return
    
End Sub


Private Sub cmdQry1_Click()

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
    
End Sub

Private Sub Form_Load()
    
    Me.Top = 1
    Me.Left = 1
    Me.Height = 7000
    Me.Width = 11200

    dtCoDate.Value = Dual_Date_Get("yyyy-MM-dd")
    
End Sub

Private Sub mnuQuit_Click()
    Unload Me
    
End Sub
Private Sub ssSpec12_DblClick(ByVal Col As Long, ByVal Row As Long)
    
    If Row > 0 Then
        ssSpec12.Row = Row
        ssSpec12.Col = 1
        txtCodeky.Text = Trim(ssSpec12.Text)
        Call txtCodeky_LostFocus
        Exit Sub
    End If
    
    If Row = 0 Then
        ssSpec12.Row = 1
        ssSpec12.Col = 1
        ssSpec12.Row2 = ssSpec12.DataRowCnt
        ssSpec12.Col2 = ssSpec12.DataColCnt
        ssSpec12.SortBy = SS_SORT_BY_ROW
        ssSpec12.SortKey(1) = Col
        ssSpec12.SortKeyOrder(1) = SS_SORT_ORDER_ASCENDING
        ssSpec12.Action = SS_ACTION_SORT
    End If
    
End Sub

Private Sub txtCodegu_LostFocus()
    
    If Not IsNumeric(txtCodegu.Text) Then
        txtCodegu.Text = ""
        txtCodeName.Text = ""
    End If
    If Len(txtCodegu.Text) > 2 Then
        txtCodegu.Text = ""
        txtCodeName.Text = ""
    End If
    
    txtCodegu.Text = Format(txtCodegu.Text, "00")
    
    Select Case txtCodegu.Text
        Case "12": txtCodeName.Text = "검사종류"
        Case "08": txtCodeName.Text = "미생물균 군분류"
        Case "13": txtCodeName.Text = "검체"
        Case "15": txtCodeName.Text = "미생물약제"
        Case "16": txtCodeName.Text = "미생물종류"
        Case "18": txtCodeName.Text = "결핵균"
        Case "21": txtCodeName.Text = "장비종류"
        Case "50": txtCodeName.Text = "해부병리"
        Case "55": txtCodeName.Text = "특수염색"
        Case Else: txtCodeName.Text = ""
                   txtCodegu.Text = ""
    End Select
    
    txtCodegu.Tag = txtCodegu.Text
    txtCodeName.Tag = txtCodeName.Text
    Call ClearForm(Me)
    
    If Trim(txtCodeName.Tag) <> "" Then
        txtCodegu.Text = txtCodegu.Tag
        txtCodeName.Text = txtCodeName.Tag
        txtCodeky.SetFocus
    End If
    
    txtCodegu.Tag = ""
    txtCodeName.Tag = ""
    
End Sub

Public Sub txtCodeky_LostFocus()
    
    If Trim(txtCodeky.Text) = "" Then Exit Sub
    
    GoSub Get_Specode12_Data
    Exit Sub
    
    
Get_Specode12_Data:
    strSql = ""
    strSql = strSql & " SELECT a.*,"
    strSql = strSql & "        TO_CHAR(a.Codate, 'YYYY-MM-DD') Codate"
    strSql = strSql & " FROM   TWEXAM_SPECODE a"
    strSql = strSql & " WHERE  a.Codegu = '" & Trim(txtCodegu.Text) & "'"
    strSql = strSql & " AND    a.Codeky = '" & Trim(txtCodeky.Text) & "'"
    
    If False = adoSetOpen(strSql, adoSet) Then
        dtCoDate.Value = Dual_Date_Get("yyyy-MM-dd")
        txtYageo.Text = ""
        txtSugaCD.Text = ""
        cmbJangbi1.ListIndex = -1
        cmbJangbi2.ListIndex = -1
        cmbJangbi3.ListIndex = -1
        cmbJangbi4.ListIndex = -1
        cmbJangbi5.ListIndex = -1
        Return
    Else
        txtCodenm.Text = Trim(adoSet.Fields("Codenm").Value & "")
        If Not IsNull(adoSet.Fields("Codate").Value) Then
            dtCoDate.Value = adoSet.Fields("Codate").Value
        Else
            dtCoDate.Value = Dual_Date_Get("yyyy-MM-dd")
        End If
        txtYageo.Text = Trim(adoSet.Fields("Yageo").Value & "")
        cmbJangbi1.Text = Trim(adoSet.Fields("Jangbi1").Value & "")
        cmbJangbi2.Text = Trim(adoSet.Fields("Jangbi2").Value & "")
        cmbJangbi3.Text = Trim(adoSet.Fields("Jangbi3").Value & "")
        cmbJangbi4.Text = Trim(adoSet.Fields("Jangbi4").Value & "")
        cmbJangbi5.Text = Trim(adoSet.Fields("Jangbi5").Value & "")
        Call adoSetClose(adoSet)
    End If
    
    Return
    
        
Clear_SubObject:
    txtCodegu.Tag = txtCodegu.Text
    txtCodeName.Tag = txtCodeName.Text
    
    Call ClearForm(Me)
    
    txtCodegu.Text = txtCodegu.Tag
    txtCodeName.Text = txtCodeName.Tag
    
    Return
    
End Sub
