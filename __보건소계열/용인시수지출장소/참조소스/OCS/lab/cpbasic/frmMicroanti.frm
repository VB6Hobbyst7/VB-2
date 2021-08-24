VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Begin VB.Form frmMicroanti 
   Caption         =   "미생물 반응약제 List 관리"
   ClientHeight    =   5700
   ClientLeft      =   1500
   ClientTop       =   1545
   ClientWidth     =   9555
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5700
   ScaleWidth      =   9555
   WindowState     =   2  '최대화
   Begin Threed.SSPanel SSPanel1 
      Height          =   1995
      Left            =   120
      TabIndex        =   10
      Top             =   180
      Width           =   11715
      _Version        =   65536
      _ExtentX        =   20664
      _ExtentY        =   3519
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelInner      =   1
      Begin VB.ComboBox cmbAntiGroup 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1800
         TabIndex        =   31
         Top             =   1140
         Width           =   855
      End
      Begin VB.TextBox txtRowID 
         Appearance      =   0  '평면
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   270
         Left            =   2940
         TabIndex        =   30
         Top             =   180
         Width           =   2295
      End
      Begin VB.TextBox txtOrgname 
         Height          =   315
         Left            =   1800
         TabIndex        =   26
         Top             =   832
         Width           =   1695
      End
      Begin VB.ComboBox cmbGram 
         Height          =   300
         Left            =   7920
         Style           =   2  '드롭다운 목록
         TabIndex        =   23
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtSource 
         Height          =   315
         Left            =   7920
         TabIndex        =   8
         Top             =   1140
         Width           =   855
      End
      Begin VB.TextBox txtHimic 
         Height          =   315
         Left            =   5580
         TabIndex        =   7
         Top             =   1470
         Width           =   1275
      End
      Begin VB.TextBox txtLomic 
         Height          =   315
         Left            =   5580
         TabIndex        =   6
         Top             =   1140
         Width           =   1275
      End
      Begin VB.TextBox txtHizone 
         Height          =   315
         Left            =   5580
         TabIndex        =   5
         Top             =   810
         Width           =   1275
      End
      Begin VB.TextBox txtLozone 
         Height          =   315
         Left            =   5580
         TabIndex        =   4
         Top             =   480
         Width           =   1275
      End
      Begin VB.TextBox txtPotency 
         Height          =   315
         Left            =   1800
         TabIndex        =   3
         Top             =   1470
         Width           =   1515
      End
      Begin VB.TextBox txtSeqno 
         Height          =   315
         Left            =   7920
         TabIndex        =   2
         Top             =   510
         Width           =   855
      End
      Begin VB.TextBox txtAntiName 
         Height          =   315
         Left            =   1800
         TabIndex        =   1
         Top             =   506
         Width           =   2595
      End
      Begin VB.TextBox txtAntiCode 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   1800
         TabIndex        =   0
         Top             =   180
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "AntiGroup"
         Height          =   315
         Left            =   840
         TabIndex        =   27
         Top             =   1170
         Width           =   915
      End
      Begin VB.Label Label11 
         Caption         =   "보조명"
         Height          =   315
         Left            =   840
         TabIndex        =   25
         Top             =   860
         Width           =   915
      End
      Begin VB.Label labelGram 
         Caption         =   "GramData"
         Height          =   255
         Left            =   7080
         TabIndex        =   24
         Top             =   900
         Width           =   735
      End
      Begin MSForms.CommandButton cmdClear 
         Height          =   495
         Left            =   9600
         TabIndex        =   22
         Top             =   1320
         Width           =   1575
         Caption         =   "화면정리"
         PicturePosition =   327683
         Size            =   "2778;873"
         Picture         =   "frmMicroanti.frx":0000
         FontName        =   "굴림체"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdDelete 
         Height          =   495
         Left            =   9600
         TabIndex        =   21
         Top             =   780
         Width           =   1575
         Caption         =   "삭제확인"
         PicturePosition =   327683
         Size            =   "2778;873"
         Picture         =   "frmMicroanti.frx":1792
         FontName        =   "굴림체"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdInsert 
         Height          =   495
         Left            =   9600
         TabIndex        =   20
         Top             =   240
         Width           =   1575
         Caption         =   "입력확인"
         PicturePosition =   327683
         Size            =   "2778;873"
         Picture         =   "frmMicroanti.frx":206C
         FontName        =   "굴림체"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin VB.Label Label10 
         Caption         =   "Source"
         Height          =   195
         Left            =   7080
         TabIndex        =   19
         Top             =   1200
         Width           =   675
      End
      Begin VB.Label Label9 
         Caption         =   "HiMic"
         Height          =   315
         Left            =   4620
         TabIndex        =   18
         Top             =   1530
         Width           =   915
      End
      Begin VB.Label Label8 
         Caption         =   "LoMic"
         Height          =   315
         Left            =   4620
         TabIndex        =   17
         Top             =   1200
         Width           =   915
      End
      Begin VB.Label Label7 
         Caption         =   "HiZone"
         Height          =   315
         Left            =   4620
         TabIndex        =   16
         Top             =   870
         Width           =   915
      End
      Begin VB.Label Label6 
         Caption         =   "LoZone"
         Height          =   315
         Left            =   4620
         TabIndex        =   15
         Top             =   540
         Width           =   915
      End
      Begin VB.Label Label5 
         Caption         =   "Potency"
         Height          =   315
         Left            =   840
         TabIndex        =   14
         Top             =   1500
         Width           =   915
      End
      Begin VB.Label Label3 
         Caption         =   "일련번호"
         Height          =   195
         Left            =   7080
         TabIndex        =   13
         Top             =   585
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "약제명"
         Height          =   315
         Left            =   840
         TabIndex        =   12
         Top             =   550
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "약제코드"
         Height          =   315
         Left            =   840
         TabIndex        =   11
         Top             =   240
         Width           =   915
      End
   End
   Begin FPSpreadADO.fpSpread ssAntiList 
      Height          =   5235
      Left            =   120
      TabIndex        =   9
      Top             =   2700
      Width           =   11715
      _Version        =   196608
      _ExtentX        =   20664
      _ExtentY        =   9234
      _StockProps     =   64
      BackColorStyle  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   13
      ScrollBars      =   2
      SpreadDesigner  =   "frmMicroanti.frx":382E
      Appearance      =   1
   End
   Begin MSForms.CommandButton cmdPrint 
      Height          =   435
      Left            =   1920
      TabIndex        =   29
      Top             =   2220
      Width           =   1755
      Caption         =   "Print"
      PicturePosition =   327683
      Size            =   "3096;767"
      Picture         =   "frmMicroanti.frx":765D
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdQryAll 
      Height          =   435
      Left            =   120
      TabIndex        =   28
      Top             =   2220
      Width           =   1755
      Caption         =   " 조회(aLL)"
      PicturePosition =   327683
      Size            =   "3096;767"
      Picture         =   "frmMicroanti.frx":7F37
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin VB.Menu mnuQuit 
      Caption         =   "Quit"
   End
End
Attribute VB_Name = "frmMicroanti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClear_Click()
    
    For i = 0 To Me.Count - 1
        If TypeOf Me.Controls(i) Is TextBox Then
            Me.Controls(i).Text = ""
        ElseIf TypeOf Me.Controls(i) Is ComboBox Then
            Me.Controls(i).ListIndex = -1
        End If
    Next
    
End Sub

Private Sub cmdDelete_Click()
    
    If Trim(txtAntiCode.Text) = "" Then Exit Sub
    
    If vbNo = MsgBox("선택하신 약제코드를 삭제하시겠습니까?", _
                      vbYesNo + vbQuestion, _
                      Trim(txtAntiCode.Text) & "." & StrConv(Trim(txtAntiName.Text), vbProperCase)) Then Exit Sub
    
    
    strSql = " DELETE FROM TWEXAM_ANTILIST WHERE Codeky = '" & Trim(txtAntiCode.Text) & "'"
    
    If adoExec(strSql) Then
        MsgBox "삭제되었습니다!"
    Else
        MsgBox "어떤 이유로 인하여 삭제되지 않았습니다!.."
    End If
    
End Sub

Private Sub cmdInsert_Click()
    Dim sRowid      As String
    
    sRowid = Trim(txtRowID.Text)
    
    If sRowid = "" Then
        GoSub antiList_Insert
    Else
        GoSub antiList_Update
    End If
    
    Exit Sub
    
antiList_Insert:
    strSql = ""
    strSql = strSql & " INSERT"
    strSql = strSql & " INTO   TWEXAM_ANTILIST"
    strSql = strSql & "       (Codeky, CodeNm, Orgname, Antigroup, Potency, Lozone, Hizone,"
    strSql = strSql & "        Lomic,  Himic,  Seqno,   Gram,      Source )"
    strSql = strSql & " VALUES('" & Trim(txtAntiCode.Text) & "',"
    strSql = strSql & "        '" & Quot_Conv(Trim(txtAntiName.Text)) & "',"
    strSql = strSql & "        '" & Quot_Conv(Trim(txtOrgname.Text)) & "',"
    strSql = strSql & "        '" & Trim(cmbAntiGroup.Text) & "',"
    strSql = strSql & "        '" & Trim(txtPotency.Text) & "',"
    strSql = strSql & "         " & Val(txtLozone.Text) & ","
    strSql = strSql & "         " & Val(txtHizone.Text) & ","
    strSql = strSql & "         " & Val(txtLomic.Text) & ","
    strSql = strSql & "         " & Val(txtHimic.Text) & ","
    strSql = strSql & "         " & Val(txtSeqno.Text) & ","
    strSql = strSql & "        '" & cmbGram.Text & "',"
    strSql = strSql & "        '" & txtSource.Text & "')"
    
    If adoExec(strSql) Then
        Call cmdClear_Click
    End If
    Return
    
antiList_Update:
    strSql = ""
    strSql = strSql & " UPDATE TWEXAM_ANTILIST"
    strSql = strSql & " SET    Codeky    = '" & Trim(txtAntiCode.Text) & "',"
    strSql = strSql & "        Codenm    = '" & Trim(txtAntiName.Text) & "',"
    strSql = strSql & "        Orgname   = '" & Trim(txtOrgname.Text) & "',"
    strSql = strSql & "        AntiGroup = '" & Trim(cmbAntiGroup.Text) & "',"
    strSql = strSql & "        Potency   = '" & Trim(txtPotency.Text) & "',"
    strSql = strSql & "        Lozone    =  " & Val(txtLozone.Text) & ","
    strSql = strSql & "        Hizone    =  " & Val(txtHizone.Text) & ","
    strSql = strSql & "        Lomic     =  " & Val(txtLomic.Text) & ","
    strSql = strSql & "        Himic     =  " & Val(txtHimic.Text) & ","
    strSql = strSql & "        Seqno     =  " & Val(txtSeqno.Text) & ","
    strSql = strSql & "        Gram      = '" & Trim(cmbGram.Text) & "',"
    strSql = strSql & "        Source    = '" & Trim(txtSource.Text) & "'"
    strSql = strSql & " WHERE  ROWID     =  '" & sRowid & "'"
    If adoExec(strSql) Then
        Call cmdClear_Click
    End If
    Return
    
End Sub

Private Sub cmdPrint_Click()
    
    If ssAntiList.DataRowCnt = 0 Then Exit Sub
    If vbNo = MsgBox("반응 약제 Data 의 Print 작업을 하시겠습니까?", _
                     vbYesNo + vbQuestion, _
                     "출력 작업 확인MessageBox") Then Exit Sub
    
    Dim strFont(1)        As String
    Dim strHead(1)        As String
    
    strFont(0) = "/fn""굴림체"" /fz""16"" /fb1 /fi0 /fu0 /fk0 /fs1"
    strFont(1) = "/fn""굴림체"" /fz""10"" /fb0 /fi0 /fu0 /fk0 /fs2"
    strHead(0) = "/f1" & "/c" & " 미생물 반응 약제 Data "
    'strHead(1) = "/f2" & "Page : " & "/p" & " of " & ssAntiList.PrintPageCount & "/r"
    
    ssAntiList.PrintHeader = strFont(0) + strHead(0) + "/n/n" + strFont(1) + strHead(1) + "/n" + strFont(1)
    ssAntiList.PrintFooter = "/f2" & "/c" & "Page : " & "/p" & " of " & ssAntiList.PrintPageCount
    ssAntiList.PrintMarginLeft = 300
    ssAntiList.PrintMarginRight = 100
    ssAntiList.PrintMarginTop = 250
    ssAntiList.PrintMarginBottom = 400
    ssAntiList.PrintColHeaders = True
    ssAntiList.PrintRowHeaders = True
    ssAntiList.PrintBorder = True
    ssAntiList.PrintColor = False
    ssAntiList.PrintGrid = True
    ssAntiList.PrintShadows = True
    ssAntiList.PrintUseDataMax = False
    ssAntiList.Row = 1
    ssAntiList.Row2 = ssAntiList.DataRowCnt
    ssAntiList.Col = 1
    ssAntiList.Col2 = ssAntiList.MaxCols
    ssAntiList.PrintType = PrintTypeCellRange
    ssAntiList.Action = ActionPrint
            
End Sub

Private Sub cmdQryAll_Click()
    
    strSql = ""
    strSql = strSql & " SELECT a.*, a.RowID"
    strSql = strSql & " FROM   TWEXAM_ANTILIST a"
    strSql = strSql & " ORDER  By Seqno"
    
    ssAntiList.MaxRows = 0
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    ssAntiList.MaxRows = adoSet.RecordCount
    
    Do Until adoSet.EOF
        ssAntiList.Row = ssAntiList.DataRowCnt + 1
        ssAntiList.Col = 1:  ssAntiList.Text = adoSet.Fields("Codeky").Value & ""
        ssAntiList.Col = 2:  ssAntiList.Text = Trim(adoSet.Fields("Codenm").Value & "")
        ssAntiList.Col = 3:  ssAntiList.Text = Trim(adoSet.Fields("Orgname").Value & "")
        ssAntiList.Col = 4:  ssAntiList.Text = Trim(adoSet.Fields("AntiGroup").Value & "")
        ssAntiList.Col = 5:  ssAntiList.Text = Trim(adoSet.Fields("Potency").Value & "")
        
        ssAntiList.Col = 6:  ssAntiList.Text = adoSet.Fields("LoZone").Value & ""
        ssAntiList.Col = 7:  ssAntiList.Text = adoSet.Fields("HiZone").Value & ""
        ssAntiList.Col = 8:  ssAntiList.Text = adoSet.Fields("LoMic").Value & ""
        ssAntiList.Col = 9:  ssAntiList.Text = adoSet.Fields("HiMic").Value & ""
        ssAntiList.Col = 10: ssAntiList.Text = adoSet.Fields("Seqno").Value & ""
        ssAntiList.Col = 11: ssAntiList.Text = adoSet.Fields("Gram").Value & ""
        ssAntiList.Col = 12: ssAntiList.Text = Trim(adoSet.Fields("Source").Value & "")
        ssAntiList.Col = 13: ssAntiList.Text = adoSet.Fields("RowID").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    
End Sub

Private Sub cmdQryAll1_Click()

End Sub

Private Sub CommandButton1_Click()

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
    End If

End Sub

Private Sub Form_Load()
    
    ssAntiList.RowHeight(-1) = 12
    GoSub cmbGram_Setting
    GoSub ANTI_GroupCode_Get
    Exit Sub
    
'/________________________________________
cmbGram_Setting:
    cmbGram.AddItem "+"
    cmbGram.AddItem "-"
    cmbGram.AddItem "a"
    cmbGram.AddItem "b"
    cmbGram.AddItem "f"
    cmbGram.AddItem "m"
    cmbGram.AddItem "o"
    cmbGram.AddItem "w"
    cmbGram.AddItem ""
    Return
    
    
ANTI_GroupCode_Get:
    strSql = ""
    strSql = strSql & " SELECT org_AntiGr"
    strSql = strSql & " FROM   TWEXAM_ORGLIST"
    strSql = strSql & " GROUP  BY Org_AntiGr"
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    Do Until adoSet.EOF
        cmbAntiGroup.AddItem Trim(adoSet.Fields("org_antiGr").Value & "")
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    Return

End Sub

Private Sub mnuQuit_Click()
    Unload Me
    
End Sub

Private Sub ssAntiList_DblClick(ByVal Col As Long, ByVal Row As Long)
    
    If Row = 0 Then
        If Col = 0 Then Exit Sub
        
        ssAntiList.SortBy = SS_SORT_BY_ROW
        ssAntiList.SortKey(1) = Col
        If ssAntiList.SortKeyOrder(1) = SortKeyOrderDescending Then
            ssAntiList.SortKeyOrder(1) = SS_SORT_ORDER_ASCENDING
        Else
            ssAntiList.SortKeyOrder(1) = SortKeyOrderDescending
        End If
        ssAntiList.Col = 1: ssAntiList.Col2 = ssAntiList.DataColCnt
        ssAntiList.Row = 1: ssAntiList.Row2 = ssAntiList.DataRowCnt
        ssAntiList.Action = SS_ACTION_SORT
    ElseIf Row > 0 Then
        If Col = 0 Then Exit Sub
        GoSub Moveup_Data
    End If
    Exit Sub
    
Moveup_Data:
    ssAntiList.Row = Row
    ssAntiList.Col = 1:  Me.txtAntiCode.Text = Trim(ssAntiList.Text)
    ssAntiList.Col = 2:  Me.txtAntiName.Text = Trim(ssAntiList.Text)
    ssAntiList.Col = 3:  Me.txtOrgname.Text = Trim(ssAntiList.Text)
    ssAntiList.Col = 4:  Call SetComboBox(cmbAntiGroup, Trim(ssAntiList.Text))
    'ssAntiList.Col = 4:  Me.txtAntiGroup.Text = Trim(ssAntiList.Text)
    ssAntiList.Col = 5:  Me.txtPotency.Text = Trim(ssAntiList.Text)
    ssAntiList.Col = 6:  Me.txtLozone.Text = Trim(ssAntiList.Text)
    ssAntiList.Col = 7:  Me.txtHizone.Text = Trim(ssAntiList.Text)
    ssAntiList.Col = 8:  Me.txtLomic.Text = Trim(ssAntiList.Text)
    ssAntiList.Col = 9:  Me.txtHimic.Text = Trim(ssAntiList.Text)
    ssAntiList.Col = 10: Me.txtSeqno.Text = Trim(ssAntiList.Text)
    ssAntiList.Col = 11: Call SetComboBox(Me.cmbGram, Trim(ssAntiList.Text))
    ssAntiList.Col = 12: Me.txtSource.Text = Trim(ssAntiList.Text)
    ssAntiList.Col = 13: Me.txtRowID.Text = Trim(ssAntiList.Text)
    
    Return
    
    
End Sub

Private Sub txtAntiCode_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        GoSub Get_AntiList_Select
    End If
    Exit Sub
    
Get_AntiList_Select:
    strSql = ""
    strSql = strSql & " SELECT a.*, a.RowID"
    strSql = strSql & " FROM   TWEXAM_ANTILIST a"
    strSql = strSql & " WHERE  a.Codeky  =  '" & txtAntiCode.Text & "'"
    strSql = strSql & " ORDER  By Seqno"
    
    ssAntiList.MaxRows = 0
    If False = adoSetOpen(strSql, adoSet) Then Return
    ssAntiList.MaxRows = adoSet.RecordCount
    
    Do Until adoSet.EOF
        ssAntiList.Row = ssAntiList.DataRowCnt + 1
        ssAntiList.Col = 1:  ssAntiList.Text = adoSet.Fields("Codeky").Value & ""
        ssAntiList.Col = 2:  ssAntiList.Text = Trim(adoSet.Fields("Codenm").Value & "")
        ssAntiList.Col = 3:  ssAntiList.Text = Trim(adoSet.Fields("Orgname").Value & "")
        ssAntiList.Col = 4:  ssAntiList.Text = Trim(adoSet.Fields("AntiGroup").Value & "")
        ssAntiList.Col = 5:  ssAntiList.Text = Trim(adoSet.Fields("Potency").Value & "")
        
        ssAntiList.Col = 6:  ssAntiList.Text = adoSet.Fields("LoZone").Value & ""
        ssAntiList.Col = 7:  ssAntiList.Text = adoSet.Fields("HiZone").Value & ""
        ssAntiList.Col = 8:  ssAntiList.Text = adoSet.Fields("LoMic").Value & ""
        ssAntiList.Col = 9:  ssAntiList.Text = adoSet.Fields("HiMic").Value & ""
        ssAntiList.Col = 10: ssAntiList.Text = adoSet.Fields("Seqno").Value & ""
        ssAntiList.Col = 11: ssAntiList.Text = adoSet.Fields("Gram").Value & ""
        ssAntiList.Col = 12: ssAntiList.Text = Trim(adoSet.Fields("Source").Value & "")
        ssAntiList.Col = 13: ssAntiList.Text = adoSet.Fields("RowID").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    Return
    
    
Get_Seqno_Plusone:
    Dim adoSeq      As ADODB.Recordset
    
    strSql = ""
    strSql = strSql & " SELECT MAX(Seqno) + 1 PlusSeqno"
    strSql = strSql & " FROM   TWEXAM_ANTILIST"
    If False = adoSetOpen(strSql, adoSeq) Then
        txtSeqno.Text = "0"
    Else
        txtSeqno.Text = adoSeq.Fields("PlusSeqno").Value & ""
        txtAntiName.SetFocus
        Call adoSetClose(adoSeq)
    End If
    
    Return

End Sub

