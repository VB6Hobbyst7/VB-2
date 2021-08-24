VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Begin VB.Form frmSample 
   Caption         =   "검체코드 관리"
   ClientHeight    =   7830
   ClientLeft      =   435
   ClientTop       =   780
   ClientWidth     =   11460
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7830
   ScaleWidth      =   11460
   WindowState     =   2  '최대화
   Begin FPSpreadADO.fpSpread sprSample 
      Height          =   5190
      Left            =   180
      TabIndex        =   23
      Top             =   2835
      Width           =   10860
      _Version        =   196608
      _ExtentX        =   19156
      _ExtentY        =   9155
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
      MaxCols         =   8
      MaxRows         =   200
      ScrollBars      =   2
      SpreadDesigner  =   "frmSample.frx":0000
      Appearance      =   1
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   1995
      Left            =   180
      TabIndex        =   0
      Top             =   60
      Width           =   10575
      _Version        =   65536
      _ExtentX        =   18653
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
      Begin Threed.SSCommand cmdClear 
         Height          =   1335
         Left            =   8310
         TabIndex        =   18
         Top             =   180
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   2355
         _StockProps     =   78
         Caption         =   "화면정리"
         BevelWidth      =   1
         Outline         =   0   'False
         Picture         =   "frmSample.frx":1A5A
      End
      Begin Threed.SSCommand cmdSeqno 
         Height          =   300
         Left            =   2040
         TabIndex        =   17
         Top             =   1470
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   529
         _StockProps     =   78
         Caption         =   "Maxno+1"
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin VB.TextBox txtAbbr 
         Height          =   315
         Left            =   3600
         TabIndex        =   8
         Top             =   1155
         Width           =   2115
      End
      Begin VB.TextBox txtCodenm 
         Height          =   315
         Left            =   2760
         TabIndex        =   7
         Top             =   825
         Width           =   2955
      End
      Begin VB.ComboBox cmbRegion 
         Height          =   300
         Left            =   1380
         TabIndex        =   6
         Top             =   1155
         Width           =   1395
      End
      Begin VB.ComboBox cmbClass1 
         Height          =   300
         ItemData        =   "frmSample.frx":31EC
         Left            =   1380
         List            =   "frmSample.frx":31EE
         TabIndex        =   5
         Top             =   180
         Width           =   2295
      End
      Begin VB.TextBox txtSeqno 
         Height          =   315
         Left            =   1380
         TabIndex        =   4
         Top             =   1455
         Width           =   615
      End
      Begin VB.TextBox txtCode 
         Height          =   315
         Left            =   1380
         TabIndex        =   3
         Top             =   825
         Width           =   1395
      End
      Begin VB.ComboBox cmbClass2 
         Height          =   300
         Left            =   1380
         TabIndex        =   2
         Top             =   495
         Width           =   2295
      End
      Begin VB.TextBox txtRowID 
         Appearance      =   0  '평면
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   270
         Left            =   3720
         TabIndex        =   1
         Top             =   180
         Width           =   2175
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   1335
         Left            =   7395
         TabIndex        =   9
         Top             =   180
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   2355
         _StockProps     =   78
         Caption         =   "삭제확인"
         BevelWidth      =   1
         Outline         =   0   'False
         Picture         =   "frmSample.frx":31F0
      End
      Begin Threed.SSCommand cmdInsert 
         Height          =   1335
         Left            =   6480
         TabIndex        =   10
         Top             =   180
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   2355
         _StockProps     =   78
         Caption         =   "입력확인"
         BevelWidth      =   1
         Outline         =   0   'False
         Picture         =   "frmSample.frx":3ACA
      End
      Begin VB.Label Label7 
         Caption         =   "부위"
         Height          =   255
         Left            =   180
         TabIndex        =   16
         Top             =   1185
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "코드/검체명"
         Height          =   255
         Left            =   180
         TabIndex        =   15
         Top             =   840
         Width           =   1035
      End
      Begin VB.Label Label2 
         Caption         =   "분류2"
         Height          =   195
         Left            =   180
         TabIndex        =   14
         Top             =   540
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "분류1"
         Height          =   195
         Left            =   180
         TabIndex        =   13
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label Label4 
         Caption         =   "약어"
         Height          =   195
         Left            =   3060
         TabIndex        =   12
         Top             =   1245
         Width           =   555
      End
      Begin VB.Label Label8 
         Caption         =   "일련번호"
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   1545
         Width           =   915
      End
   End
   Begin Threed.SSPanel panelTree 
      Height          =   7230
      Left            =   705
      TabIndex        =   21
      Top             =   90
      Visible         =   0   'False
      Width           =   5625
      _Version        =   65536
      _ExtentX        =   9922
      _ExtentY        =   12753
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
      BorderWidth     =   2
      BevelInner      =   1
      FloodColor      =   8421504
      Begin MSComctlLib.TreeView tvSample 
         Height          =   6915
         Left            =   120
         TabIndex        =   22
         Top             =   165
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   12197
         _Version        =   393217
         LineStyle       =   1
         Style           =   7
         Appearance      =   1
      End
   End
   Begin MSForms.CommandButton cmdPrint 
      Height          =   495
      Left            =   1860
      TabIndex        =   20
      Top             =   2160
      Width           =   1635
      Caption         =   "Print"
      PicturePosition =   327683
      Size            =   "2884;873"
      Picture         =   "frmSample.frx":528C
      FontName        =   "굴림"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdQry 
      Height          =   495
      Left            =   180
      TabIndex        =   19
      Top             =   2160
      Width           =   1635
      Caption         =   "조회확인"
      PicturePosition =   327683
      Size            =   "2884;873"
      Picture         =   "frmSample.frx":5B66
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin VB.Menu mnuQuit 
      Caption         =   "Quit"
   End
   Begin VB.Menu mnuTree 
      Caption         =   "TreeView(Open)"
   End
End
Attribute VB_Name = "frmSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbClass1_Click()
    
    If cmbClass1.ListIndex = -1 Then Exit Sub
    
    cmbClass2.ListIndex = -1
    txtCode.Text = ""
    txtCodenm.Text = ""
    cmbRegion.ListIndex = -1
    txtAbbr.Text = ""
    txtSeqno.Text = ""
    
    
End Sub

Private Sub cmdClear_Click()
    
    cmbClass1.ListIndex = -1
    cmbClass2.ListIndex = -1
    txtCode.Text = ""
    txtCodenm.Text = ""
    cmbRegion.ListIndex = -1
    txtAbbr.Text = ""
    txtSeqno.Text = ""
    txtRowID.Text = ""
    
    If sprSample.DataRowCnt = 0 Then Exit Sub
    
    If vbNo = MsgBox("아래 Spread 의 Data 까지 Clear 하시겠습니까?", _
                      vbYesNo + vbQuestion, _
                     "Spread Reset?") Then Exit Sub
                     
    sprSample.Row = 1
    sprSample.Row2 = sprSample.DataRowCnt
    sprSample.Col = 1
    sprSample.Col2 = sprSample.DataColCnt
    sprSample.BlockMode = True
    sprSample.Action = ActionClear
    sprSample.BlockMode = False
    
    
End Sub

Private Sub cmdDelete_Click()
    
    If cmbClass1.ListIndex = -1 Then Exit Sub
    
    
    strSql = ""
    strSql = strSql & " SELECT *"
    strSql = strSql & " FROM   TWEXAM_Sample"
    strSql = strSql & " WHERE  RowID = '" & txtRowID.Text & "'"
    
    If False = adoSetOpen(strSql, adoSet) Then
        MsgBox "이미 DataBase 에 존재하지 않는 Data 입니다"
        Exit Sub: End If
    
    
    If vbNo = MsgBox("선택하신 검체 Data 를 삭제하시겠습니까?", _
                      vbYesNo + vbQuestion + vbDefaultButton2, _
                     "삭제 확인 MessageBox") Then Exit Sub
    
    strSql = ""
    strSql = strSql & " DELETE"
    strSql = strSql & " FROM   TWEXAM_Sample"
    strSql = strSql & " WHERE  RowID  =  '" & txtRowID.Text & "'"
    Call adoExec(strSql)
    
    cmbClass1.ListIndex = -1
    cmbClass2.ListIndex = -1
    txtCode.Text = ""
    txtCodenm.Text = ""
    cmbRegion.ListIndex = -1
    txtAbbr.Text = ""
    txtSeqno.Text = ""
    txtRowID.Text = ""

    
End Sub


Private Sub cmdInsert_Click()

    If cmbClass1.ListIndex = -1 Then Exit Sub
    
    
    If Trim(txtRowID.Text) = "" Then
        GoSub Sampling_Insert
    Else
        GoSub Sampling_Update
    End If
    Exit Sub
    
Sampling_Insert:
    strSql = ""
    strSql = strSql & " INSERT "
    strSql = strSql & " INTO   TWEXAM_Sample"
    strSql = strSql & "       (Class1, Class2, Code, Codenm, Anatomy, Abbr, Seqno)"
    strSql = strSql & " VALUES('" & Left(Me.cmbClass1.Text, 1) & "',"
    strSql = strSql & "        '" & cmbClass2.Text & "',"
    strSql = strSql & "        '" & Trim(txtCode.Text) & "',"
    strSql = strSql & "        '" & Quot_Conv(Trim(txtCodenm.Text)) & "',"
    strSql = strSql & "        '" & cmbRegion.Text & "',"
    strSql = strSql & "        '" & Trim(txtAbbr.Text) & "',"
    strSql = strSql & "         " & Val(txtSeqno.Text) & ")"
    Call adoExec(strSql)
    
    GoSub Clear_Form_a
    Return
    
Sampling_Update:
    strSql = ""
    strSql = strSql & " UPDATE TWEXAM_Sample"
    strSql = strSql & " SET    Class1  =  '" & Left(cmbClass1.Text, 1) & "',"
    strSql = strSql & "        Class2  =  '" & cmbClass2.Text & "',"
    strSql = strSql & "        Code    =  '" & Trim(txtCode.Text) & "',"
    strSql = strSql & "        Codenm  =  '" & Quot_Conv(Trim(txtCodenm.Text)) & "',"
    strSql = strSql & "        Anatomy =  '" & cmbRegion.Text & "',"
    strSql = strSql & "        Abbr    =  '" & Trim(txtAbbr.Text) & "',"
    strSql = strSql & "        Seqno   =   " & Val(txtSeqno.Text)
    strSql = strSql & " WHERE  RowID   =  '" & txtRowID.Text & "'"
    Call adoExec(strSql)
    
    GoSub Clear_Form_a
    Return
    
    
    
Clear_Form_a:
    cmbClass1.ListIndex = -1
    cmbClass2.ListIndex = -1
    txtCode.Text = ""
    txtCodenm.Text = ""
    cmbRegion.ListIndex = -1
    txtAbbr.Text = ""
    txtSeqno.Text = ""
    txtRowID.Text = ""
    Return
    
    
End Sub

Private Sub cmdPrint_Click()
    Dim strFont1        As String
    Dim strFont2        As String
    Dim strHead1        As String
    Dim strHead2        As String
    
    If sprSample.DataRowCnt = 0 Then Exit Sub
    
    If vbNo = MsgBox("Spread 의 Data 를 Print 하시겠습니까?", _
                      vbYesNo + vbQuestion, _
                      "Printing continue...") Then Exit Sub
                      
    strFont1 = "/fn""굴림체"" /fz""16"" /fb1 /fi0 /fu0 /fk0 /fs1"
    strFont2 = "/fn""굴림체"" /fz""10"" /fb0 /fi0 /fu0 /fk0 /fs2"
    strHead1 = "/f1" & "/c" & "임상병리과 검체코드 List"
    strHead2 = "/f2" & " PAGE : " & "/p" & " of " & sprSample.PrintPageCount
    
    sprSample.PrintHeader = strFont1 + strHead1 + "/n/n" + strFont2 + strHead2 + "/n" + strFont2
    sprSample.PrintMarginLeft = 300
    sprSample.PrintMarginRight = 0
    sprSample.PrintMarginTop = 150
    sprSample.PrintMarginBottom = 500
    sprSample.PrintColHeaders = True
    sprSample.PrintRowHeaders = True
    sprSample.PrintBorder = True
    sprSample.PrintColor = False
    sprSample.PrintGrid = True
    sprSample.PrintShadows = True
    sprSample.PrintUseDataMax = False
    sprSample.Row = 1
    sprSample.Row2 = sprSample.DataRowCnt
    sprSample.Col = 1
    sprSample.Col2 = sprSample.MaxCols
    sprSample.PrintType = PrintTypeCellRange
    sprSample.Action = ActionPrint
    
    

End Sub

Private Sub cmdQry_Click()
    
    strSql = ""
    strSql = strSql & " SELECT a.* , a.RowID"
    strSql = strSql & " FROM   TWEXAM_Sample a"
    
    If cmbClass1.ListIndex > -1 Then
        strSql = strSql & " WHERE CLass1 = '" & Left(cmbClass1.Text, 1) & "'"
    End If
    
    strSql = strSql & " Order  By a.Code, a.Class1, a.Seqno"
    
    sprSample.MaxRows = 0
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    sprSample.MaxRows = adoSet.RecordCount
    sprSample.RowHeight(-1) = 10.23
    
    Do Until adoSet.EOF
        sprSample.Row = sprSample.DataRowCnt + 1
        sprSample.Col = 1: sprSample.Text = Trim(adoSet.Fields("RowID").Value & "")
        sprSample.Col = 2: sprSample.Text = Trim(adoSet.Fields("Class1").Value & "")
        sprSample.Col = 3: sprSample.Text = Trim(adoSet.Fields("Class2").Value & "")
        sprSample.Col = 4: sprSample.Text = Trim(adoSet.Fields("Code").Value & "")
        sprSample.Col = 5: sprSample.Text = Trim(adoSet.Fields("Codenm").Value & "")
        sprSample.Col = 6: sprSample.Text = Trim(adoSet.Fields("Abbr").Value & "")
        sprSample.Col = 7: sprSample.Text = Trim(adoSet.Fields("anatomy").Value & "")
        sprSample.Col = 8: sprSample.Text = Format(Trim(adoSet.Fields("Seqno").Value & ""), "000")
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)

End Sub


Private Sub cmdSeqno_Click()
    
    strSql = " SELECT NVL(MAX(Seqno), 0) + 1 Sqno FROM TWEXAM_Sample"
    
    If False = adoSetOpen(strSql, adoSet) Then
        txtSeqno.Text = "1": Exit Sub: End If
        
    txtSeqno.Text = adoSet.Fields("Sqno").Value & ""
    Call adoSetClose(adoSet)
    
    
End Sub


Private Sub Form_Load()
    
    
    Me.sprSample.RowHeight(-1) = 11.5
    GoSub Sample_Gubun_Code_Set
    GoSub Sample_Region_Code_Set
    GoSub Sample_Class2_Code_Set
    Exit Sub
    
    
Sample_Gubun_Code_Set:
    cmbClass1.AddItem "a. 일반검체"
    cmbClass1.AddItem "m. 미생물 검체"
    Return
    
Sample_Region_Code_Set:
    strSql = " SELECT anatomy From TWEXAM_Sample Group By anatomy"
    If False = adoSetOpen(strSql, adoSet) Then Return
    cmbRegion.Clear
    
    Do Until adoSet.EOF
        cmbRegion.AddItem Trim(adoSet.Fields("Anatomy").Value & "")
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    Return

Sample_Class2_Code_Set:
    strSql = " SELECT Class2 FROM TWEXAM_Sample GROUP BY Class2"
    If False = adoSetOpen(strSql, adoSet) Then Return
    cmbClass2.Clear
    
    Do Until adoSet.EOF
        cmbClass2.AddItem Trim(adoSet.Fields("Class2").Value & "")
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    Return
End Sub

Private Sub mnuQuit_Click()
    Unload Me
    
End Sub

Private Sub mnuTree_Click()
    
    Dim sText       As String
    Dim sRowid      As String
    Dim NodeX       As Node
    
    If mnuTree.Tag = "OPEN" Then
        mnuTree.Caption = "TreeView(Open)"
        mnuTree.Tag = ""
        panelTree.Visible = False
        Exit Sub
    Else
        mnuTree.Caption = "TreeView(Close)"
        mnuTree.Tag = "OPEN"
        panelTree.Visible = True
        panelTree.ZOrder 0
    End If
    
    DoEvents
    GoSub TreeView_Select
    Exit Sub
    
    
TreeView_Select:
    tvSample.Nodes.Clear
    Set NodeX = tvSample.Nodes.Add(, , "A0", "검체 분류")
    
    strSql = ""
    strSql = strSql & " SELECT Class2, Max(RowID) RWID"
    strSql = strSql & " FROM   TWEXAM_Sample"
    strSql = strSql & " GROUP  BY Class2"
    
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    
    Do Until adoSet.EOF
        sRowid = adoSet.Fields("RwID").Value & ""
        sText = Trim(adoSet.Fields("Class2").Value & "")
        Set NodeX = tvSample.Nodes.Add("A0", tvwChild, "A1" & sRowid, sText)
        GoSub Load_SubCode
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    tvSample.Nodes("A0").Expanded = True
    Return
    
Load_SubCode:
    Dim adoSubCode1     As ADODB.Recordset
    Dim sSubText1       As String
    Dim sSubText2       As String
    
    strSql = ""
    strSql = strSql & " SELECT class1, Codenm, anatomy, RowID"
    strSql = strSql & " FROM   TWEXAM_Sample"
    strSql = strSql & " WHERE  Class2 = '" & tvSample.Nodes("A1" & sRowid).Text & "'"
    
    If False = adoSetOpen(strSql, adoSubCode1) Then Return
    
    Do Until adoSubCode1.EOF
        sSubText1 = "B2" & adoSubCode1.Fields("RowID")
        
        sSubText2 = Trim(adoSubCode1.Fields("Class1").Value & ".") & _
                   StrConv(Trim(adoSubCode1.Fields("Codenm").Value & ""), vbProperCase)
        If Trim(adoSubCode1.Fields("Anatomy").Value & "") <> "" Then
            sSubText2 = sSubText2 & "(" & adoSubCode1.Fields("anatomy").Value & ")"
        End If
        
        Set NodeX = tvSample.Nodes.Add("A1" & sRowid, tvwChild, sSubText1, sSubText2)
        adoSubCode1.MoveNext
    Loop
    Call adoSetClose(adoSubCode1)
    Return
    
End Sub

Private Sub sprSample_DblClick(ByVal Col As Long, ByVal Row As Long)
    
    If Row = 0 Then
        GoSub Spread_Sort_Sub
        Exit Sub
    End If
    
    sprSample.Row = Row
    sprSample.Col = 1: txtRowID.Text = sprSample.Text
    sprSample.Col = 2: GoSub ComboClass1_Set
    sprSample.Col = 3: Call SetComboBox(cmbClass2, sprSample.Text)
    sprSample.Col = 4: Call SetComboBox(cmbRegion, sprSample.Text)
    
    sprSample.Col = 4: txtCode.Text = sprSample.Text
    sprSample.Col = 5: txtCodenm.Text = sprSample.Text
    sprSample.Col = 6: txtAbbr.Text = sprSample.Text
    sprSample.Col = 8: txtSeqno.Text = sprSample.Text
    Exit Sub
    
ComboClass1_Set:
    For I = 0 To cmbClass1.ListCount - 1
        If Left(Trim(cmbClass1.List(I)), 1) = Trim(sprSample.Text) Then
            cmbClass1.ListIndex = I
            Exit For
        End If
    Next
    Return

Spread_Sort_Sub:
    sprSample.Col = 1: sprSample.Col2 = sprSample.DataColCnt
    sprSample.Row = 1: sprSample.Row2 = sprSample.DataRowCnt
    
    sprSample.SortBy = SS_SORT_BY_ROW
    sprSample.SortKey(1) = Col
    If sprSample.SortKeyOrder(1) = SortKeyOrderDescending Then
        sprSample.SortKeyOrder(1) = SS_SORT_ORDER_ASCENDING
    Else
        sprSample.SortKeyOrder(1) = SortKeyOrderDescending
    End If
    sprSample.Action = SS_ACTION_SORT

    Return
End Sub


