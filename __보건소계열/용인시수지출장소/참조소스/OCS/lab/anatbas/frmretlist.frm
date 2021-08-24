VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Begin VB.Form frmRetList 
   Caption         =   "결과Data 관리화면"
   ClientHeight    =   6825
   ClientLeft      =   135
   ClientTop       =   1335
   ClientWidth     =   11655
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   11655
   Begin Threed.SSPanel panelMicro 
      Height          =   6210
      Left            =   90
      TabIndex        =   10
      Top             =   1440
      Visible         =   0   'False
      Width           =   3645
      _Version        =   65536
      _ExtentX        =   6429
      _ExtentY        =   10954
      _StockProps     =   15
      Caption         =   "SSPanel3"
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
      Begin MSComctlLib.TreeView tvSample 
         Height          =   5940
         Left            =   45
         TabIndex        =   11
         Top             =   90
         Width           =   3480
         _ExtentX        =   6138
         _ExtentY        =   10478
         _Version        =   393217
         Indentation     =   653
         LineStyle       =   1
         Style           =   7
         BorderStyle     =   1
         Appearance      =   1
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   555
      Left            =   3780
      TabIndex        =   7
      Top             =   180
      Width           =   5415
      _Version        =   65536
      _ExtentX        =   9551
      _ExtentY        =   979
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
      Begin VB.TextBox txtItemCode 
         Appearance      =   0  '평면
         BackColor       =   &H00C0E0FF&
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   120
         Width           =   1215
      End
      Begin VB.TextBox txtItemName 
         Appearance      =   0  '평면
         BackColor       =   &H00C0E0FF&
         Height          =   315
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   120
         Width           =   3615
      End
   End
   Begin Threed.SSPanel panelItem 
      Height          =   6210
      Left            =   90
      TabIndex        =   4
      Top             =   180
      Width           =   3645
      _Version        =   65536
      _ExtentX        =   6429
      _ExtentY        =   10954
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
      Begin VB.TextBox txtSlipno 
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   120
         Width           =   435
      End
      Begin VB.TextBox txtSlipname 
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   120
         Width           =   2595
      End
      Begin VB.ListBox lstItemList 
         BackColor       =   &H00C0FFFF&
         Height          =   5460
         Left            =   135
         TabIndex        =   1
         Top             =   540
         Width           =   3315
      End
      Begin Threed.SSCommand cmdHelp 
         Height          =   315
         Left            =   600
         TabIndex        =   0
         ToolTipText     =   "Slip 종류를 조회선택할수 있습니다."
         Top             =   120
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   556
         _StockProps     =   78
         Caption         =   "&H"
      End
   End
   Begin FPSpreadADO.fpSpread ssRet 
      Height          =   5010
      Left            =   3780
      TabIndex        =   2
      Top             =   1395
      Width           =   7860
      _Version        =   196608
      _ExtentX        =   13864
      _ExtentY        =   8837
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
      RestrictCols    =   -1  'True
      ScrollBarExtMode=   -1  'True
      ScrollBars      =   2
      SpreadDesigner  =   "frmRetList.frx":0000
      VisibleCols     =   6
      VisibleRows     =   500
      Appearance      =   2
   End
   Begin MSForms.CommandButton cmdClear 
      Height          =   495
      Left            =   7020
      TabIndex        =   13
      Top             =   855
      Width           =   1635
      Caption         =   "화면정리"
      PicturePosition =   327683
      Size            =   "2884;873"
      Picture         =   "frmRetList.frx":2715
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdSave 
      Height          =   495
      Left            =   5400
      TabIndex        =   12
      Top             =   855
      Width           =   1635
      Caption         =   "Data 저장"
      PicturePosition =   327683
      Size            =   "2884;873"
      Picture         =   "frmRetList.frx":3EA7
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdAdd 
      Height          =   495
      Left            =   3780
      TabIndex        =   3
      Top             =   855
      Width           =   1635
      Caption         =   "Data 추가"
      PicturePosition =   327683
      Size            =   "2884;873"
      Picture         =   "frmRetList.frx":5669
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin VB.Menu mnuQuit 
      Caption         =   "Quit"
   End
   Begin VB.Menu mnuJob 
      Caption         =   "작업구분"
      Begin VB.Menu mnuItem 
         Caption         =   "ItemCode"
      End
      Begin VB.Menu mnuSample 
         Caption         =   "검체Code"
      End
   End
End
Attribute VB_Name = "frmRetList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
    Dim nTmpSeqno       As Integer
    
    If Trim(txtItemCode.Text) = "" Then
        MsgBox "선택된 ItemCode 가 없습니다", vbInformation
        Exit Sub
    End If
    
    ssRet.MaxRows = ssRet.MaxRows + 1
    
    
    ssRet.Row = ssRet.DataRowCnt + 1
    ssRet.Col = 3
    If panelMicro.Visible = True Then ssRet.TypeComboBoxCurSel = 1
    If panelItem.Visible = True Then ssRet.TypeComboBoxCurSel = 0
    
    ssRet.Col = 4
    ssRet.Text = txtSlipno.Text
    If panelMicro.Visible = True Then ssRet.Text = "42"
    
    ssRet.Col = 5
    ssRet.Text = txtItemCode.Text
    
    nTmpSeqno = 0
    For I = 1 To ssRet.DataRowCnt
        ssRet.Row = I
        ssRet.Col = 7
        If nTmpSeqno < Val(ssRet.Text) Then
            nTmpSeqno = Val(ssRet.Text)
        End If
    Next

    ssRet.Row = ssRet.DataRowCnt
    ssRet.Col = 7
    ssRet.Text = nTmpSeqno + 1
    
    ssRet.SetFocus
    ssRet.Row = ssRet.DataRowCnt
    ssRet.Col = 6
    ssRet.Action = SS_ACTION_ACTIVE_CELL
    
    
End Sub

Public Sub cmdClear_Click()
    
    ssRet.Row = 1
    ssRet.Row2 = ssRet.DataRowCnt
    ssRet.Col = 1
    ssRet.Col2 = ssRet.DataColCnt
    ssRet.BlockMode = True
    ssRet.Action = SS_ACTION_CLEAR_TEXT
    ssRet.BlockMode = False
    
End Sub

Private Sub cmdHelp_Click()
    Dim sCode       As String * 8
    
    mdiMain.stbMain.Panels(1).Text = ""
    
    gCallWin = 2
    
    frmSlipQry.Show vbModal
    If Trim(txtSlipname.Text) <> "" Then
        txtSlipno.SetFocus
    End If

    If txtSlipno.Text = "" Then Exit Sub
    
    StrSql = ""
    StrSql = StrSql & " SELECT Codeky, ItemNM"
    StrSql = StrSql & " FROM   TWEXAM_iTemML"
    StrSql = StrSql & " WHERE  Codeky  LIKE '" & txtSlipno.Text & "%'"
    StrSql = StrSql & " ORDER  BY Codeky"
    
    If False = adoSetOpen(StrSql, adoSet) Then Exit Sub
    lstItemList.Clear
    
    Do Until adoSet.EOF
        sCode = adoSet.Fields("Codeky").Value & ""
        lstItemList.AddItem sCode & " " & adoSet.Fields("ItemNM").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    lstItemList.SetFocus
    
    
End Sub

Private Sub cmdSave_Click()
    Dim sRowid      As String
    Dim bValue      As Boolean
    Dim sSlipno     As String
    Dim sItemCD     As String
    Dim sRet        As String
    Dim nPutSeq     As Integer
    Dim sNormal     As String
    Dim sGubun      As String
    Dim iSeq        As Integer

    
    
    Screen.MousePointer = vbHourglass
    GoSub Renumber_Set
    
    For I = 1 To ssRet.DataRowCnt
        ssRet.Row = I
        ssRet.Col = 1: sRowid = ssRet.Text
        ssRet.Col = 2:
        If ssRet.Value = True Then
            bValue = True
        Else
            bValue = False
        End If
        
        If sRowid <> "" And bValue = True Then
            GoSub RET_Delete_Sub
        ElseIf sRowid <> "" And bValue = False Then
            GoSub RET_Update_Sub
        ElseIf sRowid = "" And bValue = False Then
            GoSub RET_Insert_Sub
        End If
    Next
    
    GoSub ReRead_Sub
    Screen.MousePointer = vbDefault
    
    Exit Sub
    
'/-------------------------------------------------------
Renumber_Set:
    Dim nSeq        As Integer
    
    nSeq = 1
    For I = 1 To ssRet.DataRowCnt
        ssRet.Row = I
        ssRet.Col = 2
        If ssRet.Value = False Then
            ssRet.Col = 7
            ssRet.Text = nSeq
            nSeq = nSeq + 1
        End If
    Next
    Return

'@
RET_Delete_Sub:
    StrSql = ""
    StrSql = StrSql & " DELETE "
    StrSql = StrSql & " FROM    TWEXAM_Ret"
    StrSql = StrSql & " WHERE   RowID  =  '" & sRowid & "'"
    adoConnect.BeginTrans
    If adoExec(StrSql) Then
        adoConnect.CommitTrans
    Else
        adoConnect.RollbackTrans
    End If

    Return

'@
RET_Update_Sub:
    
    
    ssRet.Col = 6: sRet = Trim(ssRet.Text)
    ssRet.Col = 7: iSeq = Val(ssRet.Text)
    ssRet.Col = 8: sNormal = ssRet.Text
    
    StrSql = ""
    StrSql = StrSql & " UPDATE  TWEXAM_Ret"
    StrSql = StrSql & " SET     Ret    = '" & RTrim(sRet) & "',"
    StrSql = StrSql & "         Seqno  =  " & iSeq & ","
    StrSql = StrSql & "         Normal = '" & sNormal & "'"
    StrSql = StrSql & " WHERE   RowID  =  '" & sRowid & "'"
    adoConnect.BeginTrans
    If adoExec(StrSql) Then
        adoConnect.CommitTrans
    Else
        adoConnect.RollbackTrans
    End If
    Return

'@
RET_Insert_Sub:
    
    ssRet.Col = 3: sGubun = Left(ssRet.Text, 1)
    ssRet.Col = 4: sSlipno = Trim(ssRet.Text)
    ssRet.Col = 5: sItemCD = Trim(ssRet.Text)
    ssRet.Col = 6: sRet = RTrim(ssRet.Text)
    ssRet.Col = 7: nPutSeq = Val(ssRet.Text)
    ssRet.Col = 8: sNormal = ssRet.Text
    
    StrSql = ""
    StrSql = StrSql & " INSERT  INTO  TWEXAM_Ret"
    StrSql = StrSql & "       (RetGb, Slipno, ItemCd, Ret, Seqno, Normal)"
    StrSql = StrSql & " VALUES('" & sGubun & "',"
    StrSql = StrSql & "        '" & sSlipno & "',"
    StrSql = StrSql & "        '" & sItemCD & "',"
    StrSql = StrSql & "        '" & sRet & "',"
    StrSql = StrSql & "         " & nPutSeq & ","
    StrSql = StrSql & "        '" & sNormal & "')"
    
    adoConnect.BeginTrans
    If adoExec(StrSql) Then
        adoConnect.CommitTrans
    Else
        adoConnect.RollbackTrans
    End If
    
    Return

ReRead_Sub:

    Dim sCode       As String * 8
    
    If panelMicro.Visible = True Then
        If Left(tvSample.SelectedItem.Key, 2) = "B2" Then
            sCode = Left(tvSample.SelectedItem.Text, 8)
        Else
            Exit Sub
        End If
        txtItemCode.Text = sCode
        txtItemName.Text = Mid(Trim(tvSample.SelectedItem.Text), 10, Len(Trim(tvSample.SelectedItem.Text)) - 9)
        
        StrSql = ""
        StrSql = StrSql & " SELECT a.*, a.RowID"
        StrSql = StrSql & " FROM   TWEXAM_Ret a"
        StrSql = StrSql & " WHERE  a.Slipno = '42'"
        StrSql = StrSql & " AND    a.ItemCd = '" & sCode & "'"
        StrSql = StrSql & " ORDER  BY Seqno"
    Else
        StrSql = ""
        StrSql = StrSql & " SELECT a.*, a.RowID"
        StrSql = StrSql & " FROM   TWEXAM_Ret a"
        StrSql = StrSql & " WHERE  a.Slipno = '" & Trim(txtSlipno.Text) & "'"
        StrSql = StrSql & " AND    a.ItemCd = '" & Trim(txtItemCode.Text) & "'"
        StrSql = StrSql & " Order  by a.Seqno"
    End If
    
    ssRet.MaxRows = 0
    If False = adoSetOpen(StrSql, adoSet) Then Return
    ssRet.MaxRows = adoSet.RecordCount
    
    Do Until adoSet.EOF
        ssRet.Row = ssRet.DataRowCnt + 1
        ssRet.Col = 1: ssRet.Text = adoSet.Fields("RowID").Value & ""
        ssRet.Col = 2: ssRet.Value = False
        If adoSet.Fields("RetGB").Value & "" = "M" Then
            ssRet.Col = 3: ssRet.TypeComboBoxCurSel = 1
        Else
            ssRet.Col = 3: ssRet.TypeComboBoxCurSel = 0
        End If
        ssRet.Col = 4: ssRet.Text = adoSet.Fields("Slipno").Value & ""
        ssRet.Col = 5: ssRet.Text = adoSet.Fields("itemCD").Value & ""
        ssRet.Col = 6: ssRet.Text = adoSet.Fields("Ret").Value & ""
        ssRet.Col = 7: ssRet.Text = adoSet.Fields("Seqno").Value & ""
        ssRet.Col = 8: ssRet.Text = adoSet.Fields("Normal").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    Return
    
End Sub

Private Sub Form_Load()
    Me.Top = 1
    Me.Left = 1
    Me.Height = 7515
    Me.Width = 11775
    

End Sub

Private Sub lstItemList_DblClick()
    Dim sCode       As String * 8
    
    sCode = Left(lstItemList.Text, 8)
    txtItemCode.Text = sCode
    txtItemName.Text = Mid(Trim(lstItemList.Text), 10, Len(Trim(lstItemList.Text)) - 9)
    
    StrSql = ""
    StrSql = StrSql & " SELECT a.*, a.RowID"
    StrSql = StrSql & " FROM   TWEXAM_Ret a"
    StrSql = StrSql & " WHERE  a.Slipno = '" & Trim(txtSlipno.Text) & "'"
    StrSql = StrSql & " AND    a.ItemCd = '" & sCode & "'"
    StrSql = StrSql & " ORDER  BY Seqno"
    
    ssRet.MaxRows = 0
    If False = adoSetOpen(StrSql, adoSet) Then Exit Sub
    ssRet.MaxRows = adoSet.RecordCount
    
    Do Until adoSet.EOF
        ssRet.Row = ssRet.DataRowCnt + 1
        ssRet.Col = 1: ssRet.Text = adoSet.Fields("RowID").Value & ""
        ssRet.Col = 2: ssRet.Value = False
        If adoSet.Fields("RetGB").Value & "" = "M" Then
            ssRet.Col = 3: ssRet.TypeComboBoxCurSel = 1
        Else
            ssRet.Col = 3: ssRet.TypeComboBoxCurSel = 0
        End If
        ssRet.Col = 4: ssRet.Text = adoSet.Fields("Slipno").Value & ""
        ssRet.Col = 5: ssRet.Text = adoSet.Fields("itemCD").Value & ""
        ssRet.Col = 6: ssRet.Text = adoSet.Fields("Ret").Value & ""
        ssRet.Col = 7: ssRet.Text = adoSet.Fields("Seqno").Value & ""
        ssRet.Col = 8: ssRet.Text = adoSet.Fields("Normal").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
End Sub

Private Sub lstItemList_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        Call lstItemList_DblClick
    End If
    
End Sub

Private Sub mnuItem_Click()
    
    DoEvents
    Call cmdClear_Click
    panelMicro.Visible = False
    panelItem.Visible = True
    panelItem.Left = 45
    panelItem.Top = 180
    panelItem.ZOrder 0
    
End Sub

Private Sub mnuQuit_Click()
    Unload Me
    
End Sub

Private Sub mnuSample_Click()
    Dim sText       As String
    Dim sRowid      As String
    Dim NodeX       As Node
    
    
        
    DoEvents
    Call cmdClear_Click
    panelItem.Visible = False
    panelMicro.Visible = True
    panelMicro.Left = 45
    panelMicro.Top = 180
    panelMicro.ZOrder 0
    
    
    DoEvents
    GoSub TreeView_Select
    
    Exit Sub
    
    
TreeView_Select:
    tvSample.Nodes.Clear
    Set NodeX = tvSample.Nodes.Add(, , "A0", "미생물검체 분류")

    
    StrSql = ""
    StrSql = StrSql & " SELECT Class2, Max(RowID) RWID"
    StrSql = StrSql & " FROM   TWEXAM_Sample"
    StrSql = StrSql & " WHERE  CLass1 = 'm'"
    StrSql = StrSql & " GROUP  BY Class2"
    
    If False = adoSetOpen(StrSql, adoSet) Then Exit Sub
    
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
    
    StrSql = ""
    StrSql = StrSql & " SELECT Code, class1, Codenm, anatomy, RowID"
    StrSql = StrSql & " FROM   TWEXAM_Sample"
    StrSql = StrSql & " WHERE  Class2 = '" & tvSample.Nodes("A1" & sRowid).Text & "'"
    StrSql = StrSql & " AND    CLass1 = 'm'"
    StrSql = StrSql & " ORDER BY Code"
    
    If False = adoSetOpen(StrSql, adoSubCode1) Then Return
    
    Do Until adoSubCode1.EOF
        sSubText1 = "B2" & adoSubCode1.Fields("RowID")
        
        sSubText2 = Trim(adoSubCode1.Fields("Code").Value & ".") & _
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

Private Sub ssRet_Click(ByVal Col As Long, ByVal Row As Long)
    
    If Col = 8 Then
        If Row > 0 Then
            GoSub Display_Default_Set
        End If
    End If
    Exit Sub
    
Display_Default_Set:
    ssRet.Row = Row
    ssRet.Col = 8
    
    If ssRet.Text = "Y" Then
        ssRet.Text = ""
    Else
        ssRet.Text = "Y"
    End If
    
    Return
    
End Sub

Private Sub tvSample_DblClick()
    Dim sCode       As String * 8
    
    
    If Left(tvSample.SelectedItem.Key, 2) = "B2" Then
        sCode = Left(tvSample.SelectedItem.Text, 8)
    Else
        Exit Sub
    End If
    txtItemCode.Text = sCode
    txtItemName.Text = Mid(Trim(tvSample.SelectedItem.Text), 10, Len(Trim(tvSample.SelectedItem.Text)) - 9)
    
    StrSql = ""
    StrSql = StrSql & " SELECT a.*, a.RowID"
    StrSql = StrSql & " FROM   TWEXAM_Ret a"
    StrSql = StrSql & " WHERE  a.Slipno = '42'"
    StrSql = StrSql & " AND    a.ItemCd = '" & sCode & "'"
    StrSql = StrSql & " ORDER  BY Seqno"
    
    ssRet.MaxRows = 0
    If False = adoSetOpen(StrSql, adoSet) Then Exit Sub
    ssRet.MaxRows = adoSet.RecordCount
    
    Do Until adoSet.EOF
        ssRet.Row = ssRet.DataRowCnt + 1
        ssRet.Col = 1: ssRet.Text = adoSet.Fields("RowID").Value & ""
        ssRet.Col = 2: ssRet.Value = False
        If adoSet.Fields("RetGB").Value & "" = "M" Then
            ssRet.Col = 3: ssRet.TypeComboBoxCurSel = 1
        Else
            ssRet.Col = 3: ssRet.TypeComboBoxCurSel = 0
        End If
        ssRet.Col = 4: ssRet.Text = adoSet.Fields("Slipno").Value & ""
        ssRet.Col = 5: ssRet.Text = adoSet.Fields("itemCD").Value & ""
        ssRet.Col = 6: ssRet.Text = adoSet.Fields("Ret").Value & ""
        ssRet.Col = 7: ssRet.Text = adoSet.Fields("Seqno").Value & ""
        ssRet.Col = 8: ssRet.Text = adoSet.Fields("Normal").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    
End Sub

Private Sub txtSlipno_Change()

End Sub
