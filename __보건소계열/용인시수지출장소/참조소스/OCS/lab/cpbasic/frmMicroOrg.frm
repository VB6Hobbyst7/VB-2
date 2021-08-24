VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Begin VB.Form frmMicroOrg 
   Caption         =   "세균코드 관리"
   ClientHeight    =   7410
   ClientLeft      =   525
   ClientTop       =   930
   ClientWidth     =   10875
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7410
   ScaleWidth      =   10875
   Begin Threed.SSPanel SSPanel1 
      Height          =   1155
      Left            =   180
      TabIndex        =   2
      Top             =   120
      Width           =   9915
      _Version        =   65536
      _ExtentX        =   17489
      _ExtentY        =   2037
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
      BorderWidth     =   0
      BevelInner      =   1
      Begin VB.ComboBox cmbGrp 
         Height          =   300
         Left            =   5220
         TabIndex        =   14
         Top             =   600
         Width           =   855
      End
      Begin VB.ComboBox cmbStatus 
         Appearance      =   0  '평면
         Height          =   300
         Left            =   7200
         TabIndex        =   5
         Top             =   615
         Width           =   795
      End
      Begin VB.ComboBox cmbGram 
         Appearance      =   0  '평면
         Height          =   300
         Left            =   3240
         TabIndex        =   4
         Top             =   615
         Width           =   675
      End
      Begin VB.TextBox txtOrgNm 
         Height          =   300
         Left            =   3240
         TabIndex        =   3
         Top             =   240
         Width           =   4755
      End
      Begin VB.TextBox txtOrgCode 
         Height          =   300
         Left            =   900
         TabIndex        =   0
         Top             =   240
         Width           =   1155
      End
      Begin Threed.SSCommand cmdInsert 
         Height          =   735
         Left            =   8100
         TabIndex        =   6
         Top             =   240
         Width           =   795
         _Version        =   65536
         _ExtentX        =   1402
         _ExtentY        =   1296
         _StockProps     =   78
         Caption         =   "입력확인"
         BevelWidth      =   1
         Outline         =   0   'False
         Picture         =   "frmMicroOrg.frx":0000
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   735
         Left            =   9060
         TabIndex        =   7
         Top             =   240
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1296
         _StockProps     =   78
         Caption         =   "화면정리"
         BevelWidth      =   1
         Outline         =   0   'False
         Picture         =   "frmMicroOrg.frx":17C2
      End
      Begin VB.Line Line1 
         X1              =   2100
         X2              =   2100
         Y1              =   240
         Y2              =   960
      End
      Begin VB.Label Label6 
         Appearance      =   0  '평면
         BackColor       =   &H80000004&
         Caption         =   "Status"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6600
         TabIndex        =   12
         Top             =   660
         Width           =   555
      End
      Begin VB.Label Label5 
         Appearance      =   0  '평면
         BackColor       =   &H80000004&
         Caption         =   "약제그룹코드"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4020
         TabIndex        =   11
         Top             =   675
         Width           =   1155
      End
      Begin VB.Label Label4 
         Appearance      =   0  '평면
         BackColor       =   &H80000004&
         Caption         =   "GramData"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2280
         TabIndex        =   10
         Top             =   645
         Width           =   855
      End
      Begin VB.Label Label3 
         Appearance      =   0  '평면
         BackColor       =   &H80000004&
         Caption         =   "세균명"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2280
         TabIndex        =   9
         Top             =   270
         Width           =   675
      End
      Begin VB.Label Label2 
         Appearance      =   0  '평면
         BackColor       =   &H80000004&
         Caption         =   "세균코드"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   300
         Width           =   735
      End
   End
   Begin FPSpreadADO.fpSpread ssOrgList 
      Height          =   5535
      Left            =   180
      TabIndex        =   1
      Top             =   1680
      Width           =   9915
      _Version        =   196608
      _ExtentX        =   17489
      _ExtentY        =   9763
      _StockProps     =   64
      BackColorStyle  =   1
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   5
      ScrollBarExtMode=   -1  'True
      ScrollBars      =   2
      SpreadDesigner  =   "frmMicroOrg.frx":2F54
      UserResize      =   0
      VisibleCols     =   5
      VisibleRows     =   500
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Caption         =   "▼→ RowHeader 를 Double Click 하면 해당 세균코드를 삭제할 수 있습니다!."
      Height          =   195
      Left            =   360
      TabIndex        =   13
      Top             =   1380
      Width           =   6375
   End
   Begin VB.Menu mnuQuit 
      Caption         =   "Quit"
   End
   Begin VB.Menu mnuQryMain 
      Caption         =   "세균조회"
   End
End
Attribute VB_Name = "frmMicroOrg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Sub cmdClear_Click()
    
    txtOrgCode.Text = ""
    txtOrgNm.Text = ""
    cmbGrp.ListIndex = -1
    cmbGram.ListIndex = -1
    cmbStatus.ListIndex = -1
    
    If vbNo = MsgBox("아래의 Spread 화면까지 정리를 하시겠습니까?", vbYesNo + vbQuestion, "화면정리 요청Box") Then
        Exit Sub
    End If
    
    ssOrgList.Row = 1
    ssOrgList.Row2 = ssOrgList.DataRowCnt
    ssOrgList.Col = 1
    ssOrgList.Col2 = ssOrgList.DataColCnt
    ssOrgList.BlockMode = True
    ssOrgList.Action = ActionClear
    ssOrgList.BlockMode = False
    
    
End Sub


Private Sub cmdInsert_Click()
    If Me.txtOrgCode.Text = "" Then Exit Sub
    
    strSql = ""
    strSql = strSql & " SELECT *"
    strSql = strSql & " FROM   TWEXAM_ORGLIST"
    strSql = strSql & " WHERE  ORG_CODE = '" & Trim(txtOrgCode.Text) & "'"
    
    If False = adoSetOpen(strSql, adoSet) Then
        GoSub OrgList_Insert
    Else
        Call adoSetClose(adoSet)
        GoSub OrgList_Update
    End If
    
    Exit Sub
    

OrgList_Insert:
    strSql = ""
    strSql = strSql & " INSERT INTO TWEXAM_ORGLIST"
    strSql = strSql & "       (      Org_Code, Org_Name, Org_Gram, Org_AntiGr, Org_Status)"
    strSql = strSql & " VALUES( '" & txtOrgCode.Text & "',"
    strSql = strSql & "         '" & Quot_Conv(Trim(txtOrgNm.Text)) & "',"
    strSql = strSql & "         '" & Trim(cmbGram.Text) & "',"
    strSql = strSql & "         '" & Trim(cmbGrp.Text) & "',"
    strSql = strSql & "         '" & Trim(cmbStatus.Text) & "')"
    Call adoExec(strSql)
    GoSub frmOrgList_Clear_Reset
    Return

OrgList_Update:
    strSql = ""
    strSql = strSql & " UPDATE TWEXAM_ORGLIST"
    strSql = strSql & " SET    Org_Name   =  '" & Quot_Conv(Trim(txtOrgNm.Text)) & "',"
    strSql = strSql & "        Org_Gram   =  '" & Trim(cmbGram.Text) & "',"
    strSql = strSql & "        Org_AntiGr =  '" & Trim(cmbGrp.Text) & "',"
    strSql = strSql & "        Org_Status =  '" & Trim(cmbStatus.Text) & "'"
    strSql = strSql & " WHERE  Org_Code   =  '" & Trim(txtOrgCode.Text) & "'"
    Call adoExec(strSql)
    GoSub Spread_Refrash
    GoSub frmOrgList_Clear_Reset
    Return
    
    
frmOrgList_Clear_Reset:
    txtOrgCode.Text = ""
    txtOrgNm.Text = ""
    cmbGram.ListIndex = -1
    cmbGrp.ListIndex = -1
    cmbStatus.ListIndex = -1
    
    Return
    
Spread_Refrash:
    For i = 1 To frmMicroOrg.ssOrgList.DataRowCnt
        frmMicroOrg.ssOrgList.Row = i
        frmMicroOrg.ssOrgList.Col = 1
        If Trim(frmMicroOrg.ssOrgList.Text) = Trim(Me.txtOrgCode.Text) Then
            frmMicroOrg.ssOrgList.Col = 4
            frmMicroOrg.ssOrgList.TypeButtonText = Trim(Me.cmbGrp.Text)
            Exit For
        End If
    Next

    Return
    
End Sub

Private Sub Form_Load()
    
    Me.Left = 1
    Me.Top = 1
    Me.Height = 7900
    Me.Width = 11000
    
    GoSub SET_Gram_Data
    GoSub SET_Status_Data
    GoSub SET_Group_Data
    
    Exit Sub
    
        
'/----------------------------------

SET_Gram_Data:
    cmbGram.Clear
    
    cmbGram.AddItem "-"
    cmbGram.AddItem "+"
    cmbGram.AddItem "a"
    cmbGram.AddItem "b"
    cmbGram.AddItem "f"
    cmbGram.AddItem "m"
    cmbGram.AddItem "o"
    cmbGram.AddItem "w"
    
    Return
    
    
SET_Status_Data:
    cmbStatus.Clear
    cmbStatus.AddItem "C"
    cmbStatus.AddItem "O"
    Return
    
SET_Group_Data:
    strSql = ""
    strSql = strSql & " SELECT Org_AntiGr"
    strSql = strSql & " FROM   TWEXAM_ORGLIST"
    strSql = strSql & " GROUP  BY Org_AntiGr"
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    Do Until adoSet.EOF
        cmbGrp.AddItem Trim(adoSet.Fields("Org_AntiGr").Value & "")
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    Return
    
    
End Sub

Private Sub mnuDelete_Click()
    
    Call ssOrgList_DblClick(1, ssOrgList.ActiveRow)
    
End Sub

Private Sub mnuGroup_Click()
    
    hWndReturn = vbNull
    frmMicroGroup.Show vbModal
    
End Sub

Private Sub mnuQryMain_Click()
    frmMicroQryOrg.Show vbModal
    
End Sub

Private Sub mnuQuit_Click()
    Unload Me
    
End Sub



Private Sub ssOrgList_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    
    If Row > 0 And Col = 4 Then
        DoEvents
        ssOrgList.Row = Row
        ssOrgList.Col = 4: gSText = ssOrgList.TypeButtonText
        gSOrgCall = "SPREAD"
        frmMicroGroup.Show vbModal
    End If
    
End Sub

Private Sub ssOrgList_DblClick(ByVal Col As Long, ByVal Row As Long)
    
    If Row > 0 Then
        If Col = 0 Then
            GoSub OrgList_Delete
        Else
            ssOrgList.Row = Row
            ssOrgList.Col = 1: txtOrgCode.Text = ssOrgList.Text
            ssOrgList.Col = 2: txtOrgNm.Text = ssOrgList.Text
            ssOrgList.Col = 3: cmbGram.Tag = ssOrgList.Text
                               GoSub SET_GramListIndex
            ssOrgList.Col = 4: Call SetComboBox(cmbGrp, Trim(ssOrgList.Text))
            ssOrgList.Col = 5: Call SetComboBox(cmbStatus, Trim(ssOrgList.Text))
            
            'ssOrgList.Col = 5: cmbStatus.Tag = ssOrgList.Text
            '                   GoSub SET_StatusIndex
        End If
        Exit Sub
    End If
        
    Exit Sub
    
    
OrgList_Delete:
    Dim cCode       As String
    Dim cName       As String
    
    ssOrgList.Row = Row
    ssOrgList.Col = 1: cCode = Trim(ssOrgList.Text)
    ssOrgList.Col = 2: cName = Trim(ssOrgList.Text)
    
    sMsg = cName & vbCrLf & " 의 Data 를 삭제하시겠습니까?"
    
    If vbNo = MsgBox(sMsg, vbYesNo + vbQuestion, "삭제 확인 Box") Then
        Exit Sub
    End If
    
    strSql = ""
    strSql = strSql & " DELETE "
    strSql = strSql & " FROM  TWEXAM_ORGLIST"
'y  strSql = strSql & " WHERE Codeky = '" & cCode & "'"
    strSql = strSql & " WHERE ORG_code = '" & cCode & "'"
    Call adoExec(strSql)
    Return
    
    
SET_GramListIndex:
    For i = 0 To cmbGram.ListCount - 1
        If Trim(cmbGram.List(i)) = Trim(cmbGram.Tag) Then
            cmbGram.ListIndex = i
            Exit For
        End If
    Next
    Return

SET_StatusIndex:
    For i = 0 To cmbStatus.ListCount - 1
        If Trim(cmbStatus.List(i)) = Trim(cmbStatus.Tag) Then
            cmbStatus.ListIndex = i
            Exit For
        End If
    Next
    Return
    
End Sub

Private Sub txtOrgCode_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        If txtOrgCode.Text <> "" Then
            GoSub Org_List_Select
        End If
    End If
    Exit Sub
    
    
Org_List_Select:
    strSql = ""
    strSql = strSql & " SELECT *"
    strSql = strSql & " FROM   TWEXAM_ORGLIST"
    strSql = strSql & " WHERE  ORG_CODE = '" & Trim(txtOrgCode.Text) & "'"
    If False = adoSetOpen(strSql, adoSet) Then
        txtOrgNm.Text = ""
        cmbGram.ListIndex = -1
        cmbGrp.ListIndex = -1
        cmbStatus.ListIndex = -1
        Return
    Else
        txtOrgNm.Text = Trim(adoSet.Fields("Org_name").Value & "")
        cmbGram.ListIndex = SetComboBox(cmbGram, Trim(adoSet.Fields("Org_Gram").Value & ""))
        Call SetComboBox(cmbGrp, Trim(adoSet.Fields("Org_AntiGr").Value & ""))
        cmbStatus.ListIndex = SetComboBox(cmbGram, Trim(adoSet.Fields("Org_Status").Value & ""))
    End If
    Return
    
    
End Sub
