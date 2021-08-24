VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Begin VB.Form frmNormal 
   Caption         =   "정상수치 관리"
   ClientHeight    =   6495
   ClientLeft      =   585
   ClientTop       =   1920
   ClientWidth     =   11085
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
   MDIChild        =   -1  'True
   ScaleHeight     =   6495
   ScaleWidth      =   11085
   Begin Threed.SSPanel SSPanel1 
      Height          =   975
      Left            =   180
      TabIndex        =   2
      Top             =   240
      Width           =   7995
      _Version        =   65536
      _ExtentX        =   14102
      _ExtentY        =   1720
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
         Height          =   855
         Left            =   6900
         TabIndex        =   12
         Top             =   60
         Width           =   795
         _Version        =   65536
         _ExtentX        =   1402
         _ExtentY        =   1508
         _StockProps     =   78
         Caption         =   "화면정리"
         BevelWidth      =   1
         Outline         =   0   'False
         Picture         =   "frmNormal.frx":0000
      End
      Begin Threed.SSCommand cmdInsert 
         Height          =   855
         Left            =   6060
         TabIndex        =   11
         Top             =   60
         Width           =   795
         _Version        =   65536
         _ExtentX        =   1402
         _ExtentY        =   1508
         _StockProps     =   78
         Caption         =   "입력확인"
         BevelWidth      =   1
         Outline         =   0   'False
         Picture         =   "frmNormal.frx":1792
      End
      Begin Threed.SSCommand cmdQry 
         Height          =   855
         Left            =   4860
         TabIndex        =   1
         Top             =   60
         Width           =   795
         _Version        =   65536
         _ExtentX        =   1402
         _ExtentY        =   1508
         _StockProps     =   78
         Caption         =   "조회확인"
         BevelWidth      =   1
         Outline         =   0   'False
         Picture         =   "frmNormal.frx":2F54
      End
      Begin Threed.SSCommand cmdHelpItem 
         Height          =   315
         Left            =   2220
         TabIndex        =   6
         Top             =   120
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   556
         _StockProps     =   78
         Caption         =   "&H"
      End
      Begin VB.TextBox txtItemName 
         Height          =   315
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   420
         Width           =   3615
      End
      Begin VB.TextBox txtItemCode 
         Height          =   315
         Left            =   1080
         MaxLength       =   8
         TabIndex        =   0
         Top             =   120
         Width           =   1155
      End
      Begin VB.Label Label1 
         Caption         =   "ITEMCODE:"
         Height          =   195
         Left            =   180
         TabIndex        =   4
         Top             =   180
         Width           =   855
      End
   End
   Begin FPSpreadADO.fpSpread ssRefData 
      Height          =   5055
      Left            =   180
      TabIndex        =   5
      Top             =   1320
      Width           =   7995
      _Version        =   196608
      _ExtentX        =   14102
      _ExtentY        =   8916
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
      MaxCols         =   11
      ScrollBars      =   2
      SpreadDesigner  =   "frmNormal.frx":382E
      Appearance      =   1
   End
   Begin Threed.SSPanel panelItemList 
      Height          =   6135
      Left            =   5220
      TabIndex        =   7
      Top             =   240
      Visible         =   0   'False
      Width           =   5655
      _Version        =   65536
      _ExtentX        =   9975
      _ExtentY        =   10821
      _StockProps     =   15
      BackColor       =   8421376
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   2
      FloodColor      =   0
      Begin VB.ComboBox cmbSlip 
         Height          =   300
         Left            =   240
         Style           =   2  '드롭다운 목록
         TabIndex        =   9
         Top             =   240
         Width           =   3555
      End
      Begin Threed.SSCommand cmdHide 
         Height          =   315
         Left            =   3900
         TabIndex        =   10
         Top             =   240
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   556
         _StockProps     =   78
         Caption         =   "Quit"
      End
      Begin FPSpreadADO.fpSpread ssItem 
         Height          =   5475
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   4995
         _Version        =   196608
         _ExtentX        =   8811
         _ExtentY        =   9657
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
         GridShowHoriz   =   0   'False
         MaxCols         =   2
         ScrollBars      =   2
         SpreadDesigner  =   "frmNormal.frx":7715
         Appearance      =   1
      End
   End
   Begin VB.Menu mnuQuit 
      Caption         =   "Quit"
   End
End
Attribute VB_Name = "frmNormal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbSlip_Click()
    Dim sCodeky     As String
    
    If cmbSLip.ListIndex = -1 Then Exit Sub
    
    sCodeky = Left(cmbSLip.List(cmbSLip.ListIndex), 2)
    
    strSql = ""
    strSql = strSql & " SELECT Codeky, Itemnm"
    strSql = strSql & " FROM   TWEXAM_ITEMML"
    strSql = strSql & " WHERE  SubStr(Codeky, 1,2) = '" & sCodeky & "'"
    strSql = strSql & " ORDER  BY Codeky"
    
    ssItem.MaxRows = 0
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    ssItem.MaxRows = adoSet.RecordCount
    
    Do Until adoSet.EOF
        ssItem.Row = ssItem.DataRowCnt + 1
        ssItem.Col = 1: ssItem.Text = Trim(adoSet.Fields("Codeky").Value & "")
        ssItem.Col = 2: ssItem.Text = Trim(adoSet.Fields("Itemnm").Value & "")
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    
    
End Sub

Private Sub cmdClear_Click()
    
    txtItemCode.Text = ""
    txtItemName.Text = ""
    Call Spread_Set_Clear(ssRefData)
    txtItemCode.SetFocus
    
End Sub


Private Sub cmdHelpItem_Click()
    
    panelItemList.Top = 200
    panelItemList.Left = 5000
    panelItemList.Visible = True
    panelItemList.ZOrder 0
    cmbSLip.SetFocus
    
End Sub

Private Sub cmdHide_Click()
    panelItemList.Visible = False
    
End Sub

Private Sub cmdInsert_Click()
    Dim sAppdate        As String
    Dim sAppGubun       As String
    Dim sItemCD         As String
    Dim nAgeMin         As Integer
    Dim nAgeMax         As Integer
    Dim sMmin           As String
    Dim sMmax           As String
    Dim sFmin           As String
    Dim sFmax           As String
    Dim sRowid          As String
    Dim bDelCheck       As Boolean
    
    For I = 1 To ssRefData.DataRowCnt
        ssRefData.Row = I
        GoSub SpreadData_Move
        If bDelCheck = True Then              'Col =1: Del Check
            If Trim(sRowid) <> "" Then        'Col =2: RowID Check
                GoSub RefData_Delete
            End If
        Else
            If Trim(sRowid) = "" Then
                GoSub RefData_Insert
            Else
                GoSub RefData_Update
            End If
        End If
    Next
    Call cmdQry_Click
    Exit Sub
    
    
RefData_Delete:
    strSql = ""
    strSql = strSql & " DELETE"
    strSql = strSql & " FROM   TWEXAM_RefData"
    strSql = strSql & " WHERE  ROWID  =  '" & sRowid & "'"
    adoConnect.BeginTrans
    If adoExec(strSql) Then
        adoConnect.CommitTrans
    Else
        adoConnect.RollbackTrans
    End If
    
    Return
    

SpreadData_Move:
    ssRefData.Row = I
    ssRefData.Col = 1
    If ssRefData.Value = True Then
        bDelCheck = True
    Else
        bDelCheck = False
    End If
    ssRefData.Col = 2:  sRowid = Trim(ssRefData.Text)
    ssRefData.Col = 3:  sAppdate = ssRefData.Text
    ssRefData.Col = 4:  sItemCD = Trim(ssRefData.Text)
    ssRefData.Col = 5:  sAppGubun = Trim(ssRefData.Text)
    ssRefData.Col = 6:  nAgeMin = Val(ssRefData.Text)
    ssRefData.Col = 7:  nAgeMax = Val(ssRefData.Text)
    ssRefData.Col = 8:  sMmin = ssRefData.Text
    ssRefData.Col = 9:  sMmax = ssRefData.Text
    ssRefData.Col = 10: sFmin = ssRefData.Text
    ssRefData.Col = 11: sFmax = ssRefData.Text
    Return
    
'/___________________________________________________________________
RefData_Insert:
    strSql = ""
    strSql = strSql & " INSERT INTO TWEXAM_RefData"
    strSql = strSql & "        (iTemCode, appDate, appGubun, ageMin, ageMax, "
    strSql = strSql & "         M_min,    M_max,   F_min,    F_max)"
    strSql = strSql & " VALUES ('" & txtItemCode.Text & "',"
    strSql = strSql & "              TO_DATE('" & sAppdate & "','YYYY-MM-DD'),"
    strSql = strSql & "         '" & sAppGubun & "',"
    strSql = strSql & "          " & nAgeMin & ","
    strSql = strSql & "          " & nAgeMax & ","
    strSql = strSql & "         '" & sMmin & "',"
    strSql = strSql & "         '" & sMmax & "',"
    strSql = strSql & "         '" & sFmin & "',"
    strSql = strSql & "         '" & sFmax & "')"
    
    adoConnect.BeginTrans
    If adoExec(strSql) Then
        adoConnect.CommitTrans
    Else
        adoConnect.RollbackTrans
    End If
    Return
    
RefData_Update:
    strSql = ""
    strSql = strSql & " UPDATE TWEXAM_RefData"
    strSql = strSql & " SET    appDate  = TO_DATE('" & sAppdate & "','yyyy-MM-dd'),"
    strSql = strSql & "        appGubun = '" & sAppGubun & "',"
    strSql = strSql & "        ageMin   =  " & nAgeMin & ","
    strSql = strSql & "        ageMax   =  " & nAgeMax & ","
    strSql = strSql & "        M_min    = '" & sMmin & "',"
    strSql = strSql & "        M_max    = '" & sMmax & "',"
    strSql = strSql & "        F_min    = '" & sFmin & "',"
    strSql = strSql & "        F_max    = '" & sFmax & "'"
    strSql = strSql & " WHERE  RowID    = '" & sRowid & "'"
    adoConnect.BeginTrans
    If adoExec(strSql) Then
        adoConnect.CommitTrans
    Else
        adoConnect.RollbackTrans
    End If
    
    Return
    
    
End Sub

Private Sub cmdQry_Click()
    
    strSql = " SELECT ItemNm FROM  TWEXAM_iTemML WHERE Codeky = '" & txtItemCode.Text & "'"
    If False = adoSetOpen(strSql, adoSet) Then
        MsgBox "해당 코드가 없습니다!"
        txtItemCode.Text = ""
        txtItemName.Text = ""
        Exit Sub
    End If
    txtItemName.Text = adoSet.Fields("iTemNM").Value & ""
    Call adoSetClose(adoSet)
    
    strSql = ""
    strSql = strSql & " SELECT a.*, a.RowID, "
    strSql = strSql & "        TO_CHAR(a.appDate,'YYYY-MM-DD') appDate"
    strSql = strSql & " FROM   TWEXAM_RefData a"
    strSql = strSql & " WHERE  a.iTemCode  =  '" & Me.txtItemCode.Text & "'"
    strSql = strSql & " ORDER  BY a.appDate DESC, a.AppGubun ASC, a.AgeMin ASC"
    
       
    Call Spread_Set_Clear(ssRefData)
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    
    
    Do Until adoSet.EOF
        ssRefData.Row = ssRefData.DataRowCnt + 1
        ssRefData.Col = 2:  ssRefData.Text = adoSet.Fields("RowID").Value & ""
        ssRefData.Col = 3:  ssRefData.Text = adoSet.Fields("appDate").Value & ""
        ssRefData.Col = 4:  ssRefData.Text = adoSet.Fields("ItemCode").Value & ""
        ssRefData.Col = 5:  ssRefData.Text = adoSet.Fields("appGubun").Value & ""
        ssRefData.Col = 6:  ssRefData.Text = adoSet.Fields("ageMin").Value & ""
        ssRefData.Col = 7:  ssRefData.Text = adoSet.Fields("ageMax").Value & ""
        ssRefData.Col = 8:  ssRefData.Text = Trim(adoSet.Fields("M_min").Value & "")
        ssRefData.Col = 9:  ssRefData.Text = Trim(adoSet.Fields("M_max").Value & "")
        ssRefData.Col = 10: ssRefData.Text = Trim(adoSet.Fields("F_min").Value & "")
        ssRefData.Col = 11: ssRefData.Text = Trim(adoSet.Fields("F_max").Value & "")
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)

End Sub

Private Sub Form_Load()
        
    GoSub Form_Size_Check
    GoSub Spread_Size_Set
    GoSub Get_SlipData
    If Trim(gSText) <> "" Then
        txtItemCode.Text = gSText
        Call cmdQry_Click
    End If
    
    Exit Sub
    
    
Form_Size_Check:
    Me.Top = 1
    Me.Left = 1
    Me.Height = 7185
    Me.Width = 11200
    Return
    
Spread_Size_Set:
    ssItem.RowHeight(-1) = 11
    ssRefData.RowHeight(-1) = 12
    Return

Get_SlipData:
    strSql = ""
    strSql = strSql & " SELECT *"
    strSql = strSql & " FROM   TWEXAM_SPECODE"
    strSql = strSql & " WHERE  Codegu = '12'"
    strSql = strSql & " ORDER  By Codeky"
    
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    
    Do Until adoSet.EOF
        cmbSLip.AddItem Trim(adoSet.Fields("Codeky").Value & "") & ". " & _
                        Trim(adoSet.Fields("Codenm").Value & "")
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    Return

End Sub


Private Sub mnuQuit_Click()
    Unload Me
    
End Sub


Private Sub ssItem_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim sCodeky     As String
    Dim sCodenm     As String
    
    ssItem.Row = Row
    ssItem.Col = 1: txtItemCode.Text = ssItem.Text
                    sCodeky = ssItem.Text
    ssItem.Col = 2: txtItemName.Text = ssItem.Text
                    sCodenm = ssItem.Text
                    
                    
    GoSub Get_RefData
    DoEvents
    panelItemList.Visible = False
    mdiMain.stbMain.Panels(1).Text = ""
    Exit Sub
    
Get_RefData:
    strSql = ""
    strSql = strSql & " SELECT a.*, a.RowID, "
    strSql = strSql & "        TO_CHAR(a.appDate,'YYYY-MM-DD') appDate"
    strSql = strSql & " FROM   TWEXAM_RefData a"
    strSql = strSql & " WHERE  a.iTemCode  =  '" & sCodeky & "'"
    strSql = strSql & " ORDER  BY a.appDate DESC "
    
       
    Call Spread_Set_Clear(ssRefData)
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    
    Do Until adoSet.EOF
        ssRefData.Row = ssRefData.DataRowCnt + 1
        ssRefData.Col = 2:  ssRefData.Text = adoSet.Fields("RowID").Value & ""
        ssRefData.Col = 3:  ssRefData.Text = adoSet.Fields("appDate").Value & ""
        ssRefData.Col = 4:  ssRefData.Text = adoSet.Fields("ItemCode").Value & ""
        ssRefData.Col = 5:  ssRefData.Text = adoSet.Fields("appGubun").Value & ""
        ssRefData.Col = 6:  ssRefData.Text = adoSet.Fields("ageMin").Value & ""
        ssRefData.Col = 7:  ssRefData.Text = adoSet.Fields("ageMax").Value & ""
        ssRefData.Col = 8:  ssRefData.Text = Trim(adoSet.Fields("M_min").Value & "")
        ssRefData.Col = 9:  ssRefData.Text = Trim(adoSet.Fields("M_max").Value & "")
        ssRefData.Col = 10: ssRefData.Text = Trim(adoSet.Fields("F_min").Value & "")
        ssRefData.Col = 11: ssRefData.Text = Trim(adoSet.Fields("F_max").Value & "")
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    Return
    
    
    
End Sub

Private Sub ssRefData_DblClick(ByVal Col As Long, ByVal Row As Long)
    
    If Row = 0 Then
        If Col = 0 Then Exit Sub      'Row Header
        If Col = 11 Then Exit Sub     'Button Cell
    
        ssRefData.SortBy = SS_SORT_BY_ROW
        ssRefData.SortKey(1) = Col
        If ssRefData.SortKeyOrder(1) = SortKeyOrderDescending Then
            ssRefData.SortKeyOrder(1) = SS_SORT_ORDER_ASCENDING
        Else
            ssRefData.SortKeyOrder(1) = SortKeyOrderDescending
        End If
        ssRefData.Col = 1: ssRefData.Col2 = ssRefData.DataColCnt
        ssRefData.Row = 1: ssRefData.Row2 = ssRefData.DataRowCnt
        ssRefData.Action = SS_ACTION_SORT
    End If
    
End Sub

Private Sub ssRefData_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    
    If Col = 3 Then
        GoSub Ref_Date_Check
    End If
    Exit Sub
    
Ref_Date_Check:
    ssRefData.Row = Row
    ssRefData.Col = 3
    If False = IsDate(ssRefData.Text) Then
        ssRefData.Text = Dual_Date_Get("yyyy-MM-dd")
    End If
    Return
        
    
End Sub

Private Sub txtItemCode_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
        
End Sub
