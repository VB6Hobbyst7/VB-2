VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{48932A52-981F-101B-A7FB-4A79242FD97B}#2.0#0"; "TAB32X20.OCX"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Begin VB.Form frmQrySuga 
   Caption         =   "수가코드 조회"
   ClientHeight    =   7875
   ClientLeft      =   2925
   ClientTop       =   2055
   ClientWidth     =   7065
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
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   7065
   Begin VB.Frame Frame1 
      Caption         =   "한/영선택"
      Height          =   780
      Left            =   5085
      TabIndex        =   11
      Top             =   45
      Width           =   1770
      Begin VB.OptionButton Option1 
         Caption         =   "영문명"
         Height          =   195
         Left            =   225
         TabIndex        =   13
         Top             =   270
         Value           =   -1  'True
         Width           =   1155
      End
      Begin VB.OptionButton Option2 
         Caption         =   "한글명"
         Height          =   195
         Left            =   225
         TabIndex        =   12
         Top             =   525
         Width           =   1155
      End
   End
   Begin Threed.SSPanel panelBun 
      Height          =   6090
      Left            =   1260
      TabIndex        =   1
      Top             =   1350
      Visible         =   0   'False
      Width           =   4875
      _Version        =   65536
      _ExtentX        =   8599
      _ExtentY        =   10742
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
      BevelOuter      =   1
      BevelInner      =   2
      Begin VB.TextBox txtBname 
         BackColor       =   &H00C0E0FF&
         Height          =   315
         Left            =   255
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   555
         Width           =   2895
      End
      Begin VB.TextBox txtCgb 
         BackColor       =   &H00C0E0FF&
         Height          =   315
         Left            =   255
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   240
         Width           =   2055
      End
      Begin Threed.SSCommand cmdQryOk 
         Height          =   735
         Left            =   3660
         TabIndex        =   3
         Top             =   180
         Width           =   915
         _Version        =   65536
         _ExtentX        =   1614
         _ExtentY        =   1296
         _StockProps     =   78
         Caption         =   "조회확인"
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Outline         =   0   'False
      End
      Begin MSComctlLib.TreeView tvBun 
         Height          =   4920
         Left            =   240
         TabIndex        =   2
         Top             =   945
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   8678
         _Version        =   393217
         LineStyle       =   1
         Style           =   7
         BorderStyle     =   1
         Appearance      =   1
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   705
      Left            =   90
      TabIndex        =   6
      Top             =   135
      Width           =   4875
      _Version        =   65536
      _ExtentX        =   8599
      _ExtentY        =   1244
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
      Begin TabproLib.vaTabPro vaTabPro1 
         Height          =   600
         Left            =   135
         OleObjectBlob   =   "frmQrySuga.frx":0000
         TabIndex        =   7
         Top             =   45
         Width           =   3345
         Begin VB.TextBox txtSucode 
            Appearance      =   0  '평면
            Height          =   285
            Left            =   1455
            TabIndex        =   9
            Top             =   3525
            Width           =   1680
         End
         Begin VB.TextBox txtQrySuname 
            Appearance      =   0  '평면
            Enabled         =   0   'False
            Height          =   270
            Left            =   -18074
            TabIndex        =   8
            Top             =   -15539
            Width           =   1905
         End
      End
      Begin Threed.SSCommand cmdQry 
         Height          =   585
         Left            =   3690
         TabIndex        =   10
         Top             =   45
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
         _ExtentY        =   1032
         _StockProps     =   78
         Caption         =   "조회확인"
         BevelWidth      =   1
         Outline         =   0   'False
      End
   End
   Begin FPSpreadADO.fpSpread ssSuga 
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   900
      Width           =   6795
      _Version        =   196608
      _ExtentX        =   11986
      _ExtentY        =   11880
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
      MaxCols         =   4
      ScrollBarExtMode=   -1  'True
      ScrollBars      =   2
      SpreadDesigner  =   "frmQrySuga.frx":0277
      UserResize      =   0
      VisibleCols     =   3
      VisibleRows     =   500
      Appearance      =   1
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
   Begin VB.Menu mnuBun 
      Caption         =   "분류별조회"
   End
End
Attribute VB_Name = "frmQrySuga"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdQry_Click()
    Dim sSucode      As String
    Dim sSugba       As String * 2
    
    Select Case vaTabPro1.ActiveTab
        Case 0:
                If Trim(txtSucode.Text) = "" Then
                    txtSucode.Text = UCase(txtSucode.Text)
                    Exit Sub: End If
        Case 1: If Trim(txtQrySuname.Text) = "" Then Exit Sub
    End Select
    
    strSql = ""
    strSql = strSql & " SELECT a.Sucode, b.Sunamee, b.Sunamek, b.Bcode, a.SugbA"
    strSql = strSql & " FROM   TWBAS_TSUGA a,"
    strSql = strSql & "        TWBAS_NSUGA b "
    strSql = strSql & " WHERE  a.Sunext = b.Sunext(+)"
    
    Select Case vaTabPro1.ActiveTab
        Case 0
            strSql = strSql & " AND    a.Sucode    Like '" & txtSucode.Text & "%'"
        Case 1
            If Option1.Value = True Then
                strSql = strSql & " AND   (upper(b.sunamek)   Like '" & UCase(txtQrySuname.Text) & "%'"
                strSql = strSql & "     Or upper(b.sunamee)   Like '" & UCase(txtQrySuname.Text) & "%')"
            Else
                strSql = strSql & " AND   (b.sunamek   Like '" & txtQrySuname.Text & "%'"
                strSql = strSql & "     Or b.sunamee   Like '" & txtQrySuname.Text & "%')"
            End If
    End Select
    
    ssSuga.MaxRows = 0
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    ssSuga.MaxRows = adoSet.RecordCount
    
    Do Until adoSet.EOF
        ssSuga.Row = ssSuga.DataRowCnt + 1
        sSugba = adoSet.Fields("SugbA").Value & ""
        ssSuga.Col = 1: ssSuga.Text = Trim(adoSet.Fields("Sucode").Value & "")
        ssSuga.Col = 2: ssSuga.Text = sSugba & Trim(adoSet.Fields("Sunamee").Value & "")
        ssSuga.Col = 3: ssSuga.Text = sSugba & Trim(adoSet.Fields("SunameK").Value & "")
        ssSuga.Col = 4: ssSuga.Text = Trim(adoSet.Fields("Bcode").Value & "")
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)

        

    
End Sub

Private Sub cmdQryOk_Click()
    Dim sBunCode        As String
    Dim sSugba       As String * 2
    
    panelBun.Visible = False
    If Trim(txtBname.Text) = "" Then Exit Sub

    sBunCode = Left(txtBname.Text, 2)
    
    strSql = ""
    strSql = strSql & " SELECT a.Sucode, b.Sunamee, b.Sunamek, b.Bcode, a.SugbA"
    strSql = strSql & " FROM   TWBAS_TSUGA a,"
    strSql = strSql & "        TWBAS_NSUGA b "
    strSql = strSql & " WHERE  a.Bun    =  '" & Trim(sBunCode) & "'"
    strSql = strSql & " AND    a.Sunext = b.Sunext(+)"
    
    ssSuga.MaxRows = 0
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    ssSuga.MaxRows = adoSet.RecordCount
    
    Do Until adoSet.EOF
        ssSuga.Row = ssSuga.DataRowCnt + 1
        ssSuga.Col = 1: ssSuga.Text = Trim(adoSet.Fields("Sucode").Value & "")
        sSugba = adoSet.Fields("SugbA").Value & ""
        ssSuga.Col = 1: ssSuga.Text = Trim(adoSet.Fields("Sucode").Value & "")
        ssSuga.Col = 1: ssSuga.Text = Trim(adoSet.Fields("Sucode").Value & "")
        ssSuga.Col = 2: ssSuga.Text = sSugba & Trim(adoSet.Fields("Sunamee").Value & "")
        ssSuga.Col = 3: ssSuga.Text = sSugba & Trim(adoSet.Fields("SunameK").Value & "")
        ssSuga.Col = 4: ssSuga.Text = Trim(adoSet.Fields("Bcode").Value & "")
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)

    
End Sub

Private Sub Form_Load()
    vaTabPro1.ActiveTab = 1
    
End Sub

Private Sub mnuBun_Click()
    Dim sText       As String
    Dim sRowid      As String
    Dim NodeX       As Node
    
    
    Screen.MousePointer = vbHourglass
    
    DoEvents
    panelBun.Visible = True
    panelBun.ZOrder 0
    
    DoEvents
    tvBun.Nodes.Clear
    
    strSql = ""
    strSql = strSql & " SELECT Cgb, MAX(RowID) RwID"
    strSql = strSql & " FROM   TWBAS_Bun"
    strSql = strSql & " GROUP  BY Cgb"
    
    If False = adoSetOpen(strSql, adoSet) Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    Do Until adoSet.EOF
        sRowid = adoSet.Fields("RwID").Value & ""
        sText = Trim(adoSet.Fields("Cgb").Value & "")
        Set NodeX = tvBun.Nodes.Add(, , "A1" & sRowid, sText)
        GoSub Load_SubCode
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    
    tvBun.SetFocus
    Screen.MousePointer = vbDefault
    
    Exit Sub
    
Load_SubCode:
    Dim adoSubCode1     As ADODB.Recordset
    Dim sSubText1       As String
    Dim sSubText2       As String
    
    
    strSql = ""
    strSql = strSql & " SELECT a.*, a.RowID"
    strSql = strSql & " FROM   TWBAS_Bun a"
    strSql = strSql & " WHERE  a.cgb = '" & tvBun.Nodes("A1" & sRowid).Text & "'"
    If False = adoSetOpen(strSql, adoSubCode1) Then Return
    
    Do Until adoSubCode1.EOF
        sSubText1 = "B2" & adoSubCode1.Fields("RowID")
        sSubText2 = adoSubCode1.Fields("Bun").Value & "." & adoSubCode1.Fields("Bname").Value & ""
        Set NodeX = tvBun.Nodes.Add("A1" & sRowid, tvwChild, sSubText1, sSubText2)
        adoSubCode1.MoveNext
    Loop
    Call adoSetClose(adoSubCode1)
    
    Return
    

End Sub

Private Sub mnuExit_Click()
    Unload Me
    
End Sub

Private Sub Option1_Click()
    
    Me.ssSuga.ReDraw = False
    Me.ssSuga.ColWidth(2) = 33.75
    Me.ssSuga.ColWidth(3) = 33.75
        
    ssSuga.Col = 2: ssSuga.ColHidden = False
    ssSuga.Col = 3: ssSuga.ColHidden = True
    Me.ssSuga.ReDraw = True
    
End Sub

Private Sub Option2_Click()
    
    Me.ssSuga.ReDraw = False
    Me.ssSuga.ColWidth(2) = 33.75
    Me.ssSuga.ColWidth(3) = 33.75
        
    ssSuga.Col = 2: ssSuga.ColHidden = True
    ssSuga.Col = 3: ssSuga.ColHidden = False
    Me.ssSuga.ReDraw = True

End Sub


Private Sub ssSuga_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim sSuname     As String
    
    Select Case Row
        Case 0
            ssSuga.Row = 1
            ssSuga.Row2 = ssSuga.DataRowCnt
            ssSuga.Col = 1
            ssSuga.Col2 = ssSuga.DataColCnt
            ssSuga.SortBy = SS_SORT_BY_ROW
            ssSuga.SortKey(1) = Col
            If ssSuga.SortKeyOrder(1) = SortKeyOrderDescending Then
                ssSuga.SortKeyOrder(1) = SS_SORT_ORDER_ASCENDING
            Else
                ssSuga.SortKeyOrder(1) = SS_SORT_ORDER_DESCENDING
            End If
            ssSuga.Action = SS_ACTION_SORT
        Case Is > 0
            
            ssSuga.Row = Row
            ssSuga.Col = 2
            
            sSuname = Mid(Trim(ssSuga.Text), 3, Len(Trim(ssSuga.Text)) - 2)
            If vbYes = MsgBox(sSuname & " 을(를) 선택하시겠습니까?", vbYesNo + vbQuestion, "보험수가 선택Box") Then
                ssSuga.Col = 1
                Call SetWindowText(hWndReturn, Trim(ssSuga.Text))
                Unload Me
            End If

        Case Else
    End Select
    
End Sub

Private Sub tvBun_NodeClick(ByVal Node As MSComctlLib.Node)
    
    If Left(Node.Key, 2) = "B2" Then
        txtCgb.Text = tvBun.SelectedItem.Parent.Text
        txtBname.Text = tvBun.SelectedItem.Text
    Else
        txtCgb.Text = tvBun.SelectedItem.Text
        txtBname.Text = ""
    End If
    

End Sub

Private Sub txtQrySuname_GotFocus()
    
    txtQrySuname.BackColor = RGB(255, 255, 220)
    
End Sub

Private Sub txtQrySuname_LostFocus()
    txtQrySuname.BackColor = RGB(255, 255, 255)
End Sub

Private Sub txtSucode_GotFocus()
    txtQrySuname.BackColor = RGB(255, 255, 220)
End Sub

Private Sub txtSucode_LostFocus()
    
    txtQrySuname.BackColor = RGB(255, 255, 255)
    
End Sub

Private Sub vaTabPro1_TabShown(ActiveTab As Integer)
    
    Select Case ActiveTab
        Case 0: txtSucode.SetFocus
        Case 1: txtQrySuname.SetFocus
    End Select
    
End Sub
