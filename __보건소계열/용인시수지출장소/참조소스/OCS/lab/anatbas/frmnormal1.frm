VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Begin VB.Form frmNormal1 
   Caption         =   "Form1"
   ClientHeight    =   6840
   ClientLeft      =   1290
   ClientTop       =   1005
   ClientWidth     =   10245
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
   ScaleHeight     =   6840
   ScaleWidth      =   10245
   WindowState     =   2  '최대화
   Begin Threed.SSCommand cmdIndexReset 
      Height          =   315
      Left            =   1860
      TabIndex        =   5
      Top             =   120
      Width           =   315
      _Version        =   65536
      _ExtentX        =   556
      _ExtentY        =   556
      _StockProps     =   78
      Caption         =   "&R"
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   540
      Width           =   4815
   End
   Begin VB.TextBox txtCode 
      Height          =   375
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   540
      Width           =   1155
   End
   Begin VB.TextBox txtSearch 
      Height          =   330
      Left            =   60
      TabIndex        =   1
      Top             =   120
      Width           =   1755
   End
   Begin MSComctlLib.TreeView tvRef 
      Height          =   3075
      Left            =   60
      TabIndex        =   0
      Top             =   540
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   5424
      _Version        =   393217
      LineStyle       =   1
      Style           =   7
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin FPSpreadADO.fpSpread ssRefData 
      Height          =   4875
      Left            =   5280
      TabIndex        =   2
      Top             =   960
      Width           =   6615
      _Version        =   196608
      _ExtentX        =   11668
      _ExtentY        =   8599
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
      MaxCols         =   10
      ScrollBars      =   2
      SpreadDesigner  =   "frmNormal1.frx":0000
      Appearance      =   1
   End
   Begin VB.Menu mnuQuit 
      Caption         =   "Quit"
   End
   Begin VB.Menu mnuJob 
      Caption         =   "작업"
      Begin VB.Menu mnuNew 
         Caption         =   "신규추가"
      End
      Begin VB.Menu mnuUpdate 
         Caption         =   "수정"
      End
   End
End
Attribute VB_Name = "frmNormal1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdIndexReset_Click()
    
    nIndex = 0
    txtSearch.SetFocus
    
End Sub

Private Sub Form_Load()
    Dim sRowID      As String
    Dim sText       As String
    Dim NodeX       As Node
    
    
    ssRefData.RowHeight(-1) = 12
    
    tvRef.Nodes.Clear
    
    strSql = ""
    strSql = strSql & " SELECT a.*, a.RowID"
    strSql = strSql & " FROM   TWEXAM_SPECODE a"
    strSql = strSql & " WHERE  a.Codegu = '12'"
    strSql = strSql & " Order  By Codeky"
    
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    
    Do Until adoSet.EOF
        sRowID = adoSet.Fields("RowID").Value & ""
        sText = Trim(adoSet.Fields("Codeky").Value & "") & ". " & _
                Trim(adoSet.Fields("Codenm").Value & "")
        Set NodeX = tvRef.Nodes.Add(, , "A1" & sRowID, sText)
        GoSub Load_SubCode
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    Exit Sub
    
Load_SubCode:
    Dim adoSubCode1     As New ADODB.Recordset
    Dim sSubText1       As String
    Dim sSubText2       As String
    
    
    strSql = ""
    strSql = strSql & " SELECT a.Itemnm, a.Codeky, a.RowID "
    strSql = strSql & " FROM   TWEXAM_iTemML a"
    strSql = strSql & " WHERE  SubStr(a.Codeky, 1,2) = '" & Left(tvRef.Nodes("A1" & sRowID).Text, 2) & "'"
    strSql = strSql & " ORDER  BY Codeky"
    If False = adoSetOpen(strSql, adoSubCode1) Then Return
    
    Do Until adoSubCode1.EOF
        sSubText1 = "B2" & adoSubCode1.Fields("RowID")
        sSubText2 = adoSubCode1.Fields("ItemNM").Value & ""
        Set NodeX = tvRef.Nodes.Add("A1" & sRowID, tvwChild, sSubText1, sSubText2)
        
        tvRef.Nodes.Item(sSubText1).Tag = adoSubCode1.Fields("Codeky").Value & ""
        adoSubCode1.MoveNext
    Loop
    Call adoSetClose(adoSubCode1)
    
    Return

End Sub

Private Sub mnuNew_Click()
    
    If Trim(frmNormal1.txtCode.Text) = "" Then Exit Sub
        
    frmRefNew.Show vbModal
    
    
    strSql = ""
    strSql = strSql & " SELECT a.*, a.RowID, "
    strSql = strSql & "        TO_CHAR(a.appDate,'YYYY-MM-DD') appDate"
    strSql = strSql & " FROM   TWEXAM_RefData a"
    strSql = strSql & " WHERE  a.iTemCode  =  '" & tvRef.SelectedItem.Tag & "'"
    strSql = strSql & " ORDER  BY a.appDate DESC "
    
    ssRefData.MaxRows = 0
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    ssRefData.MaxRows = adoSet.RecordCount
    
    
    Do Until adoSet.EOF
        ssRefData.Row = ssRefData.DataRowCnt + 1
        ssRefData.Col = 1:  ssRefData.Text = adoSet.Fields("RowID").Value & ""
        ssRefData.Col = 2:  ssRefData.Text = adoSet.Fields("appDate").Value & ""
        ssRefData.Col = 3:  ssRefData.Text = adoSet.Fields("ItemCode").Value & ""
        ssRefData.Col = 4:  ssRefData.Text = adoSet.Fields("appGubun").Value & ""
        ssRefData.Col = 5:  ssRefData.Text = adoSet.Fields("ageMin").Value & ""
        ssRefData.Col = 6:  ssRefData.Text = adoSet.Fields("ageMax").Value & ""
        ssRefData.Col = 7:  ssRefData.Text = Trim(adoSet.Fields("M_min").Value & "")
        ssRefData.Col = 8:  ssRefData.Text = Trim(adoSet.Fields("M_max").Value & "")
        ssRefData.Col = 9:  ssRefData.Text = Trim(adoSet.Fields("F_min").Value & "")
        ssRefData.Col = 10: ssRefData.Text = Trim(adoSet.Fields("F_max").Value & "")
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
End Sub

Private Sub mnuQuit_Click()
    Unload Me
    
End Sub

Private Sub mnuUpdate_Click()
    If Trim(frmNormal1.txtCode.Text) = "" Then Exit Sub
    
    frmRefNew.Show vbModal
    
    strSql = ""
    strSql = strSql & " SELECT a.*, a.RowID, "
    strSql = strSql & "        TO_CHAR(a.appDate,'YYYY-MM-DD') appDate"
    strSql = strSql & " FROM   TWEXAM_RefData a"
    strSql = strSql & " WHERE  a.iTemCode  =  '" & tvRef.SelectedItem.Tag & "'"
    strSql = strSql & " ORDER  BY a.appDate DESC "
    
    ssRefData.MaxRows = 0
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    ssRefData.MaxRows = adoSet.RecordCount
    
    
    Do Until adoSet.EOF
        ssRefData.Row = ssRefData.DataRowCnt + 1
        ssRefData.Col = 1:  ssRefData.Text = adoSet.Fields("RowID").Value & ""
        ssRefData.Col = 2:  ssRefData.Text = adoSet.Fields("appDate").Value & ""
        ssRefData.Col = 3:  ssRefData.Text = adoSet.Fields("ItemCode").Value & ""
        ssRefData.Col = 4:  ssRefData.Text = adoSet.Fields("appGubun").Value & ""
        ssRefData.Col = 5:  ssRefData.Text = adoSet.Fields("ageMin").Value & ""
        ssRefData.Col = 6:  ssRefData.Text = adoSet.Fields("ageMax").Value & ""
        ssRefData.Col = 7:  ssRefData.Text = Trim(adoSet.Fields("M_min").Value & "")
        ssRefData.Col = 8:  ssRefData.Text = Trim(adoSet.Fields("M_max").Value & "")
        ssRefData.Col = 9:  ssRefData.Text = Trim(adoSet.Fields("F_min").Value & "")
        ssRefData.Col = 10: ssRefData.Text = Trim(adoSet.Fields("F_max").Value & "")
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)

End Sub

Private Sub SSCommand1_Click()

End Sub

Private Sub ssRefData_DblClick(ByVal Col As Long, ByVal Row As Long)
    
    If Row = 0 Then Exit Sub
    
    
    ssRefData.Row = Row
    ssRefData.Col = 1
    
    If vbNo = MsgBox("해당 행의 Data 를 삭제하시겠습니까?", vbYesNo + vbQuestion, "삭제 확인MessageBox") Then
        Exit Sub
    End If
    
    strSql = ""
    strSql = strSql & " Delete "
    strSql = strSql & " From   TWEXAM_RefData"
    strSql = strSql & " WHERE  RowID  =  '" & ssRefData.Text & "'"
    
    adoConnect.BeginTrans
    If adoExecute(strSql) Then
        adoConnect.CommitTrans
        GoSub Delete_Reset_Routine
    Else
        adoConnect.RollbackTrans
    End If
    Exit Sub
    
    
Delete_Reset_Routine:
    strSql = ""
    strSql = strSql & " SELECT a.*, a.RowID, "
    strSql = strSql & "        TO_CHAR(a.appDate,'YYYY-MM-DD') appDate"
    strSql = strSql & " FROM   TWEXAM_RefData a"
    strSql = strSql & " WHERE  a.iTemCode  =  '" & tvRef.SelectedItem.Tag & "'"
    strSql = strSql & " ORDER  BY a.appDate DESC "
    
    ssRefData.MaxRows = 0
    If False = adoSetOpen(strSql, adoSet) Then Return
    ssRefData.MaxRows = adoSet.RecordCount
    
    
    Do Until adoSet.EOF
        ssRefData.Row = ssRefData.DataRowCnt + 1
        ssRefData.Col = 1:  ssRefData.Text = adoSet.Fields("RowID").Value & ""
        ssRefData.Col = 2:  ssRefData.Text = adoSet.Fields("appDate").Value & ""
        ssRefData.Col = 3:  ssRefData.Text = adoSet.Fields("ItemCode").Value & ""
        ssRefData.Col = 4:  ssRefData.Text = adoSet.Fields("appGubun").Value & ""
        ssRefData.Col = 5:  ssRefData.Text = adoSet.Fields("ageMin").Value & ""
        ssRefData.Col = 6:  ssRefData.Text = adoSet.Fields("ageMax").Value & ""
        ssRefData.Col = 7:  ssRefData.Text = Trim(adoSet.Fields("M_min").Value & "")
        ssRefData.Col = 8:  ssRefData.Text = Trim(adoSet.Fields("M_max").Value & "")
        ssRefData.Col = 9:  ssRefData.Text = Trim(adoSet.Fields("F_min").Value & "")
        ssRefData.Col = 10: ssRefData.Text = Trim(adoSet.Fields("F_max").Value & "")
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    Return
End Sub

Private Sub ssRefData_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button = 2 Then
        hWndReturn = ssRefData.hwnd
        mnuUpdate.Visible = True
        mnuNew.Visible = False
        PopupMenu mnuJob
    End If
    
End Sub

Private Sub tvRef_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button = 2 Then
        If Left(tvRef.SelectedItem.Key, 2) = "B2" Then
            hWndReturn = tvRef.hwnd
            mnuNew.Visible = True
            mnuUpdate.Visible = False
            PopupMenu mnuJob
        End If
    End If
    
    
End Sub

Private Sub tvRef_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim sCodeky     As String * 8
    
    Select Case Trim(Left(Node.Key, 2))
        Case "A1": Exit Sub
        Case "B2": sCodeky = tvRef.SelectedItem.Tag
        Case Else: Exit Sub
    End Select
        
    'MsgBox tvRef.SelectedItem.Tag
    
    txtCode.Text = sCodeky
    txtName.Text = Node.Text
    
    strSql = ""
    strSql = strSql & " SELECT a.*, a.RowID, "
    strSql = strSql & "        TO_CHAR(a.appDate,'YYYY-MM-DD') appDate"
    strSql = strSql & " FROM   TWEXAM_RefData a"
    strSql = strSql & " WHERE  a.iTemCode  =  '" & sCodeky & "'"
    strSql = strSql & " ORDER  BY a.appDate DESC "
    
    ssRefData.MaxRows = 0
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    ssRefData.MaxRows = adoSet.RecordCount
    
    
    Do Until adoSet.EOF
        ssRefData.Row = ssRefData.DataRowCnt + 1
        ssRefData.Col = 1:  ssRefData.Text = adoSet.Fields("RowID").Value & ""
        ssRefData.Col = 2:  ssRefData.Text = adoSet.Fields("appDate").Value & ""
        ssRefData.Col = 3:  ssRefData.Text = adoSet.Fields("ItemCode").Value & ""
        ssRefData.Col = 4:  ssRefData.Text = adoSet.Fields("appGubun").Value & ""
        ssRefData.Col = 5:  ssRefData.Text = adoSet.Fields("ageMin").Value & ""
        ssRefData.Col = 6:  ssRefData.Text = adoSet.Fields("ageMax").Value & ""
        ssRefData.Col = 7:  ssRefData.Text = Trim(adoSet.Fields("M_min").Value & "")
        ssRefData.Col = 8:  ssRefData.Text = Trim(adoSet.Fields("M_max").Value & "")
        ssRefData.Col = 9:  ssRefData.Text = Trim(adoSet.Fields("F_min").Value & "")
        ssRefData.Col = 10: ssRefData.Text = Trim(adoSet.Fields("F_max").Value & "")
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        GoSub Search_TreeView_Text
    End If
    
    Exit Sub
    
Search_TreeView_Text:

    
    nIndex = tvRef.SelectedItem.Index
    
    For i = nIndex + 1 To Me.tvRef.Nodes.Count
        If InStr(1, UCase(tvRef.Nodes(i).Text), UCase(txtSearch.Text), vbTextCompare) > 0 Then
            tvRef.SetFocus
            tvRef.Nodes(i).Expanded = True
            tvRef.Nodes(i).Selected = True
            txtSearch.SetFocus
            Exit For
        End If
    Next
    
    Return

End Sub

