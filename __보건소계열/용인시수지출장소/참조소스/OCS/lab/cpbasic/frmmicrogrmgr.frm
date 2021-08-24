VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Begin VB.Form frmMicroGrmgr 
   Caption         =   "약제코드 Grouping Data 관리화면"
   ClientHeight    =   6450
   ClientLeft      =   645
   ClientTop       =   1215
   ClientWidth     =   10785
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
   ScaleHeight     =   6450
   ScaleWidth      =   10785
   WindowState     =   2  '최대화
   Begin MSComctlLib.ImageList imgTree 
      Left            =   5280
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMicroGrmgr.frx":0000
            Key             =   "N"
            Object.Tag             =   "Normal"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMicroGrmgr.frx":27B4
            Key             =   "O"
            Object.Tag             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMicroGrmgr.frx":4F68
            Key             =   "F"
            Object.Tag             =   "Full"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMicroGrmgr.frx":771C
            Key             =   "S"
            Object.Tag             =   "Sylinge"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvGroup 
      Height          =   7395
      Left            =   240
      TabIndex        =   1
      Top             =   180
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   13044
      _Version        =   393217
      Style           =   7
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      ImageList       =   "imgTree"
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin FPSpreadADO.fpSpread ssAntiList 
      Height          =   7455
      Left            =   5760
      TabIndex        =   0
      Top             =   180
      Width           =   5295
      _Version        =   196608
      _ExtentX        =   9340
      _ExtentY        =   13150
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
      MaxCols         =   10
      ScrollBars      =   2
      SpreadDesigner  =   "frmMicroGrmgr.frx":8000
      Appearance      =   1
   End
   Begin VB.Menu mnuQuit 
      Caption         =   "Quit"
   End
   Begin VB.Menu mnuJob 
      Caption         =   "Job"
      Begin VB.Menu mnuNew 
         Caption         =   "신규입력"
      End
      Begin VB.Menu mnuDel 
         Caption         =   "삭제"
      End
   End
End
Attribute VB_Name = "frmMicroGrmgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim sText       As String
    Dim sRowid      As String
    Dim NodeX       As Node
    
    ssAntiList.RowHeight(-1) = 12
    GoSub TreeView_Select
    GoSub Get_AntiList
    Exit Sub
    
    
TreeView_Select:
    tvGroup.Nodes.Clear
    Set NodeX = tvGroup.Nodes.Add(, , "A0", "세균그룹코드", "N", "F")
    
    strSql = ""
    strSql = strSql & " SELECT Grp_Code, MAX(RowID) RWID"
    strSql = strSql & " FROM   TW_MIS_EXAM.TWEXAM_ANTIGROUP"
    strSql = strSql & " GROUP  BY Grp_Code"
    
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    
    Do Until adoSet.EOF
        sRowid = adoSet.Fields("RwID").Value & ""
        sText = Trim(adoSet.Fields("Grp_Code").Value & "")
        Set NodeX = tvGroup.Nodes.Add("A0", tvwChild, "A1" & sRowid, sText, "N", "F")
        GoSub Load_SubCode
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    tvGroup.Nodes("A0").Expanded = True
    
    Return
    
    
Load_SubCode:
    Dim adoSubCode1     As ADODB.Recordset
    Dim sSubText1       As String
    Dim sSubText2       As String
    
    
    strSql = ""
    strSql = strSql & " SELECT a.Anti_Code, a.RowID, b.Codenm"
    strSql = strSql & " FROM   TW_MIS_EXAM.TWEXAM_AntiGroup a,"
    strSql = strSql & "        TWEXAM_ANTILIST  b "
    strSql = strSql & " WHERE  a.Grp_Code  = '" & tvGroup.Nodes("A1" & sRowid).Text & "'"
    strSql = strSql & " AND    a.Anti_Code IS NOT NULL"
    strSql = strSql & " AND    a.Anti_Code = b.Codeky(+)"
    
    If False = adoSetOpen(strSql, adoSubCode1) Then Return
    
    Do Until adoSubCode1.EOF
        sSubText1 = "B2" & adoSubCode1.Fields("RowID")
        sSubText2 = Trim(adoSubCode1.Fields("Anti_Code").Value) & "." & _
                    StrConv(Trim(adoSubCode1.Fields("Codenm").Value & ""), vbProperCase)
        Set NodeX = tvGroup.Nodes.Add("A1" & sRowid, tvwChild, sSubText1, sSubText2, "S")
        adoSubCode1.MoveNext
    Loop
    Call adoSetClose(adoSubCode1)
    
    Return

Get_AntiList:
    strSql = ""
    strSql = strSql & " SELECT *"
    strSql = strSql & " FROM   TWEXAM_ANTILIST"
    strSql = strSql & " ORDER  By Seqno"
    
    ssAntiList.MaxRows = 0
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    ssAntiList.MaxRows = adoSet.RecordCount
    Do Until adoSet.EOF
        ssAntiList.Row = ssAntiList.DataRowCnt + 1
        ssAntiList.Col = 1:  ssAntiList.Text = Trim(adoSet.Fields("Codeky").Value & "")
        ssAntiList.Col = 2:  ssAntiList.Text = StrConv(Trim(adoSet.Fields("Codenm").Value & ""), vbProperCase)
        ssAntiList.Col = 3:  ssAntiList.Text = Trim(adoSet.Fields("Seqno").Value & "")
        'ssAntiList.Col = 4:  ssAntiList.Text = Trim(adoSet.Fields("Letter").Value & "")
        ssAntiList.Col = 5:  ssAntiList.Text = Trim(adoSet.Fields("Potency").Value & "")
        ssAntiList.Col = 6:  ssAntiList.Text = Trim(adoSet.Fields("Lozone").Value & "")
        ssAntiList.Col = 7:  ssAntiList.Text = Trim(adoSet.Fields("Hizone").Value & "")
        ssAntiList.Col = 8:  ssAntiList.Text = Trim(adoSet.Fields("Lomic").Value & "")
        ssAntiList.Col = 9:  ssAntiList.Text = Trim(adoSet.Fields("Himic").Value & "")
        ssAntiList.Col = 10: ssAntiList.Text = Trim(adoSet.Fields("Source").Value & "")
        
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    Return
    

End Sub

Private Sub mnuDel_Click()
    
    Select Case Left(tvGroup.SelectedItem.Key, 2)
        Case "A0": GoSub A0_DELETE
        Case "A1": GoSub A1_DELETE
        Case "B2": GoSub B2_DELETE
    End Select
    Exit Sub
    
A0_DELETE:
    Call MsgBox("모든 Data 를 지울수는 없습니다!.", vbCritical + vbOKOnly, "경고 Message")
    Return
    

A1_DELETE:
    If vbNo = MsgBox("GroupCode: " & Trim(tvGroup.SelectedItem.Text) & " 의 하부코드를 모두 삭제하시겠습니까?", _
                      vbYesNo + vbQuestion, "삭제확인 Box") Then Exit Sub
    strSql = ""
    strSql = strSql & " DELETE "
    strSql = strSql & " FROM   TW_MIS_EXAM.TWEXAM_ANTIGROUP"
    strSql = strSql & " WHERE  GRP_CODE  = '" & Trim(tvGroup.SelectedItem.Text) & "'"
    Call adoExec(strSql)
    Call tvGroup.Nodes.Remove(tvGroup.SelectedItem.Index)
    
    Return
    
B2_DELETE:
    Dim sCapString      As String
    Dim nFirst          As Integer
    Dim sKey            As String
    
    sCapString = Trim(tvGroup.SelectedItem.Text)
    
    nFirst = InStr(1, sCapString, ".", vbTextCompare)
    sCapString = Left(sCapString, nFirst - 1)
    
    If vbNo = MsgBox("약제코드: " & Trim(tvGroup.SelectedItem.Text) & " 를 삭제하시겠습니까?", _
                      vbYesNo + vbQuestion, "삭제확인 Box") Then Exit Sub
    strSql = ""
    strSql = strSql & " DELETE "
    strSql = strSql & " FROM   TW_MIS_EXAM.TWEXAM_ANTIGROUP"
    strSql = strSql & " WHERE  GRP_CODE  = '" & Trim(tvGroup.SelectedItem.Parent.Text) & "'"
    strSql = strSql & " AND    ANTI_CODE = '" & sCapString & "'"
    Call adoExec(strSql)
    
    Call tvGroup.Nodes.Remove(tvGroup.SelectedItem.Index)

    Return
    
    
End Sub

Private Sub mnuNew_Click()
    Dim sGrpCode    As String
    Dim sKey        As String
    Dim sText       As String
    Dim sRowid      As String
    Dim NodeX       As Node
    
    sGrpCode = InputBox("새로운 Group Code 를 입력하세요!..", "신규 GroupCode 입력 Box")
    
    If sGrpCode = "" Then Exit Sub
    
    strSql = ""
    strSql = strSql & " INSERT INTO TW_MIS_EXAM.TWEXAM_ANTIGROUP"
    strSql = strSql & "       (GRP_Code)"
    strSql = strSql & " VALUES('" & sGrpCode & "')"
    Call adoExec(strSql)
    
    GoSub TreeView_Select
    Exit Sub
    
    
TreeView_Select:
    tvGroup.Nodes.Clear
    Set NodeX = tvGroup.Nodes.Add(, , "A0", "세균그룹코드", "N", "F")
    
    strSql = ""
    strSql = strSql & " SELECT Grp_Code, MAX(RowID) RWID"
    strSql = strSql & " FROM   TW_MIS_EXAM.TWEXAM_ANTIGROUP"
    strSql = strSql & " GROUP  BY Grp_Code"
    
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    
    Do Until adoSet.EOF
        sRowid = adoSet.Fields("RwID").Value & ""
        sText = Trim(adoSet.Fields("Grp_Code").Value & "")
        Set NodeX = tvGroup.Nodes.Add("A0", tvwChild, "A1" & sRowid, sText, "N", "F")
        GoSub Load_SubCode
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    tvGroup.Nodes("A0").Expanded = True
    
    Return
    
    
Load_SubCode:
    Dim adoSubCode1     As ADODB.Recordset
    Dim sSubText1       As String
    Dim sSubText2       As String
    
    
    strSql = ""
    strSql = strSql & " SELECT a.Anti_Code, a.RowID, b.Codenm"
    strSql = strSql & " FROM   TW_MIS_EXAM.TWEXAM_AntiGroup a,"
    strSql = strSql & "        TWEXAM_ANTILIST  b "
    strSql = strSql & " WHERE  a.Grp_Code  = '" & tvGroup.Nodes("A1" & sRowid).Text & "'"
    strSql = strSql & " AND    a.Anti_Code IS NOT NULL"
    strSql = strSql & " AND    a.Anti_Code = b.Codeky(+)"
    
    If False = adoSetOpen(strSql, adoSubCode1) Then Return
    
    Do Until adoSubCode1.EOF
        sSubText1 = "B2" & adoSubCode1.Fields("RowID")
        sSubText2 = Trim(adoSubCode1.Fields("Anti_Code").Value) & "." & _
                    StrConv(Trim(adoSubCode1.Fields("Codenm").Value & ""), vbProperCase)
        Set NodeX = tvGroup.Nodes.Add("A1" & sRowid, tvwChild, sSubText1, sSubText2, "S")
        adoSubCode1.MoveNext
    Loop
    Call adoSetClose(adoSubCode1)
    
    Return


End Sub

Private Sub mnuQuit_Click()
    Unload Me
    
End Sub

Private Sub ssAntiList_DblClick(ByVal Col As Long, ByVal Row As Long)
    
    
    Select Case Left(tvGroup.SelectedItem.Key, 2)
        Case "A0": MsgBox "약제코드를 입력할 Group Code 를 선택하세요!"
        Case "A1": GoSub AddItem_AntiList
        Case "B2": tvGroup.SelectedItem.Parent.Selected = True
    End Select
    Exit Sub
    
    
AddItem_AntiList:
    Dim sAnti_Code      As String
    Dim sZoneGb         As String
    Dim nLo             As Long
    Dim nHi             As Long
    Dim sSource         As String
    Dim sKey            As String
    Dim adoRecall       As ADODB.Recordset
        
    If vbNo = MsgBox("새로운 코드를 입력하시겠습니까?", vbYesNo + vbQuestion, "입력확인") Then
        Return
    End If
        
    ssAntiList.Row = Row
    ssAntiList.Col = 1: sAnti_Code = Trim(ssAntiList.Text)
    frmMicroZone.Show vbModal
    
    sZoneGb = Trim(frmMicroGrmgr.tvGroup.Tag)
    
    Select Case Trim(frmMicroGrmgr.tvGroup.Tag)
        Case "Zone"
            sZoneGb = "ZONE"
            ssAntiList.Col = 6:  nLo = Val(ssAntiList.Text)
            ssAntiList.Col = 7:  nHi = Val(ssAntiList.Text)
            ssAntiList.Col = 10: sSource = Trim(ssAntiList.Text)
        Case "Mic"
            sZoneGb = "MIC"
            ssAntiList.Col = 8:  nLo = Val(ssAntiList.Text)
            ssAntiList.Col = 9:  nHi = Val(ssAntiList.Text)
            ssAntiList.Col = 10: sSource = Trim(ssAntiList.Text)
        Case "Cancel": Exit Sub
            
    End Select
        
    strSql = ""
    strSql = strSql & " INSERT INTO TW_MIS_EXAM.TWEXAM_ANTIGROUP"
    strSql = strSql & "       (Grp_Code, Anti_Code, ZoneGb, Lo, Hi, Source)"
    strSql = strSql & " VALUES('" & tvGroup.SelectedItem.Text & "',"
    strSql = strSql & "        '" & sAnti_Code & "',"
    strSql = strSql & "        '" & sZoneGb & "',"
    strSql = strSql & "         " & nLo & ","
    strSql = strSql & "         " & nHi & ","
    strSql = strSql & "        '" & Trim(sSource) & "')"
    Call adoExec(strSql)
    
    
    
    strSql = ""
    strSql = strSql & " SELECT a.RowiD, a.anti_Code, b.Codenm "
    strSql = strSql & " FROM   TW_MIS_EXAM.TWEXAM_ANTIGROUP a,"
    strSql = strSql & "        TWEXAM_ANTILIST  b "
    strSql = strSql & " WHERE  a.Grp_code  = '" & tvGroup.SelectedItem.Text & "'"
    strSql = strSql & " AND    a.Anti_Code = '" & sAnti_Code & "'"
    strSql = strSql & " AND    a.Anti_Code = b.Codeky(+)"
    If False = adoSetOpen(strSql, adoRecall) Then Return
    
    sKey = tvGroup.SelectedItem.Key
    Call tvGroup.Nodes.Add(sKey, tvwChild, "B2" & adoRecall.Fields("RowID").Value, _
                           Trim(adoRecall.Fields("anti_Code").Value) & "." & _
                           StrConv(Trim(adoRecall.Fields("Codenm").Value & ""), vbProperCase), "S")
    Call adoSetClose(adoRecall)
    Return
    
    
End Sub

Private Sub tvGroup_AfterLabelEdit(Cancel As Integer, NewString As String)
    
    sNewGroupString = NewString
    Select Case Left(tvGroup.SelectedItem.Key, 2)
        Case "A0": Cancel = 1
        Case "A1": GoSub Group_Code_Update
        Case "B2": Cancel = 1
        Case Else: Exit Sub
    End Select
    Exit Sub
    

Group_Code_Update:
    strSql = ""
    strSql = strSql & " UPDATE TW_MIS_EXAM.TWEXAM_ANTIGROUP"
    strSql = strSql & " SET    Grp_Code  =  '" & sNewGroupString & "'"
    strSql = strSql & " WHERE  Grp_Code  =  '" & Trim(sOldGroupString) & "'"
    Call adoExec(strSql)
    Return
    
End Sub

Private Sub tvGroup_BeforeLabelEdit(Cancel As Integer)
    
    sOldGroupString = tvGroup.SelectedItem.Text
    
    'Select Case Left(tvGroup.SelectedItem.Key, 2)
    '    Case "A0": tvGroup.LabelEdit = tvwManual
    '    Case "A1": tvGroup.LabelEdit = tvwAutomatic
    '               sOldGroupString = tvGroup.SelectedItem.Text
    '    Case "B2": tvGroup.LabelEdit = tvwManual
    '    Case Else: Exit Sub
    'End Select
    
End Sub

Private Sub tvGroup_Expand(ByVal Node As MSComctlLib.Node)
    
    If Node.Expanded = True Then
        Node.ExpandedImage = "F"
    Else
        Node.ExpandedImage = "N"
    End If

    
    
    
End Sub

Private Sub tvGroup_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button <> 2 Then Exit Sub
    
    If Left(tvGroup.SelectedItem.Key, 2) = "A0" Then
        mnuNew.Visible = True
    Else
        mnuNew.Visible = False
    End If
    
    PopupMenu mnuJob
    

End Sub

