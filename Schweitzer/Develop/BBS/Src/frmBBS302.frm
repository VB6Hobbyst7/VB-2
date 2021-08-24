VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frmBBS302 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   1  '단일 고정
   Caption         =   "Blood Split"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4065
   Icon            =   "frmBBS302.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   4065
   StartUpPosition =   1  '소유자 가운데
   Begin VB.CheckBox chkBar 
      BackColor       =   &H00DBE6E6&
      Caption         =   "바코드로 처리"
      Height          =   195
      Left            =   75
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   105
      Width           =   1575
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   2670
      Style           =   1  '그래픽
      TabIndex        =   4
      Tag             =   "128"
      Top             =   4110
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      Height          =   510
      Left            =   1350
      Style           =   1  '그래픽
      TabIndex        =   3
      Tag             =   "124"
      Top             =   4110
      Width           =   1320
   End
   Begin VB.CommandButton cmdGenerate 
      BackColor       =   &H00F4F0F2&
      Caption         =   "분획(&S)"
      Height          =   510
      Left            =   30
      Style           =   1  '그래픽
      TabIndex        =   2
      Tag             =   "15101"
      Top             =   4110
      Width           =   1320
   End
   Begin MSComctlLib.ListView lvwComp 
      Height          =   2835
      Left            =   45
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1170
      Width           =   3945
      _ExtentX        =   6959
      _ExtentY        =   5001
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Component"
         Object.Width           =   4762
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "코드"
         Object.Width           =   1235
      EndProperty
   End
   Begin VB.ComboBox cboComp 
      Height          =   300
      Left            =   1290
      Style           =   2  '드롭다운 목록
      TabIndex        =   1
      Top             =   780
      Width           =   2730
   End
   Begin VB.TextBox txtBldNo 
      Height          =   375
      Left            =   1275
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   390
      Width           =   2745
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   360
      Index           =   1
      Left            =   60
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   780
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   635
      BackColor       =   10392451
      ForeColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "Component"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   360
      Index           =   2
      Left            =   60
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   390
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   635
      BackColor       =   10392451
      ForeColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "혈액번호"
      Appearance      =   0
   End
End
Attribute VB_Name = "frmBBS302"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ObjDic As New clsDictionary
Private objComp As New clsDictionary

Private Sub cboComp_Click()
    
    If txtBldNo = "" Then Exit Sub
    
    Dim objGetSql   As New clsGetSqlStatement
    Dim DrRS        As New Recordset
    Dim RsColNm     As New Recordset
    Dim strCompocd  As String
    Dim strBldno    As String
    
'    objGetSql.setDbConn DBConn
    
    strCompocd = medGetP(cboComp.List(cboComp.ListIndex), 1, vbTab)
    
    If chkBar.value = 1 Then
        strBldno = Mid(txtBldNo, 1, 2) & "-" & _
                   Mid(txtBldNo, 3, 2) & "-" & _
                   Mid(Mid(txtBldNo, 5), 1, Len(Mid(txtBldNo, 5)) - 2)
    Else
        strBldno = txtBldNo
    End If
        
    
    Set DrRS = objGetSql.Get_BLOOD_INFORMATION(strBldno, strCompocd)
    
    If DrRS.EOF Then
        ObjDic.Clear
    Else
        With DrRS
            ObjDic.Clear
            ObjDic.FieldInialize "bldsrc,bldyy,bldno,compocd", "volumn,abo,rh,ptid,reserved,autofg,coldt," & _
                             "coltm,colid,available,expdt,exptm,entdt,enttm,entid,centercd,hosfg,donorid,donoraccdt"
            ObjDic.AddNew Join(Array(.Fields("bldsrc").value & "", .Fields("bldyy").value & "", .Fields("bldno").value & "", .Fields("compocd").value & ""), COL_DIV), _
                          Join(Array(.Fields("volumn").value & "", .Fields("abo").value & "", .Fields("rh").value & "", .Fields("ptid") & "", .Fields("reserved").value & "", _
                                     .Fields("autofg").value & "", .Fields("coldt").value & "", .Fields("coltm").value & "", .Fields("coltm").value & "", _
                                     .Fields("available").value & "", .Fields("expdt").value & "", .Fields("exptm").value & "", .Fields("entdt").value & "", _
                                     .Fields("enttm").value & "", .Fields("entid").value & "", .Fields("centercd").value & "", .Fields("hosfg").value & "", _
                                     .Fields("donorid").value & "", .Fields("donoraccdt").value & ""), COL_DIV)
        End With
        SendKeys "{tab}"
    End If
    
    Set DrRS = Nothing
    Set objGetSql = Nothing
End Sub

Private Sub cmdClear_Click()
    Dim iTmx As ListItem
    
    txtBldNo = ""
    cboComp.Clear
    
    Display_Comp
End Sub

Private Sub cmdExit_Click()
    Set objComp = Nothing
    Set ObjDic = Nothing
    Unload Me
End Sub
Private Function Redim_Cnt() As Long
    Dim iTmx As ListItem
    
    With lvwComp
        For Each iTmx In .ListItems
            If iTmx.Checked = True Then
                Redim_Cnt = Redim_Cnt + 1
            End If
        Next iTmx
    End With
End Function
Private Sub cmdGenerate_Click()
'------------
'혈액분획시작
'------------
    If Redim_Cnt = 0 Then Exit Sub
    
    Dim objExcute  As New clsBeginTrans
    Dim iTmx       As ListItem
    Dim strCompocd As String
    Dim SSQL       As String
    
'    objExcute.setDbConn DBConn
On Error GoTo SplitOut_Save_Error
    DBConn.BeginTrans
    
    With ObjDic
        .MoveFirst
        For Each iTmx In lvwComp.ListItems
            If iTmx.Checked = True Then
                strCompocd = iTmx.SubItems(1)
                SSQL = objExcute.Set_SplitInsertNew(.Fields("bldsrc"), .Fields("bldyy"), .Fields("bldno"), _
                                         strCompocd, .Fields("volumn"), .Fields("abo"), .Fields("rh"), _
                                        .Fields("ptid"), .Fields("reserved"), .Fields("autofg"), .Fields("coldt"), _
                                        .Fields("coltm"), .Fields("colid"), .Fields("available"), .Fields("expdt"), _
                                        .Fields("exptm"), .Fields("entdt"), .Fields("enttm"), .Fields("entid"), _
                                        .Fields("centercd"), .Fields("hosfg"), .Fields("donorid"), .Fields("donoraccdt"))
                
                DBConn.Execute SSQL
                iTmx.Checked = False
            End If
        Next iTmx
        SSQL = objExcute.Set_SplitUpdate(.Fields("bldsrc"), .Fields("bldyy"), .Fields("bldno"), _
                                         .Fields("compocd"))
        DBConn.Execute SSQL
    End With
    
    DBConn.CommitTrans
    MsgBox " 혈액 분획처리 되었습니다", vbInformation + vbOKOnly, "혈액분획"
    Call cmdClear_Click
    cmdGenerate.Enabled = False
    Set objExcute = Nothing
    Exit Sub
    
SplitOut_Save_Error:
    DBConn.RollbackTrans
    Set objExcute = Nothing
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub Form_Load()
    Dim objGetSql As New clsGetSqlStatement
    Dim DrRS As New Recordset
    
'    objGetSql.setDbConn DBConn
    Set DrRS = objGetSql.Get_CompoRecordSet
    txtBldNo = ""
    If DrRS.EOF Then
    Else
        objComp.Clear
        objComp.FieldInialize "compocd", "componm"
        With DrRS
            Do Until .EOF
                objComp.AddNew .Fields("compocd").value & "", .Fields("abbrnm").value & ""
                .MoveNext
            Loop
        End With
        Display_Comp
    End If
    Display_Comp
    cmdGenerate.Enabled = False

    Set DrRS = Nothing
    Set objGetSql = Nothing
    
End Sub
Private Sub Display_Comp()
    Dim iTmx As ListItem
    
    With lvwComp.ListItems
        .Clear
        objComp.MoveFirst
        While (Not objComp.EOF)
            Set iTmx = .Add()
            iTmx.Text = objComp.Fields("componm")
            iTmx.SubItems(1) = objComp.Fields("compocd")
            objComp.MoveNext
        Wend
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set ObjDic = Nothing
    Set objComp = Nothing
End Sub

Private Sub lvwComp_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.Checked = True Then
        If lvwComp.ListItems(Item.Index).ForeColor = vbRed Then
            Call medSleep(10)
            DoEvents
            Item.Checked = False
        End If
    End If
        
End Sub

Private Sub txtBldNo_Change()
    If chkBar.value = 1 Then Exit Sub
    Dim lngLen As Long
    
    With txtBldNo
        lngLen = Len(Trim(.Text))
        If lngLen = 2 Then
                .Text = .Text & "-"
                .SelStart = Len(.Text)
        End If
        If lngLen > 2 And lngLen = 5 Then
            .Text = .Text & "-"
            .SelStart = Len(.Text)
        End If
    End With
End Sub

Private Sub txtBldNo_GotFocus()
    txtBldNo.tag = txtBldNo
    txtBldNo.SelStart = 0
    txtBldNo.SelLength = Len(txtBldNo)
End Sub

Private Sub txtBldNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub txtBldNo_KeyPress(KeyAscii As Integer)
    If chkBar.value = 1 Then Exit Sub
    
    If Len(txtBldNo) <> 3 Or Len(txtBldNo) <> 6 Then
        If KeyAscii = vbKeyInsert Then KeyAscii = 0
    End If
        
    If KeyAscii = vbKeyBack Then
        With txtBldNo
            If .Text = "" Then Exit Sub
            If Mid(.Text, Len(.Text)) = "-" Then
                .Text = Mid(.Text, 1, Len(.Text) - 2)
                .SelStart = Len(.Text)
                KeyAscii = 0
            End If
        End With
    End If
End Sub

Private Sub txtBldNo_LostFocus()
'혈액번호별 History 조회
'혈액번호를 입력 받으면 bbs401에서 그에 해당되는 Component를 콤보박스에 보여준다.
    If chkBar.value <> 1 Then
        If Len(txtBldNo) <= 6 Then
            cboComp.Clear
            Exit Sub
        End If
    End If
    
    If txtBldNo.tag = txtBldNo Then Exit Sub
    
    Call Blood_Compo_Detail
    txtBldNo.tag = txtBldNo
End Sub
Private Sub Blood_Compo_Detail()
    Dim objGetSql As New clsGetSqlStatement
    Dim DrRS      As Recordset
    Dim strBldno  As String
    
    Display_Comp                        '혈액제제 보여주기
'    objGetSql.setDbConn DBConn
    cboComp.Clear
    
    If chkBar.value = 1 Then
        strBldno = Mid(txtBldNo, 1, 2) & "-" & _
                   Mid(txtBldNo, 3, 2) & "-" & _
                   Mid(txtBldNo, 5, 6)
    Else
        strBldno = txtBldNo.Text
    End If
    
    
    Set DrRS = objGetSql.Get_Split_TarGet(strBldno, ObjSysInfo.BuildingCd, True)
    If DrRS.EOF Then
        MsgBox "분획대상 혈액이 아닙니다.", vbInformation + vbOKOnly, "혈액분획"
        txtBldNo = ""
        Set objGetSql = Nothing
        Exit Sub
    Else
        Dim itmFound As ListItem
        
        With DrRS
            Do Until .EOF
                cboComp.AddItem .Fields("compocd").value & "" & vbTab & .Fields("abbrnm").value & ""
                Set itmFound = lvwComp.FindItem(.Fields("compocd").value & "", lvwSubItem)
                If Not itmFound Is Nothing Then
                    lvwComp.ListItems.Remove (itmFound.Index)
                End If
                .MoveNext
            Loop
        End With
        cmdGenerate.Enabled = True
        cboComp.ListIndex = 0
        Set itmFound = Nothing
        Set DrRS = Nothing
    End If
    
    Set DrRS = objGetSql.Get_Split_TarGet(strBldno, ObjSysInfo.BuildingCd, False)
    
    If Not DrRS.EOF Then
        Do Until DrRS.EOF
            Set itmFound = lvwComp.FindItem(DrRS.Fields("compocd").value & "", lvwSubItem)
            If Not itmFound Is Nothing Then
                lvwComp.ListItems(itmFound.Index).ForeColor = vbRed
            End If
            DrRS.MoveNext
        Loop
    End If
        
    Set objGetSql = Nothing
End Sub
