VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmRemark 
   Caption         =   "Remark 관리"
   ClientHeight    =   6180
   ClientLeft      =   555
   ClientTop       =   1665
   ClientWidth     =   9885
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
   ScaleHeight     =   6180
   ScaleWidth      =   9885
   Begin VB.CheckBox chkRef 
      Caption         =   "Auto Refrash"
      Height          =   255
      Left            =   8040
      TabIndex        =   12
      Top             =   840
      Value           =   1  '확인
      Width           =   1395
   End
   Begin Threed.SSPanel SSPanel3 
      Height          =   315
      Left            =   3900
      TabIndex        =   4
      Top             =   840
      Width           =   4095
      _Version        =   65536
      _ExtentX        =   7223
      _ExtentY        =   556
      _StockProps     =   15
      Caption         =   "Remark 내용"
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
      BorderWidth     =   1
      BevelOuter      =   1
      BevelInner      =   2
      Alignment       =   1
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   315
      Left            =   300
      TabIndex        =   3
      Top             =   840
      Width           =   3555
      _Version        =   65536
      _ExtentX        =   6271
      _ExtentY        =   556
      _StockProps     =   15
      Caption         =   "Remark Code"
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
      BorderWidth     =   1
      BevelOuter      =   1
      BevelInner      =   2
      Alignment       =   1
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   615
      Left            =   300
      TabIndex        =   1
      Top             =   120
      Width           =   7695
      _Version        =   65536
      _ExtentX        =   13573
      _ExtentY        =   1085
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
      BorderWidth     =   1
      BevelInner      =   1
      Begin VB.TextBox txtAbbCode 
         Appearance      =   0  '평면
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  '없음
         Height          =   270
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   120
         Width           =   1995
      End
      Begin VB.TextBox txtExName 
         Appearance      =   0  '평면
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  '없음
         Height          =   270
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   120
         Width           =   2055
      End
      Begin VB.TextBox txtExgubun 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  '없음
         Height          =   270
         Left            =   1680
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000007&
         Caption         =   "Label2"
         Height          =   255
         Left            =   4620
         TabIndex        =   10
         Top             =   180
         Width           =   1995
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000007&
         Caption         =   "Label2"
         Height          =   255
         Left            =   2580
         TabIndex        =   9
         Top             =   180
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000007&
         Caption         =   "Label2"
         Height          =   255
         Left            =   1740
         TabIndex        =   8
         Top             =   180
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Slip구분 :"
         Height          =   195
         Left            =   420
         TabIndex        =   2
         Top             =   180
         Width           =   1035
      End
   End
   Begin VB.TextBox txtAbbName 
      BackColor       =   &H00C0C0C0&
      Height          =   3135
      Left            =   3900
      Locked          =   -1  'True
      MaxLength       =   400
      MultiLine       =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1200
      Width           =   5535
   End
   Begin MSComctlLib.TreeView tvRmk 
      Height          =   4815
      Left            =   270
      TabIndex        =   13
      Top             =   1200
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   8493
      _Version        =   393217
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin MSForms.CommandButton cmdRefrash 
      Height          =   615
      Left            =   8040
      TabIndex        =   11
      Top             =   120
      Width           =   1395
      Caption         =   "Refrash"
      PicturePosition =   327683
      Size            =   "2461;1085"
      Picture         =   "frmRemark.frx":0000
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
      Caption         =   "Job"
      Visible         =   0   'False
      Begin VB.Menu mnuNew 
         Caption         =   "Data편집"
      End
      Begin VB.Menu mnuDel 
         Caption         =   "Data삭제"
      End
   End
End
Attribute VB_Name = "frmRemark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim NodeX           As Node

Private Sub chkRef_Click()
    
    If chkRef.Value = "1" Then
        chkRef.Caption = "Auto Refrash"
    Else
        chkRef.Caption = "Manu Refrash"
    End If
    
End Sub

Public Sub cmdRefrash_Click()
    Dim sRowid      As String
    Dim sText       As String
    
    
    Screen.MousePointer = vbHourglass
    
    tvRmk.Nodes.Clear
    StrSql = ""
    StrSql = StrSql & " SELECT a.*, a.RowID"
    StrSql = StrSql & " FROM   TWEXAM_SPECODE a"
    StrSql = StrSql & " WHERE  a.Codegu = '12'"
    
    If False = adoSetOpen(StrSql, adoSet) Then Screen.MousePointer = vbDefault: Exit Sub
    
    Do Until adoSet.EOF
        sRowid = adoSet.Fields("RowID").Value & ""
        sText = Trim(adoSet.Fields("Codeky").Value & "") & ". " & _
                Trim(adoSet.Fields("Codenm").Value & "")
        Set NodeX = tvRmk.Nodes.Add(, , "A1" & sRowid, sText)
        GoSub Load_SubCode
        adoSet.MoveNext
    Loop
    
    txtAbbName.Text = ""
    tvRmk.SetFocus
    
    For I = 1 To tvRmk.Nodes.Count
        If tvRmk.Nodes(I).Text = Trim(Me.txtExgubun.Text) & ". " & Trim(txtExName.Text) Then
            tvRmk.Nodes(I).Expanded = True
            Exit For
        End If
    Next
    
    
    If adoSet.RecordCount > 0 Then
        tvRmk.Nodes(1).Selected = True
    End If
    
    Screen.MousePointer = vbDefault
    Call adoSetClose(adoSet)
    Exit Sub
    
Load_SubCode:
    Dim adoSubCode1     As ADODB.Recordset
    Dim sSubText1       As String
    Dim sSubText2       As String
    
    
    StrSql = ""
    StrSql = StrSql & " SELECT a.*, a.RowID"
    StrSql = StrSql & " FROM   TWEXAM_Remark a"
    StrSql = StrSql & " WHERE  a.ExGubun = '" & Left(tvRmk.Nodes("A1" & sRowid).Text, 2) & "'"
    If False = adoSetOpen(StrSql, adoSubCode1) Then Return
    
    Do Until adoSubCode1.EOF
        sSubText1 = "B2" & adoSubCode1.Fields("RowID")
        sSubText2 = adoSubCode1.Fields("AbbCode").Value
        Set NodeX = tvRmk.Nodes.Add("A1" & sRowid, tvwChild, sSubText1, sSubText2)
        adoSubCode1.MoveNext
    Loop
    Call adoSetClose(adoSubCode1)
    
    Return
    
End Sub


Private Sub Form_Load()

    Dim sRowid      As String
    Dim sText       As String
    
    Me.Top = 1
    Me.Left = 1
    Me.Height = 7000
    Me.Width = 10000
    
   
    tvRmk.Nodes.Clear
    
    StrSql = ""
    StrSql = StrSql & " SELECT a.*, a.RowID"
    StrSql = StrSql & " FROM   TWEXAM_SPECODE a"
    StrSql = StrSql & " WHERE  a.Codegu = '12'"
    
    If False = adoSetOpen(StrSql, adoSet) Then Exit Sub
    
    Do Until adoSet.EOF
        sRowid = adoSet.Fields("RowID").Value & ""
        sText = Trim(adoSet.Fields("Codeky").Value & "") & ". " & _
                Trim(adoSet.Fields("Codenm").Value & "")
        Set NodeX = tvRmk.Nodes.Add(, , "A1" & sRowid, sText)
        GoSub Load_SubCode
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    Exit Sub
    
Load_SubCode:
    Dim adoSubCode1     As ADODB.Recordset
    Dim sSubText1       As String
    Dim sSubText2       As String
    
    
    StrSql = ""
    StrSql = StrSql & " SELECT a.*, a.RowID"
    StrSql = StrSql & " FROM   TWEXAM_Remark a"
    StrSql = StrSql & " WHERE  a.ExGubun = '" & Left(tvRmk.Nodes("A1" & sRowid).Text, 2) & "'"
    If False = adoSetOpen(StrSql, adoSubCode1) Then Return
    
    Do Until adoSubCode1.EOF
        sSubText1 = "B2" & adoSubCode1.Fields("RowID")
        sSubText2 = adoSubCode1.Fields("AbbCode").Value
        Set NodeX = tvRmk.Nodes.Add("A1" & sRowid, tvwChild, sSubText1, sSubText2)
        adoSubCode1.MoveNext
    Loop
    Call adoSetClose(adoSubCode1)
    
    Return
    
End Sub

Private Sub mnuDel_Click()
    Dim sCodeky     As String
    Dim sCodeGu     As String
    Dim sDataCk     As Integer
    
    sCodeky = Trim(tvRmk.SelectedItem.Text)
    sCodeGu = Left(tvRmk.SelectedItem.Parent.Text, 2)
    
    GoSub Data_Check_Routine
    If True = sDataCk Then
        GoSub Delete_Routine
    Else
        MsgBox "아마도 이미 지워진 Data 일수도 있습니다!." & vbCrLf & " Refrash Botton 을 눌러보세요!"
        Exit Sub
    End If
    
    Exit Sub
    
    
Data_Check_Routine:
    StrSql = ""
    StrSql = StrSql & " SELECT *"
    StrSql = StrSql & " FROM   TWEXAM_Remark"
    StrSql = StrSql & " WHERE  ExGubun  =  '" & Trim(sCodeGu) & "'"
    StrSql = StrSql & " AND    AbbCode  =  '" & Trim(sCodeky) & "'"
    If False = adoSetOpen(StrSql, adoSet) Then
        sDataCk = False
        Return
    Else
        sDataCk = True
    End If
    Call adoSetClose(adoSet)
    
    Return
    
Delete_Routine:
    sCodeky = Trim(tvRmk.SelectedItem.Text)
    sCodeGu = Left(tvRmk.SelectedItem.Parent.Text, 2)
    
    sMsg = sCodeky & " 를 삭제하시겠습니까?"
    If vbNo = MsgBox(sMsg, vbQuestion + vbYesNo, "삭제 확인 Box") Then
        Return
    End If
    
    StrSql = ""
    StrSql = StrSql & " DELETE"
    StrSql = StrSql & " FROM   TWEXAM_Remark"
    StrSql = StrSql & " WHERE  ExGubun  =  '" & Trim(sCodeGu) & "'"
    StrSql = StrSql & " AND    AbbCode  =  '" & Trim(sCodeky) & "'"
    
    If adoExec(StrSql) Then
        MsgBox "삭제하였습니다!.", vbOKOnly + vbInformation, "삭제 Complete Information"
        If chkRef.Value = "1" Then
            Call cmdRefrash_Click
        End If
    Else
        MsgBox "어떤 오류로 인하여 삭제하지 못하였습니다!", vbOKOnly + vbInformation, "삭제Miss Information"
        Return
    End If
    
    Return
        
End Sub

Private Sub mnuNew_Click()
    
    frmRmk.Show vbModal
    If frmRemark.chkRef.Value = "1" Then
        Call cmdRefrash_Click
    End If
    
End Sub

Private Sub mnuQuit_Click()
    Unload Me
    
End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub tvRmk_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sDataCk     As Integer
    
    
    If KeyCode = vbKeyDelete Then
        If Left(tvRmk.SelectedItem.Key, 2) = "B2" Then
            GoSub Delte_Routine
        End If
    End If
    tvRmk.SetFocus
    Exit Sub
    
    
    
Delte_Routine:
    Dim sCodeky     As String
    Dim sCodeGu     As String
    
    sCodeGu = Left(tvRmk.SelectedItem.Parent.Text, 2)
    sCodeky = tvRmk.SelectedItem.Text
    
    StrSql = ""
    StrSql = StrSql & " SELECT *"
    StrSql = StrSql & " FROM   TWEXAM_Remark"
    StrSql = StrSql & " WHERE  ExGubun  =  '" & Trim(sCodeGu) & "'"
    StrSql = StrSql & " AND    AbbCode  =  '" & Trim(sCodeky) & "'"
    If False = adoSetOpen(StrSql, adoSet) Then
        sDataCk = False
    Else
        sDataCk = True
    End If
    Call adoSetClose(adoSet)
    
    If sDataCk = False Then
        MsgBox "아마도 이미 지워진 Data 일수도 있습니다!." & vbCrLf & " Refrash Botton 을 눌러보세요!"
        Return
    End If
    
    sMsg = sCodeGu & " 의 " & vbCrLf & "Code = " & sCodeky & " 를 삭제하시겠습니까?"
    If vbNo = MsgBox(sMsg, vbQuestion + vbYesNo, "삭제확인 Box") Then
        tvRmk.SetFocus
        Return
    End If
    
    StrSql = ""
    StrSql = StrSql & " DELETE"
    StrSql = StrSql & " FROM   TWEXAM_Remark"
    StrSql = StrSql & " WHERE  Exgubun  =  '" & Trim(sCodeGu) & "'"
    StrSql = StrSql & " AND    AbbCode  =  '" & Trim(sCodeky) & "'"
    If adoExec(StrSql) Then
        sMsg = "삭제하였습니다!.." & vbCrLf & "Data Tree 를 Refrash 시키시겠습니까?"
        If chkRef.Value = "1" Then
            Call cmdRefrash_Click
        End If
    Else
        MsgBox "어떤 오류로 인하여 삭제하지 못하였습니다!", vbOKOnly + vbInformation, "삭제Miss Information"
        Return
    End If
    
    Return
    
    
End Sub

Private Sub tvRmk_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button <> 2 Then Exit Sub
    If Left(tvRmk.SelectedItem.Key, 2) = "A1" Then
        mnuDel.Visible = False
        mnuNew.Caption = "신규입력"
    Else
        mnuNew.Caption = "Data수정"
        mnuDel.Visible = True
    End If
    
    PopupMenu mnuJob
    
    
End Sub

Private Sub tvRmk_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim sCodeGu     As String
    Dim sCodeky     As String
    
    
    If Left(Node.Key, 2) = "B2" Then
        sCodeky = tvRmk.SelectedItem.Text
        sCodeGu = tvRmk.SelectedItem.Parent.Text
        GoSub Get_Remark_Text
        Me.txtExgubun = Left(sCodeGu, 2)
        Me.txtExName = Trim(Mid(sCodeGu, 4, Len(sCodeGu) - 3))
        Me.txtAbbCode = Trim(sCodeky)
    Else
        sCodeky = tvRmk.SelectedItem.Text
        Me.txtExgubun = Left(sCodeky, 2)
        Me.txtExName = Trim(Mid(sCodeky, 4, Len(sCodeky) - 3))
        txtAbbCode.Text = ""
        txtAbbName.Text = ""
    End If
    Exit Sub
    
Get_Remark_Text:
    txtAbbName.Text = ""
    StrSql = ""
    StrSql = StrSql & " SELECT *"
    StrSql = StrSql & " FROM   TWEXAM_Remark"
    StrSql = StrSql & " WHERE  ExGubun  =  '" & Left(sCodeGu, 2) & "'"
    StrSql = StrSql & " AND    AbbCode  =  '" & Trim(sCodeky) & "'"
    If False = adoSetOpen(StrSql, adoSet) Then
        MsgBox "지워진 Data 인지도 모르겠습니다!.. Refrash Botton 을 눌러보세요!"
        Return
    End If
    
    txtAbbName.Text = adoSet.Fields("abbName").Value & ""
    Call adoSetClose(adoSet)
    
    Return
End Sub

