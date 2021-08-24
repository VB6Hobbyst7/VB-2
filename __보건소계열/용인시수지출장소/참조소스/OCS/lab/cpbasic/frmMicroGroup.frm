VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMicroGroup 
   Caption         =   "약제그룹코드 Match"
   ClientHeight    =   6090
   ClientLeft      =   5955
   ClientTop       =   2565
   ClientWidth     =   5670
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
   ScaleHeight     =   6090
   ScaleWidth      =   5670
   Begin Threed.SSCommand cmdUpdate 
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   660
      Width           =   1995
      _Version        =   65536
      _ExtentX        =   3519
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "Update"
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   555
      Left            =   120
      TabIndex        =   1
      Top             =   60
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
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0E0FF&
         Height          =   330
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   120
         Width           =   915
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0E0FF&
         Height          =   330
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   120
         Width           =   3255
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H80000007&
         Height          =   330
         Left            =   4260
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   120
         Width           =   1035
      End
   End
   Begin MSComctlLib.TreeView tvOrgList 
      Height          =   4875
      Left            =   90
      TabIndex        =   0
      Top             =   1080
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   8599
      _Version        =   393217
      LineStyle       =   1
      Style           =   7
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.Menu mnuQuit 
      Caption         =   "Quit"
   End
End
Attribute VB_Name = "frmMicroGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdUpdate_Click()
    
    strSql = ""
    strSql = strSql & " UPDATE TWEXAM_ORGLIST"
    strSql = strSql & " SET    Org_AntiGr =  '" & Trim(Text3.Text) & "'"
    strSql = strSql & " WHERE  Org_Code   =  '" & Trim(Text1.Text) & "'"
    Call adoExec(strSql)
    Call MsgBox("세균 그룹코드가 " & Trim(Text3.Text) & " 로 수정되었습니다", vbInformation)
    
    Select Case gSOrgCall
        Case "TEXT"
            frmMicroOrg.cmbGrp.Text = Trim(Text3.Text)
        Case "SPREAD"
            For I = 1 To frmMicroOrg.ssOrgList.DataRowCnt
                frmMicroOrg.ssOrgList.Row = I
                frmMicroOrg.ssOrgList.Col = 1
                If Trim(frmMicroOrg.ssOrgList.Text) = Trim(Text1.Text) Then
                    frmMicroOrg.ssOrgList.Col = 4
                    frmMicroOrg.ssOrgList.TypeButtonText = Trim(Text3.Text)
                    Exit For
                End If
            Next
    End Select

    
End Sub

Private Sub Form_Load()
    Dim sText       As String
    Dim sRowid      As String
    Dim NodeX       As Node
    
    
    Select Case gSOrgCall
        Case "TEXT"
            Text1.Text = Trim(frmMicroOrg.txtOrgCode.Text)
            Text2.Text = Trim(frmMicroOrg.txtOrgNm.Text)
            Text3.Text = Trim(frmMicroOrg.cmbGrp.Text)
        Case "SPREAD"
            frmMicroOrg.ssOrgList.Row = frmMicroOrg.ssOrgList.ActiveRow
            frmMicroOrg.ssOrgList.Col = 1: Me.Text1.Text = Trim(frmMicroOrg.ssOrgList.Text)
            frmMicroOrg.ssOrgList.Col = 2: Me.Text2.Text = Trim(frmMicroOrg.ssOrgList.Text)
            frmMicroOrg.ssOrgList.Col = 4: Me.Text3.Text = Trim(frmMicroOrg.ssOrgList.TypeButtonText)
    End Select
    
    tvOrgList.Nodes.Clear
    
    strSql = ""
    strSql = strSql & " SELECT Grp_Code, MAX(RowID) RWID"
    strSql = strSql & " FROM   TW_MIS_EXAM.TWEXAM_ANTIGROUP"
    strSql = strSql & " GROUP  BY Grp_Code"
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    
    Do Until adoSet.EOF
        sRowid = adoSet.Fields("RwID").Value & ""
        sText = Trim(adoSet.Fields("Grp_Code").Value & "")
        Set NodeX = tvOrgList.Nodes.Add(, , "A1" & sRowid, sText)
        GoSub Load_SubCode
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    If Trim(gSText) <> "" Then
        GoSub Text_Check
    End If
    
    Exit Sub
    
Load_SubCode:
    Dim adoSubCode1     As ADODB.Recordset
    Dim sSubText1       As String
    Dim sSubText2       As String
    
    strSql = ""
    strSql = strSql & " SELECT a.Anti_Code, a.RowID, b.Codenm"
    strSql = strSql & " FROM   TW_MIS_EXAM.TWEXAM_AntiGroup a,"
    strSql = strSql & "        TWEXAM_ANTILIST  b "
    strSql = strSql & " WHERE  a.Grp_Code  = '" & tvOrgList.Nodes("A1" & sRowid).Text & "'"
    strSql = strSql & " AND    a.Anti_Code IS NOT NULL"
    strSql = strSql & " AND    a.Anti_Code = b.Codeky(+)"
    
    If False = adoSetOpen(strSql, adoSubCode1) Then Return
    
    Do Until adoSubCode1.EOF
        sSubText1 = "B2" & adoSubCode1.Fields("RowID")
        sSubText2 = Trim(adoSubCode1.Fields("Anti_Code").Value) & "." & _
                    StrConv(Trim(adoSubCode1.Fields("Codenm").Value & ""), vbProperCase)
        Set NodeX = tvOrgList.Nodes.Add("A1" & sRowid, tvwChild, sSubText1, sSubText2)
        adoSubCode1.MoveNext
    Loop
    Call adoSetClose(adoSubCode1)
    
    Return

Text_Check:
    For I = 1 To tvOrgList.Nodes.Count
        If Left(tvOrgList.Nodes.Item(I).Key, 2) = "A1" Then
            tvOrgList.Nodes(I).Selected = True
            If Trim(tvOrgList.SelectedItem.Text) = Trim(gSText) Then
                tvOrgList.Nodes(I).Expanded = True
                Exit For
            End If
        End If
    Next
    Return
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    gSText = ""
End Sub

Private Sub mnuQuit_Click()
    Unload Me
    
End Sub

Private Sub mnuSet_Click()
    
    frmMicroOrg.cmbGrp.Text = Trim(Text3.Text)
    Unload Me
    
End Sub

Private Sub tvOrgList_NodeClick(ByVal Node As MSComctlLib.Node)
    
    If Left(Node.Key, 2) = "B2" Then
        Text3.Text = tvOrgList.SelectedItem.Parent.Text
    Else
        Text3.Text = tvOrgList.SelectedItem.Text
    End If

End Sub
