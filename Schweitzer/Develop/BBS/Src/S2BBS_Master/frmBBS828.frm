VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frmBBS828 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "XM결과 필수입력 제제"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   10845
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "저장(&S)"
      Height          =   495
      Left            =   4245
      Style           =   1  '그래픽
      TabIndex        =   1
      Top             =   5925
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5535
      Style           =   1  '그래픽
      TabIndex        =   0
      Tag             =   "128"
      Top             =   5925
      Width           =   1215
   End
   Begin MedControls1.LisLabel lblTitle 
      Height          =   375
      Left            =   1822
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2190
      Width           =   7200
      _ExtentX        =   12700
      _ExtentY        =   661
      BackColor       =   8421504
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "혈액제제 리스트"
      Appearance      =   0
   End
   Begin MSComctlLib.ListView lvwCompo 
      Height          =   3135
      Left            =   1800
      TabIndex        =   3
      Top             =   2580
      Width           =   7200
      _ExtentX        =   12700
      _ExtentY        =   5530
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
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "제제코드"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "제제명"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Sel"
         Object.Width           =   0
      EndProperty
   End
End
Attribute VB_Name = "frmBBS828"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private Const BC2_COMPO$ = "B004"
'Private Const BC2_XM_COMPO$ = "B032"

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim strSql1 As String
    Dim strSql2 As String
    Dim strSelCompo As String
    Dim itmX As ListItem
    
    If lvwCompo.ListItems.Count = 0 Then
        MsgBox "제제마스터를 먼저 등록하십시오.", vbExclamation
        Exit Sub
    End If
    
    For Each itmX In lvwCompo.ListItems
        If itmX.Checked Then
            strSelCompo = strSelCompo & itmX.Text & ";"
        End If
    Next
    
    If strSelCompo = "" Then
        MsgBox "XM결과 필수입력 제제를 선택하십시오.", vbExclamation
        Exit Sub
    End If
    
    strSql1 = " delete from " & T_COM003 & _
              " where " & DBW("cdindex=", BC2_XM_COMPO) & _
              " and " & DBW("cdval1=", BC2_XM_COMPO)
    
    strSql2 = " insert into " & T_COM003 & " (cdindex,cdval1,text1) " & _
              " values (" & DBV("cdindex", BC2_XM_COMPO, 1) & DBV("cdval1", BC2_XM_COMPO, 1) & DBV("text1", strSelCompo) & ")"

    On Error GoTo ErrTrap
    
    DBConn.BeginTrans
    
    DBConn.Execute strSql1
    DBConn.Execute strSql2
    
    DBConn.CommitTrans
    
    MsgBox "정상적으로 처리되었습니다.", vbInformation
    
    Call LoadSelectedCompo
    
    Exit Sub
    
ErrTrap:
    DBConn.RollbackTrans
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    
    DoEvents
    
    Call LoadCompo
    Call LoadSelectedCompo
    Screen.MousePointer = vbDefault
End Sub

Private Sub LoadCompo()
    Dim Rs As Recordset
    Dim itmX As ListItem
    Dim strSql As String
    
'    strSql = " select cdval1,field1 from " & T_COM003 & _
'             " where " & DBW("cdindex=", BC2_COMPO)
    strSql = " select compocd, componm from " & T_BBS006 & _
             " where (expdt is null or expdt='') " & _
             " order by compocd "
           
    Set Rs = New Recordset
    
    Rs.Open strSql, DBConn
    
    Dim i As Long
    
    lvwCompo.ListItems.Clear
    Do Until Rs.EOF
        Set itmX = lvwCompo.ListItems.Add()
'        itmX.Text = Rs.Fields("cdval1").Value & ""
'        itmX.SubItems(1) = Rs.Fields("field1").Value & ""
        itmX.Text = Rs.Fields("compocd").Value & ""
        itmX.SubItems(1) = Rs.Fields("componm").Value & ""
        
        Rs.MoveNext
    Loop
    
    If lvwCompo.ListItems.Count = 0 Then
        MsgBox "제제마스터를 먼저 등록하십시오.", vbExclamation
    End If
    
    Set Rs = Nothing
End Sub

Private Sub LoadSelectedCompo()
    Dim itmX As ListItem
    Dim Rs As Recordset
    Dim strSql As String
    Dim aryCompo() As String
    Dim i As Long
    
    For Each itmX In lvwCompo.ListItems
        itmX.SubItems(2) = ""
        itmX.ForeColor = vbGrayText
        itmX.ListSubItems(1).ForeColor = vbGrayText
    Next
    
    strSql = " select text1 from " & T_COM003 & _
             " where " & DBW("cdindex=", BC2_XM_COMPO) & _
             " and " & DBW("cdval1=", BC2_XM_COMPO)
    
    Set Rs = New Recordset
    Rs.Open strSql, DBConn
    
    If Rs.EOF = False Then
        aryCompo = Split(Rs.Fields("text1").Value & "", ";")
        
        For i = LBound(aryCompo) To UBound(aryCompo)
            For Each itmX In lvwCompo.ListItems
                If itmX.Text = aryCompo(i) Then
                    itmX.Checked = True
                    
                    itmX.SubItems(2) = "1"
                    itmX.ForeColor = vbBlack
                    itmX.ListSubItems(1).ForeColor = vbBlack
                End If
            Next
        Next
    End If
    
    lvwCompo.Refresh
    
    Set Rs = Nothing
End Sub

Private Sub lvwCompo_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.SubItems(2) = "1" Then '등록된 아이템인 경우
        Item.ForeColor = IIf(Item.ForeColor = vbBlack, vbRed, vbBlack)
        Item.ListSubItems(1).ForeColor = IIf(Item.ListSubItems(1).ForeColor = vbBlack, vbRed, vbBlack)
    Else
        Item.ForeColor = IIf(Item.ForeColor = vbGrayText, vbRed, vbGrayText)
        Item.ListSubItems(1).ForeColor = IIf(Item.ListSubItems(1).ForeColor = vbGrayText, vbRed, vbGrayText)
    End If
End Sub

Private Sub lvwCompo_ItemClick(ByVal Item As MSComctlLib.ListItem)
'등록된 아이템(저장된)에 대해서 변경할 경우 색깔 변경
    Item.Checked = Not Item.Checked
    
    If Item.SubItems(2) = "1" Then '등록된 아이템인 경우
        Item.ForeColor = IIf(Item.ForeColor = vbBlack, vbRed, vbBlack)
        Item.ListSubItems(1).ForeColor = IIf(Item.ListSubItems(1).ForeColor = vbBlack, vbRed, vbBlack)
    Else
        Item.ForeColor = IIf(Item.ForeColor = vbGrayText, vbRed, vbGrayText)
        Item.ListSubItems(1).ForeColor = IIf(Item.ListSubItems(1).ForeColor = vbGrayText, vbRed, vbGrayText)
    End If
End Sub
