VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form Anato_User 
   Caption         =   "사용자관리"
   ClientHeight    =   7020
   ClientLeft      =   375
   ClientTop       =   1290
   ClientWidth     =   11295
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7020
   ScaleWidth      =   11295
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00C0C0C0&
      Height          =   6135
      Left            =   3825
      ScaleHeight     =   6075
      ScaleWidth      =   5265
      TabIndex        =   15
      Top             =   405
      Width           =   5325
      Begin VB.TextBox txtDept 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1890
         MaxLength       =   4
         TabIndex        =   4
         Top             =   4155
         Width           =   2850
      End
      Begin VB.TextBox txtPassWord 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         IMEMode         =   3  '사용 못함
         Left            =   1890
         MaxLength       =   4
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   2940
         Width           =   2850
      End
      Begin VB.TextBox txtRange 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1890
         MaxLength       =   10
         TabIndex        =   3
         Top             =   3525
         Width           =   2850
      End
      Begin VB.TextBox txtUserId 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1890
         MaxLength       =   6
         TabIndex        =   0
         Top             =   1320
         Width           =   1100
      End
      Begin VB.TextBox txtName 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1890
         MaxLength       =   20
         TabIndex        =   1
         Top             =   2355
         Width           =   2850
      End
      Begin VB.Label Label8 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00808000&
         BorderStyle     =   1  '단일 고정
         Caption         =   "사 용 자 정 보 등 록"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   5280
      End
      Begin VB.Label Label7 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "과 코 드"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   885
         TabIndex        =   20
         Top             =   4245
         Width           =   720
      End
      Begin VB.Label Label5 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "비밀번호"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   885
         TabIndex        =   19
         Top             =   3030
         Width           =   720
      End
      Begin VB.Label Label3 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "권    한"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   885
         TabIndex        =   18
         Top             =   3615
         Width           =   720
      End
      Begin VB.Label Label2 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "번    호"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   870
         TabIndex        =   17
         Top             =   1365
         Width           =   720
      End
      Begin VB.Label Label4 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "이    름"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   885
         TabIndex        =   16
         Top             =   2445
         Width           =   720
      End
   End
   Begin VB.ListBox lstUser 
      BackColor       =   &H00EBF5EB&
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5640
      Left            =   630
      TabIndex        =   14
      Top             =   720
      Width           =   3030
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808000&
      Height          =   4515
      Left            =   9315
      ScaleHeight     =   4455
      ScaleWidth      =   1305
      TabIndex        =   13
      Top             =   2025
      Width           =   1365
      Begin Threed.SSCommand cmdDelete 
         Height          =   900
         Left            =   0
         TabIndex        =   7
         Top             =   1800
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   1588
         _StockProps     =   78
         Caption         =   "삭 제"
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "ANATO106.frx":0000
      End
      Begin Threed.SSCommand cmdSave 
         Height          =   900
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   1588
         _StockProps     =   78
         Caption         =   "등 록"
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "ANATO106.frx":031A
      End
      Begin Threed.SSCommand cmdView 
         Height          =   900
         Left            =   0
         TabIndex        =   6
         Top             =   900
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   1588
         _StockProps     =   78
         Caption         =   "조 회"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "ANATO106.frx":076C
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   900
         Left            =   0
         TabIndex        =   8
         Top             =   2700
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   1588
         _StockProps     =   78
         Caption         =   "종 료"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "ANATO106.frx":0BBE
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00C0C0C0&
      Height          =   1500
      Left            =   9315
      ScaleHeight     =   1440
      ScaleWidth      =   1305
      TabIndex        =   9
      Top             =   405
      Width           =   1365
      Begin VB.OptionButton optCode 
         Caption         =   " 코드순"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   90
         TabIndex        =   11
         Top             =   495
         Value           =   -1  'True
         Width           =   1185
      End
      Begin VB.OptionButton optName 
         Caption         =   " 성명순"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   90
         TabIndex        =   10
         Top             =   945
         Width           =   1230
      End
      Begin VB.Label Label1 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00808000&
         BorderStyle     =   1  '단일 고정
         Caption         =   "조회순서"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   1320
      End
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  '평면
      BackColor       =   &H00808000&
      BorderStyle     =   1  '단일 고정
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   630
      TabIndex        =   22
      Top             =   405
      Width           =   3030
   End
End
Attribute VB_Name = "Anato_User"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDelete_Click()
    
    Dim Response

    Response = MsgBox(" 삭제하시겠습니까?", vbYesNo + vbCritical + vbDefaultButton2, "사용자관리")
    If Response = vbYes Then
        GoSub USER_DELETE
    End If
        Call Form_Load
    Exit Sub

USER_DELETE:
    strSQL = ""
    strSQL = strSQL & " DELETE FROM TWBAS_PASS "
    If lstUser.ListIndex <> -1 Then
        strSQL = strSQL & " WHERE IDNUMBER = '" & MidH(lstUser.List(lstUser.ListIndex), 1, 6) & "' "
    Else
        strSQL = strSQL & " WHERE IDNUMBER = '" & txtUserId & "' "
        strSQL = strSQL & "   AND Grade    = '" & txtRange & "' "
        strSQL = strSQL & "   AND Deptcode = '" & txtDept & "' "
    End If
    
    adoConnect.BeginTrans
    
    Result = AdoExecute(strSQL)
    
    If Result = True And Rowindicator > 0 Then
        adoConnect.CommitTrans
        MsgBox "삭제 완료되었습니다.", vbInformation, "진단병리과"
    Else
        adoConnect.RollbackTrans
        MsgBox "작업도중 예기치 못한 오류가 발생했습니다.", vbCritical, "오류"
    End If
    
    txtUserId.Text = ""
    txtName.Text = ""
    txtPassword.Text = ""
    txtRange.Text = ""
    txtDept.Text = ""
    
    Return
    

End Sub


Private Sub cmdExit_Click()
    
    Unload Me
    
End Sub


Private Sub cmdSave_Click()

    Dim rs                  As ADODB.Recordset
    
    If Trim(txtUserId) = "" Then Exit Sub
    
    strSQL = ""
    strSQL = strSQL & " SELECT *"
    strSQL = strSQL & " FROM   TWBAS_PASS"
    strSQL = strSQL & " WHERE  IDNUMBER = '" & txtUserId & "'"
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result Then
        AdoCloseSet rs
        GoSub USER_UPDATE
    Else
        GoSub USER_INSERT
    End If
    
    Call Form_Load
    txtUserId.SetFocus
    Exit Sub
    
USER_INSERT:
    strSQL = ""
    strSQL = strSQL & " INSERT INTO TWBAS_PASS"
    strSQL = strSQL & "       (PROGRAMID, IDNUMBER, Name, Password, GRADE, Deptcode) "
    strSQL = strSQL & " VALUES(' ',"
    strSQL = strSQL & "        '" & txtUserId.Text & "',"
    strSQL = strSQL & "        '" & txtName.Text & "',"
    strSQL = strSQL & "        '" & txtPassword.Text & "',"
    strSQL = strSQL & "        '" & txtRange.Text & "',"
    strSQL = strSQL & "        '" & Trim(txtDept.Text) & "')"
    
    adoConnect.BeginTrans
    
    Result = AdoExecute(strSQL)
    
    If Result = True And Rowindicator > 0 Then
        adoConnect.CommitTrans
        MsgBox "저장 완료되었습니다.", vbInformation, "진단병리과"
    Else
        adoConnect.RollbackTrans
        MsgBox "작업도중 예기치 못한 오류가 발생했습니다.", vbCritical, "오류"
    End If
   
    Return

'----------------------------------------------------------------------------
USER_UPDATE:
    strSQL = ""
    strSQL = strSQL & " UPDATE TWBAS_PASS"
    strSQL = strSQL & " SET    Name      = '" & txtName.Text & "',"
    strSQL = strSQL & "        Password  = '" & txtPassword.Text & "',"
    strSQL = strSQL & "        Grade     = '" & txtRange.Text & "',"
    strSQL = strSQL & "        Deptcode  = '" & Trim(txtDept.Text) & "'"
    strSQL = strSQL & " WHERE  IDNUMBER  = '" & txtUserId.Text & "'"
    
    adoConnect.BeginTrans
    
    Result = AdoExecute(strSQL)
    
    If Result = True And Rowindicator > 0 Then
        adoConnect.CommitTrans
        MsgBox "저장 완료되었습니다.", vbInformation, "진단병리과"
    Else
        adoConnect.RollbackTrans
        MsgBox "작업도중 예기치 못한 오류가 발생했습니다.", vbCritical, "오류"
    End If
   
    Return


End Sub

Private Sub cmdView_Click()
        
    Dim rs                  As ADODB.Recordset
    
    Dim I                   As Integer
    
    txtUserId = ""
    txtName = ""
    txtPassword = ""
    txtRange = ""
    txtDept = ""
    
    lstUser.Clear
    
'---------------------------------------------'
'   사용자 DB READ                             '
'---------------------------------------------'
    
    strSQL = ""
    strSQL = strSQL & " SELECT * "
    strSQL = strSQL & "   FROM TWBAS_PASS"
    strSQL = strSQL & "  WHERE IDNUMBER > '0' "
    strSQL = strSQL & "    AND DeptCode = 'AP' "
    If optCode.Value = True Then
        strSQL = strSQL & " ORDER BY IDNUMBER"
    ElseIf optName.Value = True Then
        strSQL = strSQL & " ORDER BY NAME"
    End If
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then Exit Sub
    
    Do Until rs.EOF
        lstUser.AddItem rs.Fields("IDnumber").Value & "   " & _
                        rs.Fields("name").Value & ""
        rs.MoveNext
    Loop
    AdoCloseSet rs
    
End Sub

Private Sub Form_Load()
        
    Dim rs                  As ADODB.Recordset
    
    Dim I                   As Integer
    
    lstUser.Clear
    txtUserId = ""
    txtName = ""
    txtPassword = ""
    txtRange = ""
    txtDept = ""

'---------------------------------------------'
'   사용자 DB READ                             '
'---------------------------------------------'
    strSQL = ""
    strSQL = strSQL & " SELECT * "
    strSQL = strSQL & "   FROM TWBAS_PASS "
    strSQL = strSQL & "  WHERE IDNUMBER > '0' "
    strSQL = strSQL & "    AND DeptCode = 'AP' "
    strSQL = strSQL & "  ORDER BY IDnumber"
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then Exit Sub
    
    Do Until rs.EOF
        lstUser.AddItem rs.Fields("IDnumber").Value & "   " & rs.Fields("Name").Value & ""
        rs.MoveNext
    Loop
    AdoCloseSet rs
    
End Sub

Private Sub lstUser_Click()
    
    Dim rs                  As ADODB.Recordset
    
    Dim SearchText          As String * 20
     
    SearchText = MidH(lstUser.List(lstUser.ListIndex), 1, 6)
     
    If Trim(SearchText) = "" Then Exit Sub
          
'---------------------------------------------'
'   사용자 DB READ                              '
'---------------------------------------------'
    strSQL = ""
    strSQL = strSQL & " SELECT * "
    strSQL = strSQL & "   FROM TWBAS_PASS "
    strSQL = strSQL & "  WHERE IDnumber = '" & SearchText & "'"
    
    Result = AdoOpenSet(rs, strSQL)
        
    If Result = False Then Exit Sub
  
    txtUserId = rs.Fields("IDnumber").Value & ""
    txtName = rs.Fields("NAME").Value & ""
    txtPassword = rs.Fields("PASSWORD").Value & ""
    txtRange = rs.Fields("GRADE").Value & ""
    txtDept = rs.Fields("DEPTCODE").Value & ""
    
    AdoCloseSet rs

End Sub

Private Sub txtDept_GotFocus()

    txtDept.SelStart = 0
    txtDept.SelLength = Len(txtDept.Text)
    
End Sub

Private Sub txtDept_KeyPress(KeyAscii As Integer)
    
    If KeyAscii >= 97 And KeyAscii <= 122 Then KeyAscii = KeyAscii - 32
    
    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0
    
    SendKeys "{tab}"

End Sub


Private Sub txtJikch_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0
    
    SendKeys "{tab}"

End Sub


Private Sub txtName_GotFocus()
 
    txtName.SelStart = 0
    txtName.SelLength = Len(txtName.Text)

End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0
    
    SendKeys "{tab}"

End Sub


Private Sub txtPassword_GotFocus()

    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword.Text)

End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0
    
    SendKeys "{tab}"

End Sub


Private Sub txtRange_GotFocus()

    txtRange.SelStart = 0
    txtRange.SelLength = Len(txtRange.Text)

End Sub

Private Sub txtRange_KeyPress(KeyAscii As Integer)
    
    If KeyAscii >= 97 And KeyAscii <= 122 Then KeyAscii = KeyAscii - 32
    
    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0
    
    SendKeys "{tab}"

End Sub




Private Sub txtUserId_GotFocus()
 
    txtUserId.SelStart = 0
    txtUserId.SelLength = Len(txtUserId.Text)

End Sub

Private Sub txtUserId_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then
        Exit Sub
    End If
    
    KeyAscii = 0
    
    strSQL = ""
    strSQL = strSQL & " SELECT * "
    strSQL = strSQL & "   FROM TWBAS_PASS "
    strSQL = strSQL & "  WHERE IDnumber = '" & txtUserId & "'"
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then
        SendKeys "{tab}"
        Exit Sub
    End If
    
    txtUserId = rs.Fields("IDnumber").Value & ""
    txtName = rs.Fields("NAME").Value & ""
    txtPassword = rs.Fields("PASSWORD").Value & ""
    txtRange = rs.Fields("GRADE").Value & ""
    txtDept = rs.Fields("DEPTCODE").Value & ""
    
    AdoCloseSet rs
    
    SendKeys "{tab}"

End Sub


