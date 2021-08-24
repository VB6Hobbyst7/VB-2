VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form Anato_Use_Word 
   Caption         =   "상용구관리"
   ClientHeight    =   7050
   ClientLeft      =   765
   ClientTop       =   1905
   ClientWidth     =   11085
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7050
   ScaleWidth      =   11085
   Begin VB.TextBox txtInitial 
      BackColor       =   &H00E1FAFA&
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   480
      MaxLength       =   10
      TabIndex        =   0
      Top             =   765
      Width           =   3015
   End
   Begin VB.TextBox txtcommon 
      BackColor       =   &H00DCFAFA&
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5760
      Left            =   3630
      MaxLength       =   400
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   765
      Width           =   5415
   End
   Begin VB.ListBox lstInitial 
      BackColor       =   &H00EBF5EB&
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4740
      Left            =   480
      TabIndex        =   2
      Top             =   1605
      Width           =   3030
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808000&
      Height          =   6135
      Left            =   9165
      ScaleHeight     =   6075
      ScaleWidth      =   1305
      TabIndex        =   6
      Top             =   480
      Width           =   1365
      Begin Threed.SSCommand cmdDelete 
         Height          =   885
         Left            =   0
         TabIndex        =   4
         Top             =   870
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   1561
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
         Picture         =   "ANATO105.frx":0000
      End
      Begin Threed.SSCommand cmdSave 
         Height          =   885
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   1561
         _StockProps     =   78
         Caption         =   "등 록"
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
         Picture         =   "ANATO105.frx":031A
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   975
         Left            =   0
         TabIndex        =   5
         Top             =   1740
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   1720
         _StockProps     =   78
         Caption         =   "종 료"
         ForeColor       =   0
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
         Picture         =   "ANATO105.frx":076C
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00808000&
      BorderStyle     =   1  '단일 고정
      Caption         =   "약 어 명"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   480
      TabIndex        =   9
      Top             =   480
      Width           =   3030
   End
   Begin VB.Label Label2 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00808000&
      BorderStyle     =   1  '단일 고정
      Caption         =   "상 용 구 절"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   3630
      TabIndex        =   8
      Top             =   480
      Width           =   5430
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00808000&
      BorderStyle     =   1  '단일 고정
      Caption         =   "약어 List"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   480
      TabIndex        =   7
      Top             =   1335
      Width           =   3030
   End
End
Attribute VB_Name = "Anato_Use_Word"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdDelete_Click()
    
    Dim Response

    Response = MsgBox(" 삭제하시겠습니까?", vbYesNo + vbCritical + vbDefaultButton2, "판독예문관리")
    
    If Response = vbYes Then
        strSQL = " DELETE FROM TWEXAM_REMARK WHERE Exgubun = 'AN' AND AbbCode = '" & txtInitial.Text & "'"
        
        adoConnect.BeginTrans
        
        Result = AdoExecute(strSQL)
        
        If Result = True And Rowindicator > 0 Then
            adoConnect.CommitTrans
            MsgBox "삭제 완료되었습니다.", vbInformation, "진단병리과"
        Else
            adoConnect.RollbackTrans
            MsgBox "작업도중 예기치 못한 오류가 발생했습니다.", vbCritical, "오류"
        End If
        
        txtInitial = ""
        txtcommon = ""
        
        Call Form_Load
    End If

End Sub


Private Sub cmdDelete_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0
    
    SendKeys "{tab}"


End Sub


Private Sub cmdExit_Click()

    Unload Me
    
End Sub

Private Sub cmdExit_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0
    
    SendKeys "{tab}"


End Sub


Private Sub cmdSave_Click()
    
    Dim rs                  As ADODB.Recordset
    
    If Trim(txtInitial) = "" Then Exit Sub
    If Trim(txtcommon) = "" Then Exit Sub
    

    strSQL = ""
    strSQL = strSQL & " SELECT abbName "
    strSQL = strSQL & " FROM   TWEXAM_REMARK "
    strSQL = strSQL & " WHERE  exGubun = 'AN' "
    strSQL = strSQL & " AND    abbCode = '" & txtInitial & "' "
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then
        AdoCloseSet rs
        
        strSQL = ""
        strSQL = strSQL & " INSERT INTO TWEXAM_REMARK"
        strSQL = strSQL & "       (Exgubun, abbCode, abbName)"
        strSQL = strSQL & " VALUES ( 'AN',"
        strSQL = strSQL & "          '" & txtInitial.Text & "',"
        strSQL = strSQL & "          '" & txtcommon.Text & "')"
        
        adoConnect.BeginTrans
        
        Result = AdoExecute(strSQL)
        
        If Result = True Then
            If Rowindicator > 0 Then
                adoConnect.CommitTrans
            Else
                adoConnect.RollbackTrans
                MsgBox "등록된 데이타가 없습니다.", vbCritical, "오류"
            End If
        Else
            adoConnect.RollbackTrans
            MsgBox "작업도중 예기치 못한 오류가 발생했습니다.", vbCritical, "오류"
        End If
    Else
        strSQL = ""
        strSQL = strSQL & " UPDATE  TWEXAM_REMARK"
        strSQL = strSQL & " SET     abbName  = '" & txtcommon.Text & "'"
        strSQL = strSQL & " WHERE   Exgubun  = 'AN'"
        strSQL = strSQL & " AND     abbCode  = '" & txtInitial.Text & "'"
        
        adoConnect.BeginTrans
        
        Result = AdoExecute(strSQL)
        
        If Result = True Then
            If Rowindicator > 0 Then
                adoConnect.CommitTrans
            Else
                adoConnect.RollbackTrans
                MsgBox "갱신된 데이타가 없습니다.", vbCritical, "오류"
            End If
        Else
            adoConnect.RollbackTrans
            MsgBox "작업도중 예기치 못한 오류가 발생했습니다.", vbCritical, "오류"
        End If
    
    End If
        
    Call Form_Load
    txtInitial.SetFocus
    Exit Sub
    

End Sub


Private Sub cmdView_Click()
    Dim rs                  As ADODB.Recordset
    
    Dim i                   As Integer
    
    txtInitial = ""
    lstInitial.Clear
    txtcommon = ""
    
'---------------------------------------------'
'   약어 DB READ                             '
'---------------------------------------------'
    strSQL = " SELECT * FROM TWEXAM_Remark ORDER BY abbCode"
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then Exit Sub
    
    Do Until rs.EOF
        lstInitial.AddItem rs.Fields("abbCode").Value & ""
        rs.MoveNext
    Loop
    AdoCloseSet rs
    
End Sub

Private Sub cmdSave_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0
    
    SendKeys "{tab}"


End Sub

Private Sub Form_Load()
    
    Dim rs                  As ADODB.Recordset
    
    Dim i                   As Integer
    
    txtInitial = ""
    lstInitial.Clear
    txtcommon = ""
    

'---------------------------------------------'
'   약어 DB READ                             '
'---------------------------------------------'
    strSQL = ""
    strSQL = strSQL & " SELECT *"
    strSQL = strSQL & " FROM   TWEXAM_REMARK   "
    strSQL = strSQL & " WHERE  Exgubun = 'AN' "
    strSQL = strSQL & " ORDER  BY abbCode"
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then Exit Sub
        
    Do Until rs.EOF
        lstInitial.AddItem rs.Fields("abbCode").Value & ""
        rs.MoveNext
    Loop
    AdoCloseSet rs

End Sub



Private Sub lstInitial_Click()
    
    Dim rs                  As ADODB.Recordset
    
    Dim SearchText          As String * 20
     
    SearchText = Trim(lstInitial.List(lstInitial.ListIndex))
     
    If Trim(SearchText) = "" Then Exit Sub
     
    txtInitial = SearchText

'---------------------------------------------'
'   약어 DB READ                              '
'---------------------------------------------'
    strSQL = ""
    strSQL = strSQL & " SELECT *"
    strSQL = strSQL & " FROM   TWEXAM_Remark"
    strSQL = strSQL & " WHERE  Exgubun = 'AN'"
    strSQL = strSQL & " AND    abbCode = '" & txtInitial.Text & "'"
    
    txtcommon = ""
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then Exit Sub
    txtcommon.Text = rs.Fields("abbName").Value & ""
      
    AdoCloseSet rs
    
End Sub

Private Sub lstInitial_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0
    
    SendKeys "{tab}"


End Sub


Private Sub txtcommon_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0
    
    SendKeys "{tab}"


End Sub


Private Sub txtInitial_GotFocus()
 
    txtInitial.SelStart = 0
    txtInitial.SelLength = Len(txtInitial.Text)


End Sub

Private Sub txtInitial_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0
    
    SendKeys "{tab}"


End Sub


Private Sub txtInitial_LostFocus()
    
    Dim rs                  As ADODB.Recordset
    
    If txtInitial = "" Then Exit Sub
    
'---------------------------------------------'
'   약어 DB READ                              '
'---------------------------------------------'
    strSQL = ""
    strSQL = strSQL & " SELECT *"
    strSQL = strSQL & " FROM   TWEXAM_REMARK"
    strSQL = strSQL & " WHERE  Exgubun = 'AN'"
    strSQL = strSQL & " AND    abbCode = '" & txtInitial.Text & "'"
    
    Result = AdoOpenSet(rs, strSQL)
        
    txtcommon = ""
    
    If Result = False Then Exit Sub
    
    txtcommon.Text = rs.Fields("abbName").Value & ""
    
    AdoCloseSet rs

End Sub


