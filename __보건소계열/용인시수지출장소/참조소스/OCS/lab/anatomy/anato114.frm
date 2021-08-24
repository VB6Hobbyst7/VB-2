VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form Anato_Dyeing_Code 
   Caption         =   "특수염색코드관리"
   ClientHeight    =   5625
   ClientLeft      =   1200
   ClientTop       =   1590
   ClientWidth     =   8715
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5625
   ScaleWidth      =   8715
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00C0C0C0&
      Height          =   3375
      Left            =   3240
      ScaleHeight     =   3315
      ScaleWidth      =   5265
      TabIndex        =   13
      Top             =   1176
      Width           =   5325
      Begin VB.TextBox txtAbbrev 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1890
         MaxLength       =   20
         TabIndex        =   2
         Top             =   2250
         Width           =   2850
      End
      Begin VB.TextBox txtSpeCode 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1890
         MaxLength       =   8
         TabIndex        =   0
         Top             =   810
         Width           =   1100
      End
      Begin VB.TextBox txtCodeNm 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1890
         MaxLength       =   50
         TabIndex        =   1
         Top             =   1725
         Width           =   2850
      End
      Begin VB.Label Label8 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00808000&
         BorderStyle     =   1  '단일 고정
         Caption         =   "특 수 염 색 코 드 정 보 등 록"
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
         TabIndex        =   17
         Top             =   0
         Width           =   5280
      End
      Begin VB.Label Label5 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "약    어"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   765
         TabIndex        =   16
         Top             =   2370
         Width           =   840
      End
      Begin VB.Label Label2 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "코    드"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   720
         TabIndex        =   15
         Top             =   855
         Width           =   840
      End
      Begin VB.Label Label4 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "코 드 명"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   765
         TabIndex        =   14
         Top             =   1785
         Width           =   840
      End
   End
   Begin VB.ListBox lstSpeCode 
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
      Height          =   4740
      ItemData        =   "ANATO114.frx":0000
      Left            =   120
      List            =   "ANATO114.frx":0002
      TabIndex        =   12
      Top             =   525
      Width           =   3030
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808000&
      Height          =   852
      Left            =   3255
      ScaleHeight     =   795
      ScaleWidth      =   5265
      TabIndex        =   11
      Top             =   4560
      Width           =   5325
      Begin Threed.SSCommand cmdDelete 
         Height          =   828
         Left            =   2640
         TabIndex        =   5
         Top             =   -24
         Width           =   1332
         _Version        =   65536
         _ExtentX        =   2350
         _ExtentY        =   1460
         _StockProps     =   78
         Caption         =   "삭 제"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "ANATO114.frx":0004
      End
      Begin Threed.SSCommand cmdSave 
         Height          =   828
         Left            =   0
         TabIndex        =   3
         Top             =   -24
         Width           =   1332
         _Version        =   65536
         _ExtentX        =   2350
         _ExtentY        =   1460
         _StockProps     =   78
         Caption         =   "등 록"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "ANATO114.frx":031E
      End
      Begin Threed.SSCommand cmdView 
         Height          =   828
         Left            =   1320
         TabIndex        =   4
         Top             =   -24
         Width           =   1332
         _Version        =   65536
         _ExtentX        =   2350
         _ExtentY        =   1460
         _StockProps     =   78
         Caption         =   "조 회"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "ANATO114.frx":0770
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   828
         Left            =   3960
         TabIndex        =   6
         Top             =   -24
         Width           =   1332
         _Version        =   65536
         _ExtentX        =   2350
         _ExtentY        =   1460
         _StockProps     =   78
         Caption         =   "종 료"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "ANATO114.frx":0BC2
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00C0C0C0&
      Height          =   852
      Left            =   3240
      ScaleHeight     =   795
      ScaleWidth      =   5250
      TabIndex        =   7
      Top             =   216
      Width           =   5310
      Begin VB.OptionButton optCode 
         Caption         =   "코 드 순"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   1005
         TabIndex        =   9
         Top             =   396
         Value           =   -1  'True
         Width           =   1250
      End
      Begin VB.OptionButton optName 
         Caption         =   "코드명순"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   2820
         TabIndex        =   8
         Top             =   396
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
         TabIndex        =   10
         Top             =   0
         Width           =   5265
      End
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  '평면
      BackColor       =   &H00808000&
      BorderStyle     =   1  '단일 고정
      Caption         =   "코드   코드명"
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
      Left            =   120
      TabIndex        =   18
      Top             =   210
      Width           =   3030
   End
End
Attribute VB_Name = "Anato_Dyeing_Code"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdDelete_Click()
    Dim Response

    Response = MsgBox(" 삭제하시겠습니까?", vbYesNo + vbCritical + vbDefaultButton2, "특수염색관리")
    If Response = vbYes Then
        'SPECODE_DELETE
        strSQL = ""
        strSQL = strSQL & " DELETE "
        strSQL = strSQL & " FROM   TWEXAM_SPECODE "
        strSQL = strSQL & " WHERE  Codegu = '55' "
        strSQL = strSQL & " AND    Codeky = '" & MidH(lstSpeCode.List(lstSpeCode.ListIndex), 1, 8) & "' "
    
        adoConnect.BeginTrans
        
        Result = AdoExecute(strSQL)
        
        If Result = True And Rowindicator > 0 Then
            adoConnect.CommitTrans
'            MsgBox "삭제 완료되었습니다.", vbInformation, "진단병리과"
        Else
            adoConnect.RollbackTrans
            MsgBox "작업도중 예기치 못한 오류가 발생했습니다.", vbCritical, "오류"
        End If
    
    End If
        Call Form_Load
    Exit Sub

End Sub


Private Sub cmdExit_Click()
    Unload Me

End Sub


Private Sub cmdSave_Click()

    If Trim(txtSpeCode) = "" Then Exit Sub
    
    strSQL = ""
    strSQL = strSQL & " SELECT *"
    strSQL = strSQL & " FROM   TWEXAM_SPECODE"
    strSQL = strSQL & " WHERE  Codegu  = '55'"
    strSQL = strSQL & " AND    Codeky  = '" & txtSpeCode & "'"

    Result = AdoOpenSet(rs, strSQL)
    
    If Result Then
        AdoCloseSet rs
        GoSub SPECODE_UPDATE
    Else
        GoSub SPECODE_INSERT
    End If
    
    Call Form_Load
    txtSpeCode.SetFocus
        
    Exit Sub
    

SPECODE_INSERT:
    strSQL = ""
    strSQL = strSQL & " INSERT INTO TWEXAM_SPECODE  "
    strSQL = strSQL & "       ( Codegu, Codeky, Codenm, Yageo)"
    strSQL = strSQL & " VALUES( '55', "
    strSQL = strSQL & "         '" & txtSpeCode.Text & "',"
    strSQL = strSQL & "         '" & Trim(txtCodeNm.Text) & "',"
    strSQL = strSQL & "         '" & txtAbbrev.Text & "')"
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
SPECODE_UPDATE:
    
    strSQL = ""
    strSQL = strSQL & " UPDATE TWEXAM_SPECODE"
    strSQL = strSQL & " SET    CODENM   = '" & txtCodeNm.Text & "',"
    strSQL = strSQL & "        YAGEO    = '" & txtAbbrev.Text & "'"
    strSQL = strSQL & " WHERE  CODEGU   = '55'"
    strSQL = strSQL & " AND    CODEKY   = '" & txtSpeCode.Text & "'"
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
    Dim I                   As Integer
    
    lstSpeCode.Clear
'---------------------------------------------'
'   검사코드 DB READ                             '
'---------------------------------------------'
    strSQL = " SELECT * FROM TWEXAM_Specode WHERE Codegu = '55'"
    
    If optCode.Value = True Then
        strSQL = strSQL & " ORDER BY CODEKY ASC "
    ElseIf optName.Value = True Then
        strSQL = strSQL & " ORDER BY CODENM ASC "
    End If
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then Exit Sub
    
    Do Until rs.EOF
        lstSpeCode.AddItem rs.Fields("Codeky").Value & "" & _
                           rs.Fields("Codenm").Value & ""
        rs.MoveNext
    Loop
    AdoCloseSet rs

End Sub


Private Sub Form_Load()
    Dim I                   As Integer
    
    lstSpeCode.Clear
    txtSpeCode = ""
    txtCodeNm = ""
    txtAbbrev = ""

'---------------------------------------------'
'   특수염색 DB READ                             '
'---------------------------------------------'
    strSQL = " SELECT * FROM TWEXAM_Specode Where Codegu = '55' ORDER BY Codeky"
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then Exit Sub
    
    Do Until rs.EOF
        lstSpeCode.AddItem rs.Fields("Codeky").Value & "" & _
                           rs.Fields("Codenm").Value & ""
        rs.MoveNext
    Loop
    AdoCloseSet rs

End Sub

Private Sub lstSpeCode_Click()
    Dim SearchText          As String * 8
     
    SearchText = MidH(lstSpeCode.List(lstSpeCode.ListIndex), 1, 8)
     
    If Trim(SearchText) = "" Then Exit Sub
          
'---------------------------------------------'
'   검사코드 DB READ                              '
'---------------------------------------------'
    strSQL = " SELECT * FROM TWEXAM_Specode WHERE Codegu = '55' AND Codeky = '" & SearchText & "'"
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then Exit Sub
  
    txtSpeCode = rs.Fields("Codeky").Value & ""
    txtCodeNm = rs.Fields("Codenm").Value & ""
    txtAbbrev = rs.Fields("Yageo").Value & ""

    AdoCloseSet rs

End Sub


Private Sub txtAbbrev_GotFocus()
    txtAbbrev.SelStart = 0
    txtAbbrev.SelLength = Len(txtAbbrev.Text)

End Sub


Private Sub txtAbbrev_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0
    
    SendKeys "{tab}"

End Sub


Private Sub txtCodeNm_GotFocus()
    txtCodeNm.SelStart = 0
    txtCodeNm.SelLength = Len(txtCodeNm.Text)

End Sub


Private Sub txtCodeNm_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0
    
    SendKeys "{tab}"

End Sub


Private Sub txtSpeCode_GotFocus()
    txtSpeCode.SelStart = 0
    txtSpeCode.SelLength = Len(txtSpeCode.Text)

End Sub


Private Sub txtSpeCode_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0
    
    SendKeys "{tab}"

End Sub


