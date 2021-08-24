VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form Anato_Code 
   Caption         =   "검사코드관리"
   ClientHeight    =   7170
   ClientLeft      =   1470
   ClientTop       =   2100
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7170
   ScaleWidth      =   10380
   Begin Threed.SSFrame SSFrame1 
      Height          =   660
      Left            =   144
      TabIndex        =   23
      Top             =   120
      Width           =   10044
      _Version        =   65536
      _ExtentX        =   17716
      _ExtentY        =   1164
      _StockProps     =   14
      Caption         =   "특수검사항목"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSOption SSOption1 
         Height          =   228
         Index           =   0
         Left            =   336
         TabIndex        =   24
         Top             =   312
         Width           =   1428
         _Version        =   65536
         _ExtentX        =   2519
         _ExtentY        =   402
         _StockProps     =   78
         Caption         =   "검사코드(H)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption SSOption1 
         Height          =   228
         Index           =   1
         Left            =   3508
         TabIndex        =   25
         Top             =   312
         Width           =   1428
         _Version        =   65536
         _ExtentX        =   2519
         _ExtentY        =   402
         _StockProps     =   78
         Caption         =   "특수염색"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption SSOption1 
         Height          =   228
         Index           =   2
         Left            =   5094
         TabIndex        =   26
         Top             =   312
         Width           =   1428
         _Version        =   65536
         _ExtentX        =   2519
         _ExtentY        =   402
         _StockProps     =   78
         Caption         =   "면역염색"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption SSOption1 
         Height          =   228
         Index           =   3
         Left            =   6680
         TabIndex        =   27
         Top             =   312
         Width           =   1428
         _Version        =   65536
         _ExtentX        =   2519
         _ExtentY        =   402
         _StockProps     =   78
         Caption         =   "면역형광염색"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption SSOption1 
         Height          =   228
         Index           =   4
         Left            =   8268
         TabIndex        =   28
         Top             =   312
         Width           =   1428
         _Version        =   65536
         _ExtentX        =   2519
         _ExtentY        =   402
         _StockProps     =   78
         Caption         =   "효소염색"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption SSOption1 
         Height          =   228
         Index           =   5
         Left            =   1922
         TabIndex        =   29
         Top             =   312
         Width           =   1428
         _Version        =   65536
         _ExtentX        =   2519
         _ExtentY        =   402
         _StockProps     =   78
         Caption         =   "검사코드(C)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00C0C0C0&
      Height          =   1476
      Left            =   8832
      ScaleHeight     =   1410
      ScaleWidth      =   1305
      TabIndex        =   18
      Top             =   864
      Width           =   1365
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
         Height          =   285
         Left            =   90
         TabIndex        =   20
         Top             =   945
         Width           =   1020
      End
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
         Height          =   285
         Left            =   90
         TabIndex        =   19
         Top             =   525
         Value           =   -1  'True
         Width           =   1095
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
         TabIndex        =   21
         Top             =   0
         Width           =   1320
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808000&
      Height          =   4515
      Left            =   8832
      ScaleHeight     =   4455
      ScaleWidth      =   1305
      TabIndex        =   17
      Top             =   2436
      Width           =   1365
      Begin Threed.SSCommand cmdExit 
         Height          =   1020
         Left            =   0
         TabIndex        =   8
         Top             =   3060
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   1799
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
         Picture         =   "ANATO110.frx":0000
      End
      Begin Threed.SSCommand cmdView 
         Height          =   1020
         Left            =   0
         TabIndex        =   6
         Top             =   1020
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   1799
         _StockProps     =   78
         Caption         =   "조 회"
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
         Picture         =   "ANATO110.frx":031A
      End
      Begin Threed.SSCommand cmdSave 
         Height          =   1020
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   1799
         _StockProps     =   78
         Caption         =   "등 록"
         ForeColor       =   8388736
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
         Picture         =   "ANATO110.frx":076C
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   1020
         Left            =   0
         TabIndex        =   7
         Top             =   2040
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   1799
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
         Enabled         =   0   'False
         Font3D          =   3
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "ANATO110.frx":0BBE
      End
   End
   Begin VB.ListBox lstExCode 
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
      Height          =   5520
      Left            =   144
      TabIndex        =   16
      Top             =   1200
      Width           =   3828
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00C0C0C0&
      Height          =   6108
      Left            =   4080
      ScaleHeight     =   6045
      ScaleWidth      =   4575
      TabIndex        =   9
      Top             =   864
      Width           =   4632
      Begin VB.TextBox txtItemnm 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Left            =   1224
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   2115
         Width           =   3288
      End
      Begin VB.TextBox txtItemCD 
         BackColor       =   &H00C0FFC0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1224
         MaxLength       =   8
         TabIndex        =   0
         Top             =   1080
         Width           =   1100
      End
      Begin VB.TextBox txtSuga 
         BackColor       =   &H00C0C0FF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1224
         MaxLength       =   8
         TabIndex        =   3
         Top             =   4176
         Width           =   3288
      End
      Begin VB.TextBox txtYagEo 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         Left            =   1224
         MaxLength       =   20
         TabIndex        =   2
         Top             =   3588
         Width           =   3288
      End
      Begin VB.TextBox txtitemh 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Left            =   1224
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   2844
         Width           =   3288
      End
      Begin VB.Label Label4 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "코 드 명"
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
         Left            =   168
         TabIndex        =   15
         Top             =   2208
         Width           =   1056
      End
      Begin VB.Label Label2 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "코    드"
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
         Left            =   168
         TabIndex        =   14
         Top             =   1128
         Width           =   720
      End
      Begin VB.Label Label3 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "수가코드"
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
         Left            =   168
         TabIndex        =   13
         Top             =   4248
         Width           =   1056
      End
      Begin VB.Label Label5 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "약    어"
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
         Left            =   168
         TabIndex        =   12
         Top             =   3660
         Width           =   1056
      End
      Begin VB.Label Label7 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "코드명(한)"
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
         Left            =   168
         TabIndex        =   11
         Top             =   2916
         Width           =   1056
      End
      Begin VB.Label Label8 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00808000&
         BorderStyle     =   1  '단일 고정
         Caption         =   "검 사 코 드 정 보 등 록"
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
         Width           =   5280
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
      Left            =   144
      TabIndex        =   22
      Top             =   888
      Width           =   3828
   End
End
Attribute VB_Name = "Anato_Code"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'    Dim SearchText          As String * 8
    Dim SearchText          As String


Private Sub Form_Load()
        
    lstExCode.Clear
    txtItemCD.Text = ""
    txtItemnm.Text = ""
    txtYagEo.Text = ""
    txtitemh.Text = ""
    txtSuga.Text = ""
    
    SSOption1(0).Value = True

'---------------------------------------------'
'   검사코드 DB READ                             '
'---------------------------------------------'
    
    Call cmdView_Click
  
End Sub


Private Sub cmdExit_Click()
    Unload Me

End Sub


Private Sub SSOption1_Click(Index As Integer, Value As Integer)
    
    txtItemCD.Text = ""
    txtItemnm.Text = ""
    txtYagEo.Text = ""
    txtitemh.Text = ""
    txtSuga.Text = ""
    
    Select Case Index
           Case 0
                GCodegu = "80"      '검사코드(Histology)
           Case 1
                GCodegu = "83"      '특수염색
           Case 2
                GCodegu = "87"      '면역염색
           Case 3
                GCodegu = "84"      '면역형광염색
           Case 4
                GCodegu = "86"      '효소면역
           Case 5
                GCodegu = "89"      '검사코드(Cytology)
    End Select
    
    Call cmdView_Click

End Sub


Private Sub cmdView_Click()
        
    lstExCode.Clear
    
'---------------------------------------------'
'   검사코드 DB READ                             '
'---------------------------------------------'
              '  GCodegu = "80"      '검사코드(Histology)
              '  GCodegu = "83"      '특수염색
              '  GCodegu = "87"      '면역염색
              '  GCodegu = "84"      '면역형광염색
              '  GCodegu = "86"      '효소면역
              '  GCodegu = "89"      '검사코드(Cytology)
    
    strSQL = ""
    strSQL = strSQL & " SELECT * "
    strSQL = strSQL & "   FROM TWEXAM_ITEMML "
    Select Case GCodegu
            Case "80", "89"
                strSQL = strSQL & "  WHERE Codegu = '" & GCodegu & "' "
            Case "83"
                strSQL = strSQL & "  WHERE Codeky between '853001' and '853999' "
            Case "84"
                strSQL = strSQL & "  WHERE Codeky between '854001' and '854999' "
            Case "86"
                strSQL = strSQL & "  WHERE Codeky between '856001' and '856999' "
            Case "87"
                strSQL = strSQL & "  WHERE Codeky between '857001' and '857999' "
    End Select
    
    If optCode.Value = True Then
        strSQL = strSQL & " ORDER BY CODEKY ASC "
    ElseIf optName.Value = True Then
        strSQL = strSQL & " ORDER BY ITEMNM ASC "
    End If
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then Exit Sub
    
    Do Until rs.EOF
        lstExCode.AddItem Format(Trim(rs.Fields("Codeky").Value) & "", "@@@") & " " & rs.Fields("itemNm").Value & ""
        rs.MoveNext
    Loop
    AdoCloseSet rs
  
End Sub


Private Sub lstExCode_Click()
    
    SearchText = Mid(lstExCode.List(lstExCode.ListIndex), 1, 6)
     
    If Trim(SearchText) = "" Then Exit Sub

'---------------------------------------------'
'   검사코드 DB READ                              '
'---------------------------------------------'
    strSQL = ""
    strSQL = strSQL & " SELECT * "
'    strSQL = strSQL & " FROM   TWEXAM_SPECODE "
    strSQL = strSQL & "   FROM TWEXAM_ITEMML "
    strSQL = strSQL & " WHERE  Codeky = '" & Trim(SearchText) & "' "
'
    Result = AdoOpenSet(rs, strSQL)
    If Result = False Then Exit Sub
  
    Do Until rs.EOF
       txtItemCD = rs.Fields("Codeky").Value & ""
       txtItemnm = rs.Fields("itemNm").Value & ""
       txtitemh = rs.Fields("itemko").Value & ""
       txtYagEo = rs.Fields("Yageo").Value & ""
       txtSuga = rs.Fields("Sugacd").Value & ""
'       txtEditStyle = rs.Fields("Jangbi3").Value & ""
       rs.MoveNext
    Loop
    AdoCloseSet rs
    
End Sub


Private Sub cmdSave_Click()
    Dim SearchText          As String * 8
    
    If Trim(txtItemCD.Text) = "" Then Exit Sub
    
    SearchText = Mid(lstExCode.List(lstExCode.ListIndex), 1, 6)
    
    strSQL = ""
    strSQL = strSQL & " UPDATE TWEXAM_ITEMML "
    strSQL = strSQL & " SET    ITEMNM   = '" & Trim(txtItemnm.Text) & "',"
    strSQL = strSQL & "        ITEMKO   = '" & Trim(txtitemh.Text) & "',"
    strSQL = strSQL & "        YAGEO    = '" & Trim(txtYagEo.Text) & "' "
    strSQL = strSQL & " WHERE  CODEKY   = '" & Trim(txtItemCD.Text) & "'"
    
    adoConnect.BeginTrans
    
    Result = AdoExecute(strSQL)
    
    If Result = True And Rowindicator > 0 Then
        adoConnect.CommitTrans
        MsgBox "저장 완료되었습니다.", vbInformation, "진단병리과"
    Else
        adoConnect.RollbackTrans
        MsgBox "작업도중 예기치 못한 오류가 발생했습니다.", vbCritical, "오류"
    End If
    
    Call Form_Load
'    txtItemCD.SetFocus
        
    Exit Sub


End Sub


Private Sub cmdDelete_Click()
    
    Dim Response

    Dim SearchText          As String * 8
     
    SearchText = Mid(lstExCode.List(lstExCode.ListIndex), 1, 3)
    
    Response = MsgBox(" 삭제하시겠습니까?", vbYesNo + vbCritical + vbDefaultButton2, "검사코드관리")
    If Response = vbYes Then
        'CODE_DELETE
        strSQL = ""
        strSQL = strSQL & " DELETE "
        strSQL = strSQL & " FROM   TWEXAM_SPECODE "
        strSQL = strSQL & " WHERE  CODEGU = '" & GCodegu & "' "
        strSQL = strSQL & " AND    CODEKY = '" & Trim(SearchText) & "' "
    
        adoConnect.BeginTrans
        
        Result = AdoExecute(strSQL)
        
        If Result = True And Rowindicator > 0 Then
            adoConnect.CommitTrans
            MsgBox "삭제 완료되었습니다.", vbInformation, "진단병리과"
        Else
            adoConnect.RollbackTrans
            MsgBox "작업도중 예기치 못한 오류가 발생했습니다.", vbCritical, "오류"
        End If
    
    End If
    
    Call Form_Load

End Sub


'Private Sub txtEditStyle_GotFocus()
'    txtEditStyle.SelStart = 0
'    txtEditStyle.SelLength = Len(txtEditStyle.Text)'
'
'End Sub


Private Sub txtEditStyle_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0
    
    SendKeys "{tab}"

End Sub


'Private Sub txtExCode_GotFocus()
'
'    txtexcode.SelStart = 0
'    txtexcode.SelLength = Len(txtexcode.Text)
'
'End Sub


'Private Sub txtExCode_KeyPress(KeyAscii As Integer)
'
'    If KeyAscii <> 13 Then Exit Sub
'    KeyAscii = 0
'
'    SendKeys "{tab}"
'
'End Sub


Private Sub txtITEMCD_GotFocus()

    txtItemCD.SelStart = 0
    txtItemCD.SelLength = Len(txtItemCD.Text)

End Sub


Private Sub txtITEMCD_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0
    
    SendKeys "{tab}"

End Sub


Private Sub txtItemnm_GotFocus()

    txtItemnm.SelStart = 0
    txtItemnm.SelLength = Len(txtItemnm.Text)

End Sub


Private Sub txtItemnm_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0
    
    SendKeys "{tab}"

End Sub


'Private Sub txtProcess_GotFocus()
'
'    txtProcess.SelStart = 0
'    txtProcess.SelLength = Len(txtProcess.Text)'
'
'End Sub


'Private Sub txtProcess_KeyPress(KeyAscii As Integer)
'    If KeyAscii <> 13 Then Exit Sub
'    KeyAscii = 0
'
'    SendKeys "{tab}"
'
'End Sub


'Private Sub txtTitle_GotFocus()
'    txtTitle.SelStart = 0
'    txtTitle.SelLength = Len(txtTitle.Text)'
'
'End Sub


'Private Sub txtTitle_KeyPress(KeyAscii As Integer)
'    If KeyAscii <> 13 Then Exit Sub
'    KeyAscii = 0
'
'    SendKeys "{tab}"
'
'End Sub


Private Sub txtYagEo_GotFocus()
    txtYagEo.SelStart = 0
    txtYagEo.SelLength = Len(txtYagEo.Text)

End Sub
