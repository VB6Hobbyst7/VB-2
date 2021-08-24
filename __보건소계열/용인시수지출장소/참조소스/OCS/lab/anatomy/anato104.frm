VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Anato_Macro 
   BorderStyle     =   0  '없음
   Caption         =   "매크로관리"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   12015
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8310
   ScaleWidth      =   12015
   ShowInTaskbar   =   0   'False
   WindowState     =   2  '최대화
   Begin RichTextLib.RichTextBox txtFormat 
      Height          =   6084
      Left            =   1536
      TabIndex        =   4
      Top             =   1872
      Width           =   8220
      _ExtentX        =   14499
      _ExtentY        =   10716
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"ANATO104.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00C0C0C0&
      Height          =   1890
      Left            =   9870
      ScaleHeight     =   1830
      ScaleWidth      =   1905
      TabIndex        =   20
      Top             =   945
      Width           =   1965
      Begin VB.OptionButton optClass 
         Caption         =   " 분류순"
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
         Left            =   300
         TabIndex        =   24
         Top             =   1350
         Width           =   1230
      End
      Begin VB.OptionButton optOrgan 
         Caption         =   " 조직순"
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
         Left            =   300
         TabIndex        =   22
         Top             =   945
         Width           =   1230
      End
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
         Left            =   300
         TabIndex        =   21
         Top             =   495
         Value           =   -1  'True
         Width           =   1185
      End
      Begin VB.Label Label2 
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
         TabIndex        =   23
         Top             =   0
         Width           =   1890
      End
   End
   Begin VB.TextBox txtClass 
      BackColor       =   &H00E0E0E0&
      DataField       =   "FORMAT"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8370
      MaxLength       =   2
      TabIndex        =   3
      Top             =   1200
      Width           =   1365
   End
   Begin VB.TextBox txtDisease 
      BackColor       =   &H00E0E0E0&
      DataField       =   "FORMAT"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3900
      MaxLength       =   20
      TabIndex        =   2
      Top             =   1200
      Width           =   4485
   End
   Begin VB.TextBox txtOrgan 
      BackColor       =   &H00E0E0E0&
      DataField       =   "FORMAT"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1548
      MaxLength       =   10
      TabIndex        =   1
      Top             =   1200
      Width           =   2388
   End
   Begin VB.ListBox lstCode 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6300
      Left            =   96
      TabIndex        =   14
      Top             =   1560
      Width           =   1428
   End
   Begin VB.TextBox txtCode 
      BackColor       =   &H00C0C0FF&
      DataField       =   "FORMAT"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   90
      MaxLength       =   10
      TabIndex        =   0
      Top             =   1200
      Width           =   1476
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808000&
      Height          =   4755
      Left            =   9870
      ScaleHeight     =   4695
      ScaleWidth      =   1905
      TabIndex        =   9
      Top             =   3204
      Width           =   1965
      Begin Threed.SSCommand cmdExit 
         Height          =   1155
         Left            =   60
         TabIndex        =   8
         Top             =   3480
         Width           =   1755
         _Version        =   65536
         _ExtentX        =   3096
         _ExtentY        =   2037
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
         Picture         =   "ANATO104.frx":026D
      End
      Begin Threed.SSCommand cmdView 
         Height          =   1155
         Left            =   60
         TabIndex        =   6
         Top             =   1200
         Width           =   1755
         _Version        =   65536
         _ExtentX        =   3096
         _ExtentY        =   2037
         _StockProps     =   78
         Caption         =   "조 회"
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
         Picture         =   "ANATO104.frx":0587
      End
      Begin Threed.SSCommand cmdSave 
         Height          =   1155
         Left            =   60
         TabIndex        =   5
         Top             =   60
         Width           =   1755
         _Version        =   65536
         _ExtentX        =   3096
         _ExtentY        =   2037
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
         Picture         =   "ANATO104.frx":09D9
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   1155
         Left            =   60
         TabIndex        =   7
         Top             =   2340
         Width           =   1755
         _Version        =   65536
         _ExtentX        =   3096
         _ExtentY        =   2037
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
         Picture         =   "ANATO104.frx":0E2B
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Align           =   1  '위 맞춤
      Height          =   768
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   12012
      _Version        =   65536
      _ExtentX        =   21188
      _ExtentY        =   1355
      _StockProps     =   15
      Caption         =   "ANATOMIC   PATHOLOGY"
      ForeColor       =   8388608
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelInner      =   1
      RoundedCorners  =   0   'False
      Outline         =   -1  'True
      Font3D          =   2
      Begin VB.PictureBox Picture3 
         Height          =   405
         Left            =   9660
         ScaleHeight     =   345
         ScaleWidth      =   2115
         TabIndex        =   11
         Top             =   180
         Width           =   2175
         Begin VB.Label lblUser 
            AutoSize        =   -1  'True
            Caption         =   "********"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   180
            Left            =   945
            TabIndex        =   13
            Top             =   90
            Width           =   720
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "User:"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   180
            Left            =   180
            TabIndex        =   12
            Top             =   90
            Width           =   450
         End
      End
   End
   Begin VB.Label Label8 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00808000&
      BorderStyle     =   1  '단일 고정
      Caption         =   "분   류"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   8370
      TabIndex        =   19
      Top             =   975
      Width           =   1365
   End
   Begin VB.Label Label7 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00808000&
      BorderStyle     =   1  '단일 고정
      Caption         =   "조    직"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   1548
      TabIndex        =   18
      Top             =   972
      Width           =   2388
   End
   Begin VB.Label Label5 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00808000&
      BorderStyle     =   1  '단일 고정
      Caption         =   "F   O   R   M   A   T"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   1560
      TabIndex        =   17
      Top             =   1560
      Width           =   8184
   End
   Begin VB.Label Label4 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00808000&
      BorderStyle     =   1  '단일 고정
      Caption         =   "코  드"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   96
      TabIndex        =   16
      Top             =   972
      Width           =   1476
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00808000&
      BorderStyle     =   1  '단일 고정
      Caption         =   "질   병   명"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   3900
      TabIndex        =   15
      Top             =   972
      Width           =   4488
   End
End
Attribute VB_Name = "Anato_Macro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDelete_Click()

    Dim Response            As Integer
    
    If Trim(txtCode) = "" Then Exit Sub
    
    Response = MsgBox("자료를 삭제할까요?", vbYesNo + vbQuestion + vbDefaultButton2, "진단병리")
  
    If Response = vbNo Then Exit Sub
    
    strSQL = " DELETE FROM TWANAT_MACRO WHERE Code = '" & txtCode.Text & "'"
    
    adoConnect.BeginTrans
    
    Result = AdoExecute(strSQL)
    
    If Result = True And Rowindicator > 0 Then
        adoConnect.CommitTrans
'        MsgBox "삭제 완료되었습니다.", vbInformation, "진단병리과"
    Else
        adoConnect.RollbackTrans
        MsgBox "작업도중 예기치 못한 오류가 발생했습니다.", vbCritical, "오류"
    End If
    
    Call Form_Load
 
End Sub

Private Sub cmdExit_Click()
    Unload Me

End Sub

Private Sub cmdSave_Click()

    Dim Response            As Integer
    
    If Trim(txtCode) = "" Then Exit Sub
    
    Response = MsgBox("자료를 저장하시겠습니까?", vbYesNo + vbQuestion + vbDefaultButton2, "진단병리")
    
    If Response = vbNo Then Exit Sub
    
    strSQL = " SELECT * FROM TWANAT_Macro WHERE Code = '" & txtCode.Text & "'"
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result Then
        AdoCloseSet rs
        GoSub CODE_UPDATE
    Else
        GoSub CODE_INSERT
    End If
    
    Call Form_Load
    txtCode.SetFocus
    
    Exit Sub
    
'/-------------------------------------------------------------------------------
CODE_INSERT:
    strSQL = ""
    strSQL = strSQL & " INSERT INTO TWANAT_MACRO  "
    strSQL = strSQL & "       ( Code,organ, Disease, Class, Format)"
    strSQL = strSQL & " VALUES('" & txtCode.Text & "',"
    strSQL = strSQL & "        '" & Trim(txtOrgan.Text) & "',"
    strSQL = strSQL & "        '" & Quot(Trim(txtDisease.Text)) & "',"
    strSQL = strSQL & "        '" & Trim(txtClass.Text) & "',"
    strSQL = strSQL & "        '" & Quot(Trim(txtFormat.Text)) & "')"
    
    adoConnect.BeginTrans
    
''    quot :   DATA의 중간에 "'" 값이 있을경우 data convert
    
    Result = AdoExecute(strSQL)
    
    If Result = True And Rowindicator > 0 Then
        adoConnect.CommitTrans
'        MsgBox "저장 완료되었습니다.", vbInformation, "진단병리과"
    Else
        adoConnect.RollbackTrans
        MsgBox "작업도중 예기치 못한 오류가 발생했습니다.", vbCritical, "오류"
    End If
    
    Return

'----------------------------------------------------------------------------
CODE_UPDATE:
    strSQL = ""
    strSQL = strSQL & " UPDATE TWANAT_MACRO"
    strSQL = strSQL & " SET    ORGAN      = '" & Trim(txtOrgan.Text) & "',"
    strSQL = strSQL & "        DISEASE    = '" & Quot(Trim(txtDisease.Text)) & "',"
    strSQL = strSQL & "        CLASS      = '" & Trim(txtClass.Text) & "',"
    strSQL = strSQL & "        FORMAT     = '" & Quot(Trim(txtFormat.Text)) & "'"
    strSQL = strSQL & " WHERE  CODE       = '" & txtCode.Text & "'"
    
    adoConnect.BeginTrans
    
    Result = AdoExecute(strSQL)
    
    If Result = True And Rowindicator > 0 Then
        adoConnect.CommitTrans
'        MsgBox "저장 완료되었습니다.", vbInformation, "진단병리과"
    Else
        adoConnect.RollbackTrans
        MsgBox "작업도중 예기치 못한 오류가 발생했습니다.", vbCritical, "오류"
    End If
    
    Return

End Sub

Private Sub cmdView_Click()

    Dim rs                  As ADODB.Recordset
    
    Dim i                   As Integer
        
    strSQL = ""
    strSQL = strSQL & " SELECT Code "
    strSQL = strSQL & " FROM   TWANAT_Macro "
    If optCode = True Then
        strSQL = strSQL & " ORDER BY  CODE  ASC      "
    ElseIf optOrgan = True Then
        strSQL = strSQL & " ORDER BY  ORGAN ASC      "
    ElseIf optClass = True Then
        strSQL = strSQL & " ORDER BY  CLASS ASC      "
    End If
    
    lstCode.Clear
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then Exit Sub
    
    Do Until rs.EOF
        lstCode.AddItem rs.Fields("CODE").Value & ""
        rs.MoveNext
    Loop
    AdoCloseSet rs
 
End Sub


Private Sub Form_Load()
    
    Dim rs                  As ADODB.Recordset
    
    Dim i                   As Integer
    
    lblUser = GstrPassName

    strSQL = " SELECT Code FROM TWANAT_Macro ORDER BY Code"
    
    lstCode.Clear
    txtCode = ""
    txtOrgan = ""
    txtDisease = ""
    txtClass = ""
    txtFormat.Text = ""
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then Exit Sub
    
    Do Until rs.EOF
        lstCode.AddItem rs.Fields("Code").Value & ""
        rs.MoveNext
    Loop
    AdoCloseSet rs
    
 
End Sub

Private Sub lstCode_Click()

    txtCode = lstCode.List(lstCode.ListIndex)
    Call txtCode_LostFocus
    
End Sub


Private Sub txtClass_GotFocus()

    txtClass.SelStart = 0
    txtClass.SelLength = Len(txtClass.Text)

End Sub

Private Sub TXTCLASS_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0
    
    SendKeys "{tab}"


End Sub


Private Sub txtCode_GotFocus()

    txtCode.SelStart = 0
    txtCode.SelLength = Len(txtCode.Text)

End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0
    
    SendKeys "{tab}"


End Sub


Private Sub txtCode_LostFocus()

    Dim rs                  As ADODB.Recordset
    
    strSQL = ""
    strSQL = strSQL & " SELECT * "
    strSQL = strSQL & "   FROM TWANAT_MACRO "
    strSQL = strSQL & "  WHERE Code = '" & txtCode & "'"
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then
        txtDisease = ""
        txtClass = ""
        txtFormat.Text = ""
        txtOrgan = ""
        Exit Sub
    End If
    
    txtOrgan = rs.Fields("ORGAN").Value & ""
    txtDisease = rs.Fields("DISEASE").Value & ""
    txtClass = rs.Fields("CLASS").Value & ""
    txtFormat.Text = rs.Fields("FORMAT").Value & ""
    
    AdoCloseSet rs
    
End Sub

Private Sub txtDisease_GotFocus()
    
    txtDisease.SelStart = 0
    txtDisease.SelLength = Len(txtDisease.Text)

End Sub

Private Sub txtDisease_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0
    
    SendKeys "{tab}"


End Sub



Private Sub txtOrgan_GotFocus()
    
    txtOrgan.SelStart = 0
    txtOrgan.SelLength = Len(txtOrgan.Text)

End Sub

Private Sub txtOrgan_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0
    
    SendKeys "{tab}"


End Sub


