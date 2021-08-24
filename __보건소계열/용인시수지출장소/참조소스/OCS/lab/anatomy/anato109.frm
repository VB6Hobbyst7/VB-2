VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form Anato_Dict 
   Caption         =   "병리사전관리"
   ClientHeight    =   5010
   ClientLeft      =   75
   ClientTop       =   1905
   ClientWidth     =   11550
   ControlBox      =   0   'False
   Icon            =   "ANATO109.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5010
   ScaleWidth      =   11550
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00C0C0C0&
      Height          =   1476
      Left            =   5340
      ScaleHeight     =   1410
      ScaleWidth      =   2025
      TabIndex        =   17
      Top             =   360
      Width           =   2085
      Begin VB.OptionButton optcode1 
         Caption         =   "부위코드"
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
         Index           =   1
         Left            =   240
         TabIndex        =   20
         Top             =   930
         Width           =   1095
      End
      Begin VB.OptionButton optcode1 
         Caption         =   "병리코드"
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
         Index           =   0
         Left            =   240
         TabIndex        =   18
         Top             =   525
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00808000&
         BorderStyle     =   1  '단일 고정
         Caption         =   "조회조건"
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
         TabIndex        =   19
         Top             =   0
         Width           =   2040
      End
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00C0C0C0&
      Height          =   2445
      Left            =   5340
      ScaleHeight     =   2385
      ScaleWidth      =   4575
      TabIndex        =   11
      Top             =   2250
      Width           =   4632
      Begin VB.TextBox txtCode 
         BackColor       =   &H00C0FFC0&
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
         TabIndex        =   12
         Top             =   660
         Width           =   1100
      End
      Begin VB.TextBox txtCodenm 
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
         Height          =   1005
         Left            =   1224
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   1245
         Width           =   3288
      End
      Begin VB.Label Label8 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00808000&
         BorderStyle     =   1  '단일 고정
         Caption         =   "병 리 사 전 코 드 등 록"
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
         TabIndex        =   16
         Top             =   0
         Width           =   5280
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
         Left            =   165
         TabIndex        =   15
         Top             =   705
         Width           =   720
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
         Left            =   165
         TabIndex        =   13
         Top             =   1335
         Width           =   1050
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00C0C0C0&
      Height          =   1476
      Left            =   7890
      ScaleHeight     =   1410
      ScaleWidth      =   2025
      TabIndex        =   5
      Top             =   360
      Width           =   2085
      Begin VB.OptionButton optName 
         Caption         =   "진단명순"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   7
         Top             =   945
         Width           =   1230
      End
      Begin VB.OptionButton optCode 
         Caption         =   "코드순"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   6
         Top             =   495
         Value           =   -1  'True
         Width           =   1005
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
         Height          =   330
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   2040
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808000&
      Height          =   4455
      Left            =   10080
      ScaleHeight     =   4395
      ScaleWidth      =   1305
      TabIndex        =   1
      Top             =   270
      Width           =   1365
      Begin Threed.SSCommand cmdExit 
         Height          =   1020
         Left            =   0
         TabIndex        =   3
         Top             =   3360
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
         Picture         =   "ANATO109.frx":0442
      End
      Begin Threed.SSCommand cmdView 
         Height          =   1020
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   1799
         _StockProps     =   78
         Caption         =   "조 회"
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
         Picture         =   "ANATO109.frx":075C
      End
      Begin Threed.SSCommand cmdSave 
         Height          =   1020
         Left            =   0
         TabIndex        =   9
         Top             =   1125
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
         Picture         =   "ANATO109.frx":0BAE
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   1020
         Left            =   0
         TabIndex        =   10
         Top             =   2235
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
         Font3D          =   3
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "ANATO109.frx":1000
      End
   End
   Begin VB.ListBox lstDxDict 
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
      Height          =   3960
      Left            =   75
      TabIndex        =   0
      Top             =   720
      Width           =   5145
   End
   Begin VB.Label lblTitle 
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
      Height          =   330
      Left            =   75
      TabIndex        =   4
      Top             =   360
      Width           =   5145
   End
End
Attribute VB_Name = "Anato_Dict"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Dim LsCode              As String * 10
    Dim LsName              As String * 60
    Dim LsClass             As String * 2


Private Sub cmdDelete_Click()

    Dim Response            As Integer
    
    If Trim(txtCode.Text) = "" And Trim(txtCodenm.Text) = "" Then Exit Sub
    
    Response = MsgBox("자료를 삭제할까요?", vbYesNo + vbQuestion + vbDefaultButton2, "진단병리")
  
    If Response = vbNo Then Exit Sub
    
    strSQL = " DELETE FROM TWANAT_DICT WHERE Code = '" & Trim(txtCode.Text) & "'"
    
    adoConnect.BeginTrans
    
    Result = AdoExecute(strSQL)
    
    If Result = True And Rowindicator > 0 Then
        adoConnect.CommitTrans
        MsgBox "삭제 완료되었습니다.", vbInformation, "진단병리과"
        txtCode.Text = ""
        txtCodenm.Text = ""
    Else
        adoConnect.RollbackTrans
        MsgBox "작업도중 예기치 못한 오류가 발생했습니다.", vbCritical, "오류"
    End If
    
End Sub

Private Sub cmdExit_Click()
    Unload Me
    
End Sub


Private Sub cmdSave_Click()
    
    If Trim(txtCode.Text) = "" Then Exit Sub
    
    '조회
    strSQL = ""
    strSQL = strSQL & " SELECT * "
    strSQL = strSQL & "   FROM TWANAT_Dict "
    strSQL = strSQL & "  WHERE  CODE   = '" & Trim(txtCode.Text) & "'"
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then
        'INSERT
    
        strSQL = ""
        strSQL = strSQL & " INSERT INTO TWANAT_DICT "
        strSQL = strSQL & "  ( CODE, DXDICT ) "
        strSQL = strSQL & "VALUES ('" & Trim(txtCode.Text) & "',"
        strSQL = strSQL & "        '" & Trim(txtCodenm.Text) & "')"
        
        adoConnect.BeginTrans
    
        Result = AdoExecute(strSQL)
        
        If Result = True And Rowindicator > 0 Then
            adoConnect.CommitTrans
            MsgBox "저장 완료되었습니다.", vbInformation, "진단병리과"
        Else
            adoConnect.RollbackTrans
            MsgBox "작업도중 예기치 못한 오류가 발생했습니다.", vbCritical, "오류"
        End If
    
    Else
        'UPDATE
        strSQL = ""
        strSQL = strSQL & " UPDATE TWANAT_DICT "
        strSQL = strSQL & " SET    DXDICT = '" & Trim(txtCodenm.Text) & "' "
        strSQL = strSQL & " WHERE  CODE   = '" & Trim(txtCode.Text) & "' "
        
        adoConnect.BeginTrans
        
        Result = AdoExecute(strSQL)
        
        If Result = True And Rowindicator > 0 Then
            adoConnect.CommitTrans
            MsgBox "저장 완료되었습니다.", vbInformation, "진단병리과"
        Else
            adoConnect.RollbackTrans
            MsgBox "작업도중 예기치 못한 오류가 발생했습니다.", vbCritical, "오류"
        End If
    
    End If
    
    
'    Call Form_Load
'    txtItemCD.SetFocus
        
    Exit Sub



End Sub

Private Sub cmdView_Click()

    Dim i                   As Integer
    Dim LsTitCode           As String * 10
    Dim LsTitName           As String * 54
        
    txtCode.Text = ""
    txtCodenm.Text = ""
    
    
    LsTitCode = "코 드 명"
    LsTitName = "진 단 병 리 명"
    
    If optCode = True Then
        lblTitle = LsTitCode & LsTitName
        strSQL = ""
        strSQL = strSQL & " SELECT * "
        strSQL = strSQL & "   FROM TWANAT_Dict "
        If optcode1(0).Value = True Then
            strSQL = strSQL & "  WHERE SUBSTR(CODE,1,1)  = 'M' "
        Else
            strSQL = strSQL & "  WHERE SUBSTR(CODE,1,1)  = 'T' "
        End If
        strSQL = strSQL & "  ORDER BY Code"
    Else
        lblTitle = LsTitName & LsTitCode
        strSQL = ""
        strSQL = strSQL & " SELECT * "
        strSQL = strSQL & "   FROM TWANAT_Dict "
        If optcode1(0).Value = True Then
            strSQL = strSQL & "  WHERE SUBSTR(CODE,1,1)  = 'M' "
        Else
            strSQL = strSQL & "  WHERE SUBSTR(CODE,1,1)  = 'T' "
        End If
        strSQL = strSQL & "  ORDER BY Dxdict"
    End If
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then Exit Sub
    
    lstDxDict.Clear
    Do Until rs.EOF
        LsCode = rs.Fields("Code").Value & ""
        LsName = rs.Fields("DxDict").Value & ""
        
        If optCode = True Then
            lstDxDict.AddItem LsCode & LsName
        Else
            lstDxDict.AddItem LsName & LsCode
        End If
        rs.MoveNext
    Loop
    AdoCloseSet rs
    
End Sub

Private Sub Form_Load()
    optcode1(0).Value = True

End Sub

Private Sub lstDxDict_Click()
    
    If optCode = True Then
         LsCode = Mid(lstDxDict.List(lstDxDict.ListIndex), 1, 10)
         LsName = Mid(lstDxDict.List(lstDxDict.ListIndex), 11, 60)
    Else
         LsCode = Mid(lstDxDict.List(lstDxDict.ListIndex), 11, 60)
         LsName = Mid(lstDxDict.List(lstDxDict.ListIndex), 61, 10)
    End If
  
       txtCode.Text = LsCode
       txtCodenm.Text = LsName

End Sub


Private Sub txtCode_GotFocus()
    
    If txtCode.Text <> "" Then
        MsgBox " 코드는 수정 할 수 없습니다. " & vbCrLf & vbCrLf & _
               " 코드를 수정하는 것은 신규코드를 등록하는 것과 같습니다." & vbCrLf & vbCrLf & _
               " 기존 코드가 필요 없으면 삭제하십시요."
    End If
    
    txtCode.SelStart = 0
    txtCode.SelLength = Len(txtCode.Text)

End Sub

Private Sub txtCode_LostFocus()
    txtCode.Text = UCase(txtCode.Text)

End Sub


Private Sub txtCodenm_GotFocus()
    txtCodenm.SelStart = 0
    txtCodenm.SelLength = Len(txtCodenm.Text)

End Sub
