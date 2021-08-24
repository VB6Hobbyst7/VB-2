VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmEquipExam 
   Caption         =   "장비 코드 설정"
   ClientHeight    =   8505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11925
   LinkTopic       =   "Form1"
   ScaleHeight     =   8505
   ScaleWidth      =   11925
   StartUpPosition =   2  '화면 가운데
   Begin VB.Frame Frame1 
      Height          =   5745
      Left            =   7605
      TabIndex        =   1
      Top             =   780
      Width           =   4275
      Begin VB.ComboBox cboOrdGubun 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmEquipExam.frx":0000
         Left            =   1410
         List            =   "frmEquipExam.frx":0010
         TabIndex        =   28
         Top             =   3750
         Width           =   2685
      End
      Begin VB.ComboBox ComPart 
         BeginProperty Font 
            Name            =   "새굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1410
         TabIndex        =   27
         Top             =   3360
         Width           =   2685
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "종료"
         BeginProperty Font 
            Name            =   "새굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3150
         TabIndex        =   13
         Top             =   4890
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "새굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2160
         TabIndex        =   12
         Top             =   4890
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "삭제"
         BeginProperty Font 
            Name            =   "새굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1170
         TabIndex        =   11
         Top             =   4890
         Width           =   975
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "저장"
         BeginProperty Font 
            Name            =   "새굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   210
         TabIndex        =   10
         Top             =   4890
         Width           =   975
      End
      Begin VB.TextBox txtOcsCode 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1410
         TabIndex        =   24
         Top             =   4320
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.TextBox txtSeq 
         BeginProperty Font 
            Name            =   "새굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1410
         TabIndex        =   22
         Top             =   2400
         Width           =   2655
      End
      Begin VB.TextBox txtRSCode 
         BeginProperty Font 
            Name            =   "새굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1410
         TabIndex        =   19
         Top             =   1530
         Width           =   2655
      End
      Begin VB.ComboBox cboGubun 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmEquipExam.frx":005F
         Left            =   1020
         List            =   "frmEquipExam.frx":006C
         TabIndex        =   18
         Top             =   3900
         Visible         =   0   'False
         Width           =   2685
      End
      Begin VB.TextBox txtExamName 
         BeginProperty Font 
            Name            =   "새굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1410
         TabIndex        =   16
         Top             =   1950
         Width           =   2655
      End
      Begin VB.TextBox txtEquip 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "새굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1410
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   240
         Width           =   2655
      End
      Begin VB.TextBox txtRang 
         BeginProperty Font 
            Name            =   "새굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1410
         TabIndex        =   7
         Top             =   2835
         Width           =   2655
      End
      Begin VB.TextBox txtExamCode 
         BeginProperty Font 
            Name            =   "새굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1410
         TabIndex        =   5
         Top             =   1104
         Width           =   2655
      End
      Begin VB.TextBox txtEquipCode 
         BeginProperty Font 
            Name            =   "새굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1410
         TabIndex        =   3
         Top             =   672
         Width           =   2655
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "오더구분"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   270
         TabIndex        =   29
         Top             =   3810
         Width           =   1020
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검사종류"
         BeginProperty Font 
            Name            =   "새굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   270
         TabIndex        =   26
         Top             =   3390
         Width           =   960
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "OCS 코드"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   270
         TabIndex        =   25
         Top             =   4380
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "순    번"
         BeginProperty Font 
            Name            =   "새굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   270
         TabIndex        =   23
         Top             =   2460
         Width           =   840
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "결과구분"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   270
         TabIndex        =   21
         Top             =   3870
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "결과코드"
         BeginProperty Font 
            Name            =   "새굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   270
         TabIndex        =   20
         Top             =   1590
         Width           =   960
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검 사 명"
         BeginProperty Font 
            Name            =   "새굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   270
         TabIndex        =   17
         Top             =   2010
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "장 비 명"
         BeginProperty Font 
            Name            =   "새굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   270
         TabIndex        =   8
         Top             =   315
         Width           =   900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "소수자리"
         BeginProperty Font 
            Name            =   "새굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   270
         TabIndex        =   6
         Top             =   2910
         Width           =   960
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검사코드"
         BeginProperty Font 
            Name            =   "새굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   270
         TabIndex        =   4
         Top             =   1170
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "장비코드"
         BeginProperty Font 
            Name            =   "새굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   270
         TabIndex        =   2
         Top             =   750
         Width           =   960
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   2625
      Left            =   7590
      Picture         =   "frmEquipExam.frx":008B
      ScaleHeight     =   2565
      ScaleWidth      =   4185
      TabIndex        =   14
      Top             =   5760
      Visible         =   0   'False
      Width           =   4245
   End
   Begin FPSpread.vaSpread vasList 
      Height          =   7575
      Left            =   60
      TabIndex        =   0
      Top             =   840
      Width           =   7485
      _Version        =   393216
      _ExtentX        =   13203
      _ExtentY        =   13361
      _StockProps     =   64
      ColHeaderDisplay=   0
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   10
      ScrollBars      =   2
      SpreadDesigner  =   "frmEquipExam.frx":1A415
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   675
      Left            =   60
      TabIndex        =   15
      Top             =   60
      Width           =   11805
      _Version        =   65536
      _ExtentX        =   20823
      _ExtentY        =   1191
      _StockProps     =   15
      Caption         =   "  장비 코드 설정"
      ForeColor       =   8388608
      BackColor       =   16774393
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   14.26
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      Alignment       =   1
   End
End
Attribute VB_Name = "frmEquipExam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lsExamCode As String
Dim lsExamName As String
Dim lsGubun As String
Dim lsRang As String
Dim lsEquipFlag As String

Sub ClearText()
    lsExamCode = ""
    lsExamName = ""
    lsGubun = ""
    lsRang = ""
    lsEquipFlag = ""
    
    txtEquipCode = ""
    txtExamCode = ""
    txtExamName = ""
    txtRSCode = ""
    txtOcsCode = ""
    txtRang = ""
    txtSeq = ""
    ComPart = ""
    cboOrdGubun = ""
    
    cmdSave.Caption = "저장"
End Sub

Sub DisplayList()
    ClearSpread vasList
    
    '장비, 장비코드, 검사코드, 검사구분, OCS코드, 검사명, 순번, 소수자리
    SQL = "SELECT equipno, equipcode, examcode, rscode, '', examname, seqno, resprec,'',exampart, OrdGubun " & vbCrLf & _
          "  From EquipExam " & vbCrLf & _
          " WHERE equipno = '" & gEquip & "' " '& vbCrLf & _
          " Order By 1 "
    res = db_select_Vas(gLocal, SQL, vasList)
    
    vasSort vasList, 2
    
    vasList.MaxRows = vasList.DataRowCnt
End Sub

Function Get_ExamName(asExamCode As String, asKindCode As String, asOcsCode As String) As Integer
    Get_ExamName = -1
    
    If asExamCode = "" Then
        Exit Function
    End If
    
    SQL = " Use MDCK"
    res = SendQuery(gServer, SQL)
    
    SQL = " Select NameEng From Bag_InterfaceCode " & CR & _
          " Where MedItem = '" & Trim(asExamCode) & "' "
    
    If asKindCode <> "" Then
        SQL = SQL & CR & " And Kind = '" & Trim(asKindCode) & "' "
    End If
    
    If asOcsCode <> "" Then
        SQL = SQL & CR & " And OcsCode = '" & Trim(asOcsCode) & "' "
    End If
    
    res = db_select_Col(gServer, SQL)
    
    If res = 1 Then
        txtExamName = Trim(gReadBuf(0))
        
        Get_ExamName = 1
    End If
End Function

Function ExistOfEquipCode(asEquipCode As String, Optional asExamCode As String = "") As Integer
'장비코드와 검사코드에 해당하는 데이타 존재 확인 하는 procedure

    ExistOfEquipCode = -1
    
    If asEquipCode = "" Then
        Exit Function
    End If
    
    SQL = "SELECT EquipCode, ExamCode, ExamName " & vbCrLf & _
          "  From EquipExam " & vbCrLf & _
          " WHERE EquipNo = '" & gEquip & "' " & vbCrLf & _
          "   AND EquipCode = '" & asEquipCode & "' "
          
    If Trim(asExamCode) <> "" Then
        SQL = SQL & vbCrLf & _
          "   AND ExamCode = '" & asExamCode & "' "
    End If
    
    res = db_select_Col(gLocal, SQL)
    
    If res = 0 Then
        ExistOfEquipCode = 0
        Exit Function
    ElseIf res = -1 Then
        ExistOfEquipCode = -1
        Exit Function
    End If
    
    If Trim(gReadBuf(0)) <> asEquipCode Then
        Exit Function
    End If
    
    lsExamCode = Trim(gReadBuf(1))
    lsExamName = Trim(gReadBuf(2))
'    lsGubun = Trim(gReadBuf(4))
'    lsRang = Trim(gReadBuf(5))
'    lsEquipFlag = Trim(gReadBuf(6))
'    Select Case lsEquipFlag
'    Case "0"
'        optType(0).Value = True
'    Case "1"
'        optType(1).Value = True
'    Case "2"
'        optType(2).Value = True
'    End Select
    
    ExistOfEquipCode = 1
End Function

Private Sub cboGubun_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If cboGubun.ListIndex < 0 Then
            cboGubun.SetFocus
            Exit Sub
        End If
        
        txtRang.SetFocus
    End If
End Sub

Private Sub cmdCancel_Click()
    ClearText
    txtEquipCode.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    If Trim(txtEquipCode) = "" Then
        txtEquipCode.SetFocus
        Exit Sub
    End If
    
    
    If Trim(txtExamCode) = "" Then
        txtExamCode.SetFocus
'        Exit Sub
    End If

    SQL = "Delete from EquipExam " & vbCrLf & _
          "Where EquipNo = '" & gEquip & "' " & vbCrLf & _
          "  and EquipCode = '" & Trim(txtEquipCode) & "' " & vbCrLf & _
          "  and ExamCode = '" & Trim(txtExamCode) & "' "
          
    res = SendQuery(gLocal, SQL)
    If res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If

    DisplayList
    
    cmdCancel_Click

End Sub

Private Sub cmdSave_Click()
'수가코드(검사코드) 없어도 저장되도록
Dim i As Integer
    
    If Trim(txtEquipCode) = "" Then
        txtEquipCode.SetFocus
        Exit Sub
    End If
    
    If Trim(txtExamCode) = "" Then
        txtExamCode.SetFocus
    End If
    
    If Trim(txtExamName) = "" Then
        txtExamName.SetFocus
    End If
    
    If Trim(txtRang) = "" Then
        txtRang.Text = 0
    End If
    
    res = ExistOfEquipCode(Trim(txtEquipCode), Trim(txtExamCode))
    If res = 1 Then
        'update
        SQL = " Update EquipExam " & vbCrLf & _
              " Set rscode = '" & Trim(txtRSCode.Text) & "', " & vbCrLf & _
              "     ExamName = '" & Trim(txtExamName.Text) & "', " & vbCrLf & _
              "     resprec = " & Trim(txtRang.Text) & ", " & vbCrLf & _
              "     seqno = '" & Trim(txtSeq.Text) & "', " & vbCrLf & _
              "     exampart = '" & Trim(ComPart.Text) & "', " & vbCrLf & _
              "     OrdGubun = '" & Left(cboOrdGubun.Text, 1) & "' " & vbCrLf & _
              " Where equipno = '" & Trim(txtEquip.Text) & "' " & vbCrLf & _
              " And EquipCode = '" & Trim(txtEquipCode.Text) & "' " & vbCrLf & _
              " And ExamCode = '" & Trim(txtExamCode.Text) & "' "
    
    ElseIf res = 0 Then
        'insert
        SQL = " Insert Into EquipExam(equipno, EquipCode, ExamCode, rscode,  ExamName, resprec, seqno,exampart, OrdGubun ) " & vbCrLf & _
              " Values ('" & Trim(txtEquip.Text) & "', '" & Trim(txtEquipCode.Text) & "', '" & Trim(txtExamCode.Text) & "',  " & vbCrLf & _
              "         '" & Trim(txtRSCode.Text) & "', '" & Trim(txtExamName.Text) & "', " & Trim(txtRang.Text) & ", '" & Trim(txtSeq) & "', " & vbCrLf & _
              "         '" & Trim(ComPart.Text) & "','" & Left(cboOrdGubun.Text, 1) & "') "
    End If
    
    res = SendQuery(gLocal, SQL)
    If res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If
    
    DisplayList
    
    cmdCancel_Click
End Sub

Private Sub Form_Load()
'    Me.Height = 8600
'    Me.Width = 11970
            
    ClearText
    txtEquip = gEquip
    
    ComPart.AddItem "HEMATOLOGY"
'    ComPart.AddItem "URINALYSIS"
'    ComPart.AddItem "CHEMISTRY"
'    ComPart.AddItem "IMMUNO-SERO"
'    ComPart.AddItem "MICRO-BIO"
'    ComPart.AddItem "결핵"
'    ComPart.AddItem "Others"
    
    DisplayList
End Sub

Private Sub txtEquipCode_GotFocus()
    SelectFocus txtEquipCode
End Sub

Private Sub txtEquipCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txtEquipCode = "" Then
            txtEquipCode.SetFocus
            Exit Sub
        End If
        txtExamCode.SetFocus
    End If
End Sub

Private Sub txtExamCode_GotFocus()
    SelectFocus txtExamCode
End Sub

Private Sub txtExamCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txtExamCode = "" Then
            txtExamCode.SetFocus
            Exit Sub
        End If
        
        txtExamCode.Text = UCase(txtExamCode)
        txtRSCode.SetFocus
    End If
End Sub

Private Sub txtExamName_GotFocus()
    SelectFocus txtExamName
End Sub

Private Sub txtExamName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txtExamName = "" Then
            txtExamName.SetFocus
            Exit Sub
        End If
        
        'cmdSave.SetFocus
        txtSeq.SetFocus
    End If
End Sub


Private Sub txtOcsCode_GotFocus()
    SelectFocus txtOcsCode
End Sub

Private Sub txtOcsCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
'        If txtOcsCode = "" Then
'            txtOcsCode.SetFocus
'        Else
            txtOcsCode = txtOcsCode
            
            res = Get_ExamName(txtExamCode, txtRSCode, txtOcsCode)

            txtExamName.SetFocus
'        End If
    End If
End Sub

Private Sub txtRang_GotFocus()
    SelectFocus txtRang
End Sub

Private Sub txtRang_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
'        If txtRang = "" Then
'            Exit Sub
'        End If
        
        cmdSave.SetFocus
    End If
End Sub

Private Sub txtRsCode_GotFocus()
    SelectFocus txtRSCode
End Sub

Private Sub txtRsCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then

        gReadBuf(0) = ""
        SQL = "Select RS_ALIAS from RSLT_TCD where IN_CODE = '" & Trim(txtExamCode) & "' "
        res = db_select_Col(gServer, SQL)
        
        If gReadBuf(0) <> "" Then
            txtExamName = Trim(gReadBuf(0))
        End If
        
        txtExamName.SetFocus
    End If
End Sub

Private Sub txtSeq_GotFocus()
    SelectFocus txtSeq
End Sub

Private Sub txtSeq_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txtSeq = "" Then
            txtSeq.SetFocus
        Else
            txtRang.SetFocus
        End If
    End If
End Sub

Private Sub vasList_Click(ByVal Col As Long, ByVal Row As Long)
    Dim i As Integer
    
    If Row < 1 Or Row > vasList.DataRowCnt Then
        cmdSave.Caption = "저장"
        ClearText
        Exit Sub
    End If
    
    txtEquip = Trim(GetText(vasList, Row, 1))
    txtEquipCode = Trim(GetText(vasList, Row, 2))
    txtExamCode = Trim(GetText(vasList, Row, 3))
    txtRSCode = Trim(GetText(vasList, Row, 4))
    txtOcsCode = Trim(GetText(vasList, Row, 5))
    txtExamName = Trim(GetText(vasList, Row, 6))
    
    txtSeq = Trim(GetText(vasList, Row, 7))
    
    txtRang = Trim(GetText(vasList, Row, 8))
    ComPart.Text = Trim(GetText(vasList, Row, 10))
    
    cboOrdGubun.Text = ""
    For i = 0 To cboOrdGubun.ListCount - 1
        If Left(cboOrdGubun.List(i), 1) = Trim(GetText(vasList, Row, 11)) Then
            cboOrdGubun.ListIndex = i
            Exit For
        End If
    Next i
    
    cmdSave.Caption = "수정"
End Sub
