VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmEquipExam 
   Caption         =   "검사 코드 설정"
   ClientHeight    =   8505
   ClientLeft      =   2385
   ClientTop       =   2055
   ClientWidth     =   11925
   Icon            =   "frmEquipExam.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8505
   ScaleWidth      =   11925
   Begin VB.Frame Frame1 
      Height          =   5325
      Left            =   7680
      TabIndex        =   14
      Top             =   840
      Width           =   4215
      Begin VB.CheckBox chkOrder 
         Height          =   285
         Left            =   1440
         TabIndex        =   29
         Top             =   3285
         Width           =   375
      End
      Begin VB.TextBox txtPaniclow 
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
         Left            =   1560
         TabIndex        =   5
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox txtPanichigh 
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
         Left            =   1560
         TabIndex        =   4
         Top             =   2340
         Width           =   1095
      End
      Begin VB.TextBox txtSeq 
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
         Left            =   1440
         TabIndex        =   3
         Top             =   1920
         Width           =   2655
      End
      Begin VB.TextBox txtRSCode 
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
         TabIndex        =   6
         Top             =   3615
         Visible         =   0   'False
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
         ItemData        =   "frmEquipExam.frx":08CA
         Left            =   1410
         List            =   "frmEquipExam.frx":08D7
         TabIndex        =   7
         Top             =   3900
         Visible         =   0   'False
         Width           =   2685
      End
      Begin VB.TextBox txtExamName 
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
         Left            =   1440
         TabIndex        =   2
         Top             =   1485
         Width           =   2655
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "종료"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   3150
         TabIndex        =   12
         Top             =   4560
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   2160
         TabIndex        =   11
         Top             =   4560
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "삭제"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   1170
         TabIndex        =   10
         Top             =   4560
         Width           =   975
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "저장"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   180
         TabIndex        =   9
         Top             =   4560
         Width           =   975
      End
      Begin VB.TextBox txtEquip 
         Appearance      =   0  '평면
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
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   180
         Width           =   2655
      End
      Begin VB.TextBox txtRang 
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
         TabIndex        =   8
         Top             =   4125
         Width           =   2655
      End
      Begin VB.TextBox txtExamCode 
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
         Left            =   1440
         TabIndex        =   1
         Top             =   1080
         Width           =   2655
      End
      Begin VB.TextBox txtEquipCode 
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
         Left            =   1440
         TabIndex        =   0
         Top             =   660
         Width           =   2655
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "오더 전송"
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
         Left            =   225
         TabIndex        =   28
         Top             =   3285
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "이하"
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
         Index           =   3
         Left            =   2760
         TabIndex        =   27
         Top             =   2820
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "이상"
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
         Index           =   2
         Left            =   2760
         TabIndex        =   26
         Top             =   2400
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Panic ref"
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
         Index           =   1
         Left            =   240
         TabIndex        =   25
         Top             =   2460
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "순    번"
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
         Left            =   240
         TabIndex        =   24
         Top             =   2040
         Width           =   1050
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
         TabIndex        =   23
         Top             =   3960
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "항목코드"
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
         TabIndex        =   22
         Top             =   3690
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검 사 명"
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
         Left            =   180
         TabIndex        =   21
         Top             =   1590
         Width           =   1035
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "장 비 명"
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
         Left            =   180
         TabIndex        =   18
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "소수자리"
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
         TabIndex        =   17
         Top             =   4200
         Width           =   1020
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검사코드"
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
         TabIndex        =   16
         Top             =   1170
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "장비코드"
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
         Index           =   0
         Left            =   180
         TabIndex        =   15
         Top             =   720
         Width           =   1020
      End
   End
   Begin FPSpread.vaSpread vasList 
      Height          =   7515
      Left            =   60
      TabIndex        =   13
      Top             =   840
      Width           =   7545
      _Version        =   393216
      _ExtentX        =   13309
      _ExtentY        =   13256
      _StockProps     =   64
      ColHeaderDisplay=   0
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   11
      ScrollBars      =   2
      SpreadDesigner  =   "frmEquipExam.frx":08F6
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   675
      Left            =   6240
      TabIndex        =   20
      Top             =   6660
      Visible         =   0   'False
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
   Begin VB.Shape Shape2 
      BackStyle       =   1  '투명하지 않음
      FillColor       =   &H00FF0000&
      FillStyle       =   0  '단색
      Height          =   435
      Index           =   1
      Left            =   4290
      Top             =   240
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  '투명하지 않음
      FillColor       =   &H000000FF&
      FillStyle       =   0  '단색
      Height          =   435
      Index           =   0
      Left            =   60
      Top             =   60
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "검사 코드 설정"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Index           =   9
      Left            =   420
      TabIndex        =   30
      Top             =   150
      Width           =   1980
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      FillColor       =   &H00FF8080&
      FillStyle       =   0  '단색
      Height          =   615
      Index           =   1
      Left            =   60
      Shape           =   4  '둥근 사각형
      Top             =   60
      Width           =   4365
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
    txtRang = ""
    txtSeq = ""
    
    txtPaniclow = ""
    txtPanichigh = ""
    chkOrder.Value = 0
    cmdSave.Caption = "저장"
End Sub

Sub DisplayList()
    ClearSpread vasList
    
    SQL = "SELECT EQUIPNO, EQUIPCODE, ExamCode, ExamName, SEQNO, '', '', '', paniclow, panichigh, ORDERENABLE " & vbCrLf & _
          "  From EquipExam " & vbCrLf & _
          " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf & _
          " ORDER BY SEQNO "
    res = db_SELECT_Vas(gLocal, SQL, vasList)
    
    vasList.MaxRows = vasList.DataRowCnt
End Sub

Function ExistOfEQUIPCODE(asEQUIPCODE As String, Optional asExamCode As String = "") As Integer
'장비코드와 검사코드에 해당하는 데이타 존재 확인 하는 procedure

    ExistOfEQUIPCODE = -1
    
    If asEQUIPCODE = "" Then
        Exit Function
    End If
    
    SQL = "SELECT EQUIPCODE, ExamCode, ExamName " & vbCrLf & _
          "  From EquipExam " & vbCrLf & _
          " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf & _
          "   AND EQUIPCODE = '" & asEQUIPCODE & "' "
          
    If Trim(asExamCode) <> "" Then
        SQL = SQL & vbCrLf & _
          "   AND ExamCode = '" & asExamCode & "' "
    End If
    
    res = db_SELECT_Col(gLocal, SQL)
    
    If res = 0 Then
        ExistOfEQUIPCODE = 0
        Exit Function
    ElseIf res = -1 Then
        ExistOfEQUIPCODE = -1
        Exit Function
    End If
    
    If Trim(gReadBuf(0)) <> asEQUIPCODE Then
        Exit Function
    End If
    
    lsExamCode = Trim(gReadBuf(1))
    lsExamName = Trim(gReadBuf(2))
'    lsGubun = Trim(gReadBuf(4))
'    lsRang = Trim(gReadBuf(5))
'    lsEquipFlag = Trim(gReadBuf(6))
'    SELECT Case lsEquipFlag
'    Case "0"
'        optType(0).Value = True
'    Case "1"
'        optType(1).Value = True
'    Case "2"
'        optType(2).Value = True
'    End SELECT
    
    ExistOfEQUIPCODE = 1
End Function

Function GetExamName(argExamCode As String) As Integer
'검사명 불러오기
    GetExamName = -1
    
    If argExamCode = "" Then
        Exit Function
    End If
    
    gReadBuf(0) = ""
    SQL = " SELECT ExamAlias From ExamMaster " & CR & _
          " WHERE HID = '117' " & CR & _
          " And ExamCode = '" & Trim(argExamCode) & "' "
    res = db_SELECT_Col(gServer, SQL)
    
    If gReadBuf(0) <> "" Then
        GetExamName = 1
    End If
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
        
    'db_BeginTran gServer
    
    SQL = "Delete from EquipExam " & vbCrLf & _
          "WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf & _
          "  and EQUIPCODE = '" & Trim(txtEquipCode) & "' " & vbCrLf & _
          "  and ExamCode = '" & Trim(txtExamCode) & "' "
          
    res = SendQuery(gLocal, SQL)
    If res = -1 Then
        db_RollBack gServer
        Exit Sub
    End If
    
    'db_Commit gLocal

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
'        Exit Sub
    End If
    
'    If Trim(txtRSCode) = "" Then
'        txtRSCode = "0"
'    End If
    
    If Trim(txtExamName) = "" Then
        txtExamName.SetFocus
    End If
    
'    If Trim(txtRang) = "" Then
'        txtRang.Text = 0
'    End If
    
    If Trim(txtSeq) = "" Then
        txtSeq = "99"
    End If
    
    IsolateCode cboGubun
    lsGubun = gCode

    
    'db_BeginTran gServer
    
    res = ExistOfEQUIPCODE(Trim(txtEquipCode), Trim(txtExamCode))
    If res = 1 Then
        'UPDATE
        SQL = " UPDATE EquipExam " & vbCrLf & _
              " Set ExamCode = '" & Trim(txtExamCode.Text) & "', " & vbCrLf & _
              "     ExamName = '" & Trim(txtExamName.Text) & "', " & vbCrLf & _
              "     SEQNO = " & Trim(txtSeq.Text) & ", " & vbCrLf & _
              "     paniclow = '" & Trim(txtPaniclow) & "', " & vbCrLf & _
              "     panichigh = '" & Trim(txtPanichigh) & "', " & vbCrLf & _
              "     ORDERENABLE = '" & Trim(chkOrder.Value) & "' " & vbCrLf & _
              " WHERE EQUIPNO = '" & Trim(txtEquip.Text) & "' " & vbCrLf & _
              " And EQUIPCODE = '" & Trim(txtEquipCode.Text) & "' " '& vbCrLf & _
              " And ExamCode = '" & Trim(txtExamCode.Text) & "' "
    
    ElseIf res = 0 Then
        'insert
        SQL = " Insert Into EquipExam(EQUIPNO, EQUIPCODE, ExamCode, ExamName, SEQNO, paniclow, panichigh, ORDERENABLE) " & vbCrLf & _
              " Values ('" & Trim(txtEquip.Text) & "', '" & Trim(txtEquipCode.Text) & "', '" & Trim(txtExamCode.Text) & "',  " & vbCrLf & _
              "         '" & Trim(txtExamName.Text) & "', " & Trim(txtSeq.Text) & ", '" & Trim(txtPaniclow) & "', '" & Trim(txtPanichigh) & "', '" & Trim(chkOrder.Value) & "') "
    End If
    
    res = SendQuery(gLocal, SQL)
    If res = -1 Then
        'db_RollBack gServer
        SaveQuery SQL
        Exit Sub
    End If
    
    'db_Commit gServer
    
    DisplayList
    
    cmdCancel_Click
End Sub

Private Sub Form_Load()
'    Me.Height = 8600
'    Me.Width = 11970
            
    ClearText
    txtEquip = gEquip
    
    DisplayList
    
End Sub

Private Sub txtEQUIPCODE_GotFocus()
    SELECTFocus txtEquipCode
End Sub

Private Sub txtEQUIPCODE_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtExamCode_GotFocus()
    SELECTFocus txtExamCode
End Sub

Private Sub txtExamCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtExamCode.Text = UCase(txtExamCode)
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtExamName_GotFocus()
    SELECTFocus txtExamName
End Sub

Private Sub txtExamName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtRang_GotFocus()
    SELECTFocus txtRang
End Sub

Private Sub txtRang_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txtRang = "" Then
            Exit Sub
        End If
        
        cmdSave.SetFocus
    End If
End Sub

Private Sub txtRsCode_GotFocus()
    SELECTFocus txtRSCode
End Sub

Private Sub txtRsCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txtRSCode = "" Then
            Exit Sub

        Else
            '검사항목코드로 검사명, 결과값구분, 숫자정확도 불러오기
            SQL = " SELECT MH121_NAME, MH121_GU, MH121_CORRECT From MH121_CNT WHERE MH121_CODE = '" & Trim(txtExamCode.Text) & "' " & vbCrLf & _
                  " And MH121_CD = '" & Trim(txtRSCode.Text) & "'"
                  
            res = db_SELECT_Col(gServer, SQL)
            
            If res = 1 Then
                txtExamName.Text = Trim(gReadBuf(0))
                
                cboGubun.Text = Trim(gReadBuf(1))
                 
                Select Case Trim(gReadBuf(1))
                Case "0"
                    cboGubun.ListIndex = 0
                Case "1"
                    cboGubun.ListIndex = 1
                Case "2"
                    cboGubun.ListIndex = 2
                End Select
                
                txtRang.Text = Trim(gReadBuf(2))
                
                txtExamName.SetFocus
            Else
                txtRSCode.SetFocus
            End If
        End If
    End If
    
End Sub

Private Sub txtSeq_GotFocus()
    SELECTFocus txtSeq
End Sub

Private Sub txtSeq_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub vasList_Click(ByVal Col As Long, ByVal Row As Long)
    If Row < 1 Or Row > vasList.DataRowCnt Then
        cmdSave.Caption = "저장"
        ClearText
        Exit Sub
    End If
    
    txtEquip = Trim(GetText(vasList, Row, 1))
    txtEquipCode = Trim(GetText(vasList, Row, 2))
    txtExamCode = Trim(GetText(vasList, Row, 3))
    txtExamName = Trim(GetText(vasList, Row, 4))
    txtSeq = Trim(GetText(vasList, Row, 5))
    
    txtPaniclow = Trim(GetText(vasList, Row, 9))
    txtPanichigh = Trim(GetText(vasList, Row, 10))
    
    If Trim(GetText(vasList, Row, 11)) <> "" Then
        chkOrder.Value = Trim(CInt(GetText(vasList, Row, 11)))
    Else
        chkOrder.Value = 0
    End If
    
    cmdSave.Caption = "수정"
End Sub
