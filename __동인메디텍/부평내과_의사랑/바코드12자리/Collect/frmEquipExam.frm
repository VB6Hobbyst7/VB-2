VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmEquipExam 
   Caption         =   "장비 검사코드 설정"
   ClientHeight    =   7275
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11775
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7275
   ScaleWidth      =   11775
   StartUpPosition =   2  '화면 가운데
   Begin VB.Frame Frame1 
      Height          =   5265
      Left            =   8010
      TabIndex        =   2
      Top             =   210
      Width           =   3525
      Begin VB.ComboBox cboPart 
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
         ItemData        =   "frmEquipExam.frx":0000
         Left            =   1110
         List            =   "frmEquipExam.frx":0010
         TabIndex        =   25
         Top             =   690
         Width           =   2145
      End
      Begin VB.TextBox txtGbn 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   750
         TabIndex        =   23
         Top             =   3810
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.TextBox txtEquipCode 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1110
         TabIndex        =   14
         Top             =   1125
         Width           =   2115
      End
      Begin VB.TextBox txtCode 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1110
         TabIndex        =   13
         Top             =   1560
         Width           =   2115
      End
      Begin VB.TextBox txtDec 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1110
         TabIndex        =   12
         Top             =   2400
         Width           =   2115
      End
      Begin VB.TextBox txtName 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1110
         TabIndex        =   11
         Top             =   1980
         Width           =   2115
      End
      Begin VB.TextBox txtMuch 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1110
         TabIndex        =   10
         Top             =   300
         Width           =   2115
      End
      Begin VB.TextBox txtSeq 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1110
         TabIndex        =   9
         Top             =   2820
         Width           =   585
      End
      Begin VB.PictureBox picEquip 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BorderStyle     =   0  '없음
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2820
         Picture         =   "frmEquipExam.frx":003E
         ScaleHeight     =   330
         ScaleWidth      =   330
         TabIndex        =   8
         Top             =   1170
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   495
         Left            =   150
         TabIndex        =   7
         Top             =   4620
         Width           =   1035
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   495
         Left            =   1230
         TabIndex        =   6
         Top             =   4620
         Width           =   1035
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Clear"
         Height          =   495
         Left            =   2310
         TabIndex        =   5
         Top             =   4620
         Width           =   1035
      End
      Begin VB.TextBox txtRefLow 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1110
         TabIndex        =   4
         Top             =   3240
         Width           =   585
      End
      Begin VB.TextBox txtRefHigh 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2130
         TabIndex        =   3
         Top             =   3240
         Width           =   555
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "장비구분"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   240
         TabIndex        =   24
         Top             =   795
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "장비채널"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   240
         TabIndex        =   22
         Top             =   1200
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검사코드"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   240
         TabIndex        =   21
         Top             =   1620
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "소 수 점"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   240
         TabIndex        =   20
         Top             =   2475
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검 사 명"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   240
         TabIndex        =   19
         Top             =   2055
         Width           =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "장 비 명"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   240
         TabIndex        =   18
         Top             =   375
         Width           =   720
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "순    서"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   240
         TabIndex        =   17
         Top             =   2910
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "참 고 치"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   270
         TabIndex        =   16
         Top             =   3330
         Width           =   720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "-"
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
         Left            =   1860
         TabIndex        =   15
         Top             =   3240
         Width           =   135
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "닫기"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10380
      TabIndex        =   0
      Top             =   6600
      Width           =   1125
   End
   Begin FPSpread.vaSpread vasList 
      Height          =   6945
      Left            =   120
      TabIndex        =   1
      Top             =   180
      Width           =   7755
      _Version        =   393216
      _ExtentX        =   13679
      _ExtentY        =   12250
      _StockProps     =   64
      BackColorStyle  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   9
      MaxRows         =   20
      OperationMode   =   2
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SelectBlockOptions=   0
      SpreadDesigner  =   "frmEquipExam.frx":0188
   End
End
Attribute VB_Name = "frmEquipExam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ClearText()

    txtEquipCode = ""
    txtCode = ""
    txtName = ""
    txtDec = "1"
    txtSeq = ""
    txtRefLow = ""
    txtRefHigh = ""
    cmdSave.Caption = "Save"
    
End Sub

Private Sub DisplayList()

    ClearSpread vasList

    SQL = "SELECT EQUIPNO, EXAMTYPE, EQUIPCODE, EXAMCODE, EXAMNAME, RESPREC, SEQNO, REFLOW, REFHIGH " & vbCrLf & _
          "  From EQUIPEXAM " & vbCrLf & _
          " GROUP BY EQUIPNO, EXAMTYPE, EXAMCODE, EQUIPCODE, EXAMNAME, RESPREC, SEQNO, REFLOW, REFHIGH "
          
'          " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf & _

    res = GetDBSelectVas(gLocal, SQL, vasList)
    
    vasList.MaxRows = vasList.DataRowCnt
    vasList.RowHeight(-1) = 12
    Call vasList_Click(1, 0)
    
End Sub

'-- 장비코드와 수가코드에 해당하는 데이타 존재 확인 하는 procedure
Function ExistOfEquipCode(asEquipCode As String, Optional asSuga As String = "") As Integer

    ExistOfEquipCode = -1
    
    If asEquipCode = "" Then
        Exit Function
    End If
    
    SQL = "SELECT EQUIPCODE, EXAMCODE, EXAMNAME, RESPREC, SEQNO, REFLOW, REFHIGH " & vbCrLf & _
          "  FROM EQUIPEXAM " & vbCrLf & _
          " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf & _
          "   AND EQUIPCODE = '" & asEquipCode & "' "
          
    If Trim(asSuga) <> "" Then
        SQL = SQL & vbCrLf & _
          "   AND EXAMCODE = '" & asSuga & "' "
    End If
    
    res = GetDBSelectColumn(gLocal, SQL)
    If res = 0 Then
        ExistOfEquipCode = 0
        Exit Function
    ElseIf res = -1 Then
        ExistOfEquipCode = -1
        Exit Function
    End If
    
    If Trim(gReadBuf(0)) <> asEquipCode Or Trim(gReadBuf(1)) <> asSuga Then
        Exit Function
    End If
        
    ExistOfEquipCode = 1
End Function

'-- 장비코드와 수가코드에 해당하는 데이타 존재 확인 하는 procedure
Function ExistOfEquipCode_New(asEquipNo As String, asEquipCode As String, Optional asSuga As String = "") As Integer

    ExistOfEquipCode_New = -1
    
    If asEquipCode = "" Then
        Exit Function
    End If
    
    SQL = "SELECT EQUIPCODE, EXAMCODE, EXAMNAME, RESPREC, SEQNO, REFLOW, REFHIGH " & vbCrLf & _
          "  FROM EQUIPEXAM " & vbCrLf & _
          " WHERE EQUIPNO = '" & asEquipNo & "' " & vbCrLf & _
          "   AND EQUIPCODE = '" & asEquipCode & "' "
          
    If Trim(asSuga) <> "" Then
        SQL = SQL & vbCrLf & _
          "   AND EXAMCODE = '" & asSuga & "' "
    End If
    
    res = GetDBSelectColumn(gLocal, SQL)
    If res = 0 Then
        ExistOfEquipCode_New = 0
        Exit Function
    ElseIf res = -1 Then
        ExistOfEquipCode_New = -1
        Exit Function
    End If
    
    If Trim(gReadBuf(0)) <> asEquipCode Or Trim(gReadBuf(1)) <> asSuga Then
        Exit Function
    End If
        
    ExistOfEquipCode_New = 1
    
End Function

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
    
    SQL = "DELETE FROM EQUIPEXAM " & vbCrLf & _
          "WHERE EQUIPNO = '" & Trim(txtMuch) & "' " & vbCrLf & _
          "  AND EQUIPCODE = '" & Trim(txtEquipCode) & "' " & vbCrLf & _
          "  AND EXAMCODE = '" & Trim(txtCode) & "' "
    res = SendQuery(gLocal, SQL)
    If res = -1 Then
        Exit Sub
    End If
    
    DisplayList
    
    cmdCancel_Click

End Sub

Private Sub cmdSave_Click()
    Dim lsFlag As String
    Dim lsResFlag As String
    Dim liSeqNo As Integer

    If Trim(txtEquipCode) = "" Then
        txtEquipCode.SetFocus
        MsgBox "장비코드를 입력하세요", vbInformation
        Exit Sub
    End If
    
    If Trim(txtDec) = "" Then
        txtDec.Text = 1
    End If
    
    If IsNumeric(txtSeq) Then
        liSeqNo = CInt(txtSeq)
    Else
        liSeqNo = 0
    End If
    
    res = ExistOfEquipCode_New(Trim(txtMuch), Trim(txtEquipCode), Trim(txtCode))
    If res = 1 Then
        SQL = "UPDATE EQUIPEXAM " & vbCrLf & _
              "SET RESPREC = '" & Trim(txtDec) & "', " & vbCrLf & _
              "    EXAMNAME = '" & Trim(txtName) & "', " & vbCrLf & _
              "    REFLOW = '" & Trim(txtRefLow) & "', " & vbCrLf & _
              "    REFHIGH = '" & Trim(txtRefHigh) & "', " & vbCrLf & _
              "    SEQNO = " & liSeqNo & ", " & vbCrLf & _
              "    EXAMTYPE = '" & Mid(Trim(cboPart.Text), 1, 1) & "' " & vbCrLf & _
              "WHERE EQUIPNO = '" & Trim(txtMuch) & "' " & vbCrLf & _
              "  AND EQUIPCODE = '" & Trim(txtEquipCode) & "' " & vbCrLf & _
              "  AND EXAMCODE = '" & Trim(txtCode) & "' "
    ElseIf res = 0 Then
        SQL = "INSERT INTO EQUIPEXAM (EQUIPNO,EQUIPCODE, EXAMCODE, EXAMNAME, RESPREC, SEQNO , REFLOW, REFHIGH,EXAMTYPE) " & vbCrLf & _
              "VALUES ('" & Trim(txtMuch) & "', '" & Trim(txtEquipCode) & "', '" & Trim(txtCode) & "', '" & Trim(txtName.Text) & "', '" & Trim(txtDec) & "', " & liSeqNo & ", '" & Trim(txtRefLow) & "', '" & Trim(txtRefHigh) & "','" & Mid(Trim(cboPart.Text), 1, 1) & "') "
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
'    Me.Height = 7725
'    Me.Width = 9945
            
    ClearText
    DisplayList

'    txtMuch = gEquip
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
        txtCode.SetFocus
    End If
End Sub

Private Sub txtDec_GotFocus()
    SelectFocus txtDec
End Sub

Private Sub txtDec_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txtDec = "" Then
            txtDec.SetFocus
'            Exit Sub
        End If
        
        txtRefLow.SetFocus
    End If
End Sub

Private Sub txtcode_GotFocus()
    SelectFocus txtCode
End Sub

Private Sub txtcode_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        res = ExistOfEquipCode(Trim(txtEquipCode), Trim(txtCode))
        If res = -1 Then
            txtCode.SetFocus
            Exit Sub
        ElseIf res = 0 Then
            cmdSave.Caption = "Save"
            
        ElseIf res = 1 Then
            cmdSave.Caption = "Edit"
            txtName = Trim(gReadBuf(2))
            txtDec = Trim(gReadBuf(3))
            txtSeq = Trim(gReadBuf(4))
            txtRefLow = Trim(gReadBuf(5))
            txtRefHigh = Trim(gReadBuf(6))
        End If
        
        txtName.SetFocus
    End If
    
End Sub

Private Sub txtMuch_GotFocus()

    SelectFocus txtMuch
    
End Sub

Private Sub txtMuch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Trim(txtMuch.Text) = "" Then
            txtMuch.SetFocus
            Exit Sub
        End If
        txtEquipCode.SetFocus
    End If
End Sub

Private Sub txtName_GotFocus()
    SelectFocus txtName
End Sub

Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Trim(txtName.Text) = "" Then
            txtName.SetFocus
            Exit Sub
        End If
        txtDec.SetFocus
        
    End If
End Sub

Private Sub txtSeq_GotFocus()
    SelectFocus txtSeq
End Sub

Private Sub txtSeq_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Trim(txtSeq.Text) = "" Then
            txtSeq.SetFocus
            Exit Sub
        End If

        cmdSave.SetFocus
    End If
End Sub

Private Sub vasList_Click(ByVal Col As Long, ByVal Row As Long)
    If Row = 0 Then
        Select Case Col
        Case 1
            vasSort vasList, 1, 2
        Case 2
            vasSort vasList, 2, 1
        Case 5
            vasSort vasList, 5, 1
        End Select
        Exit Sub
    End If
    
    If Row < 1 Or Row > vasList.DataRowCnt Then
        cmdSave.Caption = "Save"
        ClearText
        Exit Sub
    End If
    txtMuch = Trim(GetText(vasList, Row, 1))
'    txtGbn = Trim(GetText(vasList, Row, 2))
    cboPart = Trim(GetText(vasList, Row, 2))
    Select Case Trim(GetText(vasList, Row, 2))
    Case "C": cboPart = "Chemistry"
    Case "H": cboPart = "Hematology"
    Case "I": cboPart = "Immunology"
    Case "U": cboPart = "Urine"
    End Select
    
    txtEquipCode = Trim(GetText(vasList, Row, 3))
    txtCode = Trim(GetText(vasList, Row, 4))
    txtName = Trim(GetText(vasList, Row, 5))
    txtDec = Trim(GetText(vasList, Row, 6))
    txtSeq = Trim(GetText(vasList, Row, 7))
    txtRefLow = Trim(GetText(vasList, Row, 8))
    txtRefHigh = Trim(GetText(vasList, Row, 9))

    
    
    cmdSave.Caption = "Edit"
End Sub


