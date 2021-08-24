VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmOrderCode 
   Caption         =   "장비 코드 설정"
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12270
   LinkTopic       =   "Form1"
   ScaleHeight     =   8400
   ScaleWidth      =   12270
   StartUpPosition =   2  '화면 가운데
   Begin FPSpread.vaSpread vasList 
      Height          =   7575
      Left            =   120
      TabIndex        =   25
      Top             =   780
      Width           =   8415
      _Version        =   393216
      _ExtentX        =   14843
      _ExtentY        =   13361
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
      MaxCols         =   11
      MaxRows         =   20
      OperationMode   =   2
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SelectBlockOptions=   0
      SpreadDesigner  =   "frmOrderCode.frx":0000
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   120
      ScaleHeight     =   525
      ScaleWidth      =   12045
      TabIndex        =   23
      Top             =   120
      Width           =   12075
      Begin Threed.SSPanel SSPanel1 
         Height          =   585
         Left            =   -120
         TabIndex        =   24
         Top             =   0
         Width           =   12165
         _Version        =   65536
         _ExtentX        =   21458
         _ExtentY        =   1032
         _StockProps     =   15
         Caption         =   "   URIT 8021A 장비 코드 설정"
         ForeColor       =   8388608
         BackColor       =   16056319
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         FloodColor      =   12582912
         Alignment       =   1
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5565
      Left            =   8640
      TabIndex        =   15
      Top             =   720
      Width           =   3555
      Begin VB.ComboBox cmbResType 
         Height          =   300
         ItemData        =   "frmOrderCode.frx":0A25
         Left            =   1110
         List            =   "frmOrderCode.frx":0A2F
         TabIndex        =   6
         Top             =   2700
         Width           =   2115
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
         Height          =   330
         Left            =   2280
         TabIndex        =   9
         Top             =   3420
         Width           =   795
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
         Height          =   330
         Left            =   1110
         TabIndex        =   8
         Top             =   3420
         Width           =   795
      End
      Begin VB.ComboBox cmbGubun 
         Height          =   300
         ItemData        =   "frmOrderCode.frx":0A45
         Left            =   1110
         List            =   "frmOrderCode.frx":0A47
         TabIndex        =   2
         Top             =   1140
         Width           =   2115
      End
      Begin VB.TextBox txtSubCode 
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
         TabIndex        =   4
         Top             =   1920
         Width           =   2115
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "종료"
         Height          =   495
         Left            =   2580
         TabIndex        =   14
         Top             =   4860
         Width           =   795
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Clear"
         Height          =   495
         Left            =   1800
         TabIndex        =   13
         Top             =   4860
         Width           =   795
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "삭제"
         Height          =   495
         Left            =   1020
         TabIndex        =   12
         Top             =   4860
         Width           =   795
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "저장"
         Height          =   495
         Left            =   240
         TabIndex        =   11
         Top             =   4860
         Width           =   795
      End
      Begin VB.PictureBox picEquip 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BorderStyle     =   0  '없음
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2880
         Picture         =   "frmOrderCode.frx":0A49
         ScaleHeight     =   330
         ScaleWidth      =   330
         TabIndex        =   22
         Top             =   1500
         Visible         =   0   'False
         Width           =   330
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
         TabIndex        =   7
         Top             =   3030
         Width           =   585
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
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   300
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
         TabIndex        =   5
         Top             =   2310
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
         TabIndex        =   10
         Top             =   3810
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
         TabIndex        =   3
         Top             =   1530
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
         TabIndex        =   1
         Top             =   735
         Width           =   2115
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "결과형태"
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
         TabIndex        =   30
         Top             =   2760
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
         Left            =   2040
         TabIndex        =   29
         Top             =   3480
         Width           =   135
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "참 조 치"
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
         TabIndex        =   28
         Top             =   3480
         Width           =   720
      End
      Begin VB.Label lblGubun 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검사구분"
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
         TabIndex        =   27
         Top             =   1200
         Width           =   720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "서브코드"
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
         TabIndex        =   26
         Top             =   1980
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
         TabIndex        =   21
         Top             =   3120
         Width           =   720
      End
      Begin VB.Label Label6 
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
         TabIndex        =   20
         Top             =   375
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
         Top             =   2385
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
         TabIndex        =   18
         Top             =   3885
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
         TabIndex        =   17
         Top             =   1590
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "장비코드"
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
         TabIndex        =   16
         Top             =   810
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmOrderCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub ClearText()
'화면초기화
    txtEquipCode = ""
    txtCode = ""
    txtName = ""
    txtDec = "1"
    txtSeq = ""
    txtRefLow = ""
    txtRefHigh = ""
    txtSubCode = ""
    cmbGubun.Text = ""
    cmbResType.Text = ""
    
    cmdSave.Caption = "저장"
End Sub

Sub DisplayList()
Dim i As Integer
'검사항목 조회
    ClearSpread vasList
'equipcode, examgubun, examcode, subcode, examname, resgubun, range, seqno, reflow, refhigh
    SQL = "SELECT equipcode, examgubun, examcode, subcode, examname, resgubun, range, seqno, reflow, refhigh " & CR & _
          "  From equipexam " & CR & _
          " WHERE equipno = '" & gEquip & "' " & CR & _
          " group by equipcode, examgubun, examcode, subcode, examname, resgubun, range, seqno, reflow, refhigh " & CR & _
          " order by seqno"
          
    res = db_select_Vas(gLocal, SQL, vasList)
    
    For i = 1 To vasList.DataRowCnt
        Select Case Trim(GetText(vasList, i, 2))
        Case "1"
        SetText vasList, "검진", i, 2
        Case "2"
        SetText vasList, "진료", i, 2
        End Select
    Next i
    
    vasList.MaxRows = vasList.DataRowCnt
    vasList.RowHeight(-1) = 13
End Sub

Function ExistOfEquipCode(asEquipCode As String, Optional asSuga As String = "") As Integer
'장비코드와 수가코드에 해당하는 데이타 존재 확인 하는 procedure

    ExistOfEquipCode = -1
    
    If asEquipCode = "" Then
        Exit Function
    End If
    
    SQL = "SELECT equipcode, examcode, subcode, examname, range, seqno, reflow, refhigh, resgubun, examgubun " & CR & _
          "  From equipexam " & CR & _
          " WHERE equipno = '" & gEquip & "' " & CR & _
          "   AND equipcode = '" & asEquipCode & "' "
    If Trim(asSuga) <> "" Then
        SQL = SQL & CR & _
          "   AND examcode = '" & asSuga & "' "
    End If
    res = db_select_Col(gLocal, SQL)
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
    
'    If Trim(txtCode) = "" Then
'        txtCode.SetFocus
'        Exit Sub
'    End If
        
'    db_BeginTran gLocal
    
    SQL = "Delete From equipexam " & CR & _
          "Where equipno = '" & gEquip & "' " & CR & _
          "  and equipcode = '" & Trim(txtEquipCode) & "' " & CR & _
          "  and examcode = '" & Trim(txtCode) & "' and subcode = '" & Trim(txtSubCode) & "' "
    res = SendQuery(gLocal, SQL)
    If res = -1 Then
'        db_RollBack gLocal
        Exit Sub
    End If
    
'    db_Commit gLocal

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
    
'    db_BeginTran gLocal
    'equipno, equipcode, examcode, examname, resprec, seqno, reflow, refhigh
    res = ExistOfEquipCode(Trim(txtEquipCode), Trim(txtCode))
    If res = 1 Then
        SQL = "Update equipexam " & CR & _
              "Set subcode = '" & Trim(txtSubCode) & "', " & vbCrLf & _
              "    examname = '" & Trim(txtName) & "', " & vbCrLf & _
              "    resgubun = '" & Left(Trim(cmbResType.Text), 1) & "', " & vbCrLf & _
              "    EXAMGUBUN = '" & Left(Trim(cmbGubun.Text), 1) & "', " & vbCrLf & _
              "    seqno = '" & liSeqNo & "', " & vbCrLf & _
              "    reflow = '" & Trim(txtRefLow) & "', " & vbCrLf & _
              "    refhigh = '" & Trim(txtRefHigh) & "', " & vbCrLf & _
              "    range = " & Trim(txtDec) & " " & vbCrLf & _
              "Where equipno = '" & gEquip & "' " & vbCrLf & _
              "  and equipcode = '" & Trim(txtEquipCode) & "' " & vbCrLf & _
              "  and examcode = '" & Trim(txtCode) & "' "
    ElseIf res = 0 Then
        SQL = "Insert Into equipexam (equipno, equipcode, examgubun, examcode, subcode, examname, " & vbCrLf & _
              "resgubun, range, seqno, reflow, refhigh) " & CR & _
              "Values ('" & gEquip & "', '" & Trim(txtEquipCode) & "', '" & Left(Trim(cmbGubun.Text), 1) & "', " & vbCrLf & _
              "'" & Trim(txtCode.Text) & "', '" & Trim(txtSubCode) & "', '" & Trim(txtName) & "', " & vbCrLf & _
              "'" & cmbGubun.ListIndex & "', '" & Trim(txtDec) & "', '" & Trim(txtSeq) & "', " & vbCrLf & _
              "'" & Trim(txtRefLow) & "', '" & Trim(txtRefHigh) & "') "
    End If

    res = SendQuery(gLocal, SQL)
    If res = -1 Then
'        db_RollBack gLocal
        SaveQuery SQL
        Exit Sub
    End If
    
'    db_Commit gLocal
    
    'gEquip = txtMuch
    
    DisplayList
    
    cmdCancel_Click
End Sub

Private Sub Form_Load()
    Me.Height = 8910
    Me.Width = 12390
    
    cmbGubun.AddItem "1.검진", 0
    cmbGubun.AddItem "2.진료", 1
    
    ClearText
    DisplayList

    txtMuch = gEquip
End Sub

Private Sub Text1_Change()

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
        cmbGubun.SetFocus
    End If
End Sub

Private Sub cmbGubun_GotFocus()
    SelectFocus cmbGubun
End Sub

Private Sub cmbGubun_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If cmbGubun = "" Then
            cmbGubun.SetFocus
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
        
        cmdSave.SetFocus
    End If
End Sub

Private Sub txtcode_GotFocus()
    SelectFocus txtCode
End Sub

Private Sub txtcode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
'        txtCode = UCase(txtCode)
        res = ExistOfEquipCode(Trim(txtEquipCode), Trim(txtCode))
        If res = -1 Then
            txtCode.SetFocus
            Exit Sub
        ElseIf res = 0 Then
            cmdSave.Caption = "저장"
            
        ElseIf res = 1 Then
'            equipcode, examcode, subcode, examname, range, seqno, reflow, refhigh, resgubun, examgubun
            cmdSave.Caption = "수정"
            txtSubCode = Trim(gReadBuf(2))
            txtName = Trim(gReadBuf(3))
            txtDec = Trim(gReadBuf(4))
            txtSeq = Trim(gReadBuf(5))
            txtRefLow = Trim(gReadBuf(6))
            txtRefHigh = Trim(gReadBuf(7))
            If IsNumeric(Trim(gReadBuf(8))) = True Then
                cmbGubun.ListIndex = Trim(gReadBuf(8))
            End If
            If IsNumeric(Trim(gReadBuf(9))) = True Then
                cmbResType.ListIndex = Trim(gReadBuf(9))
            End If
            
        End If
        
        txtSubCode.SetFocus
    End If
End Sub

Private Sub txtRefHigh_GotFocus()
    SelectFocus txtRefHigh
End Sub

Private Sub txtRefHigh_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Trim(txtRefHigh.Text) = "" Then
            txtRefHigh.SetFocus
            Exit Sub
        End If
        txtDec.SetFocus
    End If
End Sub

Private Sub txtRefLow_GotFocus()
    SelectFocus txtRefLow
End Sub

Private Sub txtRefLow_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Trim(txtRefLow.Text) = "" Then
            txtRefLow.SetFocus
            Exit Sub
        End If
        txtRefHigh.SetFocus
    End If
End Sub

Private Sub txtSubCode_GotFocus()
    SelectFocus txtSubCode
End Sub

Private Sub txtSubCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Trim(txtSubCode.Text) = "" Then
            txtSubCode.SetFocus
            Exit Sub
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
        cmbResType.SetFocus
        
    End If
End Sub

Private Sub cmbResType_GotFocus()
    SelectFocus cmbResType
End Sub

Private Sub cmbResType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Trim(cmbResType.Text) = "" Then
            cmbResType.SetFocus
            Exit Sub
        End If
        txtSeq.SetFocus
        
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

        txtRefLow.SetFocus
    End If
End Sub

Private Sub vasList_Click(ByVal Col As Long, ByVal Row As Long)
    If Row = 0 Then
        Select Case Col
        Case 1
            vasSort vasList, 1, 2
        Case 2
            vasSort vasList, 2, 1
        End Select
        Exit Sub
    End If
    
    If Row < 1 Or Row > vasList.DataRowCnt Then
        cmdSave.Caption = "저장"
        ClearText
        Exit Sub
    End If
    
    txtEquipCode = Trim(GetText(vasList, Row, 1))
'    cmbGubun.ListIndex = Trim(GetText(vasList, Row, 2))
       
    Select Case Trim(GetText(vasList, Row, 2))
    Case "진료"
        cmbGubun.ListIndex = 1
    Case "검진"
        cmbGubun.ListIndex = 0
    Case Else
        cmbGubun.ListIndex = -1
    End Select
    
    txtCode = Trim(GetText(vasList, Row, 3))
    txtSubCode = Trim(GetText(vasList, Row, 4))
    txtName = Trim(GetText(vasList, Row, 5))
    
    Select Case Trim(GetText(vasList, Row, 6))
    Case "1"
        cmbResType.ListIndex = 0
    Case "2"
        cmbResType.ListIndex = 1
    Case Else
        cmbResType.ListIndex = -1
    End Select
    
    txtSeq = Trim(GetText(vasList, Row, 8))
    
    txtRefLow = Trim(GetText(vasList, Row, 9))
    txtRefHigh = Trim(GetText(vasList, Row, 10))
    txtDec = Trim(GetText(vasList, Row, 7))
    
    
    cmdSave.Caption = "수정"
End Sub
