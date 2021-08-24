VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmTestSet 
   Caption         =   "장비 코드 설정"
   ClientHeight    =   7125
   ClientLeft      =   2670
   ClientTop       =   1290
   ClientWidth     =   12495
   Icon            =   "frmTestSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7125
   ScaleWidth      =   12495
   StartUpPosition =   1  '소유자 가운데
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
      Left            =   11280
      TabIndex        =   22
      Top             =   6540
      Width           =   1125
   End
   Begin FPSpread.vaSpread vasList 
      Height          =   6945
      Left            =   90
      TabIndex        =   17
      Top             =   90
      Width           =   8685
      _Version        =   393216
      _ExtentX        =   15319
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
      MaxCols         =   8
      MaxRows         =   20
      OperationMode   =   2
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SelectBlockOptions=   0
      SpreadDesigner  =   "frmTestSet.frx":1272
   End
   Begin VB.Frame Frame1 
      Height          =   4395
      Left            =   8850
      TabIndex        =   9
      Top             =   90
      Width           =   3525
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
         Left            =   2400
         TabIndex        =   23
         Top             =   1170
         Width           =   825
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
         Left            =   2280
         TabIndex        =   19
         Top             =   2850
         Width           =   915
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
         TabIndex        =   18
         Top             =   2850
         Width           =   945
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Clear"
         Height          =   495
         Left            =   2310
         TabIndex        =   8
         Top             =   3690
         Width           =   1035
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   495
         Left            =   1230
         TabIndex        =   7
         Top             =   3690
         Width           =   1035
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   495
         Left            =   150
         TabIndex        =   6
         Top             =   3690
         Width           =   1035
      End
      Begin VB.PictureBox picEquip 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BorderStyle     =   0  '없음
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2790
         Picture         =   "frmTestSet.frx":1BF5
         ScaleHeight     =   330
         ScaleWidth      =   330
         TabIndex        =   16
         Top             =   3270
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
         TabIndex        =   5
         Top             =   2430
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
         TabIndex        =   3
         Top             =   1590
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
         TabIndex        =   4
         Top             =   2010
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
         TabIndex        =   2
         Top             =   1170
         Width           =   1275
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
         Left            =   2100
         TabIndex        =   21
         Top             =   2880
         Width           =   135
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
         Left            =   240
         TabIndex        =   20
         Top             =   2940
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
         TabIndex        =   15
         Top             =   2520
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
         TabIndex        =   14
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
         TabIndex        =   13
         Top             =   1665
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
         TabIndex        =   12
         Top             =   2085
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
         TabIndex        =   11
         Top             =   1230
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
         TabIndex        =   10
         Top             =   810
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmTestSet"
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

    SQL = "SELECT EQUIPCODE, EXAMCODE, SUBCODE, EXAMNAME, RESPREC, SEQNO, REFLOW, REFHIGH " & vbCrLf & _
          "  From EQPMASTER " & vbCrLf & _
          " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf & _
          " GROUP BY EXAMCODE, EQUIPCODE, SUBCODE, EXAMNAME, RESPREC, SEQNO, REFLOW, REFHIGH "
    SQL = SQL & " ORDER BY SEQNO "
          
    Res = GetDBSelectVas(gLocal, SQL, vasList)
    
    vasList.MaxRows = vasList.DataRowCnt
    vasList.RowHeight(-1) = 12
   ' Call vasList_Click(1, 0)
    
End Sub

'-- 장비코드와 수가코드에 해당하는 데이타 존재 확인 하는 procedure
Function ExistOfEquipCode(asEquipCode As String, Optional asSuga As String = "") As Integer

    ExistOfEquipCode = -1
    
    If asEquipCode = "" Then
        Exit Function
    End If
    
    SQL = "SELECT EQUIPCODE, EXAMCODE, SUBCODE, EXAMNAME, RESPREC, SEQNO, REFLOW, REFHIGH " & vbCrLf & _
          "  FROM EQPMASTER " & vbCrLf & _
          " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf & _
          "   AND EQUIPCODE = '" & asEquipCode & "' "
          
    If Trim(asSuga) <> "" Then
        SQL = SQL & vbCrLf & _
          "   AND EXAMCODE = '" & asSuga & "' "
    End If
    
    Res = GetDBSelectColumn(gLocal, SQL)
    If Res = 0 Then
        ExistOfEquipCode = 0
        Exit Function
    ElseIf Res = -1 Then
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
    
    SQL = "DELETE FROM EQPMASTER " & vbCrLf & _
          "WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf & _
          "  AND EQUIPCODE = '" & Trim(txtEquipCode) & "' " & vbCrLf & _
          "  AND EXAMCODE = '" & Trim(txtCode) & "' " & vbCrLf & _
          "  AND SUBCODE = '" & Trim(txtSubCode) & "' "
    Res = SendQuery(gLocal, SQL)
    If Res = -1 Then
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
    
    Res = ExistOfEquipCode(Trim(txtEquipCode), Trim(txtCode))
    If Res = 1 Then
        SQL = "UPDATE EQPMASTER " & vbCrLf & _
              "SET RESPREC = '" & Trim(txtDec) & "', " & vbCrLf & _
              "    EXAMNAME = '" & Trim(txtName) & "', " & vbCrLf & _
              "    REFLOW = '" & Trim(txtRefLow) & "', " & vbCrLf & _
              "    REFHIGH = '" & Trim(txtRefHigh) & "', " & vbCrLf & _
              "    SEQNO = " & liSeqNo & ", " & vbCrLf & _
              "    SUBCODE = '" & Trim(txtSubCode) & "' " & _
              "WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf & _
              "  AND EQUIPCODE = '" & Trim(txtEquipCode) & "' " & vbCrLf & _
              "  AND EXAMCODE = '" & Trim(txtCode) & "' "
    
    ElseIf Res = 0 Then
        SQL = "INSERT INTO EQPMASTER (EQUIPNO,EQUIPCODE, EXAMCODE, SUBCODE, EXAMNAME, RESPREC, SEQNO , REFLOW, REFHIGH) " & vbCrLf & _
              "VALUES ('" & gEquip & "', '" & Trim(txtEquipCode) & "', '" & Trim(txtCode) & "', '" & Trim(txtSubCode) & "', '" & Trim(txtName.Text) & "', '" & Trim(txtDec) & "', " & liSeqNo & ", '" & Trim(txtRefLow) & "', '" & Trim(txtRefHigh) & "') "
    End If

    Res = SendQuery(gLocal, SQL)
    If Res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If
    
    DisplayList
    
    cmdCancel_Click
End Sub


Private Sub Form_Load()
    Me.Height = 7725
    Me.Width = 12615
            
    ClearText
    DisplayList

    txtMuch = gEquip
    
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
        Res = ExistOfEquipCode(Trim(txtEquipCode), Trim(txtCode))
        If Res = -1 Then
            txtCode.SetFocus
            Exit Sub
        ElseIf Res = 0 Then
            cmdSave.Caption = "Save"
            
        ElseIf Res = 1 Then
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
    
    txtEquipCode = Trim(GetText(vasList, Row, 1))
    txtCode = Trim(GetText(vasList, Row, 2))
    txtSubCode = Trim(GetText(vasList, Row, 3))
    txtName = Trim(GetText(vasList, Row, 4))
    txtDec = Trim(GetText(vasList, Row, 5))
    txtSeq = Trim(GetText(vasList, Row, 6))
    txtRefLow = Trim(GetText(vasList, Row, 7))
    txtRefHigh = Trim(GetText(vasList, Row, 8))

    
    
    cmdSave.Caption = "Edit"
End Sub
