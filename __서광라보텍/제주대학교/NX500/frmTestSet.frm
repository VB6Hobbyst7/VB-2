VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmTestSet 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "장비코드설정"
   ClientHeight    =   7800
   ClientLeft      =   2655
   ClientTop       =   1275
   ClientWidth     =   11940
   Icon            =   "frmTestSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   11940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   10560
      Style           =   1  '그래픽
      TabIndex        =   24
      Top             =   7140
      Width           =   1125
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  '위 맞춤
      BackColor       =   &H00F8E4D8&
      Height          =   225
      Left            =   0
      ScaleHeight     =   165
      ScaleWidth      =   11880
      TabIndex        =   23
      Top             =   0
      Width           =   11940
   End
   Begin VB.CheckBox chkUnit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "검사결과에 단위포함하여 표시"
      Height          =   255
      Left            =   8520
      TabIndex        =   22
      Top             =   4860
      Width           =   3015
   End
   Begin FPSpread.vaSpread vasList 
      Height          =   7335
      Left            =   30
      TabIndex        =   17
      Top             =   300
      Width           =   8115
      _Version        =   393216
      _ExtentX        =   14314
      _ExtentY        =   12938
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
      MaxCols         =   7
      MaxRows         =   20
      OperationMode   =   2
      RetainSelBlock  =   0   'False
      ScrollBarExtMode=   -1  'True
      SelectBlockOptions=   0
      ShadowColor     =   16777215
      SpreadDesigner  =   "frmTestSet.frx":08CA
      ScrollBarTrack  =   3
      ShowScrollTips  =   3
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   4395
      Left            =   8250
      TabIndex        =   9
      Top             =   210
      Width           =   3525
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
         TabIndex        =   19
         Top             =   2850
         Width           =   675
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
         Width           =   675
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2310
         Style           =   1  '그래픽
         TabIndex        =   8
         Top             =   3690
         Width           =   1035
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1230
         Style           =   1  '그래픽
         TabIndex        =   7
         Top             =   3690
         Width           =   1035
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   150
         Style           =   1  '그래픽
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
         Left            =   2820
         Picture         =   "frmTestSet.frx":0F55
         ScaleHeight     =   330
         ScaleWidth      =   330
         TabIndex        =   16
         Top             =   1170
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

    SQL = ""
    SQL = SQL & "Select EQUIPCODE, EXAMCODE, EXAMNAME, RESPREC, SEQNO, REFLOW, REFHIGH      " & vbCrLf
    SQL = SQL & "  From EQPMASTER                                                           " & vbCrLf
    SQL = SQL & " Where EQUIPNO     = '" & gEquip & "'                                      " & vbCrLf
    SQL = SQL & " Group By EXAMCODE, EQUIPCODE, EXAMNAME, RESPREC, SEQNO, REFLOW, REFHIGH   " & vbCrLf
    SQL = SQL & " Order By SEQNO * 10                                                       " & vbCrLf
          
    Res = GetDBSelectVas(gLocal, SQL, vasList)
    
    vasList.MaxRows = vasList.DataRowCnt
    vasList.RowHeight(-1) = 12
    
End Sub

'-- 장비코드와 수가코드에 해당하는 데이타 존재 확인 하는 procedure
Function ExistOfEquipCode(asEquipCode As String, Optional asSuga As String = "") As Integer

    ExistOfEquipCode = -1
    
    If asEquipCode = "" Then
        Exit Function
    End If
    
    SQL = ""
    SQL = SQL & "Select EQUIPCODE, EXAMCODE, EXAMNAME, RESPREC, SEQNO, REFLOW, REFHIGH  " & vbCrLf
    SQL = SQL & "  From EQPMASTER                                                       " & vbCrLf
    SQL = SQL & " Where EQUIPNO     = '" & gEquip & "'                                  " & vbCrLf
    SQL = SQL & "   And EQUIPCODE   = '" & asEquipCode & "'                             " & vbCrLf
    If Trim(asSuga) <> "" Then
        SQL = SQL & "   And EXAMCODE = '" & asSuga & "'                                 " & vbCrLf
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

Private Sub chkUnit_Click()

    If chkUnit.Value = "1" Then
        Call WritePrivateProfileString("config", "IncludUnit", "Y", App.Path & "\Interface.ini")
    Else
        Call WritePrivateProfileString("config", "IncludUnit", "N", App.Path & "\Interface.ini")
    End If
    
End Sub

Private Sub cmdCancel_Click()
    Call ClearText
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
    
    SQL = ""
    SQL = SQL & "Delete From EQPMASTER                              " & vbCrLf
    SQL = SQL & " Where EQUIPNO     = '" & gEquip & "'              " & vbCrLf
    SQL = SQL & "   And EQUIPCODE   = '" & Trim(txtEquipCode) & "'  " & vbCrLf
    SQL = SQL & "   And EXAMCODE    = '" & Trim(txtCode) & "'       " & vbCrLf
    
    Res = SendQuery(gLocal, SQL)
    If Res = -1 Then
        Exit Sub
    End If
    
    Call DisplayList
    
    Call cmdCancel_Click

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
        SQL = ""
        SQL = SQL & "Update EQPMASTER " & vbCrLf
        SQL = SQL & "   Set RESPREC     = '" & Trim(txtDec) & "'        " & vbCrLf
        SQL = SQL & "     , EXAMNAME    = '" & Trim(txtName) & "'       " & vbCrLf
        SQL = SQL & "     , REFLOW      = '" & Trim(txtRefLow) & "'     " & vbCrLf
        SQL = SQL & "     , REFHIGH     = '" & Trim(txtRefHigh) & "'    " & vbCrLf
        SQL = SQL & "     , SEQNO       = " & liSeqNo & "               " & vbCrLf
        SQL = SQL & " Where EQUIPNO     = '" & gEquip & "'              " & vbCrLf
        SQL = SQL & "   And EQUIPCODE   = '" & Trim(txtEquipCode) & "'  " & vbCrLf
        SQL = SQL & "   And EXAMCODE    = '" & Trim(txtCode) & "'       " & vbCrLf
    ElseIf Res = 0 Then
        SQL = ""
        SQL = SQL & "INSERT INTO EQPMASTER (EQUIPNO, EQUIPCODE, EXAMCODE, EXAMNAME, RESPREC, SEQNO, REFLOW, REFHIGH) " & vbCrLf
        SQL = SQL & " Values " & vbCrLf
        SQL = SQL & "('" & gEquip & "', '" & Trim(txtEquipCode) & "', '" & Trim(txtCode) & "', '" & Trim(txtName.Text) & "', '" & Trim(txtDec) & "', " & liSeqNo & ", '" & Trim(txtRefLow) & "', '" & Trim(txtRefHigh) & "')"
    End If

    Res = SendQuery(gLocal, SQL)
    If Res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If
    
    Call DisplayList
    
    Call cmdCancel_Click
    
End Sub

Private Sub Form_Load()
'    Me.Height = 7725
'    Me.Width = 11700
            
    ClearText
    DisplayList

    If gUnit = True Then
        chkUnit.Value = "1"
    Else
        chkUnit.Value = "0"
    End If
    
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
    txtName = Trim(GetText(vasList, Row, 3))
    txtDec = Trim(GetText(vasList, Row, 4))
    txtSeq = Trim(GetText(vasList, Row, 5))
    txtRefLow = Trim(GetText(vasList, Row, 6))
    txtRefHigh = Trim(GetText(vasList, Row, 7))

    cmdSave.Caption = "Edit"
End Sub
