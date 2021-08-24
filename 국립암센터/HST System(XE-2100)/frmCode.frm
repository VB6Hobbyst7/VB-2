VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmCode 
   Caption         =   "장비 코드 설정"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   11850
   StartUpPosition =   2  '화면 가운데
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1905
      Left            =   7470
      Picture         =   "frmCode.frx":0000
      ScaleHeight     =   1875
      ScaleWidth      =   4275
      TabIndex        =   32
      Top             =   6180
      Width           =   4305
   End
   Begin VB.Frame Frame1 
      Height          =   5355
      Left            =   7470
      TabIndex        =   1
      Top             =   780
      Width           =   4275
      Begin VB.TextBox txtGubun 
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
         TabIndex        =   29
         Top             =   2400
         Width           =   2655
      End
      Begin VB.TextBox txtCode 
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
         TabIndex        =   23
         Top             =   660
         Width           =   2655
      End
      Begin VB.TextBox txtUnit 
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
         TabIndex        =   19
         Top             =   3300
         Width           =   2655
      End
      Begin VB.TextBox txtRefLow 
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
         TabIndex        =   17
         Top             =   2850
         Width           =   945
      End
      Begin VB.TextBox txtRefHigh 
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
         Left            =   2820
         TabIndex        =   16
         Top             =   2850
         Width           =   945
      End
      Begin VB.ComboBox cboGubun 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmCode.frx":0D49
         Left            =   1410
         List            =   "frmCode.frx":0D56
         TabIndex        =   13
         Top             =   2010
         Visible         =   0   'False
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
         Height          =   675
         Left            =   3150
         TabIndex        =   12
         Top             =   4470
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
         Height          =   675
         Left            =   2160
         TabIndex        =   11
         Top             =   4470
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
         Height          =   675
         Left            =   1170
         TabIndex        =   10
         Top             =   4470
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
         Height          =   675
         Left            =   180
         TabIndex        =   9
         Top             =   4470
         Width           =   975
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
         Top             =   2415
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.TextBox txtName 
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
         TabIndex        =   5
         Top             =   1104
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
         Left            =   1410
         TabIndex        =   3
         Top             =   225
         Width           =   2655
      End
      Begin VB.TextBox txtSeqNo 
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
         TabIndex        =   25
         Top             =   1530
         Width           =   2655
      End
      Begin VB.ComboBox cboOrdGubun 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmCode.frx":0D78
         Left            =   1410
         List            =   "frmCode.frx":0D97
         TabIndex        =   27
         Top             =   2010
         Width           =   2625
      End
      Begin VB.TextBox txtNo 
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
         Left            =   3360
         TabIndex        =   31
         Top             =   660
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "결과처리"
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
         TabIndex        =   30
         Top             =   2475
         Width           =   1020
      End
      Begin VB.Label Label9 
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
         Left            =   240
         TabIndex        =   24
         Top             =   735
         Width           =   1020
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "단    위"
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
         TabIndex        =   20
         Top             =   3375
         Width           =   1050
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
         Left            =   2520
         TabIndex        =   18
         Top             =   2925
         Width           =   135
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "보고범위"
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
         TabIndex        =   15
         Top             =   2940
         Width           =   1020
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "장비소수"
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
         TabIndex        =   7
         Top             =   2490
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "유    형"
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
         TabIndex        =   6
         Top             =   2070
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.Label Label2 
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
         Left            =   240
         TabIndex        =   4
         Top             =   1170
         Width           =   1035
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
         Left            =   240
         TabIndex        =   2
         Top             =   300
         Width           =   1020
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "순    서"
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
         TabIndex        =   26
         Top             =   1590
         Width           =   1050
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
         Left            =   240
         TabIndex        =   28
         Top             =   2070
         Width           =   1020
      End
   End
   Begin FPSpread.vaSpread vasList 
      Height          =   7215
      Left            =   60
      TabIndex        =   0
      Top             =   870
      Width           =   7365
      _Version        =   393216
      _ExtentX        =   12991
      _ExtentY        =   12726
      _StockProps     =   64
      ColHeaderDisplay=   1
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   10
      ScrollBars      =   2
      SpreadDesigner  =   "frmCode.frx":0E2F
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   675
      Left            =   60
      TabIndex        =   14
      Top             =   60
      Width           =   11685
      _Version        =   65536
      _ExtentX        =   20611
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
      Begin VB.TextBox txtEquip 
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
         Left            =   5310
         TabIndex        =   21
         Top             =   120
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "장비구분"
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
         Left            =   4140
         TabIndex        =   22
         Top             =   195
         Visible         =   0   'False
         Width           =   1020
      End
   End
End
Attribute VB_Name = "frmCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lsGubun As String
Dim lsCode As String
Dim lsName As String
Dim lsSeqNo As String
Dim lsRang As String
Dim lsRefLow As String
Dim lsRefHigh As String
Dim lsUnit As String

Sub ClearText()
    lsGubun = ""
    lsName = ""
    lsRang = ""

    
    txtEquipCode = ""
    txtCode = ""
    txtNo = "00"
    txtName = ""
    txtSeqNo = ""
    
    cboGubun.ListIndex = -1
    txtRang = ""
    txtRefLow = ""
    txtRefHigh = ""
    txtUnit = ""
    
    cmdSave.Caption = "저장"
End Sub

Sub DisplayList()
    ClearSpread vasList
    
    SQL = "SELECT EquipCode,ExamCode, ExamNo, ExamName, Seqno, OrdGubun, RSGubun,PointSize,RefLow + ' - ' + RefHigh,UnitCode " & vbCrLf & _
          "  From EquipExam " & vbCrLf & _
          " WHERE Equip = '" & gEquip & "' " & vbCrLf & _
          " Order by seqno, equipcode "
          
    db_select_Vas gLocal, SQL, vasList
    
    vasList.MaxRows = vasList.DataRowCnt
End Sub

Function ExistOfEquipCode(asEquipCode As String, Optional asExamCode As String = "") As Integer
'장비코드와 검사명에 해당하는 데이타 존재 확인 하는 procedure

    ExistOfEquipCode = -1
    
    If asEquipCode = "" Then
        Exit Function
    End If
    
    SQL = "SELECT EquipCode, ExamCode, ExamName, SeqNo, RSGubun, PointSize, RefLow, RefHigh, UnitCode " & vbCrLf & _
          "  From EquipExam " & vbCrLf & _
          " WHERE Equip = '" & gEquip & "' " & vbCrLf & _
          "   AND EquipCode = '" & asEquipCode & "' "
    If Trim(asExamCode) <> "" Then
        SQL = SQL & vbCrLf & "   AND ExamCode = '" & asExamCode & "' "
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
    
    lsCode = Trim(gReadBuf(1))
    lsName = Trim(gReadBuf(2))
    lsSeqNo = Trim(gReadBuf(3))
    lsGubun = Trim(gReadBuf(4))
    lsRang = Trim(gReadBuf(5))
    lsRefLow = Trim(gReadBuf(6))
    lsRefHigh = Trim(gReadBuf(7))
    lsUnit = Trim(gReadBuf(8))
    
    ExistOfEquipCode = 1
End Function

Function Select_Suga_Info(asSuga As String) As Integer
    Select_Suga_Info = -1
    
    If Trim(asSuga) = "" Then
        Exit Function
    End If
    
    SQL = "Select coifcode, coifrclf, coifrfit, coifrfpr, coifleng, coifcona " & vbCrLf & _
          "From ABCCOIFM" & vbCrLf & _
          "Where coifcode = '" & asSuga & "' "
    res = db_select_Col(gServer, SQL)
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    ElseIf res = 0 Then
        Select_Suga_Info = 0
        Exit Function
    End If
    If Trim(gReadBuf(0)) <> asSuga Then
        Select_Suga_Info = 0
        Exit Function
    End If
    
    txtRang = ""
    Select Case Left(Trim(gReadBuf(1)), 1)
    Case "N"
        If IsNumeric(Trim(gReadBuf(3))) = True Then
            If CInt(gReadBuf(3)) > 0 Then
                cboGubun.ListIndex = 1
                txtRang = Trim(gReadBuf(2)) & "." & Trim(gReadBuf(3))
            Else
                cboGubun.ListIndex = 0
                txtRang = Trim(gReadBuf(2))
            End If
        Else
            cboGubun.ListIndex = -1
        End If
    Case "T"
        cboGubun.ListIndex = 2
        txtRang = Trim(gReadBuf(4))
    Case Else
        cboGubun.ListIndex = -1
    End Select
    txtName = Trim(gReadBuf(5))
    
    Select_Suga_Info = 1
End Function

Private Sub cboGubun_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If cboGubun.ListIndex < 0 Then
            cboGubun.SetFocus
            Exit Sub
        End If
        
        'txtRang.SetFocus
        txtGubun.SetFocus
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
         
'    db_BeginTran gLocal
    
    SQL = "Delete from EquipExam " & vbCrLf & _
          "Where Equip = '" & gEquip & "' " & vbCrLf & _
          "  and EquipCode = '" & Trim(txtEquipCode) & "' " & vbCrLf & _
          "  and ExamCode = '" & Trim(txtCode) & "' "
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

    If Trim(txtEquipCode) = "" Then
        txtEquipCode.SetFocus
        Exit Sub
    End If
    
    If Trim(txtName) = "" Then
        txtName.SetFocus
        Exit Sub
    End If
        
    If Trim(txtRang) = "" Then
        txtRang.Text = "0"
    End If
    
    If Trim(txtNo) = "" Then
        txtNo = "00"
    End If
    
    If IsNumeric(txtSeqNo) = False And IsNumeric(txtEquipCode) = True Then txtSeqNo = txtEquipCode
    
'    db_BeginTran gLocal
    
    res = ExistOfEquipCode(Trim(txtEquipCode), Trim(txtCode))
    If res = 1 Then
        SQL = "Update EquipExam " & vbCrLf & _
              "Set ExamCode = '" & Trim(txtCode) & "', " & vbCrLf & _
              "    ExamNo = '" & Trim(txtNo) & "', " & vbCrLf & _
              "    ExamName = '" & Trim(txtName) & "', " & vbCrLf & _
              "    Seqno = " & Trim(txtSeqNo) & ", " & vbCrLf & _
              "    OrdGubun = '" & Left(cboOrdGubun.Text, 1) & "', " & vbCrLf & _
              "    RSGubun = '" & Trim(txtGubun.Text) & "', " & vbCrLf & _
              "    PointSize = " & Trim(txtRang) & ", " & vbCrLf & _
              "    RefLow = '" & Trim(txtRefLow) & "', " & vbCrLf & _
              "    RefHigh = '" & Trim(txtRefHigh) & "', " & vbCrLf & _
              "    UnitCode = '" & Trim(txtUnit) & "' " & vbCrLf & _
              "Where Equip = '" & gEquip & "' " & vbCrLf & _
              "  and EquipCode = '" & Trim(txtEquipCode) & "' " & vbCrLf & _
              "  and ExamCode = '" & Trim(txtCode) & "' "
    ElseIf res = 0 Then
        SQL = "Insert Into EquipExam (Equip,EquipCode,ExamCode,ExamNo,ExamName,SeqNo, RSGubun,PointSize,RefLow,RefHigh,UnitCode, OrdGubun ) " & vbCrLf & _
              "Values ('" & gEquip & "', '" & Trim(txtEquipCode) & "', '" & Trim(txtCode) & "', '" & Trim(txtNo) & "', '" & Trim(txtName) & "', " & Trim(txtSeqNo) & ", '" & Trim(txtGubun.Text) & "', " & Trim(txtRang) & ", '" & Trim(txtRefLow) & "', '" & Trim(txtRefHigh) & "', '" & Trim(txtUnit) & "', '" & Left(cboOrdGubun.Text, 1) & "' ) "
    End If
    res = SendQuery(gLocal, SQL)
    If res = -1 Then
'        db_RollBack gLocal
        Exit Sub
    End If
    
'    db_Commit gLocal
    
    DisplayList
    
    cmdCancel_Click
    
End Sub

Private Sub Form_Load()
'    Me.Height = 8600
'    Me.Width = 11970
            
    ClearText
    DisplayList
    
    txtEquip = gEquip
End Sub

Private Sub txtCode_GotFocus()
    SelectFocus txtCode
End Sub

Private Sub txtCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtName.SetFocus
    End If
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
        res = ExistOfEquipCode(Trim(txtEquipCode))
        If res = 1 Then
    
            txtCode = lsCode
            txtName = lsName
            txtSeqNo = lsSeqNo
            Select Case lsGubun
            Case "I"
                cboGubun.ListIndex = 0
            Case "F"
                cboGubun.ListIndex = 1
            Case "T"
                cboGubun.ListIndex = 2
            Case Else
                cboGubun.ListIndex = -1
            End Select
            txtRang = lsRang
            txtRefLow = lsRefLow
            txtRefHigh = lsRefHigh
            txtUnit = lsUnit
                
        End If
        txtCode.SetFocus
    End If
End Sub

Private Sub txtName_GotFocus()
    SelectFocus txtName
End Sub

Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txtName = "" Then
            txtName.SetFocus
            Exit Sub
        End If
        'cboGubun.SetFocus
    End If
    
End Sub

Private Sub txtRang_GotFocus()
    SelectFocus txtRang
End Sub

Private Sub txtRang_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txtRang = "" Then
            txtRang.SetFocus
            Exit Sub
        End If
        
        txtRefLow.SetFocus
    End If
End Sub

Private Sub txtRefHigh_GotFocus()
    SelectFocus txtRefHigh
End Sub

Private Sub txtRefHigh_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txtRefHigh = "" Then
            txtRefHigh.SetFocus
            Exit Sub
        End If
        
        txtUnit.SetFocus
    End If
End Sub

Private Sub txtRefLow_GotFocus()
    SelectFocus txtRefLow
End Sub

Private Sub txtRefLow_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txtRefLow = "" Then
            txtRefLow.SetFocus
            Exit Sub
        End If
        
        txtRefHigh.SetFocus
    End If
End Sub

Private Sub txtUnit_GotFocus()
    SelectFocus txtUnit
End Sub

Private Sub txtUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txtUnit = "" Then
            txtUnit.SetFocus
            Exit Sub
        End If

        cmdSave.SetFocus
    End If
End Sub

Private Sub vasList_Click(ByVal Col As Long, ByVal Row As Long)
Dim i As Integer

    If Row < 1 Or Row > vasList.DataRowCnt Then
        cmdSave.Caption = "저장"
        ClearText
        Exit Sub
    End If
    
    txtEquipCode = Trim(GetText(vasList, Row, 1))
    txtCode = Trim(GetText(vasList, Row, 2))
    txtNo = Trim(GetText(vasList, Row, 3))
    txtName = Trim(GetText(vasList, Row, 4))
    txtSeqNo = Trim(GetText(vasList, Row, 5))
    For i = 0 To cboOrdGubun.ListCount - 1
        If Left(cboOrdGubun.List(i), 1) = Trim(GetText(vasList, Row, 6)) Then
            cboOrdGubun.ListIndex = i
            Exit For
        End If
    Next i
    txtGubun = Trim(GetText(vasList, Row, 7))
'    Select Case Trim(GetText(vasList, Row, 6))
'    Case "I"
'        cboGubun.ListIndex = 0
'    Case "F"
'        cboGubun.ListIndex = 1
'    Case "T"
'        cboGubun.ListIndex = 2
'    Case Else
'        cboGubun.ListIndex = -1
'    End Select
    
    txtRang = Trim(GetText(vasList, Row, 8))
    
    
    i = InStr(1, Trim(GetText(vasList, Row, 9)), "-")
    If i > 0 Then
    txtRefLow = Trim(Mid(GetText(vasList, Row, 9), 1, i - 1))
    txtRefHigh = Trim(Mid(GetText(vasList, Row, 9), i + 1))
    End If
    
    txtUnit = Trim(GetText(vasList, Row, 10))
    
    cmdSave.Caption = "수정"
End Sub

Private Sub vasList_DblClick(ByVal Col As Long, ByVal Row As Long)
    If Row = 0 Then vasSort vasList, Col
End Sub
