VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frmQCSet 
   Caption         =   "QC 설정"
   ClientHeight    =   8730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11865
   BeginProperty Font 
      Name            =   "굴림체"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8730
   ScaleWidth      =   11865
   StartUpPosition =   2  '화면 가운데
   Begin FPSpread.vaSpread vasList 
      Height          =   5115
      Left            =   60
      TabIndex        =   19
      Top             =   3510
      Width           =   11685
      _Version        =   196608
      _ExtentX        =   20611
      _ExtentY        =   9022
      _StockProps     =   64
      ColHeaderDisplay=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   9
      MaxRows         =   20
      ScrollBars      =   2
      SpreadDesigner  =   "frmQCSet.frx":0000
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   585
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   11685
      _Version        =   65536
      _ExtentX        =   20611
      _ExtentY        =   1032
      _StockProps     =   15
      Caption         =   "       Electsys 2010  QC 설정"
      ForeColor       =   8388608
      BackColor       =   16774393
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      Alignment       =   1
   End
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   60
      TabIndex        =   1
      Top             =   690
      Width           =   11685
      Begin VB.CommandButton cmdClose 
         Caption         =   "종료"
         Height          =   405
         Left            =   9600
         TabIndex        =   18
         Top             =   2130
         Width           =   1365
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   405
         Left            =   8160
         TabIndex        =   17
         Top             =   2130
         Width           =   1365
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "삭제"
         Height          =   405
         Left            =   6720
         TabIndex        =   16
         Top             =   2130
         Width           =   1365
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "저장"
         Height          =   405
         Left            =   5280
         TabIndex        =   15
         Top             =   2130
         Width           =   1365
      End
      Begin VB.TextBox txtLevelName 
         Appearance      =   0  '평면
         Height          =   315
         Left            =   2130
         TabIndex        =   8
         Top             =   870
         Width           =   1635
      End
      Begin VB.TextBox txtLevelNo 
         Appearance      =   0  '평면
         Height          =   315
         Left            =   1620
         TabIndex        =   7
         Top             =   870
         Width           =   495
      End
      Begin VB.TextBox txtLotNo 
         Appearance      =   0  '평면
         Height          =   315
         Left            =   1620
         TabIndex        =   5
         Top             =   1320
         Width           =   1635
      End
      Begin VB.TextBox txtEquipNo 
         Appearance      =   0  '평면
         Height          =   315
         Left            =   1620
         TabIndex        =   3
         Top             =   360
         Width           =   2145
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   1620
         TabIndex        =   10
         Top             =   1800
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         Format          =   23658497
         CurrentDate     =   37676
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   315
         Left            =   1620
         TabIndex        =   11
         Top             =   2220
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         Format          =   23658497
         CurrentDate     =   37676
      End
      Begin FPSpread.vaSpread vasExam 
         Height          =   1275
         Left            =   4350
         TabIndex        =   12
         Top             =   750
         Width           =   6615
         _Version        =   196608
         _ExtentX        =   11668
         _ExtentY        =   2249
         _StockProps     =   64
         ColHeaderDisplay=   1
         EditModePermanent=   -1  'True
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   5
         MaxRows         =   10
         ScrollBars      =   2
         SelectBlockOptions=   0
         SpreadDesigner  =   "frmQCSet.frx":045E
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "까지"
         Height          =   195
         Left            =   3240
         TabIndex        =   14
         Top             =   2280
         Width           =   420
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "부터"
         Height          =   195
         Left            =   3240
         TabIndex        =   13
         Top             =   1860
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "유효기간"
         Height          =   195
         Left            =   480
         TabIndex        =   9
         Top             =   1860
         Width           =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Level"
         Height          =   195
         Left            =   480
         TabIndex        =   6
         Top             =   930
         Width           =   525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Lot No."
         Height          =   195
         Left            =   480
         TabIndex        =   4
         Top             =   1380
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "장비구분"
         Height          =   195
         Left            =   480
         TabIndex        =   2
         Top             =   420
         Width           =   840
      End
   End
End
Attribute VB_Name = "frmQCSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub Display_QC()
    SQL = "Select levelno, levelname, lotno, validstart, validend, equipcode, examname, t_mean, t_sd " & vbCrLf & _
          "from qcexam where equipno = '" & gEquip & "'"
    db_select_Vas gLocal, SQL, vasList
End Sub

Private Sub cmdClear_Click()
    txtEquipNo = gEquip
    txtLevelNo = ""
    txtLevelName = ""
    txtLotNo = ""
    
    ClearSpread vasExam
    ClearSpread vasList
    
    Display_QC
    
    SQL = "Select equipcode, examname from equipexam where equipno = '" & gEquip & "' "
    db_select_Vas gLocal, SQL, vasExam, 1, 2
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    Dim i As Integer
    Dim sValid1, sValid2 As String
    
    sValid1 = Format(DTPicker1.Value, "yyyymmdd")
    sValid2 = Format(DTPicker2.Value, "yyyymmdd")
    
    For i = 1 To vasExam.DataRowCnt
        vasExam.Row = i
        vasExam.Col = 1
        If vasExam.Value = 1 Then
            SQL = "Delete from qcexam " & vbCrLf & _
                  "where equipno = '" & gEquip & "' " & vbCrLf & _
                  "  and lotno = '" & Trim(txtLotNo.Text) & "' " & vbCrLf & _
                  "  and levelno = " & Trim(txtLevelNo.Text) & " " & vbCrLf & _
                  "  and levelname = '" & Trim(txtLevelName) & "' " & vbCrLf & _
                  "  and validstart = '" & sValid1 & "' " & vbCrLf & _
                  "  and equipcode = '" & Trim(GetText(vasExam, i, 2)) & "' "
            res = SendQuery(gLocal, SQL)
            If res = -1 Then
                SaveQuery SQL
                Exit Sub
            End If
        End If
    Next i
    
    cmdClear_Click
End Sub

Private Sub cmdSave_Click()
    Dim i As Integer
    Dim sCnt As String
    Dim sValid1, sValid2 As String
    
    sValid1 = Format(DTPicker1.Value, "yyyymmdd")
    sValid2 = Format(DTPicker2.Value, "yyyymmdd")
    
    
    db_BeginTran gLocal
    
    For i = 1 To vasExam.DataRowCnt
        vasExam.Row = i
        vasExam.Col = 1
        If vasExam.Value = 1 Then
            sCnt = ""
            
            SQL = "select count(*) from qcexam " & vbCrLf & _
                  "where equipno = '" & gEquip & "' " & vbCrLf & _
                  "  and lotno = '" & Trim(txtLotNo.Text) & "' " & vbCrLf & _
                  "  and levelno = " & Trim(txtLevelNo.Text) & " " & vbCrLf & _
                  "  and levelname = '" & Trim(txtLevelName) & "' " & vbCrLf & _
                  "  and validstart = '" & sValid1 & "' " & vbCrLf & _
                  "  and equipcode = '" & Trim(GetText(vasExam, i, 2)) & "' "
            res = db_select_Var(gLocal, SQL, sCnt)
            If IsNumeric(sCnt) = False Then
                sCnt = "0"
            End If
            If CInt(sCnt) > 0 Then
                SQL = "Delete from qcexam " & vbCrLf & _
                      "where equipno = '" & gEquip & "' " & vbCrLf & _
                      "  and lotno = '" & Trim(txtLotNo.Text) & "' " & vbCrLf & _
                      "  and levelno = " & Trim(txtLevelNo.Text) & " " & vbCrLf & _
                      "  and levelname = '" & Trim(txtLevelName) & "' " & vbCrLf & _
                      "  and validstart = '" & sValid1 & "' " & vbCrLf & _
                      "  and equipcode = '" & Trim(GetText(vasExam, i, 2)) & "' "
                res = SendQuery(gLocal, SQL)
                If res = -1 Then
                    SaveQuery SQL
                    db_RollBack gLocal
                    Exit Sub
                End If
            End If
            SQL = "Insert into qcexam (equipno, lotno, levelno, levelname, validstart, validend, " & _
                  " equipcode, examname, t_mean, t_sd ) " & vbCrLf & _
                  "values ('" & gEquip & "', '" & Trim(txtLotNo) & "', " & txtLevelNo & ", '" & Trim(txtLevelName) & "', '" & sValid1 & "', '" & sValid2 & "', " & _
                  " '" & Trim(GetText(vasExam, i, 2)) & "', '" & Trim(GetText(vasExam, i, 3)) & "', '" & Trim(GetText(vasExam, i, 4)) & "', '" & Trim(GetText(vasExam, i, 5)) & "' ) "
            res = SendQuery(gLocal, SQL)
            If res = -1 Then
                SaveQuery SQL
                db_RollBack gLocal
                Exit Sub
            End If
        End If
    Next i
    
    db_Commit gLocal
    
    cmdClear_Click
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        DTPicker2.SetFocus
    End If
End Sub

Private Sub DTPicker2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        vasActiveCell vasExam, 1, 4
        vasExam.SetFocus
    End If
End Sub

Private Sub Form_Load()
    cmdClear_Click
End Sub

Private Sub txtLevelName_GotFocus()
    SelectFocus txtLevelName
End Sub

Private Sub txtLevelName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtLotNo.SetFocus
    End If
End Sub

Private Sub txtLevelNo_GotFocus()
    SelectFocus txtLevelNo
End Sub

Private Sub txtLevelNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtLevelName.SetFocus
    End If
End Sub

Private Sub txtLotNo_GotFocus()
    SelectFocus txtLotNo
End Sub

Private Sub txtLotNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        DTPicker1.SetFocus
    End If
End Sub

Private Sub vasExam_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim iRow As Long
    Dim iCol As Long
    
    iRow = vasExam.ActiveRow
    iCol = vasExam.ActiveCol
    
    If KeyCode = vbKeyReturn Then
        If iCol = 4 Then
            vasActiveCell vasExam, iRow, 5
        ElseIf iCol = 5 Then
            If iRow < vasExam.DataRowCnt Then
                vasActiveCell vasExam, iRow + 1, 4
            Else
                cmdSave.SetFocus
            End If
        End If
    End If
End Sub

Private Sub vasList_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim i As Integer
    Dim sTmp As String
    
    If Row < 1 Or Row > vasList.DataRowCnt Then
        Exit Sub
    End If
    
    cmdClear_Click
    
    txtLevelNo = Trim(GetText(vasList, Row, 1))
    txtLevelName = Trim(GetText(vasList, Row, 2))
    txtLotNo = Trim(GetText(vasList, Row, 3))
    
    sTmp = Trim(GetText(vasList, Row, 4))
    DTPicker1.Year = Left(sTmp, 4)
    DTPicker1.Month = Mid(sTmp, 5, 2)
    DTPicker1.Day = Mid(sTmp, 7, 2)
    sTmp = Trim(GetText(vasList, Row, 5))
    DTPicker2.Year = Left(sTmp, 4)
    DTPicker2.Month = Mid(sTmp, 5, 2)
    DTPicker2.Day = Mid(sTmp, 7, 2)
    
    For i = 1 To vasExam.DataRowCnt
        If Trim(GetText(vasExam, i, 2)) = Trim(GetText(vasList, Row, 6)) Then
            vasExam.Row = i
            vasExam.Col = 1
            vasExam.Value = 1
            vasExam.Col = 4
            vasExam.Text = Trim(GetText(vasList, Row, 8))
            vasExam.Col = 5
            vasExam.Text = Trim(GetText(vasList, Row, 9))
            Exit For
        End If
    Next i
End Sub
