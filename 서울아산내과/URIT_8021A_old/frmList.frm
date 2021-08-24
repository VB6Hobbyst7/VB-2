VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmList 
   Caption         =   "List"
   ClientHeight    =   6480
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   10785
   LinkTopic       =   "Form1"
   ScaleHeight     =   6480
   ScaleWidth      =   10785
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CheckBox chkAll 
      Caption         =   "Check1"
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   3030
      TabIndex        =   20
      Top             =   240
      Width           =   225
   End
   Begin FPSpread.vaSpread vasListMach 
      Height          =   6255
      Left            =   2640
      TabIndex        =   0
      Top             =   120
      Width           =   8055
      _Version        =   393216
      _ExtentX        =   14208
      _ExtentY        =   11033
      _StockProps     =   64
      ColHeaderDisplay=   0
      ColsFrozen      =   1
      EditModeReplace =   -1  'True
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
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SpreadDesigner  =   "frmList.frx":0000
      UserResize      =   2
   End
   Begin FPSpread.vaSpread vasLExam 
      Height          =   1035
      Left            =   180
      TabIndex        =   18
      Top             =   7320
      Visible         =   0   'False
      Width           =   5595
      _Version        =   393216
      _ExtentX        =   9869
      _ExtentY        =   1826
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "frmList.frx":08C5
   End
   Begin FPSpread.vaSpread vasLTemp 
      Height          =   1755
      Left            =   180
      TabIndex        =   17
      Top             =   5880
      Visible         =   0   'False
      Width           =   5715
      _Version        =   393216
      _ExtentX        =   10081
      _ExtentY        =   3096
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "frmList.frx":0AEB
   End
   Begin VB.Frame Frame3 
      Caption         =   "[검사일자]"
      Height          =   615
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   2415
      Begin VB.Label lblExamDate 
         Caption         =   "2010-10-10"
         Height          =   255
         Left            =   180
         TabIndex        =   16
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "[매칭]"
      Height          =   1845
      Left            =   120
      TabIndex        =   12
      Top             =   3960
      Width           =   2415
      Begin VB.CommandButton cmdIFClear 
         Caption         =   "선택해제"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   180
         TabIndex        =   19
         Top             =   1380
         Width           =   2115
      End
      Begin VB.CommandButton cmdListCancel 
         Caption         =   "취소"
         Height          =   435
         Left            =   180
         TabIndex        =   14
         Top             =   840
         Width           =   2115
      End
      Begin VB.CommandButton cmdListMach 
         Caption         =   "리스트매칭"
         Height          =   435
         Left            =   180
         TabIndex        =   13
         Top             =   330
         Width           =   2115
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "[조회]"
      Height          =   3135
      Left            =   120
      TabIndex        =   1
      Top             =   780
      Width           =   2415
      Begin VB.CommandButton cmdListSch 
         Caption         =   "리스트조회"
         Height          =   435
         Left            =   180
         TabIndex        =   11
         Top             =   2460
         Width           =   2115
      End
      Begin VB.ComboBox cmbExamType 
         Height          =   300
         Left            =   240
         TabIndex        =   10
         Top             =   1920
         Width           =   2055
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  '없음
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   60
         ScaleHeight     =   315
         ScaleWidth      =   2295
         TabIndex        =   8
         Top             =   1500
         Width           =   2295
         Begin VB.Label Label4 
            BackStyle       =   0  '투명
            Caption         =   "검사선택"
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   60
            Width           =   1215
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  '없음
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   60
         ScaleHeight     =   315
         ScaleWidth      =   2295
         TabIndex        =   6
         Top             =   240
         Width           =   2295
         Begin VB.Label Label3 
            BackStyle       =   0  '투명
            Caption         =   "접수일자"
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   60
            Width           =   1215
         End
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   315
         Left            =   840
         TabIndex        =   2
         Top             =   660
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   61341697
         CurrentDate     =   40478
      End
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   315
         Left            =   840
         TabIndex        =   4
         Top             =   1020
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   61341697
         CurrentDate     =   40478
      End
      Begin VB.Label Label2 
         Caption         =   "To   :"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "From :"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   675
      End
   End
   Begin VB.Menu MnfrmList 
      Caption         =   "리스트삭제"
      Visible         =   0   'False
      Begin VB.Menu MnfrmListDel 
         Caption         =   "리스트삭제"
      End
   End
End
Attribute VB_Name = "frmList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkAll_Click()
    Dim iRow As Long
    
    If chkAll.Value = 1 Then
        For iRow = 1 To vasListMach.DataRowCnt
            vasListMach.Row = iRow
            vasListMach.Col = 1
            
            vasListMach.Value = 1
        Next iRow
    ElseIf chkAll.Value = 0 Then
        For iRow = 1 To vasListMach.DataRowCnt
            vasListMach.Row = iRow
            vasListMach.Col = 1
            
            vasListMach.Value = 0
        Next iRow
    End If
End Sub

Private Sub cmdIFClear_Click()
    Dim i As Integer
    
    For i = 1 To vasListMach.DataRowCnt
        vasListMach.Row = i
        vasListMach.Col = 1
        vasListMach.Value = 0
    Next i
    
End Sub

Private Sub cmdListCancel_Click()
    Dim i As Integer
    
    For i = 1 To vasListMach.DataRowCnt
        SetText vasListMach, "", i, 2
        
    Next
End Sub

Private Sub cmdListMach_Click()
    Dim i As Long
    Dim j As Integer
    Dim x As Integer
    Dim chServer
    Dim sExamCode As String
    Dim sSubCode As String
    Dim sEquipCode As String
    Dim sExamName As String
    Dim sSeqNo As String
    Dim msgRes
    Dim result As String
    Dim sHospital As String
    
    msgRes = MsgBox("결과매칭을 하시겠습니까?", vbYesNo, "결과매칭")
    If msgRes = 7 Then
        Exit Sub
    End If
    
    
    chServer = cmbExamType.ListIndex
    
    For i = 1 To vasListMach.DataRowCnt
        vasListMach.Col = 1
        vasListMach.Row = i
        
        If vasListMach.Value = 1 Then
        If Trim(GetText(vasListMach, i, 2)) <> "" Then
            ClearSpread vasLTemp
            Select Case chServer
            Case dpGumjin1
                SQL = "select exam_code " & vbCrLf & _
                      "from totres " & vbCrLf & _
                      "where request_date = '" & Trim(GetText(vasListMach, i, 3)) & "'  " & vbCrLf & _
                      "  and exam_no = '" & Trim(GetText(vasListMach, i, 4)) & "'  " & vbCrLf & _
                      "  and exam_code in (" & gAllExam & ") "
                res = db_select_Vas(gServer, SQL, vasLTemp)
            Case dpOCS
                SQL = "select 검사코드, 검사종류 " & vbCrLf & _
                      "from TB_진료검사 " & vbCrLf & _
                      "where 오더일련번호= '" & Trim(GetText(vasListMach, i, 5)) & "' " & vbCrLf & _
                      "and 챠트번호= '" & Trim(GetText(vasListMach, i, 4)) & "' " & vbCrLf & _
                      "and 검사코드 in (" & gAllExam_Ocs & ") "
                res = db_select_Vas(gServer_OCS, SQL, vasLTemp)
        
            Case dpGumjin2
                SQL = "select exam_code " & vbCrLf & _
                      "from twoexam " & vbCrLf & _
                      "where request_date = '" & Trim(GetText(vasListMach, i, 3)) & "'  " & vbCrLf & _
                      "  and exam_no = '" & Trim(GetText(vasListMach, i, 4)) & "'  " & vbCrLf & _
                      "  and exam_code in (" & gAllExam & ") "
                res = db_select_Vas(gServer, SQL, vasLTemp)
            End Select
            
            For j = 1 To vasLTemp.DataRowCnt
                sExamCode = ""
                sSubCode = ""
                For x = 1 To vasLExam.DataRowCnt
                    If Trim(GetText(vasLExam, x, 1)) = Trim(GetText(vasLTemp, j, 1)) And Trim(GetText(vasLExam, x, 2)) = Trim(GetText(vasLTemp, j, 2)) Then
                        sExamCode = Trim(GetText(vasLExam, x, 1))
                        sSubCode = Trim(GetText(vasLExam, x, 2))
                        sEquipCode = Trim(GetText(vasLExam, x, 3))
                        sExamName = Trim(GetText(vasLExam, x, 4))
                        sSeqNo = Trim(GetText(vasLExam, x, 5))
                        Exit For
                    End If
                Next x
                
                gReadBuf(0) = ""
                gReadBuf(1) = ""
                   
                SQL = "select result,hospital from pat_res " & vbCrLf & _
                      "where equipno = '" & gEquip & "' and sampleno = '" & Trim(GetText(vasListMach, i, 2)) & "' " & vbCrLf & _
                      "and equipcode = '" & sEquipCode & "' and examdate = '" & Format(lblExamDate.Caption, "yyyymmdd") & "' AND SENDFLAG = '0' "
                res = db_select_Col(gLocal, SQL)
                
                result = Trim(gReadBuf(0))
                sHospital = Trim(gReadBuf(1))
                
                If chServer = dpGumjin1 Or chServer = dpGumjin2 Then
                    Select Case result
                        Case "trace"
                         result = "약양성"
                        Case "norm"
                         result = "음성"
                    End Select
                End If
                
                SQL = "insert into pat_res(equipno, recedate, sampleno, receno, pid, " & vbCrLf & _
                        "pname, psex, page, pjumin, sendflag, " & vbCrLf & _
                        "examgubun,examdate, examcode, subcode, examname, equipcode, seqno,RESULT,hospital) " & vbCrLf & _
                        "values('" & gEquip & "', '" & Trim(GetText(vasListMach, i, 3)) & "', '" & Trim(GetText(vasListMach, i, 2)) & "', " & vbCrLf & _
                        "'" & Trim(GetText(vasListMach, i, 4)) & "', '" & Trim(GetText(vasListMach, i, 5)) & "', '" & Trim(GetText(vasListMach, i, 6)) & "', " & vbCrLf & _
                        "'" & Trim(GetText(vasListMach, i, 7)) & "', '" & Trim(GetText(vasListMach, i, 8)) & "', '" & Trim(GetText(vasListMach, i, 9)) & "', " & vbCrLf & _
                        "'2', '" & chServer & "','" & Format(lblExamDate.Caption, "YYYYMMDD") & "', '" & sExamCode & "', " & vbCrLf & _
                        "'" & sSubCode & "', '" & sExamName & "', '" & sEquipCode & "', '" & sSeqNo & "','" & result & "', '" & sHospital & "')"
                res = SendQuery(gLocal, SQL)
                
            Next j
            
            If res > 0 Then
                For x = 1 To frmInterface.vasRID.DataRowCnt
                    If Trim(GetText(frmInterface.vasRID, x, 3)) = Trim(GetText(vasListMach, i, 2)) Then
                        SetText frmInterface.vasRID, Trim(GetText(vasListMach, i, 3)), x, 2
                        SetText frmInterface.vasRID, Trim(GetText(vasListMach, i, 4)), x, 6
                        SetText frmInterface.vasRID, Trim(GetText(vasListMach, i, 5)), x, 7
                        SetText frmInterface.vasRID, Trim(GetText(vasListMach, i, 6)), x, 8
                        SetText frmInterface.vasRID, Trim(GetText(vasListMach, i, 7)), x, 9
                        SetText frmInterface.vasRID, Trim(GetText(vasListMach, i, 8)), x, 10
                        SetText frmInterface.vasRID, Trim(GetText(vasListMach, i, 9)), x, 11
                        SetText frmInterface.vasRID, "Result", x, 14
                        
                    End If
                    
                Next
            End If
                 
        End If
        
      End If
        
    Next
    
End Sub

Private Sub cmdListSch_Click()
    Dim chServer
    Dim iRow As Integer
    Dim sType As String

    chServer = cmbExamType.ListIndex
    ClearSpread vasListMach
    
    Select Case chServer
    Case dpGumjin1
        SQL = "select '','', a.request_date, a.exam_no, b.chart_no, b.person_name, '', '', b.personal_id " & vbCrLf & _
              "from totres a, total b " & vbCrLf & _
              "where a.request_date = b.request_date and a.exam_no = b.exam_no " & vbCrLf & _
              "  and a.request_date between '" & Format(dtpStartDate, "yyyymmdd") & "' and '" & Format(dtpEndDate, "yyyymmdd") & "' and a.result_value = '' " & vbCrLf & _
              "  and a.exam_code in (" & gAllExam & ") " & vbCrLf & _
              "group by a.request_date, a.exam_no, b.person_name, b.personal_id,b.chart_no " & vbCrLf & _
              "order by a.exam_no "
        res = db_select_Vas(gServer, SQL, vasListMach)
    Case dpOCS
        SQL = "select '', '',a.년 + a.월 + a.일, a.챠트번호, a.오더일련번호, b.수진자명, '', '', b.주민등록번호 " & vbCrLf & _
              "from TB_진료검사 a, TB_인적사항 b " & vbCrLf & _
              "where a.챠트번호 = b.챠트번호 " & vbCrLf & _
              "  and a.년 = '" & Format(dtpStartDate, "yyyy") & "' and a.월 = '" & Format(dtpStartDate, "mm") & "' and a.일 = '" & Format(dtpStartDate, "dd") & "' " & vbCrLf & _
              "  and a.검사코드 in (" & gAllExam_Ocs & " ) " & vbCrLf & _
              "  AND A.상태 = '0' AND A.오더일련번호 > '0' " & vbCrLf & _
              "group by a.년 + a.월 + a.일,a.오더일련번호, a.챠트번호, b.수진자명, b.주민등록번호"
        res = db_select_Vas(gServer_OCS, SQL, vasListMach)

    Case dpGumjin2
        SQL = "select '', '',a.request_date, a.exam_no, b.chart_no, b.person_name, '', '', b.personal_id " & vbCrLf & _
              "from twoexam a, total b, panjong2 c " & vbCrLf & _
              "where a.request_date = b.request_date and a.exam_no = b.exam_no " & vbCrLf & _
              "  and and a.request_date = c.request_date and a.exam_no = c.exam_no " & vbCrLf & _
              "  and a.request_date between '" & Format(dtpStartDate, "yyyymmdd") & "' and '" & Format(dtpEndDate, "yyyymmdd") & "' " & vbCrLf & _
              "group by a.request_date, a.exam_no, b.person_name, b.personal_id,b.chart_no " & vbCrLf & _
              "order by a.exam_no"
        res = db_select_Vas(gServer, SQL, vasListMach)
    End Select
    
    For iRow = 1 To vasListMach.DataRowCnt
        CalAgeSex Trim(GetText(vasListMach, iRow, 9)), Trim(Format(Date, "yyyy-mm-dd"))
        SetText vasListMach, gPatGen.Sex, iRow, 7
        SetText vasListMach, gPatGen.Age, iRow, 8
'        SetText vasListMach, chServer, iRow, 10
    Next iRow
    
    vasListMach.RowHeight(-1) = 13
End Sub

Private Sub Form_Load()
    cmbExamType.AddItem "AlLISS", 0
    cmbExamType.AddItem "진료", 1
    'cmbExamType.AddItem "2차검진", 2
    
    cmbExamType.ListIndex = 1
    
    dtpStartDate = Date
    dtpEndDate = Date
    
    ClearSpread vasLExam
    
    SQL = "select examcode, subcode, equipcode, examname,seqno from equipexam where equipno = '" & gEquip & "'"
    res = db_select_Vas(gLocal, SQL, vasLExam)
            
    
End Sub

Private Sub MnfrmListDel_Click()
    Dim i As Long
    Dim vasIDRow As Integer
    Dim vasResRow As Integer
    Dim x As Long
    Dim j As Long
    Dim c, r, c2, r2


    vasResRow = vasListMach.ActiveRow

    If vasListMach.IsBlockSelected Or vasListMach.SelectionCount Then

        vasListMach.BlockMode = True
'        db_BeginTran gLocal
        
        For x = 0 To vasListMach.SelectionCount - 1
            vasListMach.GetSelection x, c, r, c2, r2
            vasListMach.Col = c
            vasListMach.Col2 = c2
            vasListMach.Row = r
            vasListMach.Row2 = r2
            If IsNumeric(r) = True And IsNumeric(r2) = True Then
                If CInt(r) > 0 And CInt(r2) > 0 Then
                    DeleteRow vasListMach, r, r2
                End If
            End If
        Next x
        vasListMach.BlockMode = False

    End If
End Sub

Private Sub vasListMach_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    Dim iRow As Long
    Dim iRow2 As Long
    Dim i As Long
    
    iRow = BlockRow
    iRow2 = BlockRow2
    
    For i = iRow To iRow2
        vasListMach.Col = 1
        vasListMach.Row = i
        If vasListMach.Value = 1 Then
            vasListMach.Value = 0
        Else
            vasListMach.Value = 1
        End If
        
    Next i
End Sub

Private Sub vasListMach_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    Dim sRow As Long
    Dim sCol As Long
    Dim sStartNum As Long
    
    If KeyCode = 13 Then
        sRow = vasListMach.ActiveRow
        sCol = vasListMach.ActiveCol
        If sCol = 2 Then
            If IsNumeric(Trim(GetText(vasListMach, sRow, sCol))) = False Then
                Exit Sub
            End If
            sStartNum = Trim(GetText(vasListMach, sRow, sCol))
            For i = sRow To vasListMach.DataRowCnt
                vasListMach.Col = 1
                vasListMach.Row = i
                If vasListMach.Value = 1 Then
                    SetText vasListMach, sStartNum, i, 2
                    sStartNum = sStartNum + 1
                End If
            Next
            
        End If
        
        
    End If
    
End Sub

Private Sub vasListMach_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    PopupMenu MnfrmList
    
''    Dim i As Long
''    Dim VasidRow As Integer
''    Dim VasResRow As Integer
''    Dim x As Long
''    Dim j As Long
''    Dim c, r, c2, r2
''
''
''    VasResRow = vasListMach.ActiveRow
''
''    If vasListMach.IsBlockSelected Or vasListMach.SelectionCount Then
''
''        vasListMach.BlockMode = True
'''        db_BeginTran gLocal
''
''        For x = 0 To vasListMach.SelectionCount - 1
''            vasListMach.GetSelection x, c, r, c2, r2
''            vasListMach.Col = c
''            vasListMach.Col2 = c2
''            vasListMach.Row = r
''            vasListMach.Row2 = r2
''            If IsNumeric(r) = True And IsNumeric(r2) = True Then
''                If CInt(r) > 0 And CInt(r2) > 0 Then
''                    For j = r To r2
''                        vasListMach.SetText "1", j, "1"
''
''                    Next
''                End If
''            End If
''        Next x
''        vasListMach.BlockMode = False
''
''    End If
End Sub
