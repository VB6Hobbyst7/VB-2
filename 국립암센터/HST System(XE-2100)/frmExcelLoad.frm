VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#3.5#0"; "SPR32X35.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExcelLoad 
   Caption         =   "Form1"
   ClientHeight    =   9075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13740
   LinkTopic       =   "Form1"
   ScaleHeight     =   9075
   ScaleWidth      =   13740
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton cmdDataSave 
      BackColor       =   &H8000000E&
      Caption         =   "저장"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   9990
      MaskColor       =   &H8000000F&
      Style           =   1  '그래픽
      TabIndex        =   14
      Top             =   90
      Width           =   1275
   End
   Begin VB.CommandButton cmdDataProc 
      BackColor       =   &H8000000E&
      Caption         =   "데이타처리"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   8700
      MaskColor       =   &H8000000F&
      Style           =   1  '그래픽
      TabIndex        =   13
      Top             =   90
      Width           =   1275
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6900
      Top             =   870
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin FPSpread.vaSpread vasList 
      Height          =   7575
      Left            =   120
      TabIndex        =   11
      Top             =   1350
      Width           =   7845
      _Version        =   196613
      _ExtentX        =   13838
      _ExtentY        =   13361
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
      SpreadDesigner  =   "frmExcelLoad.frx":0000
   End
   Begin VB.CommandButton cmdLoad 
      BackColor       =   &H8000000E&
      Caption         =   "가져오기"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   7410
      MaskColor       =   &H8000000F&
      Style           =   1  '그래픽
      TabIndex        =   10
      Top             =   90
      Width           =   1275
   End
   Begin VB.TextBox txtSheet 
      Height          =   315
      Left            =   1590
      TabIndex        =   5
      Text            =   "SCL_결과샘플"
      Top             =   450
      Width           =   5565
   End
   Begin VB.TextBox txtFile 
      BackColor       =   &H00FEE7F3&
      Height          =   315
      Left            =   1590
      TabIndex        =   4
      Top             =   60
      Width           =   5565
   End
   Begin VB.TextBox txtRow 
      Height          =   315
      Left            =   2280
      TabIndex        =   3
      Text            =   "2"
      Top             =   810
      Width           =   675
   End
   Begin VB.TextBox txtCol 
      Height          =   315
      Left            =   1590
      TabIndex        =   2
      Text            =   "1"
      Top             =   810
      Width           =   675
   End
   Begin VB.TextBox txtRow2 
      Height          =   315
      Left            =   5970
      TabIndex        =   1
      Text            =   "200"
      Top             =   810
      Width           =   675
   End
   Begin VB.TextBox txtCol2 
      Height          =   315
      Left            =   5280
      TabIndex        =   0
      Text            =   "50"
      Top             =   810
      Width           =   675
   End
   Begin FPSpread.vaSpread vasExam 
      Height          =   7575
      Left            =   8100
      TabIndex        =   12
      Top             =   1350
      Width           =   5475
      _Version        =   196613
      _ExtentX        =   9657
      _ExtentY        =   13361
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
      SpreadDesigner  =   "frmExcelLoad.frx":445F
   End
   Begin MSComCtl2.DTPicker dtpExamDate 
      Height          =   345
      Left            =   8550
      TabIndex        =   15
      Top             =   780
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   24313857
      CurrentDate     =   38584
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "검사일자"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7530
      TabIndex        =   16
      Top             =   855
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "ExcelFileName"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   180
      TabIndex        =   9
      Top             =   120
      Width           =   1365
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "선택 Sheet  명"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   180
      TabIndex        =   8
      Top             =   510
      Width           =   1305
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "시작 Col && Row"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   150
      TabIndex        =   7
      Top             =   870
      Width           =   1395
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "마지막 Col && Row"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3660
      TabIndex        =   6
      Top             =   870
      Width           =   1590
   End
End
Attribute VB_Name = "frmExcelLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

End Sub

Private Sub cmdDataProc_Click()
    Dim lRow1, lRow2, lCol As Long
    Dim lsData As String
    Dim lsEquipCode, lsResult, lsFlag As String
    Dim i As Integer
    
    ClearSpread vasExam
    
    For lRow1 = 1 To vasList.DataRowCnt
        lsData = ""
        lsEquipCode = ""
        lsResult = ""
        
        For lCol = 18 To 52
            vasActiveCell vasList, lRow1, lCol
            
            lsData = Trim(GetText(vasList, lRow1, lCol))
            i = InStr(1, lsData, "|")
            If i > 0 Then
                lsEquipCode = Left(lsData, i - 1)
                If IsNumeric(lsEquipCode) Then
                    lsEquipCode = Format(CInt(lsEquipCode), "00")
                End If
                lsData = Mid(lsData, i + 1)
            
                i = InStr(1, lsData, "|")
                If i > 0 Then
                    lsResult = Left(lsData, i - 1)
                    lsData = Mid(lsData, i + 1)
                    
                    i = InStr(1, lsData, "|")
                    If i > 0 Then
                        lsFlag = Left(lsData, i - 1)
                    Else
                        lsFlag = Trim(lsData)
                    End If
                    If lsFlag = "Low" Then
                        lsFlag = "L"
                    ElseIf lsFlag = "High" Then
                        lsFlag = "H"
                    Else
                        lsFlag = ""
                    End If
                End If
            End If
            If lsEquipCode <> "" And lsResult <> "" Then
                lRow2 = lRow2 + 1
                If vasExam.MaxRows < lRow2 Then
                    vasExam.MaxRows = lRow2
                End If
                
                vasExam.SetText 1, lRow2, Trim(GetText(vasList, lRow1, 6))  'ID
                vasExam.SetText 2, lRow2, Trim(GetText(vasList, lRow1, 7))  'Rack
                vasExam.SetText 3, lRow2, Trim(GetText(vasList, lRow1, 8))  'Pos
                vasExam.SetText 4, lRow2, Trim(GetText(vasList, lRow1, 9))  'PID
                vasExam.SetText 5, lRow2, Trim(GetText(vasList, lRow1, 10))  'PName
                vasExam.SetText 6, lRow2, Trim(GetText(vasList, lRow1, 12))  'Sex
                vasExam.SetText 7, lRow2, Trim(GetText(vasList, lRow1, 13))  'Age
                vasExam.SetText 8, lRow2, lsEquipCode  'Equip
                vasExam.SetText 9, lRow2, lsResult 'result
                vasExam.SetText 10, lRow2, lsFlag  'flag
                SQL = "select examname from equipexam where equip = '" & gEquip & "' and equipcode = '" & lsEquipCode & "'"
                res = db_select_Col(gLocal, SQL)
                vasExam.SetText 11, lRow2, Trim(gReadBuf(0))  'examname
            End If
        Next lCol
    Next lRow1

End Sub

Private Sub cmdDataSave_Click()
    Dim lRow As Double
    Dim sCnt As String
    Dim sExamDate As String
    
    sExamDate = GetDateFull
    
    For lRow = 1 To vasExam.DataRowCnt
        
    '    SQL = "select examno from pat_res "
    '    res = db_select_Col(gLocal, SQL)
    '    If res = -1 Then
    '        SQL = "Alter table pat_res add examno varchar(10) "
    '        res = SendQuery(gLocal, SQL)
    '    End If
        
        sCnt = ""
        SQL = "Delete FROM pat_res " & vbCrLf & _
              "WHERE examdate = '" & Format(dtpExamDate.Value, "yyyymmdd") & "' " & vbCrLf & _
              "  AND equipno = '" & gEquip & "' " & vbCrLf & _
              "  AND equipcode = '" & CInt(CStr(GetText(vasExam, lRow, 8))) & "'" & vbCrLf & _
              "  AND barcode = '" & Trim(GetText(vasExam, lRow, 1)) & "' " & vbCrLf & _
              "  and examtype = 'XE' "
        res = SendQuery(gLocal, SQL)
        If res = -1 Then
            SaveQuery SQL
            Exit Sub
        End If
'
'                vasExam.SetText 1, lRow2, Trim(GetText(vasList, lRow1, 6))  'ID
'                vasExam.SetText 2, lRow2, Trim(GetText(vasList, lRow1, 7))  'Rack
'                vasExam.SetText 3, lRow2, Trim(GetText(vasList, lRow1, 8))  'Pos
'                vasExam.SetText 4, lRow2, Trim(GetText(vasList, lRow1, 9))  'PID
'                vasExam.SetText 5, lRow2, Trim(GetText(vasList, lRow1, 10))  'PName
'                vasExam.SetText 6, lRow2, Trim(GetText(vasList, lRow1, 12))  'Sex
'                vasExam.SetText 7, lRow2, Trim(GetText(vasList, lRow1, 13))  'Age
'                vasExam.SetText 8, lRow2, lsEquipCode  'Equip
'                vasExam.SetText 9, lRow2, lsResult 'result
'                vasExam.SetText 10, lRow2, lsFlag  'flag
'                SQL = "select examname from equipexam where equip = '" & gEquip & "' and equipcode = '" & lsEquipCode & "'"
'                res = db_select_Col(gLocal, SQL)
'                vasExam.SetText 11, lRow2, Trim(gReadBuf(0))  'examname
        If Not IsNumeric(GetText(vasExam, lRow, 7)) Then
            vasExam.SetText 7, lRow, "0"
        End If
        SQL = "INSERT INTO pat_res (examdate, equipno, " & _
                "barcode, examtype, " & _
                "receno, pid, " & _
                "pname, pjumin, page, " & _
                "psex, page1, " & _
                "WardRoom, resdate, seqno, " & _
                "diskno, posno, " & _
                "equipcode, examcode, examno, " & _
                "result, sendflag, examname, " & _
                "refflag,panicflag, deltaflag, unit, refvalue, panicvalue ) " & vbCrLf & _
              "VALUES ('" & Format(dtpExamDate.Value, "yyyymmdd") & "', '" & Trim(gEquip) & "', " & _
              "'" & Trim(GetText(vasExam, lRow, 1)) & "','XE','', " & _
              "'" & Trim(GetText(vasExam, lRow, 4)) & "', '" & GetText(vasExam, lRow, 5) & "', '', " & Trim(GetText(vasExam, lRow, 7)) & "," & _
              "'" & Trim(GetText(vasExam, lRow, 6)) & "', '" & Trim(GetText(vasExam, lRow, 7)) & "'," & _
              "'', '" & sExamDate & "', '" & Trim(GetText(vasExam, lRow, 8)) & "', " & _
              "'" & Trim(GetText(vasExam, lRow, 2)) & "', '" & Trim(GetText(vasExam, lRow, 3)) & "', " & vbCrLf & _
              "'" & Trim(GetText(vasExam, lRow, 8)) & "', '', '', " & _
              "'" & Trim(GetText(vasExam, lRow, 9)) & "', 'B', '" & Trim(GetText(vasExam, lRow, 11)) & "', " & vbCrLf & _
              "'" & Trim(GetText(vasExam, lRow, 10)) & "', '', '', '', " & _
              "'', '' ) "
        res = SendQuery(gLocal, SQL)
        If res = -1 Then
            SaveQuery SQL
            Exit Sub
        End If
        
    Next lRow
End Sub

Private Sub cmdLoad_Click()
    Dim xl As New Excel.Application
    Dim xlw As Excel.Workbook
    
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    If Trim(txtFile.Text) = "" Then
        txtFile.SetFocus
        Exit Sub
    End If
    
    Me.MousePointer = 11
    
    vasList.MaxRows = CLng(txtRow2)
    
    ClearSpread vasList
    
    ' 해당 Excel 파일을 연다.
    Set xlw = xl.Workbooks.Open(Trim(txtFile.Text))
    
    ' 가져올 데이터를 포함하고있는 Excel Sheet 를 선택한다.
    xlw.Sheets(Trim(txtSheet.Text)).Select

    i = 1
    j = 1
    k = 1
    
    For i = CLng(txtRow) To CLng(txtRow2)
        For j = CLng(txtCol) To CLng(txtCol2)
            SetText vasList, xlw.Application.Cells(i, j).Value, k, j
        Next j

        k = k + 1
        
    Next i
          
    ' Close worksheet without save changes.
    xlw.Close False
    
    Set xlw = Nothing
    Set xl = Nothing
    
    vasList.MaxRows = vasList.DataRowCnt
    Me.MousePointer = 0
End Sub

Private Sub Form_Load()
    dtpExamDate.Value = Format(CDate(GetDateFull), "yyyy-mm-dd")
End Sub

Private Sub txtFile_DblClick()
    Dim sTmp As String
    Dim iPos As Integer

    CommonDialog1.Filter = "Excel Files (*.xls)|*.xls|All (*.*)|*.*"
    CommonDialog1.ShowOpen
    txtFile.Text = CommonDialog1.FileName
    
    '2004/09/22 이상은 수정=================
    sTmp = Dir(txtFile.Text, vbDirectory)
    iPos = InStr(1, sTmp, ".")
    txtSheet.Text = Mid(sTmp, 1, iPos - 1)
    '=======================================
    
    txtSheet.SetFocus
End Sub
