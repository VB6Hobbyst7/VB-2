VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmExamSearch 
   Caption         =   "Abnormal 검체조회"
   ClientHeight    =   9060
   ClientLeft      =   4410
   ClientTop       =   1290
   ClientWidth     =   8880
   LinkTopic       =   "Form3"
   ScaleHeight     =   9060
   ScaleWidth      =   8880
   Begin VB.Frame Frame1 
      Height          =   675
      Left            =   90
      TabIndex        =   1
      Top             =   -30
      Width           =   8685
      Begin VB.CommandButton cmdClose 
         Caption         =   "닫기"
         Height          =   405
         Left            =   7590
         TabIndex        =   10
         Top             =   180
         Width           =   1005
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "출력"
         Height          =   405
         Left            =   6300
         TabIndex        =   9
         Top             =   180
         Width           =   1005
      End
      Begin VB.TextBox txtReceNo2 
         Height          =   285
         Left            =   4320
         TabIndex        =   7
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtReceNo1 
         Height          =   285
         Left            =   3360
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
      Begin MSComCtl2.DTPicker dtpReceDate 
         Height          =   285
         Left            =   990
         TabIndex        =   4
         Top             =   240
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   503
         _Version        =   393216
         Format          =   50855937
         CurrentDate     =   42096
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "검체조회"
         Height          =   405
         Left            =   5130
         TabIndex        =   2
         Top             =   180
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "-"
         Height          =   180
         Left            =   4170
         TabIndex        =   8
         Top             =   300
         Width           =   90
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "접수번호"
         Height          =   180
         Left            =   2520
         TabIndex        =   5
         Top             =   300
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "접수일자"
         Height          =   180
         Left            =   150
         TabIndex        =   3
         Top             =   300
         Width           =   720
      End
   End
   Begin FPSpread.vaSpread vasList 
      Height          =   8325
      Left            =   90
      TabIndex        =   0
      Top             =   690
      Width           =   8715
      _Version        =   393216
      _ExtentX        =   15372
      _ExtentY        =   14684
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      ColHeaderDisplay=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GridColor       =   16777215
      MaxCols         =   104
      Protect         =   0   'False
      ScrollBars      =   2
      SpreadDesigner  =   "frmExamSearch.frx":0000
   End
End
Attribute VB_Name = "frmExamSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload frmExamSearch
End Sub

Private Sub cmdPrint_Click()
    Dim strHeader As String
    
    If vasList.DataRowCnt < 1 Then
        MsgBox "출력할 리스트가 없습니다."
        Exit Sub
    End If
    
    
    strHeader = "            "
    
    vasList.PrintHeader = "/n/n/fb1/fz""15""" & strHeader & _
                           "=== 접수일자 - " & Format(dtpReceDate, "yyyy-mm-dd") & " 이상 결과 목록 ===/n/n"
    
    vasList.PrintMarginTop = 1000
    vasList.PrintMarginLeft = 1000
    'vasList.Action = ActionPrint
    vasList.PrintSheet
    cmdClose_Click
End Sub

Private Sub cmdSearch_Click()
'601과 701을 분리 하여 검사 했을때 이상 값일 경우 micro  검사를 진행하지 않게 한다.
    Dim strRece1 As String
    Dim strRece2 As String
    Dim strReceAll As String
    Dim i As Integer
    
    ClearSpread vasList
    If Trim(txtReceNo1) <> "" And IsNumeric(txtReceNo1) = True Then
        strRece1 = txtReceNo1
    ElseIf Trim(txtReceNo1) = "" Or IsNumeric(txtReceNo1) = False Then
        strRece1 = "0"
    End If
    
    If Trim(txtReceNo2) <> "" And IsNumeric(txtReceNo2) = True Then
        strRece2 = txtReceNo2
    ElseIf Trim(strRece2) = "" Or IsNumeric(txtReceNo2) = False Then
        strRece2 = "99999"
    End If
    
    SQL = ""
    SQL = SQL & vbCrLf & "select distinct labregno"
    SQL = SQL & vbCrLf & "From LabRegResult"
    SQL = SQL & vbCrLf & "where LabRegDate = '" & Format(dtpReceDate, "yyyy-mm-dd") & "'"
    SQL = SQL & vbCrLf & "  and labregno between " & strRece1 & " and " & strRece2 & ""
'''    SQL = SQL & vbCrLf & "  and machinecode in ('CM049', 'CM048')"
    SQL = SQL & vbCrLf & "  and testsubcode in (" & gAllExam_Micro & ")"
    SQL = SQL & vbCrLf & "  and (testresult01 is null or testresult01 = '')"  '결과 전송여부 체크 해야함.
    res = db_select_Row(gServer, SQL)
    
    strReceAll = ""
    
    For i = 1 To res
        If i = 1 Then
            strReceAll = "'" & Trim(gReadBuf(i - 1)) & "'"
        Else
            strReceAll = strReceAll & ", '" & Trim(gReadBuf(i - 1)) & "'"
        End If
    Next
    
    If strReceAll = "" Then
        strReceAll = "''"
    End If
    
    SQL = ""
    SQL = SQL & vbCrLf & "SELECT '', barcode, posno, pid, MID(BARCODE, 7, 5) , pname, PSEX, ' ' "
    SQL = SQL & vbCrLf & "  FROM PAT_RES"
    SQL = SQL & vbCrLf & " WHERE MID(BARCODE, 1, 6) = '" & Format(dtpReceDate, "yymmdd") & "'"
    SQL = SQL & vbCrLf & "   AND EQUIPCODE IN ('ERY', 'LEU') "
    SQL = SQL & vbCrLf & "   AND ISNUMERIC(MID(BARCODE, 7, 5)) = TRUE "
'''    SQL = SQL & vbCrLf & "   AND INT(RECENO) BETWEEN " & strRece1 & " AND " & strRece2 & ""
    SQL = SQL & vbCrLf & "   AND INT(MID(BARCODE, 7, 5)) IN (" & strReceAll & ")"
    SQL = SQL & vbCrLf & "   AND RESVALUE IN ('1+','2+','3+','4+','5+','Trace') "
'''    SQL = SQL & vbCrLf & " UNION ALL"
'''    SQL = SQL & vbCrLf & "SELECT '', barcode, posno, pid, receno, pname, PSEX, ' ' "
'''    SQL = SQL & vbCrLf & "  FROM PAT_RES"
'''    SQL = SQL & vbCrLf & " WHERE MID(BARCODE, 1, 6) = '" & Format(dtpReceDate, "yymmdd") & "'"
'''    SQL = SQL & vbCrLf & "   AND EQUIPCODE = 'LEU' "
'''    SQL = SQL & vbCrLf & "   AND RECENO <> '' "
''''''    SQL = SQL & vbCrLf & "   AND INT(RECENO) BETWEEN " & strRece1 & " AND " & strRece2 & ""
'''    SQL = SQL & vbCrLf & "   AND INT(RECENO) IN (" & strReceAll & ")"
'''
'''    SQL = SQL & vbCrLf & "   AND RESVALUE IN ('1+','2+','3+','4+','5+','Trace') "
    
    SQL = SQL & vbCrLf & "   GROUP BY barcode, posno, pid, MID(BARCODE, 7, 5), pname, PSEX "
    SQL = SQL & vbCrLf & "   order by barcode"
    
    res = db_select_Vas(gLocal, SQL, vasList)
    
End Sub

Private Sub Form_Load()
    ClearSpread vasList
    dtpReceDate = Date
End Sub
