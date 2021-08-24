VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEQ_워크리스트 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "WORK LIST"
   ClientHeight    =   7005
   ClientLeft      =   4245
   ClientTop       =   2070
   ClientWidth     =   10635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   10635
   Begin VB.CommandButton cmdClear 
      Caption         =   "초기화(&C)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9360
      Style           =   1  '그래픽
      TabIndex        =   17
      Top             =   840
      Width           =   1155
   End
   Begin VB.CommandButton cmdListDel 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6660
      TabIndex        =   13
      Top             =   5400
      Width           =   735
   End
   Begin VB.CommandButton cmdListAdd 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6660
      TabIndex        =   12
      Top             =   2460
      Width           =   735
   End
   Begin VB.CommandButton cmdOrder 
      BackColor       =   &H00FFC0C0&
      Caption         =   "ORDER  전송(&O)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7560
      Style           =   1  '그래픽
      TabIndex        =   11
      Top             =   60
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.CommandButton cmdWorkList 
      BackColor       =   &H00FFC0C0&
      Caption         =   "LIST  전송(&L)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8580
      Style           =   1  '그래픽
      TabIndex        =   10
      Top             =   60
      Width           =   915
   End
   Begin VB.Frame Frame1 
      Caption         =   "[접수일자]"
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   6255
      Begin VB.CheckBox chkResultX 
         Caption         =   "결과 X"
         Height          =   195
         Left            =   5160
         TabIndex        =   16
         Top             =   480
         Width           =   855
      End
      Begin VB.CheckBox chkResultO 
         Caption         =   "결과 O"
         Height          =   195
         Left            =   5160
         TabIndex        =   15
         Top             =   180
         Width           =   855
      End
      Begin VB.CommandButton cmdView 
         Caption         =   "조회(&V)"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3840
         Style           =   1  '그래픽
         TabIndex        =   9
         Top             =   180
         Width           =   975
      End
      Begin MSComCtl2.DTPicker dtpDateFrom 
         Height          =   315
         Left            =   840
         TabIndex        =   5
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   21430273
         CurrentDate     =   40820
      End
      Begin MSComCtl2.DTPicker dtpDateTo 
         Height          =   315
         Left            =   2400
         TabIndex        =   6
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   21430273
         CurrentDate     =   40820
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "~"
         Height          =   180
         Index           =   8
         Left            =   2220
         TabIndex        =   8
         Top             =   300
         Width           =   90
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '투명
         Caption         =   "기간"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   7
         Top             =   300
         Width           =   435
      End
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "닫기(&Q)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9600
      Style           =   1  '그래픽
      TabIndex        =   3
      Top             =   60
      Width           =   915
   End
   Begin FPSpread.vaSpread sprWorkList 
      Height          =   5295
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   6255
      _Version        =   393216
      _ExtentX        =   11033
      _ExtentY        =   9340
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
      MaxCols         =   11
      OperationMode   =   2
      SelectBlockOptions=   0
      SpreadDesigner  =   "frmEQ_워크리스트.frx":0000
   End
   Begin MSComctlLib.ProgressBar barStatus 
      Height          =   75
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   132
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin FPSpread.vaSpread sprWorkList_EXAM 
      Height          =   5295
      Left            =   7560
      TabIndex        =   14
      Top             =   1560
      Width           =   2955
      _Version        =   393216
      _ExtentX        =   5212
      _ExtentY        =   9340
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
      MaxCols         =   4
      OperationMode   =   2
      SelectBlockOptions=   0
      SpreadDesigner  =   "frmEQ_워크리스트.frx":1C00
   End
   Begin VB.Label lbl워크리스트 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "WORK LIST"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   240
      TabIndex        =   2
      Top             =   60
      Width           =   2100
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808000&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00000000&
      FillColor       =   &H0000C000&
      FillStyle       =   5  '하향 대각선
      Height          =   495
      Index           =   3
      Left            =   120
      Shape           =   4  '둥근 사각형
      Top             =   60
      Width           =   3555
   End
End
Attribute VB_Name = "frmEQ_워크리스트"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClear_Click()
    sprWorkList_EXAM.MaxRows = 0
    
End Sub

Private Sub cmdListAdd_Click()
    For intX = 1 To sprWorkList.MaxRows
        sprWorkList.Col = 1
        sprWorkList.Row = intX
        If sprWorkList.Value = 1 Then
            sprWorkList_EXAM.MaxRows = sprWorkList_EXAM.MaxRows + 1
            Call SET_CELL(sprWorkList_EXAM, 1, sprWorkList_EXAM.MaxRows, Trim(GET_CELL(sprWorkList, 2, intX)))
            Call SET_CELL(sprWorkList_EXAM, 2, sprWorkList_EXAM.MaxRows, Trim(GET_CELL(sprWorkList, 3, intX)))
            Call SET_CELL(sprWorkList_EXAM, 3, sprWorkList_EXAM.MaxRows, Trim(GET_CELL(sprWorkList, 4, intX)))
            Call SET_CELL(sprWorkList_EXAM, 4, sprWorkList_EXAM.MaxRows, Trim(GET_CELL(sprWorkList, 5, intX)))
            sprWorkList.Col = 1
            sprWorkList.Row = intX
            sprWorkList.Value = 0
        End If
    Next intX
End Sub

Private Sub cmdListDel_Click()
    Dim intDelRow   As Integer
    If sprWorkList_EXAM.ActiveRow = 0 Then Exit Sub
    intDelRow = sprWorkList_EXAM.ActiveRow
    Call sprWorkList_EXAM.DeleteRows(intDelRow, 1)
    sprWorkList_EXAM.MaxRows = sprWorkList_EXAM.MaxRows - 1
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub cmdView_Click()
    Dim intLResultTarRow    As Integer
    
    sprWorkList.MaxRows = 0
    
    Call FUNC_GET_EXCD
    
    If ConnDB_HIS = False Then Exit Sub
                   gstrQuy = "SELECT /*+ INDEX_DESC(EXAMRES EXAMRES_INDEX_5) */ "
        gstrQuy = gstrQuy & vbCrLf & "       RES.RECENO, RES.SPECIMENID, PATT.PNAME,"
        gstrQuy = gstrQuy & vbCrLf & "       REQ.PID, MAX(REQ.EMGFLAG) AS EMGFLAG, (CASE WHEN MAX(REQ.EMGFLAG) = 'Y' THEN '응급' END) AS EMGSTRING ,"
        gstrQuy = gstrQuy & vbCrLf & "       REQ.IOFLAG, (CASE WHEN MAX(REQ.IOFLAG) = 'I' THEN '입원' WHEN MAX(REQ.IOFLAG) = 'O' THEN '외래' END) AS IOSTRING "
        gstrQuy = gstrQuy & vbCrLf & "      ,(WARD.WARDNAME || ' ' || WARD.ROOM) AS WARDROOM "

        gstrQuy = gstrQuy & vbCrLf & "  FROM EXAMRES RES,"
        gstrQuy = gstrQuy & vbCrLf & "       EXAMREQ REQ,"
        gstrQuy = gstrQuy & vbCrLf & "       PATIENT PATT,"
        gstrQuy = gstrQuy & vbCrLf & "       WARD    WARD"

        gstrQuy = gstrQuy & vbCrLf & " WHERE SUBSTR(REQ.RECENO,1,8) BETWEEN '" & Format(dtpDateFrom, "yyyymmdd") & "' "
        gstrQuy = gstrQuy & vbCrLf & "                                  AND '" & Format(dtpDateTo, "yyyymmdd") & "'"
        gstrQuy = gstrQuy & vbCrLf & "   AND RES.HID = REQ.HID"
        gstrQuy = gstrQuy & vbCrLf & "   AND REQ.PID = RES.PID"
        gstrQuy = gstrQuy & vbCrLf & "   AND PATT.PID = RES.PID"
        gstrQuy = gstrQuy & vbCrLf & "   AND WARD.WARDCODE(+)= REQ.WARDCODE"
        gstrQuy = gstrQuy & vbCrLf & "   AND WARD.ROOM(+)= REQ.ROOM"
        
        gstrQuy = gstrQuy & vbCrLf & "   AND REQ.RECENO = RES.RECENO"
        gstrQuy = gstrQuy & vbCrLf & "   AND RES.LABRECYN = 'Y' "
        
        gstrQuy = gstrQuy & vbCrLf & "   AND RES.EXAMCODE IN (" & gstrEQORDREAD & ") "

        gstrQuy = gstrQuy & vbCrLf & "   AND (NVL(RES.RESEND, ' ') NOT IN ('1','2') "  '보고대상
        gstrQuy = gstrQuy & vbCrLf & "       OR RES.EXAMSTATE = 'E') "
        gstrQuy = gstrQuy & vbCrLf & "   AND RES.EXAMSTATE <> 'Q' " '서북병원 과장CONFIRM
        
        If chkResultO.Value = 1 And chkResultX.Value = 0 Then
            gstrQuy = gstrQuy & vbCrLf & "   AND NVL(RES.RESULT, ' ') <> ' ' "
        ElseIf chkResultO.Value = 0 And chkResultX.Value = 1 Then
            gstrQuy = gstrQuy & vbCrLf & "   AND NVL(RES.RESULT, ' ') = ' ' "
        End If
           
        gstrQuy = gstrQuy & vbCrLf & " GROUP BY RES.RECENO,"
        gstrQuy = gstrQuy & vbCrLf & "          RES.SPECIMENID,"
        gstrQuy = gstrQuy & vbCrLf & "          PATT.PNAME,"
        gstrQuy = gstrQuy & vbCrLf & "          REQ.PID,"
        gstrQuy = gstrQuy & vbCrLf & "          REQ.SEQNO,"
        gstrQuy = gstrQuy & vbCrLf & "          REQ.IOFLAG,"
        gstrQuy = gstrQuy & vbCrLf & "          SUBSTR(REQ.RECENO,10,4), "
        gstrQuy = gstrQuy & vbCrLf & "         (WARD.WARDNAME || ' ' || WARD.ROOM)"

        gstrQuy = gstrQuy & vbCrLf & " ORDER BY RES.RECENO"
    
'                           gstrQuy = "SELECT A.PID, B.SPECIMENID   "
'        gstrQuy = gstrQuy & vbCrLf & "  FROM PATIENT A, EXAMRES B "
'        gstrQuy = gstrQuy & vbCrLf & " WHERE A.PID = B.PID "
'        gstrQuy = gstrQuy & vbCrLf & "   AND B.EXAMCODE IN (" & gtypPAT_RES.BARCD & ") "
'        gstrQuy = gstrQuy & vbCrLf & "   AND B.EXAMCODE IN (" & gtypPAT_RES.BARCD & ") "
    If ReadSQL_HIS(gstrQuy, ADR_HIS) = False Then Call CloseDB_HIS: End

    If Not ADR_HIS Is Nothing Then
        
        Do Until ADR_HIS.EOF
            sprWorkList.MaxRows = sprWorkList.MaxRows + 1
            intLResultTarRow = sprWorkList.MaxRows
            
            Call SET_CELL(sprWorkList, 2, intLResultTarRow, Trim(ADR_HIS!RECENO & ""))
            Call SET_CELL(sprWorkList, 3, intLResultTarRow, Trim(ADR_HIS!SPECIMENID & ""))
            Call SET_CELL(sprWorkList, 4, intLResultTarRow, Trim(ADR_HIS!PNAME & ""))
            Call SET_CELL(sprWorkList, 5, intLResultTarRow, Trim(ADR_HIS!PID & ""))
            Call SET_CELL(sprWorkList, 6, intLResultTarRow, Trim(ADR_HIS!WARDROOM & ""))
            Call SET_CELL(sprWorkList, 7, intLResultTarRow, Trim(ADR_HIS!EMGFLAG & ""))
            Call SET_CELL(sprWorkList, 8, intLResultTarRow, Trim(ADR_HIS!EMGSTRING & ""))
            Call SET_CELL(sprWorkList, 9, intLResultTarRow, Trim(ADR_HIS!IOFLAG & ""))
            Call SET_CELL(sprWorkList, 10, intLResultTarRow, Trim(ADR_HIS!IOSTRING & ""))
            ADR_HIS.MoveNext
        Loop
        
        ADR_HIS.Close: Set ADR_HIS = Nothing

    End If
    
    Call CloseDB_HIS
End Sub

Private Sub cmdWorkList_Click()
    Dim BFSEQ   As String
    
With frmEQ_Main
    For intX = 1 To sprWorkList_EXAM.DataRowCnt
        For intY = 1 To .sprLResult.DataRowCnt
            If GET_CELL(.sprLResult, 1, intY) = "No Barcode" Then
                BFSEQ = GET_CELL(.sprLResult, 3, intY)
                Call SET_CELL(.sprLResult, 1, intY, Trim(GET_CELL(sprWorkList_EXAM, 2, intX)))
                Call NOVA_SAVE(GET_CELL(.sprLResult, 1, intY), intY)
                Call FUNC_LOC_DELETE_PAT_RES("No Barcode", BFSEQ)
                Exit For
            End If
        Next intY
    Next intX
End With

End Sub

Private Sub Form_Load()
    sprWorkList.MaxRows = 0
    sprWorkList_EXAM.MaxRows = 0
    barStatus.Max = 100
    barStatus.Value = 100
    dtpDateFrom = Format(Now, "yyyy/mm/dd")
    dtpDateTo = Format(Now, "yyyy/mm/dd")
End Sub

Private Sub sprWorkList_Click(ByVal Col As Long, ByVal Row As Long)
    sprWorkList.Col = 1
    sprWorkList.Row = Row
    If Row = 0 Then Exit Sub
    If sprWorkList.Value = 1 Then
         sprWorkList.Value = 0
    ElseIf sprWorkList.Value = "0" Then
        sprWorkList.Value = 1
    End If
End Sub

