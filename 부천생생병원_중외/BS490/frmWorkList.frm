VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmWorkList 
   Caption         =   "워크리스트 조회"
   ClientHeight    =   7395
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9255
   Icon            =   "frmWorkList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   9255
   StartUpPosition =   1  '소유자 가운데
   Begin VB.Timer tmrWork 
      Left            =   4650
      Top             =   120
   End
   Begin VB.CheckBox chkWAll 
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   720
      TabIndex        =   8
      Top             =   990
      Width           =   225
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "닫기"
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
      Left            =   8070
      TabIndex        =   7
      Top             =   120
      Width           =   1035
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
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
      Left            =   7020
      TabIndex        =   6
      Top             =   120
      Width           =   1035
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "워크조회"
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
      Left            =   5970
      TabIndex        =   1
      Top             =   120
      Width           =   1035
   End
   Begin FPSpread.vaSpread vasID 
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   900
      Width           =   8985
      _Version        =   393216
      _ExtentX        =   15849
      _ExtentY        =   11245
      _StockProps     =   64
      ButtonDrawMode  =   4
      ColHeaderDisplay=   0
      ColsFrozen      =   16
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
      GrayAreaBackColor=   16777215
      MaxCols         =   18
      MaxRows         =   20
      MoveActiveOnFocus=   0   'False
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "frmWorkList.frx":030A
   End
   Begin MSComCtl2.DTPicker dtpStopDt 
      Height          =   345
      Left            =   2850
      TabIndex        =   2
      Top             =   180
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   141623297
      CurrentDate     =   40248
   End
   Begin MSComCtl2.DTPicker dtpStartDt 
      Height          =   345
      Left            =   1230
      TabIndex        =   3
      Top             =   180
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   141623297
      CurrentDate     =   40248
   End
   Begin VB.Label lblStatus 
      Height          =   255
      Left            =   150
      TabIndex        =   9
      Top             =   630
      Width           =   8925
   End
   Begin VB.Label Label20 
      Caption         =   "조회일자"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   915
   End
   Begin VB.Label Label12 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2670
      TabIndex        =   4
      Top             =   270
      Width           =   105
   End
End
Attribute VB_Name = "frmWorkList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkWAll_Click()
    Dim iRow As Long
    
    With vasID
        If chkWAll.Value = 1 Then
            For iRow = 1 To .DataRowCnt
                .Row = iRow
                .Col = colCheckBox
                .Value = 1
            Next iRow
        ElseIf chkWAll.Value = 0 Then
            For iRow = 1 To .DataRowCnt
                .Row = iRow
                .Col = colCheckBox
                .Value = 0
            Next iRow
        End If
    End With
    
End Sub

Private Sub cmdClear_Click()
    
    Call frmClear

End Sub

Private Sub cmdClose_Click()
    
    Unload Me

End Sub

Private Sub frmClear()

    vasID.MaxRows = 0
    
    dtpStartDt = Now
    dtpStopDt = Now
    
    lblStatus.Caption = ""


End Sub

Private Sub cmdSearch_Click()
                
    Select Case gOCS
        Case "JWINFO":      Call GetWorkList_JWINFO(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))
    End Select
    
    vasID.RowHeight(-1) = 12

End Sub

Private Sub Form_Load()

    Call frmClear
    
    'Call SetExamCode
    
End Sub

Private Sub SetExamCode()
    Dim i As Integer
    
    
    With vasID
        .MaxCols = colWState + UBound(gArrEquip)
        
        For i = 0 To UBound(gArrEquip) - 1
            .Col = colState + (i + 1)
            .Row = -1
            .CellType = CellTypeEdit
            '.TypeEditCharSet = TypeEditCharSetAlphanumeric
            '.TypeEditCharCase = TypeEditCharCaseSetUpper
            
            .TypeEditCharSet = TypeEditCharSetASCII
            .TypeEditCharCase = TypeEditCharCaseSetNone
            
            .TypeHAlign = TypeHAlignCenter
            .TypeVAlign = TypeVAlignCenter
            Call SetText(vasID, gArrEquip(i + 1, 4), 0, colWState + (i + 1))
            .ColWidth(colWState + (i + 1)) = 6
            
        Next
    End With
    
End Sub

Private Sub GetWorkList_JWINFO(ByVal pFrDt As String, ByVal pToDt As String)
    Dim RS          As ADODB.Recordset
    Dim i           As Integer
    Dim iCnt        As Long
    Dim intRow      As Long
    Dim intCol      As Integer
    Dim strDate     As String
    Dim strChart    As String
    Dim strBarcode    As String
    Dim blnSame     As Boolean
    Dim strER       As String
    
    vasID.MaxRows = 0
    intRow = 0
    
    blnSame = False
    vasID.ReDraw = False
    
    

          SQL = "SELECT DISTINCT RECEIPTDATE as 접수일자, SPECIMENNUM as 바코드번호, RECEIPTNO as 챠트번호, IPDOPD, PTNO as 내원번호, SNAME as 이름, LABCODE as ITEM,ORDERCODE,STAT"
    SQL = SQL & vbCrLf & "  FROM SLA_LabMaster "
    SQL = SQL & vbCrLf & " WHERE RECEIPTDATE between '" & Format(pFrDt, "####-##-##") & "' and '" & Format(pToDt, "####-##-##") & "'"
    SQL = SQL & vbCrLf & "   AND LABCODE IN (" & gAllExam & ") "
    SQL = SQL & vbCrLf & "   AND JSTATUS < '3'" & vbLf
    SQL = SQL & "  ORDER BY RECEIPTDATE "
    
    Call SetSQLData("워크조회", SQL)

    '-- Record Count 가져옴
    cn_Ser.CursorLocation = adUseClient
    Set RS = cn_Ser.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then
        'frmProgress.Show
        'frmProgress.ZOrder 0
        'frmProgress.Xprog.Min = 1
        'frmProgress.Xprog.Max = RS.RecordCount + 1
        
        Do Until RS.EOF
            iCnt = iCnt + 1
            With vasID
                .ReDraw = False
                For i = 1 To .DataRowCnt
                    strDate = GetText(vasID, i, colWHOSPDATE)
                    strBarcode = GetText(vasID, i, colWBARCODE)
                    If Trim(RS("접수일자")) = strDate And Trim(RS("바코드번호")) = strBarcode Then
                        blnSame = True
                    End If
                    
'                    For intCol = colState + 1 To vasID.MaxCols
'                        If Trim(RS.Fields("ITEM")) = gArrEquip(intCol - colState, 3) Then
'                            vasID.Row = .MaxRows
'                            vasID.Col = intCol
'                            vasID.BackColor = vbYellow
'                            Exit For
'                        End If
'                    Next
                Next
                
                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    SetText vasID, "1", .MaxRows, colWCheckBox
                    
                    '- Stat : 0.일반,  1.응급,  9.외부검사
                    strER = Trim(RS.Fields("STAT")) & ""
                    
                    If strER = "1" Then
                        SetText vasID, "1", .MaxRows, colWER
                    Else
                        SetText vasID, "0", .MaxRows, colWER
                    End If
                    
                    SetText vasID, Trim(RS.Fields("접수일자")) & "", .MaxRows, colWHOSPDATE
                    SetText vasID, Trim(RS.Fields("바코드번호")) & "", .MaxRows, colWBARCODE
                    SetText vasID, Trim(RS.Fields("챠트번호")) & "", .MaxRows, colWCHARTNO
                    SetText vasID, Trim(RS.Fields("내원번호")) & "", .MaxRows, colWPID
                    SetText vasID, Trim(RS.Fields("이름")) & "", .MaxRows, colWPNAME
                    SetText vasID, IIf(Trim(RS.Fields("IPDOPD")) = 1, "입원", "외래"), .MaxRows, colWINOUT
                    SetText vasID, Trim(RS.Fields("ORDERCODE")) & "", .MaxRows, colWPSEX
                  
'                    For intCol = colState + 1 To vasID.MaxCols
'                        If Trim(RS.Fields("ITEM")) = gArrEquip(intCol - colWState, 3) Then
'                            vasID.Row = .MaxRows
'                            vasID.Col = intCol
'                            vasID.BackColor = vbYellow
'                            Exit For
'                        End If
'                    Next
                
                End If
                
                blnSame = False
            End With
            '-- 프로그레스바 진행
            frmProgress.Xprog.Value = iCnt
            DoEvents
                        
            RS.MoveNext
        Loop
        chkWAll.Value = "1"
    Else
        lblStatus.Caption = "조회 대상자가 없습니다."
        chkWAll.Value = "0"
    End If
    
    RS.Close
    
    vasID.RowHeight(-1) = 12
    vasID.ReDraw = True
    
    Screen.MousePointer = 0
    
End Sub

