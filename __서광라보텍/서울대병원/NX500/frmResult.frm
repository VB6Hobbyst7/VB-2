VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#3.5#0"; "SPR32X35.ocx"
Begin VB.Form frmResult 
   Caption         =   "결과 확인"
   ClientHeight    =   9195
   ClientLeft      =   120
   ClientTop       =   585
   ClientWidth     =   14910
   LinkTopic       =   "Form1"
   ScaleHeight     =   9195
   ScaleWidth      =   14910
   Begin VB.Frame Frame4 
      Height          =   7575
      Left            =   10140
      TabIndex        =   10
      Top             =   -30
      Width           =   4725
      Begin VB.TextBox txtPos 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   810
         TabIndex        =   27
         Top             =   2340
         Visible         =   0   'False
         Width           =   1965
      End
      Begin VB.TextBox txtDisk 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   810
         TabIndex        =   26
         Top             =   1950
         Visible         =   0   'False
         Width           =   1965
      End
      Begin FPSpread.vaSpread vasWorkList 
         Height          =   6765
         Left            =   90
         TabIndex        =   16
         Top             =   690
         Width           =   4545
         _Version        =   196613
         _ExtentX        =   8017
         _ExtentY        =   11933
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   4
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "frmResult.frx":0000
      End
      Begin MSComCtl2.DTPicker dtpReceDate 
         Height          =   315
         Left            =   1050
         TabIndex        =   11
         Top             =   270
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   11599873
         CurrentDate     =   40122
      End
      Begin Threed.SSCommand cmdWorkList 
         Height          =   405
         Left            =   2700
         TabIndex        =   28
         Top             =   210
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   714
         _StockProps     =   78
         Caption         =   "QC 접수조회"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame Frame2 
         Height          =   495
         Left            =   2730
         TabIndex        =   13
         Top             =   120
         Visible         =   0   'False
         Width           =   1905
         Begin VB.OptionButton optQC 
            Caption         =   "Sample"
            Height          =   255
            Index           =   0
            Left            =   90
            TabIndex        =   15
            Top             =   180
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optQC 
            Caption         =   "QC"
            Height          =   255
            Index           =   1
            Left            =   1110
            TabIndex        =   14
            Top             =   180
            Width           =   645
         End
      End
      Begin VB.Label Label2 
         Caption         =   "접수일자"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   180
         TabIndex        =   12
         Top             =   330
         Width           =   795
      End
   End
   Begin VB.Frame Frame1 
      Height          =   9165
      Left            =   60
      TabIndex        =   0
      Top             =   -30
      Width           =   10035
      Begin FPSpread.vaSpread vasTemp 
         Height          =   1275
         Left            =   5070
         TabIndex        =   25
         Top             =   6750
         Visible         =   0   'False
         Width           =   2295
         _Version        =   196613
         _ExtentX        =   4048
         _ExtentY        =   2249
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
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "frmResult.frx":3C11
      End
      Begin FPSpread.vaSpread vasResTemp 
         Height          =   3165
         Left            =   3510
         TabIndex        =   24
         Top             =   2940
         Visible         =   0   'False
         Width           =   6405
         _Version        =   196613
         _ExtentX        =   11298
         _ExtentY        =   5583
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
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "frmResult.frx":3E63
      End
      Begin Threed.SSCommand cmdResCall 
         Height          =   405
         Left            =   5490
         TabIndex        =   2
         Top             =   180
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   714
         _StockProps     =   78
         Caption         =   "Local Data"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker dtpExamDate 
         Height          =   315
         Left            =   1140
         TabIndex        =   1
         Top             =   240
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   11599873
         CurrentDate     =   40122
      End
      Begin FPSpread.vaSpread vasID 
         Height          =   8385
         Left            =   90
         TabIndex        =   4
         Top             =   690
         Width           =   4665
         _Version        =   196613
         _ExtentX        =   8229
         _ExtentY        =   14790
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         ColHeaderDisplay=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GridColor       =   16777215
         MaxCols         =   20
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "frmResult.frx":40B5
      End
      Begin VB.Frame Frame3 
         Height          =   465
         Left            =   2730
         TabIndex        =   5
         Top             =   120
         Width           =   2685
         Begin VB.OptionButton optTransState 
            Caption         =   "전송"
            Height          =   225
            Index           =   2
            Left            =   1890
            TabIndex        =   8
            Top             =   180
            Width           =   735
         End
         Begin VB.OptionButton optTransState 
            Caption         =   "미전송"
            Height          =   225
            Index           =   1
            Left            =   900
            TabIndex        =   7
            Top             =   180
            Width           =   945
         End
         Begin VB.OptionButton optTransState 
            Caption         =   "전체"
            Height          =   225
            Index           =   0
            Left            =   90
            TabIndex        =   6
            Top             =   180
            Value           =   -1  'True
            Width           =   735
         End
      End
      Begin FPSpread.vaSpread vasRes 
         Height          =   8385
         Left            =   4800
         TabIndex        =   9
         Top             =   690
         Width           =   5145
         _Version        =   196613
         _ExtentX        =   9075
         _ExtentY        =   14790
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         ColHeaderDisplay=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GridColor       =   16777215
         MaxCols         =   15
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "frmResult.frx":5F71
      End
      Begin Threed.SSCommand cmdSelectTrans 
         Height          =   405
         Left            =   7380
         TabIndex        =   29
         Top             =   180
         Width           =   1875
         _Version        =   65536
         _ExtentX        =   3307
         _ExtentY        =   714
         _StockProps     =   78
         Caption         =   "결과선택전송"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "검사일자"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   180
         TabIndex        =   3
         Top             =   300
         Width           =   795
      End
   End
   Begin VB.Frame Frame5 
      Height          =   1635
      Left            =   10140
      TabIndex        =   17
      Top             =   7500
      Width           =   4725
      Begin VB.TextBox txtBarcode 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2580
         TabIndex        =   21
         Top             =   630
         Width           =   1965
      End
      Begin VB.TextBox txtResRow 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2580
         TabIndex        =   20
         Top             =   210
         Width           =   1965
      End
      Begin Threed.SSCommand cmdResTrans 
         Height          =   405
         Left            =   2100
         TabIndex        =   22
         Top             =   1080
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   714
         _StockProps     =   78
         Caption         =   "결과전송"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   405
         Left            =   3570
         TabIndex        =   23
         Top             =   1080
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   714
         _StockProps     =   78
         Caption         =   "종료"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label4 
         Caption         =   "Barcode 선택"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   810
         TabIndex        =   19
         Top             =   660
         Width           =   1755
      End
      Begin VB.Label Label3 
         Caption         =   "Local 결과 선택"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   810
         TabIndex        =   18
         Top             =   270
         Width           =   1755
      End
   End
End
Attribute VB_Name = "frmResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const colCheckBox = 1
Const colBarCode = 2
Const colDisk = 3
Const colPos = 4
Const colWkNo = 5
Const colPID = 6
Const colPName = 7
Const colJumin = 8
Const colPSex = 9
Const colPAge = 10
Const colOCnt = 11
Const colRCnt = 12
Const colState = 13
Const colSampleNo = 14
Const colSampleType = 15
Const colKind = 16
Const colPriority = 17

Const colAcpNo = 18     '접수번호

Const colEquipCode = 3
Const colExamCode = 4
Const colExamName = 5
Const colEquipRes = 6
Const colResult = 7
Const colSeqNo = 8
Const colResult1 = 9
Const colDataFlag = 15

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdResCall_Click()
Dim lRow As Long

    ClearSpread vasID
    vasID.maxrows = 0
'    SQL = "select barcode, diskno, posno, receno, pid, pname, pjumin, psex, page, count(*), count(*), min(sendflag)  " & vbCrLf & _
'          "FROM pat_res " & vbCrLf & _
'          "WHERE examdate = '" & Format(CDate(dtpExamDate), "yyyymmdd") & "' " & vbCrLf & _
'          "  AND equipno = '" & gEquip & "' "
'          If optTransState(0).Value = True Then
'          ElseIf optTransState(1).Value = True Then
'            SQL = SQL & vbCrLf & "and sendflag = 'B'"
'          ElseIf optTransState(2).Value = True Then
'            SQL = SQL & vbCrLf & "and sendflag = 'C'"
'          End If
'          SQL = SQL & vbCrLf & "Group by barcode, diskno, posno, receno, pid, pname, pjumin, psex, page "

'    SQL = "select barcode, seqno, receno, diskno, posno, pid, pname, page, psex, jumin, sendflag, count(*), count(*), max(recedate)" & _
'          " from pat_res " & _
'          "where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
'          "  and equipno = '" & Trim(gEquip) & "' " & vbCrLf & _
'          "  and sendflag in ('B','C') " & vbCrLf & _
'          "group by diskno, posno, barcode, seqno, receno, pid, pname, page, psex, jumin, sendflag "
'    SQL = SQL & vbCrLf & " Union " & vbCrLf
'    SQL = SQL & vbCrLf & _
'          "select barcode, seqno, receno, diskno, posno, pid, pname, page, psex, jumin, sendflag, count(*), '0',  max(recedate)" & _
'          " from pat_res " & _
'          "where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
'          "  and equipno = '" & Trim(gEquip) & "' " & vbCrLf & _
'          "  and sendflag not in ('B','C') " & vbCrLf & _
'          "group by diskno, posno, barcode, seqno, receno,  pid, pname, page, psex, jumin, sendflag " & vbCrLf & _
'          "order by diskno,posno"
    SQL = "select barcode, diskno, posno, receno, pid, pname, jumin, psex, page, count(*), count(*), min(sendflag)  " & vbCrLf & _
          "FROM pat_res " & vbCrLf & _
          "WHERE examdate = '" & Format(CDate(dtpExamDate), "yyyymmdd") & "' " & vbCrLf & _
          "  AND equipno = '" & gEquip & "' "
          If optTransState(0).Value = True Then
          ElseIf optTransState(1).Value = True Then
            SQL = SQL & vbCrLf & "and sendflag in ('A','B') "
          ElseIf optTransState(2).Value = True Then
            SQL = SQL & vbCrLf & "and sendflag = 'C'"
          End If
          SQL = SQL & vbCrLf & "Group by barcode, diskno, posno, receno, pid, pname, jumin, psex, page "

'    res = db_select_Vas(gLocal, SQL, vasID, vasID.DataRowCnt + 1, 2)
    
    res = db_select_Vas(gLocal, SQL, vasID, 1, 2)
    vasSort vasID, colWkNo
    For lRow = 1 To vasID.DataRowCnt
        Select Case Trim(GetText(vasID, lRow, colState))
        Case "C"
            SetText vasID, "완료", lRow, colState
            SetBackColor vasID, lRow, lRow, colBarCode, colState, 202, 255, 112
        Case "B", "A"
            SetText vasID, "결과", lRow, colState
'        Case "A"
'            SetText vasID, "오더", lRow, colState
        End Select
    Next lRow
    
End Sub

Private Sub cmdResTrans_Click()
    Dim lsRow As Integer
    Dim lsbarcode As String
    Dim i As Integer
    Dim j As Integer
'    Dim lsBarcode As String
    Dim sBarCode As String
    Dim sBarCode1 As String
    Dim sExamCode As String
    Dim sResult As String
    Dim sEquip As String
    Dim sRefFlag As String
    Dim sRet As String
    Dim sErrFlag As String
    Dim sRCnt As String
    Dim lsQCGubun As Boolean
    Dim lsReceRow As Integer
    Dim sLotNo As String
    Dim sExamTime As String

    Dim sSelExamCode As String
    Dim ii As Integer
'    Dim j As Integer
    Dim k As Integer
    Dim sResMach As Boolean
    
    If MsgBox(" " & vbCrLf & "선택전송을 하시겠습니까?" & vbCrLf & " ", vbInformation + vbOKCancel, "알림:선택전송") = vbCancel Then
        Exit Sub
    End If
    
    
'    If lsRow < 1 Or lsRow > vasID.DataRowCnt Then
'        MsgBox "저장할 데이터가 없습니다."
'        Exit Sub
'    End If
    
    sResMach = False
    If IsNumeric(txtResRow) = True Then
        lsRow = CCur(txtResRow)
    Else
        MsgBox "저장할 데이터가 없습니다."
        Exit Sub
    End If
    
    
    lsbarcode = Trim(txtBarcode.Text)
    If lsbarcode = "" Then
        MsgBox "저장할 데이터가 없습니다."
        Exit Sub
    End If
    
    If Mid(lsbarcode, 1, 1) = "9" Then
        lsQCGubun = True
    Else
        lsQCGubun = False
    End If
    
    If lsQCGubun = True Then
        For i = 1 To vasWorkList.DataRowCnt
            If lsbarcode = Trim(GetText(vasWorkList, i, 1)) Then
                lsReceRow = i
                Exit For
            End If
        Next
        
        sLotNo = Trim(GetText(vasWorkList, lsReceRow, 3))
'         If Trim(GetText(vasResTemp, 1, 7)) = "" Then
'            sExamTime = Format(Date, "yyyymmddhhmmss")
'         Else
'            sExamTime = Format(Trim(GetText(vasResTemp, 1, 7)), "yyyymmddhhmmss")
'         End If
        
        gReceCode = ""
        Get_QCList Mid(lsbarcode, 1, 11), 1
        
        If lsbarcode = gQC_Info(0).BARCODE_CD Then
            sExamTime = Format(gQC_Info(0).INST_DTM, "yyyymmddhhmmss")
            
'            For i = 1 To vasRes.DataRowCnt
'                If Trim(GetText(vasRes, i, 4)) = "" Then
'                    SQL = " Select examcode from equipexam " & CR & _
'                          " Where equipno = '" & gEquip & "' " & CR & _
'                          " And equipcode = '" & Trim(GetText(vasRes, i, 3)) & "' " & CR & _
'                          " And examcode in (" & gReceCode & ") "
'                    res = db_select_Col(gLocal, SQL)
'
'                    If res > 0 Then
'                        sResMach = True
'                    End If
'
'                    SetText vasRes, Trim(lsbarcode), i, 2
'                    SetText vasRes, Trim(gReadBuf(0)), i, 4
'                End If
'            Next i
            
        End If
        
'        If sResMach = False Then
'            SetForeColor vasID, lsRow, lsRow, 1, vasID.MaxCols, 255, 0, 0
'            Exit Sub
'        End If
        
        sEquip = gEquip
           
        sExamCode = ""
        sResult = ""
        
        sRCnt = "0"

'        For i = 1 To vasResTemp.DataRowCnt
'            sRCnt = sRCnt + 1
'
'            If sExamCode = "" Then
'                sExamCode = chrTAB & Trim(GetText(vasResTemp, i, 2)) & chrTAB
'            Else
'                sExamCode = sExamCode & Trim(GetText(vasResTemp, i, 2)) & chrTAB
'            End If
'
'            If sResult = "" Then
'                sResult = chrTAB & Trim(GetText(vasResTemp, i, 3)) & chrTAB
'            Else
'                sResult = sResult & Trim(GetText(vasResTemp, i, 3)) & chrTAB
'            End If
'        Next i
        
        
        For i = 1 To vasRes.DataRowCnt
            If Trim(GetText(vasRes, i, 4)) <> "" And Trim(GetText(vasRes, i, 6)) <> "" Then
                sRCnt = sRCnt + 1
    
                If sExamCode = "" Then
                    sExamCode = chrTAB & Trim(GetText(vasRes, i, 4)) & chrTAB
                Else
                    sExamCode = sExamCode & Trim(GetText(vasRes, i, 4)) & chrTAB
                End If
                
                If sResult = "" Then
                    sResult = chrTAB & Trim(GetText(vasRes, i, 6)) & chrTAB
                Else
                    sResult = sResult & Trim(GetText(vasRes, i, 6)) & chrTAB
                End If
            End If
        Next i

         sRet = Online_QCResult(lsbarcode, gQC_Info(0).EQUIP_CD, sLotNo, sExamTime, sRCnt, sExamCode, sResult, gWorker_Info.WK_ID)
         If sRet = "N" Then
         Else
         
            SetForeColor vasID, lsRow, lsRow, 1, vasID.MaxCols, 255, 0, 0
             'QC전송성공
         End If
         
         SQL = " update pat_res set sendflag = 'C' " & vbCrLf & _
               "WHERE examdate = '" & Format(CDate(dtpExamDate), "yyyymmdd") & "' " & vbCrLf & _
               "  AND equipno = '" & gEquip & "' " & vbCrLf & _
               "  AND Barcode = '" & Trim(GetText(vasID, lsRow, colBarCode)) & "'   "
         res = SendQuery(gLocal, SQL)
        
    Else
        sBarCode1 = ""
        sExamCode = ""
        sResult = ""
        sEquip = ""
        sRefFlag = ""
        
        sRCnt = "0"
        ClearSpread vasResTemp
    
        SQL = " Select equipcode, examcode, equipres, refflag, panicflag, deltaflag, resdate " & vbCrLf & _
              " From pat_res " & vbCrLf & _
              "WHERE examdate = '" & Format(CDate(dtpExamDate), "yyyymmdd") & "' " & vbCrLf & _
              "  AND equipno = '" & gEquip & "' " & vbCrLf & _
              "  AND Barcode = '" & Trim(GetText(vasID, lsRow, colBarCode)) & "'"
        res = db_select_Vas(gLocal, SQL, vasResTemp)
        gReceCode = ""
        Get_Order lsbarcode
        For i = 1 To vasResTemp.DataRowCnt

            SQL = " Select examcode from equipexam " & CR & _
                  " Where equipno = '" & gEquip & "' " & CR & _
                  " And equipcode = '" & Trim(GetText(vasResTemp, i, 1)) & "' " & CR & _
                  " And examcode in (" & gReceCode & ") "
            res = db_select_Col(gLocal, SQL)
            
            If res > 0 Then
                sResMach = True
                SetText vasResTemp, gReadBuf(0), i, 2
    
                If Trim(GetText(vasResTemp, i, 4)) <> "" Then
                    gReadBuf(0) = ""
                    SQL = "select flag from equipflag where flag = '" & Trim(GetText(vasResTemp, i, 4)) & "' and useflag = '1'"
                    res = db_select_Col(gLocal, SQL)
                    If res > 0 Then
                    Else
                        SetText vasResTemp, "", i, 4
                    End If
                    
                    If Trim(GetText(vasResTemp, i, 4)) <> "" Then
                
                        If Trim(GetText(vasResTemp, i, 4)) = "N" Or Trim(GetText(vasResTemp, i, 4)) = "H" Or Trim(GetText(vasResTemp, i, 4)) = "L" Then
                        Else
                            SetText vasResTemp, Trim(GetText(vasResTemp, i, 3)) & "|" & Trim(GetText(vasResTemp, i, 4)), i, 3
                        End If
                    End If
                End If
                
                sRCnt = sRCnt + 1
                
                If sExamCode = "" Then
                    sExamCode = chrTAB & Trim(GetText(vasResTemp, i, 2)) & chrTAB
                Else
                    sExamCode = sExamCode & Trim(GetText(vasResTemp, i, 2)) & chrTAB
                End If
                
                If sBarCode1 = "" Then
                    sBarCode1 = chrTAB & lsbarcode & chrTAB
                Else
                    sBarCode1 = sBarCode1 & lsbarcode & chrTAB
                End If
                
                If sResult = "" Then
                    sResult = chrTAB & Trim(GetText(vasResTemp, i, 3)) & chrTAB
                Else
                    sResult = sResult & Trim(GetText(vasResTemp, i, 3)) & chrTAB
                End If
                
                If sEquip = "" Then
                    sEquip = chrTAB & gEquip & chrTAB
                Else
                    sEquip = sEquip & gEquip & chrTAB
                End If
            End If
        Next i
        
        sErrFlag = ""
        
'        If sErrFlag <> "" Then
'            sBarCode1 = chrTAB & lsBarCode & sBarCode1
'
'            sExamCode = chrTAB & "REMARK" & sExamCode
'
'            sResult = chrTAB & "※장비비고:" & sErrFlag & sResult
'
'            sEquip = chrTAB & gEquip & sEquip
'
'            sRCnt = sRCnt + 1
'        End If
        
        sRet = Online_Result_New(sBarCode1, sExamCode, sResult, sEquip, sRCnt, "", gWorker_Info.WK_ID)
        
        SQL = " update pat_res set sendflag = 'C' " & vbCrLf & _
              "WHERE examdate = '" & Format(CDate(dtpExamDate), "yyyymmdd") & "' " & vbCrLf & _
              "  AND equipno = '" & gEquip & "' " & vbCrLf & _
              "  AND Barcode = '" & Trim(GetText(vasID, lsRow, colBarCode)) & "'   "
        res = SendQuery(gLocal, SQL)
           
        If sRet = "N" Then
            '전송선공
        End If
    End If
    
    txtBarcode.Text = ""
    txtResRow.Text = ""
End Sub

Private Sub cmdSelectTrans_Click()
    Dim i As Integer
    Dim j As Integer
    Dim liRet As Integer
    
    
    If MsgBox(" " & vbCrLf & "선택전송을 하시겠습니까?" & vbCrLf & " ", vbInformation + vbOKCancel, "알림:선택전송") = vbCancel Then
        Exit Sub
    End If
    
    If (vasID.DataRowCnt < 1) Then
        MsgBox "저장할 데이터가 없습니다."
        Exit Sub
    End If
    
    For i = 1 To vasID.DataRowCnt
        vasID.Col = 1
        vasID.Row = i
        If vasID.Value = 1 Then
            For j = 1 To vasWorkList.DataRowCnt
                If Trim(GetText(vasID, i, colBarCode)) = Trim(GetText(vasWorkList, j, 4)) Then
                    liRet = -1
                    If Trim(GetText(vasID, i, colBarCode)) <> "" Then
                        SetBackColor vasID, i, i, colBarCode, colState, 255, 255, 255
                        liRet = ToServer_QC_Res(i, j)
                        If liRet = 1 Then
                            SetText vasID, "완료", i, colState
                            SetForeColor vasID, i, i, colBarCode, colState, 0, 0, 0
                            SetBackColor vasID, i, i, colBarCode, colState, 202, 255, 112
                            
                            vasID.Col = 1
                            vasID.Row = i
                            vasID.Value = 0
                        Else
                            SetText vasID, "실패", i, colState
                            SetForeColor vasID, i, i, colBarCode, colState, 255, 0, 0
                            SetBackColor vasID, i, i, colBarCode, colState, 255, 255, 255
                        End If
                    End If
                    Exit For
                End If
            Next
        End If
        
    Next
    
End Sub

Private Function ToServer_QC_Res(vasIDRow As Integer, vasWorkListRow As Integer) As Integer
    Dim lsRow As Integer
    Dim lsbarcode As String
    Dim i As Integer
    Dim j As Integer
'    Dim lsBarcode As String
    Dim sBarCode As String
    Dim sBarCode1 As String
    Dim sExamCode As String
    Dim sResult As String
    Dim sEquip As String
    Dim sRefFlag As String
    Dim sRet As String
    Dim sErrFlag As String
    Dim sRCnt As String
    Dim lsQCGubun As Boolean
    Dim lsReceRow As Integer
    Dim sLotNo As String
    Dim sExamTime As String

    Dim sSelExamCode As String
    Dim ii As Integer
'    Dim j As Integer
    Dim k As Integer
    Dim sResMach As Boolean
    Dim sMsgNum
    
    
    ToServer_QC_Res = -1
    
    lsReceRow = vasWorkListRow
    lsRow = vasIDRow
    
    
    vasID_Click colBarCode, lsRow
    
    
'    For i = 1 To vasWorkList.DataRowCnt
'        If lsbarcode = Trim(GetText(vasWorkList, i, 1)) Then
'            lsReceRow = i
'            Exit For
'        End If
'    Next
    
    lsbarcode = Trim(GetText(vasWorkList, lsReceRow, 1))
    sLotNo = Trim(GetText(vasWorkList, lsReceRow, 3))
    
    gReceCode = ""
    Get_QCList Mid(lsbarcode, 1, 11), 1
    
    If lsbarcode = gQC_Info(0).BARCODE_CD Then
        sExamTime = Format(gQC_Info(0).INST_DTM, "yyyymmddhhmmss")
        
        For i = 1 To vasRes.DataRowCnt
            If Trim(GetText(vasRes, i, 4)) = "" Then
                SQL = " Select examcode from equipexam " & CR & _
                      " Where equipno = '" & gEquip & "' " & CR & _
                      " And equipcode = '" & Trim(GetText(vasRes, i, 3)) & "' " & CR & _
                      " And examcode in (" & gReceCode & ") "
                res = db_select_Col(gLocal, SQL)
                
                If res > 0 Then
                    sResMach = True
                End If
                
                SetText vasRes, Trim(lsbarcode), i, 2
                SetText vasRes, Trim(gReadBuf(0)), i, 4
            End If
        Next i
        
    End If
    
    If sResMach = False Then
        SetForeColor vasID, lsRow, lsRow, 1, vasID.MaxCols, 255, 0, 0
        Exit Function
    End If
    
    sEquip = gEquip
       
    sExamCode = ""
    sResult = ""
    
    sRCnt = "0"


    For i = 1 To vasRes.DataRowCnt
        If Trim(GetText(vasRes, i, 4)) <> "" And Trim(GetText(vasRes, i, 6)) <> "" Then
            sRCnt = sRCnt + 1

            If sExamCode = "" Then
                sExamCode = chrTAB & Trim(GetText(vasRes, i, 4)) & chrTAB
            Else
                sExamCode = sExamCode & Trim(GetText(vasRes, i, 4)) & chrTAB
            End If

            If sResult = "" Then
                sResult = chrTAB & Trim(GetText(vasRes, i, 6)) & chrTAB
            Else
                sResult = sResult & Trim(GetText(vasRes, i, 6)) & chrTAB
            End If
        End If
    Next i

     sRet = Online_QCResult(lsbarcode, sEquip, sLotNo, sExamTime, sRCnt, sExamCode, sResult, gWorker_Info.WK_ID)
     If sRet = "N" Then
     Else
        SetForeColor vasID, lsRow, lsRow, 1, vasID.MaxCols, 255, 0, 0
        Exit Function
        
     End If
     
     SQL = " update pat_res set sendflag = 'C' " & vbCrLf & _
           "WHERE examdate = '" & Format(CDate(dtpExamDate), "yyyymmdd") & "' " & vbCrLf & _
           "  AND equipno = '" & gEquip & "' " & vbCrLf & _
           "  AND Barcode = '" & Trim(GetText(vasID, lsRow, colBarCode)) & "'   "
     res = SendQuery(gLocal, SQL)
     
     
     ToServer_QC_Res = 1
End Function

Private Sub cmdWorkList_Click()
    Dim lsbarcode As String
    Dim lRow As Long
    Dim lsState As String
    Dim i, j As Integer
    Dim lsRow As Integer
    Dim lsPatFlag As Boolean
    Dim lsStatCode As String
    Dim lsQCGubun As Boolean
    
'    If Index = 1 Then
        lsQCGubun = True
'    Else
'        lsQCGubun = False
'    End If
    
    ClearSpread vasWorkList
    
    If lsQCGubun = True Then
        
        Get_QCWorkList Format(dtpReceDate, "yyyymmdd"), gEquip
        
        For i = 0 To giIndex
            lsRow = -1
            lsPatFlag = False
            
            For j = 1 To vasWorkList.DataRowCnt
                If Trim(GetText(vasWorkList, j, 1)) = Trim(gQC_Info(i).BARCODE_CD) Then
                    lsRow = j
                    lsPatFlag = True
                    
                    Exit For
                End If
            Next
            
            If lsRow = -1 Then
                lsRow = vasWorkList.DataRowCnt + 1
            End If
            
            If vasWorkList.maxrows < lsRow Then
                vasWorkList.maxrows = lsRow
            End If
            If lsPatFlag = True Then
            Else
                
                vasWorkList.SetText 1, lsRow, gQC_Info(i).BARCODE_CD
                vasWorkList.SetText 2, lsRow, "QC"
                vasWorkList.SetText 3, lsRow, gQC_Info(i).LOT_NO
                vasWorkList.SetText 4, lsRow, gQC_Info(i).CTRL_CD
        
            End If
            
        Next
        vasWorkList.maxrows = vasWorkList.DataRowCnt
    Else
        SQL = "select examcode from equipexam " & vbCrLf & _
              "where equipno = '" & gEquip & "' "
              
        res = db_select_Row(gLocal, SQL)
        lsStatCode = ""
        For i = 1 To res
            If Trim(lsStatCode) = "" Then
                lsStatCode = "|" & gReadBuf(i - 1) '& "|"
            Else
                lsStatCode = lsStatCode & "|" & gReadBuf(i - 1) '& "|"
            End If
        Next
        lsStatCode = lsStatCode & "|"
        
        Get_WorkList Format(dtpReceDate, "yyyymmdd"), Format(dtpReceDate, "yyyymmdd"), lsStatCode, 3
        
        For i = 0 To giIndex
            lsRow = -1
            lsPatFlag = False
            
            For j = 1 To vasWorkList.DataRowCnt
                If Trim(GetText(vasWorkList, j, 1)) = Trim(gWork_Select(i).SPC_NO) Then
                    lsRow = j
                    lsPatFlag = True
                    
                    Exit For
                End If
            Next
            
            If lsRow = -1 Then
                lsRow = vasWorkList.DataRowCnt + 1
            End If
            
            If vasWorkList.maxrows < lsRow Then
                vasWorkList.maxrows = lsRow
            End If
            If lsPatFlag = True Then
            Else
                
                vasWorkList.SetText 1, lsRow, gWork_Select(i).SPC_NO
                vasWorkList.SetText 2, lsRow, gWork_Select(i).ACPT_NO
                vasWorkList.SetText 3, lsRow, gWork_Select(i).PT_NO
                vasWorkList.SetText 4, lsRow, gWork_Select(i).PT_NM
        
            End If
            vasWorkList.maxrows = vasWorkList.DataRowCnt
        Next
        
    End If
    
    vasWorkList.RowHeight(-1) = 11
End Sub

Private Sub Form_Load()

    dtpExamDate.Value = frmInterface.Text_Today.Text
    dtpReceDate.Value = frmInterface.Text_Today.Text
End Sub

Private Sub optQC_Click(Index As Integer)
    Dim lsbarcode As String
    Dim lRow As Long
    Dim lsState As String
    Dim i, j As Integer
    Dim lsRow As Integer
    Dim lsPatFlag As Boolean
    Dim lsStatCode As String
    Dim lsQCGubun As Boolean
    
    If Index = 1 Then
        lsQCGubun = True
    Else
        lsQCGubun = False
    End If
    
    ClearSpread vasWorkList
    
    If lsQCGubun = True Then
        
        Get_QCWorkList Format(dtpReceDate, "yyyymmdd"), gEquip
        
        For i = 0 To giIndex
            lsRow = -1
            lsPatFlag = False
            
            For j = 1 To vasWorkList.DataRowCnt
                If Trim(GetText(vasWorkList, j, 1)) = Trim(gQC_Info(i).BARCODE_CD) Then
                    lsRow = j
                    lsPatFlag = True
                    
                    Exit For
                End If
            Next
            
            If lsRow = -1 Then
                lsRow = vasWorkList.DataRowCnt + 1
            End If
            
            If vasWorkList.maxrows < lsRow Then
                vasWorkList.maxrows = lsRow
            End If
            If lsPatFlag = True Then
            Else
                
                vasWorkList.SetText 1, lsRow, gQC_Info(i).BARCODE_CD
                vasWorkList.SetText 2, lsRow, "QC"
                vasWorkList.SetText 3, lsRow, gQC_Info(i).LOT_NO
                vasWorkList.SetText 4, lsRow, gQC_Info(i).CTRL_CD
        
            End If
            
        Next
        vasWorkList.maxrows = vasWorkList.DataRowCnt
    Else
        SQL = "select examcode from equipexam " & vbCrLf & _
              "where equipno = '" & gEquip & "' "
              
        res = db_select_Row(gLocal, SQL)
        lsStatCode = ""
        For i = 1 To res
            If Trim(lsStatCode) = "" Then
                lsStatCode = "|" & gReadBuf(i - 1) '& "|"
            Else
                lsStatCode = lsStatCode & "|" & gReadBuf(i - 1) '& "|"
            End If
        Next
        lsStatCode = lsStatCode & "|"
        
        Get_WorkList Format(dtpReceDate, "yyyymmdd"), Format(dtpReceDate, "yyyymmdd"), lsStatCode, 3
        
        For i = 0 To giIndex
            lsRow = -1
            lsPatFlag = False
            
            For j = 1 To vasWorkList.DataRowCnt
                If Trim(GetText(vasWorkList, j, 1)) = Trim(gWork_Select(i).SPC_NO) Then
                    lsRow = j
                    lsPatFlag = True
                    
                    Exit For
                End If
            Next
            
            If lsRow = -1 Then
                lsRow = vasWorkList.DataRowCnt + 1
            End If
            
            If vasWorkList.maxrows < lsRow Then
                vasWorkList.maxrows = lsRow
            End If
            If lsPatFlag = True Then
            Else
                
                vasWorkList.SetText 1, lsRow, gWork_Select(i).SPC_NO
                vasWorkList.SetText 2, lsRow, gWork_Select(i).ACPT_NO
                vasWorkList.SetText 3, lsRow, gWork_Select(i).PT_NO
                vasWorkList.SetText 4, lsRow, gWork_Select(i).PT_NM
        
            End If
            vasWorkList.maxrows = vasWorkList.DataRowCnt
        Next
        
    End If
    
    vasWorkList.RowHeight(-1) = 11
End Sub

Private Sub SSCommand1_Click()
        
End Sub

Private Sub vasID_Click(ByVal Col As Long, ByVal Row As Long)
    Dim lsID As String
    Dim lsTmpID As String
    
    Dim i As Integer
    
    '샘플번호에 해당 하는 검사결과 Local Databse에서 가져오기
    
    If Row = 0 Then
        Select Case Col
        Case colBarCode, colWkNo, colPID, colPName
            vasSort vasID, Col
        Case colDisk, colPos
            vasSort vasID, colDisk, colPos
        Case colPSex
            vasSort vasID, colPSex, colJumin
        Case colPAge
            vasSort vasID, colPAge, colJumin
        Case colState
            vasSort vasID, colState, colDisk, colPos
        End Select
    End If
    
    If Row < 1 Or Row > vasID.DataRowCnt Then
        Exit Sub
    End If
    
    lsID = Trim(GetText(vasID, Row, colBarCode))

    ClearSpread vasRes
    vasRes.maxrows = 0
    
'    SQL = "Select a.equipcode, a.examcode, b.examname, a.result, b.seqno, a.refflag, a.result1 " & vbCrLf & _
'          "from pat_res a, equipexam b " & vbCrLf & _
'          "where a.examdate = '" & Format(CDate(Text_Today), "yyyymmdd") & "' " & vbCrLf & _
'          "  and a.equipno = '" & gEquip & "' " & vbCrLf & _
'          "  and a.barcode = '" & lsID & "' " & vbCrLf & _
'          "  and a.examcode <> a.equipcode " & vbCrLf & _
'          "  and b.equipno = a.equipno " & vbCrLf & _
'          "  and b.equipcode = a.equipcode " & vbCrLf & _
'          "  and b.examcode = a.examcode"
'    res = db_select_Vas(gLocal, SQL, vasRes)
'    SQL = "Select a.equipcode, a.examcode, max(b.examname), a.result, b.seqno, a.refflag, a.result1 " & vbCrLf & _
'          "from pat_res a, equipexam b " & vbCrLf & _
'          "where a.examdate = '" & Format(CDate(Text_Today), "yyyymmdd") & "' " & vbCrLf & _
'          "  and a.equipno = '" & gEquip & "' " & vbCrLf & _
'          "  and a.barcode = '" & lsID & "' " & vbCrLf & _
'          "  and a.examcode = a.equipcode " & vbCrLf & _
'          "  and b.equipno = a.equipno " & vbCrLf & _
'          "  and b.equipcode = a.equipcode " & vbCrLf & _
'          "group by a.equipcode, a.examcode, a.result, b.seqno, a.refflag, a.result1 "
'    res = db_select_Vas(gLocal, SQL, vasRes, vasRes.DataRowCnt + 1, 1)
'
    
    SQL = "select '', barcode, equipcode,  examcode, examname, result, result, seqno, result ,'','','','','',refflag " & vbCrLf & _
          "FROM pat_res " & vbCrLf & _
          "WHERE examdate = '" & Format(CDate(dtpExamDate), "yyyymmdd") & "' " & vbCrLf & _
          "  AND equipno = '" & gEquip & "' " & vbCrLf & _
          "  AND Barcode = '" & Trim(GetText(vasID, vasID.Row, colBarCode)) & "' " & vbCrLf & _
          "  order by equipcode"
          
    res = db_select_Vas(gLocal, SQL, vasRes)
    If res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If
    
    For i = 1 To vasRes.DataRowCnt
        If InStr(1, Trim(GetText(vasRes, i, colResult)), "Pos") > 0 Then
            vasRes.Row = i
            vasRes.Col = colResult
            vasRes.ForeColor = RGB(205, 55, 0)
        Else
            vasRes.Row = i
            vasRes.Col = colResult
            vasRes.ForeColor = RGB(0, 0, 0)
        End If
        
    Next i
    
    txtResRow.Text = Row
    
    txtDisk = ""
    txtPos = ""
    txtDisk = Trim(GetText(vasID, vasID.Row, 3))
    txtPos = Trim(GetText(vasID, vasID.Row, 4))
    
    SetBackColor vasID, 1, vasID.maxrows, 1, vasID.MaxCols, 255, 255, 255
    SetBackColor vasID, Row, Row, 1, vasID.MaxCols, 200, 200, 240
End Sub

Private Sub vasWorkList_Click(ByVal Col As Long, ByVal Row As Long)
    txtBarcode.Text = ""
    txtBarcode.Text = Trim(GetText(vasWorkList, Row, 1))
    SetBackColor vasWorkList, 1, vasWorkList.maxrows, 1, vasWorkList.MaxCols, 255, 255, 255
    SetBackColor vasWorkList, Row, Row, 1, vasWorkList.MaxCols, 200, 200, 240
End Sub
