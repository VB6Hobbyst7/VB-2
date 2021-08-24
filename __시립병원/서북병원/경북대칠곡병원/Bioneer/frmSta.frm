VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#3.5#0"; "SPR32X35.ocx"
Begin VB.Form frmSta 
   Caption         =   "Elecsys 검사 건수"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   8970
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton cmdRerun 
      Caption         =   "재검"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   8
      Top             =   150
      Width           =   945
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
      Height          =   495
      Left            =   7740
      TabIndex        =   7
      Top             =   143
      Width           =   945
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "출력"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6750
      TabIndex        =   6
      Top             =   143
      Width           =   945
   End
   Begin VB.CommandButton cmdSch 
      Caption         =   "조회"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4770
      TabIndex        =   5
      Top             =   143
      Width           =   945
   End
   Begin FPSpread.vaSpread vasList 
      Height          =   2235
      Left            =   240
      TabIndex        =   4
      Top             =   810
      Width           =   8505
      _Version        =   196613
      _ExtentX        =   15002
      _ExtentY        =   3942
      _StockProps     =   64
      ColHeaderDisplay=   1
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   7
      MaxRows         =   5
      ScrollBars      =   2
      SpreadDesigner  =   "frmSta.frx":0000
   End
   Begin MSComCtl2.DTPicker dtpSch1 
      Height          =   330
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   23789569
      CurrentDate     =   37174
   End
   Begin MSComCtl2.DTPicker dtpSch2 
      Height          =   330
      Left            =   3060
      TabIndex        =   1
      Top             =   240
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   23789569
      CurrentDate     =   37174
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "~"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2880
      TabIndex        =   3
      Top             =   300
      Width           =   120
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "조회일자"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   210
      TabIndex        =   2
      Top             =   300
      Width           =   900
   End
End
Attribute VB_Name = "frmSta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Dim sHead As String
    Dim sFoot As String
    Dim sCurDate As String
    
    If vasList.DataRowCnt < 1 Then
        MsgBox "출력할 자료가 없습니다.", , "알 림"
        Exit Sub
    End If

    sCurDate = GetDateFull
    
    sHead = dtpSch1.Value
    If (IsDate(dtpSch1.Value) = True And dtpSch1.Value <> dtpSch2.Value) Then
        sHead = sHead & " ~ " & dtpSch2.Value
    End If
     sHead = "조회 일자 : " & sHead
    sHead = "/fn""궁서체"" /fz""15"" /fb1 /fi0 /fu0 " & "/c" & "▣ Elecsys 검사 건수 통계 ▣" & "/n/n " & _
                "/fn""굴림체"" /fz""11"" /fb0 /fi0 /fu0 " & "/c" & sHead & "/n" & _
                "/fn""굴림체"" /fz""10"" /fb0 /fi0 /fu0 " & "/rPage /p" & "/n"
    sFoot = "/fn""굴림체"" /fz""10"" /fb1 /fi0 /fu0 " & "/l" & sCurDate & "/fn""궁서체"" /fz""11"" /fb1 /fi0 /fu0 /r" & "SCL 부산"
    vasList.PrintOrientation = 1   ' SS_PRINTORIENT_PORTRAIT
    vasList.PrintAbortMsg = "인쇄중 입니다 ..."
    vasList.PrintJobName = "Elecsys 검사 건수"
    vasList.PrintHeader = sHead
    vasList.PrintFooter = sFoot
    vasList.PrintMarginTop = 720
    vasList.PrintMarginBottom = 720
'현재 SS가 비대칭으로 출력함
'    vaslist.PrintMarginLeft = 720
    vasList.PrintMarginLeft = 500
    vasList.PrintMarginRight = 500
    
    vasList.PrintColor = True
    vasList.PrintGrid = True
'Set printing range
    vasList.PrintType = 0  'SS_PRINT_ALL(default)

    vasList.PrintShadows = True

    vasList.Action = 13 'SS_ACTION_PRINT

End Sub

Private Sub cmdRerun_Click()
    frmRerun.Show 1
    cmdSch_Click
End Sub

Private Sub cmdSch_Click()
    Dim sDate1 As String
    Dim sDate2 As String
    Dim a, b, c As Long
    Dim aSum, bSum, cSum As Double
    
    Dim i, j As Long
    
    sDate1 = SeperatorCls(dtpSch1.Value)
    sDate2 = SeperatorCls(dtpSch2.Value)
    
    ClearSpread vasList
    
    SQL = "Select equipcode, examname from equipexam order by equipcode "
    res = db_select_Vas(gLocal, SQL, vasList)
    
    ClearSpread Form_Main.vasTemp
    
    SQL = "Select equipcode,count(*) " & vbCrLf & _
          "from pat_res " & vbCrLf & _
          "Where examdate between '" & sDate1 & "' and '" & sDate2 & "' " & vbCrLf & _
          "group by equipcode " & vbCrLf & _
          "order by equipcode "
    res = db_select_Vas(gLocal, SQL, Form_Main.vasTemp)
    For j = 1 To Form_Main.vasTemp.DataRowCnt
        For i = 1 To vasList.DataRowCnt
            If Trim(GetText(Form_Main.vasTemp, j, 1)) = Trim(GetText(vasList, i, 1)) Then
                SetText vasList, Trim(GetText(Form_Main.vasTemp, j, 2)), i, 3
                Exit For
            End If
        Next i
    Next j
    
    ClearSpread Form_Main.vasTemp
    
    SQL = "Select equipcode,count(*) " & vbCrLf & _
          "from qc_res " & vbCrLf & _
          "Where examdate between '" & sDate1 & "' and '" & sDate2 & "' " & vbCrLf & _
          "group by equipcode " & vbCrLf & _
          "order by equipcode "
    res = db_select_Vas(gLocal, SQL, Form_Main.vasTemp)
    For j = 1 To Form_Main.vasTemp.DataRowCnt
        For i = 1 To vasList.DataRowCnt
            If Trim(GetText(Form_Main.vasTemp, j, 1)) = Trim(GetText(vasList, i, 1)) Then
                SetText vasList, Trim(GetText(Form_Main.vasTemp, j, 2)), i, 4
                Exit For
            End If
        Next i
    Next j
        
    ClearSpread Form_Main.vasTemp
    SQL = "Select equipcode,sum(r_cnt)" & vbCrLf & _
          "from rerun_cnt " & vbCrLf & _
          "Where examdate between '" & sDate1 & "' and '" & sDate2 & "' " & vbCrLf & _
          "group by equipcode " & vbCrLf & _
          "order by equipcode "
    res = db_select_Vas(gLocal, SQL, Form_Main.vasTemp)
    For j = 1 To Form_Main.vasTemp.DataRowCnt
        For i = 1 To vasList.DataRowCnt
            If Trim(GetText(Form_Main.vasTemp, j, 1)) = Trim(GetText(vasList, i, 1)) Then
                SetText vasList, Trim(GetText(Form_Main.vasTemp, j, 2)), i, 6
                Exit For
            End If
        Next i
    Next j
    
    aSum = 0
    bSum = 0
    cSum = 0
    For i = 1 To vasList.DataRowCnt
        If IsNumeric(GetText(vasList, i, 3)) Then
            a = CLng(GetText(vasList, i, 3))
        Else
            a = 0
        End If
        If IsNumeric(GetText(vasList, i, 4)) Then
            b = CLng(GetText(vasList, i, 4))
        Else
            b = 0
        End If
        If IsNumeric(GetText(vasList, i, 6)) Then
            c = CLng(GetText(vasList, i, 6))
        Else
            c = 0
        End If
        SetText vasList, Format(a, "#,##0"), i, 3
        SetText vasList, Format(b, "#,##0"), i, 4
        SetText vasList, Format(c, "#,##0"), i, 6
        SetText vasList, Format(a + b, "#,##0"), i, 5
        SetText vasList, Format(a + b + c, "#,##0"), i, 7
        
        aSum = aSum + a
        bSum = bSum + b
        cSum = cSum + c
    Next i
    i = vasList.DataRowCnt + 1
    SetText vasList, Format(aSum, "#,##0"), i + 1, 3
    SetText vasList, Format(bSum, "#,##0"), i + 1, 4
    SetText vasList, Format(cSum, "#,##0"), i + 1, 6
    SetText vasList, Format(aSum + bSum, "#,##0"), i + 1, 5
    SetText vasList, Format(aSum + bSum + cSum, "#,##0"), i + 1, 7
    
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
    Dim sDate
    
    sDate = GetDateFull
    dtpSch2.Value = Format(CDate(sDate), "yyyy/mm/dd")
    dtpSch1.Value = Format(CDate(sDate), "yyyy/mm/01")
    
End Sub
