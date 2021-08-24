VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmStatistics 
   Caption         =   "결과대장 및 검사건수 통계"
   ClientHeight    =   10650
   ClientLeft      =   -75
   ClientTop       =   315
   ClientWidth     =   15240
   Icon            =   "STATISTICS.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   10650
   ScaleWidth      =   15240
   StartUpPosition =   2  '화면 가운데
   Begin VB.Frame fraTab1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9615
      Index           =   1
      Left            =   3825
      TabIndex        =   9
      Top             =   2010
      Width           =   14610
      Begin Threed.SSFrame SSFrame4 
         Height          =   1470
         Index           =   1
         Left            =   225
         TabIndex        =   10
         Top             =   150
         Width           =   5565
         _Version        =   65536
         _ExtentX        =   9816
         _ExtentY        =   2593
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.ComboBox cboLevel 
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "STATISTICS.frx":6852
            Left            =   1620
            List            =   "STATISTICS.frx":685C
            Style           =   2  '드롭다운 목록
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   840
            Visible         =   0   'False
            Width           =   2025
         End
         Begin MSComCtl2.DTPicker DTP1 
            Height          =   330
            Index           =   0
            Left            =   1620
            TabIndex        =   11
            Top             =   360
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   582
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
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   54525955
            CurrentDate     =   37685
         End
         Begin Threed.SSPanel SSPanel1 
            Height          =   315
            Index           =   0
            Left            =   345
            TabIndex        =   14
            Top             =   360
            Width           =   1245
            _Version        =   65536
            _ExtentX        =   2196
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "작업일자"
            ForeColor       =   8454143
            BackColor       =   8388608
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.01
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   1
            BevelOuter      =   0
         End
         Begin Threed.SSPanel pnlLevel 
            Height          =   315
            Left            =   345
            TabIndex        =   15
            Top             =   840
            Visible         =   0   'False
            Width           =   1245
            _Version        =   65536
            _ExtentX        =   2196
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "구      분"
            ForeColor       =   8454143
            BackColor       =   8388608
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.01
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   1
            BevelOuter      =   0
         End
         Begin MSComCtl2.DTPicker DTP1 
            Height          =   330
            Index           =   1
            Left            =   3615
            TabIndex        =   12
            Top             =   360
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   582
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
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   54525955
            CurrentDate     =   37685
         End
         Begin VB.Label Label1 
            Alignment       =   2  '가운데 맞춤
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   12
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3225
            TabIndex        =   16
            Top             =   420
            Width           =   345
         End
      End
      Begin Threed.SSFrame SSFrame7 
         Height          =   1470
         Index           =   1
         Left            =   5790
         TabIndex        =   17
         Top             =   150
         Width           =   8670
         _Version        =   65536
         _ExtentX        =   15293
         _ExtentY        =   2593
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin Threed.SSCommand cmdSearch 
            Height          =   615
            Index           =   1
            Left            =   390
            TabIndex        =   18
            Top             =   450
            Width           =   3000
            _Version        =   65536
            _ExtentX        =   5292
            _ExtentY        =   1085
            _StockProps     =   78
            Caption         =   "결과대장 조회   F3"
            ForeColor       =   8388608
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSCommand cmdExit 
            Height          =   615
            Index           =   1
            Left            =   6795
            TabIndex        =   19
            Top             =   450
            Width           =   1695
            _Version        =   65536
            _ExtentX        =   2990
            _ExtentY        =   1085
            _StockProps     =   78
            Caption         =   "닫 기  ESC"
            ForeColor       =   4210752
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSCommand cmdPrint 
            Height          =   615
            Index           =   1
            Left            =   3420
            TabIndex        =   20
            Top             =   450
            Width           =   3000
            _Version        =   65536
            _ExtentX        =   5292
            _ExtentY        =   1085
            _StockProps     =   78
            Caption         =   "결과대장  EXCEL 변환   F5"
            ForeColor       =   16512
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   1
            RoundedCorners  =   0   'False
         End
      End
      Begin FPSpread.vaSpread spdList 
         Height          =   7710
         Left            =   240
         TabIndex        =   24
         Top             =   1770
         Width           =   14175
         _Version        =   393216
         _ExtentX        =   25003
         _ExtentY        =   13600
         _StockProps     =   64
         BackColorStyle  =   1
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
         MaxRows         =   10
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "STATISTICS.frx":6878
      End
   End
   Begin VB.Frame fraTab1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9615
      Index           =   2
      Left            =   345
      TabIndex        =   1
      Top             =   480
      Width           =   14640
      Begin Threed.SSFrame SSFrame4 
         Height          =   945
         Index           =   0
         Left            =   150
         TabIndex        =   2
         Top             =   150
         Width           =   3225
         _Version        =   65536
         _ExtentX        =   5689
         _ExtentY        =   1667
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin MSComCtl2.DTPicker dtpWDate 
            Height          =   315
            Left            =   1440
            TabIndex        =   3
            Top             =   360
            Width           =   1590
            _ExtentX        =   2805
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
            CustomFormat    =   "yyy-MM"
            Format          =   54525955
            CurrentDate     =   36165
         End
         Begin Threed.SSPanel pnlLabDate 
            Height          =   315
            Left            =   165
            TabIndex        =   4
            Top             =   360
            Width           =   1245
            _Version        =   65536
            _ExtentX        =   2196
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "작업년월"
            ForeColor       =   8454143
            BackColor       =   8388608
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.01
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   1
            BevelOuter      =   0
         End
      End
      Begin Threed.SSFrame SSFrame7 
         Height          =   945
         Index           =   0
         Left            =   3375
         TabIndex        =   5
         Top             =   150
         Width           =   11100
         _Version        =   65536
         _ExtentX        =   19579
         _ExtentY        =   1667
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin Threed.SSCommand cmdSearch 
            Height          =   585
            Index           =   2
            Left            =   285
            TabIndex        =   6
            Top             =   210
            Width           =   3465
            _Version        =   65536
            _ExtentX        =   6112
            _ExtentY        =   1032
            _StockProps     =   78
            Caption         =   "통계 DATA LIST 조회   F3"
            ForeColor       =   8388608
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSCommand cmdPrint 
            Height          =   585
            Index           =   2
            Left            =   4005
            TabIndex        =   7
            Top             =   210
            Width           =   3525
            _Version        =   65536
            _ExtentX        =   6218
            _ExtentY        =   1032
            _StockProps     =   78
            Caption         =   "통계 DATA  LIST EXCEL  F5"
            ForeColor       =   16512
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   1
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSCommand cmdExit 
            Height          =   585
            Index           =   2
            Left            =   9120
            TabIndex        =   8
            Top             =   210
            Width           =   1755
            _Version        =   65536
            _ExtentX        =   3096
            _ExtentY        =   1032
            _StockProps     =   78
            Caption         =   "닫 기  ESC"
            ForeColor       =   4210752
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   1
            RoundedCorners  =   0   'False
         End
      End
      Begin FPSpread.vaSpread spdPrint 
         Height          =   3255
         Left            =   1350
         TabIndex        =   22
         Top             =   5220
         Visible         =   0   'False
         Width           =   9015
         _Version        =   393216
         _ExtentX        =   15901
         _ExtentY        =   5741
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   50
         MaxRows         =   40
         ShadowColor     =   16777215
         ShadowDark      =   16777215
         SpreadDesigner  =   "STATISTICS.frx":6CD6
      End
      Begin FPSpread.vaSpread spdDay 
         Height          =   8250
         Left            =   165
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1215
         Width           =   14310
         _Version        =   393216
         _ExtentX        =   25241
         _ExtentY        =   14552
         _StockProps     =   64
         BackColorStyle  =   1
         ColsFrozen      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   42
         MaxRows         =   5
         OperationMode   =   1
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "STATISTICS.frx":7865
         UserResize      =   0
         VisibleCols     =   34
         TextTip         =   1
      End
   End
   Begin MSComctlLib.TabStrip Tab1 
      Height          =   10155
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   15090
      _ExtentX        =   26617
      _ExtentY        =   17912
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "결과대장 조회"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "검사건수 통계"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  '가운데 맞춤
      Height          =   225
      Left            =   135
      TabIndex        =   21
      Top             =   10335
      Width           =   15015
   End
End
Attribute VB_Name = "frmStatistics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'결과대장
Private Type IFTESTITEM
    IFSEQ   As String
    TESTNM  As String
    COL     As Integer
End Type
Dim pTestItem() As IFTESTITEM

Dim pItemCnt    As Integer

'검사항목 통계
Private Type TspdDay
    sDay    As String
    sSpdCol As Long
End Type
Private Type TspdWeek
    sSpdCol  As Integer
End Type

Dim tspdDayi(42)    As TspdDay
Dim tspdWeeki(6)    As TspdWeek


'for tab strip
Dim iCurFrame   As Integer

Private Sub Disp_RstData()
    On Error GoTo ErrHandler
    
    Dim ADORS   As New ADODB.Recordset
    Dim sSql    As String
    Dim ii      As Integer
    Dim iCol    As Integer
    Dim iRow    As Integer
    Dim vTmp    As Variant
    Dim sDate   As String
    Dim tmpJNo  As String
    Dim iSRow   As Integer
    
    spdList.MaxRows = 0
    
    MousePointer = vbHourglass
    
    sSql = " SELECT WDATE, WSEQ, JNO, IFSEQ, RESULT1, name "
    sSql = sSql & " FROM IFRESULT "
    sSql = sSql & " Where WDATE BETWEEN '" & Format(DTP1(0).Value, "YYYYMMDD") & "'"
    sSql = sSql & "   AND '" & Format(DTP1(1).Value, "YYYYMMDD") & "'"
    If cboLevel.ListIndex = 1 Then
        sSql = sSql & "   AND LEFT(jno,1) = 'Q' "
    Else
        sSql = sSql & "   AND LEN(jno) >= " & Val(gOrdCfg.sFSize(3)) & ""
        sSql = sSql & "   AND LEFT(jno,3) <> 'ERR' "
        sSql = sSql & "   AND LEFT(jno,1) <> 'Q' "
    End If
    sSql = sSql & "   and result1 <> '' "
    sSql = sSql & " ORDER BY WDATE, WSEQ, IFSEQ "
    
    ADORS.Open sSql, fGetCurDSN(gsMachineCd), adOpenForwardOnly
    
    If ADORS.EOF = True Then
        ADORS.Close: Set ADORS = Nothing
        MousePointer = vbDefault
        MsgBox "해당자료가 존재하지 않습니다.", vbInformation
        Exit Sub
    End If
    
    ADORS.MoveFirst
    iSRow = 0
    Do While Not ADORS.EOF
        iCol = 0
        For ii = 1 To spdList.MaxCols - 1
            With pTestItem(ii)
                If Trim(ADORS.Fields(3)) = Trim(.IFSEQ) Then
                    iCol = .COL
                    Exit For
                End If
            End With
        Next ii
        
        If iCol > 0 Then
            With spdList
                If Trim(sDate) = "" Or Trim(ADORS.Fields(0) & "") <> sDate Then
                    .MaxRows = .MaxRows + 1
                    .RowHeight(.MaxRows) = 11.5
                    
                    Call .SetText(2, .MaxRows, Format(ADORS.Fields(0), "####-##-##"))
                    
                    .COL = 2: .Row = .MaxRows
                    .TypeHAlign = TypeHAlignCenter
                    
                    .COL = 2: .Col2 = .MaxCols
                    .Row = .MaxRows: .Row2 = .MaxRows
                    .BlockMode = True
                    .BackColor = RGB(235, 255, 235)
                    .BlockMode = False
                    
                    iSRow = .MaxRows
                End If
                sDate = Trim(ADORS.Fields(0) & "")
            
                iRow = 0
                For ii = iSRow + 1 To .MaxRows
                    Call .GetText(1, ii, vTmp)
                    If Trim(vTmp) = Trim(ADORS.Fields(1) & "") Then
                        Call .GetText(iCol, ii, vTmp)
                        If Trim(vTmp) <> "" Then
                            Exit For
                        End If
                        
                        iRow = ii
                        
                        Exit For
                    End If
                Next ii
                
                If iRow = 0 Then
                    .MaxRows = .MaxRows + 1
                    .RowHeight(.MaxRows) = 11.5
                    
                    iRow = .MaxRows
                End If
                
                Call .SetText(1, iRow, Trim(ADORS.Fields(1) & ""))
                Call .SetText(2, iRow, Space(1) & Trim(ADORS.Fields(2) & ""))
                Call .SetText(3, iRow, Trim(ADORS.Fields(5) & ""))
                Call .SetText(iCol, iRow, Trim(ADORS.Fields(4) & ""))
            End With
        End If
        
        ADORS.MoveNext
    Loop
    
    ADORS.Close: Set ADORS = Nothing
    
    MousePointer = vbDefault
    
    Exit Sub
ErrHandler:
    MousePointer = vbDefault
    If ADORS.State = 1 Then
        ADORS.Close: Set ADORS = Nothing
    End If
    MsgBox Err.Description

End Sub

Private Sub Disp_Statistics()
    Dim i       As Integer
    Dim iCnt    As Integer
    Dim iWCnt   As Integer
    Dim iWTCnt  As Integer
    
    Dim iWeek   As Integer
    Dim sSDATE  As String
    Dim sEDATE  As String
    Dim sTmpDate    As String
    Dim iDiff   As Integer
    Dim ispdRealCnt As Integer
    Dim iCol    As Integer
    
    On Error GoTo ErrHandler

    With spdDay
        .ReDraw = False

        .MaxRows = 0
        For i = 3 To .MaxCols - 1
            .ColWidth(i) = 0
        Next i
    End With
    
    Call Get_TestItem2
    
    
    Dim sYear$, sMon$, sNYear$, sNMon$
    
    sYear = dtpWDate.Year
    sMon = dtpWDate.Month
    
    sSDATE = Format(sYear, "0000") & "-" & Format(sMon, "00") & "-" & "01"
    
    sEDATE = DateAdd("m", 1, sSDATE)
    sEDATE = DateAdd("d", -1, sEDATE)
    iWeek = 0
    
    iDiff = DateDiff("d", sSDATE, sEDATE)
    
    '/---------------------- 스프레드 설정(날짜구간)
    iWCnt = 0
    ispdRealCnt = 0
    iCol = 2
    
    For i = 0 To iDiff
        sTmpDate = DateAdd("d", i, sSDATE)
        iWeek = Weekday(sTmpDate)
'''        spdDay.ColWidth(i + 3) = 8
         
        If iWeek = 1 Then
'            If sTmpDate <> sSDATE Then
                
                iCol = iCol + 1
                spdDay.ColWidth(iCol) = 8
                Call spdDay.SetText(iCol, 0, Format(CDate(sTmpDate), "DD") & "일")
                
                ispdRealCnt = ispdRealCnt + 1
                tspdDayi(ispdRealCnt).sDay = Format(CDate(sTmpDate), "YYYYMMDD")
                tspdDayi(ispdRealCnt).sSpdCol = (iCol)
                
                iCol = iCol + 1
                spdDay.ColWidth(iCol) = 8
                ispdRealCnt = ispdRealCnt + 1
                iWCnt = iWCnt + 1
                Call spdDay.SetText(iCol, 0, iWCnt & "주합계")
                tspdWeeki(iWCnt).sSpdCol = (iCol)
                Call spdReverse(spdDay, iCol, iCol, 1, spdDay.MaxRows, 연빨강)
                
                tspdDayi(ispdRealCnt).sDay = ""
                tspdDayi(ispdRealCnt).sSpdCol = (iCol)
                
'            Else
'                spdDay.ColWidth(i + 3) = 0
'            End If
        Else
            iCol = iCol + 1
            spdDay.ColWidth(iCol) = 8
            Call spdDay.SetText(iCol, 0, Format(CDate(sTmpDate), "DD") & "일")
            ispdRealCnt = ispdRealCnt + 1
            tspdDayi(ispdRealCnt).sDay = Format(CDate(sTmpDate), "YYYYMMDD")
            tspdDayi(ispdRealCnt).sSpdCol = (iCol)
        End If
    
    Next i
    
    iWeek = Weekday(sEDATE)
    
    If iWeek <> 1 Then
        iCol = iCol + 1
        spdDay.ColWidth(iCol) = 8
        ispdRealCnt = ispdRealCnt + 1
        iWCnt = iWCnt + 1
        Call spdDay.SetText(iCol, 0, iWCnt & "주합계")
        tspdWeeki(iWCnt).sSpdCol = (iCol)
        Call spdReverse(spdDay, iCol, iCol, 1, spdDay.MaxRows, 연빨강)
    End If
    '스프레드 설정      ----------------------------/
    
    Call Get_Statistic_Data(Format(sYear, "0000") & Format(sMon, "00"))
''    Call Get_Statistic_Data(sYear & sMon)
    Call Week_Sum_Data
    
    spdDay.ReDraw = True
    
    Exit Sub
ErrHandler:
    If spdDay.ReDraw <> True Then
        spdDay.ReDraw = True
    End If
End Sub

Private Sub Format_Screen()

    spdList.MaxRows = 0
    
    spdDay.MaxRows = 0
    spdPrint.MaxRows = 0
    
End Sub

Private Sub Get_TestItem2()

    Dim adoCn   As New ADODB.Connection
    Dim ADORS   As New ADODB.Recordset
    Dim sSql    As String
    Dim i       As Integer
    Dim j       As Integer
    
    On Error GoTo ErrHandler
    
    adoCn.Open "" & fGetCurDSN(gsMachineCd) & "", "", ""
    
    sSql = "SELECT IFTESTSEQ, IFTESTNM, IFORDCD "
    sSql = sSql & " FROM IFTESTITEM "
    sSql = sSql & " WHERE IFSVRCD <> 'CO'"
    
    ADORS.Open sSql, adoCn, adOpenStatic
    
    j = ADORS.RecordCount
    
    If j < 1 Then
        Exit Sub
    End If
        
    spdDay.MaxRows = j
    
    ADORS.MoveFirst
    
    With spdDay
        For i = 1 To j
            Call .SetText(1, i, ADORS(0))
            Call .SetText(2, i, ADORS(1))
            
            ADORS.MoveNext
        Next i
    End With
    
    ADORS.Close
    Set adoCn = Nothing
    
''    Call spdDay.SetText(2, spdDay.MaxRows, "  >> 합  계 <<")
    
    Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub Get_TestItem1()
    On Error GoTo ErrHandler
    
    Dim ADORS   As New ADODB.Recordset
    Dim sSql    As String
    Dim ii      As Integer
    Dim iItemCnt%, iCalCnt%
    
    spdList.MaxCols = 3     '2
    
    '일반검사항목
    sSql = " SELECT IFTESTSEQ, IFTESTNM "
    sSql = sSql & " FROM IFTESTITEM "
    sSql = sSql & " ORDER BY IFTESTSEQ "
    
    ADORS.Open sSql, fGetCurDSN(gsMachineCd), adOpenStatic
    
    If ADORS.RecordCount = 0 Then
        ADORS.Close: Set ADORS = Nothing
        Exit Sub
    End If
    
    iItemCnt = ADORS.RecordCount
    ReDim pTestItem(iItemCnt) As IFTESTITEM
    
    ADORS.MoveFirst
    With spdList
        ii = 0
        Do While Not ADORS.EOF
            .MaxCols = .MaxCols + 1
            .COL = .MaxCols: .Row = -1
            .CellType = CellTypeStaticText
            .TypeVAlign = TypeVAlignCenter
            .TypeHAlign = TypeHAlignRight
            
            ii = ii + 1
            
            pTestItem(ii).IFSEQ = Trim(ADORS.Fields(0))
            pTestItem(ii).TESTNM = Trim(ADORS.Fields(1))
            pTestItem(ii).COL = .MaxCols
            
            Call .SetText(.MaxCols, 0, pTestItem(ii).TESTNM)
            
            ADORS.MoveNext
        Loop
    End With
    ADORS.Close: Set ADORS = Nothing
    
    '계산식 항목
    sSql = " SELECT IFTESTSEQ, IFTESTNM "
    sSql = sSql & " FROM CALTESTITEM "
    sSql = sSql & " ORDER BY IFTESTSEQ "
    
    ADORS.Open sSql, fGetCurDSN(gsMachineCd), adOpenStatic
    
    If ADORS.RecordCount = 0 Then
        ADORS.Close: Set ADORS = Nothing
        Exit Sub
    End If
    
    iCalCnt = ADORS.RecordCount
    ReDim Preserve pTestItem(iItemCnt + iCalCnt) As IFTESTITEM
    
    ADORS.MoveFirst
    With spdList
        Do While Not ADORS.EOF
            .MaxCols = .MaxCols + 1
            .COL = .MaxCols: .Row = -1
            .CellType = CellTypeStaticText
            .TypeVAlign = TypeVAlignCenter
            .TypeHAlign = TypeHAlignRight
            
            ii = ii + 1
            
            pTestItem(ii).IFSEQ = Trim(ADORS.Fields(0))
            pTestItem(ii).TESTNM = Trim(ADORS.Fields(1))
            pTestItem(ii).COL = .MaxCols
            
            Call .SetText(.MaxCols, 0, pTestItem(ii).TESTNM)
            
            ADORS.MoveNext
        Loop
    End With
    ADORS.Close: Set ADORS = Nothing
    
    Exit Sub
ErrHandler:
    If ADORS.State = 1 Then
        ADORS.Close: Set ADORS = Nothing
    End If
    MsgBox Err.Description, vbExclamation
End Sub
Private Function fExcelFilePath() As String
    Dim sTmp    As String
    
    sTmp = GetKeyValue(HKEY_CURRENT_USER, _
                "Software\Ack_if\Interface Config\" & gsMachineCd, "Excel.File.Path")
                
    fExcelFilePath = sTmp
    
End Function
Private Function fGetCurDSN(ByVal sBuf As String) As String
    Dim bRetVal As Boolean
    
    sBuf = GetKeyValue(HKEY_CURRENT_USER, "Software\Ack_if\Interface Config\" & sBuf, "DSN")
    
    If sBuf = "" Then
        bRetVal = UpdateKey(HKEY_CURRENT_USER, "Software\Ack_if\Program Config\" & sBuf, "DSN", "IFDSN")
        
        If bRetVal = True Then
            fGetCurDSN = "IFDSN"
        Else
            MsgBox "레지스트리키의 초기화 작업에 에러가 발생했습니다!!"
            fGetCurDSN = "IFDSN"
        End If
    Else
        fGetCurDSN = sBuf
    End If
End Function
Private Sub Print_RstData()
    On Error GoTo ErrRtn
    
    Dim sFileNm As String
    Dim sLogNm  As String
    Dim sPath   As String
    
    If spdList.MaxRows = 0 Then
        MsgBox "해당자료가 존재하지 않습니다.", vbInformation
        Exit Sub
    End If
    
    sPath = fExcelFilePath
    If Right(sPath, 1) <> "\" Then
        sPath = sPath & "\"
    End If
    
    sFileNm = sPath & "결과대장_" & Format(Now, "yyyymmdd") & ".xls"
    sLogNm = sPath & "결과대장.log"
    
    With spdList
        Call .ExportToHTML(sFileNm, False, sLogNm)

        If Trim(fExcelExePath) <> "" Then
            Call WinExec(fExcelExePath & " " & sFileNm, 3)
        Else
            Me.MousePointer = vbDefault
            MsgBox "해당 PC에 EXCEL 프로그램이 존재하지 않습니다!!"

            Exit Sub
        End If
    End With
    
    Exit Sub
ErrRtn:
    MsgBox Err.Description, vbExclamation
End Sub

Private Function fExcelExePath() As String
    Dim sTmp    As String
    
    sTmp = GetKeyValue(HKEY_CURRENT_USER, _
                "Software\Ack_if\Interface Config\" & gsMachineCd, "Excel.Exe.Path")
                
    fExcelExePath = sTmp
    
End Function
Private Sub Print_Statistics()
    On Error GoTo ErrHandler
    
    Dim bChk    As Boolean
    Dim sFileNm$
    Dim sPath$
    
    bChk = Create_Print_Spread
    
    If bChk = False Then
        MsgBox "출력할 데이타가 존재하지 않습니다.", vbExclamation
        Exit Sub
    End If
    
    sPath = GetExcelFilePath
    If Right(sPath, 1) <> "\" Then
        sPath = sPath & "\"
    End If
    
    sFileNm = sPath & "FST01_" & Format(dtpWDate.Value, "YYYYMM") & ".xls"
    
    With spdPrint
        Call .ExportToHTML(sFileNm, False, sPath & "FST01.log")
            
        If Trim(GetExcelExePath) <> "" Then
            Call WinExec(GetExcelExePath & " " & sFileNm, 3)
        Else
            Me.MousePointer = vbDefault
            MsgBox "해당 PC에 EXCEL 프로그램이 존재하지 않습니다!!", vbExclamation
            Exit Sub
        End If
    End With
    
ErrHandler:
    If Err <> 0 Then
        MsgBox Err.Description, vbExclamation
    End If
End Sub

Private Sub cboLevel_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{tab}"
    End If
End Sub


Private Sub cmdExit_Click(Index As Integer)
    Unload Me
End Sub
Private Sub cmdPrint_Click(Index As Integer)

    If Index = 1 Then
        Call Print_RstData
    Else
        Call Print_Statistics
    End If
    
End Sub
Private Function GetExcelExePath()
    On Error GoTo ErrHandler
    
    Dim sBuf$
    
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "Excel.Exe.Path")
    
    GetExcelExePath = sBuf
    
    Exit Function
ErrHandler:
    lblMsg = "GetExcelExePath - Err(" & Err.Description & ")"
End Function
Private Function GetExcelFilePath()
    On Error GoTo ErrHandler
    
    Dim sBuf$
    
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & gsMachineCd, "Excel.File.Path")
    
    GetExcelFilePath = sBuf
    
    Exit Function
ErrHandler:
    lblMsg = "GetExcelFilePath - Err(" & Err.Description & ")"
End Function
Private Function Create_Print_Spread() As Boolean
    Dim i       As Integer
    Dim j       As Integer
    Dim vTmp
    Dim iRow    As Integer
    Dim iCol    As Integer
    
    On Error GoTo ErrHandler
    
    Create_Print_Spread = True
    
    If spdDay.MaxRows < 1 Then
        Create_Print_Spread = False
        Exit Function
    End If
    
    With spdPrint
        spdPrint.MaxRows = 0
        For i = 1 To .MaxCols - 1
            .ColWidth(i) = 0
        Next i

        Call .SetText(1, 0, "날짜")
        
        spdPrint.MaxRows = spdDay.MaxCols
        
        iCol = 0
        
        For i = 3 To spdDay.MaxCols
            If spdDay.ColWidth(i) > 0 Then
                iCol = iCol + 1
                For j = 0 To spdDay.MaxRows
                    spdPrint.RowHeight(iCol) = 18
                    Call spdDay.GetText(i, j, vTmp)
                    Call .SetText(j, iCol, vTmp)
                Next j
            End If
        Next i
        
        spdPrint.MaxRows = iCol
        
        For i = 1 To spdDay.MaxRows
            Call spdDay.GetText(2, i, vTmp)
            .ColWidth(i) = 8
            
            Call .SetText(i, 0, vTmp)
        Next i
        
        For iRow = 1 To .MaxRows
            Call .GetText(0, iRow, vTmp)
            If Trim(vTmp) <> "" Then
                If Right(Trim(vTmp), 3) = "주합계" Then
                    Call spdReverse(spdPrint, 1, .MaxCols, iRow, iRow, 연하늘, 2)
                ElseIf Trim(vTmp) = "총합계" Then
                    Call spdReverse(spdPrint, 1, .MaxCols, iRow, iRow, 흐린파랑, 2)
                End If
            End If
        Next iRow
        
    End With

    Exit Function
ErrHandler:
    MsgBox Err.Description
End Function


Private Sub cmdSearch_Click(Index As Integer)

    If Index = 1 Then
        Call Disp_RstData
    Else
        Call Disp_Statistics
    End If
    
End Sub
Private Sub Get_Statistic_Data(ByVal sDate As String)
    Dim adoCn   As New ADODB.Connection
    Dim ADORS   As New ADODB.Recordset
    Dim sSql    As String
    Dim iCol    As Integer
    Dim iRow    As Integer
    Dim j       As Integer
    Dim k       As Integer
    Dim vTmp
    Dim sTmp    As String
    Dim bChk    As Boolean
    
    On Error GoTo ErrHandler
    
    If spdDay.MaxRows < 1 Then Exit Sub
    
    
    adoCn.Open "" & fGetCurDSN(gsMachineCd) & "", "", ""
    
    sSql = "SELECT A.WDATE, A.IFSEQ, B.IFTESTNM, COUNT(A.IFSEQ) AS CNT "
    sSql = sSql & " FROM IFRESULT AS A, IFTESTITEM AS B "
    sSql = sSql & " Where A.IFSEQ = B.IFTESTSEQ "
    sSql = sSql & " AND   left(A.WDATE, 6) = '" & sDate & "'"
    sSql = sSql & " AND   B.IFSVRCD <> 'CO' "
    sSql = sSql & " GROUP BY A.WDATE, A.IFSEQ, B.IFTESTNM "
    sSql = sSql & " ORDER BY A.WDATE, A.IFSEQ "
    
    ADORS.Open sSql, adoCn, adOpenStatic
    
    j = ADORS.RecordCount
    
    If j < 1 Then
        Exit Sub
    End If
    
    ADORS.MoveFirst
    
    With spdDay
    For k = 1 To j
        For iCol = 3 To .MaxCols
            bChk = False
            
            If tspdDayi(iCol - 2).sDay = ADORS(0) Then
                For iRow = 1 To .MaxRows
                    Call .GetText(1, iRow, vTmp)
                    If CStr(vTmp) = ADORS(1) Then
                    
                        Call .SetText(tspdDayi(iCol - 2).sSpdCol, iRow, Format(ADORS(3), "#,##0"))
                        bChk = True
                        ADORS.MoveNext
                        Exit For
                    End If
                Next iRow
            End If
            
            If bChk = True Then
                Exit For
            End If
        Next iCol
        
    Next k
    End With
    ADORS.Close
    Set adoCn = Nothing
    
    Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub Week_Sum_Data()
    On Error GoTo ErrHandler
    Dim dSubSum     As Long
    Dim dTotSum     As Long
    Dim vTmp
    Dim i, j, iRow  As Integer
    
    With spdDay
        If tspdWeeki(1).sSpdCol > 0 Then
            For iRow = 1 To .MaxRows
                dTotSum = 0
                dSubSum = 0
                For j = 3 To tspdWeeki(1).sSpdCol - 1
                
                    Call .GetText(j, iRow, vTmp)
                    If Trim(vTmp) <> "" Then
                        dSubSum = dSubSum + CDbl(vTmp)
                    End If
                Next j
                
                If dSubSum > 0 Then
                    Call .SetText(tspdWeeki(1).sSpdCol, iRow, dSubSum)
                    dTotSum = dTotSum + dSubSum
                End If
                
                For i = 2 To 6
                    dSubSum = 0
                    If CInt(tspdWeeki(i).sSpdCol) > 0 Then
                        For j = tspdWeeki(i - 1).sSpdCol + 1 To tspdWeeki(i).sSpdCol - 1
                            Call .GetText(j, iRow, vTmp)
                            If Trim(vTmp) <> "" Then
                                dSubSum = dSubSum + CDbl(vTmp)
                            End If
                        Next j
                    End If
                    
                    If dSubSum > 0 Then
                        Call .SetText(tspdWeeki(i).sSpdCol, iRow, dSubSum)
                        dTotSum = dTotSum + dSubSum
                    End If
                Next i
                
                Call .SetText(.MaxCols, iRow, dTotSum)
                
            Next iRow
        End If
    End With
    
    Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub Form_Load()
    On Error GoTo ErrHandler
    
    Call GetOrdRstCfg
    
    fraTab1(1).Left = 345: fraTab1(1).Top = 480
    fraTab1(2).Left = 345: fraTab1(2).Top = 480
    fraTab1(2).Enabled = False: fraTab1(2).Visible = False
    
    iCurFrame = 1
    
    'For 결과대장
    DTP1(0).Value = Now
    DTP1(1).Value = Now
    
    Call Get_TestItem1
    
    spdList.MaxRows = 0
    cboLevel.ListIndex = 0
    
    'For 건수통계
    dtpWDate.Value = Now
    
    spdDay.MaxRows = 0
    spdPrint.MaxRows = 0

'    If Right(gsMachineCd, 3) = "001" Or Right(gsMachineCd, 3) = "002" Then
'        pnlLevel.Visible = True
'        cboLevel.Visible = True: cboLevel.TabStop = True
'    End If

    Exit Sub
ErrHandler:
    MsgBox "LOAD ERR - (" & Err.Description & ")"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyF3
            Call cmdSearch_Click(Tab1.SelectedItem.Index)
        Case vbKeyF5
            Call cmdPrint_Click(Tab1.SelectedItem.Index)
        Case vbKeyEscape
            Call cmdExit_Click(Tab1.SelectedItem.Index)
        Case Else
    End Select
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call RegEditCurFrmTitle("Statistics", "")
    lblMsg = ""
End Sub

Private Sub Tab1_Click()
    
    fraTab1(Tab1.SelectedItem.Index).Visible = True
    fraTab1(Tab1.SelectedItem.Index).Enabled = True
    fraTab1(iCurFrame).Visible = False
    fraTab1(iCurFrame).Enabled = False
    
    iCurFrame = Tab1.SelectedItem.Index
    
    Call Format_Screen
    
End Sub
