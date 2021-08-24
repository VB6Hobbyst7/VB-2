VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmStatistics 
   Caption         =   "결과대장 및 양성율 조회"
   ClientHeight    =   12285
   ClientLeft      =   -75
   ClientTop       =   315
   ClientWidth     =   15240
   Icon            =   "STATISTICS.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   12285
   ScaleWidth      =   15240
   Begin FPSpread.vaSpread spdExcel 
      Height          =   2115
      Left            =   7725
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   9735
      Visible         =   0   'False
      Width           =   6795
      _Version        =   393216
      _ExtentX        =   11986
      _ExtentY        =   3731
      _StockProps     =   64
      BackColorStyle  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   2
      MaxRows         =   9
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "STATISTICS.frx":6852
   End
   Begin FPSpread.vaSpread spdBottom 
      Height          =   3240
      Left            =   135
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   6810
      Width           =   15015
      _Version        =   393216
      _ExtentX        =   26485
      _ExtentY        =   5715
      _StockProps     =   64
      BackColorStyle  =   1
      ColsFrozen      =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   6
      MaxRows         =   10
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "STATISTICS.frx":6C43
   End
   Begin Threed.SSFrame SSFrame4 
      Height          =   1185
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   15
      Width           =   15045
      _Version        =   65536
      _ExtentX        =   26538
      _ExtentY        =   2090
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
      Begin VB.ComboBox cboRange 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "STATISTICS.frx":7234
         Left            =   4785
         List            =   "STATISTICS.frx":7244
         Style           =   2  '드롭다운 목록
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   690
         Width           =   945
      End
      Begin VB.ComboBox cboGbn 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "STATISTICS.frx":7260
         Left            =   1455
         List            =   "STATISTICS.frx":726D
         Style           =   2  '드롭다운 목록
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   690
         Width           =   1515
      End
      Begin MSComCtl2.DTPicker DTP1 
         Height          =   330
         Index           =   0
         Left            =   1455
         TabIndex        =   0
         Top             =   255
         Width           =   1515
         _ExtentX        =   2672
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
         Format          =   56885251
         CurrentDate     =   39450
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   315
         Index           =   0
         Left            =   345
         TabIndex        =   11
         Top             =   255
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "작업구간"
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
      Begin Threed.SSPanel SSPanel1 
         Height          =   315
         Index           =   1
         Left            =   345
         TabIndex        =   12
         Top             =   675
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "조회구분"
         ForeColor       =   8454143
         BackColor       =   49152
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
      Begin Threed.SSCommand cmdSearch 
         Height          =   840
         Left            =   6600
         TabIndex        =   4
         Top             =   210
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2469
         _ExtentY        =   1482
         _StockProps     =   78
         Caption         =   "LIST 조회(&L)"
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
         Picture         =   "STATISTICS.frx":7294
      End
      Begin Threed.SSCommand cmdExcel 
         Height          =   840
         Left            =   9450
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   210
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2469
         _ExtentY        =   1482
         _StockProps     =   78
         Caption         =   "Excel 변환"
         ForeColor       =   12582912
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
         Picture         =   "STATISTICS.frx":7F6E
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   840
         Left            =   13290
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   210
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2469
         _ExtentY        =   1482
         _StockProps     =   78
         Caption         =   "닫 기"
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
         Picture         =   "STATISTICS.frx":8C48
      End
      Begin Threed.SSCommand cmdApply 
         Height          =   840
         Left            =   8025
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   210
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2469
         _ExtentY        =   1482
         _StockProps     =   78
         Caption         =   "변경사항 적용"
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
         Picture         =   "STATISTICS.frx":9C9A
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   15
         Top             =   675
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "정상범위"
         ForeColor       =   8454143
         BackColor       =   49152
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
         Left            =   3240
         TabIndex        =   1
         Top             =   255
         Width           =   1515
         _ExtentX        =   2672
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
         Format          =   56885251
         CurrentDate     =   39450
      End
      Begin VB.Label Label2 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
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
         Height          =   210
         Left            =   3030
         TabIndex        =   17
         Top             =   300
         Width           =   150
      End
      Begin VB.Label Label1 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         Caption         =   "-/+"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   4395
         TabIndex        =   16
         Top             =   750
         Width           =   345
      End
      Begin VB.Image imgMinus 
         Height          =   165
         Left            =   12195
         Picture         =   "STATISTICS.frx":A574
         Top             =   180
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Image imgPlus 
         Height          =   165
         Left            =   12015
         Picture         =   "STATISTICS.frx":A742
         Top             =   180
         Visible         =   0   'False
         Width           =   165
      End
   End
   Begin FPSpread.vaSpread spdList 
      Height          =   5790
      Left            =   135
      TabIndex        =   8
      Top             =   1290
      Width           =   15015
      _Version        =   393216
      _ExtentX        =   26485
      _ExtentY        =   10213
      _StockProps     =   64
      BackColorStyle  =   1
      ColsFrozen      =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   6
      MaxRows         =   25
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "STATISTICS.frx":A910
      TextTip         =   1
   End
   Begin MSComctlLib.StatusBar StBar1 
      Align           =   2  '아래 맞춤
      Height          =   360
      Left            =   0
      TabIndex        =   13
      Top             =   11925
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   26353
         EndProperty
      EndProperty
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
    DOT     As String
    JGBN    As String
End Type
Dim pTestItem() As IFTESTITEM

Dim pItemCnt    As Integer

Dim iSpdFixCol  As Integer

Private Function Chk_Condition() As Boolean
    
    Chk_Condition = False
    
    If Format(DTP1(0).Value, "YYYY-MM") <> Format(DTP1(1).Value, "YYYY-MM") Then
        MsgBox "조회구간 설정에 오류가 있습니다." & vbCrLf & "동일한 년/월을 선택해 주십시요.", vbExclamation, Me.Caption
        DTP1(0).SetFocus
        Exit Function
    End If
    
    Chk_Condition = True
    
End Function

Private Sub Get_TestItem1()
    On Error GoTo ErrHandler
    
    Dim ADORS   As New ADODB.Recordset
    Dim sSql    As String
    Dim ii      As Integer
    Dim iItemCnt%, iCalCnt%
    
    iSpdFixCol = 5
    
    spdList.MaxCols = iSpdFixCol
    spdBottom.MaxCols = iSpdFixCol
    spdExcel.MaxCols = iSpdFixCol - 4
    
    
    '일반검사항목
                   '0          1         2         3    4
    sSql = " SELECT IFTESTSEQ, IFTESTNM, dotdigit, lhu, judgegbn "
    sSql = sSql & " FROM IFTESTITEM "
    sSql = sSql & " ORDER BY IFTESTSEQ "
    
    ADORS.Open sSql, fGetCurDSN(gsMachineCd), adOpenForwardOnly
    
    If ADORS.EOF = True Then
        ADORS.Close: Set ADORS = Nothing
        Exit Sub
    End If
    
    Erase pTestItem()
    
    ADORS.MoveFirst
    With spdList
        .ReDraw = False
        
        Do While Not ADORS.EOF
            .MaxCols = .MaxCols + 1
            .ColWidth(.MaxCols) = 10
            
            .Col = .MaxCols: .Row = -1
            .CellType = CellTypeStaticText
            .TypeVAlign = TypeVAlignCenter: .TypeHAlign = TypeHAlignRight
            .TypeEllipses = True
            
            ii = .MaxCols
            
            ReDim Preserve pTestItem(ii)
            
            pTestItem(ii).IFSEQ = Trim(ADORS.Fields(0))
            pTestItem(ii).TESTNM = Trim(ADORS.Fields(1))
            If Trim(ADORS.Fields(3)) = "H" Then
                pTestItem(ii).DOT = Trim(ADORS.Fields(2))
            End If
            
            Select Case Trim(ADORS.Fields(4))
                Case "0", "1"
                    pTestItem(ii).JGBN = "1"
                Case "2", "3", "5", "6"
                    pTestItem(ii).JGBN = "2"
                Case Else
                    pTestItem(ii).JGBN = "0"
            End Select
            
            Call .SetText(.MaxCols, 0, pTestItem(ii).TESTNM)
            
            '<S--- spdBottom/spdExcel 표시
            With spdBottom
                .MaxCols = .MaxCols + 1
                
                Select Case pTestItem(ii).JGBN
                    Case "1"
                        Call .SetText(.MaxCols, 0, pTestItem(ii).TESTNM & vbCrLf & "(mean)")
                    Case "2"
                        Call .SetText(.MaxCols, 0, pTestItem(ii).TESTNM & vbCrLf & "(%)")
                    Case Else
                        Call .SetText(.MaxCols, 0, " ")
                End Select
                
                .Col = .MaxCols: .Row = -1
                .CellType = CellTypeStaticText
                .TypeHAlign = TypeHAlignCenter
                .TypeEllipses = True
                .ColWidth(.MaxCols) = spdList.ColWidth(.MaxCols)
            End With
            With spdExcel
                .MaxCols = .MaxCols + 1
                
                Select Case pTestItem(ii).JGBN
                    Case "1"
                        Call .SetText(.MaxCols, 0, pTestItem(ii).TESTNM & vbCrLf & "(mean)")
                    Case "2"
                        Call .SetText(.MaxCols, 0, pTestItem(ii).TESTNM & vbCrLf & "(%)")
                    Case Else
                        Call .SetText(.MaxCols, 0, " ")
                End Select
                
                .Col = .MaxCols: .Row = -1
'                .CellType = CellTypeStaticText
                .TypeHAlign = TypeHAlignCenter
                .ColWidth(.MaxCols) = spdList.ColWidth(spdList.MaxCols)
'                .BackColor = vbWhite
            End With
            '>E----------------------------
            
            ADORS.MoveNext
        Loop
        
        .ReDraw = True
    End With
    ADORS.Close: Set ADORS = Nothing
    
    
'    '계산식 항목
'    sSql = " SELECT IFTESTSEQ, IFTESTNM "
'    sSql = sSql & " FROM CALTESTITEM "
'    sSql = sSql & " ORDER BY IFTESTSEQ "
'
'    ADORS.Open sSql, fGetCurDSN(gsMachineCd), adOpenForwardOnly
'
'    If ADORS.EOF = True Then
'        ADORS.Close: Set ADORS = Nothing
'        Exit Sub
'    End If
'
'    ReDim Preserve pTestItem(iItemCnt + iCalCnt) As IFTESTITEM
'
'    ADORS.MoveFirst
'    With spdList
'        Do While Not ADORS.EOF
'            .MaxCols = .MaxCols + 1
'            .COL = .MaxCols: .Row = -1
'            .CellType = CellTypeStaticText
'            .TypeVAlign = TypeVAlignCenter: .TypeHAlign = TypeHAlignRight
'
'            ii = .MaxCols
'
'            ReDim Preserve pTestItem(ii)
'
'            pTestItem(ii).IFSEQ = Trim(ADORS.Fields(0))
'            pTestItem(ii).TESTNM = Trim(ADORS.Fields(1))
'            pTestItem(ii).COL = .MaxCols
'
'            Call .SetText(.MaxCols, 0, pTestItem(ii).TESTNM)
'
'            ADORS.MoveNext
'        Loop
'    End With
'    ADORS.Close: Set ADORS = Nothing
    
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

    Dim sRetVal As String
    
    'MS Access
    fGetCurDSN = "Driver={Microsoft Access Driver (*.mdb)};Dbq=[LOCALDB]"

    sRetVal = GetKeyValue(HKEY_CURRENT_USER, "Software\Ack_if\Interface Config\" & sBuf, "DSN")

    If sRetVal = "" Then
        sRetVal = App.Path & "\" & sBuf
        Call UpdateKey(HKEY_CURRENT_USER, "Software\Ack_if\Interface Config\" & sBuf, "DSN", sRetVal)
    End If
    fGetCurDSN = Replace(fGetCurDSN, "[LOCALDB]", Trim(sRetVal))
    
End Function
Private Sub GetPosRatio(ByVal sDate As String, ByVal iSRow As Integer, ByVal iERow As Integer)
    On Error GoTo ErrGetRatio
    
    Dim vTmp, vChk
    Dim iRow%, iCol%, iCnt%
    Dim sTmp$
    Dim sgTmp   As Single
    Dim sDispRst$, sDispRst2$
    
    With spdBottom
        .MaxRows = .MaxRows + 1
        .RowHeight(.MaxRows) = 11
        
        Call .SetText(4, .MaxRows, sDate)
    End With
    
    With spdList
        For iCol = iSpdFixCol + 1 To .MaxCols
            If Trim(pTestItem(iCol).TESTNM) <> "" Then
                sgTmp = 0
                iCnt = 0
                
                For iRow = iSRow To iERow
                    Call .GetText(2, iRow, vChk)
                    If vChk = vbChecked Then
                        '<S--- 선택된 자료만 양성율 계산
                        Call .GetText(iCol, iRow, vTmp)
                        If Trim(vTmp) <> "" Then
                            iCnt = iCnt + 1
                            
                            If pTestItem(iCol).JGBN = "1" Then
                                '정량
                                sgTmp = sgTmp + Val(vTmp)
                            ElseIf pTestItem(iCol).JGBN = "2" Then
                                '정성(Pos/Reactive 누적)
                                If InStr(UCase(vTmp), "POS") > 0 Or _
                                        (InStr(UCase(vTmp), "REACTIVE") > 0 And Left(UCase(vTmp), 3) <> "NON") Then
                                    sgTmp = sgTmp + 1
                                End If
                            End If
                        End If
                        '>E-----------------------------
                    End If
                Next iRow
                
                'DIsp spdBottom
                With spdBottom
                    If iCnt > 0 Then
                        If pTestItem(iCol).JGBN = "1" Then
                            sgTmp = sgTmp / iCnt
                            If pTestItem(iCol).DOT <> "" Then
                                If pTestItem(iCol).DOT = "0" Then
                                    sDispRst = Format(sgTmp, "0")
                                Else
                                    sDispRst = Format(sgTmp, "0." & String(Val(pTestItem(iCol).DOT), "0"))
                                End If
                            Else
                                sDispRst = Format(sgTmp, "0.00")
                            End If
                        ElseIf pTestItem(iCol).JGBN = "2" Then
                            sgTmp = sgTmp / iCnt * 100
                            sDispRst = Format(sgTmp, "0.0")
                        End If
                        Call .SetText(iCol, .MaxRows, Trim(sDispRst))
                    Else
                        Call .SetText(iCol, .MaxRows, "0.0")
                    End If
                End With
            End If
        Next iCol
    End With
    
ErrGetRatio:
    If Err <> 0 Then
        Call DispMsg("DispSpdBottom Err - " & Err.Description)
    End If
End Sub
Private Sub GetPosRatio_Total(ByVal iSRow As Integer, ByVal iERow As Integer)
    On Error GoTo ErrGetRatio
    
    Dim vTmp, vChk
    Dim iRow%, iCol%, iCnt%
    Dim sTmp$
    Dim sgTmp   As Single
    Dim sDispRst$, sDispRst2$
    Dim sDate$
    Dim sLow$, sHigh$
    Dim sgLow   As Single
    Dim sgHigh  As Single
    Dim aRef()  As String
    
    Dim sRange  As String
    sRange = Trim(Replace(cboRange.Text, "%", ""))
    
    sDate = Format(DTP1(0).Value, "YYYY년 MM월")
        
    With spdBottom
        Call .SetCellBorder(4, .MaxRows, .MaxCols, .MaxRows, 8, vbBlack, CellBorderStyleSolid)
        
        .MaxRows = .MaxRows + 1
        .RowHeight(.MaxRows) = 11
        
        .BlockMode = True
        .Col = 1: .Col2 = .MaxCols
        .Row = .MaxRows: .Row2 = .MaxRows
        .FontBold = True
        .BlockMode = False
        
        Call .SetText(4, .MaxRows, sDate)
        
        .MaxRows = .MaxRows + 1
        .RowHeight(.MaxRows) = 11
        Call .SetText(5, .MaxRows, "(-/+ " & sRange & "%)")
    End With
    
    With spdList
        ReDim aLow(.MaxCols) As Single
        ReDim aHigh(.MaxCols) As Single
        
        For iCol = iSpdFixCol + 1 To .MaxCols
            If Trim(pTestItem(iCol).TESTNM) <> "" Then
                sgTmp = 0
                iCnt = 0
                
                For iRow = iSRow To iERow
                    Call .GetText(2, iRow, vChk)
                    If vChk = vbChecked Then
                        '<S--- 선택된 자료만 양성율 계산
                        Call .GetText(iCol, iRow, vTmp)
                        If Trim(vTmp) <> "" Then
                            iCnt = iCnt + 1
                            
                            If pTestItem(iCol).JGBN = "1" Then
                                '정량
                                sgTmp = sgTmp + Val(vTmp)
                            ElseIf pTestItem(iCol).JGBN = "2" Then
                                '정성(Pos/Reactive 누적)
                                If InStr(UCase(vTmp), "POS") > 0 Or _
                                        (InStr(UCase(vTmp), "REACTIVE") > 0 And Left(UCase(vTmp), 3) <> "NON") Then
                                    sgTmp = sgTmp + 1
                                End If
                            End If
                        End If
                        '>E-----------------------------
                    End If
                Next iRow
                
                'DIsp spdBottom
                With spdBottom
                    .ReDraw = False
                    
                    If iCnt > 0 Then
                        If pTestItem(iCol).JGBN = "1" Then
                            sgTmp = sgTmp / iCnt
                            If pTestItem(iCol).DOT <> "" Then
                                If pTestItem(iCol).DOT = "0" Then
                                    sDispRst = Format(sgTmp, "0")
                                Else
                                    sDispRst = Format(sgTmp, "0." & String(Val(pTestItem(iCol).DOT), "0"))
                                End If
                            Else
                                sDispRst = Format(sgTmp, "0.00")
                            End If
                        ElseIf pTestItem(iCol).JGBN = "2" Then
                            sgTmp = sgTmp / iCnt * 100
                            sDispRst = Format(sgTmp, "0.0")
                        End If
                        Call .SetText(iCol, .MaxRows - 1, Trim(sDispRst))
                        
                        '<S--- +/- ?% 계산&표시
                        sgHigh = sgTmp + (sgTmp * (Val(sRange) / 100))
                        sHigh = Format(sgHigh, "0.00")
                        
                        sgLow = sgTmp - (sgTmp * (Val(sRange) / 100))
                        sLow = Format(sgLow, "0.00")
                        
                        Call .SetText(iCol, .MaxRows, Trim(sLow) & "/" & Trim(sHigh))
                        '>E--------------------
                        
                        '<S--- +/- ?% 벗어난 값 색상 표시
                        For iRow = 1 To .MaxRows
                            Call .GetText(iCol, iRow, vTmp)
                            If Val(vTmp) <> 0 Then
                                If Val(vTmp) > Val(sHigh) Then
                                    .Col = iCol: .Row = iRow
                                    .BackColor = RGB(255, 180, 180)
                                ElseIf Val(vTmp) < Val(sLow) Then
                                    .Col = iCol: .Row = iRow
                                    .BackColor = RGB(180, 180, 255)
                                End If
                            End If
                        Next iRow
                        '>E------------------------------
                    Else
                        Call .SetText(iCol, .MaxRows - 1, "0.0")
                        
                        '<S--- +/- ?% 표시
                        Call .SetText(iCol, .MaxRows, "0.00/0.00")
                        '>E----------------
                    End If
                    
                    .ReDraw = True
                End With
            End If
        Next iCol
    End With
    
ErrGetRatio:
    If Err <> 0 Then
        Call DispMsg("DispSpdBottom Err - " & Err.Description)
    End If
End Sub


Private Function fExcelExePath() As String
    Dim sTmp    As String
    
    sTmp = GetKeyValue(HKEY_CURRENT_USER, _
                "Software\Ack_if\Interface Config\" & gsMachineCd, "Excel.Exe.Path")
                
    fExcelExePath = sTmp
    
End Function

Private Sub cboGbn_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{tab}"
    End If
End Sub


Private Sub cboRange_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{tab}"
    End If
End Sub


Private Sub cmdApply_Click()
    
    '% 표시
    Call DispSpdBottom
    
End Sub

Private Sub cmdExcel_Click()
    On Error GoTo ErrRtn
    
    Dim ii%, kk%
    Dim vTmp
    Dim aData() As String
    Dim aColor()    As String
    Dim sDate   As String
    Dim sFileNm As String
    Dim sLogNm  As String
    Dim sPath   As String
    
    Dim sRange  As String
    sRange = Trim(Replace(cboRange.Text, "%", ""))
    
    If spdBottom.MaxRows = 0 Then
        MsgBox "해당자료가 존재하지 않습니다.", vbInformation
        Exit Sub
    End If
    
    MousePointer = vbHourglass
    
    spdExcel.MaxRows = 0
    
    With spdBottom
        For ii = 1 To .MaxRows
            Call .GetText(4, ii, vTmp): sDate = Trim(vTmp)
            
            ReDim aData(.MaxCols) As String
            ReDim aColor(.MaxCols) As String
            
            For kk = iSpdFixCol + 1 To .MaxCols
                Call .GetText(kk, ii, vTmp): aData(kk) = Trim(vTmp)
                
                .Col = kk: .Row = ii
                Select Case .BackColor
                    Case RGB(255, 180, 180): aColor(kk) = "H"
                    Case RGB(180, 180, 255): aColor(kk) = "L"
                    Case Else: aColor(kk) = ""
                End Select
            Next kk
            
            With spdExcel
                .ReDraw = False
                
                .MaxRows = .MaxRows + 1
                .RowHeight(.MaxRows) = 12
                
                If ii = spdBottom.MaxRows Then
                    Call .SetText(1, .MaxRows, "(-/+ " & sRange & "%)")
                ElseIf ii = spdBottom.MaxRows - 1 Then
                    .MaxRows = .MaxRows + 1
                    .RowHeight(.MaxRows) = 12
                    Call .SetText(1, .MaxRows, sDate)
                    
                    .MaxRows = .MaxRows + 1
                    .RowHeight(.MaxRows) = 12
                    Call .SetText(1, .MaxRows, "월평균")
                Else
                    Call .SetText(1, .MaxRows, sDate)
                End If
                
                For kk = 2 To .MaxCols
                    Call .SetText(kk, .MaxRows, aData(kk + 4))
                    
                    .Col = kk: .Row = .MaxRows
                    Select Case aColor(kk + 4)
                        Case "H": .ForeColor = vbRed
                        Case "L": .ForeColor = vbBlue
                        Case Else: .ForeColor = vbBlack
                    End Select
'                    Select Case aColor(kk + 4)
'                        Case "H": .BackColor = RGB(255, 150, 150)
'                        Case "L": .BackColor = RGB(150, 150, 255)
'                        Case Else: .BackColor = vbWhite
'                    End Select
                Next kk
                
                .ReDraw = True
            End With
        Next ii
    End With
    
    sPath = fExcelFilePath
    If Right(sPath, 1) <> "\" Then
        sPath = sPath & "\"
    End If
    
    sFileNm = sPath & "FST02_" & Format(Now, "yyyymmdd") & ".xls"
    sLogNm = sPath & "FST02.log"
    
    Call spdExcel.ExportToHTML(sFileNm, False, sLogNm)
'    Call spdExcel.ExportToExcel(sFileNm, "", sLogNm)
    
    If Trim(fExcelExePath) <> "" Then
        Call WinExec(fExcelExePath & " " & sFileNm, 3)
    Else
        Me.MousePointer = vbDefault
        MsgBox "해당 PC에 EXCEL 프로그램이 존재하지 않습니다!!", vbInformation

        Exit Sub
    End If
    
    MousePointer = vbDefault
    
ErrRtn:
    If Err <> 0 Then
        MousePointer = vbDefault
        MsgBox Err.Description, vbExclamation, Me.Caption
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSearch_Click()
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
    
    spdList.ReDraw = False
    spdList.MaxRows = 0
    spdBottom.MaxRows = 0
    
    If Chk_Condition <> True Then
        Exit Sub
    End If
    
    MousePointer = vbHourglass
    
    sSql = " SELECT WDATE, WSEQ, JNO, IFSEQ, RESULT1, name "
    sSql = sSql & " FROM IFRESULT "
    sSql = sSql & " Where WDATE BETWEEN '" & Format(DTP1(0).Value, "YYYYMMDD") & "'"
    sSql = sSql & "                 AND '" & Format(DTP1(1).Value, "YYYYMMDD") & "'"
    sSql = sSql & "   AND LEN(jno) >= " & Val(gOrdCfg.sFSize(3)) & ""
    sSql = sSql & "   AND LEFT(jno,3) <> 'ERR' "
    sSql = sSql & "   and result1 <> '' "
    If cboGbn.ListIndex = 1 Then
        sSql = sSql & " and regstate = '1' "
    ElseIf cboGbn.ListIndex = 2 Then
        sSql = sSql & " and regstate = '0' "
    End If
    sSql = sSql & " ORDER BY WDATE, WSEQ, IFSEQ "
    
    ADORS.Open sSql, fGetCurDSN(gsMachineCd), adOpenForwardOnly
    
    If ADORS.EOF = True Then
        ADORS.Close: Set ADORS = Nothing
        MousePointer = vbDefault
        MsgBox "해당자료가 존재하지 않습니다.", vbInformation, Me.Caption
        Exit Sub
    End If
    
    ADORS.MoveFirst
    iSRow = 0
    Do While Not ADORS.EOF
        iCol = 0
        For ii = 1 To spdList.MaxCols - 1
            With pTestItem(ii)
                If Trim(ADORS.Fields(3)) = Trim(.IFSEQ) Then
                    iCol = ii   '.Col
                    Exit For
                End If
            End With
        Next ii
        
        If iCol > 0 Then
            With spdList
                If Trim(sDate) = "" Or Trim(ADORS.Fields(0) & "") <> sDate Then
                    .MaxRows = .MaxRows + 1
                    .RowHeight(.MaxRows) = 11
                    
                    Call .SetText(4, .MaxRows, Format(ADORS.Fields(0), "####-##-##"))
                    
                    .Col = 4: .Row = .MaxRows
                    .TypeHAlign = TypeHAlignCenter
                    
                    .Col = 4: .Col2 = .MaxCols
                    .Row = .MaxRows: .Row2 = .MaxRows
                    .BlockMode = True
                    .BackColor = RGB(235, 255, 235)
                    .BlockMode = False
                    
                    iSRow = .MaxRows
                    
                    '<S--- CellType 지정
                    .Col = 1: .Row = .MaxRows
                    .CellType = CellTypeCheckBox
                    .TypeVAlign = TypeVAlignCenter
                    .TypeHAlign = TypeHAlignCenter
                    .TypeCheckPicture(0) = imgPlus: .TypeCheckPicture(2) = imgPlus
                    .TypeCheckPicture(1) = imgMinus: .TypeCheckPicture(3) = imgMinus
                    
                    .Col = 2: .Row = .MaxRows
                    .CellType = CellTypeStaticText
                    '>E-----------------
                End If
                sDate = Trim(ADORS.Fields(0) & "")
            
                iRow = 0
                For ii = iSRow + 1 To .MaxRows
                    Call .GetText(3, ii, vTmp)
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
                    .RowHeight(.MaxRows) = 11
                    
                    iRow = .MaxRows
                    
                    .Row = .MaxRows: .RowHidden = True
                    Call .SetText(2, .MaxRows, vbChecked)
                End If
                
                Call .SetText(3, iRow, Trim(ADORS.Fields(1) & ""))
                Call .SetText(4, iRow, Space(1) & Trim(ADORS.Fields(2) & ""))
                Call .SetText(5, iRow, Trim(ADORS.Fields(5) & ""))
                Call .SetText(iCol, iRow, Trim(ADORS.Fields(4) & ""))
            End With
        End If
        
        ADORS.MoveNext
    Loop
    
    ADORS.Close: Set ADORS = Nothing
    
    spdList.ReDraw = True
    
    '% 표시
    Call DispSpdBottom
    
    MousePointer = vbDefault
    
ErrHandler:
    If Err <> 0 Then
        MousePointer = vbDefault
        If ADORS.State = 1 Then
            ADORS.Close: Set ADORS = Nothing
        End If
        MsgBox Err.Description, vbExclamation, Me.Caption
    End If
End Sub


Private Sub DTP1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{tab}"
    End If
End Sub


Private Sub Form_Load()
    On Error GoTo ErrHandler
    
    Me.Left = 0: Me.Top = 0
    Me.Height = 11050
    
    Call GetOrdRstCfg
    
    DTP1(0).Value = Format(Now, "YYYY-MM") & "-01"
    DTP1(1).Value = Now
    
    cboGbn.ListIndex = 0
    cboRange.ListIndex = 1
    
    Call Get_TestItem1
    
    spdList.MaxRows = 0
    
    spdBottom.MaxRows = 0
    spdBottom.EditModePermanent = True
    
ErrHandler:
    If Err <> 0 Then
        MsgBox "LOAD ERR - (" & Err.Description & ")", vbExclamation
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call RegEditCurFrmTitle("Statistics", "")
    Call DispMsg("")
End Sub

Private Sub DispMsg(ByVal sMsg As String)
    StBar1.Panels(1).Text = Trim(sMsg)
End Sub

Private Sub spdBottom_ColWidthChange(ByVal Col1 As Long, ByVal Col2 As Long)

    Dim ii%
    
    With spdList
        .ReDraw = False
        
        For ii = Col1 To Col2
            .ColWidth(ii) = spdBottom.ColWidth(ii)
        Next ii
        
        .ReDraw = True
    End With
    
End Sub

Private Sub spdBottom_TopLeftChange(ByVal OldLeft As Long, ByVal OldTop As Long, ByVal NewLeft As Long, ByVal NewTop As Long)
    
    spdList.LeftCol = NewLeft
    
End Sub

Private Sub spdList_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    
    Dim vChk
    Dim ii%
    
    If Row = 0 Or Col <> 1 Then Exit Sub
    
    With spdList
        .Col = 1: .Row = Row
        If .CellType <> CellTypeCheckBox Then Exit Sub
        
        Call .GetText(1, Row, vChk)
        
        .ReDraw = False
        
        For ii = Row + 1 To .MaxRows
            .Row = ii
            If .CellType = CellTypeCheckBox Then
                Exit For
            Else
                If vChk = vbUnchecked Then
                    .RowHidden = True
                Else
                    .RowHidden = False
                End If
            End If
        Next ii
        
        .ReDraw = True
    End With
    
End Sub

Private Sub spdList_ColWidthChange(ByVal Col1 As Long, ByVal Col2 As Long)
    
    Dim ii%
    
    With spdBottom
        .ReDraw = False
        
        For ii = Col1 To Col2
            .ColWidth(ii) = spdList.ColWidth(ii)
        Next ii
        
        .ReDraw = True
    End With
    
End Sub
Private Sub DispSpdBottom()
    On Error GoTo ErrDisp
    
    Dim ii%, iSRow%, iERow%
    Dim vDate
    Dim sDate$
    
    If spdList.MaxRows = 0 Then Exit Sub
    
    spdBottom.ReDraw = False
    spdBottom.MaxRows = 0
    
    With spdList
        MousePointer = vbHourglass
    
        sDate = ""
        iSRow = 0: iERow = 0
        
        For ii = 1 To .MaxRows
            .Col = 1: .Row = ii
            If .CellType = CellTypeCheckBox Then
                If iSRow = 0 Then
                    iSRow = ii
                    
                    Call .GetText(4, ii, vDate): sDate = Trim(vDate)
                Else
                    iERow = ii
                    
                    '해당 구간 양성율 표시
                    Call GetPosRatio(sDate, iSRow, iERow - 1)
                    
                    iSRow = ii: iERow = 0
                    Call .GetText(4, ii, vDate): sDate = Trim(vDate)
                End If
            Else
                iERow = ii
            End If
        Next ii
        
        If iSRow > 0 And iERow > 0 Then
            '해당 구간 양성율 표시
            Call GetPosRatio(sDate, iSRow, iERow)
        End If
    End With

    '월평균 표시 & 색상표시
    Call GetPosRatio_Total(1, spdList.MaxRows)
    

    spdBottom.ReDraw = True
    
    MousePointer = vbDefault
    
ErrDisp:
    If Err <> 0 Then
        If spdBottom.ReDraw = False Then
            spdBottom.ReDraw = True
        End If
        MousePointer = vbDefault
        MsgBox "DispSpdBottom Err - " & Err.Description, vbExclamation, Me.Caption
    End If
End Sub

