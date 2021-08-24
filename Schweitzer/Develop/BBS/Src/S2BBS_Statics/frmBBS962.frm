VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmBBS962 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   0  '없음
   Caption         =   "월중 혈액수불현황"
   ClientHeight    =   9105
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10995
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9105
   ScaleWidth      =   10995
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00F4F0F2&
      Caption         =   "출력(&P)"
      Height          =   510
      Left            =   8175
      Style           =   1  '그래픽
      TabIndex        =   15
      Tag             =   "124"
      Top             =   8400
      Width           =   1320
   End
   Begin VB.CommandButton cmdclear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      Height          =   510
      Left            =   6855
      Style           =   1  '그래픽
      TabIndex        =   14
      Tag             =   "124"
      Top             =   8400
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   9510
      Style           =   1  '그래픽
      TabIndex        =   13
      Tag             =   "128"
      Top             =   8400
      Width           =   1320
   End
   Begin VB.CommandButton cmdQuery 
      BackColor       =   &H00F4F0F2&
      Caption         =   "조회(&Q)"
      Height          =   510
      Left            =   4215
      Style           =   1  '그래픽
      TabIndex        =   10
      Tag             =   "124"
      Top             =   8400
      Width           =   1320
   End
   Begin VB.CommandButton cmdExcel 
      BackColor       =   &H00EAE7E3&
      Caption         =   "Excel(&E)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   5535
      Style           =   1  '그래픽
      TabIndex        =   9
      Tag             =   "127"
      Top             =   8400
      Width           =   1320
   End
   Begin MSComDlg.CommonDialog DlgSave 
      Left            =   -105
      Top             =   8520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin FPSpread.vaSpread tblexcel 
      Height          =   675
      Left            =   345
      TabIndex        =   6
      Top             =   8535
      Visible         =   0   'False
      Width           =   675
      _Version        =   196608
      _ExtentX        =   1191
      _ExtentY        =   1191
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
      SpreadDesigner  =   "frmBBS962.frx":0000
   End
   Begin MedControls1.LisLabel LisLabel11 
      Height          =   315
      Left            =   75
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   45
      Width           =   6630
      _ExtentX        =   11695
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "월중 혈액수불현황"
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   7980
      Left            =   75
      TabIndex        =   3
      Top             =   285
      Width           =   6645
      Begin MSComCtl2.DTPicker dtpClose 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "gg yyyy-MM-dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   0
         EndProperty
         Height          =   315
         Left            =   1125
         TabIndex        =   4
         Top             =   195
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM"
         Format          =   68681731
         CurrentDate     =   36799
      End
      Begin FPSpread.vaSpread tblData 
         Height          =   7275
         Left            =   45
         TabIndex        =   5
         Top             =   540
         Width           =   6585
         _Version        =   196608
         _ExtentX        =   11615
         _ExtentY        =   12832
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   9
         MaxRows         =   30
         OperationMode   =   1
         ShadowColor     =   14737632
         ShadowDark      =   13818331
         SpreadDesigner  =   "frmBBS962.frx":01AB
         TextTip         =   4
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   3
         Left            =   45
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   195
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "조회기간"
         Appearance      =   0
      End
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   315
      Left            =   6735
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   45
      Width           =   3990
      _ExtentX        =   7038
      _ExtentY        =   556
      BackColor       =   8388608
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "수혈자수 합계(Irradiation)"
      Appearance      =   0
   End
   Begin VB.Frame FraCmc02 
      BackColor       =   &H00DBE6E6&
      Height          =   7980
      Left            =   6750
      TabIndex        =   0
      Top             =   285
      Width           =   4080
      Begin FPSpread.vaSpread tblTrans 
         Height          =   7290
         Left            =   30
         TabIndex        =   1
         Top             =   540
         Width           =   3915
         _Version        =   196608
         _ExtentX        =   6906
         _ExtentY        =   12859
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   5
         MaxRows         =   30
         OperationMode   =   1
         ShadowColor     =   14737632
         ShadowDark      =   13818331
         SpreadDesigner  =   "frmBBS962.frx":1EBE
         TextTip         =   4
      End
      Begin MedControls1.LisLabel lblIrrCount 
         Height          =   315
         Left            =   1575
         TabIndex        =   2
         Top             =   195
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         BackColor       =   14411494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   0
         Left            =   45
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   195
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "Irradiation건수"
         Appearance      =   0
      End
   End
End
Attribute VB_Name = "frmBBS962"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private objSql As New clsStatics
Private lngCompCnt As Long
Private Const MENU_LEFT& = 1
Private Const MENU_LEXCEL& = 2
Private Const MENU_RIGHT& = 3
Private Const MENU_REXCEL& = 4
Private Const MENU_SEP& = 5

Private Sub Forminitialize()
    Dim RS As Recordset
    Dim ii As Integer
    Dim jj As Integer
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    dtpClose.Value = Format(GetSystemDate, "yyyy-mm")
    Set RS = objSql.Statics962
    
    Call medClearTable(tblData)
    Call medClearTable(tblTrans)

    If Not RS.EOF Then
        lngCompCnt = RS.RecordCount
        
        With tblData
            .MaxRows = lngCompCnt * 5 + 1 '맨 마지막 로우에 총계 추가
            .RowHeight(-1) = 12.6
            
            'border line 없애기
            .Row = 1: .Row2 = .MaxRows
            .Col = 1: .Col2 = .MaxCols
            .BlockMode = True
            .CellBorderColor = vbWhite
            .CellBorderStyle = CellBorderStyleSolid
            .CellBorderType = 1 Or 2 Or 4 Or 8
            .Action = 16
            .BlockMode = False
        End With
        
        With tblTrans
            .MaxRows = lngCompCnt * 5 + 1 '맨 마지막 로우에 총계 추가
            .RowHeight(-1) = 12.6
            
            'border line 없애기
            .Row = 1: .Row2 = .MaxRows
            .Col = 1: .Col2 = .MaxCols
            .BlockMode = True
            .CellBorderColor = vbWhite
            .CellBorderStyle = CellBorderStyleSolid
            .CellBorderType = 1 Or 2 Or 4 Or 8
            .Action = 16
            .BlockMode = False
        End With

        For i = 0 To 4 * lngCompCnt Step lngCompCnt
            j = 0
            RS.MoveFirst
            Do Until RS.EOF
                With tblData

                    .Row = 1 + i + j: .Col = 2: .Value = RS.Fields("field1").Value & ""
                                      .Col = 9: .Value = RS.Fields("cdval1").Value & ""
                    tblTrans.Row = 1 + i + j: tblTrans.Col = 2: tblTrans.Value = RS.Fields("field1").Value & ""

                    j = j + 1
                End With
                RS.MoveNext
            Loop

'            혈액형 블럭별 색깔 변경
            With tblData
                .Row = 1 + i: .Row2 = 1 + i + lngCompCnt
                .Col = 1: .Col2 = .MaxCols
                .BlockMode = True
                .BackColor = Choose(i / lngCompCnt + 1, RGB(255, 255, 202), RGB(238, 255, 255), RGB(255, 208, 255), RGB(255, 255, 202), RGB(238, 255, 255))
                .BlockMode = False

                tblTrans.Row = 1 + i: tblTrans.Row2 = 1 + i + lngCompCnt
                tblTrans.Col = 1: tblTrans.Col2 = tblTrans.MaxCols
                tblTrans.BlockMode = True
                tblTrans.BackColor = Choose(i / lngCompCnt + 1, RGB(255, 255, 202), RGB(238, 255, 255), RGB(255, 208, 255), RGB(255, 255, 202), RGB(238, 255, 255))
                tblTrans.BlockMode = False
                
'총계부분 처리
'------------
                .Row = .MaxRows: .Row2 = .MaxRows
                .Col = 1: .Col2 = .MaxCols
                .BlockMode = True
                .BackColor = DCM_LightRed
                .BlockMode = False
                
                tblTrans.Row = tblTrans.MaxRows: tblTrans.Row2 = tblTrans.MaxRows
                tblTrans.Col = 1: tblTrans.Col2 = tblTrans.MaxCols
                tblTrans.BlockMode = True
                tblTrans.BackColor = DCM_LightRed
                tblTrans.BlockMode = False
'------------
'            혈액형 블럭 border line 그리기
                .Row = 1 + i: .Row2 = 1 + i + lngCompCnt
                .Col = 1: .Col2 = 1
                .BlockMode = True
                .CellBorderStyle = CellBorderStyleSolid
                .CellBorderType = 16 '4 Or 8
                .CellBorderColor = vbBlack
                .Action = 16
                .CellType = CellTypeStaticText
                .BlockMode = False

                tblTrans.Row = 1 + i: tblTrans.Row2 = 1 + i + lngCompCnt
                tblTrans.Col = 1: tblTrans.Col2 = 1
                tblTrans.BlockMode = True
                tblTrans.CellBorderStyle = CellBorderStyleSolid
                tblTrans.CellBorderType = 16 '4 Or 8
                tblTrans.CellBorderColor = vbBlack
                tblTrans.Action = 16
                tblTrans.CellType = CellTypeStaticText
                tblTrans.BlockMode = False
                
'총계부분 처리
'------------
                .Row = .MaxRows: .Row2 = .MaxRows
                .Col = 1: .Col2 = 1
                .BlockMode = True
                .CellBorderStyle = CellBorderStyleSolid
                .CellBorderType = 16 '4 Or 8
                .CellBorderColor = vbBlack
                .Action = 16
                .CellType = CellTypeStaticText
                .BlockMode = False

                tblTrans.Row = tblTrans.MaxRows: tblTrans.Row2 = tblTrans.MaxRows
                tblTrans.Col = 1: tblTrans.Col2 = 1
                tblTrans.BlockMode = True
                tblTrans.CellBorderStyle = CellBorderStyleSolid
                tblTrans.CellBorderType = 16 '4 Or 8
                tblTrans.CellBorderColor = vbBlack
                tblTrans.Action = 16
                tblTrans.CellType = CellTypeStaticText
                tblTrans.BlockMode = False
'------------

'            혈액형 등록하기
                Call .SetText(1, CLng(lngCompCnt / 2) + i, Choose(i / lngCompCnt + 1, "A", "B", "O", "AB", "합계"))
                Call .SetText(1, .MaxRows, "총계")
                
                Call tblTrans.SetText(1, CLng(lngCompCnt / 2) + i, Choose(i / lngCompCnt + 1, "A", "B", "O", "AB", "합계"))
                Call tblTrans.SetText(1, tblTrans.MaxRows, "총계")
            End With
        Next i

        With tblData
            'border line 그리기
            .Row = 1: .Row2 = .MaxRows
            .Col = 2: .Col2 = .MaxCols
            .BlockMode = True
            .CellBorderStyle = CellBorderStyleSolid
            .CellBorderType = 1 Or 2 Or 4 Or 8
            .CellBorderColor = vbBlack
            .Action = 16
            .BlockMode = False
            
            '혈액형 정렬
             .Row = 1: .Row2 = .MaxRows
             .Col = 1: .Col2 = 1
             .BlockMode = True
             .TypeHAlign = TypeHAlignCenter
             .TypeVAlign = TypeVAlignCenter
             .BlockMode = False
             '제제명 정렬(왼쪽)
             .Row = 1: .Row2 = .MaxRows
             .Col = 2: .Col2 = 2
             .BlockMode = True
             .TypeHAlign = TypeHAlignLeft
             .TypeVAlign = TypeVAlignCenter
             .BlockMode = False
             '수치값 정렬(오른쪽)
             .Row = 1: .Row2 = .MaxRows
             .Col = 3: .Col2 = .MaxCols
             .BlockMode = True
             .TypeHAlign = TypeHAlignRight
             .TypeVAlign = TypeVAlignCenter
             .BlockMode = False
        End With

        With tblTrans
            'border line 그리기
            .Row = 1: .Row2 = .MaxRows
            .Col = 2: .Col2 = .MaxCols
            .BlockMode = True
            .CellBorderStyle = CellBorderStyleSolid
            .CellBorderType = 1 Or 2 Or 4 Or 8
            .CellBorderColor = vbBlack
            .Action = 16
            .BlockMode = False
            
            '혈액형 정렬
             .Row = 1: .Row2 = .MaxRows
             .Col = 1: .Col2 = 1
             .BlockMode = True
             .TypeHAlign = TypeHAlignCenter
             .TypeVAlign = TypeVAlignCenter
             .BlockMode = False
             '제제명 정렬(왼쪽)
             .Row = 1: .Row2 = .MaxRows
             .Col = 2: .Col2 = 2
             .BlockMode = True
             .TypeHAlign = TypeHAlignLeft
             .TypeVAlign = TypeVAlignCenter
             .BlockMode = False
             '수치값 정렬(오른쪽)
             .Row = 1: .Row2 = .MaxRows
             .Col = 3: .Col2 = .MaxCols
             .BlockMode = True
             .TypeHAlign = TypeHAlignRight
             .TypeVAlign = TypeVAlignCenter
             .BlockMode = False
        End With
    End If
    
'    If Not RS.EOF Then
'        lngCompCnt = RS.RecordCount
'
'        With tblData
'            For ii = 1 To 30 Step 6
'                RS.MoveFirst
'                jj = ii
'                Do Until RS.EOF
'                    .Row = jj
'                    .Col = 2:  .Value = RS.Fields("field1").Value & ""
'                    .Col = 9: .Value = RS.Fields("cdval1").Value & ""
'                    jj = jj + 1
'                    RS.MoveNext
'                Loop
'            Next
'            .Col = 3: .Col2 = 9
'            .Row = 1: .Row2 = 30
'            .BlockMode = True
'            .Text = ""
'            .BlockMode = False
'        End With
'
'        With tblTrans
'            For ii = 1 To 30 Step 6
'                RS.MoveFirst
'                jj = ii
'                Do Until RS.EOF
'                    .Row = jj
'                    .Col = 2:  .Value = RS.Fields("field1").Value & ""
'                    jj = jj + 1
'                    RS.MoveNext
'                Loop
'            Next
'            .Col = 3: .Col2 = 5
'            .Row = 1: .Row2 = 30
'            .BlockMode = True
'            .Text = ""
'            .BlockMode = False
'        End With
'    End If
    
    lblIrrCount.Caption = ""
    
    Set RS = Nothing

End Sub

Private Sub cmdClear_Click()
    With tblData
        .Col = 3: .Col2 = 9
        .Row = 1: .Row2 = .MaxRows
        .BlockMode = True
        .Text = ""
        .BlockMode = False
    End With
    With tblTrans
        .Col = 3: .Col2 = 5
        .Row = 1: .Row2 = .MaxRows
        .BlockMode = True
        .Text = ""
        .BlockMode = False
    End With
    lblIrrCount.Caption = ""
End Sub

Private Sub ExportLeft()
    Dim strTmp As String
    Dim lngRows As Long
    
    If tblData.DataRowCnt = 0 Then Exit Sub
    
    With tblData
        .Row = 0: .Row2 = .MaxRows
        .Col = 1: .Col2 = .MaxCols
        .BlockMode = True
        strTmp = .Clip
        .BlockMode = False
        lngRows = .MaxRows
    End With
 
    With tblexcel
        .MaxRows = tblData.MaxRows + 1
        .MaxCols = tblData.MaxCols
        .Row = 1: .Row2 = .MaxRows
        .Col = 1: .Col2 = tblData.MaxCols
        .BlockMode = True
        .Clip = strTmp
        .BlockMode = False
    End With
    
    DlgSave.InitDir = "C:\"
    DlgSave.Filter = "ExCelFile(*.XLS)|*.XLS"
    DlgSave.FileName = "월중 수불내역 통계"
    DlgSave.ShowSave

    tblexcel.SaveTabFile (DlgSave.FileName)
End Sub

Private Sub ExportRight()
    Dim strTmp As String
    Dim lngRows As Long
    
    If tblTrans.DataRowCnt = 0 Then Exit Sub

    Call medClearTable(tblexcel)
    
    With tblTrans
        .Row = 0: .Row2 = .MaxRows
        .Col = 1: .Col2 = .MaxCols
        .BlockMode = True
        strTmp = .Clip
        .BlockMode = False
        lngRows = .MaxRows
    End With
 
    With tblexcel
        .MaxRows = tblTrans.MaxRows + 1
        .MaxCols = tblTrans.MaxCols
        .Row = 1: .Row2 = .MaxRows
        .Col = 1: .Col2 = tblTrans.MaxCols
        .BlockMode = True
        .Clip = strTmp
        .BlockMode = False
    End With
    
    DlgSave.InitDir = "C:\"
    DlgSave.Filter = "ExCelFile(*.XLS)|*.XLS"
    DlgSave.FileName = "월중 수혈자수"
    DlgSave.ShowSave

    tblexcel.SaveTabFile (DlgSave.FileName)
End Sub

Private Sub cmdExcel_Click()
    MsgBox "Excel로 Export할 내역을 마우스 우클릭으로 선택하십시오.", vbInformation
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdQuery_Click()
    Dim objProBar   As New clsProgress
    Dim ii          As Long
    
'    Set objProBar.StatusBar = MainFrm.stsbar
    objProBar.Container = MainFrm.stsbar
    
    Call cmdClear_Click
    
    Me.MousePointer = 11
    objProBar.Max = 100
    For ii = 1 To 20
        objProBar.Message = "혈액입고 건수를 집계중입니다..."
        objProBar.Value = ii
    Next
    Call Storage                '입고(혈액원/헌혈)
    For ii = 21 To 40
        objProBar.Message = "혈액출고 건수를 집계중입니다..."
        objProBar.Value = ii
    Next
    Call DeliveryBlood          '출고/폐기/반환
    For ii = 41 To 60
        objProBar.Message = "혈액반환/폐기 건수를 집계중입니다..."
        objProBar.Value = ii
    Next
    Call Inventory              '현재고
    For ii = 61 To 80
        objProBar.Message = "혈액재고를 집계중입니다..."
        objProBar.Value = ii
    Next
        
    Call TotalTransfusion   '수혈 맞은갯수
    Call IrrCount           'irradiation 건수
    
    Call TotalInventory         '합계
    For ii = 81 To 100
        objProBar.Message = "합계 계산중입니다..."
        objProBar.Value = ii
    Next
    Me.MousePointer = 0
    Set objProBar = Nothing
End Sub

Private Sub Form_Load()
    Call Forminitialize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objSql = Nothing
End Sub

Private Sub cmdPrint_Click()
    Call PrintLeftInfo
    
    Call PrintRightInfo
End Sub

Private Sub PrintLeftInfo()
    Dim strFont   As String
    Dim strYear   As String
    Dim strMonth  As String
    
    strYear = Format(dtpClose.Value, "yyyy")
    strMonth = Format(Format(dtpClose.Value, "mm"), "##")
    With tblData
        .PrintJobName = "월중 혈액수불현황"
        .PrintAbortMsg = "월중 혈액수불현황 출력중입니다. "
        .PrintColor = False
        .PrintFirstPageNumber = 1
         strFont = "/fn""굴림체""/fz""11"""
        .PrintHeader = strFont & "/n/n/n/l/fb1 " & " 【 월중 혈액수불현황 통계 보고서 】 /n/l/fb1   조회월:" & _
                       strYear & " 년 " & _
                       strMonth & " 월" & "/c/fb1/n/n/n/n/n"
        .PrintFooter = " /l " & String(125, Chr(6)) & "/n/l " & HOSPITAL_NAME & "/c/p/fb1" & " /r" & strMonth & " 월중 혈액수불현황 "
        
        .PrintMarginBottom = 100
        .PrintMarginLeft = 1000
        .PrintMarginRight = 100
        .PrintMarginTop = 300
        
        .PrintShadows = True
        .PrintNextPageBreakCol = 1
        .PrintNextPageBreakRow = 1
        .PrintRowHeaders = True
        .PrintColHeaders = True
        .PrintBorder = True
        .PrintGrid = True
'        .GridSolid = True
        .PrintType = PrintTypeAll
        .Action = ActionPrint
'        .GridSolid = False
    End With
End Sub

Private Sub PrintRightInfo()
    Dim strFont   As String
    Dim strYear   As String
    Dim strMonth  As String
    
    strYear = Format(dtpClose.Value, "yyyy")
    strMonth = Format(Format(dtpClose.Value, "mm"), "##")

    With tblTrans
'        .Position = PositionCenter

        .PrintJobName = "월중 혈액수혈자수 현황"
        .PrintAbortMsg = "월중 혈액수혈자수 현황 출력중입니다. "
        .PrintColor = False
        .PrintFirstPageNumber = 1
         strFont = "/fn""굴림체""/fz""11"""
        .PrintHeader = strFont & "/n/n/n/l/fb1 " & " 【 월중 혈액수혈자수 현황 통계 보고서 】 /n/l/fb1   Irradiation건수: " & lblIrrCount.Caption & " /n/l/fb1 " & _
                       "  조회월:" & _
                       strYear & " 년 " & _
                       strMonth & " 월" & "/c/fb1/n/n/n/n/n"
                       
        .PrintFooter = " /l " & String(125, Chr(6)) & "/n/l " & HOSPITAL_NAME & "/c/p/fb1" & " /r" & strMonth & " 월중 혈액수혈자수 현황 "
        
        .PrintMarginBottom = 100
        .PrintMarginLeft = 1000
        .PrintMarginRight = 100
        .PrintMarginTop = 300
        
        .PrintShadows = True
        .PrintNextPageBreakCol = 1
        .PrintNextPageBreakRow = 1
        .PrintRowHeaders = True
        .PrintColHeaders = True
        .PrintBorder = True
        .PrintGrid = True
'        .GridSolid = True
        .PrintType = PrintTypeAll
        .Action = ActionPrint
'        .GridSolid = False
    End With
End Sub

Private Sub Inventory()
    Dim RS      As Recordset
    Dim SSQL    As String
    Dim NewRow  As Long
    
    SSQL = " select  a.compocd,a.abo,a.groupcd, sum(a.cnt) as cnt from("
    
    SSQL = SSQL & _
           " SELECT  b.groupcd,a.compocd,a.abo,count(a.abo) as cnt " & _
           " FROM " & T_BBS006 & " b," & T_BBS401 & " a " & _
           " WHERE " & DBW("a.entdt<=", Format(dtpClose.Value, "yyyymm") & "31") & _
           " AND " & DBW("a.centercd=", ObjSysInfo.BuildingCd) & _
           " AND a.stscd != '4' " & _
           " AND a.compocd=b.compocd" & _
           " GROUP BY a.compocd,a.abo,b.groupcd"
    
    SSQL = SSQL & " union all "
    
    SSQL = SSQL & _
           " SELECT  b.groupcd,a.compocd,a.abo,count(a.abo) as cnt " & _
           " FROM " & T_BBS006 & " b," & T_BBS401 & " a " & _
           " WHERE " & DBW("a.entdt<=", Format(dtpClose.Value, "yyyymmdd") & "31") & _
           " AND " & DBW("a.centercd=", ObjSysInfo.BuildingCd) & _
           " AND a.hosfg = '1' " & _
           " AND a.stscd != '4' " & _
           " AND a.compocd=b.compocd" & _
           " GROUP BY a.compocd,a.abo,b.groupcd"
           
    SSQL = SSQL & " union all "
    
    SSQL = SSQL & _
           " (SELECT  b.groupcd,a.compocd,a.abo,count(a.abo) * -1 as cnt " & _
           " FROM " & T_BBS402 & " c, " & T_BBS006 & " b," & T_BBS401 & " a " & _
           " WHERE " & DBW("c.deliverydt<=", Format(dtpClose.Value, "yyyymmdd") & "31") & _
           " AND " & DBW("a.centercd=", ObjSysInfo.BuildingCd) & _
           " AND c.retfg is null " & _
           " AND a.stscd != '4' " & _
           " AND a.compocd=b.compocd" & _
           " AND a.bldsrc=c.bldsrc " & _
           " AND a.bldyy=c.bldyy " & _
           " AND a.bldno=c.bldno " & _
           " AND a.compocd=c.compocd " & _
           " GROUP BY a.compocd,a.abo,b.groupcd"
           
    SSQL = SSQL & " union "
    
    SSQL = SSQL & _
           " SELECT  b.groupcd,a.compocd,a.abo,count(a.abo) as cnt " & _
           " FROM " & T_BBS402 & " c, " & T_BBS006 & " b," & T_BBS401 & " a " & _
           " WHERE " & DBW("a.realexpdt<=", Format(dtpClose.Value, "yyyymmdd") & "31") & _
           " AND " & DBW("a.centercd=", ObjSysInfo.BuildingCd) & _
           " AND c.retfg is null " & _
           " AND a.stscd != '4' " & _
           " AND a.compocd=b.compocd" & _
           " AND a.bldsrc=c.bldsrc " & _
           " AND a.bldyy=c.bldyy " & _
           " AND a.bldno=c.bldno " & _
           " AND a.compocd=c.compocd " & _
           " GROUP BY a.compocd,a.abo,b.groupcd) ) a" & _
           " GROUP BY a.compocd,a.abo,a.groupcd " & _
           " having sum(a.cnt) !=0 "
           
    '-- 원본 ------------------------------------------------------------
'    SSQL = " SELECT  b.groupcd,a.compocd,a.abo,count(a.abo) as cnt " & _
           " FROM " & T_BBS006 & " b," & T_BBS401 & " a " & _
           " WHERE a.stscd in('0','1','2') " & _
           " AND " & DBW("a.centercd=", ObjSysInfo.BuildingCd) & _
           " AND a.compocd=b.compocd" & _
           " GROUP BY a.compocd,a.abo,b.groupcd"
    '--------------------------------------------------------------------
           
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    If Not RS.EOF Then
        With tblData
        Do Until RS.EOF
            Select Case RS.Fields("abo").Value & ""
'                Case "A":  NewRow = 1
'                Case "B":  NewRow = 7
'                Case "O":  NewRow = 13
'                Case "AB": NewRow = 19
                Case "A":  NewRow = 1
                Case "B":  NewRow = lngCompCnt * 1 + 1
                Case "O":  NewRow = lngCompCnt * 2 + 1
                Case "AB": NewRow = lngCompCnt * 3 + 1
            End Select
            
'            .Row = NewRow + Val(Mid(RS.Fields("groupcd").Value & "", 2)) - 1
            .Row = NewRow + Val(RS.Fields("groupcd").Value & "") - 1
            .Col = 8: .Value = RS.Fields("cnt").Value & ""
            RS.MoveNext
        Loop
        End With
    End If
    
    Set RS = Nothing
    
End Sub

Private Sub Storage()
    Dim RS          As Recordset
    Dim SMSQL       As String
    Dim SSQL        As String
    Dim SubSQL      As String
    Dim strABO      As String
    Dim NewRow      As Long
    
    
    SSQL = " SELECT  b.groupcd,a.compocd,a.abo,count(a.abo) as cnt ,'' as hosfg" & _
           " FROM " & T_BBS006 & " b," & T_BBS401 & " a " & _
           " WHERE " & _
                     DBW("a.entdt>=", Format(dtpClose.Value, "yyyymm") & "01") & _
           " AND " & DBW("a.entdt<=", Format(dtpClose.Value, "yyyymm") & "31") & _
           " AND " & DBW("a.centercd=", ObjSysInfo.BuildingCd) & _
           " AND a.compocd=b.compocd" & _
           " GROUP BY a.compocd,a.abo,b.groupcd"
    SSQL = SSQL & " UNION ALL"
    SSQL = SSQL & " SELECT  b.groupcd,a.compocd,a.abo,count(a.abo) as cnt ,'1' as hosfg" & _
           " FROM " & T_BBS006 & " b," & T_BBS401 & " a " & _
           " WHERE " & _
                     DBW("a.entdt>=", Format(dtpClose.Value, "yyyymm") & "01") & _
           " AND " & DBW("a.entdt<=", Format(dtpClose.Value, "yyyymm") & "31") & _
           " AND " & DBW("a.centercd=", ObjSysInfo.BuildingCd) & _
           " AND " & DBW("a.hosfg=", "1") & _
           " AND a.compocd=b.compocd" & _
           " GROUP BY a.compocd,a.abo,b.groupcd"
    
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    If Not RS.EOF Then
        With tblData
            Do Until RS.EOF
                Select Case RS.Fields("abo").Value & ""
'                    Case "A":  NewRow = 1
'                    Case "B":  NewRow = 7
'                    Case "O":  NewRow = 13
'                    Case "AB": NewRow = 19
                    Case "A":  NewRow = 1
                    Case "B":  NewRow = lngCompCnt * 1 + 1
                    Case "O":  NewRow = lngCompCnt * 2 + 1
                    Case "AB": NewRow = lngCompCnt * 3 + 1
                End Select
                
'                .Row = NewRow + Val(Mid(RS.Fields("groupcd").Value & "", 2)) - 1
                .Row = NewRow + Val(RS.Fields("groupcd").Value & "") - 1
                .Col = 3
                If RS.Fields("hosfg").Value & "" = "1" Then .Col = 4
                 
                .Value = RS.Fields("cnt").Value & ""
                RS.MoveNext
            Loop
        End With
    End If
    Set RS = Nothing
    
End Sub

Private Sub DeliveryBlood()
    Dim RS          As Recordset
    Dim SSQL        As String
    Dim NewRow      As Long
    
    '출고건수
    SSQL = " SELECT  b.groupcd,a.compocd,a.abo,count(a.abo) as cnt,'1' as div " & _
           " FROM " & T_BBS402 & " c," & T_BBS006 & " b," & T_BBS401 & " a " & _
           " WHERE " & _
                     DBW("c.deliverydt>=", Format(dtpClose.Value, "yyyymm") & "01") & _
           " AND " & DBW("c.deliverydt<=", Format(dtpClose.Value, "yyyymm") & "31") & _
           " AND " & DBW("a.centercd=", ObjSysInfo.BuildingCd) & _
           " AND a.compocd=b.compocd" & _
           " AND a.bldsrc=c.bldsrc AND a.bldyy=c.bldyy AND a.bldno=c.bldno AND a.compocd=c.compocd " & _
           " GROUP BY a.compocd,a.abo,b.groupcd"
    '폐기건수
    SSQL = SSQL & " UNION " & _
           " SELECT  b.groupcd,a.compocd,a.abo,count(a.abo) as cnt,'2' as div " & _
           " FROM " & T_BBS402 & " c," & T_BBS006 & " b," & T_BBS401 & " a " & _
           " WHERE " & _
                     DBW("a.realexpdt>=", Format(dtpClose.Value, "yyyymm") & "01") & _
           " AND " & DBW("a.realexpdt<=", Format(dtpClose.Value, "yyyymm") & "31") & _
           " AND " & DBW("a.centercd=", ObjSysInfo.BuildingCd) & _
           " AND a.compocd=b.compocd" & _
           " AND a.bldsrc=c.bldsrc AND a.bldyy=c.bldyy AND a.bldno=c.bldno AND a.compocd=c.compocd " & _
           " GROUP BY a.compocd,a.abo,b.groupcd"
    
'2005/05/31 modify by legends
'자체 폐기건수를 구하기 위해 sql 구문 추가
    '출고전 자체 폐기건수
    SSQL = SSQL & " UNION " & _
           " SELECT  b.groupcd,a.compocd,a.abo,count(a.abo) as cnt,'2' as div " & _
           " FROM " & T_BBS006 & " b," & T_BBS401 & " a " & _
           " WHERE " & _
                     DBW("a.realexpdt>=", Format(dtpClose.Value, "yyyymm") & "01") & _
           " AND " & DBW("a.realexpdt<=", Format(dtpClose.Value, "yyyymm") & "31") & _
           " AND " & DBW("a.centercd=", ObjSysInfo.BuildingCd) & _
           " AND a.compocd=b.compocd" & _
           " GROUP BY a.compocd,a.abo,b.groupcd"
    
    '반환건수
    SSQL = SSQL & " UNION " & _
           " SELECT  b.groupcd,a.compocd,a.abo,count(a.abo) as cnt,'3' as div " & _
           " FROM " & T_BBS402 & " c," & T_BBS006 & " b," & T_BBS401 & " a " & _
           " WHERE " & _
                     DBW("c.retdt>=", Format(dtpClose.Value, "yyyymm") & "01") & _
           " AND " & DBW("c.retdt<=", Format(dtpClose.Value, "yyyymm") & "31") & _
           " AND " & DBW("a.centercd=", ObjSysInfo.BuildingCd) & _
           " AND a.compocd=b.compocd" & _
           " AND a.bldsrc=c.bldsrc AND a.bldyy=c.bldyy AND a.bldno=c.bldno AND a.compocd=c.compocd " & _
           " GROUP BY a.compocd,a.abo,b.groupcd"
           
    
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    If Not RS.EOF Then
        With tblData
            Do Until RS.EOF
                Select Case RS.Fields("abo").Value & ""
'                    Case "A":  NewRow = 1
'                    Case "B":  NewRow = 7
'                    Case "O":  NewRow = 13
'                    Case "AB": NewRow = 19
                    Case "A":  NewRow = 1
                    Case "B":  NewRow = lngCompCnt * 1 + 1
                    Case "O":  NewRow = lngCompCnt * 2 + 1
                    Case "AB": NewRow = lngCompCnt * 3 + 1
                End Select
'                .Row = NewRow + Val(Mid(RS.Fields("groupcd").Value & "", 2)) - 1
                .Row = NewRow + Val(RS.Fields("groupcd").Value & "") - 1
                Select Case RS.Fields("div").Value & ""
                    Case "1": .Col = 5
                    Case "2": .Col = 6
                    Case "3": .Col = 7
                End Select
                .Value = RS.Fields("cnt").Value & ""
                RS.MoveNext
            Loop
        End With
    End If
    Set RS = Nothing
End Sub

Private Sub TotalInventory()
'    Dim EntCnt(0 To 5)      As Long '입고(구입)
'    Dim DonationCnt(0 To 5) As Long '헌혈입고
'    Dim DeliveryCnt(0 To 5) As Long '출고
'    Dim ExpCnt(0 To 5)      As Long '폐기
'    Dim ReturnCnt(0 To 5)   As Long '반납
'    Dim MonTotalCnt(0 To 5) As Long '금월재고
'
'    Dim Trans1(0 To 5)      As Long
'    Dim Trans2(0 To 5)      As Long
'    Dim TotTrans(0 To 5)    As Long
    Dim EntCnt()       As Long '입고(구입)
    Dim DonationCnt() As Long  '헌혈입고
    Dim DeliveryCnt() As Long   '출고
    Dim ExpCnt()       As Long '폐기
    Dim ReturnCnt()    As Long '반납
    Dim MonTotalCnt() As Long  '금월재고
    
    Dim Trans1()       As Long
    Dim Trans2()       As Long
    Dim TotTrans()     As Long
    
    Dim ii          As Long
    Dim jj          As Long
    
    '총계를 구하기 위한 변수 선언
    Dim lngEntCnt As Long
    Dim lngDonaCnt As Long
    Dim lngDelCnt As Long
    Dim lngExpCnt As Long
    Dim lngRetCnt As Long
    Dim lngMoncnt As Long
    Dim lng1PCnt As Long
    Dim lng2PCnt As Long
    Dim lngTPCnt As Long
    
    ReDim EntCnt(0 To lngCompCnt - 1)      '입고(구입)
    ReDim DonationCnt(0 To lngCompCnt - 1) '헌혈입고
    ReDim DeliveryCnt(0 To lngCompCnt - 1)  '출고
    ReDim ExpCnt(0 To lngCompCnt - 1)      '폐기
    ReDim ReturnCnt(0 To lngCompCnt - 1)   '반납
    ReDim MonTotalCnt(0 To lngCompCnt - 1) '금월재고
    
    ReDim Trans1(0 To lngCompCnt - 1)
    ReDim Trans2(0 To lngCompCnt - 1)
    ReDim TotTrans(0 To lngCompCnt - 1)
    
    
    '입출고 건수
    With tblData
        For ii = 1 To .MaxRows - 1
            .Row = ii
            
            For jj = 3 To 8
                .Col = jj
                Select Case .Col
'                    Case 3: EntCnt(Val(ii - 1) Mod 6) = EntCnt(Val(ii - 1) Mod 6) + Val(.Value)
'                    Case 4: DonationCnt(Val(ii - 1) Mod 6) = DonationCnt(Val(ii - 1) Mod 6) + Val(.Value)
'                    Case 5: DeliveryCnt(Val(ii - 1) Mod 6) = DeliveryCnt(Val(ii - 1) Mod 6) + Val(.Value)
'                    Case 6: ExpCnt(Val(ii - 1) Mod 6) = ExpCnt(Val(ii - 1) Mod 6) + Val(.Value)
'                    Case 7: ReturnCnt(Val(ii - 1) Mod 6) = ReturnCnt(Val(ii - 1) Mod 6) + Val(.Value)
'                    Case 8: MonTotalCnt(Val(ii - 1) Mod 6) = MonTotalCnt(Val(ii - 1) Mod 6) + Val(.Value)
                    Case 3: EntCnt(Val(ii - 1) Mod lngCompCnt) = EntCnt(Val(ii - 1) Mod lngCompCnt) + Val(.Value)
                    Case 4: DonationCnt(Val(ii - 1) Mod lngCompCnt) = DonationCnt(Val(ii - 1) Mod lngCompCnt) + Val(.Value)
                    Case 5: DeliveryCnt(Val(ii - 1) Mod lngCompCnt) = DeliveryCnt(Val(ii - 1) Mod lngCompCnt) + Val(.Value)
                    Case 6: ExpCnt(Val(ii - 1) Mod lngCompCnt) = ExpCnt(Val(ii - 1) Mod lngCompCnt) + Val(.Value)
                    Case 7: ReturnCnt(Val(ii - 1) Mod lngCompCnt) = ReturnCnt(Val(ii - 1) Mod lngCompCnt) + Val(.Value)
                    Case 8: MonTotalCnt(Val(ii - 1) Mod lngCompCnt) = MonTotalCnt(Val(ii - 1) Mod lngCompCnt) + Val(.Value)
                End Select
            Next
        Next
        
'        For ii = 25 To 30
        For ii = lngCompCnt * 4 + 1 To lngCompCnt * 5
            .Row = ii
            For jj = 3 To 8
                .Col = jj
                Select Case .Col
'                    Case 3: .Value = IIf(EntCnt(ii - 25) = 0, "", EntCnt(ii - 25))
'                    Case 4: .Value = IIf(DonationCnt(ii - 25) = 0, "", DonationCnt(ii - 25))
'                    Case 5: .Value = IIf(DeliveryCnt(ii - 25) = 0, "", DeliveryCnt(ii - 25)) 'DeliveryCnt(ii - 25)
'                    Case 6: .Value = IIf(ExpCnt(ii - 25) = 0, "", ExpCnt(ii - 25)) 'ExpCnt(ii - 25)
'                    Case 7: .Value = IIf(ReturnCnt(ii - 25) = 0, "", ReturnCnt(ii - 25)) 'ReturnCnt(ii - 25)
'                    Case 8: .Value = IIf(MonTotalCnt(ii - 25) = 0, "", MonTotalCnt(ii - 25)) 'MonTotalCnt(ii - 25)
                    Case 3
                        .Value = IIf(EntCnt(ii - (lngCompCnt * 4 + 1)) = 0, "", EntCnt(ii - (lngCompCnt * 4 + 1)))
                        lngEntCnt = lngEntCnt + Val(.Value)
                    Case 4
                        .Value = IIf(DonationCnt(ii - (lngCompCnt * 4 + 1)) = 0, "", DonationCnt(ii - (lngCompCnt * 4 + 1)))
                        lngDonaCnt = lngDonaCnt + Val(.Value)
                    Case 5
                        .Value = IIf(DeliveryCnt(ii - (lngCompCnt * 4 + 1)) = 0, "", DeliveryCnt(ii - (lngCompCnt * 4 + 1))) 'DeliveryCnt(ii - 25)
                        lngDelCnt = lngDelCnt + Val(.Value)
                    Case 6
                        .Value = IIf(ExpCnt(ii - (lngCompCnt * 4 + 1)) = 0, "", ExpCnt(ii - (lngCompCnt * 4 + 1))) 'ExpCnt(ii - 25)
                        lngExpCnt = lngExpCnt + Val(.Value)
                    Case 7
                        .Value = IIf(ReturnCnt(ii - (lngCompCnt * 4 + 1)) = 0, "", ReturnCnt(ii - (lngCompCnt * 4 + 1))) 'ReturnCnt(ii - 25)
                        lngRetCnt = lngRetCnt + Val(.Value)
                    Case 8
                        .Value = IIf(MonTotalCnt(ii - (lngCompCnt * 4 + 1)) = 0, "", MonTotalCnt(ii - (lngCompCnt * 4 + 1))) 'MonTotalCnt(ii - 25)
                        lngMoncnt = lngMoncnt + Val(.Value)
                End Select
            Next
        Next
        
        Call .SetText(3, .MaxRows, IIf(lngEntCnt = 0, "", lngEntCnt))
        Call .SetText(4, .MaxRows, IIf(lngDonaCnt = 0, "", lngDonaCnt))
        Call .SetText(5, .MaxRows, IIf(lngDelCnt = 0, "", lngDelCnt))
        Call .SetText(6, .MaxRows, IIf(lngExpCnt = 0, "", lngExpCnt))
        Call .SetText(7, .MaxRows, IIf(lngRetCnt = 0, "", lngRetCnt))
        Call .SetText(8, .MaxRows, IIf(lngMoncnt = 0, "", lngMoncnt))
    End With
        
    '수혈 맞은갯수
    With tblTrans
        For ii = 1 To .MaxRows - 1
            .Row = ii
            For jj = 3 To 5
                .Col = jj
                Select Case .Col
'                    Case 3: Trans1(Val(ii - 1) Mod 6) = Trans1(Val(ii - 1) Mod 6) + Val(.Value)
'                    Case 4: Trans2(Val(ii - 1) Mod 6) = Trans2(Val(ii - 1) Mod 6) + Val(.Value)
'                    Case 5: TotTrans(Val(ii - 1) Mod 6) = TotTrans(Val(ii - 1) Mod 6) + Val(.Value)
                    Case 3: Trans1(Val(ii - 1) Mod lngCompCnt) = Trans1(Val(ii - 1) Mod lngCompCnt) + Val(.Value)
                    Case 4: Trans2(Val(ii - 1) Mod lngCompCnt) = Trans2(Val(ii - 1) Mod lngCompCnt) + Val(.Value)
                    Case 5: TotTrans(Val(ii - 1) Mod lngCompCnt) = TotTrans(Val(ii - 1) Mod lngCompCnt) + Val(.Value)
                End Select
            Next jj
        Next ii
'        For ii = 25 To 30
        For ii = lngCompCnt * 4 + 1 To lngCompCnt * 5
            .Row = ii
            For jj = 3 To 8
                .Col = jj
                Select Case .Col
                    Case 3
                        .Value = IIf(Trans1(ii - (lngCompCnt * 4 + 1)) = 0, "", Trans1(ii - (lngCompCnt * 4 + 1)))
                        lng1PCnt = lng1PCnt + Val(.Value)
                    Case 4
                        .Value = IIf(Trans2(ii - (lngCompCnt * 4 + 1)) = 0, "", Trans2(ii - (lngCompCnt * 4 + 1)))
                        lng2PCnt = lng2PCnt + Val(.Value)
                    Case 5
                        .Value = IIf(TotTrans(ii - (lngCompCnt * 4 + 1)) = 0, "", TotTrans(ii - (lngCompCnt * 4 + 1))) 'DeliveryCnt(ii - 25)
                        lngTPCnt = lngTPCnt + Val(.Value)
                End Select
            Next
        Next
        
        Call .SetText(3, .MaxRows, IIf(lng1PCnt = 0, "", lng1PCnt))
        Call .SetText(4, .MaxRows, IIf(lng2PCnt = 0, "", lng2PCnt))
        Call .SetText(5, .MaxRows, IIf(lngTPCnt = 0, "", lngTPCnt))
    End With
End Sub

Private Sub TotalTransfusion()
    Dim RS          As Recordset
    Dim SSQL        As String
    Dim NewRow      As Long
    Dim NewCol      As Long
    Dim TotTrans    As Long
    Dim CLOSEDT     As String
    Dim ii          As Long
    
    CLOSEDT = Format(dtpClose.Value, "yyyymm")
    
    SSQL = "SELECT d.groupcd,(a.deliverycnt-a.retcnt) as transcnt ,c.abo,c.compocd  " & _
           "FROM " & T_BBS006 & " d," & T_BBS401 & " c," & T_BBS203 & " a," & T_BBS302 & " b " & _
           "WHERE " & _
                 DBW("b.vfydt>=", CLOSEDT & "01") & _
           " AND " & DBW("b.vfydt<=", CLOSEDT & "31") & _
           " AND " & DBW("c.centercd=", ObjSysInfo.BuildingCd) & _
           " AND a.workarea=b.workarea AND a.accdt=b.accdt AND a.accseq=b.accseq" & _
           " AND b.bldsrc=c.bldsrc AND b.bldyy=c.bldyy AND b.bldno=c.bldno AND b.compocd =c.compocd" & _
           " AND b.compocd=d.compocd" & _
           " AND (b.cancelfg='0' or b.cancelfg is null or b.cancelfg='')"
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    With tblTrans
        If Not RS.EOF Then
            Do Until RS.EOF
                Select Case RS.Fields("abo").Value & ""
'                    Case "A":  NewRow = 1
'                    Case "B":  NewRow = 7
'                    Case "O":  NewRow = 13
'                    Case "AB": NewRow = 19
                    Case "A":  NewRow = 1
                    Case "B":  NewRow = lngCompCnt * 1 + 1
                    Case "O":  NewRow = lngCompCnt * 2 + 1
                    Case "AB": NewRow = lngCompCnt * 3 + 1
                End Select
'                .Row = NewRow + Val(Mid(RS.Fields("groupcd").Value & "", 2)) - 1
                .Row = NewRow + Val(RS.Fields("groupcd").Value & "") - 1
                
                If Val(RS.Fields("transcnt").Value & "") > 0 Then
                    .Col = IIf(Val(RS.Fields("transcnt").Value & "") > 1, 4, 3)
                    .Value = Val(.Value) + 1
                End If
                RS.MoveNext
            Loop
'            For ii = 1 To 24
            For ii = 1 To lngCompCnt * 4
                .Row = ii: .Col = 3: TotTrans = Val(.Value)
                .Row = ii: .Col = 4: TotTrans = TotTrans + Val(.Value)
                .Row = ii: .Col = 5: .Value = IIf(TotTrans = 0, "", TotTrans)
            Next
        End If
    End With
End Sub

Private Sub IrrCount()
    Dim RS      As Recordset
    Dim SSQL    As String
    Dim CLOSEDT As String
    
    lblIrrCount.Caption = "0 건"
    CLOSEDT = Format(dtpClose.Value, "yyyymm")
    
    SSQL = " SELECT COUNT(*) as cnt FROM " & T_BBS401 & _
         " WHERE " & DBW("irrfg=", "1") & _
         " AND " & DBW("irrdt>=", CLOSEDT & "01") & _
         " AND " & DBW("irrdt<=", CLOSEDT & "31")
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    If Not RS.EOF Then
        lblIrrCount.Caption = RS.Fields("cnt").Value & "" & " 건"
    End If
    Set RS = Nothing
End Sub

Private Sub tblData_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    Dim objPop As clsPopupMenu
    Set objPop = New clsPopupMenu
    With objPop
        .AddMenu MENU_LEFT, "Print..."
        .AddMenu MENU_SEP, "-"
        .AddMenu MENU_LEXCEL, "To Excel"
        .PopupMenus Me.hWnd
        
        If .MenuID = MENU_LEFT Then
            Call PrintLeftInfo
        ElseIf .MenuID = MENU_LEXCEL Then
            Call ExportLeft
        End If
    End With
    
    Set objPop = Nothing
End Sub

Private Sub tblTrans_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    Dim objPop As clsPopupMenu
    Set objPop = New clsPopupMenu
    With objPop
        .AddMenu MENU_RIGHT, "Print..."
        .AddMenu MENU_SEP, "-"
        .AddMenu MENU_REXCEL, "To Excel"
        .PopupMenus Me.hWnd
        
        If .MenuID = MENU_RIGHT Then
            Call PrintRightInfo
        ElseIf .MenuID = MENU_REXCEL Then
            Call ExportRight
        End If
    End With
    
    Set objPop = Nothing
End Sub
