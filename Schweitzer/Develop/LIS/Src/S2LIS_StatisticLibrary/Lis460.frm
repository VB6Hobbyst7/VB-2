VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frm460ItemCnt 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   0  '없음
   Caption         =   "그룹별 검사항목 통계"
   ClientHeight    =   9150
   ClientLeft      =   585
   ClientTop       =   915
   ClientWidth     =   14685
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9150
   ScaleWidth      =   14685
   ShowInTaskbar   =   0   'False
   Tag             =   "검사건수 통계"
   WindowState     =   2  '최대화
   Begin VB.Frame fraPrgBar 
      BackColor       =   &H00AFBCC5&
      Caption         =   "                                                                                    "
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000F5386&
      Height          =   1035
      Left            =   4245
      TabIndex        =   13
      Top             =   4095
      Visible         =   0   'False
      Width           =   6525
      Begin MSComctlLib.ProgressBar Prgbar 
         Height          =   225
         Left            =   60
         TabIndex        =   14
         Top             =   720
         Width           =   6405
         _ExtentX        =   11298
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label5 
         BackColor       =   &H00BDC6CA&
         BackStyle       =   0  '투명
         Caption         =   "데이터를 로드중 입니다."
         Height          =   195
         Left            =   2370
         TabIndex        =   15
         Top             =   300
         Width           =   2025
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00CBD2D6&
         BackStyle       =   1  '투명하지 않음
         Height          =   1035
         Left            =   0
         Top             =   0
         Width           =   6525
      End
   End
   Begin VB.Frame fraDuration 
      BackColor       =   &H00DBE6E6&
      Height          =   780
      Index           =   0
      Left            =   1380
      TabIndex        =   4
      Top             =   -45
      Width           =   13080
      Begin VB.CommandButton cmdStart 
         BackColor       =   &H00DBE6E6&
         Caption         =   "조회(&S)"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   1
         Left            =   10155
         Style           =   1  '그래픽
         TabIndex        =   7
         Tag             =   "158"
         Top             =   165
         Width           =   1320
      End
      Begin VB.CommandButton cmdExcel 
         BackColor       =   &H00DBE6E6&
         Caption         =   "To Excel(&E)"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   1
         Left            =   11490
         Style           =   1  '그래픽
         TabIndex        =   6
         Tag             =   "127"
         Top             =   165
         Width           =   1320
      End
      Begin VB.ComboBox cboBuilding 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6930
         Style           =   2  '드롭다운 목록
         TabIndex        =   5
         Top             =   225
         Width           =   2250
      End
      Begin MSComCtl2.DTPicker dtpFrDt 
         Height          =   360
         Left            =   255
         TabIndex        =   8
         Top             =   240
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd  dddd"
         Format          =   65863680
         UpDown          =   -1  'True
         CurrentDate     =   36339
      End
      Begin MSComCtl2.DTPicker dtpToDt 
         Height          =   360
         Left            =   3465
         TabIndex        =   9
         Top             =   255
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarForeColor=   0
         CustomFormat    =   "dd  dddd"
         Format          =   65863680
         UpDown          =   -1  'True
         CurrentDate     =   36339
      End
      Begin VB.Label Label1 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "부터"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3000
         TabIndex        =   11
         Top             =   285
         Width           =   450
      End
      Begin VB.Label Label3 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "까지"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6285
         TabIndex        =   10
         Top             =   300
         Width           =   450
      End
   End
   Begin MSComDlg.CommonDialog DlgSave 
      Left            =   9330
      Top             =   8430
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   13140
      Style           =   1  '그래픽
      TabIndex        =   1
      Tag             =   "128"
      Top             =   8535
      Width           =   1320
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00F4F0F2&
      Caption         =   "출력(&P)"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   11820
      Style           =   1  '그래픽
      TabIndex        =   0
      Tag             =   "132"
      Top             =   8535
      Width           =   1320
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   675
      Left            =   75
      TabIndex        =   3
      Top             =   45
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   1191
      BackColor       =   10392451
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Caption         =   "기 간"
      Appearance      =   0
      LeftGab         =   0
   End
   Begin FPSpread.vaSpread tblAccCnt 
      Height          =   7560
      Left            =   75
      TabIndex        =   12
      Top             =   810
      Width           =   14385
      _Version        =   196608
      _ExtentX        =   25374
      _ExtentY        =   13335
      _StockProps     =   64
      BackColorStyle  =   1
      ColsFrozen      =   2
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   35
      MaxRows         =   49
      OperationMode   =   1
      ShadowColor     =   13753559
      ShadowDark      =   13753559
      ShadowText      =   0
      SpreadDesigner  =   "Lis460.frx":0000
      UserResize      =   1
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "** 주의 : 기간에 따라 많은 데이타를 검색하므로 속도가 현저히 늦어질 수 있습니다. 업무가 과중한 시간을 피해서 작업하십시오."
      ForeColor       =   &H00000080&
      Height          =   180
      Left            =   1470
      TabIndex        =   2
      Top             =   915
      Width           =   10380
   End
   Begin VB.Shape Shape1 
      Height          =   1905
      Left            =   1995
      Top             =   2910
      Width           =   6015
   End
End
Attribute VB_Name = "frm460ItemCnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ItemRow()   As Integer
Private objStsSql   As New clsLISSqlStatistic

Public Event LastFormUnload()

Private Sub cboBuilding_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"

End Sub

Private Sub cmdExcel_Click(Index As Integer)

    DlgSave.InitDir = App.Path & "\rpt"
    DlgSave.Filter = "ExCelFile(*.XLS,*.MDB)|*.XLS;*.MDB"
    DlgSave.FileName = ""
    DlgSave.ShowSave
     
    If DlgSave.FileName = "" Then Exit Sub
    tblAccCnt.SaveTabFile (DlgSave.FileName)

End Sub

Private Sub cmdExit_Click()
    Unload Me
    If IsLastForm Then RaiseEvent LastFormUnload
End Sub

Private Sub cmdPrint_Click()

    With tblAccCnt
        .PrintOrientation = PrintOrientationLandscape
        .PrintJobName = "통계 출력"
        .PrintAbortMsg = "통계결과를 출력중 입니다. "

        .PrintColor = False
        .PrintFirstPageNumber = 1
       
        .PrintHeader = "/n/n/fb1/l" & " 월간 검사건수 통계 : " & Format(dtpFrDt.Value, "YYYYMMDD") & "년 " & _
                                                                 Format(dtpFrDt.Value, "MM") & "월 " & "/l/fb1/n"
        .PrintFooter = "/c/p/fb1"
        .PrintMarginBottom = 100
        .PrintMarginLeft = 0
        .PrintMarginRight = 0
        .PrintShadows = True
        .PrintMarginTop = 100
        .PrintNextPageBreakCol = 1
        .PrintNextPageBreakRow = 1
        .PrintPageEnd = 2
        .PrintRowHeaders = True
        .PrintGrid = True
        .PrintType = PrintTypeAll
         
        .Action = ActionPrint
    End With
    
End Sub

Private Sub cmdStart_Click(Index As Integer)
    Call ClearTable
    Call CreateDataFile
End Sub

Private Sub dtpFrDt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"

End Sub

Private Sub dtpToDt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"

End Sub

Private Sub Form_Activate()
    MainFrm.lblSubMenu.Caption = Me.Caption
End Sub

Private Sub Form_Load()
    

    dtpFrDt.Value = Format(GetSystemDate, "YYYY-MM-dd")
    dtpToDt.Value = Format(DateAdd("d", 7, GetSystemDate), "YYYY-MM-dd")
    
    ReDim ItemRow(0)

    Call LoadBuildingList
    
End Sub


Private Sub ReadTable(ByVal SqlStmt As String)

    Dim RS      As Recordset
    Dim strTmp  As String
    Dim strTmp1 As String
    Dim strDay  As String
    Dim iFrom   As Integer
    Dim iTo     As Integer
    Dim i       As Integer
    Dim j       As Integer
    Dim iRow    As Integer
    
    Set RS = Nothing: Set RS = New Recordset
    RS.Open SqlStmt, dbconn
    
    Prgbar.Max = Prgbar.Max + RS.RecordCount
    
    iFrom = 0: iTo = 0: i = 1
    strTmp = "zz": strTmp1 = ""
    
    With tblAccCnt
        
        i = .MaxRows
        
        For j = 1 To RS.RecordCount
            
            If strTmp1 <> "" & RS.Fields("TestCd").Value Then
                i = i + 1
                If .MaxRows < i Then .MaxRows = i
                strTmp1 = "" & RS.Fields("TestCd").Value
                
                '검사항목명
                .Row = i:  .Col = 2
                .Value = Trim("" & RS.Fields("TestNm").Value) & "-" & Trim("" & RS.Fields("TestCd").Value)
                '검사항목코드
                .Row = i:  .Col = 35
                .Value = Trim("" & RS.Fields("TestCd").Value)
            
            End If
             
            '검사부서명
            If strTmp <> "" & RS.Fields("WAnm").Value Then
                iTo = i - 1
                If (iFrom > 0) And (iTo - iFrom) >= 0 Then
                    Call SetBorder(iFrom, iTo, 1, 1, vbBlack, CellBorderStyleSolid, 16)  'SS_BORDER_TYPE_OUTLINE
                    Call SetBorder(iFrom, iTo, 2, 2, vbBlack, CellBorderStyleSolid, 16)  'SS_BORDER_TYPE_OUTLINE
                    Call SetBorder(iFrom, iTo, 3, .MaxCols, vbBlack, CellBorderStyleSolid, 16)
                End If
                iFrom = i
                .Row = i: .Col = 1
                .Value = Trim("" & RS.Fields("WAnm").Value)
                strTmp = .Value
            End If
            
            Prgbar.Value = Prgbar.Value + 1
            'DoEvents
            strDay = Mid(Trim("" & RS.Fields("RcvDt").Value), 7, 2)
            .Col = Val(strDay) + 3
            .Row = i   'ItemRow(Val(rs.Fields("ItemSeq").Value))
            .Value = Val(.Value) + Val("" & RS.Fields("Cnt").Value)
            '세로합계
            .Row = 1
            .Value = Val(.Value) + Val("" & RS.Fields("Cnt").Value)
            '가로합계
            .Col = 3
            .Row = i  'ItemRow(Val(rs.Fields("ItemSeq").Value))
            .Value = Val(.Value) + Val("" & RS.Fields("Cnt").Value)
            '전체합계
            .Row = 1
            .Value = Val(.Value) + Val("" & RS.Fields("Cnt").Value)
            RS.MoveNext
            
        Next
                
        iTo = i
        If (iFrom > 0) And (iTo - iFrom) >= 0 Then
            Call SetBorder(iFrom, iTo, 1, 1, vbBlack, CellBorderStyleSolid, 16)  'SS_BORDER_TYPE_OUTLINE
            Call SetBorder(iFrom, iTo, 2, 2, vbBlack, CellBorderStyleSolid, 16)  'SS_BORDER_TYPE_OUTLINE
            Call SetBorder(iFrom, iTo, 3, .MaxCols, vbBlack, CellBorderStyleSolid, 16)
        End If
        
    End With
    
    Set RS = Nothing
    
End Sub


Public Sub SetBorder(ByVal Row1 As Integer, Row2 As Integer, Col1 As Integer, COL2 As Integer, _
                     ByVal BColor As Long, ByVal BStyle As Integer, ByVal BType As Integer, Optional ByVal FColor As Variant)
                     
    With tblAccCnt
        .Row = Row1: .Row2 = Row2
        .Col = Col1: .COL2 = COL2
        .BlockMode = True
        .CellBorderStyle = BStyle
        .CellBorderType = BType
        .CellBorderColor = BColor
        .Action = ActionSetCellBorder
        If Not IsMissing(FColor) Then .ForeColor = FColor
        .BlockMode = False
    End With
    
End Sub




Public Sub CreateDataFile()
    Dim strFrDt     As String
    Dim strToDt     As String
    Dim strColDt    As String
    Dim i           As Integer
    
    strFrDt = Format(dtpFrDt.Value, "YYYYMMDD")
    strToDt = Format(dtpToDt.Value, "YYYYMMDD")
    
    fraPrgBar.Visible = True
    Prgbar.Max = 1
    Prgbar.Value = 0
    DoEvents
    
    With tblAccCnt
        
        .ReDraw = False
        
        .MaxRows = 1   'rs.RecordCount + 1
        
'        For i = 1 To 3
            Call ReadTable(objStsSql.SqlTestCount(i, strColDt, strFrDt, strToDt, medGetP(cboBuilding.Text, 1, " ")))
'        Next
        
        .Row = 1:   .Col = 1
        .FontBold = True
        .Value = "        합        계   "
        .AllowCellOverflow = True
        Call SetBorder(1, 1, 1, 2, vbBlack, CellBorderStyleSolid, 16)
        Call SetBorder(1, .MaxRows, 3, 3, vbBlack, CellBorderStyleSolid, 16, &HDF6A3E)
        Call SetBorder(1, 1, 3, .MaxCols, vbBlack, CellBorderStyleSolid, 16, &H80&)
        .RowsFrozen = 1
        .ColsFrozen = 3
    
        .TopRow = 2
        .LeftCol = 4
        
        .ReDraw = True
        
    End With
    
    fraPrgBar.Visible = False
        
End Sub


Public Sub ReadTable_back(ByVal SqlStmt As String)

    Dim RS      As Recordset
    Dim strDay  As String
    
    Set RS = New Recordset
    RS.Open SqlStmt, dbconn
    
    Prgbar.Max = Prgbar.Max + RS.RecordCount
    
    With tblAccCnt
        Do Until RS.EOF
            Prgbar.Value = Prgbar.Value + 1
            'DoEvents
            strDay = Mid(Trim("" & RS.Fields("RcvDt").Value), 7, 2)
            .Col = Val(strDay) + 3
            .Row = ItemRow(Val("" & RS.Fields("Seq").Value))
            .Value = Val(.Value) + Val("" & RS.Fields("Cnt").Value)
            '세로합계
            .Row = 1
            .Value = Val(.Value) + Val("" & RS.Fields("Cnt").Value)
            '가로합계
            .Col = 3
            .Row = ItemRow(Val("" & RS.Fields("Seq").Value))
            .Value = Val(.Value) + Val("" & RS.Fields("Cnt").Value)
            '전체합계
            .Row = 1
            .Value = Val(.Value) + Val("" & RS.Fields("Cnt").Value)
            RS.MoveNext
        Loop
        .TopRow = 2
        .LeftCol = 4
    End With
    Set RS = Nothing
    
End Sub


Public Sub ClearTable()

    With tblAccCnt
        .MaxRows = 0
    End With
    
End Sub


Public Sub LoadBuildingList()
    Dim tmpRs    As Recordset
    Dim i        As Integer
    Dim SqlStmt  As String
    
    Set tmpRs = Nothing: Set tmpRs = New Recordset
    
    tmpRs.Open objStsSql.GetBuildCd, dbconn
    
    cboBuilding.Clear
    For i = 1 To tmpRs.RecordCount
       cboBuilding.AddItem Trim("" & tmpRs.Fields("BuildCd").Value) & "   " & _
                           Trim("" & tmpRs.Fields("BuildNm").Value)
       tmpRs.MoveNext
    Next
    
    Set tmpRs = Nothing
    cboBuilding.ListIndex = medComboFind(cboBuilding, ObjSysInfo.BuildingCd)
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set objStsSql = Nothing
End Sub
