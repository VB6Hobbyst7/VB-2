VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frmBBS915 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   0  '없음
   Caption         =   "폐기사유별건수"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10995
   Icon            =   "frmBBS915.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   10995
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   9510
      Style           =   1  '그래픽
      TabIndex        =   0
      Tag             =   "128"
      Top             =   8400
      Width           =   1320
   End
   Begin Crystal.CrystalReport CReport 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MedControls1.LisLabel LisLabel11 
      Height          =   315
      Left            =   75
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   45
      Width           =   10770
      _ExtentX        =   18997
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
      Caption         =   "부적격 혈액이송"
      Appearance      =   0
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DBE6E6&
      Height          =   7980
      Left            =   75
      TabIndex        =   1
      Top             =   300
      Width           =   10770
      Begin MSComctlLib.TabStrip tabRsn 
         Height          =   330
         Left            =   435
         TabIndex        =   4
         Top             =   1080
         Width           =   2820
         _ExtentX        =   4974
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "월별폐기건수"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "폐기사유별 건수"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00F4F0F2&
         Height          =   5370
         Left            =   435
         TabIndex        =   9
         Top             =   1335
         Width           =   10050
         Begin FPSpread.vaSpread tblList 
            Height          =   4560
            Left            =   60
            TabIndex        =   10
            Top             =   540
            Width           =   9885
            _Version        =   196608
            _ExtentX        =   17436
            _ExtentY        =   8043
            _StockProps     =   64
            BackColorStyle  =   1
            DisplayRowHeaders=   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "돋움"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GrayAreaBackColor=   14411494
            GridShowVert    =   0   'False
            MaxCols         =   13
            MaxRows         =   19
            OperationMode   =   2
            ScrollBars      =   2
            SelectBlockOptions=   0
            ShadowColor     =   14737632
            ShadowDark      =   14737632
            ShadowText      =   0
            SpreadDesigner  =   "frmBBS915.frx":076A
            TextTip         =   4
         End
         Begin MSComCtl2.DTPicker dtpYear 
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
            Left            =   1155
            TabIndex        =   11
            Top             =   195
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy"
            Format          =   62193667
            CurrentDate     =   36799
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   315
            Index           =   1
            Left            =   75
            TabIndex        =   12
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
            Caption         =   "기 간"
            Appearance      =   0
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00F4F0F2&
         Height          =   5370
         Left            =   435
         TabIndex        =   5
         Top             =   1335
         Width           =   10050
         Begin MSComCtl2.DTPicker dtpMonth 
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
            Left            =   1155
            TabIndex        =   6
            Top             =   195
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy-MM"
            Format          =   62193667
            CurrentDate     =   36799
         End
         Begin FPSpread.vaSpread tblList1 
            Height          =   4455
            Left            =   75
            TabIndex        =   7
            Top             =   540
            Width           =   9840
            _Version        =   196608
            _ExtentX        =   17357
            _ExtentY        =   7858
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
            MaxCols         =   7
            MaxRows         =   13
            OperationMode   =   1
            ShadowColor     =   14737632
            ShadowDark      =   13818331
            SpreadDesigner  =   "frmBBS915.frx":1192
            TextTip         =   4
         End
         Begin MedControls1.LisLabel LisLabel4 
            Height          =   315
            Index           =   0
            Left            =   90
            TabIndex        =   8
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
            Caption         =   "기 간"
            Appearance      =   0
         End
      End
      Begin VB.CommandButton cmdQuery 
         BackColor       =   &H00F4F0F2&
         Caption         =   "조회(&Q)"
         Height          =   510
         Left            =   9135
         Style           =   1  '그래픽
         TabIndex        =   3
         Tag             =   "124"
         Top             =   6705
         Width           =   1320
      End
   End
End
Attribute VB_Name = "frmBBS915"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum tblColumn
    tcRSNNM = 1
    tcMON1
    tcMON2
    tcMON3
    tcMON4
    tcMON5
    tcMON6
    tcMON7
    tcMON8
    tcMON9
    tcMON10
    tcMON11
    tcMON12
End Enum
Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdQuery_Click()
    
    If tabRsn.SelectedItem.Index = 1 Then
        Call Query1
    ElseIf tabRsn.SelectedItem.Index = 2 Then
        Dim ii As Integer
        medClearTable tblList1
        With tblList1
            
            For ii = 2 To .MaxCols
                .Row = 0: .Col = ii: .Value = ""
            Next
        End With
        Call Query2
    End If
    
End Sub
Private Sub Total_Sum()
    Dim total(1 To 12) As Long
    Dim ii           As Integer
    Dim jj           As Integer
    
    With tblList
        For ii = 1 To .MaxRows
            .Row = ii
            For jj = 2 To .MaxCols
                .Col = jj
                total(jj - 1) = total(jj - 1) + .Value
            Next
        Next
        .MaxRows = .MaxRows + 2
        .Row = .MaxRows
        For jj = 2 To .MaxCols
            .Col = jj: .Value = total(jj - 1)
        Next
        .Col = 1: .Value = " 합  계"
    End With
    
End Sub

Private Sub Query1()
    Dim objdic    As New clsDictionary
    Dim objstatic As New clsStatics
    Dim Year      As String
    Dim ii        As Integer
    

    Year = Format(dtpYear.Value, "yyyy")
    medClearTable tblList
    
    Set objdic = objstatic.Get_ExpireRecord(Year)
    If objdic.RecordCount > 0 Then
        objdic.MoveFirst
        With tblList
            .MaxRows = objdic.RecordCount
            Do Until objdic.EOF
                ii = ii + 1
                .Row = ii
                .Col = tblColumn.tcRSNNM: .Value = objdic.Fields("rsnnm")
                .Col = tblColumn.tcMON1:  .Value = objdic.Fields("mon1")
                .Col = tblColumn.tcMON2:  .Value = objdic.Fields("mon2")
                .Col = tblColumn.tcMON3:  .Value = objdic.Fields("mon3")
                .Col = tblColumn.tcMON4:  .Value = objdic.Fields("mon4")
                .Col = tblColumn.tcMON5:  .Value = objdic.Fields("mon5")
                .Col = tblColumn.tcMON6:  .Value = objdic.Fields("mon6")
                .Col = tblColumn.tcMON7:  .Value = objdic.Fields("mon7")
                .Col = tblColumn.tcMON8:  .Value = objdic.Fields("mon8")
                .Col = tblColumn.tcMON9:  .Value = objdic.Fields("mon9")
                .Col = tblColumn.tcMON10: .Value = objdic.Fields("mon10")
                .Col = tblColumn.tcMON11: .Value = objdic.Fields("mon11")
                .Col = tblColumn.tcMON12: .Value = objdic.Fields("mon12")
                objdic.MoveNext
            Loop
        End With
        Call Total_Sum
    End If
    
    Set objdic = Nothing
    Set objstatic = Nothing
End Sub

Private Sub Form_Load()
    Call Clear
End Sub
Private Sub Clear()
    dtpYear.Value = Format(GetSystemDate, "yyyy-mm-dd")
    dtpMonth.Value = Format(GetSystemDate, "YYYY-MM")
    medClearTable tblList
    medClearTable tblList1
   ' tabRef.SelectedItem.Index = 1
End Sub

Private Sub Query2()
    Dim sFDate  As String
    Dim sTDate  As String
    Dim SSQL    As String
    Dim ii      As Integer
    Dim jj      As Integer
    Dim blnChk  As Boolean
    
    Dim RS      As Recordset
    
    sFDate = Format(dtpMonth.Value, "YYYYMM") & "01"
    sTDate = Format(dtpMonth.Value, "YYYYMM") & "31"
    
    '1: 폐기사유를 먼저 화면에 Display
    '1: 폐기된 혈액에 대해서 제제별/폐기사유별로 통계를 Query
    
    
    '1
    SSQL = " SELECT * FROM " & T_COM003 & " WHERE " & DBW("cdindex", BC2_EXP_RESON, 2)
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    If Not RS.EOF Then
        Do Until RS.EOF
            With tblList1
                If ii + 2 > .MaxRows Then .MaxRows = .MaxRows + 1
                .Row = ii + 2: .Col = 1: .Value = Trim(RS.Fields("field1").Value & "")
                               .Col = 2: .Value = Trim(RS.Fields("cdval1").Value & "")
            End With
            ii = ii + 1
            RS.MoveNext
        Loop
    Else
        MsgBox "폐기사유에 대한 마스터가 존재하지 않습니다.", vbInformation + vbOKOnly, "Info"
        Set RS = Nothing
        Exit Sub
    End If
    
    SSQL = " SELECT distinct a.exprsncd ,a.compocd ,b.abbrnm,count(*) as cnt " & _
           " FROM " & T_BBS006 & " b," & T_BBS401 & " a" & _
           " WHERE " & _
                     DBW("a.realexpdt>=", sFDate) & _
           " AND " & DBW("a.realexpdt<=", sTDate) & _
           " AND " & DBW("a.stscd=", BBSBloodStatus.stsEXPIRE) & _
           " AND a.compocd=b.compocd" & _
           " GROUP BY a.exprsncd,a.compocd,b.abbrnm" & _
           " ORDER BY a.exprsncd,a.compocd"
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    If Not RS.EOF Then
        With tblList1
            .ReDraw = False
            Do Until RS.EOF
                For ii = 2 To .DataRowCnt
                    .Row = ii: .Col = 2
                    If .Value = Trim(RS.Fields("exprsncd").Value & "") Then
                        blnChk = False
                        For jj = 3 To .DataColCnt
                            .Row = 1: .Col = jj
                            If .Value = RS.Fields("compocd").Value & "" Then
                                .Row = ii: .Col = jj: .Value = Format(RS.Fields("cnt").Value & "", "#,###")
                                blnChk = True
                            End If
                        Next
                        
                        If blnChk = False Then
                            If .MaxCols = .DataColCnt Then
                                .MaxCols = .MaxCols + 1
                            End If
                            .Col = .DataColCnt + 1
                            .ColWidth(.DataColCnt + 1) = 12.25
                            .Row = 0: .Value = RS.Fields("abbrnm").Value & ""
                            .Row = 1: .Value = RS.Fields("compocd").Value & ""
                            .Row = ii: .Value = Format(RS.Fields("cnt").Value & "", "#,###")
                        End If
                        Exit For
                    End If
                    
                Next
                RS.MoveNext
            Loop
            '합계 계산

            
            Dim ComTot As Long
            Dim RsnTot As Long
            Dim RowCnt As Long
            Dim ColCnt As Long
            blnChk = False
            '제제별 건수
            For ii = 3 To .DataColCnt
                .Col = ii
                For jj = 2 To .DataRowCnt
                    .Row = jj
                    ComTot = ComTot + Val(.Value)
                Next
                
                If blnChk = False Then
                    If .DataRowCnt + 2 > .MaxRows Then
                        .MaxRows = .MaxRows + 2
                    End If
                    blnChk = True
                    RowCnt = .DataRowCnt + 2
                End If
                .Row = RowCnt: .Value = Format(ComTot, "#,###"): ComTot = 0
                .Col = 1: .Value = " 합 계"
            Next
            '폐기사유별 건수
            blnChk = False
            For ii = 2 To .DataRowCnt
                .Row = ii
                For jj = 3 To .DataRowCnt
                    .Col = jj
                    RsnTot = RsnTot + Val(.Value)
                Next
                If blnChk = False Then
                    If .DataColCnt + 1 > .MaxCols Then
                        .MaxCols = .MaxCols + 1
                    End If
                    blnChk = True
                    ColCnt = .DataColCnt + 1
                End If
                .Row = 0: .Col = ColCnt: .Value = " 합 계 ": .ColWidth(ColCnt) = 12.25
                .Row = ii
                .Col = ColCnt: .Value = IIf(RsnTot = 0, "", Format(RsnTot, "#,###")): RsnTot = 0
            Next
            Select Case ColCnt - 1
                Case 3:
                    .MaxCols = 4
                    .ColWidth(ColCnt - 1) = (12.25 * 5) / (ColCnt - 2)
                    .ColWidth(ColCnt) = (12.25 * 5) / (ColCnt - 2)
                Case 4:
                    .MaxCols = 5
                    .ColWidth(ColCnt - 1) = (12.25 * 5) / (ColCnt - 2)
                    .ColWidth(ColCnt - 2) = (12.25 * 5) / (ColCnt - 2)
                    .ColWidth(ColCnt) = (12.25 * 5) / (ColCnt - 2)
                Case 5:
                    .MaxCols = 6
                    .ColWidth(ColCnt - 1) = (12.25 * 5) / (ColCnt - 2)
                    .ColWidth(ColCnt - 2) = (12.25 * 5) / (ColCnt - 2)
                    .ColWidth(ColCnt - 3) = (12.25 * 5) / (ColCnt - 2)
                    .ColWidth(ColCnt) = (12.25 * 5) / (ColCnt - 2)
                Case 6:
                    .MaxCols = 7
                    .ColWidth(ColCnt - 1) = (12.25 * 5) / (ColCnt - 2)
                    .ColWidth(ColCnt - 2) = (12.25 * 5) / (ColCnt - 2)
                    .ColWidth(ColCnt - 3) = (12.25 * 5) / (ColCnt - 2)
                    .ColWidth(ColCnt - 4) = (12.25 * 5) / (ColCnt - 2)
                    
                    .ColWidth(ColCnt) = (12.25 * 5) / (ColCnt - 2)
            End Select
            
            .ReDraw = True
        End With
    Else
        MsgBox "폐기된 혈액이 없습니다.", vbInformation + vbOKOnly, "Info"
    End If
    
    Set RS = Nothing
End Sub

Private Sub tabRsn_Click()
    If tabRsn.SelectedItem.Index = 1 Then
        Frame1.ZOrder 0
    ElseIf tabRsn.SelectedItem.Index = 2 Then

            
        Frame2.ZOrder 0
    End If
End Sub
