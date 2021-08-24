VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frmBBS208 
   BackColor       =   &H00DBE6E6&
   Caption         =   "수혈부작용통계"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14490
   Icon            =   "frmBBS208.frx":0000
   LinkTopic       =   "Form6"
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   14490
   WindowState     =   2  '최대화
   Begin VB.CheckBox chkCompo 
      BackColor       =   &H80000005&
      Caption         =   "혈액제제별 부작용건수"
      Height          =   480
      Left            =   3825
      TabIndex        =   10
      Top             =   165
      Width           =   1395
   End
   Begin VB.CheckBox chkRCompo 
      BackColor       =   &H80000005&
      Caption         =   "수혈사유별 부작용건수"
      Height          =   480
      Left            =   7215
      TabIndex        =   9
      Top             =   165
      Width           =   1395
   End
   Begin VB.CheckBox chkReaction 
      BackColor       =   &H80000005&
      Caption         =   "부작용사유별혈액제제건수"
      Height          =   480
      Left            =   5490
      TabIndex        =   8
      Top             =   165
      Width           =   1395
   End
   Begin VB.CommandButton cmdQuery 
      BackColor       =   &H80000005&
      Caption         =   "조회(&Q)"
      Height          =   525
      Left            =   9135
      Style           =   1  '그래픽
      TabIndex        =   6
      Top             =   180
      Width           =   1080
   End
   Begin VB.CommandButton cmdExcel 
      BackColor       =   &H80000005&
      Caption         =   "To_Excel(&E)"
      Height          =   345
      Index           =   0
      Left            =   10980
      Style           =   1  '그래픽
      TabIndex        =   4
      Top             =   735
      Width           =   1485
   End
   Begin VB.CommandButton cmdExcel 
      BackColor       =   &H80000005&
      Caption         =   "To_Excel(&E)"
      Height          =   345
      Index           =   1
      Left            =   10980
      Style           =   1  '그래픽
      TabIndex        =   3
      Top             =   3210
      Width           =   1485
   End
   Begin VB.CommandButton cmdExcel 
      BackColor       =   &H80000005&
      Caption         =   "To_Excel(&E)"
      Height          =   345
      Index           =   2
      Left            =   10980
      Style           =   1  '그래픽
      TabIndex        =   2
      Top             =   5655
      Width           =   1485
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H80000005&
      Caption         =   "종료(&X)"
      Height          =   525
      Left            =   11385
      Style           =   1  '그래픽
      TabIndex        =   1
      Top             =   180
      Width           =   1080
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H80000005&
      Caption         =   "지움(&C)"
      Height          =   525
      Left            =   10260
      Style           =   1  '그래픽
      TabIndex        =   0
      Top             =   180
      Width           =   1080
   End
   Begin FPSpread.vaSpread tblBlood2 
      Height          =   2085
      Left            =   405
      TabIndex        =   5
      Top             =   6015
      Width           =   12045
      _Version        =   196608
      _ExtentX        =   21246
      _ExtentY        =   3678
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
      MaxCols         =   8
      MaxRows         =   8
      ScrollBars      =   2
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      SpreadDesigner  =   "frmBBS208.frx":076A
      Appearance      =   2
      TextTip         =   2
   End
   Begin MSComCtl2.DTPicker dtpYear 
      Height          =   315
      Left            =   2580
      TabIndex        =   7
      Top             =   270
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "yyyy 년"
      Format          =   62455811
      CurrentDate     =   37918
   End
   Begin FPSpread.vaSpread tblBlood 
      Height          =   2100
      Index           =   0
      Left            =   405
      TabIndex        =   11
      Top             =   1095
      Width           =   12045
      _Version        =   196608
      _ExtentX        =   21246
      _ExtentY        =   3704
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
      MaxCols         =   15
      MaxRows         =   7
      ScrollBars      =   2
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      SpreadDesigner  =   "frmBBS208.frx":0CBF
      Appearance      =   2
      TextTip         =   2
   End
   Begin FPSpread.vaSpread tblBlood 
      Height          =   2100
      Index           =   1
      Left            =   405
      TabIndex        =   12
      Top             =   3570
      Width           =   12045
      _Version        =   196608
      _ExtentX        =   21246
      _ExtentY        =   3704
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
      MaxCols         =   15
      MaxRows         =   7
      ScrollBars      =   2
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      SpreadDesigner  =   "frmBBS208.frx":1486
      Appearance      =   2
      TextTip         =   2
   End
   Begin MSComDlg.CommonDialog DlgSave 
      Left            =   -135
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin FPSpread.vaSpread tblexcel 
      Height          =   675
      Left            =   -300
      TabIndex        =   13
      Top             =   1800
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
      SpreadDesigner  =   "frmBBS208.frx":1C3D
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H009E9383&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H009E9383&
      FillColor       =   &H009E9383&
      Height          =   345
      Index           =   2
      Left            =   405
      Top             =   5655
      Width           =   10575
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H009E9383&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H009E9383&
      FillColor       =   &H009E9383&
      Height          =   345
      Index           =   1
      Left            =   405
      Top             =   3195
      Width           =   10575
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H009E9383&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H009E9383&
      FillColor       =   &H009E9383&
      Height          =   345
      Index           =   0
      Left            =   405
      Top             =   735
      Width           =   10575
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "조회년도 :"
      Height          =   285
      Index           =   0
      Left            =   1545
      TabIndex        =   18
      Top             =   315
      Width           =   1125
   End
   Begin VB.Label Label1 
      BackColor       =   &H009E9383&
      Caption         =   "◆  혈액제제별 부작용건수"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   1
      Left            =   405
      TabIndex        =   17
      Top             =   825
      Width           =   4380
   End
   Begin VB.Label Label1 
      BackColor       =   &H009E9383&
      Caption         =   "◆  수혈사유별 부작용건수"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   2
      Left            =   405
      TabIndex        =   16
      Top             =   3285
      Width           =   4380
   End
   Begin VB.Label Label1 
      BackColor       =   &H009E9383&
      Caption         =   "◆  수혈 사유별 혈액제제"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   3
      Left            =   405
      TabIndex        =   15
      Top             =   5745
      Width           =   4380
   End
   Begin VB.Label Label1 
      BackColor       =   &H00CCFFFF&
      Caption         =   "조회조건"
      Height          =   285
      Index           =   4
      Left            =   480
      TabIndex        =   14
      Top             =   285
      Width           =   780
   End
   Begin VB.Shape Shape3 
      Height          =   600
      Left            =   1335
      Top             =   90
      Width           =   7740
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00CCFFFF&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00000000&
      FillColor       =   &H000000C0&
      Height          =   600
      Left            =   390
      Top             =   90
      Width           =   945
   End
End
Attribute VB_Name = "frmBBS208"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const sMode1 = "0"
Private Const sMode2 = "1"
Private Const sMode3 = "2"



Private Sub cmdClear_Click()
    Call FormClear
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call FormClear
End Sub
Private Sub FormClear(Optional ByVal blnClear As Boolean = True)
    If blnClear = True Then
        dtpYear.value = GetSystemDate
        chkCompo.value = 1: chkReaction.value = 1: chkRCompo.value = 1
    End If
    Call medClearTable(tblBlood(0))
    Call medClearTable(tblBlood(1))
    Call medClearTable(tblBlood2)
    Call TblCompoSetting
    Call TblReactionSetting
End Sub

Private Sub TblCompoSetting()
    Dim objSQL  As clsTransfusion
    Dim RS      As Recordset
    Dim SSQL    As String
    
    Set RS = New Recordset
    Set objSQL = New clsTransfusion
    SSQL = objSQL.GetCompoentSQL
    RS.Open SSQL, DBConn
    
    If Not RS.EOF Then
        With tblBlood(0)
            If Not RS.EOF Then
                Do Until RS.EOF
                    If .DataRowCnt + 1 > .MaxRows Then
                        .MaxRows = .MaxRows + 1
                    End If
                    .Row = .DataRowCnt + 1
                    .Col = 1: .value = RS.Fields("componm").value & ""
                    .Col = 15: .value = RS.Fields("compocd").value & ""
                    RS.MoveNext
                Loop
            End If
        End With
    End If
    
    Set RS = Nothing
    Set objSQL = Nothing
End Sub
Private Sub TblReactionSetting()
    Dim objSQL  As clsTransfusion
    Dim RS      As Recordset
    Dim SSQL    As String
    Dim ii      As Long
    
    
    Set RS = New Recordset
    Set objSQL = New clsTransfusion
    SSQL = objSQL.GetReactionSQL
    RS.Open SSQL, DBConn
    If Not RS.EOF Then
        With tblBlood(1)
            If Not RS.EOF Then
                Do Until RS.EOF
                    If .DataRowCnt + 1 > .MaxRows Then
                        .MaxRows = .MaxRows + 1
                    End If
                    .Row = .DataRowCnt + 1
                    .Col = 1: .value = RS.Fields("field1").value & ""
                    .Col = 15: .value = RS.Fields("cdval1").value & ""
                    RS.MoveNext
                Loop
            End If
        End With
        RS.MoveFirst
        Do Until RS.EOF
            With tblBlood2
                If ii + 2 > .MaxRows Then .MaxRows = .MaxRows + 1
                .Row = ii + 2: .Col = 1: .value = Trim(RS.Fields("field1").value & "")
                               .Col = 2: .value = Trim(RS.Fields("cdval1").value & "")
            End With
            ii = ii + 1
            RS.MoveNext
        Loop
    End If
    Set RS = Nothing
    Set objSQL = Nothing
End Sub


Private Sub cmdQuery_Click()
    Call FormClear(False)
    If chkCompo.value = 1 Then Call StaticsticQuery(sMode1)
    If chkReaction.value = 1 Then Call StaticsticQuery(sMode2)
    If chkRCompo.value = 1 Then Call StaticsticQuery(sMode3)
    Call TotSumCalcu
End Sub


Private Sub StaticsticQuery(ByVal sMode As String)
    Dim objSQL  As clsTransfusion
    Dim RS      As Recordset
    Dim SSQL    As String
    Dim sYear   As String
    Dim ii      As Long
    Dim jj      As Integer
    Dim blnChk  As Boolean
    Dim kk      As Long
    
    Dim objPro  As New clsProgress
    
    
    With objPro
        .Container = Me
        .Left = Shape1(Val(sMode)).Left
        .Top = Shape1(Val(sMode)).Top
        .Width = Shape1(Val(sMode)).Width
        .Height = Shape1(Val(sMode)).Height
        .Message = "자료를 수집하고 있습니다..."
'        .Choice = True
'        .SetMyForm Me
'        .XPos = Shape1(Val(sMode)).Left
'        .YPos = Shape1(Val(sMode)).Top
'        .XWidth = Shape1(Val(sMode)).Width
'        .YHeight = Shape1(Val(sMode)).Height
'        .Appearance = aPlate
'        .Msg = "자료를 수집하고 있습니다."
    End With
    
    
    
    sYear = Format(dtpYear.value, "YYYY")
    Set RS = New Recordset
    Set objSQL = New clsTransfusion
    
    SSQL = objSQL.GetReationStatics(sYear, sMode)
    
    RS.Open SSQL, DBConn
    
    objPro.Max = RS.RecordCount
    
    If Not RS.EOF Then
        If sMode < 2 Then
            With tblBlood(CInt(sMode))
                If Not RS.EOF Then
                    Do Until RS.EOF
                        For ii = 1 To .DataRowCnt
                            .Row = ii: .Col = 15
                            If RS.Fields("querydiv").value & "" = .value Then
                                .Col = Val(RS.Fields("reactiondate").value & "") + 1:
                                .value = RS.Fields("cnt").value & ""
                                Exit For
                            End If
                        Next
                        kk = kk + 1
                        objPro.value = kk
                        RS.MoveNext
                    Loop
                End If
            End With
        Else
            With tblBlood2
                .ReDraw = False
                Do Until RS.EOF
                    For ii = 2 To .DataRowCnt
                        .Row = ii: .Col = 2
                        If .value = Trim(RS.Fields("reactioncd").value & "") Then
                            blnChk = False
                            For jj = 3 To .DataColCnt
                                .Row = 1: .Col = jj
                                If .value = RS.Fields("compocd").value & "" Then
                                    .Row = ii: .Col = jj: .value = Format(RS.Fields("cnt").value & "", "#,###")
                                    blnChk = True
                                End If
                            Next
                            
                            If blnChk = False Then
                                If .MaxCols = .DataColCnt Then
                                    .MaxCols = .MaxCols + 1
                                End If
                                .Col = .DataColCnt + 1
                                .ColWidth(.DataColCnt + 1) = 12.5
                                .Row = 0: .value = RS.Fields("abbrnm").value & ""
                                .Row = 1: .value = RS.Fields("compocd").value & ""
                                .Row = ii: .value = Format(RS.Fields("cnt").value & "", "#,###")
                            End If
                            Exit For
                        End If
                    Next
                    kk = kk + 1
                    objPro.value = kk
                    RS.MoveNext
                Loop
            End With
        End If
        
    End If
    
    Set RS = Nothing
    Set objSQL = Nothing
    Set objPro = Nothing
End Sub
Private Sub TotSumCalcu()
    Dim RowTot  As Long
    Dim ColTot  As Long
    Dim lngDataCount As Long
    
    Dim ii      As Long
    Dim jj      As Long
    Dim kk      As Long
    
    For kk = 0 To 1
        If chkCompo.value = 0 And kk = 0 Then GoTo Skip
        If chkReaction.value = 0 And kk = 1 Then GoTo Skip
        With tblBlood(kk)
            lngDataCount = .DataRowCnt
            .MaxRows = lngDataCount + 2
            For ii = 1 To lngDataCount
                .Row = ii
                RowTot = 0
                For jj = 2 To 13
                    .Col = jj
                    RowTot = RowTot + Val(.value)
                Next
                .Col = 14: .value = IIf(RowTot = 0, "", RowTot)
            Next
            
            For ii = 2 To 14
                .Col = ii
                ColTot = 0
                For jj = 1 To lngDataCount
                    .Row = jj
                    ColTot = ColTot + Val(.value)
                Next
                .Row = lngDataCount + 2
                .Col = 1:   .value = "합  계": .TypeHAlign = TypeHAlignRight: .FontBold = True
                .Col = ii:  .value = IIf(ColTot = 0, "", ColTot)
            Next
            If .DataRowCnt < 7 Then .MaxRows = 7
        End With
Skip:
    Next
    If chkRCompo.value = 0 Then Exit Sub
    Call ReactionCompoSum
End Sub
        
Private Sub ReactionCompoSum()
    Dim ComTot As Long
    Dim RsnTot As Long
    Dim RowCnt As Long
    Dim ColCnt As Long
    Dim blnChk  As Boolean
    Dim ii      As Long
    Dim jj      As Long
    
    With tblBlood2
        '제제별 건수
        For ii = 3 To .DataColCnt
            .Col = ii
            For jj = 2 To .DataRowCnt
                .Row = jj
                ComTot = ComTot + Val(.value)
            Next
            
            If blnChk = False Then
                If .DataRowCnt + 2 > .MaxRows Then
                    .MaxRows = .MaxRows + 2
                End If
                blnChk = True
                RowCnt = .DataRowCnt + 2
            End If
            .Row = RowCnt: .value = Format(ComTot, "#,###"): ComTot = 0
            .Col = 1: .value = " 합 계": .TypeHAlign = TypeHAlignRight: .FontBold = True
        Next
        '부작용별 건수
        blnChk = False
        For ii = 2 To .DataRowCnt
            .Row = ii
            For jj = 3 To .DataColCnt
                .Col = jj
                RsnTot = RsnTot + Val(.value)
            Next
            If blnChk = False Then
                If .DataColCnt + 1 > .MaxCols Then
                    .MaxCols = .MaxCols + 1
                End If
                blnChk = True
                ColCnt = .DataColCnt + 1
            End If
            .Row = 0: .Col = ColCnt: .value = " TOTAL ": .ColWidth(ColCnt) = 12.5
            .Row = ii
            .Col = ColCnt: .value = IIf(RsnTot = 0, "", Format(RsnTot, "#,###")): RsnTot = 0
        Next
        
        Select Case ColCnt - 1
            Case 3:
                .MaxCols = 4
                .ColWidth(ColCnt - 1) = (12.5 * 6) / (ColCnt - 2)
                .ColWidth(ColCnt) = (12.5 * 5) / (ColCnt - 2)
            Case 4:
                .MaxCols = 5
                .ColWidth(ColCnt - 1) = (12.5 * 6) / (ColCnt - 2)
                .ColWidth(ColCnt - 2) = (12.5 * 6) / (ColCnt - 2)
                .ColWidth(ColCnt) = (12.5 * 6) / (ColCnt - 2)
            Case 5:
                .MaxCols = 6
                .ColWidth(ColCnt - 1) = (12.5 * 6) / (ColCnt - 2)
                .ColWidth(ColCnt - 2) = (12.5 * 6) / (ColCnt - 2)
                .ColWidth(ColCnt - 3) = (12.5 * 6) / (ColCnt - 2)
                .ColWidth(ColCnt) = (12.5 * 6) / (ColCnt - 2)
            Case 6:
                .MaxCols = 7
                .ColWidth(ColCnt - 1) = (12.5 * 6) / (ColCnt - 2)
                .ColWidth(ColCnt - 2) = (12.5 * 6) / (ColCnt - 2)
                .ColWidth(ColCnt - 3) = (12.5 * 6) / (ColCnt - 2)
                .ColWidth(ColCnt - 4) = (12.5 * 6) / (ColCnt - 2)
                .ColWidth(ColCnt) = (12.5 * 6) / (ColCnt - 2)
            Case 7:
                .MaxCols = 8
                .ColWidth(ColCnt - 1) = (12.5 * 6) / (ColCnt - 2)
                .ColWidth(ColCnt - 2) = (12.5 * 6) / (ColCnt - 2)
                .ColWidth(ColCnt - 3) = (12.5 * 6) / (ColCnt - 2)
                .ColWidth(ColCnt - 4) = (12.5 * 6) / (ColCnt - 2)
                .ColWidth(ColCnt - 5) = (12.5 * 6) / (ColCnt - 2)
                .ColWidth(ColCnt) = (12.5 * 6) / (ColCnt - 2)
                
        End Select
        
    End With
End Sub
        
        
Private Sub cmdExcel_Click(Index As Integer)
    
    Dim strTmp      As String
    Dim lngMaxRow   As Long
    Dim sFileNm     As String
    Dim lngMaxCol   As Long
    
    Call medClearTable(tblexcel)
    Select Case Index
        Case 0: sFileNm = "혈액제제별 부작용건수  "
                If tblBlood(Index).DataRowCnt < 1 Then Exit Sub
        Case 1: sFileNm = "부작용사유별 부작용건수"
                If tblBlood(Index).DataRowCnt < 1 Then Exit Sub
        Case 2: sFileNm = "부작용사유별 혈액제제"
                If tblBlood2.DataRowCnt < 1 Then Exit Sub
    End Select
    If Index < 2 Then
       With tblBlood(Index)
           .Row = 0: .Row2 = .MaxRows
           .Col = 1: .COL2 = .MaxCols - 1
           .BlockMode = True
           strTmp = .Clip
           .BlockMode = False
           lngMaxCol = .MaxCols - 1
           lngMaxRow = .MaxRows
       End With
    Else
       With tblBlood2
           .Row = 0: .Row2 = .MaxRows
           .Col = 1: .COL2 = .MaxCols
           .BlockMode = True
           strTmp = .Clip
           .BlockMode = False
           lngMaxCol = .MaxCols
           lngMaxRow = .MaxRows
       End With
    End If
    
    With tblexcel
       .MaxRows = lngMaxRow + 1
       .MaxCols = lngMaxCol
       .Row = 1: .Row2 = .MaxRows
       .Col = 1: .COL2 = lngMaxCol
       .BlockMode = True
       .Clip = strTmp
       .BlockMode = False
    End With
    
    DlgSave.InitDir = "C:\"
    DlgSave.Filter = "ExCelFile(*.XLS)|*.XLS"
    DlgSave.FileName = sFileNm
    DlgSave.ShowSave

    tblexcel.SaveTabFile (DlgSave.FileName)
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    Set frmReactionStatistic = Nothing
End Sub

