VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmBBS919 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   0  '없음
   Caption         =   "Form1"
   ClientHeight    =   9015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10875
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9015
   ScaleWidth      =   10875
   ShowInTaskbar   =   0   'False
   Begin MedControls1.LisLabel LisLabel11 
      Height          =   315
      Left            =   75
      TabIndex        =   5
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
      Caption         =   "혈액제제별 입고현황"
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   1140
      Left            =   75
      TabIndex        =   6
      Top             =   285
      Width           =   10770
      Begin VB.OptionButton optDiv 
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFFF&
         Caption         =   "출고현황"
         ForeColor       =   &H80000008&
         Height          =   435
         Index           =   1
         Left            =   1245
         Style           =   1  '그래픽
         TabIndex        =   14
         Top             =   375
         Width           =   990
      End
      Begin VB.OptionButton optDiv 
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFFF&
         Caption         =   "입고현황"
         ForeColor       =   &H80000008&
         Height          =   465
         Index           =   0
         Left            =   225
         Style           =   1  '그래픽
         TabIndex        =   13
         Top             =   375
         Value           =   -1  'True
         Width           =   990
      End
      Begin VB.CheckBox chkHos 
         BackColor       =   &H00DBE6E6&
         Caption         =   "병원헌혈제외"
         Height          =   255
         Left            =   6840
         TabIndex        =   12
         Top             =   525
         Value           =   1  '확인
         Width           =   1380
      End
      Begin VB.CommandButton cmdQuery 
         BackColor       =   &H00F4F0F2&
         Caption         =   "조회(&Q)"
         Height          =   510
         Left            =   9300
         Style           =   1  '그래픽
         TabIndex        =   7
         Tag             =   "124"
         Top             =   465
         Width           =   1320
      End
      Begin MSComCtl2.DTPicker dtpFMonth 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "gg yyyy-MM-dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   0
         EndProperty
         Height          =   345
         Left            =   3750
         TabIndex        =   8
         Top             =   465
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   609
         _Version        =   393216
         CustomFormat    =   "yyyy-MM"
         Format          =   62324739
         CurrentDate     =   36799
      End
      Begin MSComCtl2.DTPicker dtpTMonth 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "gg yyyy-MM-dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   0
         EndProperty
         Height          =   345
         Left            =   5490
         TabIndex        =   9
         Top             =   465
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   609
         _Version        =   393216
         CustomFormat    =   "yyyy-MM"
         Format          =   62324739
         CurrentDate     =   36799
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   3
         Left            =   2595
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   465
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "~"
         Height          =   180
         Left            =   5160
         TabIndex        =   11
         Top             =   540
         Width           =   135
      End
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00F4F0F2&
      Caption         =   "출력(&P)"
      Height          =   510
      Left            =   8175
      Style           =   1  '그래픽
      TabIndex        =   2
      Tag             =   "15101"
      Top             =   8400
      Width           =   1320
   End
   Begin VB.CommandButton cmdExcel 
      BackColor       =   &H00F4F0F2&
      Caption         =   "Excel(&E)"
      Height          =   510
      Left            =   6855
      Style           =   1  '그래픽
      TabIndex        =   1
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
      TabIndex        =   0
      Tag             =   "128"
      Top             =   8400
      Width           =   1320
   End
   Begin FPSpread.vaSpread tblList 
      Height          =   6780
      Left            =   75
      TabIndex        =   3
      Tag             =   "10114"
      Top             =   1470
      Width           =   10725
      _Version        =   196608
      _ExtentX        =   18918
      _ExtentY        =   11959
      _StockProps     =   64
      AllowMultiBlocks=   -1  'True
      AutoCalc        =   0   'False
      AutoClipboard   =   0   'False
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      EditEnterAction =   5
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FormulaSync     =   0   'False
      GrayAreaBackColor=   16777215
      GridShowVert    =   0   'False
      MaxCols         =   10
      MaxRows         =   27
      MoveActiveOnFocus=   0   'False
      OperationMode   =   3
      ProcessTab      =   -1  'True
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   14737632
      ShadowDark      =   12632256
      ShadowText      =   0
      SpreadDesigner  =   "frmBBS919.frx":0000
      StartingColNumber=   2
      VirtualRows     =   24
      VisibleRows     =   13
   End
   Begin FPSpread.vaSpread tblexcel 
      Height          =   675
      Left            =   0
      TabIndex        =   4
      Top             =   2850
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
      SpreadDesigner  =   "frmBBS919.frx":06C7
   End
   Begin MSComDlg.CommonDialog DlgSave 
      Left            =   5070
      Top             =   2550
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmBBS919"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum tblColumn
    tcComp = 1
    tcVOL
    tcA
    tcB
    tcAB
    tcO
    tcCnt
    tcMoney
    tcTotMoney
    tcCompocd
End Enum

Private objDic  As clsDictionary

Private Sub GetCompoMoney()
    Dim SSQL    As String
    Dim RS      As Recordset
    Dim objSql  As clsMastrSQL
    
    Set objDic = New clsDictionary
    Set objSql = New clsMastrSQL
    
    objDic.Clear
    objDic.FieldInialize "compo,vol", "money"
    
    SSQL = objSql.GetCompoMoney
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    If Not RS.EOF Then
        Do Until RS.EOF
            If Not objDic.Exists(RS.Fields("cdval1").Value & "" & COL_DIV & RS.Fields("cdval2").Value & "") Then
                objDic.AddNew RS.Fields("cdval1").Value & "" & COL_DIV & RS.Fields("cdval2").Value & "", RS.Fields("field1").Value & ""
            End If
            RS.MoveNext
        Loop
    End If
    Set RS = Nothing
End Sub
Private Sub SpreadDsp(ByVal SSQL As String)
    Dim RS      As Recordset
    Dim strTmp  As String
    Dim blnFirst As Boolean
    Dim lngCnt   As Integer
    Dim strVol   As String
    
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    If Not RS.EOF Then
        With tblList
            .ReDraw = False
            Do Until RS.EOF
                If strTmp <> RS.Fields("compocd").Value & "" Then
                    If blnFirst = True Then
                        .Col = tblColumn.tcCnt:     .Value = lngCnt:
                        .Col = tblColumn.tcVOL: strVol = .Value
                        lngCnt = 0
                    End If
                    blnFirst = True
                    If .DataRowCnt + 1 > .MaxRows Then .MaxRows = .MaxRows + 1
                    .Row = .DataRowCnt + 1
                    .RowHeight(.Row) = 15
                End If
                
                .Col = tblColumn.tcComp:    .Value = RS.Fields("componm").Value & ""
                .Col = tblColumn.tcVOL:     .Value = RS.Fields("volumn").Value & ""
                .Col = tblColumn.tcCompocd: .Value = RS.Fields("compocd").Value & ""
                Select Case RS.Fields("abo").Value & ""
                    Case "A": .Col = tblColumn.tcA
                    Case "B": .Col = tblColumn.tcB
                    Case "O": .Col = tblColumn.tcO
                    Case "AB": .Col = tblColumn.tcAB
                End Select
                .Value = RS.Fields("cnt").Value & "": lngCnt = lngCnt + Val(.Value)
                If objDic.Exists(RS.Fields("compocd").Value & "" & COL_DIV & RS.Fields("volumn").Value & "") Then
                    objDic.KeyChange RS.Fields("compocd").Value & "" & COL_DIV & RS.Fields("volumn").Value & ""
                    .Col = tblColumn.tcMoney: .Value = Format(Val(objDic.Fields("money")), "#,###")
                    .Col = tblColumn.tcTotMoney: .Value = Format(Val(objDic.Fields("money")) * Val(lngCnt), "#,###")
                End If
                         
                strTmp = RS.Fields("compocd").Value & ""
                
                RS.MoveNext
            Loop
            .Row = .DataRowCnt:
            .Col = tblColumn.tcCnt: .Value = lngCnt
            .Col = tblColumn.tcVOL: strVol = .Value
            If objDic.Exists(strTmp & COL_DIV & strVol) Then
                objDic.KeyChange strTmp & COL_DIV & strVol
                .Col = tblColumn.tcMoney: .Value = Format(Val(objDic.Fields("money")), "#,###")
                .Col = tblColumn.tcTotMoney: .Value = Format(Val(objDic.Fields("money")) * Val(lngCnt), "#,###")
                
            End If
            
            .ReDraw = True
        End With
    End If
End Sub


Private Sub cmdQuery_Click()
    Dim objSql    As clsStatics
    Dim RS        As Recordset
    Dim strFrdate As String
    Dim strToDate As String
    Dim strTmp    As String
    Dim lngA      As Long
    Dim lngAB     As Long
    Dim lngO      As Long
    Dim lngB      As Long
    Dim lngTot    As Long
    
    Dim ii        As Integer
    
    Dim SSQL      As String
    
    Set objSql = New clsStatics
    Call medClearTable(tblList)


    
    strFrdate = Format(dtpFMonth.Value, "YYYYMM") & "01"
    strToDate = Format(dtpTMonth.Value, "YYYYMM") & "31"
    
    If optDiv(0).Value Then
        SSQL = objSql.GetBloodStatic(strFrdate, strToDate, "320", "0", chkHos.Value)
        Call SpreadDsp(SSQL)
        SSQL = objSql.GetBloodStatic(strFrdate, strToDate, "400", "0", chkHos.Value)
        Call SpreadDsp(SSQL)
        SSQL = objSql.GetBloodStatic(strFrdate, strToDate, "250", "0", chkHos.Value)
        Call SpreadDsp(SSQL)
    ElseIf optDiv(1).Value Then
        SSQL = objSql.GetBloodStaticForDelivery(strFrdate, strToDate, "320", "0", chkHos.Value)
        Call SpreadDsp(SSQL)
        SSQL = objSql.GetBloodStaticForDelivery(strFrdate, strToDate, "400", "0", chkHos.Value)
        Call SpreadDsp(SSQL)
        SSQL = objSql.GetBloodStaticForDelivery(strFrdate, strToDate, "250", "0", chkHos.Value)
        Call SpreadDsp(SSQL)
    End If
    
    With tblList
        .ReDraw = False
        .SortBy = SortByRow
        .SortKey(1) = tblColumn.tcCompocd
        .SortKey(2) = tblColumn.tcVOL

        .SortKeyOrder(1) = SortKeyOrderAscending
        .SortKeyOrder(2) = SortKeyOrderAscending
        
        .Col = 1:  .Col2 = .MaxCols
        .Row = 1:  .Row2 = .MaxRows
        .BlockMode = True
        .FontBold = False
        .Action = 25
        .BlockMode = False
        .ReDraw = True
        strTmp = ""
        For ii = 1 To .DataRowCnt
            .Row = ii
            .Col = tblColumn.tcCompocd
            If strTmp = .Value Then
                .Col = tblColumn.tcComp: .Value = ""
            Else
                .Col = tblColumn.tcComp: .FontBold = True
            End If
            .Col = tblColumn.tcVOL
            Select Case .Value
                Case "320": .ForeColor = DCM_LightRed
                Case "400": .ForeColor = DCM_LightBlue
                Case "250": .ForeColor = vbBlack
            End Select
            .FontBold = True
            
            .Col = tblColumn.tcA: lngA = lngA + Val(.Value)
            .Col = tblColumn.tcB: lngB = lngB + Val(.Value)
            .Col = tblColumn.tcAB: lngAB = lngAB + Val(.Value)
            .Col = tblColumn.tcO: lngO = lngO + Val(.Value)
            .Col = tblColumn.tcCnt: lngTot = lngTot + Val(.Value)
            
            
            .Col = tblColumn.tcCompocd: strTmp = .Value
        Next
        .Row = .DataRowCnt + 1: .Col = tblColumn.tcCompocd: .Value = "Tot"
        
        .Row = .DataRowCnt + 1: .RowHeight(.Row) = 15
        
        .Col = tblColumn.tcComp:    .Value = "합계": .ForeColor = DCM_LightBlue: .FontBold = True
        .Col = tblColumn.tcA:       .Value = Format(lngA, "#,###")
        .Col = tblColumn.tcB:       .Value = Format(lngB, "#,###")
        .Col = tblColumn.tcAB:      .Value = Format(lngAB, "#,###")
        .Col = tblColumn.tcO:       .Value = Format(lngO, "#,###")
        
        .Col = tblColumn.tcCnt:     .Value = Format(lngTot, "#,###")
        
        strTmp = ""
        For ii = 1 To .DataRowCnt
            .Row = ii: .Col = tblColumn.tcTotMoney
            strTmp = Val(strTmp) + Val(Replace(.Value, ",", ""))
        Next
        .Row = .DataRowCnt: .Col = tblColumn.tcTotMoney: .Value = Format(strTmp, "#,###")
        .Col = tblColumn.tcA:   .Col2 = tblColumn.tcTotMoney
        .Row = 1:               .Row2 = .DataRowCnt
        .BlockMode = True
        .FontBold = True: .TypeHAlign = TypeHAlignRight: .TypeVAlign = TypeVAlignCenter
        .BlockMode = False
    End With
    
    Set objSql = Nothing
    
End Sub

Private Sub Form_Load()
    dtpFMonth.Value = GetSystemDate
    dtpTMonth.Value = GetSystemDate
    Call GetCompoMoney
End Sub


Private Sub cmdExit_Click()
    Unload Me
    Set objDic = Nothing
End Sub
Private Sub cmdExcel_Click()
    Dim strTmp As String
    Dim lngRows As Long
    
    If tblList.DataRowCnt = 0 And tblList.DataRowCnt = 0 Then Exit Sub
    
    With tblList
        .Row = 0: .Row2 = .MaxRows
        .Col = 1: .Col2 = .MaxCols
        .BlockMode = True
        strTmp = .Clip
        .BlockMode = False
        lngRows = .MaxRows
    End With
 
    With tblexcel
        .MaxRows = tblList.MaxRows + 1
        .MaxCols = tblList.MaxCols
        .Row = 1: .Row2 = .MaxRows
        .Col = 1: .Col2 = tblList.MaxCols
        .BlockMode = True
        .Clip = strTmp
        .BlockMode = False
    End With
    
    DlgSave.InitDir = "C:\"
    DlgSave.Filter = "ExCelFile(*.XLS)|*.XLS"
    DlgSave.FileName = LisLabel11.Caption
    DlgSave.ShowSave

    tblexcel.SaveTabFile (DlgSave.FileName)
End Sub

Private Sub cmdPrint_Click()
    With tblList
    
        .Row = 1: .Row2 = .DataRowCnt
        .Col = 1: .Col2 = .MaxCols
        .BlockMode = True
        .PrintJobName = "혈액제제별 " & IIf(optDiv(0).Value, "입고현황", "출고현황") & " 출력"
        .PrintAbortMsg = "혈액제제별 " & IIf(optDiv(0).Value, "입고현황", "출고현황") & " 출력중 입니다. "

        .PrintColor = False
        .PrintFirstPageNumber = 1

        .PrintHeader = "/n/n/l/fb1 " & "♧  혈액제제별 " & IIf(optDiv(0).Value, "입고현황", "출고현황") & " 출력 (" & Format(dtpFMonth.Value, "yyyy년 MM월") & " 부터 " & _
                                                              Format(dtpFMonth.Value, "yyyy년 MM월") & " 까지 ) /c/fb1/n"
                                       
        .PrintFooter = " /l " & String(116, Chr(6)) & "/n/l " & HOSPITAL_MAIN & "/c/p/fb1"
     
        .PrintMarginBottom = 100
        .PrintMarginLeft = 200
        .PrintMarginRight = 100
        .PrintShadows = False
        .PrintMarginTop = 500
        .PrintNextPageBreakCol = 1
        .PrintNextPageBreakRow = 1
        .PrintRowHeaders = False
        .PrintColHeaders = True
        .PrintBorder = True
        .PrintGrid = True
        .GridSolid = False
        .PrintType = PrintTypeAll

        .Action = ActionPrint

        .GridSolid = True
    End With
'    Dim objReport   As clsBBSPrint
'    Dim ii          As Integer
'
'
'    Dim strHeader1 As String
'    Dim strHeader2 As String
'    Dim strHeader3 As String
'    Dim strBody    As String
'    Dim strTmp     As String
'
'    If tblList.MaxRows = 0 Then Exit Sub
'    Set objReport = New clsBBSPrint
'
'    Me.MousePointer = 11
'
'    strHeader1 = "혈액제제별 입고현황"
'
'    strHeader2 = " ♣ 조회일자 : " & Format(dtpFMonth, "yyyy-mm") & " ~ " & Format(dtpTMonth.Value, "YYYY-MM")
'    strHeader2 = strHeader2 & "    ♣ 출력일시 : " & Format(Getsystemdate, "YYYY-MM-DD HH:MM") & "    ♣ 작성자 : " & ObjSysInfo.EmpNm & COL_DIV & "5" & COL_DIV & "1"
'
'    strHeader3 = "혈액제제" & COL_DIV & "5" & COL_DIV & "0"
'    strHeader3 = strHeader3 & vbTab & "용량" & COL_DIV & "70" & COL_DIV & "0"
'    strHeader3 = strHeader3 & vbTab & "A형" & COL_DIV & "85" & COL_DIV & "0"
'    strHeader3 = strHeader3 & vbTab & "B형" & COL_DIV & "100" & COL_DIV & "0"
'    strHeader3 = strHeader3 & vbTab & "AB형" & COL_DIV & "115" & COL_DIV & "0"
'    strHeader3 = strHeader3 & vbTab & "O형" & COL_DIV & "130" & COL_DIV & "0"
'    strHeader3 = strHeader3 & vbTab & "수량" & COL_DIV & "145" & COL_DIV & "0"
'    strHeader3 = strHeader3 & vbTab & "단가" & COL_DIV & "160" & COL_DIV & "0"
'    strHeader3 = strHeader3 & vbTab & "금액" & COL_DIV & "175" & COL_DIV & "1"
'
'    With tblList
'        For ii = 1 To .DataRowCnt
'            .Row = ii
'
'            .Col = tblColumn.tcComp
'            strBody = strBody & .Value & COL_DIV & "5" & COL_DIV & "0" & COL_DIV & "0"
'
'            .Col = tblColumn.tcVOL
'            strBody = strBody & vbTab & .Value & COL_DIV & "70" & COL_DIV & "0" & COL_DIV & "0"
'
'            .Col = tblColumn.tcA
'            strBody = strBody & vbTab & .Value & COL_DIV & "85" & COL_DIV & "0" & COL_DIV & "0"
'
'            .Col = tblColumn.tcB
'            strBody = strBody & vbTab & .Value & COL_DIV & "100" & COL_DIV & "0" & COL_DIV & "0"
'
'            .Col = tblColumn.tcAB
'            strBody = strBody & vbTab & .Value & COL_DIV & "115" & COL_DIV & "0" & COL_DIV & "0"
'
'            .Col = tblColumn.tcO
'            strBody = strBody & vbTab & .Value & COL_DIV & "130" & COL_DIV & "0" & COL_DIV & "0"
'
'            .Col = tblColumn.tcCnt
'            strBody = strBody & vbTab & .Value & COL_DIV & "145" & COL_DIV & "0" & COL_DIV & "0"
'            .Col = tblColumn.tcMoney
'            strBody = strBody & vbTab & .Value & COL_DIV & "160" & COL_DIV & "0" & COL_DIV & "0"
'
''            .Col = .MaxCols
''            If strTmp <> .Value Then
''                strTmp = "1"
''            Else
''                strTmp = ""
''            End If
'            .Col = tblColumn.tcTotMoney
'            strBody = strBody & vbTab & .Value & COL_DIV & "175" & COL_DIV & "1" & COL_DIV & strTmp & vbTab
''            .Col = .MaxCols
''            strTmp = .Value
'
'
'        Next
'    End With
'
'    strBody = strBody & vbTab & "" & COL_DIV & "130" & COL_DIV & "1" & COL_DIV & "0"
'    strBody = strBody & vbTab & "" & COL_DIV & "130" & COL_DIV & "1" & COL_DIV & "0"
'    strBody = strBody & vbTab & "" & COL_DIV & "130" & COL_DIV & "1" & COL_DIV & "0"
'
'    strBody = strBody & vbTab & " 담당자 :    " & ObjSysInfo.EmpNm & " ( 인 ) " & COL_DIV & "130" & COL_DIV & "1" & COL_DIV & "0"
'    strBody = strBody & vbTab & " 확인자 :    하경임 ( 인 ) " & COL_DIV & "130" & COL_DIV & "1" & COL_DIV & "0"
'
'
'    strBody = Mid(strBody, 1, Len(strBody) - 1)
'    With objReport
'        .Header1 = strHeader1
'        .Header2 = strHeader2
'        .Header3 = strHeader3
'        .Body = strBody
'        .mvarSpace = 8
'        Call .CallPrint
'    End With
'    Set objReport = Nothing
'    Me.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objDic = Nothing
End Sub

Private Sub optDiv_Click(Index As Integer)
    If Index = 0 Then
        LisLabel11.Caption = "혈액제제별 입고현황"
        Call medClearTable(tblList)
    ElseIf Index = 1 Then
        LisLabel11.Caption = "혈액제제별 출고현황"
        Call medClearTable(tblList)
    End If
End Sub
