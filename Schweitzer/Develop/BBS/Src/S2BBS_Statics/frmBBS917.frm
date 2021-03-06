VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmBBS917 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   0  '없음
   Caption         =   "방사선 혈액제제"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10995
   Icon            =   "frmBBS917.frx":0000
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   10995
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdExcel 
      BackColor       =   &H00F4F0F2&
      Caption         =   "Excel(&E)"
      Height          =   510
      Left            =   6855
      Style           =   1  '그래픽
      TabIndex        =   13
      Tag             =   "124"
      Top             =   8440
      Width           =   1320
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00F4F0F2&
      Caption         =   "출력(&P)"
      Height          =   510
      Left            =   8175
      Style           =   1  '그래픽
      TabIndex        =   12
      Tag             =   "15101"
      Top             =   8440
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   9510
      Style           =   1  '그래픽
      TabIndex        =   11
      Tag             =   "128"
      Top             =   8440
      Width           =   1320
   End
   Begin FPSpread.vaSpread tblexcel 
      Height          =   675
      Left            =   0
      TabIndex        =   7
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
      SpreadDesigner  =   "frmBBS917.frx":000C
   End
   Begin MSComDlg.CommonDialog DlgSave 
      Left            =   15
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MedControls1.LisLabel LisLabel11 
      Height          =   315
      Left            =   75
      TabIndex        =   8
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
      Caption         =   "방사선조사 혈액제제"
      Appearance      =   0
   End
   Begin FPSpread.vaSpread tblBlood 
      Height          =   6780
      Left            =   75
      TabIndex        =   5
      Tag             =   "10114"
      Top             =   1545
      Width           =   10770
      _Version        =   196608
      _ExtentX        =   18997
      _ExtentY        =   11959
      _StockProps     =   64
      AllowUserFormulas=   -1  'True
      AutoCalc        =   0   'False
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GridShowVert    =   0   'False
      MaxCols         =   5
      MaxRows         =   27
      Protect         =   0   'False
      ScrollBars      =   2
      ShadowColor     =   14737632
      ShadowDark      =   12632256
      ShadowText      =   0
      SpreadDesigner  =   "frmBBS917.frx":01B5
      StartingColNumber=   2
      VirtualRows     =   24
      VisibleRows     =   13
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   1230
      Left            =   75
      TabIndex        =   0
      Top             =   300
      Width           =   10770
      Begin VB.ComboBox cboCenter 
         Height          =   300
         ItemData        =   "frmBBS917.frx":0795
         Left            =   1290
         List            =   "frmBBS917.frx":0797
         Style           =   2  '드롭다운 목록
         TabIndex        =   1
         Top             =   225
         Width           =   2505
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
         Height          =   315
         Left            =   1290
         TabIndex        =   2
         Top             =   660
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM"
         Format          =   62455811
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
         Height          =   315
         Left            =   2715
         TabIndex        =   3
         Top             =   660
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM"
         Format          =   62455811
         CurrentDate     =   36799
      End
      Begin VB.CommandButton cmdQuery 
         BackColor       =   &H00F4F0F2&
         Caption         =   "조회(&Q)"
         Height          =   510
         Left            =   9300
         Style           =   1  '그래픽
         TabIndex        =   6
         Tag             =   "124"
         Top             =   420
         Width           =   1320
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   3
         Left            =   150
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   630
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
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   0
         Left            =   150
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   225
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
         Caption         =   "Center"
         Appearance      =   0
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "~"
         Height          =   180
         Left            =   2460
         TabIndex        =   4
         Top             =   720
         Width           =   135
      End
   End
End
Attribute VB_Name = "frmBBS917"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum tblCol
    tcCompNm = 1
    tcCompCd
    tcEnt
    tcDel
    tcExp
End Enum



Private Sub cmdExcel_Click()
    Dim strTmp As String
    Dim lngRows As Long
    
    If tblBlood.DataRowCnt = 0 And tblBlood.DataRowCnt = 0 Then Exit Sub
    
    With tblBlood
        .Row = 0: .Row2 = .MaxRows
        .Col = 1: .Col2 = .MaxCols
        .BlockMode = True
        strTmp = .Clip
        .BlockMode = False
        lngRows = .MaxRows
    End With
 
    With tblexcel
        .MaxRows = tblBlood.MaxRows + 1
        .MaxCols = tblBlood.MaxCols
        .Row = 1: .Row2 = .MaxRows
        .Col = 1: .Col2 = tblBlood.MaxCols
        .BlockMode = True
        .Clip = strTmp
        .BlockMode = False
    End With
    
    DlgSave.InitDir = "C:\"
    DlgSave.Filter = "ExCelFile(*.XLS)|*.XLS"
    DlgSave.FileName = "방사선 혈액제제 사용량"
    DlgSave.ShowSave

    tblexcel.SaveTabFile (DlgSave.FileName)
End Sub


Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    With tblBlood
    
        .Row = 1: .Row2 = .DataRowCnt
        .Col = 1: .Col2 = .MaxCols
        .BlockMode = True
        .PrintJobName = "방사선 혈액제제 사용량 출력"
        .PrintAbortMsg = "방사선 혈액제제 사용량 출력중 입니다. "

        .PrintColor = False
        .PrintFirstPageNumber = 1

        .PrintHeader = "/n/n/l/fb1 " & "♧ 방사선 혈액제제 사용량 (" & Format(dtpFMonth.Value, CS_DateLongFormat) & " 부터 " & _
                                                              Format(dtpTMonth.Value, CS_DateLongFormat) & " 까지 ) /c/fb1/n" & _
                                       " ♧ 센 터 : " & medGetP(cboCenter.Text, 1, COL_DIV) & "/n/n"
                                       
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
End Sub

Private Sub cmdQuery_Click()
    Me.MousePointer = 11
    Call ClearAll
    Call Query
    Me.MousePointer = 0
End Sub

Private Sub Form_Load()
    dtpFMonth.Value = GetSystemDate
    dtpTMonth.Value = GetSystemDate
    Call SetCenterCombo
    ClearAll
End Sub

Private Sub ClearAll()
    medClearTable tblBlood
End Sub

Private Sub SetCenterCombo()
    Dim objcom003 As clsCom003
    Dim i As Long
    
    Set objcom003 = New clsCom003
    Call objcom003.AddComboBox(BC2_CENTER, cboCenter, True)
    Set objcom003 = Nothing
    
    cboCenter.ListIndex = -1
    
    For i = 0 To cboCenter.ListCount - 1
        If ObjSysInfo.BuildingCd = medGetP(cboCenter.List(i), 1, " ") Then
            cboCenter.ListIndex = i
            Exit For
        End If
    Next i
End Sub
Private Sub Query()
    Dim SSQL        As String
    Dim sFDate      As String
    Dim sTDate      As String
    
    Dim sCentercd   As String
    Dim EntTot      As Long
    Dim DelTot      As Long
    Dim ExpTot      As Long
    Dim blnFirst    As Boolean
    Dim blnChk      As Boolean
    
    Dim ii          As Integer
    Dim jj          As Integer
    
    Dim RS          As Recordset
    
    sCentercd = medGetP(cboCenter.Text, 1, " ")
    
    sFDate = Format(dtpFMonth.Value, "YYYYMM") & "01"
    sTDate = Format(dtpTMonth.Value, "YYYYMM") & "31"
    
    
    '입고량
    SSQL = " SELECT distinct a.compocd,c.componm,count(*) as cnt ,'E' as Div " & _
           " FROM " & T_BBS006 & " c," & T_BBS401 & " a" & _
           " WHERE " & _
                     DBW("a.entdt>=", sFDate) & _
           " AND " & DBW("a.entdt<=", sTDate) & _
           " AND " & DBW("a.irrfg=", "1")
           
    If sCentercd <> "(ALL)" Then
        SSQL = SSQL & " AND " & DBW("a.centercd=", sCentercd)
    End If
    SSQL = SSQL & _
           " AND a.compocd=c.compocd" & _
           " GROUP BY a.compocd,c.componm"
    '출고량
    SSQL = SSQL & " UNION ALL" & _
           " SELECT distinct b.compocd,c.componm,count(*) as cnt ,'D' as div " & _
           " FROM " & T_BBS006 & " c," & T_BBS401 & " b," & T_BBS402 & " a" & _
           " WHERE" & _
                     DBW("a.deliverydt>=", sFDate) & _
           " AND " & DBW("a.deliverydt<=", sTDate) & _
           " AND " & DBW("b.stscd=", BBSBloodStatus.stsDELIVERY) & _
           " AND " & DBW("b.irrfg=", "1")
    If sCentercd <> "(ALL)" Then
        SSQL = SSQL & " AND " & DBW("b.centercd=", sCentercd)
    End If
    SSQL = SSQL & _
           " AND a.bldsrc=b.bldsrc AND a.bldyy=b.bldyy AND a.bldno=b.bldno" & _
           " AND a.compocd=b.compocd" & _
           " AND b.compocd=c.compocd" & _
           " GROUP BY b.compocd,c.componm"
           
    '폐기량
    SSQL = SSQL & "  UNION ALL SELECT distinct a.compocd,c.componm,count(*) as cnt ,'X' as Div " & _
           " FROM " & T_BBS006 & " c," & T_BBS401 & " a" & _
           " WHERE " & _
                     DBW("a.realexpdt>=", sFDate) & _
           " AND " & DBW("a.realexpdt<=", sTDate) & _
           " AND " & DBW("a.stscd=", BBSBloodStatus.stsEXPIRE) & _
           " AND " & DBW("a.irrfg=", "1")
           
    If sCentercd <> "(ALL)" Then
        SSQL = SSQL & " AND " & DBW("a.centercd=", sCentercd)
    End If
    
    SSQL = SSQL & _
           " AND a.compocd=c.compocd" & _
           " GROUP BY a.compocd,c.componm"
           
    SSQL = SSQL & _
           " ORDER BY compocd"
    
    
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    If Not RS.EOF Then
        With tblBlood
            Do Until RS.EOF

                If blnFirst = False Then
                    If .DataRowCnt + 1 > .MaxRows Then .MaxRows = .MaxRows + 1
                    .Row = .DataRowCnt + 1
                    .Col = tblCol.tcCompNm: .Value = RS.Fields("componm").Value & "": .TypeHAlign = TypeHAlignLeft
                    .Col = tblCol.tcCompCd: .Value = RS.Fields("compocd").Value & ""
                    Select Case RS.Fields("div").Value & ""
                        Case "E": .Col = tblCol.tcEnt: .Value = RS.Fields("cnt").Value & ""
                        Case "D": .Col = tblCol.tcDel: .Value = RS.Fields("cnt").Value & ""
                        Case "X": .Col = tblCol.tcExp: .Value = RS.Fields("cnt").Value & ""
                    End Select
                    blnFirst = True
                Else
                    blnChk = False
                    For ii = 1 To .DataRowCnt
                        .Row = ii: .Col = tblCol.tcCompCd
                        If .Value = RS.Fields("compocd").Value & "" Then
                            Select Case RS.Fields("div").Value & ""
                                Case "E": .Col = tblCol.tcEnt: .Value = RS.Fields("cnt").Value & ""
                                Case "D": .Col = tblCol.tcDel: .Value = RS.Fields("cnt").Value & ""
                                Case "X": .Col = tblCol.tcExp: .Value = RS.Fields("cnt").Value & ""
                            End Select
                            blnChk = True
                            Exit For
                        End If
                    Next
                    If blnChk = False Then
                        If .DataRowCnt + 1 > .MaxRows Then .MaxRows = .MaxRows + 1
                        .Row = .DataRowCnt + 1
                        .Col = tblCol.tcCompNm: .Value = RS.Fields("componm").Value & "": .TypeHAlign = TypeHAlignLeft
                        .Col = tblCol.tcCompCd: .Value = RS.Fields("compocd").Value & ""
                        Select Case RS.Fields("div").Value & ""
                            Case "E": .Col = tblCol.tcEnt: .Value = RS.Fields("cnt").Value & ""
                            Case "D": .Col = tblCol.tcDel: .Value = RS.Fields("cnt").Value & ""
                            Case "X": .Col = tblCol.tcExp: .Value = RS.Fields("cnt").Value & ""
                        End Select
                    End If
                End If
                RS.MoveNext
            Loop
            '합계 계산
            For ii = 1 To .DataRowCnt
                .Row = ii
                .Col = tblCol.tcDel: DelTot = DelTot + Val(.Value)
                .Col = tblCol.tcEnt: EntTot = EntTot + Val(.Value)
                .Col = tblCol.tcExp: ExpTot = ExpTot + Val(.Value)
            Next
            If .DataRowCnt + 2 > .MaxRows Then
                .MaxRows = .MaxRows + 2
            End If
            .Row = .DataRowCnt + 2
            .Col = tblCol.tcCompNm: .Value = " 합  계"
            .Col = tblCol.tcEnt: .Value = IIf(EntTot = 0, "", EntTot)
            .Col = tblCol.tcDel: .Value = IIf(DelTot = 0, "", DelTot)
            .Col = tblCol.tcExp: .Value = IIf(ExpTot = 0, "", ExpTot)
            
        End With
    End If
    Set RS = Nothing
    
End Sub




