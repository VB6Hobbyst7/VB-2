VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Anato_Jeobsu_Print 
   BorderStyle     =   0  '없음
   Caption         =   "접수환자출력"
   ClientHeight    =   6930
   ClientLeft      =   1410
   ClientTop       =   1890
   ClientWidth     =   12315
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6930
   ScaleWidth      =   12315
   ShowInTaskbar   =   0   'False
   WindowState     =   2  '최대화
   Begin VB.Frame Frame1 
      Height          =   3015
      Left            =   10500
      TabIndex        =   11
      Top             =   4815
      Width           =   1755
      Begin Threed.SSCommand cmdPrint 
         Height          =   900
         Left            =   60
         TabIndex        =   12
         Top             =   1110
         Width           =   1635
         _Version        =   65536
         _ExtentX        =   2879
         _ExtentY        =   1587
         _StockProps     =   78
         Caption         =   "출 력"
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         Font3D          =   3
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "ANATO112.frx":0000
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   900
         Left            =   60
         TabIndex        =   13
         Top             =   2010
         Width           =   1635
         _Version        =   65536
         _ExtentX        =   2879
         _ExtentY        =   1587
         _StockProps     =   78
         Caption         =   "종 료"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         Font3D          =   3
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "ANATO112.frx":031A
      End
      Begin Threed.SSCommand cmdView 
         Height          =   900
         Left            =   60
         TabIndex        =   14
         Top             =   210
         Width           =   1635
         _Version        =   65536
         _ExtentX        =   2879
         _ExtentY        =   1587
         _StockProps     =   78
         Caption         =   "조 회"
         ForeColor       =   8388736
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         Font3D          =   3
         RoundedCorners  =   0   'False
         AutoSize        =   1
         Picture         =   "ANATO112.frx":0634
      End
   End
   Begin FPSpread.vaSpread ssResult 
      Height          =   7755
      Left            =   150
      TabIndex        =   1
      Top             =   825
      Width           =   10065
      _Version        =   196608
      _ExtentX        =   17754
      _ExtentY        =   13679
      _StockProps     =   64
      BackColorStyle  =   1
      ColsFrozen      =   5
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GridColor       =   8421376
      MaxCols         =   13
      MaxRows         =   600
      ShadowColor     =   12632256
      ShadowDark      =   8421504
      ShadowText      =   0
      SpreadDesigner  =   "ANATO112.frx":0A86
      UserResize      =   0
      VisibleCols     =   12
      VisibleRows     =   500
      ScrollBarTrack  =   3
   End
   Begin Threed.SSPanel SSPanel2 
      Align           =   1  '위 맞춤
      Height          =   765
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12315
      _Version        =   65536
      _ExtentX        =   21722
      _ExtentY        =   1349
      _StockProps     =   15
      Caption         =   "접  수  환  자  출  력"
      ForeColor       =   8388608
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   20.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelInner      =   1
      RoundedCorners  =   0   'False
      Outline         =   -1  'True
      Font3D          =   2
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   1965
      Left            =   10500
      TabIndex        =   2
      Top             =   840
      Width           =   1755
      _Version        =   65536
      _ExtentX        =   3096
      _ExtentY        =   3466
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelInner      =   1
      Begin MSComCtl2.DTPicker dtFromJeobsu 
         Height          =   315
         Left            =   180
         TabIndex        =   6
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   24444931
         CurrentDate     =   36312
      End
      Begin MSComCtl2.DTPicker dtToJeobsu 
         Height          =   315
         Left            =   180
         TabIndex        =   7
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   24444931
         CurrentDate     =   36312
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   5
         Top             =   510
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   1140
         Width           =   210
      End
      Begin VB.Label Label2 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00808000&
         BorderStyle     =   1  '단일 고정
         Caption         =   "접수일자"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   1485
      End
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   1560
      Left            =   10500
      TabIndex        =   8
      Top             =   2985
      Width           =   1755
      _Version        =   65536
      _ExtentX        =   3096
      _ExtentY        =   2752
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSOption optRefferal 
         Height          =   315
         Left            =   210
         TabIndex        =   15
         Top             =   1110
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   556
         _StockProps     =   78
         Caption         =   "Refferal"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.01
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption optCytology 
         Height          =   255
         Left            =   210
         TabIndex        =   9
         Top             =   735
         Width           =   1260
         _Version        =   65536
         _ExtentX        =   2222
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Cytology"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.01
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption optHistology 
         Height          =   255
         Left            =   210
         TabIndex        =   10
         Top             =   300
         Width           =   1260
         _Version        =   65536
         _ExtentX        =   2222
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Histology"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.01
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "Anato_Jeobsu_Print"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdExit_Click()

    Unload Me

End Sub


Private Sub Form_Load()
    
    dtFromJeobsu.Value = Dual_Date_Get("yyyy-MM-dd")
    dtToJeobsu.Value = Dual_Date_Get("yyyy-MM-dd")
        
    optHistology.Value = True
        
End Sub


Private Sub cmdView_Click()
    
    Dim LsPtNo           As String * 8
    Dim LsStatus         As String * 1
    Dim LsCodeKy         As String
    Dim LsDrCode         As String * 6
    Dim LsDeptCode       As String * 4
    Dim LiReccnt         As Integer
    Dim i                As Integer
    Dim LsRet
    
    gSFrDate = Format(dtFromJeobsu.Value, "yyyy-MM-dd")
    gSToDate = Format(dtToJeobsu.Value, "yyyy-MM-dd")
    
    Call SSInitialize(ssResult)
    
    strSQL = ""
    strSQL = strSQL & " SELECT a.*, "
    strSQL = strSQL & "        TO_CHAR(a.Jdate,   'YYYY-MM-DD') Jdate1, "
    strSQL = strSQL & "        TO_CHAR(a.Orderdt, 'YYYY-MM-DD') Orderdt,"
    strSQL = strSQL & "        b.Deptnamek, c.Drname"
    strSQL = strSQL & " FROM   TWANAT_DIAG  a,"
    strSQL = strSQL & "        TWBAS_DEPT   b,"
    strSQL = strSQL & "        TWBAS_DOCTOR c "
    strSQL = strSQL & " WHERE  a.Jdate   BETWEEN TO_DATE('" & gSFrDate & "','YYYY-MM-DD')"
    strSQL = strSQL & "                      AND TO_DATE('" & gSToDate & "','YYYY-MM-DD')"
'    strSQL = strSQL & " AND    a.GbResult = '0'"
'    strSQL = strSQL & " AND    a.GbGross  = '0'"
    strSQL = strSQL & " AND    a.Deptcode = b.DeptCode(+)"
    strSQL = strSQL & " AND    a.Drcode   = c.Drcode(+)"
    If optHistology.Value = True Then
        strSQL = strSQL & " AND    a.Class   = 'P' "
    ElseIf optCytology.Value = True Then
        strSQL = strSQL & " AND    a.Class   = 'C' "
    ElseIf optRefferal.Value = True Then
        strSQL = strSQL & " AND    a.Class   = 'R' "
    End If
    
    strSQL = strSQL & " ORDER  BY Jdate, a.Class, a.Dateyy, a.Seqnum"
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then Exit Sub
        
    Do Until rs.EOF
        ssResult.Row = ssResult.DataRowCnt + 1
        ssResult.Col = 2:  ssResult.Text = rs.Fields("Jdate1").Value & ""
        ssResult.Col = 3:  ssResult.Text = rs.Fields("Class").Value & "-" & _
                                           rs.Fields("Dateyy").Value & "-" & _
                                           rs.Fields("Seqnum").Value & ""
        ssResult.Col = 4:  ssResult.Text = rs.Fields("Ptno").Value & ""
        ssResult.Col = 5:  ssResult.Text = rs.Fields("Sname").Value & ""
        ssResult.Col = 6:  ssResult.Text = IIf(rs.Fields("Sex").Value & "" = "M", "M", "F") & "/" & (rs.Fields("AgeYY").Value & "")
        ssResult.Col = 7:  ssResult.Text = OrganE_Load(rs.Fields("OrganPart").Value & "")

'    txtOpname.Text = rs.Fields("Opname").Value & ""
        
        ssResult.Col = 8:  ssResult.Text = illNameE_Load(rs.Fields("Opname").Value & "")
        
        Dim strGBResult
        Select Case (Trim(rs.Fields("GBResult").Value & ""))
                Case 0
                     strGBResult = "접수중"
                Case 1
                     strGBResult = "육안검사"
                Case 2
                     strGBResult = "Preliminary"
                Case 3
                     strGBResult = "판독"
                Case 4
                     strGBResult = "결과완료"
                Case 9
                     strGBResult = "Additional"
                Case "X"
                     strGBResult = "취소"
        End Select
        ssResult.Col = 9:  ssResult.Text = strGBResult
        
        ssResult.Col = 10:  ssResult.Text = IIf(Mid(rs.Fields("itemcd").Value & "", 1, 4) = "8513", "유", "무")
        
        
        If Trim(rs.Fields("RoomCode").Value & "") = "" Then
            ssResult.Col = 11:  ssResult.Text = rs.Fields("Deptcode").Value & ""
        Else
            ssResult.Col = 11:  ssResult.Text = rs.Fields("RoomCode").Value & ""
        End If
        
        ssResult.Col = 12: ssResult.Text = rs.Fields("Drname").Value & ""
        
        rs.MoveNext
    Loop
    AdoCloseSet rs

End Sub


Private Sub cmdPrint_Click()

'    Dim LsLineString        As String
    Dim strHead1
    Dim strHead2
    Dim strHead3
    Dim strHead4
    
    If ssResult.DataRowCnt = 0 Then Exit Sub
    
'    For I = 1 To ssResult.DataRowCnt
'        ssResult.Row = I
'        ssResult.Col = 1
'        If ssResult.Text = "1" Then
'            Click_Check = True
'        End If
'    Next I
'
'    If Click_Check = False Then
'        MsgBox " 출력할 환자를 선택하십시요. "
'        Exit Sub
'    End If
    
'    LineCnt = 0
    
    strFont1 = "/fn""돋움체""/fz""24""/fb1/fi0/fu0/fk0/fs1"
    strFont2 = "/fn""돋움체""/fz""14""/fb0/fi0/fu0/fk0/fs1"

    strHead1 = "/l/f1/c" & "접 수 환 자 명 단     " & "/n"
    strHead2 = "/l/f1/c" & "---------------------------------        " & "/n"
    strHead3 = "/l/f1/r" & "날짜 : " & Format$(Dual_Date_Get("yyyy-MM-dd"), "YYYY 년 MM월 DD일") & "    "
    
'    Printer.Print
'    Printer.FontSize = 24
'    Printer.Print Tab(16); "접 수 환 자 명 단"
'
'    Printer.FontName = "굴림체"
'    Printer.FontSize = 12
'    Printer.FontBold = False
'    Printer.FontItalic = False
'    Printer.FontUnderline = False
'
'    LsLineString = ""
'
'    For j = 1 To 36
'        LsLineString = LsLineString & "-"
'    Next j
'    Printer.Print Tab(30); LsLineString
'
    
    With ssResult
        .PrintAbortMsg = "인쇄중 - 잠시만 기다리십시오."
        .PrintJobName = ""
        .PrintHeader = strFont1 + strHead1 + strFont2 + strHead2 + strHead3 '+ strHead4
        .PrintMarginLeft = 700 '450
        .PrintMarginRight = 0
        .PrintMarginTop = 860
        .PrintMarginBottom = 200
        
        .PrintColHeaders = True
        .PrintRowHeaders = False
        .PrintBorder = True
        .PrintColor = False
        .PrintGrid = True
        .PrintShadows = False
        .PrintUseDataMax = False 'True
        
        .Row = 1:        .Row2 = .DataRowCnt ' .MaxRows
        .Col = 2:        .Col2 = .MaxCols

        .PrintType = SS_PRINT_CELL_RANGE
        .PrintOrientation = SS_PRINTORIENT_PORTRAIT

'        For I = 0 To .MaxRows
'            .Row = I
'            .FontSize = 8
'        Next I
        
'        For I = 2 To .MaxCols
'            .ColWidth(I) = .ColWidth(I) * 0.6    '0.65
'        Next I
         
        .Action = SS_ACTION_PRINT
             
'        For I = 0 To .MaxRows
'            .Row = I
'            .FontSize = 11
'        Next I

'        For I = 2 To .MaxCols
'            .ColWidth(I) = .ColWidth(I) / 0.6    '0.65
'        Next I
    
    End With
    
'    If ssResult.DataRowCnt Mod 6 <> 0 Then
'        LiPageCnt = LiPageCnt + 1
'        Printer.Print Tab(65); "Page : " & LiPageCnt
'        Printer.EndDoc
'    End If
        
Exit Sub


End Sub


'미사용 과거 SOURCE
Private Sub ssResult_Click(ByVal Col As Long, ByVal Row As Long)
    
    Dim i                   As Integer
    
    If Row = 0 And Col = 1 Then
        ssResult.Col = 1:        ssResult.Row = 0
        If ssResult.Text = "A" Then
            ssResult.Col = 1
            ssResult.Row = 0
            ssResult.Text = "C"
            For i = 1 To ssResult.DataRowCnt
                ssResult.Row = i
                ssResult.Text = "0"
            Next i
        Else
            ssResult.Col = 1
            ssResult.Row = 0
            ssResult.Text = "A"
            For i = 1 To ssResult.DataRowCnt
                ssResult.Row = i
                ssResult.Text = "1"
            Next i
        End If
    End If

End Sub

'미사용 과거 SOURCE
Private Sub cmdPrint2_Click()

    Dim i                   As Integer
    Dim j                   As Integer
    Dim LiPageCnt           As Integer
    Dim LineCnt             As Integer
    Dim LsJDate             As String * 10
    Dim LsOrderDt           As String * 10
    Dim LsAnatNo            As String * 11
    Dim LsPtNo              As String * 8
    Dim LsSname             As String * 10
    Dim LsSex               As String * 2
    Dim LsAge               As String * 3
    Dim LsDiagdate          As String * 10
    Dim LsRoomCode          As String * 6
    Dim LsDpName            As String * 16
    Dim LsDrName            As String * 8
    Dim LsChiefName         As String * 10
    Dim LsLineString        As String
    Dim LsRemark            As String
    Dim LsRemark1           As String * 30
    Dim LsRemark2           As String * 120
    Dim LsRemark3           As String * 30
    Dim LsRemark4           As String * 120
    Dim Click_Check         As Boolean
    
    If ssResult.DataRowCnt = 0 Then Exit Sub
    
    For i = 1 To ssResult.DataRowCnt
        ssResult.Row = i
        ssResult.Col = 1
        If ssResult.Text = "1" Then
            Click_Check = True
        End If
    Next i
        
    If Click_Check = False Then
        MsgBox " 출력할 환자를 선택하십시요. "
        Exit Sub
    End If
    
    LineCnt = 0
    
'    GoSub SUB_LINE_PRINT
        
    For i = 1 To ssResult.DataRowCnt
        ssResult.Row = i
        ssResult.Col = 1
        
        If ssResult.Text = "1" Then
            ssResult.Col = 7:        LsJDate = ssResult.Text
            ssResult.Col = 2:        LsAnatNo = ssResult.Text
            ssResult.Col = 3:        LsPtNo = ssResult.Text
            ssResult.Col = 8:        LsOrderDt = ssResult.Text
        
            strSQL = ""
            strSQL = strSQL & " SELECT REMARK4 "
            strSQL = strSQL & "   FROM TWOCS_OCLINICAL "
            strSQL = strSQL & "  WHERE PTNO  = '" & LsPtNo & "' "
            strSQL = strSQL & "    AND BDATE = TO_DATE('" & LsOrderDt & "','YYYY-MM-DD')"
            
            Result = AdoOpenSet(rs, strSQL)
            If Result = False Then
                LsRemark1 = ""
                LsRemark2 = ""
                LsRemark3 = ""
                LsRemark4 = ""
            Else
                LsRemark = rs.Fields("REMARK4").Value & ""
                LsRemark1 = MidH(LsRemark, 1, 30)
                LsRemark2 = MidH(LsRemark, 31, 120)
                LsRemark3 = MidH(LsRemark, 151, 30)
                LsRemark4 = MidH(LsRemark, 181, 120)
            End If
            
            AdoCloseSet rs
            
            ssResult.Col = 4:        LsSname = ssResult.Text
            ssResult.Col = 5:        LsAge = ssResult.Text
            ssResult.Col = 6:        LsSex = ssResult.Text
            ssResult.Col = 9:        LsRoomCode = ssResult.Text
            ssResult.Col = 10:       LsDpName = ssResult.Text
            ssResult.Col = 11:       LsDrName = ssResult.Text
                
            Printer.FontName = "굴림체"
            Printer.FontSize = 11
            Printer.FontBold = False
            Printer.FontItalic = False
            Printer.FontUnderline = False
            
            
            Printer.Print
            Printer.Print Tab(1); "병리번호 : " & LsAnatNo
            Printer.Print Tab(1); "환자번호 : " & LsPtNo;
            Printer.Print Tab(25); "환 자 명 : " & LsSname;
            Printer.Print Tab(50); "성    별 : " & LsSex;
            Printer.Print Tab(70); "연    령 : " & LsAge
            Printer.Print Tab(1); "의 뢰 과 : " & LsDpName;
            Printer.Print Tab(25); "의뢰의사 : " & LsDrName;
            Printer.Print Tab(50); "병    실 : " & LsRoomCode
            Printer.Print Tab(1); "병력사항 : ";
            Printer.Print Tab(13); "Organ : " & LsRemark1
            Printer.Print Tab(13); "Clinical History : " & LsRemark2
            Printer.Print Tab(13); "Procedure : " & LsRemark3
            Printer.Print Tab(13); "Clinical Impression : " & LsRemark4
            Printer.Print
            
'            GoSub SUB_LINE_PRINT
'SUB_LINE_PRINT:
            Printer.FontName = "굴림체"
            Printer.FontSize = 12
            Printer.FontBold = False
            Printer.FontItalic = False
            Printer.FontUnderline = False
            
            LsLineString = ""
                
            For j = 1 To 80
                LsLineString = LsLineString & "-"
            Next j
            Printer.Print LsLineString
            
'            Return
                                               
            LineCnt = LineCnt + 1
            If LineCnt >= 6 Then
                LiPageCnt = LiPageCnt + 1
                Printer.Print Tab(81); "Page : " & LiPageCnt
                Printer.EndDoc
                LineCnt = O
            End If
        
        End If
        
    Next i
    
    If ssResult.DataRowCnt Mod 6 <> 0 Then
        LiPageCnt = LiPageCnt + 1
        Printer.Print Tab(65); "Page : " & LiPageCnt
        Printer.EndDoc
    End If
        
Exit Sub


End Sub

