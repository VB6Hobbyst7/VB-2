VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frmBBS920 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   0  '없음
   Caption         =   "부적격혈액이송"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10995
   Icon            =   "frmBBS920.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   10995
   ShowInTaskbar   =   0   'False
   Begin MedControls1.LisLabel LisLabel11 
      Height          =   315
      Left            =   75
      TabIndex        =   3
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   1440
      Left            =   75
      TabIndex        =   4
      Top             =   285
      Width           =   10770
      Begin VB.ComboBox cboVol 
         Height          =   300
         ItemData        =   "frmBBS920.frx":076A
         Left            =   1410
         List            =   "frmBBS920.frx":0777
         Style           =   2  '드롭다운 목록
         TabIndex        =   10
         Top             =   1005
         Width           =   2415
      End
      Begin VB.CommandButton cmdQuery 
         BackColor       =   &H00F4F0F2&
         Caption         =   "조회(&Q)"
         Height          =   480
         Left            =   7785
         Style           =   1  '그래픽
         TabIndex        =   9
         Tag             =   "124"
         Top             =   675
         Width           =   1245
      End
      Begin VB.ComboBox cboCompo 
         Height          =   300
         ItemData        =   "frmBBS920.frx":078A
         Left            =   1410
         List            =   "frmBBS920.frx":078C
         Style           =   2  '드롭다운 목록
         TabIndex        =   7
         Top             =   660
         Width           =   2415
      End
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
         Left            =   1410
         TabIndex        =   5
         Top             =   285
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM"
         Format          =   62455811
         CurrentDate     =   36799
      End
      Begin VB.CheckBox chkALL 
         BackColor       =   &H00DBE6E6&
         Caption         =   "전체용량"
         Height          =   255
         Index           =   1
         Left            =   3960
         TabIndex        =   12
         Top             =   1035
         Width           =   1050
      End
      Begin VB.CheckBox chkALL 
         BackColor       =   &H00DBE6E6&
         Caption         =   "전체제제"
         Height          =   255
         Index           =   0
         Left            =   3945
         TabIndex        =   13
         Top             =   705
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  '오른쪽 맞춤
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "조회용량"
         Height          =   180
         Left            =   165
         TabIndex        =   11
         Top             =   1050
         Width           =   720
      End
      Begin VB.Label Label17 
         Alignment       =   1  '오른쪽 맞춤
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Component"
         Height          =   180
         Left            =   150
         TabIndex        =   8
         Top             =   660
         Width           =   975
      End
      Begin VB.Label Label5 
         Alignment       =   1  '오른쪽 맞춤
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "조회기간"
         Height          =   180
         Left            =   165
         TabIndex        =   6
         Top             =   345
         Width           =   720
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
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   9510
      Style           =   1  '그래픽
      TabIndex        =   1
      Tag             =   "128"
      Top             =   8400
      Width           =   1320
   End
   Begin FPSpread.vaSpread tblList 
      Height          =   6525
      Left            =   75
      TabIndex        =   0
      Top             =   1740
      Width           =   10770
      _Version        =   196608
      _ExtentX        =   18997
      _ExtentY        =   11509
      _StockProps     =   64
      DisplayColHeaders=   0   'False
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
      MaxCols         =   12
      MaxRows         =   27
      OperationMode   =   1
      ScrollBars      =   2
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      ShadowText      =   0
      SpreadDesigner  =   "frmBBS920.frx":078E
      TextTip         =   4
   End
   Begin Crystal.CrystalReport CReport 
      Left            =   3960
      Top             =   7740
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frmBBS920"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum tblColumn
    tcNUM = 1
    TcBLOODNO
    tcALT
    tcBINFECTION
    
    tcCINFECTION
    
    tcSYPHILIS
    tcAIDS
    tcETC
    tcTESTDT
    tcCOMPONM
    
    tcrmk
    tcvol
End Enum

Private Sub cmdExit_Click()
    Unload Me
    
End Sub



Private Sub cmdQuery_Click()
    Dim objstatic As New clsStatics
    
    Dim Fdt       As String
    Dim Tdt       As String
    
    Table_Clear
    Fdt = Format(dtpMonth.Value, "yyyymm") & "01"
    Tdt = Format(dtpMonth.Value, "yyyymm") & "31"
    
    Call Not_List(Fdt, Tdt)
    
End Sub
Private Sub Not_List(ByVal Fdt As String, ByVal Tdt As String)
    Dim objstatic As New clsStatics
    Dim objProBar As New clsProgress
    Dim RS        As New Recordset
    Dim RsnRs     As New Recordset
    Dim Rsncd()   As String
    Dim strTmp    As String
    Dim sCompoCd    As String
    Dim sVol        As String
    Dim ii        As Integer
    Dim jj        As Integer
    
    If chkALL(0).Value = 0 Then
        sCompoCd = medGetP(cboCompo.Text, 1, " ")
    Else
        sCompoCd = ""
    End If
    
    If chkALL(1).Value = 0 Then
        sVol = cboVol.Text
    Else
        sVol = ""
    End If
    
    Set RS = objstatic.Get_NotDonorList(Fdt, Tdt, sCompoCd, sVol)
    
    If Not RS.EOF Then
'        Set objProBar.StatusBar = MainFrm.stsbar
        objProBar.Container = MainFrm.stsbar
        objProBar.Max = RS.RecordCount
        
        Set RsnRs = objstatic.NotAcceptRsncd
        
        
        With tblList
            .MaxRows = RS.RecordCount + 2
            .ReDraw = False
            Do Until RS.EOF
                ii = ii + 1
                .Row = ii + 2
                .Col = tblColumn.tcNUM: .Value = ii
                .Col = tblColumn.TcBLOODNO: .Value = RS.Fields("bldsrc").Value & "" & "-" & _
                                                     RS.Fields("bldyy").Value & "" & "-" & _
                                                     Format(RS.Fields("bldno").Value & "", "00000#")
                
                Rsncd() = Split(objstatic.Get_NotDonorRsncd(RS.Fields("donorid").Value & "", RS.Fields("donoraccdt").Value & ""), COL_DIV)
                
                If UBound(Rsncd) > -1 Then
                    For jj = 0 To UBound(Rsncd)
                        RsnRs.MoveFirst
                        Do Until RsnRs.EOF
                            strTmp = RsnRs.Fields("field1").Value & "" & COL_DIV & RsnRs.Fields("field2").Value & ""
                            If InStr(strTmp, Rsncd(jj)) > 0 Then
                                Select Case medGetP(strTmp, 1, COL_DIV)
                                    Case "ALT":     .Col = tblColumn.tcALT:       .Value = "○"
                                    Case "B형간염": .Col = tblColumn.tcBINFECTION: .Value = "○"
                                    Case "C형간염": .Col = tblColumn.tcCINFECTION: .Value = "○"
                                    Case "매독":    .Col = tblColumn.tcSYPHILIS:  .Value = "○"
                                    Case "AIDS":    .Col = tblColumn.tcAIDS:      .Value = "○"
                                    Case "기타":    .Col = tblColumn.tcETC:       .Value = "○"
                                End Select
                            End If
                            RsnRs.MoveNext
                        Loop
                    Next
                End If
                
                .Col = tblColumn.tcTESTDT: .Value = objstatic.Get_NotdonorVdt(RS.Fields("donorid").Value & "", RS.Fields("donoraccdt").Value & "")
                .Col = tblColumn.tcCOMPONM: .Value = RS.Fields("abbrnm").Value & ""
                .Col = tblColumn.tcvol: .Value = RS.Fields("volumn").Value & ""
                    If .Value <> "" Then .Value = .Value & "cc"
                objProBar.Value = ii
                RS.MoveNext
            Loop
            .ReDraw = True
        End With
        cmdPrint.Enabled = True
    Else
        cmdPrint.Enabled = False
        tblList.MaxRows = 2
    End If
    
    Set RS = Nothing
    Set RsnRs = Nothing
    Set objProBar = Nothing
    Set objstatic = Nothing
End Sub
Private Sub Form_Load()
    Dim objSql    As New clsGetSqlStatement
    
    dtpMonth.Value = Format(GetSystemDate, "yyyy-mm")
    tblList.MaxRows = 2

    Call objSql.Compolist(cboCompo)
    Set objSql = Nothing
    cboVol.ListIndex = 0
    cboCompo.ListIndex = 0
    
End Sub
Private Sub Table_Clear()
    Dim ii As Integer
    Dim jj As Integer
    
    With tblList
        For ii = 3 To .MaxRows
            .Row = ii
            For jj = tblColumn.tcNUM To tblColumn.tcrmk
                .Col = jj
                .Value = ""
            Next
        Next
    End With
End Sub

    
    
Private Sub cmdPrint_Click()

    Dim ii      As Integer
    Dim lngCnt  As Integer
    
    Call P_PrtSet
    With tblList
        If .MaxRows < 3 Then Exit Sub
        Call PrinterHeader
        For ii = 3 To .MaxRows
            lngCnt = lngCnt + 1
            If lngCnt > 15 And lngCnt Mod 15 = 1 Then
                Printer.NewPage
                Call PrinterHeader
            End If
            .Row = ii
            Call Print_Setting(lngCnt, 6, 8, 7, , , False)
            .Col = 2: Call Print_Setting(.Value, 13, 8, 27, , , False)
            .Col = 3: Call Print_Setting(.Value, 40, 8, 15, , , False)
            .Col = 4: Call Print_Setting(.Value, 55, 8, 15, , , False)
            .Col = 5: Call Print_Setting(.Value, 70, 8, 15, , , False)
            .Col = 6: Call Print_Setting(.Value, 85, 8, 15, , , False)
            .Col = 7: Call Print_Setting(.Value, 100, 8, 15, , , False)
            .Col = 8: Call Print_Setting(.Value, 115, 8, 15, , , False)
            .Col = 9: Call Print_Setting(.Value, 130, 8, 20, , , False)
            .Col = 10: Call Print_Setting(.Value, 150, 8, 20, , , False)
            .Col = 12: Call Print_Setting(.Value, 170, 8, 20)  ', , , False)
            
        Next
    End With
    Printer.EndDoc
    
End Sub


Private Sub PrinterHeader()
    Dim lngY        As Long
    Dim lngX        As Long
    Dim ii          As Integer
    Dim lngLineY    As Integer
    
    
    lngY = 8
    lngX = 6
    
    
    
    lngCurYPos = 15
    Printer.DrawWidth = 4
    Printer.FontSize = 20: Printer.FontBold = True
    Call Print_Setting(HOSPITAL_MAIN, 6, lngY, Printer.ScaleWidth)
    Call Print_Setting("", lngX, lngY)
    
    Printer.FontSize = 10: Printer.FontBold = False
    Printer.Line (6, lngCurYPos)-(192, lngCurYPos + 16), , B
    Call Print_Setting(HOSPITAL_ADDR & " / 전화 " & HOSPITAL_TEL1 & "(대)/전송" & HOsPITAL_FAX, lngX, lngY, 192, "L")
    Call Print_Setting("담당부서명 : " & HOSPITAL_NAME & "      담당자:  ", lngX, lngY, 192, "L")
    
    Call Print_Setting("", lngX, lngY)
    
    Call Print_Setting("문서번호:", lngX, lngY, lngY, "L", , False)
    Call Print_Setting("시행일자: " & Format(Now, "YYYY.MM.dD"), 100, lngY, lngY, "L")
    Call Print_Setting("수    신:", lngX, lngY, lngY, "L", , False)
    Call Print_Setting("발    신:              (인)", 100, lngY, lngY, "L")
    Call Print_Setting("참    조:", lngX, lngY, lngY, "L")
    Call Print_Setting("", lngX * 2, lngY)
    Call Print_Setting("제    목:", lngX, lngY, lngY, "L", , False)
    Printer.FontSize = 15: Printer.FontBold = True
    Call Print_Setting(" 부적격 혈액 이송", lngX + CLng(Printer.TextWidth("제    목:")), lngY, lngY, "L")
    Printer.FontSize = 10: Printer.FontBold = False
    Printer.Line (lngX, lngCurYPos)-(192, lngCurYPos)
    
    lngCurYPos = lngCurYPos + lngY
    
    Call Print_Setting("  혈액관리법 제 8조 제 2항 및 동법시행규칙 제 10조 제 3호의 규정에 의하여 다음과 같이 부적격혈액", lngX, lngY, lngY, "L")
    Call Print_Setting("을 이송합니다.", lngX, lngY, lngY, "L")
    Call Print_Setting("", lngX * 2, lngY)
    

    
    Printer.Line (6, 130)-(190, 266), , B
    lngLineY = 130
    For ii = 1 To 16
        lngLineY = lngLineY + 8
        If ii = 1 Then
            Printer.Line (40, lngLineY)-(130, lngLineY)
        Else
            Printer.Line (lngX, lngLineY)-(190, lngLineY)
        End If
    Next
    lngCurYPos = 130
    Printer.Line (13, 130)-(13, 266)
    Printer.Line (40, 130)-(40, 266)
    Printer.Line (55, 138)-(55, 266)
    Printer.Line (70, 138)-(70, 266)
    Printer.Line (85, 138)-(85, 266)
    Printer.Line (100, 138)-(100, 266)
    Printer.Line (115, 138)-(115, 266)
    Printer.Line (130, 130)-(130, 266)
    Printer.Line (150, 130)-(150, 266)
    Printer.Line (170, 130)-(170, 266)
    Call Print_Setting("일련", lngX, lngY, 7, , , False)
    Call Print_Setting("혈    액", 13, lngY, 27, , , False)
    Call Print_Setting("부 적 격 사 유", 40, lngY, 90, , , False)
    Call Print_Setting("검사일", 130, lngY * 2, 20, , , False)
    Call Print_Setting("혈   액", 150, lngY, 20, , , False)
    Call Print_Setting("비    고", 170, lngY * 2, 20, , , False)
    Call Print_Setting("", lngX, lngY)
    Call Print_Setting("번호", lngX, lngY, 7, , , False)
    Call Print_Setting("번    호", 13, lngY, 27, , , False)
    Call Print_Setting("ALT", 40, lngY, 15, , , False)
    Call Print_Setting("B형감염", 55, lngY, 15, , , False)
    Call Print_Setting("C형감염", 70, lngY, 15, , , False)
    Call Print_Setting("매  독", 85, lngY, 15, , , False)
    Call Print_Setting("에이즈", 100, lngY, 15, , , False)
    Call Print_Setting("기  타", 115, lngY, 15, , , False)
    Call Print_Setting("제제명", 150, lngY, 20)
End Sub

