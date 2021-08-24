VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frmBBS921 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   0  '없음
   Caption         =   "헌혈실적통지서"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10995
   Icon            =   "frmBBS921.frx":0000
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
      Caption         =   "헌혈 실적 통지서"
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   825
      Left            =   75
      TabIndex        =   4
      Top             =   285
      Width           =   10770
      Begin VB.CommandButton cmdQuery 
         BackColor       =   &H00F4F0F2&
         Caption         =   "조회(&Q)"
         Height          =   510
         Left            =   9300
         Style           =   1  '그래픽
         TabIndex        =   5
         Tag             =   "124"
         Top             =   225
         Width           =   1320
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
         Left            =   1215
         TabIndex        =   6
         Top             =   285
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM"
         Format          =   64094211
         CurrentDate     =   36799
      End
      Begin VB.Label Label5 
         Alignment       =   1  '오른쪽 맞춤
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "조회기간"
         Height          =   180
         Index           =   1
         Left            =   195
         TabIndex        =   7
         Top             =   345
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   9510
      Style           =   1  '그래픽
      TabIndex        =   2
      Tag             =   "128"
      Top             =   8400
      Width           =   1320
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00F4F0F2&
      Caption         =   "출력(&P)"
      Height          =   510
      Left            =   8175
      Style           =   1  '그래픽
      TabIndex        =   1
      Tag             =   "15101"
      Top             =   8400
      Width           =   1320
   End
   Begin FPSpread.vaSpread tblList 
      Height          =   3405
      Left            =   1485
      TabIndex        =   0
      Top             =   2715
      Width           =   8430
      _Version        =   196608
      _ExtentX        =   14870
      _ExtentY        =   6006
      _StockProps     =   64
      AutoSize        =   -1  'True
      BackColorStyle  =   1
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
      GridShowVert    =   0   'False
      MaxCols         =   8
      MaxRows         =   8
      ScrollBars      =   0
      SelectBlockOptions=   0
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      ShadowText      =   0
      SpreadDesigner  =   "frmBBS921.frx":076A
      TextTip         =   4
   End
   Begin Crystal.CrystalReport CReport 
      Left            =   1260
      Top             =   6300
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DBE6E6&
      Height          =   7245
      Left            =   75
      TabIndex        =   8
      Top             =   1020
      Width           =   10770
   End
End
Attribute VB_Name = "frmBBS921"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum tblColumn
    tcDONORCNT = 2
    tcTOTAL
    tc32Occ
    tc400cc
    tcPLASMA
    tcPLATELET
    tcLEUKOCYTE
End Enum
Dim Papercnt  As Long               '헌혈증서재고량
    

Private Sub cmdExit_Click()
    Unload Me
End Sub


Private Sub cmdQuery_Click()
    Dim objProBar As New clsProgress
    Dim objstatic As New clsStatics
    
    Dim Fdt       As String
    Dim Tdt       As String
    Dim ii        As Integer
    Dim jj        As Integer
    Dim tot       As Long
    
    Fdt = Format(dtpMonth.Value, "yyyymm") & "01"
    Tdt = Format(dtpMonth.Value, "yyyymm") & "31"
    cmdPrint.Enabled = True
    Table_Clear
    
'    Set objProBar.StatusBar = MainFrm.stsbar
    objProBar.Container = MainFrm.stsbar
    objProBar.Max = 100
    
    For ii = 1 To 30
        objProBar.Value = ii
    Next
    '-----------------
    '헌혈지원자 구하기
    '-----------------
    Call Donor_Support(Fdt, Tdt)            '헌혈지원자 구하기
    
    For ii = 31 To 60
        objProBar.Value = ii
    Next
    '-------------
    '전혈자 구하기
    '-------------
    Call Whole_BloodCnt(Fdt, Tdt)           '전혈자 구하기
    '-----------------
    
    '성분헌혈자 구하기
    '-----------------
    Call Pheresis_BloodCnt(Fdt, Tdt)        '성분헌혈자 구하기
    
    For ii = 61 To 100
        objProBar.Value = ii
    Next
    
'    objstatic.setDbConn DBConn
    Papercnt = objstatic.DonorPaperCnt
    
    With tblList
        .Row = 8
        .Col = tblColumn.tcTOTAL: .Value = Format(GetSystemDate, "yyyy") & "년"
        .Col = tblColumn.tc32Occ: .Value = Format(GetSystemDate, "mm") & " 월" & Format(GetSystemDate, "dd") & " 일"
        .Col = tblColumn.tcPLASMA: .Value = "총 " & Papercnt & "매"
        For ii = 4 To 7
            .Row = ii
            For jj = tblColumn.tc32Occ To tblColumn.tcLEUKOCYTE
                .Col = jj
                tot = tot + Val(.Value)
            Next
            .Col = tblColumn.tcTOTAL: .Value = tot
            tot = 0
        Next
    End With
    
   
    Set objstatic = Nothing
    Set objProBar = Nothing
End Sub
Private Sub Pheresis_BloodCnt(ByVal Fdt As String, ByVal Tdt As String)
'----------------------
'전혈자
'----------------------
    Dim objstatic As New clsStatics
    Dim PRs       As New Recordset
    Dim RS        As New Recordset
    Dim lngFind   As Long
    Dim lngFind1  As Long
    Dim lngFind2  As Long
    Dim MPlasma   As Long
    Dim MPlete    As Long
    Dim MLeuk     As Long
    Dim FPlasma   As Long
    Dim FPlete    As Long
    Dim FLeuk     As Long
    Dim MOPlasma  As Long
    Dim MOPlete   As Long
    Dim MOLeuk    As Long
    Dim YPlasma   As Long
    Dim YPlete    As Long
    Dim YLeuk     As Long
    Dim strTmp    As String
    
'    objstatic.setDbConn DBConn
    Set RS = objstatic.PheresisCnt(Fdt, Tdt)
    Set PRs = objstatic.PheresisList
    If Not RS.EOF Then
        Do Until RS.EOF
            PRs.MoveFirst
            Do Until PRs.EOF
                strTmp = PRs.Fields("field1").Value & "" & COL_DIV & PRs.Fields("field2").Value & ""
                lngFind = InStr(strTmp, RS.Fields("compocd").Value & "")
                If lngFind > 0 Then
                    Select Case medGetP(strTmp, 1, COL_DIV)
                        Case "혈  장"
                            Select Case RS.Fields("div").Value & ""
                                Case "sexdiv"
                                    If RS.Fields("sex").Value & "" = "M" Then
                                        MPlasma = MPlasma + Val(RS.Fields("cnt").Value & "")
                                    Else
                                        FPlasma = FPlasma + Val(RS.Fields("cnt").Value & "")
                                    End If
                                Case "M": MOPlasma = MOPlasma + Val(RS.Fields("cnt").Value & "")
                                Case "Y": YPlasma = YPlasma + Val(RS.Fields("cnt").Value & "")
                            End Select
                        Case "혈소판"
                            Select Case RS.Fields("div").Value & ""
                                Case "sexdiv"
                                    If RS.Fields("sex").Value & "" = "M" Then
                                        MPlete = MPlete + Val(RS.Fields("cnt").Value & "")
                                    Else
                                        FPlete = FPlete + Val(RS.Fields("cnt").Value & "")
                                    End If
                                Case "M": MOPlete = MOPlete + Val(RS.Fields("cnt") & "")
                                Case "Y": YPlete = YPlete + Val(RS.Fields("cnt").Value & "")
                            End Select
                        Case "백혈구"
                            Select Case RS.Fields("div").Value & ""
                                Case "sexdiv"
                                    If RS.Fields("sex").Value & "" = "M" Then
                                        MLeuk = MLeuk + Val(RS.Fields("cnt").Value & "")
                                    Else
                                        FLeuk = FLeuk + Val(RS.Fields("cnt").Value & "")
                                    End If
                                Case "M": MOLeuk = MOLeuk + Val(RS.Fields("cnt").Value & "")
                                Case "Y": YLeuk = YLeuk + Val(RS.Fields("cnt").Value & "")
                            End Select
                    End Select
                End If
                PRs.MoveNext
            Loop
            RS.MoveNext
        Loop
        YPlasma = YPlasma + MOPlasma
        YPlete = YPlete + MOPlete
        YLeuk = YLeuk + MOLeuk
    End If
    With tblList
        .Row = 4: .Col = tblColumn.tcPLASMA:    .Value = MPlasma
                  .Col = tblColumn.tcPLATELET:  .Value = MPlete
                  .Col = tblColumn.tcLEUKOCYTE: .Value = MLeuk
        .Row = 5: .Col = tblColumn.tcPLASMA:    .Value = FPlasma
                  .Col = tblColumn.tcPLATELET:  .Value = FPlete
                  .Col = tblColumn.tcLEUKOCYTE: .Value = FLeuk
        .Row = 6: .Col = tblColumn.tcPLASMA:    .Value = MOPlasma
                  .Col = tblColumn.tcPLATELET:  .Value = MOPlete
                  .Col = tblColumn.tcLEUKOCYTE: .Value = MOLeuk
        .Row = 7: .Col = tblColumn.tcPLASMA:    .Value = YPlasma
                  .Col = tblColumn.tcPLATELET:  .Value = YPlete
                  .Col = tblColumn.tcLEUKOCYTE: .Value = YLeuk
    End With
    
    Set RS = Nothing
    Set PRs = Nothing
    Set objstatic = Nothing
End Sub

Private Sub Whole_BloodCnt(ByVal Fdt As String, ByVal Tdt As String)
'----------------------
'전혈자
'----------------------
    Dim objstatic As New clsStatics
    Dim RS       As New Recordset
    Dim MCnt320   As Long
    Dim MCnt400   As Long
    Dim Fcnt320      As Long
    Dim Fcnt400      As Long
    Dim MonthCnt320  As Long
    Dim MonthCnt400  As Long
    Dim YearCnt320   As Long
    Dim YearCnt400   As Long
    
'    objstatic.setDbConn DBConn
    Set RS = objstatic.WholeBloodCnt(Fdt, Tdt)
    If RS.RecordCount > 0 Then
        Do Until RS.EOF
            Select Case RS.Fields("div").Value & ""
                Case "sexdiv"
                    If RS.Fields("sex").Value & "" = "M" Then
                        If RS.Fields("volumn").Value & "" = "320" Then
                            MCnt320 = Val(RS.Fields("cnt").Value & "")
                        ElseIf RS.Fields("volumn").Value & "" = "400" Then
                            MCnt400 = Val(RS.Fields("cnt").Value & "")
                        End If
                    ElseIf RS.Fields("sex").Value & "" = "F" Then
                        If RS.Fields("volumn").Value & "" = "320" Then
                            Fcnt320 = Val(RS.Fields("cnt").Value & "")
                        ElseIf RS.Fields("volumn").Value & "" = "400" Then
                            Fcnt400 = Val(RS.Fields("cnt").Value & "")
                        End If
                    End If
                Case "M":
                    If RS.Fields("volumn").Value & "" = "320" Then
                        MonthCnt320 = Val(RS.Fields("cnt").Value & "")
                    ElseIf RS.Fields("volumn").Value & "" = "400" Then
                        MonthCnt400 = Val(RS.Fields("cnt").Value & "")
                    End If
                Case "Y":
                    If RS.Fields("volumn").Value & "" = "320" Then
                        YearCnt320 = Val(RS.Fields("cnt").Value & "")
                    ElseIf RS.Fields("volumn").Value & "" = "400" Then
                        YearCnt400 = Val(RS.Fields("cnt").Value & "")
                    End If
            End Select
            RS.MoveNext
        Loop
        YearCnt320 = YearCnt320 + MonthCnt320
        YearCnt400 = YearCnt400 + MonthCnt400
    End If
    With tblList
        .Row = 4: .Col = tblColumn.tc32Occ: .Value = MCnt320
                  .Col = tblColumn.tc400cc: .Value = MCnt400
        .Row = 5: .Col = tblColumn.tc32Occ: .Value = Fcnt320
                  .Col = tblColumn.tc400cc: .Value = Fcnt400
        .Row = 6: .Col = tblColumn.tc32Occ: .Value = MonthCnt320
                  .Col = tblColumn.tc400cc: .Value = MonthCnt400
        .Row = 7: .Col = tblColumn.tc32Occ: .Value = YearCnt320
                  .Col = tblColumn.tc400cc: .Value = YearCnt400
    End With
    Set RS = Nothing
    Set objstatic = Nothing
    
End Sub
Private Sub Donor_Support(ByVal Fdt As String, ByVal Tdt As String)
'----------------------
'헌혈지원자
'----------------------
    Dim objstatic As New clsStatics
    Dim RS       As New Recordset
    Dim MCnt      As Long
    Dim Fcnt      As Long
    Dim MonthCnt  As Long
    Dim YearCnt   As Long
    
'    objstatic.setDbConn DBConn
    Set RS = objstatic.Get_DonorRecord(Fdt, Tdt)
    If RS.RecordCount > 0 Then
        Do Until RS.EOF
            Select Case RS.Fields("div").Value & ""
                Case "sexdiv"
                    If RS.Fields("sex").Value & "" = "M" Then
                        MCnt = Val(RS.Fields("cnt").Value & "")
                    ElseIf RS.Fields("sex").Value & "" = "F" Then
                        Fcnt = Val(RS.Fields("cnt").Value & "")
                    End If
                Case "M": MonthCnt = Val(RS.Fields("cnt").Value & "")
                Case "Y": YearCnt = Val(RS.Fields("cnt").Value & "") + MonthCnt
            End Select
            RS.MoveNext
        Loop
    End If
    With tblList
        .Row = 4: .Col = tblColumn.tcDONORCNT: .Value = MCnt
        .Row = 5: .Col = tblColumn.tcDONORCNT: .Value = Fcnt
        .Row = 6: .Col = tblColumn.tcDONORCNT: .Value = MonthCnt
        .Row = 7: .Col = tblColumn.tcDONORCNT: .Value = YearCnt
    End With
    
    Set RS = Nothing
    Set objstatic = Nothing
End Sub

Private Sub Form_Load()
    dtpMonth.Value = Format(GetSystemDate, "yyyy-mm")
    cmdPrint.Enabled = False
End Sub
Private Sub Table_Clear()
    Dim ii As Integer
    Dim jj As Integer
    
    With tblList
        For ii = 4 To 7
            .Row = ii
            For jj = tblColumn.tcDONORCNT To tblColumn.tcLEUKOCYTE
                .Col = jj
                .Value = ""
            Next
        Next
    End With
End Sub


Private Sub cmdPrint_Click()
'--------
'출력
'--------
    Call VbPrint

End Sub

Private Sub VbPrint()
    Dim lngY        As Long
    Dim lngX        As Long
    Dim lngLineY    As Long
    Dim ii          As Integer
    Dim jj          As Long
    Dim strPage As String
    Dim lngFirst As Long
    
    Call P_PrtSet
    
    lngY = 8
    lngX = 12
    lngCurYPos = 20
    
    Printer.DrawWidth = 4
    Printer.DrawStyle = 0
    
    Printer.FontSize = 20: Printer.FontBold = True
    Call Print_Setting(" 헌 혈 실 적 통 지 서 ", 0, lngY * 2, Printer.ScaleWidth)
    Printer.Line (60, lngCurYPos - 3)-(139, lngCurYPos - 3)
    
    Printer.FontSize = 10: Printer.FontBold = False
    
    Call Print_Setting("", lngX, lngY)
    Call Print_Setting("혈액원 명칭 : " & HOSPITAL_MAIN, lngX, lngY, , "L")
    
    Call Print_Setting("혈액원 주소 : " & HOSPITAL_ADDR & "  전화번호: " & HOSPITAL_TEL1, lngX, lngY, , "L")
    
    Call Print_Setting("", lngX, lngY * 2)
    
    Call Print_Setting(Format(Now, "yyyy년 mm월분"), lngX, lngY - 3, lngX, "L")
    Printer.FontSize = 7
    Call Print_Setting("(단위 : 명)", 160, lngY - 3, , "L")
    Printer.FontSize = 10
    lngCurYPos = 88
    '86
    lngLineY = lngCurYPos
    Printer.Line (lngX, lngCurYPos)-(180, 250), , B
    For ii = 1 To 8
        lngLineY = lngLineY + 8
        Select Case ii
            Case 1
                Printer.Line (60, lngLineY)-(180, lngLineY)
            Case 2
                Printer.Line (80, lngLineY)-(180, lngLineY)
            Case 8
                lngLineY = lngLineY + 2
                Printer.Line (12, lngLineY)-(180, lngLineY)
            Case Else
                Printer.Line (12, lngLineY)-(180, lngLineY)
        End Select
    Next
    Printer.Line (35, 88)-(35, 154)
    Printer.Line (60, 88)-(60, 144)
    Printer.Line (80, 96)-(80, 144)
    Printer.Line (100, 104)-(100, 144)
    
    Printer.Line (120, 96)-(120, 144)
    Printer.Line (140, 104)-(140, 144)
    Printer.Line (160, 104)-(160, 144)
    
    With tblList
        Call Print_Setting("구  분", lngX, lngY * 3, 23, , , False)
        Call Print_Setting("헌혈지원자", 35, lngY * 3, 25, , , False)
        Call Print_Setting("헌   혈   자", 60, lngY, 120)
        Call Print_Setting("계", 60, lngY * 2, 20, , , False)
        Call Print_Setting("전혈", 80, lngY, 40, , , False)
        Call Print_Setting("성분헌혈", 120, lngY, 60)
        Call Print_Setting("320㎖", 80, lngY, 20, , , False)
        Call Print_Setting("400㎖", 100, lngY, 20, , , False)
        Call Print_Setting("혈장", 120, lngY, 20, , , False)
        Call Print_Setting("혈소판", 140, lngY, 20, , , False)
        Call Print_Setting("백혈구", 160, lngY, 20)
        For ii = 4 To 7
            .Row = ii
            For jj = 1 To 8
                .Col = jj
                Select Case jj
                    Case 1: Call Print_Setting(.Value, lngX, lngY, 23, , , False)
                    Case 2: Call Print_Setting(.Value, 35, lngY, 25, , , False)
                    Case 3: Call Print_Setting(.Value, 60, lngY, 20, , , False)
                    Case 4: Call Print_Setting(.Value, 80, lngY, 20, , , False)
                    Case 5: Call Print_Setting(.Value, 100, lngY, 20, , , False)
                    Case 6: Call Print_Setting(.Value, 120, lngY, 20, , , False)
                    Case 7: Call Print_Setting(.Value, 140, lngY, 20, , , False)
                    Case 8: Call Print_Setting(.Value, 160, lngY, 20, , , False)
                End Select
            Next
            Call Print_Setting("", lngX, lngY)
        Next
        Call Print_Setting("헌혈증서", lngX, lngY - 4, 20)
        Call Print_Setting("재고량", lngX, lngY - 4, 20, , , False)
        Call Print_Setting("", lngX, -4)
                    
        .Col = 6
        .Row = 8
        strPage = .Text
        lngFirst = InStr(strPage, "총")
        If lngFirst > 0 Then
            strPage = Mid(strPage, lngFirst + 1)
            strPage = Replace(strPage, "매", "")
        End If
                            
        Call Print_Setting(Space(5) & Format(Now, "yyyy 년 mm 월 dd 일") & _
                           "    현재     총 " & strPage & "  매", 35, lngY, 145)
        Call Print_Setting("", lngX, lngY * 3)
        
        Call Print_Setting("위와 같이 헌혈실적을 통지합니다.", lngX, lngY, 168)
        Call Print_Setting("", lngX, lngY * 2)
        Call Print_Setting(Format(Now, "yyyy 년 mm 월 dd 일"), 100, lngY, , "L")
        Call Print_Setting("발 신 인 : " & HOSPITAL_MAIN & " (인)", 100, lngY, , "L")
        Call Print_Setting("", lngX, lngY * 3)
        
        Printer.FontSize = 15: Printer.FontBold = True
        Call Print_Setting("대한적십자사 총재 귀하", lngX, lngY, 168)
        
    End With
    Printer.EndDoc
End Sub
