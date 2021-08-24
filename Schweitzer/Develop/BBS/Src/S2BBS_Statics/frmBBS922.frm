VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frmBBS922 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   0  '없음
   Caption         =   "헌혈경력조회서"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10995
   Icon            =   "frmBBS922.frx":0000
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
      Caption         =   "헌혈 경력 조회서"
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
         Left            =   9285
         Style           =   1  '그래픽
         TabIndex        =   6
         Tag             =   "124"
         Top             =   225
         Width           =   1320
      End
      Begin MSComCtl2.DTPicker dtpTo 
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
         Left            =   2760
         TabIndex        =   5
         Top             =   300
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   77463555
         CurrentDate     =   36799
      End
      Begin MSComCtl2.DTPicker dtpFrom 
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
         Left            =   1200
         TabIndex        =   7
         Top             =   300
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   77463555
         CurrentDate     =   36799
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   285
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
         Caption         =   "조회 일자"
         Appearance      =   0
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "~"
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   1
         Left            =   2580
         TabIndex        =   8
         Tag             =   "40304"
         Top             =   360
         Visible         =   0   'False
         Width           =   90
      End
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   480
      Left            =   9510
      Style           =   1  '그래픽
      TabIndex        =   1
      Tag             =   "128"
      Top             =   8400
      Width           =   1320
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00F4F0F2&
      Caption         =   "출력(&P)"
      Height          =   480
      Left            =   8175
      Style           =   1  '그래픽
      TabIndex        =   0
      Tag             =   "15101"
      Top             =   8400
      Width           =   1320
   End
   Begin Crystal.CrystalReport CReport 
      Left            =   405
      Top             =   6120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin FPSpread.vaSpread tblList 
      Height          =   7155
      Left            =   75
      TabIndex        =   2
      Top             =   1125
      Width           =   10770
      _Version        =   196608
      _ExtentX        =   18997
      _ExtentY        =   12621
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
      MaxCols         =   9
      MaxRows         =   30
      OperationMode   =   2
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      ShadowText      =   0
      SpreadDesigner  =   "frmBBS922.frx":076A
      TextTip         =   4
   End
End
Attribute VB_Name = "frmBBS922"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum tblColumn
    tcChk = 1
    tcName
    tcSSN
    TcAddress
    tcABO
    TcCount
    tcOK
    tcNOT
End Enum
    
Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub TableClear()
    With tblList
        .MaxRows = 2
        .MaxRows = 30
    End With
End Sub

Private Sub cmdPrint_Click()

    Dim blnPrint As Boolean
    Dim Name As String
    Dim SSN  As String
    Dim ABO  As String
    Dim ADD  As String
    Dim sOK  As String
    Dim sNot As String
    
    Dim ii   As Long
    Call P_PrtSet
    With tblList
        For ii = 3 To .MaxRows
            .Row = ii
            .Col = tblColumn.tcChk
            If .Value = 0 Then
                If blnPrint = True Then Printer.NewPage
                .Col = tblColumn.tcName: Name = .Value
                .Col = tblColumn.tcSSN: SSN = .Value
                .Col = tblColumn.tcABO: ABO = .Value
                .Col = tblColumn.TcAddress: ADD = .Value
                .Col = tblColumn.tcOK: sOK = .Value
                .Col = tblColumn.tcNOT: sNot = .Value
                
                Call ReportPrint(Name, SSN, ADD, ABO, sOK, sNot)
                blnPrint = True
            End If
        Next
        If blnPrint Then Printer.EndDoc
    End With
            
        
End Sub

Private Sub Form_Load()
    dtpFrom = Format(GetSystemDate, "yyyy-MM")
    dtpTo = Format(GetSystemDate, "yyyy-MM-dd")
    tblList.MaxRows = 2
    cmdPrint.Enabled = False
End Sub


Private Sub cmdQuery_Click()
    Dim objDonor As New clsStatics
    Dim objPro   As New clsProgress
    Dim RS       As New Recordset
    Dim strJudge As String
    Dim Fdt      As String
    Dim Tdt      As String
    Dim ii       As Integer
    
    Fdt = Format(dtpFrom.Value, "YYYYMMdd")
    Tdt = Format(dtpTo.Value, "YYYYMMdd")

    Set RS = objDonor.Get_DonorHistory(Fdt, Tdt)
    
    If RS.RecordCount > 0 Then
'        Set objPro.StatusBar = MainFrm.stsbar
        objPro.Container = MainFrm.stsbar
        objPro.Max = RS.RecordCount
        
        RS.MoveFirst
        With tblList
            .MaxRows = RS.RecordCount + 2
            ii = 3
            Do Until RS.EOF
                .Row = ii
                .Col = tblColumn.tcChk:     .Value = 0
                .Col = tblColumn.tcName:    .Value = "":  .Value = RS.Fields("donornm").Value & ""
                .Col = tblColumn.tcSSN:     .Value = "":   .Value = Mid(RS.Fields("ssn").Value & "", 1, 6) & "-" & Mid(RS.Fields("ssn").Value & "", 7)
                .Col = tblColumn.TcAddress: .Value = "": .Value = RS.Fields("addr1").Value & "" & " " & vbTab & RS.Fields("addr2").Value & ""
                .Col = tblColumn.tcABO:     .Value = "": .Value = RS.Fields("abo").Value & "" & RS.Fields("rh").Value & ""
                .Col = tblColumn.TcCount:   .Value = "": .Value = Val(RS.Fields("cnt").Value & "")
                strJudge = RS.Fields("okdiv3").Value & ""
                If strJudge <> "" Then
                    Select Case strJudge
                        Case "1"
                            .Col = tblColumn.tcOK: .Value = "○": .TypeHAlign = TypeHAlignCenter
                        Case "0"
                            .Col = tblColumn.tcNOT: .Value = "○": .TypeHAlign = TypeHAlignCenter
                        End Select
                        
                End If
                    
                
                ii = ii + 1
                objPro.Value = ii - 3
                RS.MoveNext
            Loop
        End With
        Set objPro = Nothing
        cmdPrint.Enabled = True
    Else
        tblList.MaxRows = 2
        cmdPrint.Enabled = False
    End If
End Sub

Private Sub ReportPrint(ByVal nm As String, ByVal SSN As String, ByVal ADD As String, _
                        ByVal ABO As String, ByVal sOK As String, ByVal sNot As String)
    Dim lngY As Long
    lngY = 8
    
    
    
    lngCurYPos = 10
    
    Printer.DrawWidth = 4
    Printer.Line (6, 10)-(190, 274), , B
    Printer.Line (6, 40)-(190, 40)
    Printer.Line (6, 48)-(190, 48)
    Printer.Line (6, 56)-(190, 56)
    Printer.Line (6, 64)-(190, 64)
    Printer.Line (6, 72)-(190, 72)
    Printer.Line (6, 80)-(190, 80)
    Printer.Line (36, 48)-(36, 80)
    Printer.Line (100, 40)-(100, 80)
    Printer.Line (130, 48)-(130, 80)
    Printer.Line (162, 10)-(190, 22), , B
    Printer.Line (162, 16)-(190, 16)
    Printer.Line (6, 144)-(190, 144)
    Printer.Line (6, 152)-(190, 152)
    Printer.Line (100, 160)-(190, 160)
    Printer.Line (6, 168)-(190, 168)
    Printer.Line (6, 176)-(190, 176)
    Printer.Line (130, 160)-(130, 176)
    Printer.Line (162, 184)-(190, 184)
    Printer.Line (6, 190)-(100, 190)
    Printer.Line (6, 200)-(190, 200)
    Printer.Line (162, 176)-(162, 200)
    Printer.Line (100, 152)-(100, 200)
    
    
'    Call Print_Setting(" 헌 혈 경 력 조 회 서", 10, 30, 186, "C", "C", False)
    
    
    Call Print_Setting("처리기간", 166, 6, 22, "C", "C")
    Call Print_Setting("즉    시", 166, 6, 22, "C", "C")
    Printer.FontBold = True: Printer.FontSize = 20
    Call Print_Setting(" 헌 혈 경 력 조 회 서", 10, 18, 186, "C", "C")
    Printer.FontBold = False: Printer.FontSize = 10
    Call Print_Setting(" 신   청   인 ", 6, lngY, 94, "C", "C", False)
    Call Print_Setting(" 헌 혈 자 인 적 사 항 ", 100, 8, 92, "C", "C")
    Call Print_Setting(" 명      칭 ", 6, lngY, "30", "C", "C", False)
    
    Call Print_Setting(HOSPITAL_MAIN, 30, lngY, 70, "C", "C", False)
    
    Call Print_Setting(" 성        명 ", 100, lngY, 30, "C", "C", False)
    
    Call Print_Setting(nm, 130, lngY, 62)
    
    
    Call Print_Setting(" 대표자이름", 6, lngY, 30, "C", "C", False)
    Call Print_Setting(" 주민등록번호", 100, lngY, 30, "C", "C", False)
    Call Print_Setting(SSN, 130, lngY, 62)
    
    
    Call Print_Setting(" 주      소", 6, lngY, 30, "C", "C", False)
    
    Call Print_Setting("인천 남동구 구월동 1198", 36, 8, 64, "C", "C", False)
    
    Call Print_Setting(" 주        소", 100, lngY, 30, "C", "C", False)
    
    
'    Call Print_Setting(ADD, 130, lngY, 62)
    If medGetP(ADD, 2, vbTab) <> "" Then
        Printer.FontSize = 8
        Call Print_Setting(medGetP(ADD, 1, vbTab), 130, lngY - 4, 62)
        Call Print_Setting(medGetP(ADD, 2, vbTab), 130, lngY - 4, 62, , , False)
        Call Print_Setting("", 130, 4)
        Printer.FontSize = 10
    Else
        Call Print_Setting(ADD, 130, lngY, 62)
    End If
    
    Call Print_Setting(" 연  락  처", 6, lngY, 30, "C", "C", False)
    
    Call Print_Setting(HOSPITAL_TEL2, 36, lngY, 64, "C", "C", False)
    
    Call Print_Setting(" 혈   액   형", 100, lngY, 30, "C", "C", False)
    Call Print_Setting(ABO, 130, lngY, 62)
    
    Call Print_Setting("", 6, lngY * 2)
    Call Print_Setting("혈액관리법 제8조제1항 및 동법시행규칙 제9조제2항의 규정에 의하여", 6, lngY, 186)
    Call Print_Setting("위 헌혈자의 경력조회서를 신청합니다.", 6, lngY, 186)
    Call Print_Setting("", 6, lngY * 2)
    
    Call Print_Setting(Space(50) & Format(Now, "yyyy 년 mm 월  dd 일"), 6, 6, 180)
    Call Print_Setting("인천광역시적십자혈액원장 귀하", 10, lngY, 50, "L", "C", False)
    Call Print_Setting(Space(20) & " 신청인", 6, lngY, 180, , , False)
    Printer.FontSize = 6
    Call Print_Setting(Space(120) & "(서명 또는 인)", 6, lngY, 180)
    Printer.FontSize = 10
    
    Call Print_Setting("", 6, 2)
    Call Print_Setting(" 조     회     결     과", 6, lngY, 180)
    Call Print_Setting(" 과거 헌혈경력이 있는 경우", 6, lngY, 94, , , False)
    Call Print_Setting("검사결과", 100, lngY, 92)
    Call Print_Setting("(아래칸에 헌혈일자 기재)", 6, lngY, 94, , , False)
    Call Print_Setting("적      격", 100, lngY, 30, , , False)
    Call Print_Setting("부   적   격", 130, lngY, 62)
    
    Call Print_Setting("", 100, lngY, 30, , , False)
    Call Print_Setting("", 130, lngY, 62)
    
    
    'Call Print_Setting("", 6, lngY)
    Call Print_Setting(" 과거 헌혈경력이 있는 경우", 6, lngY, 94, , , False)
    Call Print_Setting("수수료", 162, lngY, 30)
    Call Print_Setting("(아래칸에 없음 기재)", 6, lngY, 94, , , False)
    
    Call Print_Setting("없음", 162, lngY * 2, 30)
    
    Call Print_Setting("", 6, lngY * 2)
    
    Call Print_Setting("혈액관리법 시행규칙 제9조제3항의 규정에 의하여", 6, lngY, 180)
    Call Print_Setting("위 헌혈자의 헌혈경력조회를 통보합니다.", 6, lngY, 180)
    Call Print_Setting("", 6, lngY * 1)
    Call Print_Setting(Space(50) & Format(Now, "yyyy 년 mm 월  dd 일"), 6, lngY + 2, 180)
    Call Print_Setting(Space(10) & "인천광역적십자혈액원장 " & Space(30) & "(인)", 6, lngY, 180)
'    Call Print_Setting("", 6, lngY)
    
    Call Print_Setting("귀하", 6, lngY, 50)
End Sub


