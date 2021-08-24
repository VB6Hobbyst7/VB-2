VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmWorkList 
   Caption         =   "LASC Order"
   ClientHeight    =   8610
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9810
   Icon            =   "frmWorkList.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   9810
   StartUpPosition =   1  '소유자 가운데
   Begin MSComCtl2.DTPicker dtpSDate 
      Height          =   315
      Left            =   1410
      TabIndex        =   26
      Top             =   60
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   99745793
      CurrentDate     =   40205
   End
   Begin HST_WorkList_국립암센터.MDButton cmdClear 
      Height          =   525
      Left            =   7170
      TabIndex        =   23
      Top             =   7980
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   926
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Clear"
   End
   Begin HST_WorkList_국립암센터.MDButton cmdClose 
      Height          =   525
      Left            =   8430
      TabIndex        =   24
      Top             =   7980
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   926
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "닫기"
   End
   Begin HST_WorkList_국립암센터.MDButton cmdWorkList 
      Height          =   525
      Left            =   5790
      TabIndex        =   22
      Top             =   7980
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   926
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "대기검체전송"
   End
   Begin HST_WorkList_국립암센터.MDButton CmdPortOpen 
      Height          =   525
      Left            =   4530
      TabIndex        =   21
      Top             =   7980
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   926
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "PortOpen"
   End
   Begin VB.TextBox txtMsg 
      Height          =   2475
      Left            =   8790
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   20
      Top             =   5310
      Visible         =   0   'False
      Width           =   4425
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   525
      Left            =   6180
      TabIndex        =   19
      Top             =   1680
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox txtTemp 
      Height          =   1695
      Left            =   7440
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   2655
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   8940
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      InputLen        =   1
      OutBufferSize   =   1024
      RThreshold      =   1
      SThreshold      =   1
   End
   Begin FPSpread.vaSpread vasTemp 
      Height          =   2115
      Left            =   6330
      TabIndex        =   14
      Top             =   2820
      Visible         =   0   'False
      Width           =   3495
      _Version        =   393216
      _ExtentX        =   6165
      _ExtentY        =   3731
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
      SpreadDesigner  =   "frmWorkList.frx":0442
   End
   Begin VB.CheckBox Check1 
      Caption         =   "SP"
      Height          =   285
      Left            =   11040
      TabIndex        =   13
      Top             =   1050
      Value           =   1  '확인
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   120
      TabIndex        =   11
      Top             =   8220
      Visible         =   0   'False
      Width           =   8775
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   8040
      Top             =   90
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7095
      Left            =   30
      TabIndex        =   5
      Top             =   450
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   12515
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "대기검체"
      TabPicture(0)   =   "frmWorkList.frx":0624
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "vasList"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdReLoad"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "오더 완료"
      TabPicture(1)   =   "frmWorkList.frx":0640
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "vasOrder"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdReLoad1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.CommandButton cmdReLoad1 
         Caption         =   "ReLoad"
         Height          =   255
         Left            =   -69900
         TabIndex        =   10
         Top             =   60
         Width           =   885
      End
      Begin VB.CommandButton cmdReLoad 
         Caption         =   "ReLoad"
         Height          =   255
         Left            =   2040
         TabIndex        =   9
         Top             =   60
         Visible         =   0   'False
         Width           =   885
      End
      Begin FPSpread.vaSpread vasList 
         Height          =   6615
         Left            =   90
         TabIndex        =   6
         Top             =   360
         Width           =   5955
         _Version        =   393216
         _ExtentX        =   10504
         _ExtentY        =   11668
         _StockProps     =   64
         ArrowsExitEditMode=   -1  'True
         ColHeaderDisplay=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   15
         ScrollBars      =   2
         SpreadDesigner  =   "frmWorkList.frx":065C
      End
      Begin FPSpread.vaSpread vasOrder 
         Height          =   7005
         Left            =   -74910
         TabIndex        =   15
         Top             =   375
         Width           =   5955
         _Version        =   393216
         _ExtentX        =   10504
         _ExtentY        =   12356
         _StockProps     =   64
         ColHeaderDisplay=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   10
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "frmWorkList.frx":20BC
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   8400
      Top             =   60
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   5000
   End
   Begin VB.CheckBox chkReal 
      Caption         =   "자동오더전송"
      Height          =   255
      Left            =   9750
      TabIndex        =   3
      Top             =   1050
      Value           =   1  '확인
      Visible         =   0   'False
      Width           =   1425
   End
   Begin FPSpread.vaSpread vasExam 
      Height          =   6285
      Left            =   6420
      TabIndex        =   2
      Top             =   1230
      Width           =   3225
      _Version        =   393216
      _ExtentX        =   5689
      _ExtentY        =   11086
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
      ScrollBars      =   2
      SpreadDesigner  =   "frmWorkList.frx":3ACD
   End
   Begin VB.CommandButton cmdListen 
      Caption         =   "처방 받기"
      Height          =   525
      Left            =   2190
      TabIndex        =   0
      Top             =   9090
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   " 수기 검체 등록 "
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   6420
      TabIndex        =   16
      Top             =   60
      Width           =   3225
      Begin VB.TextBox txtReOrd 
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1050
         TabIndex        =   17
         Top             =   270
         Width           =   2085
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검체번호"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   60
         TabIndex        =   18
         Top             =   360
         Width           =   900
      End
   End
   Begin MSComCtl2.DTPicker dtpEDate 
      Height          =   315
      Left            =   3180
      TabIndex        =   27
      Top             =   60
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   99745793
      CurrentDate     =   40205
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '투명
      Caption         =   "~"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2940
      TabIndex        =   28
      Top             =   120
      Width           =   225
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '투명
      Caption         =   "※ 조회일자"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   25
      Top             =   120
      Width           =   1245
   End
   Begin VB.Label lblWinsock 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "winsock 상태:"
      Height          =   180
      Left            =   9750
      TabIndex        =   12
      Top             =   3750
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Label lblBarCode 
      AutoSize        =   -1  'True
      Caption         =   "    "
      Height          =   180
      Left            =   7590
      TabIndex        =   8
      Top             =   915
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "검체번호"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   6480
      TabIndex        =   7
      Top             =   930
      Width           =   900
   End
   Begin VB.Label lblMsg 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "MSG"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   90
      TabIndex        =   4
      Top             =   7665
      Width           =   465
   End
End
Attribute VB_Name = "frmWorkList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'2010.01. 국립암센터
'타이머 interval 100 -> 1000

Dim lsTotRsv As String
Dim sckState As Integer
Dim MyPort As String
Dim gbOrdering As Boolean

Dim gsExamAll As String

Sub GetExamCode()
    gsExamAll = ""
    
    SQL = "Select examcode from equipexam where ordgubun in ('C','D','R','N','P') "
    gsExamAll = db_select_RowALL(gLocal, SQL)
    
End Sub

Private Sub chkReal_Click()
'    If chkReal.Value = 0 Then
'        If MsgBox("실시간으로 받은 처방을 장비로 전송할 수 없습니다" & vbCrLf & vbCrLf & "실시간 전송을 원하십니까? ", vbCritical + vbYesNo + vbDefaultButton1, "알림") = vbYes Then
'            chkReal.Value = 1
'        Else
'            chkReal.Value = 0
'        End If
'    End If
End Sub

Private Sub cmdClear_Click()
    ClearSpread vasExam
    ClearSpread vasList
    ClearSpread vasOrder
    
    SetBackColor vasOrder, 1, vasOrder.MaxRows, 1, vasOrder.MaxCols, 255, 255, 255
    
    lblBarCode = ""
    gRow = 1
End Sub

Private Sub cmdClose_Click()
    If MsgBox("프로그램을 종료하면 검사 처방 받기 및 장비로 처방 전송이 되지 않습니다" & vbCrLf & vbCrLf & "프로그램을 종료하시겠습니까?", vbCritical + vbYesNo + vbDefaultButton2, "종료알림") = vbNo Then
        Exit Sub
    End If
    
    Unload Me
End Sub

Private Sub cmdPortOpen_Click()
    If MSComm1.PortOpen = True Then
        If MsgBox("실시간으로 받은 처방을 장비로 전송할 수 없습니다" & vbCrLf & vbCrLf & "실시간 전송을 원하십니까? ", vbCritical + vbYesNo + vbDefaultButton1, "알림") = vbYes Then
            Exit Sub
        End If
        
        Timer1.Enabled = False
        MSComm1.PortOpen = False
        CmdPortOpen.Caption = "PortOpen"
        
        lblMsg.Caption = "[Message] 포트가 닫혀있습니다"
        'chkReal.Value = 0
'        Timer1.Enabled = True
    Else
        CmdPortOpen.Caption = "PortClose"
                
        LASCPortOpen
        
        Timer1.Interval = gTimer
        Timer1.Enabled = True
    End If
End Sub

Private Sub cmdReLoad_Click()
    ClearSpread vasList
    
    SQL = "Select barcode, OrdFlag, PID, PName, ReceDate,RemoteIP  from WorkList where OrdFlag = 'A' "
    res = db_select_Vas(gLocal, SQL, vasList)
    
    vasList.MaxRows = vasList.DataRowCnt
End Sub

Private Sub cmdReLoad1_Click()
    Dim lsReceDate  As String
    Dim lRow        As Long
    
    lsReceDate = Format(CDate(GetDateFull), "yyyymmdd")
    
    ClearSpread vasOrder
    
    SQL = "Select barcode, OrdFlag, PID, PName, ReceDate, ReceTime, wkno from WorkList where OrdFlag <> 'A' "
    res = db_select_Vas(gLocal, SQL, vasOrder)
    If res = -1 Then
        SaveQuery SQL
    End If
    
    For lRow = 1 To vasOrder.DataRowCnt
        Select Case Trim(GetText(vasOrder, lRow, 2))
        Case "C" '결과
            SetBackColor vasOrder, lRow, lRow, 1, 1, 255, 255, 185
        Case "D" '완료
            SetBackColor vasOrder, lRow, lRow, 1, 1, 170, 255, 170
        Case Else
            SetBackColor vasOrder, lRow, lRow, 1, 1, 255, 255, 255
        End Select
    Next lRow
    
    vasOrder.MaxRows = vasOrder.DataRowCnt
End Sub

Private Sub cmdWorkList_Click()

    If MSComm1.CTSHolding = False Then
        lblMsg.Caption = "[Message] LASC 의 포트가 준비되지 않았습니다"
        lblMsg.ForeColor = RGB(255, 0, 0)
    Else
        lblMsg.Caption = "[Message] LASC 의 포트가 준비되었습니다"
        lblMsg.ForeColor = RGB(0, 0, 0)
    End If
    

    Dim lsID, lsID1 As String
    Dim i, j, k As Integer
    Dim AdoRs_Exam As ADODB.Recordset
    Dim Ord(7)      As String
    Dim lsOrder     As String
    Dim lRow        As Long
    Dim lRow1       As Long
    Dim lsDate      As String
    
    Dim lsWkNo      As String
    Dim lsPID       As String
    Dim lsPName     As String
    
    Dim lsSlideOrd  As String
    Dim lsExamDate  As String
    Dim lsReceNo    As String
    Dim lsSlideName As String
    
'    If gbOrdering Then
'        SaveOrdLog "Ordering 중 Timer"
'        'Exit Sub
'    End If
    
'    If gbOrdering = True Then
'        SaveOrdLog "TIMER"
'    Else
'        SaveOrdLog "Ordering"
'        Exit Sub
'    End If
    
    SaveOrdLog "TIMER"
    
    If MSComm1.PortOpen = False Then
        LASCPortOpen
    End If
        
    If MSComm1.CTSHolding = False Then
        lblMsg.Caption = "[Message] LASC 의 포트가 준비되지 않았습니다"
        lblMsg.ForeColor = RGB(255, 0, 0)
        
        Exit Sub
    Else
        lblMsg.Caption = "[Message] LASC 의 포트가 준비되었습니다"
        lblMsg.ForeColor = RGB(0, 0, 0)
    End If
            
    DisConnect_Server
    
    If Connect_Server Then
        cn_Server_Flag = True
    End If
    
    ClearSpread vasList
    
    'res = Get_WorkList("20091001", Format(Date, "yyyymmdd"), gsExamAll, 2)
    
    res = Online_TLA(gXml_S05, dtpSDate.Value, dtpEDate.Value)
    
    lsID1 = ""
    
    With vasList
        For lRow1 = 0 To UBound(gTLA_Info_Select)
            lsID = Trim(gTLA_Info_Select(lRow1).SPCNO)
            
            res = Online_XML(gXml_S03, lsID)
            lsReceNo = Trim(gPat_Info_Select.ACPTNO_1)
            If lsID <> lsID1 Then
                SQL = "select barcode, OrdFlag from worklist where barcode = '" & lsID & "' and WkNo = '" & lsReceNo & "' "
                res = db_select_Col(gLocal, SQL)
                If Trim(gReadBuf(0)) <> "" Then
                    'txtMsg = txtMsg & Trim(GetText(vasList, lRow, 1)) & " : delete"
                    'DeleteRow vasList, lRow, lRow
                Else
                    SQL = "select barcode from pat_res where barcode = '" & lsID & "'and receno = '" & lsReceNo & "'  "
                    res = db_select_Col(gLocal, SQL)
                    If Trim(gReadBuf(0)) = lsID Then
                        'txtMsg = txtMsg & Trim(GetText(vasList, lRow, 1)) & " : delete"
                        'DeleteRow vasList, lRow, lRow
                    Else
                        lRow = .DataRowCnt + 1
                        
                        .SetText 1, lRow, lsID
                        .SetText 8, lRow, gPat_Info_Select.ACPTNO_1
                        
                        .SetText 3, lRow, gPat_Info_Select.PT_NO
                        .SetText 4, lRow, gPat_Info_Select.PT_NM
                        .SetText 5, lRow, gPat_Info_Select.ACPT_DTETM   '날짜
                        .SetText 6, lRow, ""    '시간
                        .SetText 7, lRow, "" '접수코드
                        .SetText 7, lRow, gPat_Info_Select.ACPTNO_1
                        .SetText 9, lRow, gPat_Info_Select.Sex
                        .SetText 10, lRow, gPat_Info_Select.Age
                        .SetText 11, lRow, ""   'slip
                    End If
                End If
            End If
            lsID1 = lsID
        Next lRow1
    End With
    
    If vasList.DataRowCnt > 0 Then
        Timer1.Enabled = False
'        gbOrdering = True
    Else
        Exit Sub
    End If
    
    If gSetup.Protocol = "B" Then
        OrderEntry_1 1
        Exit Sub
    Else
        For lRow = 1 To vasList.DataRowCnt
            OrderEntry_1 lRow
        Next lRow
    End If

    Timer1.Enabled = True
    
End Sub

Sub CopyRecord(ByVal asRow As Long)
    Dim llRow As Long
    Dim llCol As Long
    
    If asRow < 1 Or asRow > vasList.DataRowCnt Then Exit Sub
    
    llRow = vasOrder.DataRowCnt + 1
    If llRow > vasOrder.MaxRows Then
        vasOrder.MaxRows = llRow
    End If
    
    For llCol = 1 To 6
        SetText vasOrder, Trim(GetText(vasList, asRow, llCol)), llRow, llCol
    Next llCol
    vasList.DeleteRows asRow, 1
End Sub

Private Sub Command1_Click()
    Dim lRow, lRow1 As Long
    Dim lsID As String
    
    ClearSpread vasTemp
    ClearSpread vasList
    
    txtMsg = ""
    
    SQL = "select barcode from res_flag " & vbCrLf & _
          "where examdate = '" & Format(Date, "yyyymmdd") & "'  " & vbCrLf & _
          "  and SampleJudg = '1' "
    res = db_select_Vas(gLocal, SQL, vasTemp)
    
    txtMsg = txtMsg & "flag 개수 : " & vasTemp.DataRowCnt
    
    If vasTemp.DataRowCnt > 0 Then Timer1.Enabled = False
    
    For lRow = 1 To vasTemp.DataRowCnt
        lsID = Trim(GetText(vasTemp, lRow, 1))
        
        lRow1 = vasList.DataRowCnt + 1
        
        txtMsg = txtMsg & lsID & " : " & lRow1
        
        If Len(lsID) > 9 And IsNumeric(lsID) = True Then
        
            SQL = "Select wnifsmyr || LPAD(to_char(wnifsmsn), 7, '0') || to_char(wnifsms1), " & _
                  " '', wnifidno, wnifname, " & vbCrLf & _
                  " WNIFACDT, WNIFACTM, wnifjscd, wnifwkno, WNIFRRNF, WNIFRSEX " & vbCrLf & _
                  "from arcwnifh a " & vbCrLf & _
                  "Where wnifdpcd = 'CP' " & vbCrLf & _
                  "  And wnifdate = to_char(sysdate, 'yyyymmdd') " & vbCrLf & _
                  "  And wnifslip in ('L20', 'L60', 'L67') " & vbCrLf & _
                  "  AND wnifitem = '00' " & vbCrLf & _
                  "  And wnifstat <> 'X' " & vbCrLf & _
                  "  and wnifsmyr = '" & Left(lsID, 2) & "'  " & vbCrLf & _
                  "  and wnifsmsn = " & Mid(lsID, 3, 7) & vbCrLf & _
                  "Order by WNIFACDT, wnifwkno "
            
            res = db_select_Vas(gServer, SQL, vasList, lRow1, 1)
            txtMsg = txtMsg & lsID & " : " & lRow1 & "[" & res & "]"
            If res > 0 Then
                vasList.SetText 2, lRow1, "1"
            End If
        End If
    Next lRow
    txtMsg = txtMsg & "flag 가져오기 : " & vasList.DataRowCnt
    
    SQL = "Select wnifsmyr || LPAD(to_char(wnifsmsn), 7, '0') || to_char(wnifsms1), " & _
          " '', wnifidno, wnifname, " & vbCrLf & _
          " WNIFACDT, WNIFACTM, wnifjscd, wnifwkno, WNIFRRNF, WNIFRSEX " & vbCrLf & _
          "from arcwnifh a " & vbCrLf & _
          "Where wnifdpcd = 'CP' " & vbCrLf & _
          "  And wnifdate = to_char(sysdate, 'yyyymmdd') " & vbCrLf & _
          "  And wnifslip in ('L20', 'L60', 'L67') " & vbCrLf & _
          "  AND wnifitem = '00' " & vbCrLf & _
          "  And wnifstat in ('0', '9') " & vbCrLf & _
          "Order by WNIFACDT, wnifwkno "
    
    res = db_select_Vas(gServer, SQL, vasList, vasList.DataRowCnt + 1, 1)
    txtMsg = txtMsg & "작업예정 : " & vasList.DataRowCnt
    
    lRow = 1
    Do While lRow <= vasList.DataRowCnt
        txtMsg = txtMsg & Trim(GetText(vasList, lRow, 1)) & " : " & Trim(GetText(vasList, lRow, 2))
        If Trim(GetText(vasList, lRow, 2)) = "1" Then
            lRow = lRow + 1
        Else
            SQL = "select barcode, OrdFlag from worklist where barcode = '" & Trim(GetText(vasList, lRow, 1)) & "' "
            res = db_select_Col(gLocal, SQL)
            If Trim(gReadBuf(0)) <> "" Then
                txtMsg = txtMsg & Trim(GetText(vasList, lRow, 1)) & " : delete"
                DeleteRow vasList, lRow, lRow
            Else
                SQL = "select barcode from pat_res where barcode = '" & Trim(GetText(vasList, lRow, 1)) & "' "
                res = db_select_Col(gLocal, SQL)
                If Trim(gReadBuf(0)) = Trim(GetText(vasList, lRow, 1)) Then
                    txtMsg = txtMsg & Trim(GetText(vasList, lRow, 1)) & " : delete"
                    DeleteRow vasList, lRow, lRow
                Else
                    lRow = lRow + 1
                End If
            End If
        End If
    Loop
End Sub

Private Sub Form_Load()
    Dim db_tmp  As String * 100
    Dim lsData  As String
    Dim i       As Integer
        
    cn_Local_Flag = False
    cn_Server_Flag = False
    
    GetSetup
    
    If Connect_Local Then
        cn_Local_Flag = True
    End If
    
    gRow = 1
    
    If Not IsNumeric(gExpireDate) Then
        gExpireDate = 15
    End If
    
    'gExpireDate = Format(DateAdd("d", 0 - CInt(gExpireDate), CDate(GetDateFull)), "yyyy-mm-dd") & " 00:00:00"
    'gExpireDate = Format(DateAdd("d", 0 - CInt(gExpireDate), CDate(GetDateFull)), "yyyymmdd") & " 00:00:00"
    
'    SQL = "Delete from WorkList "
'    res = SendQuery(gLocal, SQL)
    
    dtpSDate.Value = CDate(Date)
    dtpEDate.Value = CDate(Date)
    
    SQL = "Select RemoteIP from WorkList "
    res = db_select_Col(gLocal, SQL)
    If res = -1 Then
        SQL = "Alter table worklist add column RemoteIP varchar(20) "
        res = SendQuery(gLocal, SQL)
    End If
    
    SQL = "Select ReceTime from WorkList "
    res = db_select_Col(gLocal, SQL)
    If res = -1 Then
        SQL = "Alter table worklist add column ReceTime varchar(20) "
        res = SendQuery(gLocal, SQL)
    End If
    
    SQL = "Select WkNo from WorkList "
    res = db_select_Col(gLocal, SQL)
    If res = -1 Then
        SQL = "Alter table worklist add column WkNo varchar(20) "
        res = SendQuery(gLocal, SQL)
    End If
    
    SQL = "Select SlideOrd from res_flag "
    res = db_select_Col(gLocal, SQL)
    If res = -1 Then
        SQL = "Alter table res_flag add column SlideOrd varchar(2) "
        res = SendQuery(gLocal, SQL)
    End If
        
    If Not IsNumeric(gTimer) Then
        gTimer = 30
    End If
    
    GetExamCode
    
    gbOrdering = False
    
    cmdPortOpen_Click
    'Timer1.Interval = gTimer
    'cmdReLoad_Click
End Sub

Sub LASCPortOpen()
    
    GetSetup_LASC
    
    MSComm1.CommPort = gSetup.Port
    MSComm1.Settings = gSetup.Speed & "," & gSetup.Parity & "," & gSetup.DataBit & "," & gSetup.StopBit
    If gSetup.DTREnable = "1" Then
        MSComm1.DTREnable = True
    Else
        MSComm1.DTREnable = False
    End If
    If gSetup.RTSEnable = "1" Then
        MSComm1.RTSEnable = True
    Else
        MSComm1.RTSEnable = False
    End If
    
    MSComm1.PortOpen = True

    If MSComm1.CTSHolding = False Then
        lblMsg.Caption = "[MSG]LASC 의 포트가 준비되지 않았습니다"
        lblMsg.ForeColor = RGB(255, 0, 0)

        'MSComm1.PortOpen = False
        CmdPortOpen.Caption = "PortOpen"

        lblMsg.Caption = "[MSG]포트가 닫혀있습니다"

        Timer1.Enabled = False
    Else
        lblMsg.Caption = "[MSG]LASC 의 포트가 준비되었습니다"
        lblMsg.ForeColor = RGB(0, 0, 0)
    End If
    
    Timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Timer1.Enabled = False
    
    If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
    
    
    DisConnect_Local
    
    End
End Sub

Private Sub MSComm1_OnComm()
    Dim s As String
    
    s = MSComm1.Input
    
    If s = chrACK Then
        SaveOrdLog s
        
        gbOrdering = False
        'SaveOrdLog s
        If gSetup.Protocol = "B" Then
            If vasList.DataRowCnt > 0 Then
                OrderEntry_1 1
            Else
                Timer1.Enabled = True
            End If
        End If
    End If
End Sub


Private Sub Timer1_Timer()
    
    On Error Resume Next
    
    If MSComm1.CTSHolding = False Then
        lblMsg.Caption = "[Message] LASC 의 포트가 준비되지 않았습니다"
        lblMsg.ForeColor = RGB(255, 0, 0)
    Else
        lblMsg.Caption = "[Message] LASC 의 포트가 준비되었습니다"
        lblMsg.ForeColor = RGB(0, 0, 0)
    End If
    

    Dim lsID        As String
    Dim lsID1       As String
    Dim i, j, k     As Integer
    Dim ii          As Integer
    Dim lsExamCode  As String
    
    Dim AdoRs_Exam As ADODB.Recordset
    Dim Ord(7)  As String
    Dim lsOrder As String
    Dim lRow    As Long
    Dim lRow1   As Long
    Dim lsDate  As String
    
    Dim lsWkNo  As String
    Dim lsPID   As String
    Dim lsPName As String
    
    Dim lsSlideOrd As String
    Dim lsExamDate As String
    
    Dim lsSlideName As String
    
    Dim lsReceNo    As String
    
'    If gbOrdering Then
'        SaveOrdLog "Ordering 중 Timer"
'        'Exit Sub
'    End If
    
'    If gbOrdering = True Then
'        SaveOrdLog "TIMER"
'    Else
'        SaveOrdLog "Ordering"
'        Exit Sub
'    End If
    
    'SaveOrdLog "TIMER"
    
    If MSComm1.PortOpen = False Then
        LASCPortOpen
    End If
        
    If MSComm1.CTSHolding = False Then
        lblMsg.Caption = "[Message] LASC 의 포트가 준비되지 않았습니다"
        lblMsg.ForeColor = RGB(255, 0, 0)
        
        'Exit Sub
    Else
        lblMsg.Caption = "[Message] LASC 의 포트가 준비되었습니다"
        lblMsg.ForeColor = RGB(0, 0, 0)
    End If
    
            
'    DisConnect_Server
'
'    If Connect_Server Then
'        cn_Server_Flag = True
'    End If
    
    ClearSpread vasList
    ClearSpread vasTemp
    
    'SampleJudg = '1' 인 것은 슬라이드 오더 들어가도록
    SQL = "select barcode from res_flag " & vbCrLf & _
          "where examdate = '" & Format(Date, "yyyymmdd") & "'  " & vbCrLf & _
          "  and SampleJudg = '1' "
    res = db_select_Vas(gLocal, SQL, vasTemp)
    
    For lRow = 1 To vasTemp.DataRowCnt
        lsID = Trim(GetText(vasTemp, lRow, 1))
        
        lRow1 = vasList.DataRowCnt + 1
        
        'txtMsg = txtMsg & lsID & " : " & lRow1 & vbCrLf
        
        If Len(lsID) = 11 And IsNumeric(lsID) = True Then
            res = Online_XML(gXml_S03, lsID)
            'txtMsg = txtMsg & lsID & " : " & lRow1 & "[" & res & "]"
            'res = 1
            If res > 0 Then
                vasList.SetText 1, lRow1, lsID
                vasList.SetText 2, lRow1, "1"
                
                vasList.SetText 3, lRow1, gPat_Info_Select.PT_NO
                vasList.SetText 4, lRow1, gPat_Info_Select.PT_NM
                If IsDate(gPat_Info_Select.ACPT_DTETM) Then
                    vasList.SetText 5, lRow1, Format(CDate(gPat_Info_Select.ACPT_DTETM), "yyyy-mm-dd")    '날짜
                Else
                    vasList.SetText 5, lRow1, gPat_Info_Select.ACPT_DTETM   '날짜
                End If
                vasList.SetText 6, lRow1, ""    '시간
                vasList.SetText 7, lRow1, ""    '접수코드
                vasList.SetText 8, lRow1, gPat_Info_Select.ACPTNO_1
                vasList.SetText 9, lRow1, gPat_Info_Select.Sex
                vasList.SetText 10, lRow1, gPat_Info_Select.Age
                vasList.SetText 11, lRow1, ""   'slip
            End If
        End If
    Next lRow
    'txtMsg = txtMsg & "flag 가져오기 : " & vasList.DataRowCnt
    
    '이전
    'res = Get_WorkList(Format(Date, "yyyymmdd"), Format(Date, "yyyymmdd"), gsExamAll, 3)
    
    res = Online_TLA(gXml_S06, dtpSDate.Value, dtpEDate.Value)
    
    If res = 1 Then
        lsID1 = ""
        
        With vasList
            For lRow1 = 0 To UBound(gTLA_Info_Select)
                lsID = Trim(gTLA_Info_Select(lRow1).SPCNO)
                
    '            '해당 검사항목 확인하기
    '            Dim sCnt As String
    '
    '            sCnt = "0"
    '            lsExamCode = ""
    '            res = Online_XML(gXml_S07, lsID)
    '            For ii = 0 To UBound(gExam_Select)
    '                If lsExamCode = "" Then
    '                    lsExamCode = "'" & gExam_Select(ii).TST_CD & "'"
    '                Else
    '                    lsExamCode = lsExamCode & ",'" & gExam_Select(ii).TST_CD & "'"
    '                End If
    '            Next ii
    '
    '            SQL = "Select count(examcode) from equipexam where examcode in (" & lsExamCode & ")"
    '            res = db_select_Var(gLocal, SQL, sCnt)
    '            If sCnt = "" Then sCnt = "0"
    '
    '            If sCnt > "0" Then
                If Len(lsID) = 11 Then
                    '접수번호 가져오기
                    res = Online_XML(gXml_S03, lsID)
                    
                    lsReceNo = Trim(gPat_Info_Select.ACPTNO_1)
                    If lsID <> lsID1 Then
                        SQL = "select barcode, OrdFlag from worklist where barcode = '" & lsID & "' and WkNo = '" & lsReceNo & "' "
                        res = db_select_Col(gLocal, SQL)
                        If Trim(gReadBuf(0)) <> "" Then
                            'txtMsg = txtMsg & Trim(GetText(vasList, lRow, 1)) & " : delete"
                            'DeleteRow vasList, lRow, lRow
                        Else
                            SQL = "select barcode from pat_res where barcode = '" & lsID & "'and receno = '" & lsReceNo & "'  "
                            res = db_select_Col(gLocal, SQL)
                            If Trim(gReadBuf(0)) = lsID Then
                                'txtMsg = txtMsg & Trim(GetText(vasList, lRow, 1)) & " : delete"
                                'DeleteRow vasList, lRow, lRow
                            Else
                                lRow = .DataRowCnt + 1
                                
                                .SetText 1, lRow, lsID
                                .SetText 8, lRow, Trim(gPat_Info_Select.ACPTNO_1)
                                
                                .SetText 3, lRow, gPat_Info_Select.PT_NO
                                .SetText 4, lRow, gPat_Info_Select.PT_NM
                                '.SetText 5, lRow, gPatient_Info.ACPT_DTM    '날짜
                                If IsDate(gPat_Info_Select.ACPT_DTETM) Then
                                    .SetText 5, lRow, Format(CDate(gPat_Info_Select.ACPT_DTETM), "yyyy-mm-dd")    '날짜
                                Else
                                    .SetText 5, lRow, gPat_Info_Select.ACPT_DTETM    '날짜
                                End If
                                .SetText 6, lRow, ""    '시간
                                .SetText 7, lRow, "" '접수코드
                                .SetText 8, lRow, gPat_Info_Select.ACPTNO_1
                                .SetText 9, lRow, gPat_Info_Select.Sex
                                .SetText 10, lRow, gPat_Info_Select.Age
                                .SetText 11, lRow, ""   'slip
                                
                                '2011.10.31 이상은 추가
                                .SetText 12, lRow, gPat_Info_Select.CAUTION_YN
                                .SetText 13, lRow, gPat_Info_Select.ORD_SITE
        
                                '2012.04.25 오세원 추가
                                .SetText 14, lRow, gPat_Info_Select.PATSECT
        
                            End If
                        End If
                    End If
                End If
                lsID1 = lsID
            Next lRow1
        End With
    End If
    
    If vasList.DataRowCnt > 0 Then
        Timer1.Enabled = False
'        gbOrdering = True
    Else
        Timer1.Enabled = True
        Exit Sub
    End If
    
    If gSetup.Protocol = "B" Then
        OrderEntry_1 1
        Exit Sub
    Else
        For lRow = 1 To vasList.DataRowCnt
            OrderEntry_1 1
        Next lRow
    End If

    Timer1.Enabled = True
End Sub

Function OrderEntry(asRow As Long) As Integer
    'Group Order Format
    
    Dim lsID As String
    Dim i, j, k As Integer
    Dim AdoRs_Exam As ADODB.Recordset
    Dim Ord(7) As String
    Dim lsOrder As String
    Dim lRow, lRow1 As Long
    Dim lsDate As String
    
    Dim lsPatInfo As String
    Dim lsWkNo As String
    Dim lsPID As String
    Dim lsPName As String
    Dim lsPEName As String
    Dim lsPAge As String
    Dim lsPSex As String
    Dim lsPBirth As String
    Dim lsWard As String
    
    Dim lsSlideName1 As String
    Dim lsSlideOrd As String
    Dim lsExamDate As String
    
    Dim lsSlideName As String
    Dim iMOR As Integer
    Dim iPBS As Integer
    
    Dim iSP As Integer
    
    Dim lsReceNo As String
    
    lRow = asRow
    
    If lRow < 1 Or lRow > vasList.DataRowCnt Then Exit Function
    
    lsExamDate = Format(Date, "yyyymmdd")
    
    SQL = "select equipcode, examcode, examname, OrdGubun, examno from equipexam where equip = '" & gEquip & "' and OrdGubun <> 'N' order by 2, 5 "
    Set AdoRs_Exam = db_select_rs(gLocal, SQL)
        
    lsSlideName = ""
        
    lsOrder = ""
    For i = 1 To 7
        Ord(i) = "0"
    Next i
                
    lsSlideOrd = ""
        
    lsID = Trim(GetText(vasList, lRow, 1))
    
    lblBarCode.Caption = lsID
    
    If Trim(GetText(vasList, lRow, 2)) = "1" Then
        iSP = 1
    Else
        iSP = 0
    End If
    
    'lsWkNo = SetSpace(Trim(gReadBuf(1)), 5)
    lsPID = Trim(GetText(vasList, lRow, 3))
    lsPName = Trim(GetText(vasList, lRow, 4))
    
    lsSlideName = lsPName
    lsSlideName = Conv_Kor_Eng(lsSlideName)
    lsPEName = lsSlideName
    lsWkNo = Trim(GetText(vasList, lRow, 8))
    lsPBirth = Trim(GetText(vasList, lRow, 9))
    lsPSex = Trim(GetText(vasList, lRow, 10))
    lsWard = ""
    
    Select Case lsPSex
    Case "3", "4"
        lsPBirth = "20" & lsPBirth
    Case "1", "2", "5", "6", "7", "8", "9", "0"
        lsPBirth = "19" & lsPBirth
    Case Else
        lsPBirth = ""
    End Select
    
    If IsNumeric(lsPSex) Then
        If CInt(lsPSex) Mod 2 = 0 Then
            lsPSex = "F"
        Else
            lsPSex = "M"
        End If
    End If
    
    lsDate = GetDateFull

    
    ClearSpread vasTemp
    ClearSpread vasExam
    
    lsReceNo = lsWkNo
    
    lsSlideName1 = Mid(lsExamDate, 3, 2) & "/" & Mid(lsExamDate, 5, 2) & "/" & Mid(lsExamDate, 7, 2)
'    If vasExam.DataRowCnt > 0 Then
'        lsSlideName1 = lsSlideName1 & "  " & Trim(GetText(vasExam, 1, 3))
'    End If
    
    iMOR = -1
    iPBS = -1
    
    'res = Get_Order(lsID)
    res = Online_XML(gXml_S07, lsID)
    
    For j = 0 To UBound(gExam_Select)
        
        If Not AdoRs_Exam Is Nothing Then
            AdoRs_Exam.MoveFirst
            
            Debug.Print (GetText(vasExam, j + 1, 1)) & "  " & Trim(AdoRs_Exam("examcode"))
            
            Do Until AdoRs_Exam.EOF
                If Trim(AdoRs_Exam("examcode")) = gExam_Select(j).TST_CD Then
                    Select Case Trim(AdoRs_Exam("OrdGubun"))
                    Case "C": Ord(1) = "1"
                    Case "D": Ord(2) = "1"
                    Case "R"
                        Ord(3) = "1"
                        'Ord(1) = "1"
                    Case "P"
                        Ord(4) = "1"
                        lsSlideOrd = "SP"
                    Case "S"
                        Ord(5) = "1"
                        lsSlideOrd = "SC"

                    Case "X": Ord(6) = "1"
                    Case "B": Ord(7) = "1"
                    End Select
                    
                    If gExam_Select(j).TST_CD = "L2023" Or gExam_Select(j).TST_CD = "L20231" Or gExam_Select(j).TST_CD = "L20232" Then
                        iPBS = 1
                    End If

'                        Case "CP0106"   'Morphology
'                            lsSlideName = "MOR." & lsSlideName
'                        Case "CP0131"   'PB Smear
'                            lsSlideName = "PBS." & lsSlideName
'                        End Select
                    Exit Do
                End If
                
                AdoRs_Exam.MoveNext
            Loop
        End If
    Next j
        
    'If Ord(5) = "1" Then Ord(2) = "1"
    
    If iSP = 1 Then
        Ord(4) = "1"
        lsSlideOrd = "SP"
        SQL = "update res_flag set SampleJudg = '0' " & vbCrLf & _
              "where examdate = '" & Format(Date, "yyyymmdd") & "'  " & vbCrLf & _
              "  and barcode = '" & lsID & "' "
        res = SendQuery(gLocal, SQL)
    End If
    
    If iPBS = 1 Then
        lsSlideOrd = "SP"
        Ord(1) = "1"
        Ord(2) = "1"
    End If
    
    If Ord(4) = "1" Then Ord(5) = "0"
    
    lsOrder = ""
    For i = 1 To 7
        lsOrder = lsOrder & Ord(i)
    Next i
    
'    If iPBS = 1 Then
'        lsSlideName = "PBS." & lsSlideName
'    Else
'        If iMOR = 1 Then
'            lsSlideName = "MOR." & lsSlideName
'        End If
'    End If
    
    'MsgBox lsOrder
    
    If lsOrder <> "0000000" And lsOrder <> "" Then
        'DoSleep gOrdGap
        
        lsSlideName = Left(lsSlideName, 13)
        lsSlideName = SetChar(lsSlideName, 13, 1, " ")
        lsPatInfo = lsPID
        lsPatInfo = Left(lsPatInfo, 13)
        lsPatInfo = SetChar(lsPatInfo, 13, 1, " ")
        lsSlideName1 = lsSlideName1 & " " & lsWkNo
        lsSlideName1 = Left(lsSlideName1, 13)
        lsSlideName1 = SetChar(lsSlideName1, 13, 1, " ")

        '동아대 작업 : Slide 이름에 (MOR)+환자영문이름 => 최대 자리수 잘라넣기
        '2006년 9월 29일 환자 정보 더 넣기
'            lsOrder = chrSTX & "S" & "00000000" & SetChar(lsID, 15, 1, " ") & "********" & lsOrder & "00000"
'            lsOrder = lsOrder & lsSlideName
'            lsOrder = lsOrder & "0000000000000"
'            lsOrder = lsOrder & "0000000000000"
'            lsOrder = lsOrder & "000****************************************" & chrETX
        
        lsOrder = chrSTX & "S" & "00000000" & SetChar(lsID, 15, 1, " ") & "********" & lsOrder & "00000"
        lsOrder = lsOrder & "0000000000000"
        lsOrder = lsOrder & "0000000000000"
'        lsOrder = lsOrder & lsSlideName1
'        lsOrder = lsOrder & lsPatInfo
        lsOrder = lsOrder & lsSlideName
        lsOrder = lsOrder & "000"
        'lsOrder = lsOrder & SetSpace(lsPID, 16, 1)
'        Select Case lsPSex
'        Case "M"
'            lsOrder = lsOrder & "1"
'        Case "F"
'            lsOrder = lsOrder & "2"
'        Case Else
'            lsOrder = lsOrder & "3"
'        End Select
'        lsOrder = lsOrder & SetSpace(Trim(lsPBirth), 8, 1)
        lsOrder = lsOrder & "****************"
        lsOrder = lsOrder & "*********"
        lsOrder = lsOrder & "***************" & chrETX
    
            
        OrderOutput lsOrder
        

'                MSComm1.Output = lsOrder
'                SaveOrdLog lsOrder
        
        SQL = "Select barcode from res_flag where examdate = '" & Format(Date, "yyyymmdd") & "' and Barcode = '" & lsID & "'"
        res = db_select_Col(gLocal, SQL)
        If Trim(gReadBuf(0)) = lsID Then
            SQL = "update res_flag set SlidOrd = '" & lsSlideOrd & "', SampleJudg = '0' " & vbCrLf & _
                  "where examdate = '" & Format(Date, "yyyymmdd") & "' and Barcode = '" & lsID & "'"
            res = SendQuery(gLocal, SQL)
        Else
            SQL = "Insert Into res_flag (examdate, barcode, SampleJudg, PosDiff, PosMorph, PosCnt, ErrFunc, ErrRes, WBCAbnor, WBCSusp, RBCAbnor, RBCSusp, PLTAbnor, PLTSusp, InfoWBC, InfoPLT, PBSFlag, SlideOrd ) " & vbCrLf & _
                  "Values ('" & Format(Date, "yyyymmdd") & "', '" & lsID & "', '0', '', '', '', " & _
                  "'', '', '', '', '', '', " & _
                  "'', '', '', '', '', '" & lsSlideOrd & "') "
            res = SendQuery(gLocal, SQL)
        End If

        SQL = "select barcode from worklist where barcode = '" & lsID & "' "
        res = db_select_Col(gLocal, SQL)
        If Trim(gReadBuf(0)) = Trim(lsID) Then
            SQL = "delete from worklist where barcode = '" & lsID & "' "
            res = SendQuery(gLocal, SQL)
        End If
        SQL = "Insert Into WorkList(ReceDate, ReceTime, Barcode, PID, PName, OrdFlag, OrdDateTime, ResDateTime, OrdCnt, RemoteIP, WkNo ) " & vbCrLf & _
              "Values ('" & Trim(GetText(vasList, lRow, 5)) & "', '" & Trim(GetText(vasList, lRow, 6)) & "', '" & lsID & "', '" & Trim(GetText(vasList, lRow, 3)) & "','" & Trim(GetText(vasList, lRow, 4)) & "', 'B','" & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "hh:nn:ss") & "','', 0, 'cbcworklist', '" & lsReceNo & "') "
        res = SendQuery(gLocal, SQL)
    Else
        SQL = "select barcode from worklist where barcode = '" & lsID & "' "
        res = db_select_Col(gLocal, SQL)
        If Trim(gReadBuf(0)) = Trim(lsID) Then
            SQL = "delete from worklist where barcode = '" & lsID & "' "
            res = SendQuery(gLocal, SQL)
        End If
        SQL = "Insert Into WorkList(ReceDate, ReceTime, Barcode, PID, PName, OrdFlag, OrdDateTime, ResDateTime, OrdCnt, RemoteIP, WkNo ) " & vbCrLf & _
              "Values ('" & Trim(GetText(vasList, lRow, 5)) & "', '" & Trim(GetText(vasList, lRow, 6)) & "', '" & lsID & "', '" & Trim(GetText(vasList, lRow, 3)) & "','" & Trim(GetText(vasList, lRow, 4)) & "', 'B','" & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "hh:nn:ss") & "','', 0, 'cbcworklist', '" & lsReceNo & "') "
        res = SendQuery(gLocal, SQL)
        
        Exit Function
    End If
        
    
    lRow1 = vasOrder.DataRowCnt + 1
    If lRow1 > vasOrder.MaxRows Then vasOrder.MaxRows = lRow1
    
    InsertRow vasOrder, 1

    vasOrder.SetText 1, 1, Trim(GetText(vasList, lRow, 1))
    vasOrder.SetText 2, 1, "B"
    vasOrder.SetText 3, 1, Trim(GetText(vasList, lRow, 3))
    vasOrder.SetText 4, 1, Trim(GetText(vasList, lRow, 4))
    vasOrder.SetText 5, 1, Trim(GetText(vasList, lRow, 5))
    vasOrder.SetText 6, 1, Trim(GetText(vasList, lRow, 6))
    vasOrder.SetText 7, 1, Trim(GetText(vasList, lRow, 7))
    vasOrder.SetText 8, 1, Trim(GetText(vasList, lRow, 8))
    
        
    DeleteRow vasList, lRow, lRow
End Function

Function OrderEntry_1(asRow As Long) As Integer
    'Individual Order Format
    
    Dim lsID    As String
    Dim i, j, k As Integer
    Dim AdoRs_Exam As ADODB.Recordset
    
    Dim Ord(7)  As String
    
    Dim lsOrder     As String
    Dim lsOrder1    As String
    Dim lsOrder2    As String
    Dim lRow, lRow1 As Long
    Dim lsDate      As String
    
    Dim lsPatInfo   As String
    Dim lsWkNo      As String
    Dim lsPID       As String
    Dim lsPName     As String
    Dim lsPEName    As String
    Dim lsPAge      As String
    Dim lsPSex      As String
    Dim lsPBirth    As String
    Dim lsWard      As String
    
    Dim lsSlideName1 As String
    Dim lsSlideOrd  As String
    Dim lsExamDate  As String
    
    Dim lsSlideName As String
    Dim iMOR As Integer
    Dim iPBS As Integer
    
    Dim iSP As Integer
    
    Dim lsReceNo    As String
    
    Dim sParam      As String
    
    Dim s1, s2, s3, s4, s5, s6, s7, s8, s9, s10 As String
    Dim s11, s12, s13, s14, s15, s16, s17, s18, s19, s20 As String
    Dim s21, s22, s23, s24, s25, s26, s27, s28, s29, s30 As String
    Dim s31, s32, s33, s34, s35, s36, s37, s38, s39, s40 As String
    
    Dim iEosoFlag As Integer
    
    Dim lsPATSECT    As String
    Dim m As Integer
    
    lRow = asRow
    
    If lRow < 1 Or lRow > vasList.DataRowCnt Then Exit Function
    
    lsExamDate = Trim(Format(Date, "yyyymmdd"))
    
    SQL = "select equipcode, examcode, examname, OrdGubun, examno from equipexam where equip = '" & gEquip & "' and OrdGubun <> 'N' order by 2, 5 "
    Set AdoRs_Exam = db_select_rs(gLocal, SQL)
        
    lsSlideName = ""
    lsSlideName1 = ""
    
    lsOrder = ""
'    For i = 1 To 7
'        Ord(i) = "0"
'    Next i
                
    lsSlideOrd = ""
        
    lsID = Trim(GetText(vasList, lRow, 1))
    
    '-- 2012.04.25 오세원 추가
    lsPATSECT = Trim(GetText(vasList, lRow, 14))
    
    lblBarCode.Caption = lsID
    
    '바코드자리수에 안 맞으면 프로시저 타지 말것
    If Len(lsID) <> 11 Then
        Exit Function
    End If
    
    If Trim(GetText(vasList, lRow, 2)) = "1" Then
        iSP = 1
    Else
        iSP = 0
    End If
    
    'lsWkNo = SetSpace(Trim(gReadBuf(1)), 5)
    lsPID = Trim(GetText(vasList, lRow, 3))
    lsPName = Trim(GetText(vasList, lRow, 4))
    
    lsSlideName = lsPName
    lsSlideName = Conv_Kor_Eng(lsSlideName)
    lsPEName = lsSlideName
    lsWkNo = Trim(GetText(vasList, lRow, 8))
    lsPBirth = Trim(GetText(vasList, lRow, 9))
    lsPSex = Trim(GetText(vasList, lRow, 10))
    lsWard = ""
    
    '-- 2012.04.25 오세원 추가
    lsPATSECT = Trim(GetText(vasList, lRow, 14))
    If Trim(lsPATSECT) = "" Then
        lsPATSECT = gPat_Info_Select.PATSECT
    End If
    
    Select Case lsPSex
    Case "3", "4"
        lsPBirth = "20" & lsPBirth
    Case "1", "2", "5", "6", "7", "8", "9", "0"
        lsPBirth = "19" & lsPBirth
    Case Else
        lsPBirth = ""
    End Select
    
    If IsNumeric(lsPSex) Then
        If CInt(lsPSex) Mod 2 = 0 Then
            lsPSex = "F"
        Else
            lsPSex = "M"
        End If
    End If
    
    lsDate = GetDateFull

    
    ClearSpread vasTemp
    ClearSpread vasExam
    
    lsReceNo = lsWkNo
    
    'lsSlideName1 = Mid(lsExamDate, 3, 2) & "/" & Mid(lsExamDate, 5, 2) & "/" & Mid(lsExamDate, 7, 2)
    lsSlideName1 = lsExamDate
    
'    If vasExam.DataRowCnt > 0 Then
'        lsSlideName1 = lsSlideName1 & "  " & Trim(GetText(vasExam, 1, 3))
'    End If
    
    s1 = "0"
    s2 = "0"
    s3 = "0"
    s4 = "0"
    s5 = "0"
    s6 = "0"
    s7 = "0"
    s8 = "0"
    s9 = "0"
    s10 = "0"
    s11 = "0"
    s12 = "0"
    s13 = "0"
    s14 = "0"
    s15 = "0"
    s16 = "0"
    s17 = "0"
    s18 = "0"
    s19 = "0"
    s20 = "0"
    s21 = "0"
    s22 = "0"
    s23 = "0"
    s24 = "0"
    s25 = "0"
    s26 = "0"
    s27 = "0"
    s28 = "0"
    s29 = "0"
    s30 = "0"
    s31 = "0"
    s32 = "0"
    s33 = "0"
    s34 = "0"
    s35 = "0"
    s36 = "0"
    s37 = "0"
    s38 = "0"
    s39 = "0"
    s40 = "0"
    
    
    iMOR = -1
    iPBS = -1
    
    iEosoFlag = -1
    
    res = Online_XML(gXml_S07, lsID)
    
    For j = 0 To UBound(gExam_Select)
        
        If Not AdoRs_Exam Is Nothing Then
            AdoRs_Exam.MoveFirst
            
            Debug.Print (GetText(vasExam, j + 1, 1)) & "  " & Trim(AdoRs_Exam("examcode"))
            
            Do Until AdoRs_Exam.EOF
                If gExam_Select(j).TST_CD <> "" Then
                    If Trim(AdoRs_Exam("examcode")) = gExam_Select(j).TST_CD Then
                        Select Case AdoRs_Exam("equipcode")
                        Case "01"
                            s1 = "1"
                            s8 = "1"    'PLT
                            s23 = "1"   'P-LCR
                        Case "02"
                            s2 = "1"
                        Case "03"
                            s3 = "1"
                        Case "04"
                            s4 = "1"
                        Case "05"
                            s5 = "1"
                        Case "06"
                            s6 = "1"
                        Case "07"
                            s7 = "1"
                        Case "08"
                            s8 = "1"
                        Case "09"
                            s9 = "1"
                            s14 = "1"   'LYMPH#
                        Case "10"
                            s10 = "1"
                            s15 = "1"   'MONO#
                        Case "11"
                            s11 = "1"
                            s16 = "1"   'NEUT#
                        Case "12"
                            s12 = "1"
                            s17 = "1"   'EO#
                        Case "13"
                            s13 = "1"
                            s18 = "1"   'BASO#
                            
                        Case "14"
                            s14 = "1"
                        Case "15"
                            s15 = "1"
                        Case "16"
                            s16 = "1"
                        Case "17"       'Eosino
                            s17 = "1"
                            'iEosoFlag = 1
                            iSP = 1
                            iPBS = 1
                        Case "18"
                            If s13 = "1" Then
                                s18 = "1"
                            End If

                        Case "19"
                            s19 = "1"   'RDW-CV
                            s20 = "1"   'RDW-SD
'                        Case "20"
'                            If s19 = "1" Then
'                                s20 = "1"
'                            End If
                        
                        Case "21"
                            s21 = "1"
                        Case "22"
                            s22 = "1"
                        Case "23"
                            s23 = "1"
                        Case "24"       'Reti
                            s24 = "1"
                            s25 = "1"
                            s26 = "1"
                            s27 = "1"
                            s28 = "1"
                            s29 = "1"
                        Case "25"
                            s25 = "1"
                        Case "26"
                            s26 = "1"
                        Case "27"
                            s27 = "1"
                        Case "28"
                            s28 = "1"
                        Case "29"
                            s29 = "1"
                        Case "30"
                            s30 = "1"
                        Case "31"
                            s31 = "1"
                        Case "32"
                            s32 = "1"
                        Case "33"
                            s33 = "1"
                        Case "34"
                            s34 = "1"
                        Case "35"
                            s35 = "1"
                        Case "36"
                            s36 = "1"
                        
                        '**********************************
                        Case "37"   'PB Smear(L2112)
                            s37 = "1"
                            iSP = 1
                            iPBS = 1
                        Case "38"   '말라리아(L2125)    CBC+DIFF+슬라이드
                            s38 = "1"
                            iSP = 1
                            iPBS = 1
                        Case "39"   '말라리아(L8616)    CBC+DIFF+슬라이드
                            s39 = "1"
                            iSP = 1
                            iPBS = 1
                        '**********************************
                        
                        Case "40"
                            s40 = "1"
                        End Select
                    End If
                End If
                
                '-- 2012.04.25 오세원 추가
                'If gExam_Select(j).TST_CD = "L2111" Then
                
                '-- 2012.05.07 오세원 수정 L2111 >> L2112
                If gExam_Select(j).TST_CD = "L2112" Then
                    iPBS = 1
                    iSP = 1
                End If
                
                AdoRs_Exam.MoveNext
            Loop
        End If
    Next j
        
    '2011.10.31 이상은 - 무조건 슬라이드 밀기
'    If Trim(GetText(vasList, lRow, 12)) = "Y" And Trim(GetText(vasList, lRow, 13)) = "SPHO" Then
    
    '2012.04.26 오세원 - 무조건 슬라이드 밀기 ==> 조건수정
    If Trim(GetText(vasList, lRow, 12)) = "Y" Or Trim(GetText(vasList, lRow, 13)) = "SPHO" Then
        iSP = 1
        iPBS = 1
    End If
    
    If iSP = 1 Then
        lsSlideOrd = "SP"
        SQL = "update res_flag set SampleJudg = '0' " & vbCrLf & _
              "where examdate = '" & Format(Date, "yyyymmdd") & "'  " & vbCrLf & _
              "  and barcode = '" & lsID & "' "
        res = SendQuery(gLocal, SQL)
    End If
    
    If iPBS = 1 Then
        lsSlideOrd = "SP"
'        Ord(1) = "1"
'        Ord(2) = "1"
    End If
    
    'If Ord(4) = "1" Then Ord(5) = "0"
    
    lsOrder = ""
'    For i = 1 To 7
'        lsOrder = lsOrder & Ord(i)
'    Next i
    
'    If iPBS = 1 Then
'        lsSlideName = "PBS." & lsSlideName
'    Else
'        If iMOR = 1 Then
'            lsSlideName = "MOR." & lsSlideName
'        End If
'    End If
    
    lsOrder = s1 & s2 & s3 & s4 & s5 & s6 & s7 & s8
    lsOrder = lsOrder & s9 & s10 & s11 & s12 & s13 & s14 & s15 & s16 & s17 & s18 & s19 & s20 & s21 & s22 & s23
    lsOrder = lsOrder & "00"
    lsOrder = lsOrder & s24 & s25 & s26 & s27 & s28 & s29
    lsOrder = lsOrder & "0" & s30 & s31 & s32
    lsOrder = lsOrder & "00000000000000"
    
    'If lsOrder <> "0000000" And lsOrder <> "" Then
    If lsOrder <> "" Then
        'DoSleep gOrdGap
        
        '20100127-206410012700336 ILee gi suk
        
        '환자이름
        lsSlideName = lsSlideName
        lsSlideName = Left(lsSlideName, 13)
        lsSlideName = SetChar(lsSlideName, 13, 2, " ")
        
        '-- 2012.05.07 오세원 수정
        lsSlideName = Replace(lsSlideName, " ", "_")
'        If Right(lsSlideName, 1) = " " Then
'            lsSlideName = Mid(lsSlideName, 1, 12) & "_"
'        End If
        
        '바코드번호 외래/입원구분
        'lsPatInfo = lsPID
        lsPatInfo = lsID
        lsPatInfo = Left(lsPatInfo, 12)
        
        '-- 2012.04.25 오세원 수정
        'lsPatInfo = SetChar(lsPatInfo, 12, 2, " ") & "I"
'        If lsPATSECT <> "" Then
'            lsPatInfo = SetChar(lsPatInfo, 12, 2, " ") & Mid(lsPATSECT, 1, 1)
'        Else
'            lsPatInfo = SetChar(lsPatInfo, 12, 2, " ") & " "
'        End If

        
        If Trim(lsPATSECT) = "" Then lsPATSECT = " "
        '-- 2012.05.07 오세원 수정 >> 입.외 구분이 없으면 'O' 외래로 한다.
        'If Trim(lsPATSECT) = "" Then lsPATSECT = "O"
        
        lsPatInfo = SetChar(lsPatInfo, 12, 2, " ") & lsPATSECT
        
        
        '작업일자-작업번호
        lsSlideName1 = lsSlideName1 & "-" & SetChar(lsWkNo, 4, 2, "")
        lsSlideName1 = Left(lsSlideName1, 13)
        lsSlideName1 = SetChar(lsSlideName1, 13, 1, " ")
        
        lsOrder1 = ""
        
        'MsgBox lsOrder
        
        
'        Select Case lsOrder
'        Case "1000000"      'CBC
                                '12345678901234567890123            456     789 012     34567890123456789
'            lsOrder1 = "001" & "11111111000000000011111" & "00" & "000" & "0000100" & "00000000000000"
'        Case "1000100"      'CBC+SC
'            lsOrder1 = "001" & "11111111000000000011111" & "00" & "000" & "0000100" & "00000000000000"
'
'        Case "1100000"      'CBC+Diff
'            lsOrder1 = "001" & "11111111111111111111111" & "00" & "000" & "0000100" & "00000000000000"
'        Case "1100100"      'CBC+Diff+SC
'            lsOrder1 = "001" & "11111111111111111111111" & "00" & "000" & "0000100" & "00000000000000"
'
'        Case "1110000"      'CBC+Diff+Reti
'            lsOrder1 = "001" & "11111111111111111111111" & "00" & "111" & "1111100" & "00000000000000"
'        Case "1110100"      'CBC+Diff+Reti+SC
'            lsOrder1 = "001" & "11111111111111111111111" & "00" & "111" & "1111100" & "00000000000000"
'
'        Case "0010000"
'            lsOrder1 = "001" & "00000000000000000000000" & "00" & "111" & "1110000" & "00000000000000"
'        End Select
        
        'RET/SP/SC*****************************************
        'If iEosoFlag = 1 Then       'Eosinophil count
        If iSP = 1 Then     'Slide 밀기
            lsOrder1 = "010" & lsOrder
        Else
            lsOrder1 = "001" & lsOrder
        End If
        '**************************************************
        
        
        '동아대 작업 : Slide 이름에 (MOR)+환자영문이름 => 최대 자리수 잘라넣기
        '2006년 9월 29일 환자 정보 더 넣기
'            lsOrder = chrSTX & "S" & "00000000" & SetChar(lsID, 15, 1, " ") & "********" & lsOrder & "00000"
'            lsOrder = lsOrder & lsSlideName
'            lsOrder = lsOrder & "0000000000000"
'            lsOrder = lsOrder & "0000000000000"
'            lsOrder = lsOrder & "000****************************************" & chrETX
        
        lsOrder = chrSTX & "S" & "00000000" & SetChar(lsID, 15, 1, " ") & "********" & lsOrder1 & "00000"

        lsOrder = lsOrder & lsSlideName1    'Print Number: 1st
        lsOrder = lsOrder & lsPatInfo       'Print Number: 2nd
        lsOrder = lsOrder & lsSlideName     'Print Number: 3rd
        lsOrder = lsOrder & "100"           'Number of Films / Slide Glass
        

        'Patient ID
        lsOrder = lsOrder & "****************"
        
        'Sex(1), Date of Birht(8)
        lsOrder = lsOrder & "*********"
        
        'HCT(4), WBC(6), RBC(5)
        lsOrder = lsOrder & "***************" & chrETX
    
        Debug.Print lsOrder
        OrderOutput lsOrder
        
        SQL = "Select barcode from res_flag where examdate = '" & Format(Date, "yyyymmdd") & "' and Barcode = '" & lsID & "'"
        res = db_select_Col(gLocal, SQL)
        If Trim(gReadBuf(0)) = lsID Then
            SQL = "update res_flag set SlidOrd = '" & lsSlideOrd & "', SampleJudg = '0' " & vbCrLf & _
                  "where examdate = '" & Format(Date, "yyyymmdd") & "' and Barcode = '" & lsID & "'"
            res = SendQuery(gLocal, SQL)
        Else
            SQL = "Insert Into res_flag (examdate, barcode, SampleJudg, PosDiff, PosMorph, PosCnt, ErrFunc, ErrRes, WBCAbnor, WBCSusp, RBCAbnor, RBCSusp, PLTAbnor, PLTSusp, InfoWBC, InfoPLT, PBSFlag, SlideOrd ) " & vbCrLf & _
                  "Values ('" & Format(Date, "yyyymmdd") & "', '" & lsID & "', '0', '', '', '', " & _
                  "'', '', '', '', '', '', " & _
                  "'', '', '', '', '', '" & lsSlideOrd & "') "
            res = SendQuery(gLocal, SQL)
        End If

        SQL = "select barcode from worklist where barcode = '" & lsID & "' "
        res = db_select_Col(gLocal, SQL)
        If Trim(gReadBuf(0)) = Trim(lsID) Then
            SQL = "delete from worklist where barcode = '" & lsID & "' "
            res = SendQuery(gLocal, SQL)
        End If
        SQL = "Insert Into WorkList(ReceDate, ReceTime, Barcode, PID, PName, OrdFlag, OrdDateTime, ResDateTime, OrdCnt, RemoteIP, WkNo ) " & vbCrLf & _
              "Values ('" & Trim(GetText(vasList, lRow, 5)) & "', '" & Trim(GetText(vasList, lRow, 6)) & "', '" & lsID & "', '" & Trim(GetText(vasList, lRow, 3)) & "','" & Trim(GetText(vasList, lRow, 4)) & "', 'B','" & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "hh:nn:ss") & "','', 0, 'cbcworklist', '" & lsReceNo & "') "
        res = SendQuery(gLocal, SQL)
    Else
        SQL = "select barcode from worklist where barcode = '" & lsID & "' "
        res = db_select_Col(gLocal, SQL)
        If Trim(gReadBuf(0)) = Trim(lsID) Then
            SQL = "delete from worklist where barcode = '" & lsID & "' "
            res = SendQuery(gLocal, SQL)
        End If
        SQL = "Insert Into WorkList(ReceDate, ReceTime, Barcode, PID, PName, OrdFlag, OrdDateTime, ResDateTime, OrdCnt, RemoteIP, WkNo ) " & vbCrLf & _
              "Values ('" & Trim(GetText(vasList, lRow, 5)) & "', '" & Trim(GetText(vasList, lRow, 6)) & "', '" & lsID & "', '" & Trim(GetText(vasList, lRow, 3)) & "','" & Trim(GetText(vasList, lRow, 4)) & "', 'B','" & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "hh:nn:ss") & "','', 0, 'cbcworklist', '" & lsReceNo & "') "
        res = SendQuery(gLocal, SQL)
        
        Exit Function
    End If
        
    '2010.03.17 이상은 - TLA 조회 후 TLA Flag 업데이트***************
'    sParam = ""
'
'    sParam = sParam & CR & _
'            "<Table>" & CR & _
'            "<QID><![CDATA[PG_SRL.SLP91_U02]]></QID>" & CR & _
'            "<QTYPE><![CDATA[Package]]></QTYPE>" & CR & _
'            "<USERID><![CDATA[" & gServerID & "]]></USERID>" & CR & _
'            "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>" & CR & _
'            "<TABLENAME><![CDATA[]]></TABLENAME>" & CR & _
'            "<P0><![CDATA[" & lsID & "]]></P0>" & CR & _
'            "<P1><![CDATA[" & CDate(Date) & "]]></P1>" & CR & _
'            "<P2><![CDATA[" & CDate(Date) & "]]></P2>" & CR & _
'            "</Table>"
'    sParam = "<NewDataSet>" & sParam & "</NewDataSet>"
'
'    Online_TLA_Qry sParam
    
    '****************************************************************
    
    lRow1 = vasOrder.DataRowCnt + 1
    If lRow1 > vasOrder.MaxRows Then vasOrder.MaxRows = lRow1
    
    InsertRow vasOrder, 1

    vasOrder.SetText 1, 1, Trim(GetText(vasList, lRow, 1))
    vasOrder.SetText 2, 1, "B"
    vasOrder.SetText 3, 1, Trim(GetText(vasList, lRow, 3))
    vasOrder.SetText 4, 1, Trim(GetText(vasList, lRow, 4))
    vasOrder.SetText 5, 1, Trim(GetText(vasList, lRow, 5))
    vasOrder.SetText 6, 1, Trim(GetText(vasList, lRow, 6))
    vasOrder.SetText 7, 1, Trim(GetText(vasList, lRow, 7))
    vasOrder.SetText 8, 1, Trim(GetText(vasList, lRow, 8))
    
        
    DeleteRow vasList, lRow, lRow
End Function


Function OrderEntry_1_기존(asRow As Long) As Integer
    'Individual Order Format
    
    Dim lsID    As String
    Dim i, j, k As Integer
    Dim AdoRs_Exam As ADODB.Recordset
    
    Dim Ord(7)  As String
    
    Dim lsOrder     As String
    Dim lsOrder1    As String
    Dim lRow, lRow1 As Long
    Dim lsDate      As String
    
    Dim lsPatInfo As String
    Dim lsWkNo As String
    Dim lsPID As String
    Dim lsPName As String
    Dim lsPEName As String
    Dim lsPAge As String
    Dim lsPSex As String
    Dim lsPBirth As String
    Dim lsWard As String
    
    Dim lsSlideName1 As String
    Dim lsSlideOrd As String
    Dim lsExamDate As String
    
    Dim lsSlideName As String
    Dim iMOR As Integer
    Dim iPBS As Integer
    
    Dim iSP As Integer
    
    Dim lsReceNo    As String
    
    Dim sParam      As String
    
    lRow = asRow
    
    If lRow < 1 Or lRow > vasList.DataRowCnt Then Exit Function
    
    lsExamDate = Trim(Format(Date, "yyyymmdd"))
    
    SQL = "select equipcode, examcode, examname, OrdGubun, examno from equipexam where equip = '" & gEquip & "' and OrdGubun <> 'N' order by 2, 5 "
    Set AdoRs_Exam = db_select_rs(gLocal, SQL)
        
    lsSlideName = ""
    lsSlideName1 = ""
    
    lsOrder = ""
    For i = 1 To 7
        Ord(i) = "0"
    Next i
                
    lsSlideOrd = ""
        
    lsID = Trim(GetText(vasList, lRow, 1))
    
    lblBarCode.Caption = lsID
    
    If Trim(GetText(vasList, lRow, 2)) = "1" Then
        iSP = 1
    Else
        iSP = 0
    End If
    
    'lsWkNo = SetSpace(Trim(gReadBuf(1)), 5)
    lsPID = Trim(GetText(vasList, lRow, 3))
    lsPName = Trim(GetText(vasList, lRow, 4))
    
    lsSlideName = lsPName
    lsSlideName = Conv_Kor_Eng(lsSlideName)
    lsPEName = lsSlideName
    lsWkNo = Trim(GetText(vasList, lRow, 8))
    lsPBirth = Trim(GetText(vasList, lRow, 9))
    lsPSex = Trim(GetText(vasList, lRow, 10))
    lsWard = ""
    
    Select Case lsPSex
    Case "3", "4"
        lsPBirth = "20" & lsPBirth
    Case "1", "2", "5", "6", "7", "8", "9", "0"
        lsPBirth = "19" & lsPBirth
    Case Else
        lsPBirth = ""
    End Select
    
    If IsNumeric(lsPSex) Then
        If CInt(lsPSex) Mod 2 = 0 Then
            lsPSex = "F"
        Else
            lsPSex = "M"
        End If
    End If
    
    lsDate = GetDateFull

    
    ClearSpread vasTemp
    ClearSpread vasExam
    
    lsReceNo = lsWkNo
    
    'lsSlideName1 = Mid(lsExamDate, 3, 2) & "/" & Mid(lsExamDate, 5, 2) & "/" & Mid(lsExamDate, 7, 2)
    lsSlideName1 = lsExamDate
    
'    If vasExam.DataRowCnt > 0 Then
'        lsSlideName1 = lsSlideName1 & "  " & Trim(GetText(vasExam, 1, 3))
'    End If
    
    iMOR = -1
    iPBS = -1
    
    'res = Get_Order(lsID)
    res = Online_XML(gXml_S07, lsID)
    
    For j = 0 To UBound(gExam_Select)
        
        If Not AdoRs_Exam Is Nothing Then
            AdoRs_Exam.MoveFirst
            
            Debug.Print (GetText(vasExam, j + 1, 1)) & "  " & Trim(AdoRs_Exam("examcode"))
            
            Do Until AdoRs_Exam.EOF
                If gExam_Select(j).TST_CD <> "" Then
                    If Trim(AdoRs_Exam("examcode")) = gExam_Select(j).TST_CD Then
                        Select Case Trim(AdoRs_Exam("OrdGubun"))
                        Case "C": Ord(1) = "1"
                        Case "D": Ord(2) = "1"
                        Case "R"
                            Ord(3) = "1"
                        Case "P"
                            Ord(4) = "1"
                            lsSlideOrd = "SP"
                        Case "S"
                            Ord(5) = "1"
                            lsSlideOrd = "SC"
    
                        Case "X": Ord(6) = "1"
                        Case "B": Ord(7) = "1"
                        End Select
                        
    '                    If gExam_Select(j).TST_CD = "L2023" Or gExam_Select(j).TST_CD = "L20231" Or gExam_Select(j).TST_CD = "L20232" Then
    '                        iPBS = 1
    '                    End If
    
    '                        Case "CP0106"   'Morphology
    '                            lsSlideName = "MOR." & lsSlideName
    '                        Case "CP0131"   'PB Smear
    '                            lsSlideName = "PBS." & lsSlideName
    '                        End Select
                        Exit Do
                    End If
                End If
                
                AdoRs_Exam.MoveNext
            Loop
        End If
    Next j
        
    'If Ord(5) = "1" Then Ord(2) = "1"
    
    If iSP = 1 Then
        Ord(4) = "1"
        lsSlideOrd = "SP"
        SQL = "update res_flag set SampleJudg = '0' " & vbCrLf & _
              "where examdate = '" & Format(Date, "yyyymmdd") & "'  " & vbCrLf & _
              "  and barcode = '" & lsID & "' "
        res = SendQuery(gLocal, SQL)
    End If
    
    If iPBS = 1 Then
        lsSlideOrd = "SP"
        Ord(1) = "1"
        Ord(2) = "1"
    End If
    
    If Ord(4) = "1" Then Ord(5) = "0"
    
    lsOrder = ""
    For i = 1 To 7
        lsOrder = lsOrder & Ord(i)
    Next i
    
'    If iPBS = 1 Then
'        lsSlideName = "PBS." & lsSlideName
'    Else
'        If iMOR = 1 Then
'            lsSlideName = "MOR." & lsSlideName
'        End If
'    End If
    
    'MsgBox lsOrder
    
    If lsOrder <> "0000000" And lsOrder <> "" Then
        'DoSleep gOrdGap
        
        '20100127-206410012700336 ILee gi suk
        
        '환자이름
        lsSlideName = lsSlideName
        lsSlideName = Left(lsSlideName, 13)
        lsSlideName = SetChar(lsSlideName, 13, 2, " ")
        
        '바코드번호 외래/입원구분
        'lsPatInfo = lsPID
        lsPatInfo = lsID
        lsPatInfo = Left(lsPatInfo, 12)
        lsPatInfo = SetChar(lsPatInfo, 12, 2, " ") & "I"
        
        '작업일자-작업번호
        lsSlideName1 = lsSlideName1 & "-" & SetChar(lsWkNo, 4, 2, "")
        lsSlideName1 = Left(lsSlideName1, 13)
        lsSlideName1 = SetChar(lsSlideName1, 13, 1, " ")
        
        lsOrder1 = ""
        
        'MsgBox lsOrder
        
        'RET/SP/SC
        Select Case lsOrder
        Case "1000000"      'CBC
            lsOrder1 = "001" & "11111111000000000011111" & "00" & "000" & "0000100" & "00000000000000"
        Case "1000100"      'CBC+SC
            lsOrder1 = "001" & "11111111000000000011111" & "00" & "000" & "0000100" & "00000000000000"
            
        Case "1100000"      'CBC+Diff
            lsOrder1 = "001" & "11111111111111111111111" & "00" & "000" & "0000100" & "00000000000000"
        Case "1100100"      'CBC+Diff+SC
            lsOrder1 = "001" & "11111111111111111111111" & "00" & "000" & "0000100" & "00000000000000"
            
        Case "1110000"      'CBC+Diff+Reti
            lsOrder1 = "001" & "11111111111111111111111" & "00" & "111" & "1111100" & "00000000000000"
        Case "1110100"      'CBC+Diff+Reti+SC
            lsOrder1 = "001" & "11111111111111111111111" & "00" & "111" & "1111100" & "00000000000000"
            
        Case "0010000"
            lsOrder1 = "001" & "00000000000000000000000" & "00" & "111" & "1110000" & "00000000000000"
        End Select
        
        '동아대 작업 : Slide 이름에 (MOR)+환자영문이름 => 최대 자리수 잘라넣기
        '2006년 9월 29일 환자 정보 더 넣기
'            lsOrder = chrSTX & "S" & "00000000" & SetChar(lsID, 15, 1, " ") & "********" & lsOrder & "00000"
'            lsOrder = lsOrder & lsSlideName
'            lsOrder = lsOrder & "0000000000000"
'            lsOrder = lsOrder & "0000000000000"
'            lsOrder = lsOrder & "000****************************************" & chrETX
        
        lsOrder = chrSTX & "S" & "00000000" & SetChar(lsID, 15, 1, " ") & "********" & lsOrder1 & "00000"

        lsOrder = lsOrder & lsSlideName1
        lsOrder = lsOrder & lsPatInfo
        lsOrder = lsOrder & lsSlideName
        lsOrder = lsOrder & "100"
        
        'Patient ID
        'lsOrder = lsOrder & SetSpace(lsPID, 16, 1)
        
'        Select Case lsPSex
'        Case "M"
'            lsOrder = lsOrder & "1"
'        Case "F"
'            lsOrder = lsOrder & "2"
'        Case Else
'            lsOrder = lsOrder & "3"
'        End Select
'        lsOrder = lsOrder & SetSpace(Trim(lsPBirth), 8, 1)

        'Patient ID
        lsOrder = lsOrder & "****************"
        
        'Sex(1), Date of Birht(8)
        lsOrder = lsOrder & "*********"
        
        'HCT(4), WBC(6), RBC(5)
        lsOrder = lsOrder & "***************" & chrETX
    
        Debug.Print lsOrder
        OrderOutput lsOrder
        

'                MSComm1.Output = lsOrder
'                SaveOrdLog lsOrder
        
        SQL = "Select barcode from res_flag where examdate = '" & Format(Date, "yyyymmdd") & "' and Barcode = '" & lsID & "'"
        res = db_select_Col(gLocal, SQL)
        If Trim(gReadBuf(0)) = lsID Then
            SQL = "update res_flag set SlidOrd = '" & lsSlideOrd & "', SampleJudg = '0' " & vbCrLf & _
                  "where examdate = '" & Format(Date, "yyyymmdd") & "' and Barcode = '" & lsID & "'"
            res = SendQuery(gLocal, SQL)
        Else
            SQL = "Insert Into res_flag (examdate, barcode, SampleJudg, PosDiff, PosMorph, PosCnt, ErrFunc, ErrRes, WBCAbnor, WBCSusp, RBCAbnor, RBCSusp, PLTAbnor, PLTSusp, InfoWBC, InfoPLT, PBSFlag, SlideOrd ) " & vbCrLf & _
                  "Values ('" & Format(Date, "yyyymmdd") & "', '" & lsID & "', '0', '', '', '', " & _
                  "'', '', '', '', '', '', " & _
                  "'', '', '', '', '', '" & lsSlideOrd & "') "
            res = SendQuery(gLocal, SQL)
        End If

        SQL = "select barcode from worklist where barcode = '" & lsID & "' "
        res = db_select_Col(gLocal, SQL)
        If Trim(gReadBuf(0)) = Trim(lsID) Then
            SQL = "delete from worklist where barcode = '" & lsID & "' "
            res = SendQuery(gLocal, SQL)
        End If
        SQL = "Insert Into WorkList(ReceDate, ReceTime, Barcode, PID, PName, OrdFlag, OrdDateTime, ResDateTime, OrdCnt, RemoteIP, WkNo ) " & vbCrLf & _
              "Values ('" & Trim(GetText(vasList, lRow, 5)) & "', '" & Trim(GetText(vasList, lRow, 6)) & "', '" & lsID & "', '" & Trim(GetText(vasList, lRow, 3)) & "','" & Trim(GetText(vasList, lRow, 4)) & "', 'B','" & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "hh:nn:ss") & "','', 0, 'cbcworklist', '" & lsReceNo & "') "
        res = SendQuery(gLocal, SQL)
    Else
        SQL = "select barcode from worklist where barcode = '" & lsID & "' "
        res = db_select_Col(gLocal, SQL)
        If Trim(gReadBuf(0)) = Trim(lsID) Then
            SQL = "delete from worklist where barcode = '" & lsID & "' "
            res = SendQuery(gLocal, SQL)
        End If
        SQL = "Insert Into WorkList(ReceDate, ReceTime, Barcode, PID, PName, OrdFlag, OrdDateTime, ResDateTime, OrdCnt, RemoteIP, WkNo ) " & vbCrLf & _
              "Values ('" & Trim(GetText(vasList, lRow, 5)) & "', '" & Trim(GetText(vasList, lRow, 6)) & "', '" & lsID & "', '" & Trim(GetText(vasList, lRow, 3)) & "','" & Trim(GetText(vasList, lRow, 4)) & "', 'B','" & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "hh:nn:ss") & "','', 0, 'cbcworklist', '" & lsReceNo & "') "
        res = SendQuery(gLocal, SQL)
        
        Exit Function
    End If
        
    '2010.03.17 이상은 - TLA 조회 후 TLA Flag 업데이트***************
    sParam = ""

    sParam = sParam & CR & _
            "<Table>" & CR & _
            "<QID><![CDATA[PG_SRL.SLP91_U02]]></QID>" & CR & _
            "<QTYPE><![CDATA[Package]]></QTYPE>" & CR & _
            "<USERID><![CDATA[" & gServerID & "]]></USERID>" & CR & _
            "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>" & CR & _
            "<TABLENAME><![CDATA[]]></TABLENAME>" & CR & _
            "<P0><![CDATA[" & lsID & "]]></P0>" & CR & _
            "<P1><![CDATA[" & CDate(Date) & "]]></P1>" & CR & _
            "<P2><![CDATA[" & CDate(Date) & "]]></P2>" & CR & _
            "</Table>"
    sParam = "<NewDataSet>" & sParam & "</NewDataSet>"

    Online_TLA_Qry sParam
    
    '****************************************************************
    
    lRow1 = vasOrder.DataRowCnt + 1
    If lRow1 > vasOrder.MaxRows Then vasOrder.MaxRows = lRow1
    
    InsertRow vasOrder, 1

    vasOrder.SetText 1, 1, Trim(GetText(vasList, lRow, 1))
    vasOrder.SetText 2, 1, "B"
    vasOrder.SetText 3, 1, Trim(GetText(vasList, lRow, 3))
    vasOrder.SetText 4, 1, Trim(GetText(vasList, lRow, 4))
    vasOrder.SetText 5, 1, Trim(GetText(vasList, lRow, 5))
    vasOrder.SetText 6, 1, Trim(GetText(vasList, lRow, 6))
    vasOrder.SetText 7, 1, Trim(GetText(vasList, lRow, 7))
    vasOrder.SetText 8, 1, Trim(GetText(vasList, lRow, 8))
    
        
    DeleteRow vasList, lRow, lRow
End Function

Sub OrderOutput(asOrder As String)
    If gSetup.Protocol = "B" Then gbOrdering = True

    DoSleep 5
    MSComm1.Output = asOrder
            
    SaveOrdLog asOrder
End Sub

Private Sub txtReOrd_GotFocus()
    SelectFocus txtReOrd
End Sub

Private Sub txtReOrd_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lsID As String
    Dim i, j, k As Integer
    Dim AdoRs_Exam As ADODB.Recordset
    Dim Ord(7) As String
    Dim lsOrder As String
    Dim lRow As Long
    Dim lsDate As String
    
    Dim lsWkNo As String
    Dim lsPID As String
    Dim lsPName As String
    
    Dim lsSlideOrd As String
    Dim lsExamDate As String
    Dim lsSlideName As String
    
    If KeyCode = vbKeyReturn Then
        If Trim(txtReOrd) = "" Then
            txtReOrd.SetFocus
            Exit Sub
        End If
        
        lsSlideName = ""
        
        If MSComm1.CTSHolding = False Then
            lblMsg.Caption = "[Message] LASC 의 포트가 준비되지 않았습니다"
            lblMsg.ForeColor = RGB(255, 0, 0)
        Else
            lblMsg.Caption = "[Message] LASC 의 포트가 준비되었습니다"
            lblMsg.ForeColor = RGB(0, 0, 0)
        End If
        
        If MSComm1.PortOpen = False Then
            LASCPortOpen
        End If

        If MSComm1.CTSHolding = False Then
            lblMsg.Caption = "[Message] LASC 의 포트가 준비되지 않았습니다"
            lblMsg.ForeColor = RGB(255, 0, 0)

            Exit Sub
        Else
            lblMsg.Caption = "[Message] LASC 의 포트가 준비되었습니다"
            lblMsg.ForeColor = RGB(0, 0, 0)
        End If
        
        lsExamDate = Format(CDate(GetDateFull), "yyyymmdd")
                
        lsID = Trim(txtReOrd)
        lRow = vasList.DataRowCnt + 1
        
        With vasList
            .SetText 1, lRow, lsID
            .SetText 2, lRow, "1"
            
            'res = Get_PatInfo(lsID)
            res = Online_XML(gXml_S03, lsID)
            If res > 0 Then
                .SetText 3, lRow, gPat_Info_Select.PT_NO
                .SetText 4, lRow, gPat_Info_Select.PT_NM
                .SetText 5, lRow, gPat_Info_Select.ACPT_DTETM     '날짜
                .SetText 6, lRow, ""    '시간
                .SetText 7, lRow, "" '접수코드
                .SetText 8, lRow, gPat_Info_Select.ACPTNO_1
                .SetText 9, lRow, gPat_Info_Select.Sex
                .SetText 10, lRow, gPat_Info_Select.Age
                .SetText 11, lRow, ""   'slip
            End If
        End With
                        
        OrderEntry_1 lRow
        
        txtReOrd = ""
        txtReOrd.SetFocus
        
        Exit Sub
            
    End If
    
    
    Exit Sub
    
ErrHandle:
    SaveQuery "[개별전송]" & Err.Number & ": " & Err.Description
    Resume Next

End Sub

Private Sub vasList_Click(ByVal Col As Long, ByVal Row As Long)
    If Row = 0 Then
        If Col > 0 Then
            vasSort vasList, Col
        End If
    End If
End Sub

Private Sub vasList_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim lsID    As String
    Dim lRow    As Integer
    Dim i       As Integer
    Dim mExam   As Variant
    
    If Row < 1 Or Row > vasList.DataRowCnt Then Exit Sub
    
    ClearSpread vasExam
    
    lsID = Trim(GetText(vasList, Row, 1))
    lblBarCode.Caption = lsID
'    res = Get_Order(lsID)
'    For i = 0 To UBound(gOrder_List)
'        vasExam.SetText 1, i + 1, gOrder_List(i).TST_CD
'    Next i

    res = Online_XML(gXml_S07, lsID)
    For i = 0 To UBound(gExam_Select)
        vasExam.SetText 1, i + 1, gExam_Select(i).TST_CD
        
        SQL = " Select examname From equipexam Where examcode = '" & gExam_Select(i).TST_CD & "' "
        res = db_select_Col(gLocal, SQL)
        vasExam.SetText 2, i + 1, Trim(gReadBuf(0))
    Next i
End Sub

Private Sub vasList_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lRow As Long
    
    If KeyCode = vbKeyDelete Then
        lRow = vasList.ActiveRow
        
        If lRow < 1 Or lRow > vasList.DataRowCnt Then Exit Sub
            
        If MsgBox("검체코드 " & Trim(GetText(vasList, lRow, 1)) & " " & _
                  Trim(GetText(vasList, lRow, 4)) & " 검체를 삭제하시겠습니까? ", vbCritical + vbYesNo, "알림") = vbNo Then
            Exit Sub
        End If
        
        SQL = "Delete from worklist where barcode = '" & Trim(GetText(vasList, lRow, 1)) & "' "
        res = SendQuery(gLocal, SQL)
        If res = 1 Then
            DeleteRow vasList, lRow, lRow
        End If
    End If
End Sub

Private Sub vasOrder_Click(ByVal Col As Long, ByVal Row As Long)
    If Row = 0 Then
        If Col > 0 Then
            vasSort vasOrder, Col
        End If
    End If
End Sub

Private Sub vasOrder_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim lsID As String
    Dim i, k As Integer
    
    If Row < 1 Or Row > vasOrder.DataRowCnt Then Exit Sub
    
    lsID = Trim(GetText(vasOrder, Row, 1))
    
    ClearSpread vasExam
    
    lblBarCode.Caption = lsID

'    res = Get_Order(lsID)
'    For i = 0 To UBound(gOrder_List)
'        vasExam.SetText 1, i + 1, gOrder_List(i).TST_CD
'    Next i

    res = Online_XML(gXml_S07, lsID)
    For i = 0 To UBound(gExam_Select)
        vasExam.SetText 1, i + 1, gExam_Select(i).TST_CD
        
        SQL = " Select examname From equipexam Where examcode = '" & gExam_Select(i).TST_CD & "' "
        res = db_select_Col(gLocal, SQL)
        vasExam.SetText 2, i + 1, Trim(gReadBuf(0))
    Next i
    
    
End Sub

Private Sub vasOrder_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lRow As Long
    
    If KeyCode = vbKeyDelete Then
        lRow = vasOrder.ActiveRow
        
        If lRow < 1 Or lRow > vasOrder.DataRowCnt Then Exit Sub
            
        If MsgBox("검체코드 " & Trim(GetText(vasOrder, lRow, 1)) & " " & _
                  Trim(GetText(vasOrder, lRow, 4)) & " 검체를 삭제하시겠습니까? ", vbCritical + vbYesNo, "알림") = vbNo Then
            Exit Sub
        End If
        
        SQL = "Delete from worklist where barcode = '" & Trim(GetText(vasOrder, lRow, 1)) & "' "
        res = SendQuery(gLocal, SQL)
        If res = 1 Then
            DeleteRow vasOrder, lRow, lRow
        End If
    End If
End Sub

Function GetSetup() As Boolean
'---------------------------------------------------------------------------------------------------------------------
'                       Setup  File을 읽어온다.
'---------------------------------------------------------------------------------------------------------------------
    Dim db_tmp As String * 100

    db_tmp = ""
    
    GetSetup = False
    
    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "driver", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    gDB_Parm.Driver = Trim(txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "uid", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    gDB_Parm.User = Trim(txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "pwd", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    gDB_Parm.Passwd = Trim(txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "server", "", db_tmp, 100, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    gDB_Parm.Server = Trim(txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "database", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    gDB_Parm.DB = Trim(txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "hostname", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    gDB_Parm.HostName = Trim(txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("Data", "WorkListExpire", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    gExpireDate = Trim(txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("Option", "Timer", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    gTimer = Trim(txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("Option", "Order_Gap", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    gOrdGap = Trim(txtTemp)
    If Not IsNumeric(gOrdGap) Then
        gOrdGap = 10
    End If
    
    '2010.01.15 이상은
    db_tmp = ""
    Call GetPrivateProfileString("Server", "ServerPath", "", db_tmp, 100, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    gServerPath = Trim(txtTemp)
    
    '2010.04.14 이상은
    db_tmp = ""
    Call GetPrivateProfileString("Server", "ServerID", "", db_tmp, 100, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    gServerID = Trim(txtTemp)
    
    GetSetup = True

End Function

Public Sub GetSetup_LASC()
    Dim db_tmp As String * 20
    Dim i As Integer
    Dim lRow As Long
       
    lRow = 0
    For i = 1 To 10
        db_tmp = ""
        Call GetPrivateProfileString("COM " & CStr(i), "Use", "", db_tmp, 20, App.Path & "\Interface.ini")
        txtTemp = Trim(db_tmp)
        If Trim(txtTemp) <> "" Then
'            lRow = lRow + 1
'
'            vasComList.Row = lRow
'            vasComList.Col = 1
'            If Trim(txtTemp) = "1" Then
'                vasComList.Value = 1
'            Else
'                vasComList.Value = 0
'            End If
            
            db_tmp = ""
            Call GetPrivateProfileString("COM " & CStr(i), "Gubun", "", db_tmp, 20, App.Path & "\Interface.ini")
            txtTemp = Trim(db_tmp)
            
            If Left(Trim(txtTemp), 4) = "LASC" Then
                
                gSetup.Port = i
                
                db_tmp = ""
                Call GetPrivateProfileString("COM " & CStr(i), "Speed", "", db_tmp, 20, App.Path & "\Interface.ini")
                txtTemp = Trim(db_tmp)
                gSetup.Speed = Trim(txtTemp)
            
                db_tmp = ""
                Call GetPrivateProfileString("COM " & CStr(i), "Parity", "", db_tmp, 20, App.Path & "\Interface.ini")
                txtTemp = Trim(db_tmp)
                gSetup.Parity = Trim(txtTemp)
            
                db_tmp = ""
                Call GetPrivateProfileString("COM " & CStr(i), "DataBit", "", db_tmp, 20, App.Path & "\Interface.ini")
                txtTemp = Trim(db_tmp)
                gSetup.DataBit = Trim(txtTemp)
            
                db_tmp = ""
                Call GetPrivateProfileString("COM " & CStr(i), "StopBit", "", db_tmp, 20, App.Path & "\Interface.ini")
                txtTemp = Trim(db_tmp)
                gSetup.StopBit = Trim(txtTemp)
            
                db_tmp = ""
                Call GetPrivateProfileString("COM " & CStr(i), "RTSEnable", "", db_tmp, 20, App.Path & "\Interface.ini")
                txtTemp = Trim(db_tmp)
                gSetup.RTSEnable = Trim(txtTemp)
            
                db_tmp = ""
                Call GetPrivateProfileString("COM " & CStr(i), "DTREnable", "", db_tmp, 20, App.Path & "\Interface.ini")
                txtTemp = Trim(db_tmp)
                gSetup.DTREnable = Trim(txtTemp)
            
            
                db_tmp = ""
                Call GetPrivateProfileString("COM " & CStr(i), "Protocol", "", db_tmp, 20, App.Path & "\Interface.ini")
                txtTemp = Trim(db_tmp)
                gSetup.Protocol = Trim(txtTemp)
            
            End If
        End If
    Next i

End Sub

Public Function STS(ByVal strStmt As String) As String
    Dim strTmp As String
    
    strTmp = Replace(strStmt, "'", "''")
    
    STS = "'" & strTmp & "'"
End Function


Public Sub SaveWinsockLog(argSQL As String)
'argSQL의 내용을 파일로 저장
    Dim FilNum
    
    FilNum = FreeFile
    
    If Dir(App.Path & "\Log", vbDirectory) = "" Then
        MkDir App.Path & "\Log"
    End If
    
    Open App.Path & "\Log\Socket.log" For Append As FilNum
    Print #FilNum, Time & " " & argSQL
    Close FilNum
End Sub

Public Sub SaveOrdLog(argSQL As String)
'argSQL의 내용을 파일로 저장
    Dim FilNum
    
    FilNum = FreeFile
    
    If Dir(App.Path & "\Log", vbDirectory) = "" Then
        MkDir App.Path & "\Log"
    End If
    
    Open App.Path & "\Log\Ord" & Date & ".log" For Append As FilNum
    Print #FilNum, Time & " " & argSQL
    Close FilNum
End Sub

