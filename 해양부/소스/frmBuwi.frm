VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmBuwi 
   Caption         =   "해양관측부이"
   ClientHeight    =   11565
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   21945
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11565
   ScaleWidth      =   21945
   Begin VB.Frame fra1 
      BackColor       =   &H00FFFFFF&
      Height          =   10065
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   20355
      Begin VB.Frame fraSearch 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   675
         Left            =   6540
         TabIndex        =   8
         Top             =   120
         Width           =   13695
         Begin VB.ComboBox cboToHour 
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            ItemData        =   "frmBuwi.frx":0000
            Left            =   7470
            List            =   "frmBuwi.frx":0002
            Style           =   2  '드롭다운 목록
            TabIndex        =   13
            Top             =   210
            Width           =   795
         End
         Begin VB.CommandButton cmdSearch2 
            BackColor       =   &H00E0E0E0&
            Caption         =   "조회"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8370
            Style           =   1  '그래픽
            TabIndex        =   12
            Top             =   180
            Width           =   1095
         End
         Begin VB.ComboBox cboSearchCount 
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            ItemData        =   "frmBuwi.frx":0004
            Left            =   11490
            List            =   "frmBuwi.frx":0006
            Style           =   2  '드롭다운 목록
            TabIndex        =   11
            Top             =   240
            Width           =   1005
         End
         Begin VB.ComboBox cboFromHour 
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            ItemData        =   "frmBuwi.frx":0008
            Left            =   4980
            List            =   "frmBuwi.frx":000A
            Style           =   2  '드롭다운 목록
            TabIndex        =   10
            Top             =   210
            Width           =   795
         End
         Begin VB.ComboBox cboVPNList2 
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            ItemData        =   "frmBuwi.frx":000C
            Left            =   870
            List            =   "frmBuwi.frx":000E
            Style           =   2  '드롭다운 목록
            TabIndex        =   9
            Top             =   210
            Width           =   2025
         End
         Begin MSComCtl2.DTPicker dtpFromDate 
            Height          =   345
            Left            =   3540
            TabIndex        =   14
            Top             =   210
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   43319297
            CurrentDate     =   43884
         End
         Begin MSComCtl2.DTPicker dtpToDate 
            Height          =   345
            Left            =   6060
            TabIndex        =   15
            Top             =   210
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   43319297
            CurrentDate     =   43884
         End
         Begin VB.Label Label6 
            Alignment       =   2  '가운데 맞춤
            BackStyle       =   0  '투명
            Caption         =   "~"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   5820
            TabIndex        =   19
            Top             =   270
            Width           =   195
         End
         Begin VB.Label Label5 
            BackStyle       =   0  '투명
            Caption         =   "출력건수"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   10620
            TabIndex        =   18
            Top             =   330
            Width           =   975
         End
         Begin VB.Label Label4 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "기간"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   3000
            TabIndex        =   17
            Top             =   270
            Width           =   435
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '투명
            Caption         =   "관측소"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   180
            TabIndex        =   16
            Top             =   270
            Width           =   585
         End
      End
      Begin VB.Frame fraTimer 
         BackColor       =   &H00FFFFFF&
         Height          =   675
         Left            =   120
         TabIndex        =   1
         Top             =   9240
         Width           =   6345
         Begin VB.Timer tmrResult 
            Enabled         =   0   'False
            Interval        =   100
            Left            =   9000
            Top             =   120
         End
         Begin VB.CommandButton cmdSave 
            BackColor       =   &H00E0E0E0&
            Caption         =   "설정 저장"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   4740
            Style           =   1  '그래픽
            TabIndex        =   7
            Top             =   210
            Width           =   1395
         End
         Begin VB.TextBox txtInterval60 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   7830
            MaxLength       =   10
            TabIndex        =   6
            Top             =   180
            Visible         =   0   'False
            Width           =   750
         End
         Begin VB.ComboBox cboIntervalGrade 
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            ItemData        =   "frmBuwi.frx":0010
            Left            =   1950
            List            =   "frmBuwi.frx":0012
            Style           =   2  '드롭다운 목록
            TabIndex        =   5
            Top             =   210
            Width           =   1005
         End
         Begin VB.CheckBox chkRefresh 
            BackColor       =   &H00FFFFFF&
            Caption         =   "자동갱신"
            Height          =   255
            Left            =   150
            TabIndex        =   4
            Top             =   270
            Width           =   1065
         End
         Begin VB.TextBox txtInterval 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1290
            MaxLength       =   10
            TabIndex        =   3
            Top             =   210
            Width           =   630
         End
         Begin VB.CommandButton cmdSearch 
            BackColor       =   &H00E0E0E0&
            Caption         =   "조회"
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   3060
            Style           =   1  '그래픽
            TabIndex        =   2
            Top             =   210
            Width           =   1395
         End
      End
      Begin FPSpread.vaSpread spdVPNList2 
         Height          =   8985
         Left            =   120
         TabIndex        =   20
         Top             =   210
         Width           =   6345
         _Version        =   393216
         _ExtentX        =   11192
         _ExtentY        =   15849
         _StockProps     =   64
         ColsFrozen      =   2
         DisplayRowHeaders=   0   'False
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         GridColor       =   15921919
         GridShowVert    =   0   'False
         MaxCols         =   2
         MaxRows         =   20
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         ScrollBarExtMode=   -1  'True
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   16773087
         SpreadDesigner  =   "frmBuwi.frx":0014
         ScrollBarTrack  =   3
         ShowScrollTips  =   3
      End
      Begin FPSpread.vaSpread spdBUWIList 
         Height          =   9075
         Left            =   6540
         TabIndex        =   21
         Top             =   840
         Width           =   13695
         _Version        =   393216
         _ExtentX        =   24156
         _ExtentY        =   16007
         _StockProps     =   64
         ColsFrozen      =   2
         DisplayRowHeaders=   0   'False
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "맑은 고딕"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         GridColor       =   15921919
         GridShowVert    =   0   'False
         MaxCols         =   2
         MaxRows         =   20
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         ScrollBarExtMode=   -1  'True
         SelectBlockOptions=   0
         ShadowColor     =   16773087
         SpreadDesigner  =   "frmBuwi.frx":045E
         ScrollBarTrack  =   3
         ShowScrollTips  =   3
      End
   End
End
Attribute VB_Name = "frmBuwi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim intOneMinute As Integer

'-- 컨트롤초기화
Private Sub CtlInitializing()
    Dim iHour   As Integer
    
    With spdVPNList2
        .MaxCols = 3
        .MaxRows = 0
        
        Call SetText(spdVPNList2, "관측소ID", 0, 1):         .ColWidth(1) = 0
        Call SetText(spdVPNList2, "관측소", 0, 2):           .ColWidth(2) = 20
        Call SetText(spdVPNList2, "관측시간", 0, 3):         .ColWidth(3) = 30
    End With
    
    '------------------------------------------------------------------------
    With spdBUWIList
        .MaxCols = 29
        .MaxRows = 0
        
        Call SetText(spdBUWIList, "관측소ID", 0, 1):        .ColWidth(1) = 0
        Call SetText(spdBUWIList, "관측소명", 0, 2):        .ColWidth(2) = 10
        Call SetText(spdBUWIList, "경도", 0, 3):            .ColWidth(3) = 10
        Call SetText(spdBUWIList, "위도", 0, 4):            .ColWidth(4) = 10
        Call SetText(spdBUWIList, "관측시간", 0, 5):        .ColWidth(5) = 10
        Call SetText(spdBUWIList, "등록일자", 0, 6):        .ColWidth(6) = 10
        Call SetText(spdBUWIList, "풍속", 0, 7):            .ColWidth(7) = 10
        Call SetText(spdBUWIList, "풍향", 0, 8):            .ColWidth(8) = 10
        Call SetText(spdBUWIList, "돌풍(최대풍속)", 0, 9):  .ColWidth(9) = 10
        Call SetText(spdBUWIList, "기온", 0, 10):           .ColWidth(10) = 10
        Call SetText(spdBUWIList, "기압", 0, 11):           .ColWidth(11) = 10
        Call SetText(spdBUWIList, "부이방향", 0, 12):       .ColWidth(12) = 10
        Call SetText(spdBUWIList, "파고", 0, 13):           .ColWidth(13) = 10
        Call SetText(spdBUWIList, "파주기", 0, 14):         .ColWidth(14) = 10
        Call SetText(spdBUWIList, "유속", 0, 15):           .ColWidth(15) = 10
        Call SetText(spdBUWIList, "유향", 0, 16):           .ColWidth(16) = 10
        Call SetText(spdBUWIList, "수온", 0, 17):           .ColWidth(17) = 10
        Call SetText(spdBUWIList, "전도도", 0, 18):         .ColWidth(18) = 10
        Call SetText(spdBUWIList, "염분", 0, 19):           .ColWidth(19) = 10
        Call SetText(spdBUWIList, "장비ID", 0, 20):         .ColWidth(20) = 10
        Call SetText(spdBUWIList, "장비풍향", 0, 21):       .ColWidth(21) = 10
        Call SetText(spdBUWIList, "시정", 0, 22):           .ColWidth(22) = 10
        Call SetText(spdBUWIList, "배터리", 0, 23):         .ColWidth(23) = 10
        Call SetText(spdBUWIList, "레퍼런스", 0, 24):       .ColWidth(24) = 10
        Call SetText(spdBUWIList, "추적번호", 0, 25):       .ColWidth(25) = 10
        Call SetText(spdBUWIList, "최대파주기", 0, 26):     .ColWidth(26) = 10
        Call SetText(spdBUWIList, "최대파고", 0, 27):       .ColWidth(27) = 10
        Call SetText(spdBUWIList, "파향", 0, 28):           .ColWidth(28) = 10
        Call SetText(spdBUWIList, "풍향원본", 0, 29):       .ColWidth(29) = 10
        
    End With
    
    
    cboIntervalGrade.Clear
    cboIntervalGrade.AddItem "초"
    cboIntervalGrade.AddItem "분"
    cboIntervalGrade.ListIndex = 0
    
    dtpFromDate.Value = Now
    dtpToDate.Value = Now

    For iHour = 1 To 24
        cboFromHour.AddItem iHour
        cboToHour.AddItem iHour
    Next
    
    cboFromHour.ListIndex = 0
    cboToHour.ListIndex = 23
    
    cboSearchCount.Clear
    For iHour = 1 To 10
        cboSearchCount.AddItem iHour * 10
    Next
    cboSearchCount.ListIndex = 0
    
    gSORT = 0

End Sub

Private Sub chkRefresh_Click()

    If chkRefresh.Value = "1" Then
        tmrResult.Interval = 1000
        tmrResult.Enabled = True
    Else
        tmrResult.Enabled = False
    End If
    
End Sub

Private Sub cmdSave_Click()
    
    If chkRefresh.Value = "1" Then
        Call WritePrivateProfileString("USER", "AUTOREFREH", "1", App.PATH & "\MARINE.ini")
    Else
        Call WritePrivateProfileString("USER", "AUTOREFREH", "0", App.PATH & "\MARINE.ini")
    End If
    
    Call WritePrivateProfileString("USER", "INTERVAL", txtInterval.Text, App.PATH & "\MARINE.ini")
    Call WritePrivateProfileString("USER", "INTERGBN", cboIntervalGrade.Text, App.PATH & "\MARINE.ini")

End Sub

Private Sub cmdSearch_Click()
    
    If cn_Server_Flag = True Then
        '부이관측소-리스트 가져오기
        Call GetVPNList2("")
        
    End If

End Sub

'조위관측소 상세자료가져오기
Private Sub cmdSearch2_Click()
        
    Call GetBOUYList_Detail(mGetP(cboVPNList2.Text, 2, "|"), dtpFromDate.Value, cboFromHour.Text, dtpToDate.Value, cboToHour.Text, cboSearchCount.Text)

End Sub

Private Sub GetBOUYList_Detail(Optional ByVal pDT_TS_ID As String, _
                            Optional ByVal pDT_F_TIME As String, _
                            Optional ByVal pDT_F_HOUR As String, _
                            Optional ByVal pDT_T_TIME As String, _
                            Optional ByVal pDT_T_HOUR As String, _
                            Optional ByVal pCount As Integer)
    
    Dim pAdoRS  As ADODB.Recordset
    Dim intRow  As Integer
    
    Set pAdoRS = Get_BUOYList_Detail(pDT_TS_ID, pDT_F_TIME, pDT_F_HOUR, pDT_T_TIME, pDT_T_HOUR, pCount)
    
    spdBUWIList.MaxRows = 0

    If pAdoRS Is Nothing Then
        '등록된 정보 없음
    Else
        With spdBUWIList
            Do Until pAdoRS.EOF
                .MaxRows = .MaxRows + 1
                intRow = .MaxRows
    'SQL = SQL & "SELECT a.STATION_ID, a.OBS_TIME, b.STATION_NAME, b.EQUIP_ID        " & vbCrLf
                
                
'''                Call SetText(spdBUWIList, pAdoRS.Fields("STATION_ID").Value & "", intRow, 1)      '관측소ID
'''                Call SetText(spdBUWIList, pAdoRS.Fields("STATION_NAME").Value & "", intRow, 2)       '관측소명
'''                Call SetText(spdBUWIList, pAdoRS.Fields("OBS_TIME").Value & "", intRow, 3)       '관측시간
'''                Call SetText(spdBUWIList, pAdoRS.Fields("REG_DATE").Value & "", intRow, 4)      '로그기록시간
'''                Call SetText(spdBUWIList, pAdoRS.Fields("LOG_CONTENT").Value & "", intRow, 5)   '로그내용
                
                Call SetText(spdBUWIList, pAdoRS.Fields("STATION_ID").Value & "", intRow, 1)      '관측소ID
                Call SetText(spdBUWIList, pAdoRS.Fields("STATION_NAME").Value & "", intRow, 2)       '관측소명
                Call SetText(spdBUWIList, pAdoRS.Fields("BUOY_LON").Value & "", intRow, 3)       '
                Call SetText(spdBUWIList, pAdoRS.Fields("BUOY_LAT").Value & "", intRow, 4)       '
                Call SetText(spdBUWIList, Format(pAdoRS.Fields("OBS_TIME").Value, "YYYY/MM/DD hh:mm:ss"), intRow, 5)        '
                Call SetText(spdBUWIList, Format(pAdoRS.Fields("REG_TIME").Value, "YYYY/MM/DD hh:mm:ss"), intRow, 6)
                Call SetText(spdBUWIList, pAdoRS.Fields("WIND_S").Value & "", intRow, 7)
                Call SetText(spdBUWIList, pAdoRS.Fields("WIND_D").Value & "", intRow, 8)
                Call SetText(spdBUWIList, pAdoRS.Fields("WIND_G").Value & "", intRow, 9)
                Call SetText(spdBUWIList, pAdoRS.Fields("AIR_TEMP").Value & "", intRow, 10)
                Call SetText(spdBUWIList, pAdoRS.Fields("AIR_PRES").Value & "", intRow, 11)
                Call SetText(spdBUWIList, pAdoRS.Fields("BUOY_ORIENTATION").Value & "", intRow, 12)
                Call SetText(spdBUWIList, pAdoRS.Fields("WH").Value & "", intRow, 13)
                Call SetText(spdBUWIList, pAdoRS.Fields("WP").Value & "", intRow, 14)
                Call SetText(spdBUWIList, pAdoRS.Fields("CURRENT_S").Value & "", intRow, 15)
                Call SetText(spdBUWIList, pAdoRS.Fields("CURRENT_D").Value & "", intRow, 16)
                Call SetText(spdBUWIList, pAdoRS.Fields("W_TEMP").Value & "", intRow, 17)
                Call SetText(spdBUWIList, pAdoRS.Fields("CONDUCTIVITY").Value & "", intRow, 18)
                Call SetText(spdBUWIList, pAdoRS.Fields("SAL").Value & "", intRow, 19)
                Call SetText(spdBUWIList, pAdoRS.Fields("EQUIP_ID").Value & "", intRow, 20)
                Call SetText(spdBUWIList, pAdoRS.Fields("WIND_D_RAW").Value & "", intRow, 21)
                Call SetText(spdBUWIList, pAdoRS.Fields("VISIBILITY").Value & "", intRow, 22)
                Call SetText(spdBUWIList, pAdoRS.Fields("BATTERY").Value & "", intRow, 23)
                Call SetText(spdBUWIList, pAdoRS.Fields("REFERENCE").Value & "", intRow, 24)
                Call SetText(spdBUWIList, pAdoRS.Fields("TRACK_SEQ").Value & "", intRow, 25)
                Call SetText(spdBUWIList, pAdoRS.Fields("MAX_WAVE_PERIOD").Value & "", intRow, 26)
                Call SetText(spdBUWIList, pAdoRS.Fields("MAX_WAVE_HEIGHT").Value & "", intRow, 27)
                Call SetText(spdBUWIList, pAdoRS.Fields("WAVE_DIRECT").Value & "", intRow, 28)
                Call SetText(spdBUWIList, pAdoRS.Fields("ORIGINAL_WIND_D").Value & "", intRow, 29)
                pAdoRS.MoveNext
            Loop
        End With
    End If
    
    pAdoRS.Close

End Sub


Private Sub Form_Load()

    Call CtlInitializing
    
    txtInterval.Text = gInterVal
    cboIntervalGrade.Text = gInterGbn
    If gAutoRefresh = "1" Then
        chkRefresh.Value = "1"
    Else
        chkRefresh.Value = "0"
    End If
    
'    If gInterGbn = "분" Then
'        txtInterval60.Visible = True
'    Else
'        txtInterval60.Visible = False
'    End If
    
    tmrResult.Interval = 1000
    tmrResult.Enabled = True

    intOneMinute = 0
    
    If cn_Server_Flag = True Then
        '부이관측소-리스트 가져오기
        Call GetVPNList2("")
        
        Call GetVPNList2_Combo(cboVPNList2, "")
    
    End If
    
End Sub

Private Sub GetVPNList2(Optional ByVal pSTATION_ID As String)
    Dim pAdoRS  As ADODB.Recordset
    Dim intRow  As Integer
    Dim str10Min    As String
    Dim str24Hour   As String
    Dim strData(3)  As String
    
    '10분전 시간 반환
    str10Min = DateAdd("n", -10, frmMDI.dtpTotime.Value)
    str10Min = Format(str10Min, "YYYY/MM/DD hh:mm:ss")
    '24시간전 시간 반환
    str24Hour = DateAdd("h", -24, frmMDI.dtpToday.Value)
    str24Hour = Format(str24Hour, "YYYY/MM/DD hh:mm:ss")
    
    Set pAdoRS = Get_VPNList2(pSTATION_ID)
    
    spdVPNList2.MaxRows = 0
    
    If pAdoRS Is Nothing Then
        '등록된 정보 없음
    Else
        With spdVPNList2
            Do Until pAdoRS.EOF
                .MaxRows = .MaxRows + 1
                intRow = .MaxRows
                Call SetText(spdVPNList2, pAdoRS.Fields("STATION_ID").Value & "", intRow, 1)
                Call SetText(spdVPNList2, pAdoRS.Fields("STATION_NAME").Value & "", intRow, 2)
                Call SetText(spdVPNList2, Format(pAdoRS.Fields("OBS_TIME").Value, "YYYY/MM/DD hh:mm:ss"), intRow, 3)
                pAdoRS.MoveNext
            Loop
        End With
    End If
    pAdoRS.Close

    With spdVPNList2
        For intRow = 1 To .MaxRows
            If GetText(spdVPNList2, intRow, 3) <= str10Min Then
                strData(0) = GetText(spdVPNList2, intRow, 1)
                strData(1) = GetText(spdVPNList2, intRow, 2)
                strData(2) = GetText(spdVPNList2, intRow, 3)
                
                Call DeleteRow(spdVPNList2, intRow, intRow)
                Call InsertRow(spdVPNList2, 1, False)
                
                Call SetText(spdVPNList2, strData(0), 1, 1)
                Call SetText(spdVPNList2, strData(1), 1, 2)
                Call SetText(spdVPNList2, strData(2), 1, 3)
                
                Call SetBackColor(spdVPNList2, 1, 1, 1, 3, 255, 255, 0)
            End If
        Next
        For intRow = 1 To .MaxRows
            If GetText(spdVPNList2, intRow, 3) <= str24Hour Then
                strData(0) = GetText(spdVPNList2, intRow, 1)
                strData(1) = GetText(spdVPNList2, intRow, 2)
                strData(2) = GetText(spdVPNList2, intRow, 3)
                
                Call DeleteRow(spdVPNList2, intRow, intRow)
                Call InsertRow(spdVPNList2, 1, False)
                
                Call SetText(spdVPNList2, strData(0), 1, 1)
                Call SetText(spdVPNList2, strData(1), 1, 2)
                Call SetText(spdVPNList2, strData(2), 1, 3)
                
                Call SetBackColor(spdVPNList2, 1, 1, 1, 3, 255, 0, 0)
            End If
        Next
    End With
End Sub

Private Sub GetVPNList2_Combo(ByVal obj As Object, Optional ByVal pSTATION_ID As String)
    Dim pAdoRS  As ADODB.Recordset
    
    Set pAdoRS = Get_VPNList2(pSTATION_ID, True)
    
    If pAdoRS Is Nothing Then
        '등록된 정보 없음
    Else
        obj.Clear
        Do Until pAdoRS.EOF
            obj.AddItem pAdoRS.Fields("STATION_NAME").Value & Space(30) & "|" & pAdoRS.Fields("STATION_ID").Value
            pAdoRS.MoveNext
        Loop
        obj.ListIndex = 0
    End If
    
    pAdoRS.Close

End Sub

'종합해양관측부이
Private Sub GetBUOYList(Optional ByVal pSTATION_ID As String)
    Dim pAdoRS  As ADODB.Recordset
    Dim intRow  As Integer
    
    Set pAdoRS = Get_BUOYList(pSTATION_ID)
    
'    spdBUOYList.MaxRows = 0
'
'    If pAdoRS Is Nothing Then
'        '등록된 정보 없음
'    Else
'        With spdBUOYList
'            Do Until pAdoRS.EOF
'                .MaxRows = .MaxRows + 1
'                intRow = .MaxRows
'                Call SetText(spdBUOYList, pAdoRS.Fields("STATION_NAME").Value & "", intRow, 1)
'                Call SetText(spdBUOYList, pAdoRS.Fields("OBS_TIME").Value & "", intRow, 2)
'                Call SetText(spdBUOYList, pAdoRS.Fields("EQUIP_ID").Value & "", intRow, 3)
'                Call SetText(spdBUOYList, pAdoRS.Fields("STATION_ID").Value & "", intRow, 4)
'                pAdoRS.MoveNext
'            Loop
'        End With
'    End If
    
    pAdoRS.Close

End Sub

Private Sub GetBOUYViewList(Optional ByVal pSTATION_ID As String, Optional ByVal pOBS_TIME As String, Optional ByVal pCount As Integer)
    
    Dim pAdoRS  As ADODB.Recordset
    Dim intRow  As Integer
    
    Set pAdoRS = Get_BOUYViewList(pSTATION_ID, pOBS_TIME, pCount)
    
    spdBUWIList.MaxRows = 0

    If pAdoRS Is Nothing Then
        '등록된 정보 없음
    Else
        With spdBUWIList
            Do Until pAdoRS.EOF
                .MaxRows = .MaxRows + 1
                intRow = .MaxRows
                Call SetText(spdBUWIList, pAdoRS.Fields("STATION_ID").Value & "", intRow, 1)      '관측소ID
                Call SetText(spdBUWIList, pAdoRS.Fields("STATION_NAME").Value & "", intRow, 2)       '관측소명
                Call SetText(spdBUWIList, pAdoRS.Fields("BUOY_LON").Value & "", intRow, 3)       '
                Call SetText(spdBUWIList, pAdoRS.Fields("BUOY_LAT").Value & "", intRow, 4)       '
                Call SetText(spdBUWIList, Format(pAdoRS.Fields("OBS_TIME").Value, "YYYY/MM/DD hh:mm:ss"), intRow, 5)        '
                Call SetText(spdBUWIList, Format(pAdoRS.Fields("REG_TIME").Value, "YYYY/MM/DD hh:mm:ss"), intRow, 6)
                Call SetText(spdBUWIList, pAdoRS.Fields("WIND_S").Value & "", intRow, 7)
                Call SetText(spdBUWIList, pAdoRS.Fields("WIND_D").Value & "", intRow, 8)
                Call SetText(spdBUWIList, pAdoRS.Fields("WIND_G").Value & "", intRow, 9)
                Call SetText(spdBUWIList, pAdoRS.Fields("AIR_TEMP").Value & "", intRow, 10)
                Call SetText(spdBUWIList, pAdoRS.Fields("AIR_PRES").Value & "", intRow, 11)
                Call SetText(spdBUWIList, pAdoRS.Fields("BUOY_ORIENTATION").Value & "", intRow, 12)
                Call SetText(spdBUWIList, pAdoRS.Fields("WH").Value & "", intRow, 13)
                Call SetText(spdBUWIList, pAdoRS.Fields("WP").Value & "", intRow, 14)
                Call SetText(spdBUWIList, pAdoRS.Fields("CURRENT_S").Value & "", intRow, 15)
                Call SetText(spdBUWIList, pAdoRS.Fields("CURRENT_D").Value & "", intRow, 16)
                Call SetText(spdBUWIList, pAdoRS.Fields("W_TEMP").Value & "", intRow, 17)
                Call SetText(spdBUWIList, pAdoRS.Fields("CONDUCTIVITY").Value & "", intRow, 18)
                Call SetText(spdBUWIList, pAdoRS.Fields("SAL").Value & "", intRow, 19)
                Call SetText(spdBUWIList, pAdoRS.Fields("EQUIP_ID").Value & "", intRow, 20)
                Call SetText(spdBUWIList, pAdoRS.Fields("WIND_D_RAW").Value & "", intRow, 21)
                Call SetText(spdBUWIList, pAdoRS.Fields("VISIBILITY").Value & "", intRow, 22)
                Call SetText(spdBUWIList, pAdoRS.Fields("BATTERY").Value & "", intRow, 23)
                Call SetText(spdBUWIList, pAdoRS.Fields("REFERENCE").Value & "", intRow, 24)
                Call SetText(spdBUWIList, pAdoRS.Fields("TRACK_SEQ").Value & "", intRow, 25)
                Call SetText(spdBUWIList, pAdoRS.Fields("MAX_WAVE_PERIOD").Value & "", intRow, 26)
                Call SetText(spdBUWIList, pAdoRS.Fields("MAX_WAVE_HEIGHT").Value & "", intRow, 27)
                Call SetText(spdBUWIList, pAdoRS.Fields("WAVE_DIRECT").Value & "", intRow, 28)
                Call SetText(spdBUWIList, pAdoRS.Fields("ORIGINAL_WIND_D").Value & "", intRow, 29)
                pAdoRS.MoveNext
            Loop
        End With
    End If
    
    pAdoRS.Close

End Sub

Private Sub Form_Resize()

    'Exit Sub
    On Error Resume Next

    If Me.ScaleHeight = 0 Then Exit Sub

    fra1.WIDTH = Me.ScaleWidth - 100
    fra1.HEIGHT = Me.ScaleHeight - 100
    
    spdVPNList2.HEIGHT = fra1.HEIGHT - fraTimer.HEIGHT - 300
    fraTimer.TOP = spdVPNList2.HEIGHT + 200
    
    fraSearch.WIDTH = fra1.WIDTH - spdVPNList2.WIDTH - 300
    
    spdBUWIList.WIDTH = fraSearch.WIDTH
    spdBUWIList.HEIGHT = (fra1.HEIGHT - fraSearch.HEIGHT) - 300
    
End Sub

Private Sub spdVPNList2_Click(ByVal Col As Long, ByVal Row As Long)
    Dim strDTTSID   As String
    Dim strDTTIME   As String
    Dim intIdx      As Integer
    
    If Row = 0 Then
        Call SetSpreadSort(spdVPNList2, gSORT)
        Exit Sub
    End If

    strDTTSID = GetText(spdVPNList2, Row, 1)
    strDTTIME = GetText(spdVPNList2, Row, 3)
    
    For intIdx = 0 To cboVPNList2.ListCount
        If strDTTSID = mGetP(cboVPNList2.List(intIdx), 2, "|") Then
            cboVPNList2.ListIndex = intIdx
            Exit For
        End If
    Next
    
    '조위관측자료수집로그 가져오기
    Call GetBOUYViewList(strDTTSID, strDTTIME, cboSearchCount.Text)

End Sub



Private Sub tmrResult_Timer()
    
    If chkRefresh.Value = "1" Then
        If cboIntervalGrade.Text = "초" Then
            txtInterval.Text = txtInterval.Text - 1
            If txtInterval.Text = "0" Then
                '자동갱신
                If chkRefresh.Value = "1" Then
                    If cn_Server_Flag = True Then
                        '조위관측소-VPN 리스트 가져오기
                        Call GetVPNList2("")
                    End If

                End If
                
                txtInterval.Enabled = False
                txtInterval.Text = gInterVal
            End If
        Else
            intOneMinute = intOneMinute + 1
            txtInterval60.Text = intOneMinute
            If intOneMinute = 60 Then
                intOneMinute = 0
                txtInterval.Text = txtInterval.Text - 1
                If txtInterval.Text = "0" Then
                    '자동갱신
                    If chkRefresh.Value = "1" Then
                        If cn_Server_Flag = True Then
                            '조위관측소-VPN 리스트 가져오기
                            Call GetVPNList2("")
                        End If
                    End If
                    
                    txtInterval.Enabled = False
                    txtInterval.Text = gInterVal
                End If
            End If
        End If
    End If

End Sub


