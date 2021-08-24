VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmJouiDetail 
   Caption         =   "조위관측소자료"
   ClientHeight    =   11940
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   21435
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11940
   ScaleWidth      =   21435
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
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   20115
         Begin VB.CommandButton cmdExcel 
            BackColor       =   &H00E0E0E0&
            Caption         =   "엑셀"
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
            Left            =   9510
            Style           =   1  '그래픽
            TabIndex        =   14
            Top             =   180
            Width           =   1095
         End
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
            ItemData        =   "frmJouiDetail.frx":0000
            Left            =   7470
            List            =   "frmJouiDetail.frx":0002
            Style           =   2  '드롭다운 목록
            TabIndex        =   6
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
            TabIndex        =   5
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
            ItemData        =   "frmJouiDetail.frx":0004
            Left            =   15300
            List            =   "frmJouiDetail.frx":0006
            Style           =   2  '드롭다운 목록
            TabIndex        =   4
            Top             =   60
            Visible         =   0   'False
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
            ItemData        =   "frmJouiDetail.frx":0008
            Left            =   4980
            List            =   "frmJouiDetail.frx":000A
            Style           =   2  '드롭다운 목록
            TabIndex        =   3
            Top             =   210
            Width           =   795
         End
         Begin VB.ComboBox cboVPNList 
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
            ItemData        =   "frmJouiDetail.frx":000C
            Left            =   870
            List            =   "frmJouiDetail.frx":000E
            Style           =   2  '드롭다운 목록
            TabIndex        =   2
            Top             =   210
            Width           =   2025
         End
         Begin MSComCtl2.DTPicker dtpFromDate 
            Height          =   345
            Left            =   3540
            TabIndex        =   7
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
            Format          =   131465217
            CurrentDate     =   43884
         End
         Begin MSComCtl2.DTPicker dtpToDate 
            Height          =   345
            Left            =   6060
            TabIndex        =   8
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
            Format          =   131465217
            CurrentDate     =   43884
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   12120
            Top             =   90
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
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
            TabIndex        =   12
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
            Left            =   14430
            TabIndex        =   11
            Top             =   150
            Visible         =   0   'False
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
            TabIndex        =   10
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
            TabIndex        =   9
            Top             =   270
            Width           =   585
         End
      End
      Begin FPSpread.vaSpread spdDT 
         Height          =   9045
         Left            =   120
         TabIndex        =   13
         Top             =   870
         Width           =   20115
         _Version        =   393216
         _ExtentX        =   35481
         _ExtentY        =   15954
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
         SpreadDesigner  =   "frmJouiDetail.frx":0010
         ScrollBarTrack  =   3
         ShowScrollTips  =   3
      End
   End
End
Attribute VB_Name = "frmJouiDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim intOneMinute As Integer

'-- 컨트롤초기화
Private Sub CtlInitializing()
    Dim iHour   As Integer
    
    With spdDT
        .MaxCols = 22
        .MaxRows = 0

'        Call SetText(spdDT, "관측소ID", 0, 1):          .ColWidth(1) = 0
'        Call SetText(spdDT, "관측소명", 0, 2):          .ColWidth(2) = 10
'        Call SetText(spdDT, "관측시간", 0, 3):          .ColWidth(3) = 10
'        Call SetText(spdDT, "풍속", 0, 4):              .ColWidth(4) = 10
'        Call SetText(spdDT, "최대풍속(돌풍)", 0, 5):    .ColWidth(5) = 10
'        Call SetText(spdDT, "풍향", 0, 6):              .ColWidth(6) = 10
'        Call SetText(spdDT, "기온", 0, 7):              .ColWidth(7) = 10
'        Call SetText(spdDT, "기압", 0, 8):              .ColWidth(8) = 10
'        Call SetText(spdDT, "조위", 0, 9):              .ColWidth(9) = 10
'        Call SetText(spdDT, "표류식조위(OTT)", 0, 10):  .ColWidth(10) = 10
'        Call SetText(spdDT, "압력식조위(WLS)", 0, 11):  .ColWidth(11) = 10
'        Call SetText(spdDT, "염분", 0, 12):             .ColWidth(12) = 10
'        Call SetText(spdDT, "수온", 0, 13):             .ColWidth(13) = 10
'        Call SetText(spdDT, "시정", 0, 14):             .ColWidth(14) = 10
'        Call SetText(spdDT, "참조", 0, 15):             .ColWidth(15) = 10
'        Call SetText(spdDT, "MIROS 조위", 0, 16):       .ColWidth(16) = 10
'        Call SetText(spdDT, "전도도", 0, 17):           .ColWidth(17) = 10
'        Call SetText(spdDT, "년월일시분", 0, 18):       .ColWidth(18) = 10
'        Call SetText(spdDT, "전송플래그", 0, 19):       .ColWidth(19) = 10
'        Call SetText(spdDT, "등록일", 0, 20):           .ColWidth(20) = 10
'        Call SetText(spdDT, "추적번호", 0, 21):         .ColWidth(21) = 10
'        Call SetText(spdDT, "유의파고", 0, 22):         .ColWidth(22) = 10
'        Call SetText(spdDT, "최대파고", 0, 23):         .ColWidth(23) = 10
'        Call SetText(spdDT, "유의파고주기", 0, 24):     .ColWidth(24) = 10
'        Call SetText(spdDT, "최대파고주기", 0, 25):     .ColWidth(25) = 10
'        Call SetText(spdDT, "일사", 0, 26):             .ColWidth(26) = 10
'        Call SetText(spdDT, "레이저조위", 0, 27):       .ColWidth(27) = 10
'        Call SetText(spdDT, "D_OTT", 0, 28):            .ColWidth(28) = 10
    
        Call SetText(spdDT, "관측소ID", 0, 1):          .ColWidth(1) = 0
        Call SetText(spdDT, "관측소명", 0, 2):          .ColWidth(2) = 10
        Call SetText(spdDT, "관측시간", 0, 3):          .ColWidth(3) = 10
        Call SetText(spdDT, "표류식조위", 0, 4):       .ColWidth(4) = 10
        Call SetText(spdDT, "압력식조위", 0, 5):       .ColWidth(5) = 10
        Call SetText(spdDT, "미러스조위", 0, 6):       .ColWidth(6) = 10
        Call SetText(spdDT, "D_OTT", 0, 7):            .ColWidth(7) = 10
        Call SetText(spdDT, "LASER", 0, 8):            .ColWidth(8) = 10
        Call SetText(spdDT, "풍속", 0, 9):              .ColWidth(9) = 10
        Call SetText(spdDT, "최대풍속", 0, 10):          .ColWidth(10) = 10
        Call SetText(spdDT, "풍향", 0, 11):             .ColWidth(11) = 10
        Call SetText(spdDT, "기온", 0, 12):              .ColWidth(12) = 10
        Call SetText(spdDT, "기압", 0, 13):              .ColWidth(13) = 10
        Call SetText(spdDT, "염분", 0, 14):             .ColWidth(14) = 10
        Call SetText(spdDT, "수온", 0, 15):             .ColWidth(15) = 10
        Call SetText(spdDT, "전도도", 0, 16):           .ColWidth(16) = 10
        Call SetText(spdDT, "유의파고", 0, 17):         .ColWidth(17) = 10
        Call SetText(spdDT, "최대파고", 0, 18):         .ColWidth(18) = 10
        Call SetText(spdDT, "유의파고주기", 0, 19):     .ColWidth(19) = 10
        Call SetText(spdDT, "최대파고주기", 0, 20):     .ColWidth(20) = 10
        Call SetText(spdDT, "시정", 0, 21):             .ColWidth(21) = 10
        Call SetText(spdDT, "일사", 0, 22):             .ColWidth(22) = 10
    
    
    End With
    
    
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


Private Sub cmdExcel_Click()
    Dim sFileName As String
            
On Error GoTo ErrHandler

    If spdDT.DataRowCnt < 1 Then
        MsgBox "저장할 자료가 없습니다.", , "알 림"
        Exit Sub
    Else
        With CommonDialog1
            .CancelError = True
            .Flags = cdlOFNHideReadOnly
            .InitDir = App.PATH
            .Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*"
            .Filename = App.PATH & "\" & Format(Now, "yyyy-mm-dd") & "_결과대장.xlsx"
            .ShowSave
            sFileName = CommonDialog1.Filename
            SaveExcel sFileName, spdDT
            MsgBox "엑셀 저장완료", vbOKOnly + vbInformation, Me.Caption
        End With
    End If

Exit Sub
  
ErrHandler:
      
    ' 사용자가 [취소] 단추를 눌렀습니다.
    Exit Sub

End Sub

Private Sub cmdSearch2_Click()
        
    Call GetDTList_Detail(mGetP(cboVPNList.Text, 2, "|"), dtpFromDate.Value, cboFromHour.Text, dtpToDate.Value, cboToHour.Text, cboSearchCount.Text)
        
End Sub

Private Sub Form_Load()

    Call CtlInitializing
    
    
    If cn_Server_Flag = True Then
        Call GetVPNList_Combo(cboVPNList, "")
    End If
    
End Sub

'관측정보
Private Sub GetDTList_Detail(Optional ByVal pDT_TS_ID As String, _
                            Optional ByVal pDT_F_TIME As String, _
                            Optional ByVal pDT_F_HOUR As String, _
                            Optional ByVal pDT_T_TIME As String, _
                            Optional ByVal pDT_T_HOUR As String, _
                            Optional ByVal pCount As Integer)
    
    Dim pAdoRS  As ADODB.Recordset
    Dim intRow  As Integer
    
    Set pAdoRS = Get_DTList_Detail(pDT_TS_ID, pDT_F_TIME, pDT_F_HOUR, pDT_T_TIME, pDT_T_HOUR, pCount)
    
    spdDT.MaxRows = 0

    If pAdoRS Is Nothing Then
        '등록된 정보 없음
    Else
        With spdDT
            Do Until pAdoRS.EOF
                .MaxRows = .MaxRows + 1
                intRow = .MaxRows
                
'                Call SetText(spdDT, pAdoRS.Fields("DT_TS_ID").Value & "", intRow, 1)      '관측소ID
'                Call SetText(spdDT, pAdoRS.Fields("TS_NAME").Value & "", intRow, 2)       '관측시간
'                Call SetText(spdDT, pAdoRS.Fields("DT_TIME").Value & "", intRow, 3)   '
'                Call SetText(spdDT, pAdoRS.Fields("DT_WSPEED").Value & " ", intRow, 4)   '
'                Call SetText(spdDT, pAdoRS.Fields("DT_WMSPEED").Value & " ", intRow, 5)   '
'                Call SetText(spdDT, pAdoRS.Fields("DT_WDIR").Value & " ", intRow, 6)   '
'                Call SetText(spdDT, pAdoRS.Fields("DT_ATEMP").Value & " ", intRow, 7)   '
'                Call SetText(spdDT, pAdoRS.Fields("DT_APRESS").Value & " ", intRow, 8)   '
'                Call SetText(spdDT, pAdoRS.Fields("DT_TIDE").Value & " ", intRow, 9)   '
'                Call SetText(spdDT, pAdoRS.Fields("DT_FTIDE").Value & " ", intRow, 10)   '
'                Call SetText(spdDT, pAdoRS.Fields("DT_PTIDE").Value & " ", intRow, 11)   '
'                Call SetText(spdDT, pAdoRS.Fields("DT_SAL").Value & " ", intRow, 12)   '
'                Call SetText(spdDT, pAdoRS.Fields("DT_WTEMP").Value & " ", intRow, 13)   '
'                Call SetText(spdDT, pAdoRS.Fields("DT_VISIBILITY").Value & " ", intRow, 14)   '
'                Call SetText(spdDT, pAdoRS.Fields("REFERENCE").Value & " ", intRow, 15)   '
'                Call SetText(spdDT, pAdoRS.Fields("DT_TIDE3").Value & " ", intRow, 16)   '
'                Call SetText(spdDT, pAdoRS.Fields("DT_CONDUCTIVITY").Value & " ", intRow, 17)   '
'                Call SetText(spdDT, pAdoRS.Fields("DT_REF_TIME").Value & " ", intRow, 18)   '
'                Call SetText(spdDT, pAdoRS.Fields("DT_FLAG").Value & " ", intRow, 19)   '
'                Call SetText(spdDT, pAdoRS.Fields("REG_DATE").Value & " ", intRow, 20)   '
'                Call SetText(spdDT, pAdoRS.Fields("TRACK_SEQ").Value & " ", intRow, 21)   '
'                Call SetText(spdDT, pAdoRS.Fields("DT_SIGNIFI_WAVE").Value & " ", intRow, 22)   '
'                Call SetText(spdDT, pAdoRS.Fields("DT_MAX_WAVE").Value & " ", intRow, 23)   '
'                Call SetText(spdDT, pAdoRS.Fields("DT_SIGNIFI_WAVE_PERIOD").Value & " ", intRow, 24)   '
'                Call SetText(spdDT, pAdoRS.Fields("DT_MAX_WAVE_PERIOD").Value & " ", intRow, 25)   '
'                Call SetText(spdDT, pAdoRS.Fields("DT_INSOLATION").Value & " ", intRow, 26)   '
'                Call SetText(spdDT, pAdoRS.Fields("DT_LTIDE").Value & " ", intRow, 27)   '
'                Call SetText(spdDT, pAdoRS.Fields("DT_LTIDE2").Value & " ", intRow, 28)   '
                
                Call SetText(spdDT, pAdoRS.Fields("DT_TS_ID").Value & "", intRow, 1)      '관측소ID
                Call SetText(spdDT, pAdoRS.Fields("TS_NAME").Value, intRow, 2)         '관측시간
                Call SetText(spdDT, Format(pAdoRS.Fields("DT_TIME").Value, "YYYY/MM/DD hh:mm:ss"), intRow, 3)   '
                Call SetText(spdDT, pAdoRS.Fields("DT_FTIDE").Value & "", intRow, 4)   '
                Call SetText(spdDT, pAdoRS.Fields("DT_PTIDE").Value & "", intRow, 5)   '
                Call SetText(spdDT, pAdoRS.Fields("DT_TIDE3").Value & "", intRow, 6)   '
                Call SetText(spdDT, pAdoRS.Fields("DT_LTIDE2").Value & "", intRow, 7)   '
                Call SetText(spdDT, pAdoRS.Fields("DT_LTIDE").Value & "", intRow, 8)   '
                Call SetText(spdDT, pAdoRS.Fields("DT_WSPEED").Value & "", intRow, 9)   '
                Call SetText(spdDT, pAdoRS.Fields("DT_WMSPEED").Value & "", intRow, 10)   '
                Call SetText(spdDT, pAdoRS.Fields("DT_WDIR").Value & "", intRow, 11)   '
                Call SetText(spdDT, pAdoRS.Fields("DT_ATEMP").Value & "", intRow, 12)   '
                Call SetText(spdDT, pAdoRS.Fields("DT_APRESS").Value & "", intRow, 13)   '
                Call SetText(spdDT, pAdoRS.Fields("DT_SAL").Value & "", intRow, 14)   '
                Call SetText(spdDT, pAdoRS.Fields("DT_WTEMP").Value & "", intRow, 15)   '
                Call SetText(spdDT, pAdoRS.Fields("DT_CONDUCTIVITY").Value & "", intRow, 16)   '
                Call SetText(spdDT, pAdoRS.Fields("DT_SIGNIFI_WAVE").Value & "", intRow, 17)   '
                Call SetText(spdDT, pAdoRS.Fields("DT_MAX_WAVE").Value & "", intRow, 18)   '
                Call SetText(spdDT, pAdoRS.Fields("DT_SIGNIFI_WAVE_PERIOD").Value & "", intRow, 19)   '
                Call SetText(spdDT, pAdoRS.Fields("DT_MAX_WAVE_PERIOD").Value & "", intRow, 20)   '
                Call SetText(spdDT, pAdoRS.Fields("DT_VISIBILITY").Value & "", intRow, 21)   '
                Call SetText(spdDT, pAdoRS.Fields("DT_INSOLATION").Value & "", intRow, 22)   '
                
                pAdoRS.MoveNext
            Loop
        End With
    End If
    
    pAdoRS.Close

End Sub



Private Sub GetVPNList_Combo(ByVal obj As Object, Optional ByVal pDT_TS_ID As String)
    Dim pAdoRS  As ADODB.Recordset
    
    Set pAdoRS = Get_VPNList(pDT_TS_ID, True)
    
    If pAdoRS Is Nothing Then
        '등록된 정보 없음
    Else
        obj.Clear
        Do Until pAdoRS.EOF
            obj.AddItem pAdoRS.Fields("TS_NAME").Value & Space(30) & "|" & pAdoRS.Fields("DT_TS_ID").Value
            pAdoRS.MoveNext
        Loop
        obj.ListIndex = 0
    End If
    
    pAdoRS.Close

End Sub


''관측정보
'Private Sub GetDTList(Optional ByVal pDT_TS_ID As String, Optional ByVal pDT_TIME As String)
'
'    Dim pAdoRS  As ADODB.Recordset
'    Dim intRow  As Integer
'
'    Set pAdoRS = Get_DTList(pDT_TS_ID, pDT_TIME)
'
'    spdDT.MaxRows = 0
'
'    If pAdoRS Is Nothing Then
'        '등록된 정보 없음
'    Else
'        With spdDT
'            Do Until pAdoRS.EOF
'                .MaxRows = .MaxRows + 1
'                intRow = .MaxRows
'                Call SetText(spdDT, pAdoRS.Fields("DT_TS_ID").Value & "", intRow, 1)      '관측소ID
'                Call SetText(spdDT, pAdoRS.Fields("DT_TIME").Value & "", intRow, 2)       '관측시간
'                Call SetText(spdDT, pAdoRS.Fields("").Value & "", intRow, 3)   '
'                Call SetText(spdDT, pAdoRS.Fields("").Value & "", intRow, 4)   '
'                Call SetText(spdDT, pAdoRS.Fields("").Value & "", intRow, 5)   '
'                Call SetText(spdDT, pAdoRS.Fields("").Value & "", intRow, 6)   '
'                Call SetText(spdDT, pAdoRS.Fields("").Value & "", intRow, 7)   '
'                Call SetText(spdDT, pAdoRS.Fields("").Value & "", intRow, 5)   '
'                Call SetText(spdDT, pAdoRS.Fields("").Value & "", intRow, 5)   '
'                Call SetText(spdDT, pAdoRS.Fields("").Value & "", intRow, 5)   '
'                pAdoRS.MoveNext
'            Loop
'        End With
'    End If
'
'    pAdoRS.Close
'
'End Sub



Private Sub Form_Resize()

    On Error Resume Next

    If Me.ScaleHeight = 0 Then Exit Sub

    fra1.WIDTH = Me.ScaleWidth - 100
    fra1.HEIGHT = Me.ScaleHeight - 100
    
    fraSearch.WIDTH = fra1.WIDTH - 300
    
    spdDT.WIDTH = fraSearch.WIDTH
    spdDT.HEIGHT = (fra1.HEIGHT - fraSearch.HEIGHT) - 400
    spdDT.TOP = fraSearch.TOP + fraSearch.HEIGHT + 100

End Sub

