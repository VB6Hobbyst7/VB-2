VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form CT500 
   Caption         =   "Clinitek 500"
   ClientHeight    =   8250
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   ScaleHeight     =   8250
   ScaleWidth      =   11850
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  '위 맞춤
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11850
      _ExtentX        =   20902
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   5
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "자료수신을 시작합니다"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "자료수신을 종료합니다"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "통신환경을 설정합니다 "
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "데이타를 변경할 수 있습니다"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "프로그램을 종료합니다"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   9840
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      Handshaking     =   2
   End
   Begin FPSpread.vaSpread SS 
      Height          =   6780
      Left            =   0
      TabIndex        =   4
      Top             =   1110
      Width           =   6135
      _Version        =   196608
      _ExtentX        =   10821
      _ExtentY        =   11959
      _StockProps     =   64
      BackColorStyle  =   1
      ColsFrozen      =   5
      EditEnterAction =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   5
      SpreadDesigner  =   "CT500.frx":0000
      UserResize      =   1
      VisibleCols     =   5
      VisibleRows     =   120
   End
   Begin VB.Frame Frame2 
      Height          =   1308
      Left            =   6780
      TabIndex        =   10
      Top             =   1170
      Width           =   4524
      Begin Threed.SSPanel SSPan 
         Height          =   492
         Left            =   144
         TabIndex        =   11
         Top             =   168
         Width           =   4236
         _Version        =   65536
         _ExtentX        =   7472
         _ExtentY        =   868
         _StockProps     =   15
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   8.95
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   0
         BevelOuter      =   1
         BevelInner      =   2
      End
      Begin VB.Label lblDate 
         BorderStyle     =   1  '단일 고정
         Height          =   492
         Left            =   144
         TabIndex        =   13
         Top             =   720
         Width           =   2076
      End
      Begin VB.Label lblTime 
         Alignment       =   2  '가운데 맞춤
         BorderStyle     =   1  '단일 고정
         ForeColor       =   &H00FF0000&
         Height          =   492
         Left            =   2304
         TabIndex        =   12
         Top             =   720
         Width           =   2076
      End
   End
   Begin VB.ListBox ErrList 
      BeginProperty Font 
         Name            =   "돋움체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5280
      Left            =   6510
      TabIndex        =   9
      Top             =   2490
      Width           =   5130
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Left            =   6510
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   6300
      Visible         =   0   'False
      Width           =   636
   End
   Begin VB.FileListBox File1 
      Height          =   270
      Left            =   7155
      TabIndex        =   7
      Top             =   6300
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "돋움체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   2145
      TabIndex        =   6
      Top             =   2355
      Width           =   2868
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "돋움체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   195
      TabIndex        =   5
      Top             =   2355
      Width           =   2868
   End
   Begin VB.Timer Timer_RRequest 
      Left            =   6990
      Top             =   675
   End
   Begin VB.Timer Timer_RCheck 
      Left            =   6660
      Top             =   675
   End
   Begin VB.Timer Timer_Picture 
      Interval        =   2000
      Left            =   6030
      Top             =   675
   End
   Begin VB.Timer Timer1 
      Left            =   5640
      Top             =   675
   End
   Begin VB.PictureBox picResult 
      Height          =   6675
      Left            =   6180
      ScaleHeight     =   6615
      ScaleWidth      =   5565
      TabIndex        =   1
      Top             =   1200
      Visible         =   0   'False
      Width           =   5625
      Begin FPSpread.vaSpread SSR 
         Height          =   6600
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   5565
         _Version        =   196608
         _ExtentX        =   9816
         _ExtentY        =   11642
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   4
         MaxRows         =   50
         SpreadDesigner  =   "CT500.frx":1E38
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8145
      Top             =   6300
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   435
      Left            =   0
      TabIndex        =   14
      Top             =   660
      Width           =   5805
      _Version        =   65536
      _ExtentX        =   10239
      _ExtentY        =   767
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
      BorderWidth     =   1
      BevelOuter      =   1
      BevelInner      =   2
      Begin MSComCtl2.DTPicker GeomDate 
         Height          =   315
         Left            =   60
         TabIndex        =   15
         Top             =   60
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   25362433
         CurrentDate     =   36892
      End
      Begin Threed.SSOption SSOpt_Ptno 
         Height          =   225
         Index           =   0
         Left            =   2250
         TabIndex        =   16
         Top             =   90
         Width           =   1485
         _Version        =   65536
         _ExtentX        =   2619
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "병록번호"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "System"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   -1  'True
      End
      Begin Threed.SSOption SSOpt_Ptno 
         Height          =   225
         Index           =   1
         Left            =   3960
         TabIndex        =   17
         Top             =   90
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Slip Number"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "System"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin Threed.SSPanel Label1 
      Height          =   495
      Left            =   6420
      TabIndex        =   3
      Top             =   690
      Width           =   5205
      _Version        =   65536
      _ExtentX        =   9181
      _ExtentY        =   873
      _StockProps     =   15
      Caption         =   "CLINITEK 500"
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelInner      =   1
      Font3D          =   3
   End
   Begin VB.Label lblPort 
      Alignment       =   2  '가운데 맞춤
      Height          =   285
      Left            =   6180
      TabIndex        =   19
      Top             =   7920
      Width           =   1965
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   6090
      Picture         =   "CT500.frx":2A3C
      Top             =   660
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label2 
      Caption         =   "date 변경용 obj"
      Height          =   225
      Left            =   6510
      TabIndex        =   18
      Top             =   6825
      Visible         =   0   'False
      Width           =   1860
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   10530
      Top             =   690
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   7
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CT500.frx":2D46
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CT500.frx":3060
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CT500.frx":337A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CT500.frx":3694
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CT500.frx":39AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CT500.frx":3CC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CT500.frx":3FE2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu MnuOption 
      Caption         =   "옵션(&O)"
      Begin VB.Menu MnuRack 
         Caption         =   "Rack/Position Set "
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu MnuReceive 
      Caption         =   "자료수신 (&R)"
   End
   Begin VB.Menu MnuEnd 
      Caption         =   "수신종료 (&E)"
   End
   Begin VB.Menu MnuSet 
      Caption         =   "환경설정 (&S)"
   End
   Begin VB.Menu MnuChange 
      Caption         =   "결과입력 (&C)"
   End
   Begin VB.Menu MnuExit 
      Caption         =   "종      료 (&X)"
   End
End
Attribute VB_Name = "CT500"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
    Dim InputLineData()         As String
    Dim CasePTNO                As Boolean '/병록번호로 또는 슬립번호를 체크 할꺼냐...
    Dim DataUpdate_Commit       As Boolean
    Dim RPoint                                  'row 위치 지정용
    Dim CPoint                                  'col 위치 지정용
    Dim RSequence                               'Result Record Sequence Check용
    Dim GnRow                   As Integer
    Dim Temp_OrderNo            As String
    Dim Temp_K()                        As String                            ' item table data 입력용 buffer
    
'    Dim RResult
    Dim FileSaveDirChk      As Boolean
    Dim Update_Check        As Boolean          ' 종료 check용
    Dim Update_Check_Force  As Boolean          ' 종료 check용
    Dim Receive_Check       As Boolean
    
    Dim GBTransmit          As String
    Dim strBiDirect_Trans   As Boolean          ' batch로 data 수신시 사용
    
    Dim Tcounter                                ' 수신표시 image count용
    Dim QCounter            As Integer
    Dim RCounter            As Integer          ' data 송신시 record count check 용
    
    Dim Ser                 As Integer
    Dim ResultText          As String
    Dim i
    Dim j
    

    
    Dim Temp_Jeobsu(100, 50)            As String           ' Spread의 Data를 Temp_Jeobsu Array로 Move
    Dim Temp_Request_Code(100, 50)      As String           ' input test code A1~Z6
    Dim Temp_Result(100, 50)            As String           ' Result Input Data
    
    Dim Temp_JCount                     As String           ' Spread의 Data를 Temp_Jeobsu Array로 Move
    Dim Temp_Quality(10, 100)           As String
    
    
    Dim FileClose                       As Boolean
    
    Dim LNormal                         As Boolean          'working list msg termination flag
    
    Dim N                   As Integer
    Dim JeobsuCheck         As Boolean          'Data_Update
    Dim Pflag               As Boolean          'Data_Update
    
    Dim RecordCountSum                          'Data_Update
    Dim RecordCountBit                          'Data_Update
    
    Dim MaxRecordCount      As Long             '검사 code count
    
    Dim MaxDataRowCnt       As Integer
    
    Dim SColumn
    Dim SRow
    
    Dim temp_file           As String           'file directory
    Dim temp_file_Order     As String           'file directory
    
    Dim hSaveFile
    Dim hSaveFile_Order
    
    Dim RBuffer             As String
    Dim RBufferSum          As String
    Dim SOHBuffer           As Boolean
    Dim STXBuffer           As Boolean
    Dim ETXBuffer           As Boolean
    Dim EOTBuffer           As Boolean
    Dim ENQBuffer           As Boolean
    Dim ACKBuffer           As Boolean
    Dim NACKBuffer          As Boolean
    
    Dim SOH                 As String
    Dim STX                 As String
    Dim ETX                 As String
    Dim EOT                 As String
    Dim ENQ                 As String
    Dim ACK                 As String
    Dim NACK                As String
    Dim ETB                 As String
    Dim FF                  As String
    Dim CR                  As String
    Dim LF                  As String
    
    Dim SRS                 As String
    Dim SGS                 As String
    Dim SFS                 As String
    
    Dim C1                  As Integer          ' work buffer column1
    Dim R1                  As Integer          ' work buffer row1
'    Dim Order_Data_Seq      As Integer
'    Dim Or_Seq              As Integer
    
    Dim timerx              As Boolean
    Dim PortOpen            As Boolean
    Dim SSCheck             As Boolean
    
    Dim BC                  As String           ' Block Code
    Dim LC                  As Integer          ' Data Line Line Code
    Dim DataLine            As String           ' Data Line Type
    
    Dim RecordCount

    Dim StrGBER             As String       '응급구분 Check

'    Dim CNT_MaxResultItem'
'
'Private udGblResult(1 To CNT_MaxResultItem) As ResultData                           ' 기타 파트


'
'
Public Sub GotoSpreadSet()
    SS.SetFocus                             'clear후 cell active 상태로 변경
    SS.Row = 1
    SS.Col = 1
    SS.Action = SS_ACTION_GOTO_CELL
    SS.Action = SS_ACTION_ACTIVE_CELL
    
End Sub




Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
    Me.WindowState = 0
    Me.Top = 0
    Me.Left = 0
    Me.Width = 800 * 15
    Me.Height = 600 * 15
End Sub




Private Sub MnuChange_Click()
    
    FrmChange.Show vbModal
    
End Sub

Private Sub mnuRack_Click()
'    frmRackCnt.Show vbModal
'    Call GotoSpreadSet                            ' spread cell active

End Sub

Private Sub picResult_Click()
     picResult.Visible = False
End Sub


Private Sub SSOpt_Ptno_Click(Index As Integer, Value As Integer)
    Select Case Index
        Case 0
            GeomDate.Enabled = True
        Case 1
            GeomDate.Enabled = False
    End Select
End Sub





Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    On Error Resume Next
        
    Select Case Button.Index
            Case 1
                   Call mnuReceive_Click
            Case 2
                   Call mnuEnd_Click
            Case 3
                   Call mnuSet_Click
            Case 4
                Call MnuChange_Click
                   'Call NResult_Delete
            Case 5
                   Call mnuExit_Click
    End Select
    
End Sub




Private Sub Form_Initialize()
    Dim Title$
    Dim i           As Integer
    ' 기존에 이미 Window에 해당 Program이 Loading 되었을 경우
    '        Loading 되어있는 Program이 Activate 되도록 하는 Routine
    '        새로 Loading 하려는 Program 은 End 시킨다
    If App.PrevInstance Then
        Title$ = App.Title
        App.Title = "Temp"
        AppActivate Title$
        End
    End If

End Sub


Private Sub Form_Load()
    ' db connect 초기 작업
    
    DoEvents
    Me.Show
    
    Dim Title$
    
    DoEvents
    Me.Show
    
    
    Call DbAdoConnect("TW_MIS_EXAM", "HOSPITAL", "V2MTS")
    
    SSPan.Caption = "Server 컴퓨터에 접속되었습니다."
    SSPan.ForeColor = Val("&H000000FF&")
    Call SysDate_Get
    GeomDate.Value = GstrSysDate
    Call Parmini                            ' spread 초기화 작업
    Call vaSpread_Clear(SS, 1, 1, SS.MaxCols, SS.MaxRows)
    Call GetIniFile
    
    Call CodeKy_Search                      ' codeky Read from twexam_itemml
    Label1.BackColor = &HC0C0C0
    Label1.FontSize = 15
    Label1.Caption = "CLINITEK 500"
    
'    Call Kdelete                                '30일 경과된 누적 file 삭제
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Update_Check_Force = True Then Exit Sub
    If Update_Check = True Then
        Cancel = 1                          ' cancel = 0 (true) 일 경우만 종료됨
    End If

End Sub


Private Sub Form_Unload(Cancel As Integer)
    If MSComm1.PortOpen = True Then
       MSComm1.PortOpen = False
    End If
    Call DbAdoDisConnect
    End
    
End Sub


Private Sub ErrCancel_Click()
'자료수신취소
    For i = 1 To SS.MaxRows
        For j = 1 To 5
            SS.Row = i
            SS.Col = j
            SS.Lock = False
        Next j
    Next i
'    SS.Enabled = True
    
    Timer_Picture.Interval = 0                  'Timer_Picture_Timer End
'    Timer_Order.Interval = 0                  'Timer_order_Timer End
'    Timer_RCheck.Interval = 0                   'Timer_RCheck_Timer End
    timerx = False
    
    Call WorkDisplay(0)
    Close #hSaveFile
    FileClose = True

End Sub

Private Sub SS_Click(ByVal Col As Long, ByVal Row As Long)
    Dim TempJeobDate                    As String
    Dim TempSlno1                       As String
    Dim TempSlno2                       As String
    
    If Row > 0 And (Col = 3 Or Col = 4) Then
        With SS
            .Row = Row:
            .Col = 1:       TempJeobDate = Format(.Text, "YYYY-MM-DD")
            .Col = 4:       TempSlno1 = .Text
            .Col = 5:       TempSlno2 = .Text
            If TempJeobDate <> "" And TempSlno1 <> "" And TempSlno2 <> "" Then
                Call Result_View(TempJeobDate, TempSlno1, TempSlno2)
                picResult.Visible = True:       picResult.ZOrder 0
            End If
        End With
    Else
        picResult.Visible = False:       picResult.ZOrder 1
    End If
End Sub

Private Sub Result_View(strJDate As String, strSlno1 As String, strSlno2 As String)
    Dim i               As Integer
    
    Call vaSpread_Clear(SSR, 1, 1, SSR.MaxCols, SSR.MaxRows)
    
    strSQL = ""
    strSQL = strSQL & " SELECT A.ITEMKO , A.ITEMNM, B.RESULT1                           " & vbLf
    strSQL = strSQL & " FROM   TWEXAM_ITEMML A,                                         " & vbLf
    strSQL = strSQL & "     TWEXAM_GENERAL_SUB B                                        " & vbLf
    strSQL = strSQL & " WHERE B.JEOBSUDT  = TO_DATE('" & strJDate & "', 'YYYY-MM-DD')   " & vbLf
    strSQL = strSQL & "     AND B.SLIPNO1 = " & strSlno1 & vbLf
    strSQL = strSQL & "     AND B.SLIPNO2 = " & strSlno2 & vbLf
    strSQL = strSQL & "     AND A.CODEKY  = B.ITEMCD                                    " & vbLf
    strSQL = strSQL & " ORDER BY A.CODEKY                                               " & vbLf
    
    Result = adoSQL(strSQL)
    
    If Result <> 0 Or rowindicator = 0 Then
        MsgBox "결과 입력작업이 이루어지지 않은 데이타입니다", vbOKOnly + vbInformation, "결과"
        Exit Sub
    End If
    With SSR
        For i = 0 To rowindicator - 1
            .Row = i + 1
            .Col = 1:           .Text = AdoGetString(Rs, "ITEMNM", i)
            .Col = 2:           .Text = AdoGetString(Rs, "RESULT1", i)
            .Col = 3:           .Text = AdoGetString(Rs, "ITEMKO", i)
        Next i
    End With
End Sub

Private Sub SS_DblClick(ByVal Col As Long, ByVal Row As Long)
'    Dim SSDatex             As String
'
'    If Col <> 1 Or Row > SS.DataRowCnt Then
'        For i = 1 To 100
'            Beep
'        Next i
'        Exit Sub
'    End If
'
'    SS.Col = Col
'    SS.Row = Row

End Sub


'Private Sub SS_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    
'    Dim Rs                  As ADODB.Recordset
'
'    Dim strPt               As String
'    Dim Bjeobsudt           As String
'    Dim Bslipno1            As String
'    Dim Bslipno2            As String
'
'    Dim Checkdouble         As String
'    Dim Ptnolen
'    Dim Temp_ptno
'    Dim Temp_name
'
'    SS.Row = Row
'    SS.Col = Col
'
'    If Col >= 2 Then
'        SS.Text = ""
'        SS.SetFocus                             'clear후 cell active 상태로 변경
'        Exit Sub
'    End If
'
'    strPt = SS.Text
'
'    If Row = 500 Then
'        MsgBox " Max Sequence Number Reached." & _
'               " 프로그램을 재시작하여 주시기 바랍니다. "
'        Exit Sub
'    End If
'
'    If Len(strPt) = 12 Then
'        Bjeobsudt = Mid(strPt, 1, 5)
'        Bslipno1 = Mid(strPt, 6, 2)
'        Bslipno2 = Mid(strPt, 8, 5)
'
'        If Col = 1 And strPt <> "" Then
'            SS.Row = Row - 1
'            If Bslipno2 = SS.Text Then
'
'                SS.Row = Row
'                SS.Text = ""
'                SS.SetFocus                             'clear후 cell active 상태로 변경
'                SS.Action = SS_ACTION_ACTIVE_CELL
'                SS.BackColor = &HFF&
'
'                Exit Sub
'            End If
'        End If
'
'        strSQL = ""
'        strSQL = strSQL & " SELECT PTNO "
'        strSQL = strSQL & "   FROM TWEXAM_GENERAL_SUB "                  ' 고객 MASTER
'        strSQL = strSQL & "  WHERE JEOBSUDT = TO_DATE('" & Bjeobsudt & "','YYYY-MM-DD')"
'        strSQL = strSQL & "    AND SLIPNO1 =   '" & Bslipno1 & "'"        ' 일련번호
'        strSQL = strSQL & "    AND SLIPNO2 =   '" & Bslipno2 & "'"        ' 일련번호
'
'        Result = AdoOpenSet(Rs, strSQL)
'
'        If Result Then
'            Rs.MoveFirst
'            Do While Not Rs.EOF
'                Temp_ptno = Trim$(Rs.Fields("ptno")) & ""
'                Rs.MoveNext
'            Loop
'        Else
'
'            SS.Row = Row
'            SS.Text = ""
'            SS.SetFocus                             'clear후 cell active 상태로 변경
'
'            MsgBox " DATABASE에 등록된 내용이 없거나 접수가 잘못되었습니다." & vbCrLf & vbCrLf & _
'                   " DATA를 재입력 하십시요." & vbCrLf & vbCrLf & _
'                   " 재입력 후에도 같은 ERROR가 발생할 경우 전산실로 연락 바랍니다."
'            SS.Action = SS_ACTION_ACTIVE_CELL
'            Exit Sub
'        End If
'
'        AdoCloseSet Rs
'
'
'        strSQL = ""
'        strSQL = strSQL & " SELECT SNAME "
'        strSQL = strSQL & "   FROM TWBAS_PATIENT "                  ' 고객 MASTER
'        strSQL = strSQL & "  WHERE PTNO = '" & Temp_ptno & "' "     ' PATIENT NO
'
'        If AdoOpenSet(Rs, strSQL) Then
'            Rs.MoveFirst
'            Do While Not Rs.EOF
'                Temp_name = Trim$(Rs.Fields("sname")) & ""
'                Rs.MoveNext
'            Loop
'        Else
'
'            SS.Row = Row
'            SS.Text = ""
'            SS.SetFocus                             'clear후 cell active 상태로 변경
'            SS.Action = SS_ACTION_ACTIVE_CELL
'            MsgBox "DATABASE에 등록된 이름이 없거나 접수가 잘못되었습니다." & vbCrLf & _
'                   "PTNO를 재입력 하십시요." & vbCrLf & vbCrLf & _
'                   "재입력 후에도 계속 같은 ERROR가 발생할 경우 전산실로 연락 바랍니다."
'            Exit Sub
'        End If
'
'
'        SS.Row = Row
'
'        SS.Col = 2
'        SS.Text = Temp_ptno
'
'        SS.Col = 3
'        SS.Text = Temp_name
'
'    End If

'End Sub

'
'Private Sub SS_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    Dim SS_M_col            As Long             ' spread mouse down col val
'    Dim SS_M_row            As Long             ' spread mouse down row val
'
'    Dim Msg, Style, Title, Response
'
'    If Receive_Check = True Then
'        Exit Sub
'    End If
'
'    Call SS.GetCellFromScreenCoord(SS_M_col, SS_M_row, X, Y)
'
'    If (SS_M_col <> 1 Or SS_M_row > SS.DataRowCnt) And Button = vbRightButton Then
'        For i = 1 To 50
'            Beep
'        Next i
'        Exit Sub
'    End If
'
'    If Update_Check = True Then Exit Sub
'
'    If Button = vbRightButton Then                                      ' Value = 2
''        Call SS.GetCellFromScreenCoord(SS_M_col, SS_M_row, x, y)
'        SS.Col = SS_M_col
'        SS.Row = SS_M_row
'        If SS.ActiveRow = SS_M_row And SS.ActiveCol = 1 And SS_M_row <= SS.DataRowCnt Then
'            Msg = SS_M_row & " 번째 DATA를 삭제 하시겠습니까?" & vbCrLf & _
'                             " DATA를 확인하셨습니까?"
'            Style = vbYesNo + vbQuestion + vbDefaultButton2             ' Define buttons.
'            Title = "DATA 삭제"                                         ' 기본 제목.
'            Response = MsgBox(Msg, Style, Title)
'            If Response = vbYes Then                                    ' 사용자가 예를 선택.
'                SS.Action = SS_ACTION_DELETE_ROW                        ' value = 5
'
'                'Work 변수 재편성 작업
'
'            End If
'        End If
'    End If
'
'End Sub



'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Private Sub mnuReceive_Click()
'1)검사시작
    
    On Error Resume Next

    If FileSaveDirChk = False Then
        GoSub FileSaveDirShow
    End If
    
    If MSComm1.PortOpen = False Then
       ' MSComm1.Handshaking = comRTS        'SF-3000은 요걸로 해야 해여
        
        MSComm1.InBufferSize = 1024         '8192         ' InBufferSize 변경은 portopen = false일 경우만 가능  ' default = 1024
        MSComm1.PortOpen = True
    End If
    
    MSComm1.RThreshold = 1                  ' MSComm 컨트롤은 수신 버퍼에 한 문자가 들어 올 때마다 OnComm 이벤트를 발생시킵니다.
    MSComm1.InputLen = 1
    
    Call WorkDisplay(1)                     '  "파일로 수신 중입니다." MsgBox Display
    
'******************** Part 4 **********************************
'*  Output File OPen                                          *
'**************************************************************
    FileClose = False
    Timer_Picture.Interval = 1000                            'Timer_Picture_Timer 1000mS
    Receive_Check = True
    hSaveFile = FreeFile
    
    Open temp_file For Append As hSaveFile
    
    If Err Then
        MsgBox Error$, vbExclamation
        Close hSaveFile
        hSaveFile = 0
        Call WorkDisplay(0)
        Exit Sub
    End If
    
    SSPan = "DATA 수신 대기중입니다."
    SSPan.ForeColor = &H0&                  ' black  &H00000000&
    
    Exit Sub
'/==================================================================
'/==================================================================
FileSaveDirShow:

    'temp_file = App.Path & App.Path & "\InitData\" & "H" & Format$(lblDate, "yyyymmdd") & ".ifc"
    GnRow = 0
    
    If SS.DataRowCnt = 0 Then       ' Transmition of Patient File to STA
        strBiDirect_Trans = False   ' batch mode
    Else
        strBiDirect_Trans = True    ' Query Mode
    End If

    RSequence = 0
    QCounter = 0
    RCounter = 1

    ErrList.Clear

    FileClose = False
    Timer_Picture.Interval = 1000                            'Timer_Picture_Timer 1000mS
    N = 0

    SColumn = 6                                             'spread sheet 초기화 위치
    SRow = 1

'******************** Part 3 **********************************
'*    Output용 File Name Set Open/Save 처리                   *
'**************************************************************
    On Error GoTo ErrorMsg
    'Ser = Ser + 1
    CommonDialog1.InitDir = App.Path & "\InitData\"
    CommonDialog1.FileName = "C" & Format$(lblDate, "yyyymmdd") '& Ser

    'CommonDialog1.Flags = "&H2"                             '파일중복 check & msgbox
    CommonDialog1.Filter = "All Files (*.*)|*.*|" & _
                           "Text Files" & "(*.txt)|*.txt|" & _
                           "Ifc Files" & "(*.ifc)|*.ifc|"
    CommonDialog1.FilterIndex = 3
    CommonDialog1.CancelError = True                        '취소확인을 위해 사용

    On Error GoTo ErrCancel
    CommonDialog1.ShowSave                                  'dialog show
    CommonDialog1.CancelError = True                        'cancel error reset

    On Error GoTo ErrorMsg

    DoEvents
    Me.Show

    temp_file = CommonDialog1.FileName
    FileSaveDirChk = True

    Return
'/============================================================================================
ErrCancel:

    MsgBox "자료수신을 취소하였습니다."
    Call ErrCancel_Click                    ' 자료수신 종료
    Exit Sub
Exit Sub
'/============================================================================================
ErrorMsg:

    MsgBox "Error " & "Code = " & Err.Number & vbLf & vbLf & Err.Description
    If FileClose = False Then
        Close #hSaveFile                    ' error 발생시 file close
    End If


'    SS.Enabled = True
    Timer_Picture.Interval = 0              'Timer_Picture_Timer End
'    Timer_Order.Interval = 0              'Timer_order_Timer End
    timerx = False                          '수신표시 정지용 flag

    Call WorkDisplay(0)
    If MSComm1.PortOpen = True Then
        MSComm1.PortOpen = False
    End If
    Exit Sub
End Sub



Private Function Data_Analyze(ArgIDNumber As String) As String
    Dim i                           As Integer
    Dim j                           As Integer
    Dim Ars                         As New ADODB.Recordset
    Dim SBarSlno1               As String
    Dim SbarSLno2               As String
    Dim TempName                As String
        
    
'/====================================================================================================
'/Data_Analyze_Clear:
    
    For i = 0 To UBound(nReciveData)
        nReciveData(i).strRecord = ""
    Next i
'/===================================================================================================
'/Data_Analyze:
    
    For i = 1 To UBound(nReciveData)
        For j = 1 To UBound(InputLineData)
            With nReciveData(i)
                TempName = Mid(InputLineData(j), 2, 3)
                
                If UCase(Trim(TempName)) = UCase(Trim(.StrFullName)) Then
                    .strRecord = Trim(Mid(InputLineData(j), 10, Len(InputLineData(j))))
                    Exit For
                End If
            End With
        Next j
    Next i
    
'//RETURN BARCODENO'/============================================================================

    Data_Analyze = Val(ArgIDNumber)

    If CasePTNO = True Then
        SBarSlno1 = "":        SbarSLno2 = ""
        
        Call GetBarCodeNumber(Data_Analyze, GeomDate.Value, GGCODE)
        
        If Result <> 0 Or rowindicator = 0 Then
            Data_Analyze = ArgIDNumber
        Else
            SBarSlno1 = AdoGetString(Rs, "SlipNo1", rowindicator - 1)      ' 오더가 여러개인 경우 가장 최근에 내려진
            SbarSLno2 = AdoGetString(Rs, "SlipNo2", rowindicator - 1)      ' 오더를 기준으로 한다.
            Data_Analyze = convLabnoToComp(Format(GeomDate.Value, "YYYYMMDD")) & Format(SBarSlno1, "00") & Format(SbarSLno2, "00000")
        End If
        
    End If
    Data_Analyze = Format(Data_Analyze, "000000000000")
End Function

'{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}

Private Sub MSComm1_OnComm() ' DATA 수신 처리
    Dim EVMsg$
    Dim ERMsg$
    Dim TestFile            As String
    
    Select Case MSComm1.CommEvent           'CommEvent 속성에 따른 항목
       '이벤트 메시지
        Case comEvReceive                   ' 포트로부터 데이터가 들어왔음...
             RBuffer = MSComm1.Input
        Case Else:              ERMsg$ = " 오류 이벤트"
    End Select
  
    Select Case RBuffer                                 ' Message Block 단위로 편집 출력
           Case STX, "@":                       ' STX는 날라오는 군여
                STXBuffer = True                ' STX [] Check용
           Case ETX, FF:                        ' ETX 체크가 않되네여...그래서... NEW PAGE 문자를 ETX체크용으로...
                ETXBuffer = True                ' ETX [] Check용
    End Select
    
    RBufferSum = RBufferSum & RBuffer                   ' comm port 에서 입력한 data 누적
    
    If ETXBuffer = True Then                            'ETX Check
        If FileClose = False Then
            Print #hSaveFile, "[" & Format(Time, "hh:mm:ss") & "] : " & RBufferSum
        End If
        
        Call DataReceive(RBufferSum)                             '수신한 data record 분석
        
        RBufferSum = ""
        RBuffer = ""
        ETXBuffer = False
        SSPan = "DATA 수신 중입니다."
    End If
    
End Sub

Private Function DataReceive(strPrmData As String) As Boolean
    '일단은 특별한... 처리는 없지만... 데이타 처리가 제대로 됐는지에 대한 체크는 하고자 한다.
    '이후에... 문제가생기면... 처리를 해야 할거 아닌가.. 그치~~~
    '처리가 제대로 않되면..False 반환
    Dim StrIdNumber                 As String
    Dim BslipNo1                    As String
    Dim BslipNo2                    As String
    Dim OrderDate                   As String
    Dim OrderNumber                 As String
    Dim Temp_Result                 As String
    Dim nCol                        As Integer
    Dim i                           As Integer
    Dim NLine                       As Integer
    Dim NCount                      As Integer
    Dim NSeqNo                      As String
    On Error Resume Next
    
    If SSOpt_Ptno(0).Value = True Then
        CasePTNO = True
    ElseIf SSOpt_Ptno(1).Value = True Then
        CasePTNO = False
    End If
    
    For i = 0 To UBound(InputLineData)
        InputLineData(i) = ""
    Next i
    
    NLine = 0
    
    For i = InStr(1, strPrmData, SRS) To Len(strPrmData)
        If UBound(InputLineData) = NLine Then Exit For
        If Mid(strPrmData, i, 2) = vbCrLf Then
            Debug.Print InputLineData(NLine)
            If UCase(Mid(InputLineData(NLine), 2, 3)) = "ID=" Then
                StrIdNumber = Val(Mid(InputLineData(NLine), 5, 12))
            End If
            If UCase(Mid(InputLineData(NLine), 2, 1)) = "#" Then
                NSeqNo = Mid(InputLineData(NLine), 3, 11)
            End If
            NLine = NLine + 1
        Else
            InputLineData(NLine) = InputLineData(NLine) & Mid(strPrmData, i, 1)
            'MsgBox Mid(strPrmData, i, 2)
        End If
    Next i
    If StrIdNumber = "" Then
        ErrList.AddItem "환자정보 미입력 샘플번호 : " & NSeqNo
        ErrList.ListIndex = ErrList.ListCount - 1
        SSPan = "등록되지 않은 환자입니다."
        DataReceive = False:        Exit Function
    Else
        StrIdNumber = Data_Analyze(StrIdNumber)
    End If
    
    GoSub Spread_View
    GoSub DataUpdate_Rtn
    
    
    Exit Function

'/===================================================================================================
Spread_View:
    
    OrderDate = convLabnoToExpand(Mid(StrIdNumber, 1, 5))
    BslipNo1 = Mid(StrIdNumber, 6, 2)
    BslipNo2 = Mid(StrIdNumber, 8, 5)
    
    If PreResultChk(StrIdNumber) = True Then
        If MsgBox("기존에 입력된 데이타가 있습니다. 새로 입력하시겠습니까?", vbQuestion + vbYesNo, "확인") = vbNo Then
           DataReceive = True:      Exit Function  '/ 데이터의 에러없이 종료
        End If
    End If
    
    strSQL = ""
    strSQL = strSQL & " SELECT DISTINCT PT.SNAME, TG.PTNO, TO_CHAR(TG.JEOBSUDT , 'YYYY-MM-DD') JUBSUDATE,   " & vbLf
    strSQL = strSQL & "        TG.SLIPNO1, TG.SLIPNO2                                                       " & vbLf        ', TG.ROUTINCD , TG.RESULT1
    strSQL = strSQL & " FROM TWEXAM_GENERAL_SUB TG,                                                         " & vbLf
    strSQL = strSQL & "   TW_MIS_PMPA.TWBAS_PATIENT PT                                                      " & vbLf
    strSQL = strSQL & " WHERE JEOBSUDT = TO_DATE('" & OrderDate & "', 'YYYYMMDD')                           " & vbLf
    strSQL = strSQL & "    AND PT.PTNO = TG.PTNO                                                            " & vbLf
    strSQL = strSQL & "    AND TG.SLIPNO1 = " & BslipNo1 & vbLf
    strSQL = strSQL & "    AND TG.SLIPNO2 = " & BslipNo2
    
    Result = adoSQL(strSQL)
    
    If Result <> 0 Or rowindicator = 0 Then
        ErrList.AddItem "Unregistered Patient  일련번호 " & StrIdNumber
        ErrList.ListIndex = ErrList.ListCount - 1
        SSPan = "등록되지 않은 환자입니다."
        DataReceive = False:        Exit Function
    End If
    
    
    For i = 0 To rowindicator - 1
        GnRow = GnRow + 1
        
        If SS.MaxRows < GnRow Then SS.MaxRows = GnRow + 1
        
        With SS
            .Row = GnRow
            .Col = 1:       .Text = AdoGetString(Rs, "JUBSUDATE", i)
            .Col = 2:       .Text = AdoGetString(Rs, "PTNO", i)
            .Col = 3:       .Text = AdoGetString(Rs, "SNAME", i)
            .Col = 4:       .Text = AdoGetString(Rs, "SLIPNO1", i)
            .Col = 5:       .Text = AdoGetString(Rs, "SLIPNO2", i)
        End With
    Next i
    Return

'/====================================================================================================

DataUpdate_Rtn:
    
    strSQL = ""
    strSQL = strSQL & "SELECT PTNO, ITEMCD                                          " & vbLf
    strSQL = strSQL & "FROM TWEXAM_GENERAL_SUB                                      " & vbLf                  ' 고객 MASTER
    strSQL = strSQL & "WHERE JEOBSUDT = TO_DATE('" & OrderDate & "','YYYYMMDD')     " & vbLf
    strSQL = strSQL & "    AND SLIPNO1 =   " & Val(BslipNo1) & "                    " & vbLf        ' 일련번호
    strSQL = strSQL & "    AND SLIPNO2 =   " & Val(BslipNo2) & "                    " & vbLf        ' 일련번호

    Result = adoSQL(strSQL)

    If Result <> 0 Or rowindicator = 0 Then
        DataReceive = False:         Exit Function    'SKIP
    End If
    DataUpdate_Commit = True
    
    adoConnect.BeginTrans
    
    For i = 0 To rowindicator - 1
        If DataUpdate_Commit = False Then Exit For
        Temp_Result = Result_ItemCd(AdoGetString(Rs, "ITEMCD", i))
        If Trim(Temp_Result) <> Trim("No Result") Then         '아이템 코드가 등록되지 않았거나 장비에 항목이 없으면 스킴
            Call Save_Result(StrIdNumber, AdoGetString(Rs, "ITEMCD", i), Temp_Result)
        End If
    Next i
'/==================================================================================================
' TWEXAM_GENERAL 테이블 변경 ... 검사구분... 'E'로 셋팅... 결과 입력 완료...
    
    strSQL = ""
    strSQL = strSQL & "Update Twexam_General           Set                          " & vbLf
    strSQL = strSQL & "         GeomsaGu = 'E'                                      " & vbLf
    strSQL = strSQL & "WHERE JEOBSUDT = TO_DATE('" & OrderDate & "','YYYYMMDD')     " & vbLf
    strSQL = strSQL & "    AND SLIPNO1 =   " & Val(BslipNo1) & "                    " & vbLf        ' 일련번호
    strSQL = strSQL & "    AND SLIPNO2 =   " & Val(BslipNo2) & "                    " & vbLf        ' 일련번호
    
    Result = adoSQL(strSQL)
    
    If Result <> 0 Then
        adoConnect.RollbackTrans
        DataUpdate_Commit = False
        ErrList.AddItem "Unregistered Patient  일련번호 " & StrIdNumber
        ErrList.ListIndex = ErrList.ListCount - 1
        SSPan = "데이타 갱신중 에러 발생 !!"
        DataReceive = False:        Exit Function
    End If
'/==================================================================================================
    If DataUpdate_Commit = True Then adoConnect.CommitTrans
    Return
'/===================================================================================================

End Function


Public Function Result_ItemCd(StrItemCD) As String

    For i = 0 To UBound(nReciveData)
        If Trim(StrItemCD) = Trim(nReciveData(i).StrKeyCode) Then
            Result_ItemCd = nReciveData(i).strRecord
            Exit For
        Else
            Result_ItemCd = "No Result"
        End If
    Next i
    

End Function

Private Function PreResultChk(StrIdNumber As String) As Boolean
    Dim BslipNo1                    As String
    Dim BslipNo2                    As String
    Dim OrderDate                   As String
    Dim Prs                         As New ADODB.Recordset
    
    OrderDate = convLabnoToExpand(Mid(StrIdNumber, 1, 5))
    BslipNo1 = Mid(StrIdNumber, 6, 2)
    BslipNo2 = Mid(StrIdNumber, 8, 5)
    
    strSQL = ""
    strSQL = strSQL & " SELECT STATUS                                             " & vbLf
    strSQL = strSQL & " FROM TWEXAM_GENERAL                                         " & vbLf
    strSQL = strSQL & " WHERE JEOBSUDT = To_Date('" & OrderDate & "', 'YYYYMMDD')   " & vbLf
    strSQL = strSQL & "     AND SLIPNO1 = " & BslipNo1 & vbLf
    strSQL = strSQL & "     AND SLIPNO2 = " & BslipNo2 & vbLf
    
    Result = AdoOpenSet(Prs, strSQL)
    
    If rowindicator > 0 Then
        If AdoGetString(Prs, "STATUS", 0) = "C" Then PreResultChk = True
    End If
    AdoCloseSet Prs
End Function
    

Private Sub Save_Result(JeobsuPT, itemcd1, resultu)
    Dim Urs                     As New ADODB.Recordset
    Dim Bdt
    Dim Bno1
    Dim Bno2
        
    Bdt = convLabnoToExpand(Mid(JeobsuPT, 1, 5))
    Bno1 = Mid(JeobsuPT, 6, 2)
    Bno2 = Mid(JeobsuPT, 8, 5)
    
    strSQL = ""
    strSQL = strSQL & "UPDATE TWEXAM_GENERAL_SUB "
    strSQL = strSQL & "   SET RESULT1  =   '" & resultu & "'"
    strSQL = strSQL & " WHERE JEOBSUDT =  TO_DATE('" & Bdt & "','YYYYMMDD') "    '입력된 날자로 검색
    strSQL = strSQL & "   AND SLIPNO1  =   '" & Bno1 & "' "                        '구분
    strSQL = strSQL & "   AND SLIPNO2  =   '" & Bno2 & "' "                        '구분
    strSQL = strSQL & "   AND ITEMCD   =    '" & itemcd1 & "'"                     'ITEMCODE
    strSQL = strSQL & "   AND VERIFY   =  'N'"                                    ' 접수결과에서 VERIFY OK한경우에는 UPDATE하지않음
    
    Result = AdoExecute(strSQL)
    
    If Result = 0 Then
        SSPan = "DATABASE에 저장 되었습니다. "
        DataUpdate_Commit = True
    Else
        ErrList.AddItem "    Verify Data       " & JeobsuPT
        ErrList.AddItem "    or Update Error   " & itemcd1 & "  " & resultu
        ErrList.ListIndex = ErrList.ListCount - 1
        adoConnect.RollbackTrans                                                 ' TRANSACTION ERROR시 ROLLBACK 시킴
        SSPan = "DB에 갱신중 ERROR가 발생하였습니다." & vbCrLf & _
                "VERIFY된 DATA인지 확인하십시요."
        DataUpdate_Commit = False
    End If
    
    
End Sub


'Private Sub ini_check()
'
'
'    Dim Bjeobsudt
'    Dim BslipNo1
'    Dim BslipNo2
'
'    Dim Verify_Check
''******************** Part 1 **********************************
''*      Spread의 pt no와 slipno2를 Temp_Jeobsu Array로 Move   *
''**************************************************************
'
'    R1 = RCounter
'        C1 = 1
'
'        SS.Row = R1
'        SS.Col = C1
'        SS.Text = Temp_Jeobsu(R1, C1)
'
'                       'Rn  Cn
'        Bjeobsudt = convLabnoToExpand(Mid(Temp_Jeobsu(R1, C1), 1, 5))
'        BslipNo1 = Mid(Temp_Jeobsu(R1, C1), 6, 2)
'        BslipNo2 = Mid(Temp_Jeobsu(R1, C1), 8, 5)
'
'        For C1 = 2 To 4
'            SS.Row = R1
'            Select Case C1
'                   Case 2   'date
'                        SS.Col = C1
'                        Temp_Jeobsu(R1, C1) = Bjeobsudt
'                        SS.Text = Temp_Jeobsu(R1, C1)
'
'                   Case 3   'PTNO
'                        SS.Col = C1
'                        Temp_Jeobsu(R1, C1) = PTNOSearch(Temp_Jeobsu(R1, 1))
'                        SS.Text = Temp_Jeobsu(R1, C1)
'                   Case 4   'Name
'                        SS.Col = C1
'                        Temp_Jeobsu(R1, C1) = NameSearch(Temp_Jeobsu(R1, 3))
'                        SS.Text = Temp_Jeobsu(R1, C1)
'            End Select
'
'        Next C1
'
'        strSQL = ""
'        strSQL = strSQL & " SELECT ITEMCD, GEOMJAN2, GEOMJAN3,GBER "
'        strSQL = strSQL & "   FROM TWEXAM_GENERAL_SUB A, "                   ' 검사접수결과 세부사항
'        strSQL = strSQL & "        TWEXAM_ITEMML B, "                        ' 검사 ITEM MASTER
'        strSQL = strSQL & "        TWEXAM_GENERAL C "                        ' 검사접수결과
'        strSQL = strSQL & "  WHERE A.JEOBSUDT = TO_DATE('" & Bjeobsudt & "','YYYY-MM-DD')"
'        strSQL = strSQL & "    AND A.SLIPNO1 =   '" & BslipNo1 & "'"        ' 일련번호
'        strSQL = strSQL & "    AND A.SLIPNO2 =   '" & BslipNo2 & "'"        ' 일련번호
'        strSQL = strSQL & "    AND A.ITEMCD = B.CODEKY "
'        strSQL = strSQL & "    AND B.GBROUTINE = 'I'   "
'        strSQL = strSQL & "    AND A.PTNO = C.PTNO "
'        strSQL = strSQL & "    AND A.JEOBSUDT = C.JEOBSUDT "
'        strSQL = strSQL & "    AND A.SLIPNO1 = C.SLIPNO1 "
'        strSQL = strSQL & "    AND A.SLIPNO2 = C.SLIPNO2 "
'        strSQL = strSQL & "    AND B.GEOMJAN1 = '" & GGJCODE & "' "
'
'        Result = AdoOpenSet(Rs, strSQL)
'
'        'Debug.Print Rowindicator
'
'        If Result Then
'            Do While Not Rs.EOF
'                If Val(Trim(Rs.Fields("GEOMJAN2") & "")) >= "11" Then
'                    Temp_Jeobsu(R1, Val(Trim(Rs.Fields("GEOMJAN2") & ""))) = Trim(Rs.Fields("ITEMCD") & "")
'                    If Trim(Rs.Fields("GBER") & "") = "E" Then
'                        StrGBER = "S"
'                    Else
'                        StrGBER = "R"
'                    End If
'
'                End If
'
'                Rs.MoveNext
'            Loop
'        End If
'
'    If StrGBER = "S" Then
'        SS.Row = R1
'        For i = 1 To 6
'            SS.Col = i
'            SS.ForeColor = RGB(255, 0, 0)
'        Next i
'    End If
'
'    SS.Col = 5
'    Temp_Jeobsu(R1, 5) = rowindicator
'    SS.Text = rowindicator
'
'    SS.SetFocus                             ' cell active 상태로 변경
'    SS.Action = SS_ACTION_ACTIVE_CELL       ' 지정된 위치로 cursor 이동
'
'End Sub


Private Sub Save_Result_Flag(JeobsuPT2)

    Dim Bdt
    Dim Bno1
    Dim Bno2
        
    Bdt = convLabnoToExpand(Mid(JeobsuPT2, 1, 5))
    Bno1 = Mid(JeobsuPT2, 6, 2)
    Bno2 = Mid(JeobsuPT2, 8, 5)
    
    strSQL = ""
    strSQL = strSQL & " SELECT JEOBSUDT, SLIPNO1, SLIPNO2, STATUS               "
    strSQL = strSQL & "   FROM TWEXAM_GENERAL                                   "
    strSQL = strSQL & "  WHERE JEOBSUDT = TO_DATE('" & Bdt & "','YYYY-MM-DD')   "
    strSQL = strSQL & "    AND SLIPNO1 =   '" & Bno1 & "'                       "
    strSQL = strSQL & "    AND SLIPNO2 =   '" & Bno2 & "'                       "
    strSQL = strSQL & "    AND (STATUS  = 'R' OR STATUS = 'U')                  "
    
    Result = AdoOpenSet(Rs, strSQL)
        
    If Result Then
        adoConnect.BeginTrans                          ' TRANSACTION의 종료시에 COMMITTRANS를 지정함
        
        strSQL = ""
        strSQL = strSQL & "UPDATE TWEXAM_GENERAL "
        strSQL = strSQL & "   SET STATUS   = 'U' "
        strSQL = strSQL & " WHERE JEOBSUDT = TO_DATE('" & Bdt & "','YYYY-MM-DD') "    '입력된 날자로 검색
        strSQL = strSQL & "   AND SLIPNO1  = '" & Bno1 & "' "                        '구분
        strSQL = strSQL & "   AND SLIPNO2  = '" & Bno2 & "' "                        '구분
        strSQL = strSQL & "   AND (STATUS  = 'R' OR STATUS = 'U') "
        
        Result = AdoExecute(strSQL)
        If Result = True And rowindicator > 0 Then
            SSPan = "DATABASE에 저장 되었습니다. "
            adoConnect.CommitTrans                                                   ' TRANSACTION 종료시에 COMMIT 시킴
        Else
            ErrList.AddItem "    Verify OK         " & JeobsuPT2
            ErrList.AddItem "    or 접수취소       "
            ErrList.ListIndex = ErrList.ListCount - 1
            
            adoConnect.RollbackTrans                                                 ' TRANSACTION ERROR시 ROLLBACK 시킴
            SSPan = " DATABASE에 갱신중 ERROR가 발생하였습니다." & vbCrLf & _
                    " 결과완료된 DATA인지 확인하십시요."
        End If
        
    End If

End Sub
Private Sub mnuEnd_Click()
'2)수신종료
    
    On Error Resume Next
'    If SS.DataRowCnt < 1 Then Exit Sub

    If Receive_Check = False Then Exit Sub
    
    
    Dim Msg, Style, Title, Response
    RecordCount = 0
    Msg = " 검사를 종료 하시겠습니까?" '& vbCrLf & " 미수신된 자료를 확인하셨습니까?"
    Style = vbYesNo + vbQuestion + vbDefaultButton2     ' Define buttons.
    Title = "검사 종료 확인"                          ' 기본 제목.
    Response = MsgBox(Msg, Style, Title)

    If Response = vbNo Then Exit Sub                    ' 사용자가 아니오 선택시 무동작.

    Receive_Check = False

'    For i = 1 To SS.MaxRows
'        For j = 1 To 6
'            SS.Row = i
'            SS.Col = j
'            SS.Lock = False
'        Next j
'    Next i

'        SS.Enabled = True
    Timer_Picture.Interval = 0                       'Timer_Picture_Timer End
'    Timer_Order.Interval = 0                       'Timer_order_Timer End
    timerx = False

    Call WorkDisplay(0)
'
''@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'    Print #hSaveFile, " "
'    ResultText = ""
'    For R1 = 1 To MaxDataRowCnt
'        For C1 = 7 To SS.MaxCols + 6
'            If Temp_Result(R1, C1) <> "" Then
'                ResultText = ResultText & "  " & C1 & " = " & Temp_Result(R1, C1)
'            End If
'        Next C1
'        Print #hSaveFile, Format$(R1, "000") & " " & ResultText
'        ResultText = ""
'    Next R1
''@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'    Print #hSaveFile, " " & vbCrLf & "@@@@@  Spread Data" & vbCrLf & " "
'    ResultText = ""
'    For R1 = 0 To MaxDataRowCnt
'        For C1 = 1 To SS.MaxCols
'            SS.Row = R1
'            SS.Col = C1
'            If C1 <> 3 Then
'                If SS.Text = "" Then SS.Text = "0"
'                ResultText = ResultText & Format$(Trim$(SS.Text), "@@@@@@@@@@@@@@") & " : "
'            End If
'        Next C1
'        Print #hSaveFile, ResultText
'        ResultText = ""
'    Next R1
''@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
    If FileClose = False Then
        Close #hSaveFile
        FileClose = True
    End If
    
    If MSComm1.PortOpen = True Then
        MSComm1.PortOpen = False
    End If

End Sub


'Private Sub mnuWrite_Click()
''3)자료저장
''Spread sheet의 data를 Server에 저장
'
'    If strBiDirect_Trans = True Then
'        Exit Sub
'    End If
'
'    On Error Resume Next
'
'    If SS.DataRowCnt < 1 Then Exit Sub
'    Dim Msg, Style, Title, Response
'    Msg = " 자료를 DATABASE에 " & vbCrLf & "저장하시겠습니까?"
'    Style = vbYesNo + vbQuestion + vbDefaultButton2 ' Define buttons.
'    Title = "DATABASE UPDATE"                       ' 기본 제목.
'    Response = MsgBox(Msg, Style, Title)
'    If Response = vbYes Then                        ' 사용자가 예를 선택.
'        Call Data_Update                            ' 임상병리 검사종류 11(생화학)검사에대한 ITEM CODE 검색 & SET
'    End If
'
'End Sub
'

'Private Sub Data_Update()
'
'
'    SSPan = "DATABASE에 저장하고 있습니다."
'    Pflag = False
'    JeobsuCheck = True
'
'''''''''''''''''''''''''''''''''''''''''''''''' TRANSACTION 의 시작위치 지정
'    adoConnect.BeginTrans                          ' TRANSACTION의 종료시에 COMMITTRANS를 지정함
'
'    ' DATABASE UPDATE                                                                       '
'    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'    '임상병리 검사종류 11(생화학 자동분석)검사에대한 ITEM CODE SETTING
'    ' Temp_Jeobsu(R1,C1)에 settting 되어있음
'
'    For R1 = 1 To MaxDataRowCnt
'        For C1 = 7 To MaxRecordCount + 6        'itemcd를 check하기위한 for next
'            If Trim$(Temp_Result(R1, C1)) <> "" Then
'
''                '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
''                List3.AddItem C1 & "  " & Temp_Result(R1, C1)
''                List3.ListIndex = List3.ListCount - 1
''                Debug.Print Temp_Result(R1, C1)
'                '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
'                strSQL = ""
'                strSQL = strSQL & "UPDATE TWEXAM_GENERAL_SUB1                                                   "
'                strSQL = strSQL & "   SET RESULT1  =   '" & Format(Val(Temp_Result(R1, C1)), "###0.000") & "'   "
'                strSQL = strSQL & " WHERE JEOBSUDT =  TO_DATE('" & Temp_Jeobsu(R1, 4) & "','YYYY-MM-DD')        "      '입력된 날자로 검색
'                strSQL = strSQL & "   AND SLIPNO1  =   11                                                       "                                                      '구분
'                strSQL = strSQL & "   AND VERIFY   =  'N'                                                       "                                                      ' 접수결과에서 VERIFY OK한경우에는 UPDATE하지않음
'                strSQL = strSQL & "   AND SLIPNO2  =     " & Temp_Jeobsu(R1, 2)                                 '일일 2건이상일 경우 CHECK
'                strSQL = strSQL & "   AND PTNO     =    '" & Trim$(Temp_Jeobsu(R1, 1)) & "'                     "                    'PATIENT NUMBER
'                strSQL = strSQL & "   AND ITEMCD   =    '" & Trim$(Temp_K(C1 - 6, 0)) & "'                      "                     'ITEMCODE
'
'                Result = AdoExecute(strSQL)
'                If Result >= 0 And rowindicator > 0 Then
'                    RecordCountBit = 1
'                ElseIf Result = -1 Then
'                    MsgBox "Check Error" & vbCrLf & R1 & "번째 data를 확인하십시요 "
'                    JeobsuCheck = False
'                End If
'            End If
'        Next C1
'        RecordCountSum = RecordCountSum + RecordCountBit
'        RecordCountBit = 0
'    Next R1
'
'    If Result Then
'        SSPan = "DATABASE에 저장 되었습니다. ( " & RecordCountSum & " 건)"
'        If RecordCountSum = 0 Then SSPan = " 저장된 Data가 없습니다."
'        adoConnect.CommitTrans                                                   ' TRANSACTION 종료시에 COMMIT 시킴
'        RecordCountSum = 0
'        Update_Check = False
'    Else
'        MsgBox "   Update Error     "
'        adoConnect.RollbackTrans                                                 ' TRANSACTION ERROR시 ROLLBACK 시킴
'        SSPan = "DATABASE에 갱신중 ERROR가 발생하였습니다."
'        Update_Check = False
'    End If
'
'End Sub


Private Sub mnuClear_Click()
'4)화면 Clear
    Dim Msg, Style, Title, Response
    If Update_Check = True Then
        Msg = " 수신한 DATA를 저장하지 않았습니다." & vbCrLf & _
              " DATA를 저장한다음 화면을 CLEAR하십시요." & vbCrLf & _
              "                                " & vbCrLf & _
              " 화면을 CLEAR하시겠습니까?"
        Style = vbYesNo + vbDefaultButton2 + vbCritical     ' Define buttons.
        Title = "화면 CLEAR"                                ' 기본 제목.
        Response = MsgBox(Msg, Style, Title)
        
        If Response = vbYes Then                            ' 사용자가 예를 선택.
            Call SS_INIT(SS, 1, 1)
            Call GotoSpreadSet
            SSPan = ""
            Update_Check = False
            ErrList.Clear
        End If
    Else
        Msg = " 화면을 CLEAR하시겠습니까?"
        Style = vbYesNo + vbQuestion + vbDefaultButton2     ' Define buttons.
        Title = "화면 CLEAR"                                ' 기본 제목.
        Response = MsgBox(Msg, Style, Title)
        If Response = vbYes Then                            ' 사용자가 예를 선택.
            Call SS_INIT(SS, 1, 1)
            Call GotoSpreadSet
            SSPan = ""
            StrGBER = "R"
            GBTransmit = ""
            For i = 0 To 100
                For j = 0 To 50
                    Temp_Jeobsu(i, j) = ""
                    Temp_Request_Code(i, j) = ""
                    Temp_Result(i, j) = ""
                Next j
            Next i
        End If
    End If
    
End Sub


Private Sub mnuSet_Click()
'5)통신환경설정
    Dim CodeCheck

    frmSetComm.Show vbModal
    
    CodeCheck = GetSetting("LabInterface", "SetPC", "GGJCODE" & GGJCODE)
    GGCODE = Mid(CodeCheck, 1, 2)
    If Mid(CodeCheck, 6, 1) = "1" Then
        ComPort = GetSetting("LabInterface", "SetComm", "ComPort1")
        Settings = GetSetting("LabInterface", "SetComm", "ComSettings1")
    ElseIf Mid(CodeCheck, 6, 1) = "2" Then
        ComPort = GetSetting("LabInterface", "SetComm", "ComPort2")
        Settings = GetSetting("LabInterface", "SetComm", "ComSettings2")
    End If
    
    If ComPort = "" Then
        MsgBox " COM PORT 지정이 잘못 되었습니다."
        Exit Sub
    End If
    
    If MSComm1.PortOpen = False Then
        MSComm1.CommPort = ComPort
        MSComm1.Settings = Settings
    End If
    Call GotoSpreadSet

End Sub


Private Sub mnuExit_Click()
'6)종료
    Dim Msg, Style, Title, Response
    If Update_Check = True Then
        Msg = " 수신한 DATA를 저장하지 않았습니다." & vbCrLf & _
              " DATA를 저장한다음 종료하십시요." & vbCrLf & _
              "                                " & vbCrLf & _
              " 프로그램을 종료하시겠습니까?"
        Style = vbYesNo + vbDefaultButton2 + vbCritical     ' Define buttons.
        Title = "프로그램 종료"                             ' 기본 제목.
        Response = MsgBox(Msg, Style, Title)
        If Response = vbYes Then                            ' 사용자가 예를 선택.
            If MSComm1.PortOpen = True Then
                MSComm1.PortOpen = False
            End If
            Update_Check_Force = True                       '강제종료시 사용
            Unload Me
        End If
    Else
        Msg = " 프로그램을 종료하시겠습니까?"
        Style = vbYesNo + vbDefaultButton2 + vbQuestion     ' Define buttons.
        Title = "프로그램 종료"                             ' 기본 제목.
        Response = MsgBox(Msg, Style, Title)
        If Response = vbYes Then                            ' 사용자가 예를 선택.
            If MSComm1.PortOpen = True Then
                MSComm1.PortOpen = False
            End If
            Update_Check_Force = True                       '강제종료시 사용
            Unload Me
        End If
    End If

End Sub


Private Sub lblDate_Click()
    FrmCalendar.Show vbModal
    
    lblDate = FrmCalendar.Caption
    Call GotoSpreadSet
    
End Sub


Private Sub vaSpread_Display(ResultText, ssR1, ssC1)
    SS.Row = ssR1
    SS.Col = ssC1
    SS.Text = ResultText
    
    SS.Col = 6
    SS.Action = SS_ACTION_ACTIVE_CELL
    
End Sub





Sub GetIniFile()
    Dim CodeCheck
'Registry 저장위치
'HKEY_CURRENT_USER\Software\VB and VBA Program Settings\LabInterface

'Rack number / Position number Set
'    If (GetSetting("LabInterface", "SetPc", "MaxRCntNo")) = "" Then
'        Call SaveSetting("LabInterface", "SetPc", "MaxRCntNo", "1")
'    End If
'    GnRCntNo = Val(GetSetting("LabInterface", "SetPc", "MaxRCntNo"))              ' register에서 serial number get
'
'    If GetSetting("LabInterface", "Setpc", "MaxPCntNo") = "" Then
'        Call SaveSetting("LabInterface", "SetPc", "MaxPCntNo", "1")
'    End If
'    GnPCntNo = Val(GetSetting("LabInterface", "SetPc", "MaxPCntNo"))              ' register에서 serial number get
    
    
'통신환경 초기설정 확인및 기본 환경 설정
    
    GGJCODE = "09"          '혈액분석
    
'    Call SaveSetting("LabInterface", "SetPC", "GGJCODE" & GGJCODE, GGJCODE)
    
    CodeCheck = GetSetting("LabInterface", "SetPC", "GGJCODE" & GGJCODE)
    
    GGCODE = Mid(CodeCheck, 1, 2)
    
    If GGCODE = "" Then
        Call mnuSet_Click
    End If
    
    If Mid(CodeCheck, 6, 1) = "1" Then
        ComPort = GetSetting("LabInterface", "SetComm", "ComPort1")
        Settings = GetSetting("LabInterface", "SetComm", "ComSettings1")
    ElseIf Mid(CodeCheck, 6, 1) = "2" Then
        ComPort = GetSetting("LabInterface", "SetComm", "ComPort2")
        Settings = GetSetting("LabInterface", "SetComm", "ComSettings2")
    End If
    
    If MSComm1.PortOpen = False Then
        MSComm1.CommPort = ComPort
        MSComm1.Settings = Settings
    End If
    
    lblPort.Caption = "Com" & ComPort & "," & Settings
    
End Sub


Sub Parmini()                              '프로그램 초기설정 파라미터
    Timer1.Interval = 500                  'Timer1_Timer 1sec
    
    lblDate = GstrSysDate
    
    lblDate.Alignment = 2
    lblDate.FontSize = 14
    lblDate.BorderStyle = 1
    
    lblTime.Alignment = 2
    lblTime.FontSize = 14
    lblTime.BorderStyle = 1
    
    SSPan.FontSize = 13     '14
    
    Update_Check = False
    Update_Check_Force = False
    
    
    '통신용 Definition Character
    SOH = Chr(1)                '<SOH> []
    STX = Chr(2)                '<STX> []
    ETX = Chr(3)                '<ETX> []
    EOT = Chr(4)                '<EOT> []
    ENQ = Chr(5)                '<ENQ>
    ACK = Chr(6)                '<ACK>
    LF = Chr(10)
    FF = Chr(12)                '[]
    CR = Chr(13)
    NACK = Chr(21)              '<NACK>
    ETB = Chr(23)               '<ETB>
    SRS = Chr(30)               '   []
    SGS = Chr(29)               '   []
    SFS = Chr(28)               '   [] ' LINE CHANGE
    
    
End Sub


Sub CodeKy_Search()
    Dim i               As Integer
    Dim ArrCnt          As Integer
    Dim nINDEX          As Integer
    'code key data 검색
    '/ 먼저 ItemMl Table 의 ItmeNm과 장비에서내보내는 일정한 이름들(Full Name으로 배설터군)과 테이블을 일치시켜야함..
    '/ 물론 필드를 하나 더 만들어서 정렬순서를 정하면 좋겠지요...
    '/ 그리고.. 아이템코드가 장비에서 내보내는 데이타와 같은지 점검 하시고... 각 장비별 특성에 맞게
    '/ 테이블의 장비, 순번등을 정하심이...          2000.12.22
    
    strSQL = ""
    strSQL = strSQL & " SELECT CODEKY,GEOMJAN2 , TO_NUMBER(GeomJan3) NSEQNO " & vbLf
    strSQL = strSQL & " FROM TWEXAM_ITEMML                                  " & vbLf
    strSQL = strSQL & " WHERE GEOMJAN1 = '" & GGCODE & "'                   " & vbLf                                   'STAii 장비 code
    strSQL = strSQL & "     AND GEOMJAN3 <> '99'                            " & vbLf        '/검사항목이 정해지지 않은경우 순번을 99로 셋팅 읽지 않음
    strSQL = strSQL & "  ORDER BY TO_NUMBER(GeomJan3)  "
    
    Result = adoSQL(strSQL)
    
    If Result <> 0 Or rowindicator = 0 Then
        MsgBox "CODEKY 검색 ERROR" & vbCrLf & "CODEKY가 없습니다.", vbCritical
        Exit Sub
    End If
    
    ArrCnt = AdoGetNumber(Rs, "NSEQNO", rowindicator - 1)
    
    'Erase InputLineData
    
    ReDim nReciveData(ArrCnt)
    ReDim InputLineData(ArrCnt + 5)     '위아래로 몇개 더 들어오더군... 여유있게 잡아놓는것이....
    
    For i = 0 To rowindicator - 1
        nINDEX = Val(AdoGetString(Rs, "NSEQNO", i))
        With nReciveData(nINDEX)
            .StrKeyCode = AdoGetString(Rs, "CodeKy", i)
            .StrFullName = AdoGetString(Rs, "GEOMJAN2", i)
        End With
    Next i
    MaxRecordCount = rowindicator                                   ' 검사항목 갯수
    SS.MaxCols = rowindicator + 6                                     'Record Count를 check 하여 max columns를 결정한다.
    
End Sub

Sub WorkDisplay(i)
    Select Case i
        Case 0
            timerx = True
            Label1.BackColor = &HC0C0C0
            Label1.FontSize = 15
            Label1.Font = "Arial Black"
            Label1.Caption = "CLINITEK 500"
            Timer1.Tag = ""
            SSPan = "수신이 종료되었습니다."
            
            MnuReceive.Enabled = True
            MnuEnd.Enabled = False
            MnuChange.Enabled = True
            MnuSet.Enabled = True
            MnuExit.Enabled = True
            
            Toolbar1.Buttons(1).Enabled = True
            Toolbar1.Buttons(2).Enabled = False
            Toolbar1.Buttons(3).Enabled = True
            Toolbar1.Buttons(4).Enabled = True
            Toolbar1.Buttons(5).Enabled = True
'           Toolbar1.Buttons(6).Enabled = True
           
'           Frame1.Enabled = True
           
           lblDate.Enabled = True
        Case 1
            timerx = True
            Timer1.Tag = "ON"
            SSPan = "DATA 수신 중입니다."
            
            MnuReceive.Enabled = False
            MnuEnd.Enabled = True
            MnuChange.Enabled = False
            MnuSet.Enabled = False
            MnuExit.Enabled = False
            
            Toolbar1.Buttons(1).Enabled = False
            Toolbar1.Buttons(2).Enabled = True
            Toolbar1.Buttons(3).Enabled = False
            Toolbar1.Buttons(4).Enabled = False
            Toolbar1.Buttons(5).Enabled = False
            lblDate.Enabled = False
    End Select
End Sub


Sub Kdelete()

    Dim i
    Dim Wfile               As String
    Dim Kfile               As String
    Dim Kdate               As Date
    Dim Rdate               As Date

    On Error Resume Next
        
        File1.Pattern = "*.*"
        File1.Path = App.Path & "\intdown"
        Kdate = Format(Date, "yyyy-mm-dd") 'Date
        Rdate = Format(Date, "yyyy-mm-dd") 'Date
        
        For i = 1 To 999
            If File1.ListCount = 0 Then
                Exit For
            End If
            File1.ListIndex = File1.ListIndex + 1
            
            Text1.Text = File1.Path & "\" & File1.FileName
            Kfile = Text1.Text
            Wfile = LTrim$(Right$(ExtractTime(Text1.Text), 21))
            If Mid(Wfile, 1, 2) = "20" Then
                Date = Mid(Wfile, 1, 10)
                Kdate = Format(Date, "yyyy-mm-dd") 'Date
            Else
                Date = Mid(Wfile, 1, 8)
                Kdate = Format(Date, "yyyy-mm-dd") 'Date
            End If
            
            If DateAdd("m", 1, Kdate) < Rdate Then                             'file 작성 날자 check
                Kill (Kfile)                                       '30일 이전에 작성된 file 삭제
            End If
            
            If File1.ListIndex = File1.ListCount - 1 Then
                File1.ListIndex = 0
                Exit For
            End If
        Next i
        
    Date = Rdate

End Sub


Private Sub Timer1_Timer()
    lblTime = Time
End Sub


Private Sub Timer_Picture_Timer()
    If timerx = False Then Exit Sub
    Tcounter = Tcounter + 1
    Select Case Tcounter
        Case 1
                'Image1(0).Visible = True
                Label1.Font = "굴림체"
                Label1.BackColor = &HFFFFC0
                Label1.FontSize = 15
                Label1.Caption = "투 윈 정 보 시 스 템"
        Case 2
                'Image1(0).Visible = False
                Label1.Font = "Arial Black"
                Label1.BackColor = &HC0FFFF
                Tcounter = 0
                Label1.FontSize = 16
                Label1.Caption = "http://twowin.co.kr"
    End Select



End Sub





