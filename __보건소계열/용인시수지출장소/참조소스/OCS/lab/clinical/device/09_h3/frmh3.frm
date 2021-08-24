VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmh3 
   Caption         =   "Technicon H3"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   2115
   ClientWidth     =   11880
   Icon            =   "frmh3.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   WindowState     =   2  '최대화
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  '위 맞춤
      Height          =   660
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   6
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "자료수신을 시작합니다"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "자료수신을 종료합니다"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Key             =   ""
            Object.ToolTipText     =   "화면의 데이타를 DB에 저장합니다"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "화면의 데이타를 지웁니다"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "통신환경을 설정합니다 "
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "프로그램을 종료합니다"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.PictureBox picResult 
      Height          =   6015
      Left            =   6300
      ScaleHeight     =   5955
      ScaleWidth      =   5115
      TabIndex        =   13
      Top             =   1140
      Visible         =   0   'False
      Width           =   5175
      Begin FPSpread.vaSpread SSR 
         Height          =   5910
         Left            =   30
         TabIndex        =   14
         Top             =   30
         Width           =   5055
         _Version        =   196608
         _ExtentX        =   8916
         _ExtentY        =   10425
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
         MaxCols         =   3
         MaxRows         =   50
         SpreadDesigner  =   "frmh3.frx":030A
      End
   End
   Begin FPSpread.vaSpread SS 
      Height          =   6030
      Left            =   240
      TabIndex        =   12
      Top             =   1140
      Width           =   11175
      _Version        =   196608
      _ExtentX        =   19711
      _ExtentY        =   10636
      _StockProps     =   64
      BackColorStyle  =   1
      ColsFrozen      =   6
      EditEnterAction =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   30
      SpreadDesigner  =   "frmh3.frx":0ED6
      UserResize      =   1
      VisibleCols     =   23
      VisibleRows     =   120
   End
   Begin VB.TextBox txtBarCode 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   16
      Top             =   720
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Left            =   8976
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   2184
      Visible         =   0   'False
      Width           =   636
   End
   Begin VB.FileListBox File1 
      Height          =   270
      Left            =   9612
      TabIndex        =   5
      Top             =   2184
      Visible         =   0   'False
      Width           =   996
   End
   Begin VB.Timer Timer_RRequest 
      Left            =   10320
      Top             =   1425
   End
   Begin VB.Timer Timer_RCheck 
      Left            =   9984
      Top             =   1425
   End
   Begin VB.Timer Timer_Picture 
      Left            =   9312
      Top             =   1425
   End
   Begin VB.Timer Timer1 
      Left            =   8976
      Top             =   1425
   End
   Begin VB.Timer Timer_Order 
      Left            =   9648
      Top             =   1425
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
      Height          =   780
      Left            =   4968
      TabIndex        =   3
      Top             =   7248
      Width           =   6444
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
      Left            =   3312
      TabIndex        =   2
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
      Left            =   432
      TabIndex        =   1
      Top             =   2355
      Width           =   2868
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10608
      Top             =   2184
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   11055
      Top             =   1425
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Frame Frame2 
      Height          =   1308
      Left            =   288
      TabIndex        =   8
      Top             =   7152
      Width           =   4524
      Begin Threed.SSPanel SSPan 
         Height          =   492
         Left            =   144
         TabIndex        =   9
         Top             =   168
         Width           =   4236
         _Version        =   65536
         _ExtentX        =   7472
         _ExtentY        =   868
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
         BevelOuter      =   1
         BevelInner      =   2
      End
      Begin VB.Label lblTime 
         Alignment       =   2  '가운데 맞춤
         BorderStyle     =   1  '단일 고정
         ForeColor       =   &H00FF0000&
         Height          =   492
         Left            =   2304
         TabIndex        =   11
         Top             =   720
         Width           =   2076
      End
      Begin VB.Label lblDate 
         BorderStyle     =   1  '단일 고정
         Height          =   492
         Left            =   144
         TabIndex        =   10
         Top             =   720
         Width           =   2076
      End
   End
   Begin VB.Label lblPort 
      Alignment       =   2  '가운데 맞춤
      Height          =   285
      Left            =   4980
      TabIndex        =   17
      Top             =   8100
      Width           =   1965
   End
   Begin VB.Label Label3 
      Caption         =   "BarCode입력"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   240
      TabIndex        =   15
      Top             =   780
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   5250
      Picture         =   "frmh3.frx":2F12
      Top             =   660
      Visible         =   0   'False
      Width           =   480
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   11004
      Top             =   2184
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   6
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmh3.frx":321C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmh3.frx":3536
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmh3.frx":3850
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmh3.frx":3B6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmh3.frx":3E84
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmh3.frx":419E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "date 변경용 obj"
      Height          =   228
      Left            =   8976
      TabIndex        =   4
      Top             =   2712
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   6420
      TabIndex        =   0
      Top             =   660
      Width           =   4935
   End
   Begin VB.Menu mnuOption 
      Caption         =   "옵션(&M)"
      Visible         =   0   'False
      Begin VB.Menu mnuRack 
         Caption         =   "Rack/Position Set"
      End
   End
   Begin VB.Menu mnuReceive 
      Caption         =   "자료수신(&R)"
   End
   Begin VB.Menu mnuEnd 
      Caption         =   "수신종료(&E)"
   End
   Begin VB.Menu mnuWrite 
      Caption         =   "자료저장(&W)"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuClear 
      Caption         =   "화면지우기(&C)"
   End
   Begin VB.Menu mnuSet 
      Caption         =   "환경설정(&S)"
   End
   Begin VB.Menu mnuExit 
      Caption         =   "종료(&X)"
   End
End
Attribute VB_Name = "frmh3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
    Dim RPoint                                  'row 위치 지정용
    Dim CPoint                                  'col 위치 지정용
    Dim RSequence                               'Result Record Sequence Check용
    
'    Dim RResult
    
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
    
    Dim C1                  As Integer          ' work buffer column1
    Dim R1                  As Integer          ' work buffer row1
    
    Dim Temp_Jeobsu(100, 50)            As String           ' Spread의 Data를 Temp_Jeobsu Array로 Move
    Dim Temp_Request_Code(100, 50)      As String           ' input test code A1~Z6
    Dim Temp_Result(100, 50)            As String           ' Result Input Data
    
    Dim Temp_JCount                     As String           ' Spread의 Data를 Temp_Jeobsu Array로 Move
    Dim Temp_Quality(10, 100)           As String
    
    Dim Temp_K(60, 3)       As String           ' item table data 입력용 buffer
    Dim FileClose           As Boolean
    
    Dim LNormal             As Boolean          'working list msg termination flag
    
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
'    Dim RCheckSum           As String
'    Dim RCheckSumD          As String
    
    
'    Dim SBuffer             As String
'    Dim SBufferSum          As String
'    Dim SBufferSumD         As String
    
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
    
'    Dim Order_Data_Seq      As Integer
'    Dim Or_Seq              As Integer
    
    Dim timerx              As Boolean
    Dim PortOpen            As Boolean
    Dim SSCheck             As Boolean
    
    Dim BC                  As String           ' Block Code
    Dim LC                  As Integer          ' Data Line Line Code
    Dim DataLine            As String           ' Data Line Type
    
'    Dim Receive_STX_Check   As Boolean
'    Dim Receive_Data_Check  As Boolean
'    Dim Receive_STA_Check   As Boolean
'    Dim Receive_STA_Seq     As Integer
    
'    Dim RackNo_Result       As String
'    Dim PosiNo_Result       As String
'    Dim Infostr_Result      As String
    
'    Dim LCpos               As Integer
    
'    Dim MaxRackNo
'    Dim MaxPosiNo
'    Dim End_check           As Boolean
'    Dim TimerRNo
'    Dim TimerPNo
'    Dim TimerTCode
    
'    Dim Test_Code           As String           'line code 12의 test code  check
    
'    Dim Error_Message       As String
'    Dim Error_Message_Block As String
    
'    Dim Sample_Result(5)    As String
    Dim RecordCount

'    Dim SendTime
'    Dim SendBuffW           As String
'    Dim SendBuffT           As String

'    Dim R_Check             As Boolean
'    Dim ME_Check            As Boolean
'    Dim MA_Check            As Boolean
    
    'data receive part 정의
'    Dim Mheader
'    Dim Mdate
    
'    Dim Lterminator
    
'    Dim Fcheck              As Integer
    
'    Dim RJeobsuNo           As String
    
'    Dim Pns                 As Integer      'Patient information length set Start
'    Dim Pne                 As Integer      'Patient information length set Start
'    Dim Pinfo1                              'Patient information 16Byte
'    Dim Pinfo2                              'Patient information 12Byte
'    Dim Pinfo3                              'Patient information  6Byte
'    Dim Pinfo4                              'Patient information  4Byte
    
'    Dim Optno               As String       '수신 Data에서 ptno를 check하여 move
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




Private Sub mnuRack_Click()
'    frmRackCnt.Show vbModal
'    Call GotoSpreadSet                            ' spread cell active

End Sub



Private Sub picResult_Click()
        picResult.Visible = False

End Sub

Private Sub SSR_Click(ByVal Col As Long, ByVal Row As Long)
'        picResult.Visible = False

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    On Error Resume Next
    Select Case Button.Index
            Case 1
                   Call mnuReceive_Click
            Case 2
                   Call mnuEnd_Click
            Case 3
                   Call mnuWrite_Click
            Case 4
                   Call mnuClear_Click
            Case 5
                   Call mnuSet_Click
            Case 6
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
    
    
    Call DbAdoConnect("TW_MIS_EXAM", "HOSPITAL", "kuh2")
    
    SSPan.Caption = "Server 컴퓨터에 접속되었습니다."
    SSPan.ForeColor = Val("&H000000FF&")
    
    Call Parmini                            ' spread 초기화 작업
    Call vaSpread_Clear(SS, 1, 1, 0, 0)
    Call GetIniFile
    
    Call CodeKy_Search                      ' codeky Read from twexam_itemml
    
    Label1.FontSize = 24
    Label1.BorderStyle = 0
    Label1.Caption = " Technicon H3 "
    
    For i = 7 To SS.MaxCols                 ' spread header를 약어 code로 초기화
        SS.Row = 0
        SS.Col = i
        If Temp_K(i - 6, 1) <> "" Then
            SS.Text = Temp_K(i - 6, 1)
        Else
            SS.Text = "_"
        End If
    
    Next i
    
    For i = 1 To SS.MaxRows
        For j = 1 To 6
            SS.Row = i
            SS.Col = j
            SS.Lock = True
        Next j
    Next i
    
'    For i = 1 To MaxRecordCount                 ' spread header를 약어 code로 초기화
'        SSR.Col = 1
'        SSR.Row = i
'        SSR.Text = "   " & Temp_K(i, 0)
'
'        SSR.Col = 2
'        SSR.Row = i
'        SSR.Text = "   " & Temp_K(i, 1)
'    Next i
'    SSR.MaxRows = MaxRecordCount
    
    Call Kdelete                                '30일 경과된 누적 file 삭제
            
    txtBarCode.SetFocus

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
    Timer_Order.Interval = 0                  'Timer_order_Timer End
    Timer_RCheck.Interval = 0                   'Timer_RCheck_Timer End
    timerx = False
    
    Call WorkDisplay(0)
    Close #hSaveFile
    FileClose = True

End Sub


Private Sub SS_Click(ByVal Col As Long, ByVal Row As Long)
    
    If Col = 3 Or Col = 4 Then
        picResult.Visible = True
       i = 0
       j = 0
       
       SSR.MaxRows = 0
       
       SS.Col = 6
       SS.Row = Row
       
       SSR.MaxRows = Val(SS.Text)
       
       'Temp_Jeobsu(RPoint, CPoint)
        For i = 1 To MaxRecordCount
            If Temp_Result(Row, i + 10) <> "" Then
                j = j + 1
                
                SSR.Col = 1
                SSR.Row = j
                SSR.Text = "  " & Temp_K(i, 0)
                
                SSR.Col = 2
                SSR.Row = j
                SSR.Text = "  " & Temp_K(i, 1)
                
                
                SSR.Col = 3
                SSR.Row = j
                SSR.Text = Temp_Result(Row, i + 10)
            End If
        Next i
'        SSR.MaxRows = MaxRecordCount
        SSR.MaxRows = SSR.DataRowCnt
    Else
        SSR.Col = 1:    SSR.Col2 = SSR.MaxCols
        SSR.Row = 1:    SSR.Row2 = SSR.DataRowCnt
         
        SSR.BlockMode = True
        SSR.Action = SS_ACTION_CLEAR_TEXT
        SSR.BlockMode = False
        picResult.Visible = False
    End If
    
    
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


Private Sub SS_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    
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

End Sub


Private Sub SS_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim SS_M_col            As Long             ' spread mouse down col val
    Dim SS_M_row            As Long             ' spread mouse down row val
    
    Dim Msg, Style, Title, Response
    
    If Receive_Check = True Then
        Exit Sub
    End If
    
    Call SS.GetCellFromScreenCoord(SS_M_col, SS_M_row, x, y)
    
    If (SS_M_col <> 1 Or SS_M_row > SS.DataRowCnt) And Button = vbRightButton Then
        For i = 1 To 50
            Beep
        Next i
        Exit Sub
    End If
    
    If Update_Check = True Then Exit Sub
    
    If Button = vbRightButton Then                                      ' Value = 2
'        Call SS.GetCellFromScreenCoord(SS_M_col, SS_M_row, x, y)
        SS.Col = SS_M_col
        SS.Row = SS_M_row
        If SS.ActiveRow = SS_M_row And SS.ActiveCol = 1 And SS_M_row <= SS.DataRowCnt Then
            Msg = SS_M_row & " 번째 DATA를 삭제 하시겠습니까?" & vbCrLf & _
                             " DATA를 확인하셨습니까?"
            Style = vbYesNo + vbQuestion + vbDefaultButton2             ' Define buttons.
            Title = "DATA 삭제"                                         ' 기본 제목.
            Response = MsgBox(Msg, Style, Title)
            If Response = vbYes Then                                    ' 사용자가 예를 선택.
                SS.Action = SS_ACTION_DELETE_ROW                        ' value = 5
                
                'Work 변수 재편성 작업
                
            End If
        End If
    End If
    
End Sub



'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Private Sub mnuReceive_Click()
'1)검사시작
    On Error Resume Next

    Receive_Check = True
    
    If SS.DataRowCnt = 0 Then       ' Transmition of Patient File to STA
        strBiDirect_Trans = False   ' batch mode
    Else
        strBiDirect_Trans = True    ' Query Mode
    End If
    
    RSequence = 0
    QCounter = 0
    RCounter = 1
    

    ErrList.Clear
    
'    R_Check = False
'    ME_Check = False
'    MA_Check = False
    
    timerx = True                                           '수신표시 정지용 flag
    FileClose = False
    Timer_Picture.Interval = 1000                            'Timer_Picture_Timer 1000mS
    N = 0
    
    SColumn = 6                                             'spread sheet 초기화 위치
    SRow = 1
    Call SS_INIT(SS, SColumn, SRow)

'******************** Part 3 **********************************
'*    Output용 File Name Set Open/Save 처리                   *
'**************************************************************
    
    On Error GoTo ErrorMsg
    Ser = Ser + 1
    CommonDialog1.InitDir = "C:\intdown"
    CommonDialog1.FileName = "H" & Format$(lblDate, "yyyymmdd") & Ser
    
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
    Call WorkDisplay(1)                     '  "파일로 수신 중입니다." MsgBox Display
    
    If MSComm1.PortOpen = False Then
        MSComm1.InBufferSize = 8192         ' InBufferSize 변경은 portopen = false일 경우만 가능  ' default = 1024
        MSComm1.PortOpen = True
    End If
    
    MSComm1.RThreshold = 1                  ' MSComm 컨트롤은 수신 버퍼에 한 문자가 들어 올 때마다 OnComm 이벤트를 발생시킵니다.
    MSComm1.InputLen = 1
    
'******************** Part 4 **********************************
'*  Output File OPen                                          *
'**************************************************************
    hSaveFile = FreeFile
    Open temp_file For Append As hSaveFile
        If Err Then
            MsgBox Error$, vbExclamation
            Close hSaveFile
            hSaveFile = 0
            Call WorkDisplay(0)
            Exit Sub
        End If
    
    Timer_Order.Interval = 30000           'Order 를 30초 단위로 전송
    
    SSPan = "DATA 수신 대기중입니다."
    SSPan.ForeColor = &H0&                  ' black  &H00000000&

Exit Sub

ErrCancel:
    MsgBox "자료수신을 취소하였습니다."
    Call ErrCancel_Click                    ' 자료수신 종료
Exit Sub


ErrorMsg:
    MsgBox "Error " & "Code = " & Err.Number & vbLf & vbLf & Err.Description
    If FileClose = False Then
        Close #hSaveFile                    ' error 발생시 file close
    End If


'    SS.Enabled = True
    Timer_Picture.Interval = 0              'Timer_Picture_Timer End
    Timer_Order.Interval = 0              'Timer_order_Timer End
    timerx = False                          '수신표시 정지용 flag
    
    Call WorkDisplay(0)
    If MSComm1.PortOpen = True Then
        MSComm1.PortOpen = False
    End If

End Sub


'{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}
'{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}
'{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}
'{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}

Private Sub MSComm1_OnComm() ' DATA 수신 처리
    Dim EVMsg$
    Dim ERMsg$
    Select Case MSComm1.CommEvent           'CommEvent 속성에 따른 항목
       '이벤트 메시지
        Case comEvReceive                   ' 포트로부터 데이터가 들어왔음...
             RBuffer = MSComm1.Input
        Case Else:              ERMsg$ = " 오류 이벤트"
    End Select
    
    Select Case RBuffer                                 ' Message Block 단위로 편집 출력
           Case STX:    STXBuffer = True                ' STX [] Check용
           Case ETX:    ETXBuffer = True                ' ETX [] Check용
    End Select
    
    RBufferSum = RBufferSum & RBuffer                   ' comm port 에서 입력한 data 누적
    
    If ETXBuffer = True Then                            'ETX Check
        If FileClose = False Then
            Print #hSaveFile, "[ Rx " & Format(Time, "hh:mm:ss") & " ]"
            Print #hSaveFile, RBufferSum   ' write omitting cr lf
        End If
        
        Call DataReceive(RBufferSum)                                '수신한 data record 분석
        
'        DataReceive strTemp
        
        RBufferSum = ""
        RBuffer = ""
        ETXBuffer = False
'        Call Ack_Send
    End If
    
End Sub

Private Function DataReceive(strPrmData As String) As Boolean
    
    Dim RJeobsuPT
    Dim Ritemcd1
    Dim RResultU        As String

    On Error Resume Next

    If Len(strPrmData) <> 335 Then Exit Function
    
    Dim udH3receive As H3ReceiveRecord
    
    H3_InitReceiveLength udH3receive
    H3_ReadReceiveBuffer strPrmData, udH3receive
        
    
    RSequence = RSequence + 1                                'Result Record Sequence Check용
    
    Print #hSaveFile, ""
    For i = 1 To 39
        '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
        Print #hSaveFile, Format(i, "@@") & " " & _
                          Format(udH3receive.intPosition(i), "@@@") & " " & _
                          udH3receive.intLength(i) & " " & _
                          udH3receive.strRecord(i)
        '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
        
        SS.Col = 1
        SS.Row = RSequence
        RJeobsuPT = Trim(SS.Text)
        
        RPoint = RSequence                                              'row 위치 지정용
        
        Select Case i
               Case 1, 2, 3, 4, 5, 6, 7
                    Ritemcd1 = Temp_K(i, 0)                             'ITEM CODE
                    If Val(udH3receive.strRecord(i)) < 1 Then
                        RResultU = Val(udH3receive.strRecord(i))        '결과
                    Else
                        RResultU = udH3receive.strRecord(i)
                    End If
                    CPoint = i + 10                                     'col 위치 지정용
                    Temp_Result(RPoint, CPoint) = RResultU
                    Call Save_Result(RJeobsuPT, Ritemcd1, RResultU)
               
               Case 8
                    Ritemcd1 = Temp_K(9, 0)
                    If Val(udH3receive.strRecord(i)) < 1 Then
                        RResultU = Val(udH3receive.strRecord(i))        '결과
                    Else
                        RResultU = udH3receive.strRecord(i)
                    End If
                    CPoint = 19
                    Temp_Result(RPoint, CPoint) = RResultU
                    Call Save_Result(RJeobsuPT, Ritemcd1, RResultU)
               
               Case 9
                    Ritemcd1 = Temp_K(8, 0)
                    If Val(udH3receive.strRecord(i)) < 1 Then
                        RResultU = Val(udH3receive.strRecord(i))        '결과
                    Else
                        RResultU = udH3receive.strRecord(i)
                    End If
                    CPoint = 18
                    Temp_Result(RPoint, CPoint) = RResultU
                    Call Save_Result(RJeobsuPT, Ritemcd1, RResultU)
               
               Case 10
                    Ritemcd1 = Temp_K(13, 0)
                    If Val(udH3receive.strRecord(i)) < 1 Then
                        RResultU = Val(udH3receive.strRecord(i))        '결과
                    Else
                        RResultU = udH3receive.strRecord(i)
                    End If
                    CPoint = 23
                    Temp_Result(RPoint, CPoint) = RResultU
                    Call Save_Result(RJeobsuPT, Ritemcd1, RResultU)
               
               Case 11
                    Ritemcd1 = Temp_K(12, 0)
                    If Val(udH3receive.strRecord(i)) < 1 Then
                        RResultU = Val(udH3receive.strRecord(i))        '결과
                    Else
                        RResultU = udH3receive.strRecord(i)
                    End If
                    CPoint = 22
                    Temp_Result(RPoint, CPoint) = RResultU
                    Call Save_Result(RJeobsuPT, Ritemcd1, RResultU)
               
               Case 12
                    Ritemcd1 = Temp_K(10, 0)
                    If Val(udH3receive.strRecord(i)) < 1 Then
                        RResultU = Val(udH3receive.strRecord(i))        '결과
                    Else
                        RResultU = udH3receive.strRecord(i)
                    End If
                    CPoint = 20
                    Temp_Result(RPoint, CPoint) = RResultU
                    Call Save_Result(RJeobsuPT, Ritemcd1, RResultU)
               
               Case 19, 20, 21, 22, 23, 24
                    Ritemcd1 = Temp_K(i - 5, 0)
                    If Val(udH3receive.strRecord(i)) < 1 Then
                        RResultU = Val(udH3receive.strRecord(i))        '결과
                    Else
                        RResultU = udH3receive.strRecord(i)
                    End If
                    CPoint = i + 5
                    Temp_Result(RPoint, CPoint) = RResultU
                    Call Save_Result(RJeobsuPT, Ritemcd1, RResultU)
        
        End Select
        
    Next i
    

End Function


'Private Sub Save_Result(ByVal Rresult1 As String, ByVal itemcd1 As String, ByVal Optno1 As String)
Private Sub Save_Result(JeobsuPT, itemcd1, resultu)
    
'    Exit Sub
    
    Dim Bdt
    Dim Bno1
    Dim Bno2
        
    Bdt = convLabnoToExpand(Mid(JeobsuPT, 1, 5))
    Bno1 = Mid(JeobsuPT, 6, 2)
    Bno2 = Mid(JeobsuPT, 8, 5)
    
    adoConnect.BeginTrans                          ' TRANSACTION의 종료시에 COMMITTRANS를 지정함
    
    strSQL = ""
    strSQL = strSQL & "UPDATE TWEXAM_GENERAL_SUB "
    strSQL = strSQL & "   SET RESULT1  =   '" & resultu & "'"
    strSQL = strSQL & " WHERE JEOBSUDT =  TO_DATE('" & Bdt & "','YYYY-MM-DD') "    '입력된 날자로 검색
    strSQL = strSQL & "   AND SLIPNO1  =   '" & Bno1 & "' "                        '구분
    strSQL = strSQL & "   AND SLIPNO2  =   '" & Bno2 & "' "                        '구분
    strSQL = strSQL & "   AND ITEMCD   =    '" & itemcd1 & "'"                     'ITEMCODE
    strSQL = strSQL & "   AND VERIFY   =  'N'"                                    ' 접수결과에서 VERIFY OK한경우에는 UPDATE하지않음
    
    Result = AdoExecute(strSQL)
    If Result = True And Rowindicator > 0 Then
'        SSPan = "DATABASE에 저장 되었습니다. ( " & RecordCountSum & " 건)"
        SSPan = "DATABASE에 저장 되었습니다. "
        
        adoConnect.CommitTrans                                                   ' TRANSACTION 종료시에 COMMIT 시킴
    Else
        ErrList.AddItem "    Verify Data       " & JeobsuPT
        ErrList.AddItem "    or Update Error   " & itemcd1 & "  " & resultu
        ErrList.ListIndex = ErrList.ListCount - 1
        
        'file write routine insert
        
        adoConnect.RollbackTrans                                                 ' TRANSACTION ERROR시 ROLLBACK 시킴
        SSPan = "DB에 갱신중 ERROR가 발생하였습니다." & vbCrLf & _
                "VERIFY된 DATA인지 확인하십시요."
    End If
    
    
    If strBiDirect_Trans = True Then
        SS.Row = Val(RPoint)
        SS.Col = Val(CPoint) - 4
        SS.Text = Temp_Result(RPoint, CPoint)
    
        SS.Row = Val(RPoint)
        SS.Col = 6
        SS.Text = Val(SS.Text) + 1
        
        If CPoint >= 29 Then
            SS.Col = 6
            SS.BackColor = RGB(0, 255, 0)
            
            'general_sub update
            Call Save_Result_Flag(JeobsuPT)
        
        End If
    Else
        
        SS.Row = Val(RPoint)
        SS.Col = Val(CPoint) - 4
        SS.Text = Temp_Result(RPoint, CPoint)
    
        SS.Row = Val(RPoint)
        SS.Col = 6
        SS.Text = Val(SS.Text) + 1
        
        If CPoint >= 29 Then
            SS.Col = 6
            SS.BackColor = RGB(0, 255, 0)
            
            'general_sub update
            Call Save_Result_Flag(JeobsuPT)
        
        End If
        
'        SS.Row = Val(RPoint)
'
'        SS.Col = 1
'        SS.Text = JeobsuPT
'
'        SS.Col = Val(CPoint) - 4
'        SS.Text = resultu
'
'        SS.Col = 6
'        SS.Text = Val(SS.Text) + 1
'
    End If
    
    
End Sub


Private Sub ini_check()
    Dim Rs                  As ADODB.Recordset
    
    Dim Bjeobsudt
    Dim Bslipno1
    Dim Bslipno2
    
    Dim Verify_Check
'******************** Part 1 **********************************
'*      Spread의 pt no와 slipno2를 Temp_Jeobsu Array로 Move   *
'**************************************************************
     
    R1 = RCounter
        C1 = 1
        
        SS.Row = R1
        SS.Col = C1
        SS.Text = Temp_Jeobsu(R1, C1)
        
                       'Rn  Cn
        Bjeobsudt = convLabnoToExpand(Mid(Temp_Jeobsu(R1, C1), 1, 5))
        Bslipno1 = Mid(Temp_Jeobsu(R1, C1), 6, 2)
        Bslipno2 = Mid(Temp_Jeobsu(R1, C1), 8, 5)
            
        For C1 = 2 To 4
            SS.Row = R1
            Select Case C1
                   Case 2   'date
                        SS.Col = C1
                        Temp_Jeobsu(R1, C1) = Bjeobsudt
                        SS.Text = Temp_Jeobsu(R1, C1)
                   
                   Case 3   'PTNO
                        SS.Col = C1
                        Temp_Jeobsu(R1, C1) = PTNOSearch(Temp_Jeobsu(R1, 1))
                        SS.Text = Temp_Jeobsu(R1, C1)
                   Case 4   'Name
                        SS.Col = C1
                        Temp_Jeobsu(R1, C1) = NameSearch(Temp_Jeobsu(R1, 3))
                        SS.Text = Temp_Jeobsu(R1, C1)
            End Select

        Next C1
        
        strSQL = ""
        strSQL = strSQL & " SELECT ITEMCD, GEOMJAN2, GEOMJAN3,GBER "
        strSQL = strSQL & "   FROM TWEXAM_GENERAL_SUB A, "                   ' 검사접수결과 세부사항
        strSQL = strSQL & "        TWEXAM_ITEMML B, "                        ' 검사 ITEM MASTER
        strSQL = strSQL & "        TWEXAM_GENERAL C "                        ' 검사접수결과
        strSQL = strSQL & "  WHERE A.JEOBSUDT = TO_DATE('" & Bjeobsudt & "','YYYY-MM-DD')"
        strSQL = strSQL & "    AND A.SLIPNO1 =   '" & Bslipno1 & "'"        ' 일련번호
        strSQL = strSQL & "    AND A.SLIPNO2 =   '" & Bslipno2 & "'"        ' 일련번호
        strSQL = strSQL & "    AND A.ITEMCD = B.CODEKY "
        strSQL = strSQL & "    AND B.GBROUTINE = 'I'   "
        strSQL = strSQL & "    AND A.PTNO = C.PTNO "
        strSQL = strSQL & "    AND A.JEOBSUDT = C.JEOBSUDT "
        strSQL = strSQL & "    AND A.SLIPNO1 = C.SLIPNO1 "
        strSQL = strSQL & "    AND A.SLIPNO2 = C.SLIPNO2 "
        strSQL = strSQL & "    AND B.GEOMJAN1 = '" & GGJCODE & "' "
        
        Result = AdoOpenSet(Rs, strSQL)
        
        'Debug.Print Rowindicator
        
        If Result Then
            Do While Not Rs.EOF
                If Val(Trim(Rs.Fields("GEOMJAN2") & "")) >= "11" Then
                    Temp_Jeobsu(R1, Val(Trim(Rs.Fields("GEOMJAN2") & ""))) = Trim(Rs.Fields("ITEMCD") & "")
                    If Trim(Rs.Fields("GBER") & "") = "E" Then
                        StrGBER = "S"
                    Else
                        StrGBER = "R"
                    End If
                
                End If

                Rs.MoveNext
            Loop
        End If
        
    If StrGBER = "S" Then
        SS.Row = R1
        For i = 1 To 6
            SS.Col = i
            SS.ForeColor = RGB(255, 0, 0)
        Next i
    End If
    
    SS.Col = 5
    Temp_Jeobsu(R1, 5) = Rowindicator
    SS.Text = Rowindicator

    SS.SetFocus                             ' cell active 상태로 변경
    SS.Action = SS_ACTION_ACTIVE_CELL       ' 지정된 위치로 cursor 이동

End Sub


Private Sub Save_Result_Flag(JeobsuPT2)

    Dim Bdt
    Dim Bno1
    Dim Bno2
        
    Bdt = convLabnoToExpand(Mid(JeobsuPT2, 1, 5))
    Bno1 = Mid(JeobsuPT2, 6, 2)
    Bno2 = Mid(JeobsuPT2, 8, 5)
    
    strSQL = ""
    strSQL = strSQL & " SELECT JEOBSUDT, SLIPNO1, SLIPNO2, STATUS "
    strSQL = strSQL & "   FROM TWEXAM_GENERAL"
    strSQL = strSQL & "  WHERE JEOBSUDT = TO_DATE('" & Bdt & "','YYYY-MM-DD')"
    strSQL = strSQL & "    AND SLIPNO1 =   '" & Bno1 & "'"
    strSQL = strSQL & "    AND SLIPNO2 =   '" & Bno2 & "'"
    strSQL = strSQL & "    AND (STATUS  = 'R' OR STATUS = 'U') "
    
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
        If Result = True And Rowindicator > 0 Then
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
        
    For i = 1 To SS.MaxRows
        For j = 1 To 6
            SS.Row = i
            SS.Col = j
            SS.Lock = False
        Next j
    Next i
    
'        SS.Enabled = True
    Timer_Picture.Interval = 0                       'Timer_Picture_Timer End
    Timer_Order.Interval = 0                       'Timer_order_Timer End
    timerx = False
    
    Call WorkDisplay(0)

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    Print #hSaveFile, " "
    ResultText = ""
    For R1 = 1 To MaxDataRowCnt
        For C1 = 7 To SS.MaxCols + 6
            If Temp_Result(R1, C1) <> "" Then
                ResultText = ResultText & "  " & C1 & " = " & Temp_Result(R1, C1)
            End If
        Next C1
        Print #hSaveFile, Format$(R1, "000") & " " & ResultText
        ResultText = ""
    Next R1
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    Print #hSaveFile, " " & vbCrLf & "@@@@@  Spread Data" & vbCrLf & " "
    ResultText = ""
    For R1 = 0 To MaxDataRowCnt
        For C1 = 1 To SS.MaxCols
            SS.Row = R1
            SS.Col = C1
            If C1 <> 3 Then
                If SS.Text = "" Then SS.Text = "0"
                ResultText = ResultText & Format$(Trim$(SS.Text), "@@@@@@@@@@@@@@") & " : "
            End If
        Next C1
        Print #hSaveFile, ResultText
        ResultText = ""
    Next R1
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    
    If FileClose = False Then
        Close #hSaveFile
    End If
    FileClose = True
    If MSComm1.PortOpen = True Then
        MSComm1.PortOpen = False
    End If

End Sub


Private Sub mnuWrite_Click()
'3)자료저장
'Spread sheet의 data를 Server에 저장
    
    If strBiDirect_Trans = True Then
        Exit Sub
    End If
    
    On Error Resume Next
    
    If SS.DataRowCnt < 1 Then Exit Sub
    Dim Msg, Style, Title, Response
    Msg = " 자료를 DATABASE에 " & vbCrLf & "저장하시겠습니까?"
    Style = vbYesNo + vbQuestion + vbDefaultButton2 ' Define buttons.
    Title = "DATABASE UPDATE"                       ' 기본 제목.
    Response = MsgBox(Msg, Style, Title)
    If Response = vbYes Then                        ' 사용자가 예를 선택.
        Call Data_Update                            ' 임상병리 검사종류 11(생화학)검사에대한 ITEM CODE 검색 & SET
    End If

End Sub


Private Sub Data_Update()
    Dim Rs                  As ADODB.Recordset
 
    SSPan = "DATABASE에 저장하고 있습니다."
    Pflag = False
    JeobsuCheck = True

''''''''''''''''''''''''''''''''''''''''''''''' TRANSACTION 의 시작위치 지정
    adoConnect.BeginTrans                          ' TRANSACTION의 종료시에 COMMITTRANS를 지정함
    
    ' DATABASE UPDATE                                                                       '
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '임상병리 검사종류 11(생화학 자동분석)검사에대한 ITEM CODE SETTING
    ' Temp_Jeobsu(R1,C1)에 settting 되어있음
    
    For R1 = 1 To MaxDataRowCnt
        For C1 = 7 To MaxRecordCount + 6        'itemcd를 check하기위한 for next
            If Trim$(Temp_Result(R1, C1)) <> "" Then
            
'                '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'                List3.AddItem C1 & "  " & Temp_Result(R1, C1)
'                List3.ListIndex = List3.ListCount - 1
'                Debug.Print Temp_Result(R1, C1)
                '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
                
                strSQL = ""
                strSQL = strSQL & "UPDATE TWEXAM_GENERAL_SUB1 "
                strSQL = strSQL & "   SET RESULT1  =   '" & Format(Val(Temp_Result(R1, C1)), "###0.000") & "'"
                strSQL = strSQL & " WHERE JEOBSUDT =  TO_DATE('" & Temp_Jeobsu(R1, 4) & "','YYYY-MM-DD') "      '입력된 날자로 검색
                strSQL = strSQL & "   AND SLIPNO1  =   11"                                                      '구분
                strSQL = strSQL & "   AND VERIFY   =  'N'"                                                      ' 접수결과에서 VERIFY OK한경우에는 UPDATE하지않음
                strSQL = strSQL & "   AND SLIPNO2  =     " & Temp_Jeobsu(R1, 2)                                 '일일 2건이상일 경우 CHECK
                strSQL = strSQL & "   AND PTNO     =    '" & Trim$(Temp_Jeobsu(R1, 1)) & "'"                    'PATIENT NUMBER
                strSQL = strSQL & "   AND ITEMCD   =    '" & Trim$(Temp_K(C1 - 6, 0)) & "'"                     'ITEMCODE
                
                Result = AdoExecute(strSQL)
                If Result >= 0 And Rowindicator > 0 Then
                    RecordCountBit = 1
                ElseIf Result = -1 Then
                    MsgBox "Check Error" & vbCrLf & R1 & "번째 data를 확인하십시요 "
                    JeobsuCheck = False
                End If
            End If
        Next C1
        RecordCountSum = RecordCountSum + RecordCountBit
        RecordCountBit = 0
    Next R1
          
    If Result Then
        SSPan = "DATABASE에 저장 되었습니다. ( " & RecordCountSum & " 건)"
        If RecordCountSum = 0 Then SSPan = " 저장된 Data가 없습니다."
        adoConnect.CommitTrans                                                   ' TRANSACTION 종료시에 COMMIT 시킴
        RecordCountSum = 0
        Update_Check = False
    Else
        MsgBox "   Update Error     "
        adoConnect.RollbackTrans                                                 ' TRANSACTION ERROR시 ROLLBACK 시킴
        SSPan = "DATABASE에 갱신중 ERROR가 발생하였습니다."
        Update_Check = False
    End If
    
End Sub


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
            txtBarCode.Text = ""
            txtBarCode.SetFocus
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
            txtBarCode.Text = ""
            txtBarCode.SetFocus
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
    End If

End Sub


Private Sub lblDate_Click()
    Call FrmCalendar.Calendar_Show(lblDate)
    Call GotoSpreadSet
    
End Sub


Private Sub vaSpread_Display(ResultText, ssR1, ssC1)
    SS.Row = ssR1
    SS.Col = ssC1
    SS.Text = ResultText
    
    SS.Col = 6
    SS.Action = SS_ACTION_ACTIVE_CELL
    
End Sub

Sub vaSpread_Clear(SS, SColumn, SRow, EColumn, ERow)

    SS.Col = SColumn: SS.Col2 = SS.MaxCols
    SS.Row = SRow: SS.Row2 = -1
     
    SS.BlockMode = True
    SS.Action = SS_ACTION_CLEAR_TEXT
    SS.BlockMode = False
    
    SS.Col = 1:     SS.Row = 1
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
    
    GGJCODE = "09"
    
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
    
    lblDate = SysDate_Get
    
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
    NACK = Chr(21)              '<NACK>
    ETB = Chr(23)               '<ETB>
    
    
End Sub


Sub CodeKy_Search()
    Dim Rs                  As ADODB.Recordset
    
    'code key data 검색
    strSQL = ""
    strSQL = strSQL & " SELECT CODEKY,YAGEO,GEOMJAN2,GEOMJAN3 "
    strSQL = strSQL & "   FROM TWEXAM_ITEMML "
    strSQL = strSQL & "  WHERE GEOMJAN1 = " & GGCODE                                  'STAii 장비 code
    strSQL = strSQL & "  ORDER BY CODEKY ASC "
    
    If AdoOpenSet(Rs, strSQL) Then
        R1 = 0
        Do Until Rs.EOF
            Temp_K(R1 + 1, 0) = Trim$(Rs.Fields("CODEKY") & "")         'CODEKY
            Temp_K(R1 + 1, 1) = Trim$(Rs.Fields("YAGEO") & "")          'YAGEO
            
            Temp_K(R1 + 1, 2) = Trim$(Rs.Fields("GEOMJAN2") & "")       'GEOMJAN2 (Serial number or Position number)
            Temp_K(R1 + 1, 3) = Trim$(Rs.Fields("GEOMJAN3") & "")       'GEOMJAN3 (test code)
            
            Rs.MoveNext: R1 = R1 + 1
        Loop
        MaxRecordCount = Rowindicator                                   ' 검사항목 갯수
    Else
        MsgBox "CODEKY 검색 ERROR" & vbCrLf & "CODEKY가 없습니다.", vbCritical
    End If
    SS.MaxCols = Rowindicator + 6                                     'Record Count를 check 하여 max columns를 결정한다.
    
End Sub



Sub WorkDisplay(i)
    Select Case i
        Case 0
           Timer1.Tag = ""
           SSPan = "수신이 종료되었습니다."
           
           mnuReceive.Enabled = True
           mnuEnd.Enabled = True
           mnuWrite.Enabled = True
           mnuClear.Enabled = True
           mnuSet.Enabled = True
           mnuExit.Enabled = True
           
           Toolbar1.Buttons(1).Enabled = True
           Toolbar1.Buttons(3).Enabled = True
           Toolbar1.Buttons(4).Enabled = True
           Toolbar1.Buttons(5).Enabled = True
           Toolbar1.Buttons(6).Enabled = True
           
'           Frame1.Enabled = True
           
           lblDate.Enabled = True
        Case 1
           Timer1.Tag = "ON"
           SSPan = "DATA 수신 중입니다."
           
           mnuReceive.Enabled = False
'          mnuEnd.Enabled = False
           mnuWrite.Enabled = False
           mnuClear.Enabled = False
           mnuSet.Enabled = False
           mnuExit.Enabled = False
           
           Toolbar1.Buttons(1).Enabled = False
           Toolbar1.Buttons(3).Enabled = False
           Toolbar1.Buttons(4).Enabled = False
           Toolbar1.Buttons(5).Enabled = False
           Toolbar1.Buttons(6).Enabled = False
           
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
        File1.Path = "c:\intdown"
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
                Image1(0).Visible = True
        Case 2
                Image1(0).Visible = False
                Tcounter = 0
    End Select



End Sub



Private Sub Dir_Open()
    
    On Error GoTo ErrorMsg
    Ser = Ser + 1
    CommonDialog1.InitDir = "C:\Vitros"
    CommonDialog1.FileName = "V" & Format$(lblDate, "yyyymmdd") & Ser
    
    CommonDialog1.Filter = "All Files (*.*)|*.*|" & _
                           "Text Files" & "(*.txt)|*.txt|"
    CommonDialog1.FilterIndex = 2
    CommonDialog1.CancelError = True                        '취소확인을 위해 사용
    
    On Error GoTo ErrCancel
'    CommonDialog1.ShowSave                                  'dialog show
    CommonDialog1.CancelError = True                        'cancel error reset
    On Error GoTo ErrorMsg
    
    DoEvents
    Me.Show
    
    temp_file_Order = CommonDialog1.FileName
    
    Exit Sub

ErrCancel:
    MsgBox "자료수신을 취소하였습니다."
'    Call ErrCancel_Click                    ' 자료수신 종료
Exit Sub


ErrorMsg:
    MsgBox "Error " & "Code = " & Err.Number & vbLf & vbLf & Err.Description
    If FileClose = False Then
        Close #hSaveFile                    ' error 발생시 file close
    End If
Return

End Sub


Private Sub File_Open()
    
    On Error GoTo ErrorMsg
    
    hSaveFile_Order = FreeFile
    Open temp_file_Order For Append As hSaveFile_Order
        If Err Then
            MsgBox Error$, vbExclamation
            Close hSaveFile_Order
            hSaveFile_Order = 0
            Exit Sub
        End If
    
Exit Sub
     

ErrorMsg:
    MsgBox "Error " & "Code = " & Err.Number & vbLf & vbLf & Err.Description
    If FileClose = False Then
        Close #hSaveFile_Order
    End If
Return

End Sub


Private Sub txtBarCode_Click()
    txtBarCode.SelStart = 0
    txtBarCode.SelLength = Len(txtBarCode.Text)

End Sub

Private Sub txtBarCode_GotFocus()
    txtBarCode.SelStart = 0
    txtBarCode.SelLength = Len(txtBarCode.Text)

End Sub


Private Sub txtBarCode_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Dim Rs                  As ADODB.Recordset
    
    Dim strPt               As String
    Dim Bjeobsudt           As String
    Dim Bslipno1            As String
    Dim Bslipno2            As String
    
    Dim Checkdouble         As String
    Dim Ptnolen
    Dim Temp_ptno
    Dim Temp_name
    
    Dim Temp_Jeobsu
    
    
    If KeyCode = 13 Then
        If Len(txtBarCode.Text) <> 12 Then
            txtBarCode.SelStart = 0
            txtBarCode.SelLength = Len(txtBarCode.Text)
            txtBarCode.SetFocus
            MsgBox " 접수Number Length Error "
            Exit Sub
        End If
        If Len(txtBarCode.Text) = 12 Then
        
            SS.Col = 1
            SS.Row = SS.DataRowCnt
            Temp_Jeobsu = SS.Text
            
            If Temp_Jeobsu = txtBarCode.Text Then
                txtBarCode.SelStart = 0
                txtBarCode.SelLength = Len(txtBarCode.Text)
                txtBarCode.SetFocus
                Exit Sub
            End If
            
            SS.Col = 1
            SS.Row = SS.DataRowCnt + 1
            
            SS.Text = txtBarCode.Text
            strPt = txtBarCode.Text
        
            Bjeobsudt = convLabnoToExpand(Mid(strPt, 1, 5))
            Bslipno1 = Mid(strPt, 6, 2)
            Bslipno2 = Mid(strPt, 8, 5)
            
            SS.Col = 2
            SS.Text = Bjeobsudt

            strSQL = ""
            strSQL = strSQL & " SELECT PTNO "
            strSQL = strSQL & "   FROM TWEXAM_GENERAL_SUB "                  ' 고객 MASTER
            strSQL = strSQL & "  WHERE JEOBSUDT = TO_DATE('" & Bjeobsudt & "','YYYY-MM-DD')"
            strSQL = strSQL & "    AND SLIPNO1 =   '" & Bslipno1 & "'"        ' 일련번호
            strSQL = strSQL & "    AND SLIPNO2 =   '" & Bslipno2 & "'"        ' 일련번호
            
            Result = AdoOpenSet(Rs, strSQL)
            
            If Result Then
                Rs.MoveFirst
                Do While Not Rs.EOF
                    Temp_ptno = Trim$(Rs.Fields("ptno")) & ""
                    Rs.MoveNext
                Loop
            Else
            
                SS.Text = ""
                SS.SetFocus                             'clear후 cell active 상태로 변경
                
                MsgBox " DATABASE에 등록된 내용이 없거나 접수가 잘못되었습니다." & vbCrLf & vbCrLf & _
                       " DATA를 재입력 하십시요." & vbCrLf & vbCrLf & _
                       " 재입력 후에도 같은 ERROR가 발생할 경우 전산실로 연락 바랍니다."
                Exit Sub
            End If
            
            AdoCloseSet Rs
            
            
            strSQL = ""
            strSQL = strSQL & " SELECT SNAME "
            strSQL = strSQL & "   FROM TWBAS_PATIENT "                  ' 고객 MASTER
            strSQL = strSQL & "  WHERE PTNO = '" & Temp_ptno & "' "     ' PATIENT NO
            
            If AdoOpenSet(Rs, strSQL) Then
                Rs.MoveFirst
                Do While Not Rs.EOF
                    Temp_name = Trim$(Rs.Fields("sname")) & ""
                    Rs.MoveNext
                Loop
            Else
                SS.Text = ""
                SS.SetFocus                             'clear후 cell active 상태로 변경
                SS.Action = SS_ACTION_ACTIVE_CELL
                MsgBox " DATABASE에 등록된 이름이 없거나 접수가 잘못되었습니다. " & vbCrLf & _
                       " PTNO를 재입력 하십시요. " & vbCrLf & vbCrLf & _
                       " 재입력 후에도 계속 같은 ERROR가 발생할 경우 전산실로 연락 바랍니다. "
                Exit Sub
            End If
            
            SS.Col = 3
            SS.Text = Temp_ptno
        
            SS.Col = 4
            SS.Text = Temp_name
            
            SS.Col = 1
            SS.SetFocus                             ' cell active 상태로 변경
            SS.Action = SS_ACTION_ACTIVE_CELL       ' 지정된 위치로 cursor 이동
            
            txtBarCode.SelStart = 0
            txtBarCode.SelLength = Len(txtBarCode.Text)
            txtBarCode.SetFocus
            
        End If
    Else
    
    End If
    
'    txtBarCode.SelStart = 0
'    txtBarCode.SelLength = Len(txtBarCode.Text)

End Sub


