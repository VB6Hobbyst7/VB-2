VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmaxsym 
   Caption         =   "Axsym"
   ClientHeight    =   8490
   ClientLeft      =   90
   ClientTop       =   2115
   ClientWidth     =   11880
   Icon            =   "frmaxsym.frx":0000
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
      TabIndex        =   6
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
      Height          =   5685
      Left            =   6360
      ScaleHeight     =   5625
      ScaleWidth      =   5115
      TabIndex        =   11
      Top             =   1440
      Visible         =   0   'False
      Width           =   5175
      Begin FPSpread.vaSpread SSR 
         Height          =   5490
         Left            =   30
         TabIndex        =   12
         Top             =   30
         Width           =   5055
         _Version        =   196608
         _ExtentX        =   8916
         _ExtentY        =   9684
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
         SpreadDesigner  =   "frmaxsym.frx":030A
      End
   End
   Begin FPSpread.vaSpread SS 
      Height          =   5700
      Left            =   330
      TabIndex        =   0
      Top             =   1440
      Width           =   11175
      _Version        =   196608
      _ExtentX        =   19711
      _ExtentY        =   10054
      _StockProps     =   64
      BackColorStyle  =   1
      ColsFrozen      =   5
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
      MaxCols         =   25
      SpreadDesigner  =   "frmaxsym.frx":0ED6
      UserResize      =   1
      VisibleCols     =   23
      VisibleRows     =   120
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Left            =   8310
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   2184
      Visible         =   0   'False
      Width           =   636
   End
   Begin VB.FileListBox File1 
      Height          =   270
      Left            =   8955
      TabIndex        =   4
      Top             =   2184
      Visible         =   0   'False
      Width           =   996
   End
   Begin VB.Timer Timer_Picture 
      Left            =   9312
      Top             =   1728
   End
   Begin VB.Timer Timer1 
      Left            =   8976
      Top             =   1728
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
      TabIndex        =   2
      Top             =   7248
      Width           =   6444
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9945
      Top             =   2190
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   10950
      Top             =   2220
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Timer Timer_30Sec 
      Left            =   9780
      Top             =   1710
   End
   Begin VB.Frame Frame2 
      Height          =   1308
      Left            =   288
      TabIndex        =   7
      Top             =   7152
      Width           =   4524
      Begin Threed.SSPanel SSPan 
         Height          =   492
         Left            =   144
         TabIndex        =   8
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
         TabIndex        =   10
         Top             =   720
         Width           =   2076
      End
      Begin VB.Label lblDate 
         BorderStyle     =   1  '단일 고정
         Height          =   492
         Left            =   144
         TabIndex        =   9
         Top             =   720
         Width           =   2076
      End
   End
   Begin VB.ListBox lstComm_R 
      Height          =   420
      Left            =   450
      TabIndex        =   14
      Top             =   6360
      Width           =   2715
   End
   Begin VB.ListBox lstComm_S 
      Height          =   420
      Left            =   3360
      TabIndex        =   15
      Top             =   6360
      Width           =   2715
   End
   Begin VB.Label lblPort 
      Alignment       =   2  '가운데 맞춤
      Height          =   225
      Left            =   4950
      TabIndex        =   13
      Top             =   8160
      Width           =   1965
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   5
      Left            =   6060
      Picture         =   "frmaxsym.frx":2EE3
      Top             =   810
      Visible         =   0   'False
      Width           =   480
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   10410
      Top             =   2190
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
            Picture         =   "frmaxsym.frx":31ED
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmaxsym.frx":3507
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmaxsym.frx":3821
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmaxsym.frx":3B3B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmaxsym.frx":3E55
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmaxsym.frx":416F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "date 변경용 obj"
      Height          =   225
      Left            =   6330
      TabIndex        =   3
      Top             =   2250
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
      Height          =   435
      Left            =   6750
      TabIndex        =   1
      Top             =   810
      Width           =   3945
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
Attribute VB_Name = "frmaxsym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
    'data Frame Set
    
    Dim GSendBuffT          As String
    Dim Query_Mode
    
    Dim RPoint                                  'row 위치 지정용
    Dim CPoint                                  'col 위치 지정용

    Dim Update_Check        As Boolean          ' 종료 check용
    Dim Update_Check_Force  As Boolean          ' 종료 check용
    Dim Receive_Check       As Boolean
    
    Dim strBiDirect_Trans   As Boolean          ' batch로 data 수신시 사용
    
    Dim Tcounter                                ' 수신표시 image count용
    Dim Qcounter            As Integer
    Dim Rcounter            As Integer
    
    Dim Ser                 As Integer
    Dim ResultText          As String
    Dim i
    Dim j
    
    Dim C1                  As Integer          ' work buffer column1
    Dim R1                  As Integer          ' work buffer row1
    
    Dim Temp_Jeobsu(100, 50)            As String           ' Spread의 Data를 Temp_Jeobsu Array로 Move
    Dim Temp_Request_Code(100, 50)      As String           ' input test code A1~Z6
    Dim Temp_Result(100, 50)            As String           ' Result Input Data
    
    Dim Temp_K(50, 2)       As String           ' item table data 입력용 buffer
    Dim FileOpen            As Boolean
    
    Dim MaxRecordCount      As Long
    
    Dim SColumn
    Dim SRow
    
    Dim temp_file           As String           'file directory
    
    Dim hSaveFile
    
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
    
    Dim timerx              As Boolean
    Dim PortOpen            As Boolean
    Dim SSCheck             As Boolean
    
    Dim RecordCount

    Dim SendBuffT           As String

    Dim StrGBER             As String       '응급구분 Check



Public Sub GotoSpreadSet()
    SS.SetFocus                             'clear후 cell active 상태로 변경
    SS.Row = 1
    SS.Col = 1
    SS.Action = SS_ACTION_GOTO_CELL
    SS.Action = SS_ACTION_ACTIVE_CELL
    
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    On Error Resume Next
    Select Case Button.Index
            Case 1
                   Call mnuReceive_Click
            Case 2
                   Call mnuEnd_Click
'            Case 3
'                   Call mnuWrite_Click
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
    
    Call CodeKy_Search_A                      ' codeky Read from twexam_itemml
    
    Label1.FontSize = 24
    Label1.BorderStyle = 0
    Label1.Caption = " Axsym "
    
    For i = 7 To SS.MaxCols                 ' spread header를 약어 code로 초기화
        SS.Row = 0
        SS.Col = i
        If Temp_K(i - 6, 1) <> "" Then
            SS.Text = Temp_K(i - 6, 1)
        Else
            SS.Text = "_"
        End If
    Next i
    
    Call Kdelete                     '30일 경과된 누적 file 삭제

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
        For j = 1 To 6
            SS.Row = i
            SS.Col = j
            SS.Lock = False
        Next j
    Next i
'    SS.Enabled = True
    
    Timer_Picture.Interval = 0                  'Timer_Picture_Timer End
    timerx = False
    Image1(5).Visible = False
    
    Call WorkDisplay(0)
    Close #hSaveFile
    FileOpen = False

End Sub


Private Sub SS_Click(ByVal Col As Long, ByVal Row As Long)
    Dim Rcount
    
    Rcount = 0
    
    If Col = 3 Or Col = 4 Then
        picResult.Visible = True
       i = 0
       
       SSR.MaxRows = 0
       
       SS.Col = 5
       SS.Row = Row
       
       SSR.MaxRows = Val(SS.Text)
       
        For i = 1 To MaxRecordCount
            If Temp_Jeobsu(Row, i + 10) <> "" Then
                
                Rcount = Rcount + 1
                
                SSR.Col = 1
                SSR.Row = Rcount  'i
                SSR.Text = "  " & Temp_K(i, 0)
                
                SSR.Col = 2
                SSR.Row = Rcount 'i
                SSR.Text = "  " & Temp_K(i, 1)
                
                SSR.Col = 3
                SSR.Row = Rcount 'i
                SSR.Text = Temp_Result(Row, i + 10)
            
            End If
        Next i
    Else
        picResult.Visible = False
    End If
    
End Sub


Private Sub SS_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    
    Dim Rs                  As ADODB.Recordset
    
    Dim strPt               As String
    Dim Bjeobsudt           As String
    Dim Bslipno1            As String
    Dim Bslipno2            As String
    
    Dim Checkdouble         As String
    Dim Ptnolen
    Dim Temp_ptno
    Dim Temp_name
    
    SS.Row = Row
    SS.Col = Col
    
    strPt = SS.Text
    
    Exit Sub

End Sub


Private Sub SS_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim SS_M_col            As Long             ' spread mouse down col val
    Dim SS_M_row            As Long             ' spread mouse down row val
    
    Dim Msg, Style, Title, Response
    
    Exit Sub
    
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
        strBiDirect_Trans = True    ' Query Mode
    Else
        strBiDirect_Trans = False   ' batch mode
    End If
    
    Query_Mode = 0
    
    Qcounter = 0
    Rcounter = 1

    ErrList.Clear
    
    timerx = True                                           '수신표시 정지용 flag
    FileOpen = True
    Timer_Picture.Interval = 500                            'Timer_Picture_Timer 1000mS
    
    SColumn = 6                                             'spread sheet 초기화 위치
    SRow = 1
    Call SS_INIT(SS, SColumn, SRow)

'******************** Part 3 **********************************
'*    Output용 File Name Set Open/Save 처리                   *
'**************************************************************
    On Error GoTo ErrorMsg
    Ser = Ser + 1
    CommonDialog1.InitDir = "C:\intdown"
    CommonDialog1.FileName = "X" & Format$(lblDate, "yyyymmdd") & Ser
    
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
    
    SSPan = "DATA 수신 대기중입니다."
    SSPan.ForeColor = &H0&                  ' black  &H00000000&

    For i = 1 To SS.MaxRows
        For j = 1 To SS.MaxCols
            SS.Row = i
            SS.Col = j
            SS.Lock = True
        Next j
    Next i

Exit Sub
     

ErrCancel:
    MsgBox "자료수신을 취소하였습니다."
    Call ErrCancel_Click                    ' 자료수신 종료
Exit Sub


ErrorMsg:
    MsgBox "Error " & "Code = " & Err.Number & vbLf & vbLf & Err.Description
    If FileOpen = True Then
        Close #hSaveFile                    ' error 발생시 file close
    End If

    For i = 1 To SS.MaxRows
        For j = 1 To SS.MaxCols
            SS.Row = i
            SS.Col = j
            SS.Lock = False
        Next j
    Next i

    Timer_Picture.Interval = 0              'Timer_Picture_Timer End
    timerx = False                          '수신표시 정지용 flag
    Image1(5).Visible = False               '수신표시 image
    Call WorkDisplay(0)
    If MSComm1.PortOpen = True Then
        MSComm1.PortOpen = False
    End If

End Sub



'{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}
'{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}
'{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}
'{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}

Private Sub MSComm1_OnComm()                            ' DATA 수신 처리
    Dim EVMsg$
    Dim ERMsg$
    Select Case MSComm1.CommEvent                       'CommEvent 속성에 따른 항목
        Case comEvReceive                               ' 포트로부터 데이터가 들어왔음...
             RBuffer = MSComm1.Input
        Case Else
             Exit Sub
    End Select
    
    Select Case RBuffer
           Case STX:    STXBuffer = True                ' STX [] Check용
           Case ETX:    ETXBuffer = True                ' ETX [] Check용
           Case EOT:    EOTBuffer = True                ' EOT [] Check용
           Case ENQ:    ENQBuffer = True                ' ENQ     Check용
           Case ACK:    ACKBuffer = True                ' ACK     Check용
           Case NACK:   NACKBuffer = True               ' NACK    Check용
    End Select
    
    RBufferSum = RBufferSum & RBuffer                   ' comm port 에서 입력한 data 누적
    
    If STXBuffer = True And RBuffer = vbLf Then         'STX & LF  Check
        Call RPrint(RBufferSum)                         'write omitting cr lf
        If Mid(RBufferSum, 3, 1) = "Q" Then             '장비로부터 QUERY DATA 수신
            Query_Mode = 1
        ElseIf Mid(RBufferSum, 3, 1) = "P" Then         '장비로부터 PATIENT RESULT DATA 수신
            Query_Mode = 0
        End If
        lstComm_R.AddItem RBufferSum
        Call Ack_Send
    End If
    
    If EOTBuffer = True Then                            'DATA 수신 종료 Flag
        If Query_Mode = 1 Then
            Print #hSaveFile, ""
            Call RPrint(RBufferSum)                     'write omitting cr lf
            Call DataReceive_Query
'            Timer_30Sec.Interval = 1000                 'max 1sec (10000)
            Timer_30Sec.Interval = 100                 'max 1sec (10000)
        Else
            Print #hSaveFile, ""
            Call RPrint(RBufferSum)                     'write omitting cr lf
            Call DataReceive_Result
        End If
        EOTBuffer = False
        RBufferSum = ""
    End If
    
    If ACKBuffer = True Then
        GSendBuffT = ""
        Call RPrint(RBufferSum)
        Call Order_Data_Send                            'DATA 송신 시작 Flag
        ACKBuffer = False
    
        ENQBuffer = False
        RBufferSum = ""
    End If
    
    If ENQBuffer = True Then
        Print #hSaveFile, ""
        Print #hSaveFile, "   Receive data"
        Call RPrint(RBufferSum)                         ' write omitting cr lf
        Call Ack_Send                                   'DATA 송신 시작 Flag from STA
        ENQBuffer = False
        RBufferSum = ""
    End If
    
    If NACKBuffer = True Then
        Print #hSaveFile, "NACK Received"
        Call Order_Data_Send                            'DATA 송신 시작 Flag
        
        NACKBuffer = False
        RBufferSum = ""
    End If
    
End Sub

Private Sub RPrint(RData)
    If FileOpen = True Then
        If Len(RData) >= 3 Then
            Print #hSaveFile, "Rx " & Format(lblTime, "hh:mm:ss") & " ]  " & Mid(RData, 1, Len(RData) - 2)
        Else
            Print #hSaveFile, "Rx " & Format(lblTime, "hh:mm:ss") & " ]  " & RData
        End If
    End If

End Sub


Private Sub DataReceive_Query()
    On Error Resume Next
    
    Dim i, line_cnt             As Integer
    Dim cc, SID, ANO, OrderSeq  As String
'    Dim R1, R2, R3              As String
    
    line_cnt = lstComm_R.ListCount
    OrderSeq = 0
    
    For i = 0 To line_cnt - 1
       cc = Mid(lstComm_R.List(i), 3, 1)
       Select Case cc
         Case "Q"
                   SID = QGetSID(lstComm_R.List(i))
         Case "L"
                   If SID = "" Then Exit Sub
                   Call Save_Query(SID)             'data display & DB Write

       End Select
       
    Next i

End Sub


Private Sub Save_Query(aaa) '(QSID, QANO)
    
    Dim QSID            As String
    Dim QANO            As String
    Dim QOrder          As String
    
    O_NO = 0
    P_NO = 0
    MOD_8 = 0
    QSID = aaa
    
    lstComm_S.Clear
    
    
    
    For i = 1 To SS.DataRowCnt
        
        SS.Col = 1
        SS.Row = i
        
        If Trim(QSID) = Trim(SS.Text) Then          ' Axsym에서 중복 Run을 check 하여 중복 수신된 data를 clear 하여 무시한다.
    
            SSPan = " 중복 실행으로 Order 무시함 " & QSID & " !!!!!"
            ErrList.AddItem "  중복실행 " & QSID
            
            lstComm_R.Clear
            Timer_30Sec.Interval = 0
            Exit Sub
        End If
    Next i
    
    
    Qcounter = Qcounter + 1
    Call ini_check(QSID)
    
    lstComm_S.AddItem MakeH                                         'Header
    lstComm_S.AddItem MakeP(QSID)                                         'Patient
    
    'order select
    
    For i = 11 To 50                                                ''temp_k의 항목수
        
        If Trim(Temp_Jeobsu(Qcounter, i)) <> "" Then
            QOrder = Trim(Temp_K(i - 10, 2))
            If QOrder <> "" Then
                lstComm_S.AddItem MakeO(QSID, QOrder)      'Order
            End If

        End If
    Next i
    
    lstComm_S.AddItem MakeL                                         'End
    
    lstComm_S.AddItem EOT                                           'EOT
    
    lstComm_R.Clear
    
    Timer_30Sec.Interval = 0
    

End Sub


Private Sub Timer_30Sec_Timer()
    Dim SendBuff            As String
    
    SSPan.ForeColor = &HFF&                                 ' red  &H000000FF&
    SSPan = "Patient Order Data 송신중입니다..........."
    
    SendBuff = ENQ
    If MSComm1.PortOpen = True Then
        MSComm1.Output = SendBuff
        Print #hSaveFile, ""
        Print #hSaveFile, "   Order Send"
        Print #hSaveFile, "Tx " & Format(lblTime, "hh:mm:ss") & " ]  " & SendBuff
    End If
    ENQBuffer = False

    Timer_30Sec.Interval = 0

End Sub


Private Sub Ack_Send()
    Dim SendBuff            As String
    
    SendBuff = ACK
    
    If MSComm1.PortOpen = True Then
        MSComm1.Output = SendBuff
        Print #hSaveFile, "Tx " & Format(lblTime, "hh:mm:ss") & " ]  " & SendBuff
    End If
    
    STXBuffer = False
    ENQBuffer = False
    RBufferSum = ""
    RBuffer = ""

End Sub


Private Sub Order_Data_Send()
    On Error Resume Next

    If SS.DataRowCnt < 1 Then
        Exit Sub
    End If
    
    If lstComm_S.ListCount = 0 Then
        Exit Sub
    End If
        
    SSPan.ForeColor = &HFF&                         ' red  &H000000FF&
    SSPan = "ORDER DATA 전송중입니다..........."
    
    If GSendBuffT = "" Then
        SendBuffT = lstComm_S.List(0)
        If MSComm1.PortOpen = True Then
            GSendBuffT = SendBuffT
            MSComm1.Output = SendBuffT               'data send to com port
            Print #hSaveFile, TSaveRecord(SendBuffT)
        End If
        lstComm_S.RemoveItem (0)
    Else
        If MSComm1.PortOpen = True Then
            MSComm1.Output = GSendBuffT               '재송신
            Print #hSaveFile, TSaveRecord(GSendBuffT)
        End If
        
        DELAY (1)
    
    End If
    
    SSPan.ForeColor = &H0&                       ' black  &H000000FF&
    SSPan = "ORDER DATA 수신대기중입니다..........."
    
End Sub


Private Sub ini_check(QRSID)
    Dim Rs                  As ADODB.Recordset
    
    Dim Bjeobsudt
    Dim Bslipno1
    Dim Bslipno2
    
    Dim Verify_Check

    Dim Order_Count

'******************** Part 1 **********************************
'*      Spread의 pt no와 slipno2를 Temp_Jeobsu Array로 Move   *
'**************************************************************
    Order_Count = 0
    
    R1 = Qcounter
        C1 = 1
        
        SS.Row = R1
        SS.Col = C1
        SS.Text = QRSID
        
        Temp_Jeobsu(R1, C1) = QRSID
        
                       'Rn  Cn
        Bjeobsudt = convLabnoToExpand(Mid(Temp_Jeobsu(R1, C1), 1, 5))
        Bslipno1 = Mid(Temp_Jeobsu(R1, C1), 6, 2)
        Bslipno2 = Mid(Temp_Jeobsu(R1, C1), 8, 5)
        
        SS.Row = R1
        
        'date
        Temp_Jeobsu(R1, 2) = Bjeobsudt
        SS.Col = 2
        SS.Text = Temp_Jeobsu(R1, 2)
        
        'PTNO
        Temp_Jeobsu(R1, 3) = PTNOSearch(Temp_Jeobsu(R1, 1))
        SS.Col = 3
        SS.Text = Temp_Jeobsu(R1, 3)
        
        'Name
        Temp_Jeobsu(R1, 4) = NameSearch(Temp_Jeobsu(R1, 3))
        SS.Col = 4
        SS.Text = Temp_Jeobsu(R1, 4)
        
        strSQL = ""
        strSQL = strSQL & " SELECT ITEMCD,GEOMJAN3,GBER "
        strSQL = strSQL & "   FROM TWEXAM_GENERAL_SUB A, "                   ' 검사접수결과 세부사항
        strSQL = strSQL & "        TWEXAM_ITEMML B, "                        ' 검사 ITEM MASTER
        strSQL = strSQL & "        TWEXAM_GENERAL C "                        ' 검사접수결과
        strSQL = strSQL & "  WHERE A.JEOBSUDT = TO_DATE('" & Bjeobsudt & "','YYYY-MM-DD')"
        strSQL = strSQL & "    AND A.SLIPNO1   = '" & Bslipno1 & "'"         ' 일련번호
        strSQL = strSQL & "    AND A.SLIPNO2   = '" & Bslipno2 & "'"         ' 일련번호
        strSQL = strSQL & "    AND A.ITEMCD    = B.CODEKY "
        strSQL = strSQL & "    AND B.GBROUTINE = 'I'   "
'        strSQL = strSQL & "    AND B.GEOMJAN3  IS NOT NULL "
        strSQL = strSQL & "    AND A.PTNO      = C.PTNO "
        strSQL = strSQL & "    AND A.JEOBSUDT  = C.JEOBSUDT "
        strSQL = strSQL & "    AND A.SLIPNO1   = C.SLIPNO1 "
        strSQL = strSQL & "    AND A.SLIPNO2   = C.SLIPNO2 "
        
        Result = AdoOpenSet(Rs, strSQL)
        
        'Debug.Print Rowindicator
        
        If Result Then
            Do While Not Rs.EOF
                For i = 1 To 50    'temp_k의 항목수
                
                    If Temp_K(i, 0) = Trim(Rs.Fields("ITEMCD") & "") Then
                        
                        Order_Count = Order_Count + 1
                        
                        Temp_Jeobsu(R1, i + 10) = Trim(Rs.Fields("ITEMCD") & "")
                        If Trim(Rs.Fields("GBER") & "") = "E" Then
                            StrGBER = "S"
                        Else
                            StrGBER = "R"
                        End If
                    End If
                Next i
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
    Temp_Jeobsu(R1, 5) = Order_Count ' Rowindicator       'order 갯수
    SS.Text = Order_Count               ' Rowindicator

    SS.SetFocus                             ' cell active 상태로 변경
    SS.Action = SS_ACTION_ACTIVE_CELL       ' 지정된 위치로 cursor 이동

End Sub


Private Sub DataReceive_Result()

    On Error Resume Next

    Dim i, line_cnt             As Integer
    Dim cc, SID, ANO, OrderSeq  As String
    Dim R1, R2, R3              As String
    
    line_cnt = lstComm_R.ListCount
    OrderSeq = 0
    SID = ""
    ANO = ""
    
    For i = 0 To line_cnt - 1
       cc = Mid(lstComm_R.List(i), 3, 1)
       Select Case cc
         Case "P"
                   'DO NOT USE
'                   SID = GetSID(lstComm_R.List(i))
'                   ANO = GetANO(lstComm_R.List(i))
         Case "O"
                   If SID <> "" Then
                       'ErrList.AddItem SID & "  /" & ANO & "  /" & R1 & "  /" & R2 & "  /" & R3
                       Print #hSaveFile, SID & "  /" & ANO & "  /" & R1 & "  /" & R2 & "  /" & R3
                       If Len(Trim(SID)) = 12 Then Call Save_Result(SID, ANO, R1, R2, R3)
                       R1 = ""
                       R2 = ""
                       R3 = ""
                   End If
                   
                   SID = GetSID(lstComm_R.List(i))
                   ANO = GetANO(lstComm_R.List(i))
         
         Case "R"
                   OrderSeq = Mid(lstComm_R.List(i), 5, 1)
                   Select Case OrderSeq
                     Case "1": R1 = GetResult(lstComm_R.List(i))
                     Case "2": R2 = GetResult(lstComm_R.List(i))
                     Case "3": R3 = GetResult(lstComm_R.List(i))
                   End Select
                   
         Case "L"
                   'ErrList.AddItem SID & "  /" & ANO & "  /" & R1 & "  /" & R2 & "  /" & R3
                   Print #hSaveFile, SID & "  /" & ANO & "  /" & R1 & "  /" & R2 & "  /" & R3
                   If Len(Trim(SID)) = 12 Then Call Save_Result(SID, ANO, R1, R2, R3)
                   
                   SID = ""
                   ANO = ""
                   R1 = ""
                   R2 = ""
                   R3 = ""
                   
       End Select
       
    Next i

    lstComm_R.Clear

End Sub


'Private Sub Save_Result(RSID, jcode, ResultU1)
Private Sub Save_Result(RSID, Jcode, ResultU1, ResultU2, ResultU3)
    
    Dim Bdt
    Dim Bno1
    Dim Bno2
    Dim ResultU
    Dim itemcd1
    
    
    Bdt = convLabnoToExpand(Mid(RSID, 1, 5))
    Bno1 = Mid(RSID, 6, 2)
    Bno2 = Mid(RSID, 8, 5)
    
    For i = 1 To SS.DataRowCnt
        SS.Row = i
        SS.Col = 1
        If Trim(SS.Text) = RSID Then
            RPoint = i
            Exit For
        End If
    Next i
    
    
    For i = 1 To 50
        If Temp_K(i, 2) = Jcode Then
            itemcd1 = Temp_K(i, 0)
            CPoint = i + 10
            Exit For
        End If
    Next i
    
    If Trim(ResultU3) = "" Then
        ResultU = ResultU1
    Else
        If ResultU3 = "NONREACTIVE" Then
            ResultU3 = "NEGATIVE"
        ElseIf ResultU3 = "REACTIVE" Then
            ResultU3 = "POSITIVE"
        End If
        ResultU = ResultU3
    End If
    
    
    adoConnect.BeginTrans                          ' TRANSACTION의 종료시에 COMMITTRANS를 지정함
    
    strSQL = ""
    strSQL = strSQL & "UPDATE TWEXAM_GENERAL_SUB "
    strSQL = strSQL & "   SET RESULT1  =   '" & ResultU & "'"
    strSQL = strSQL & " WHERE JEOBSUDT =  TO_DATE('" & Bdt & "','YYYY-MM-DD') "    '입력된 날자로 검색
    strSQL = strSQL & "   AND SLIPNO1  =   '" & Bno1 & "' "                        '구분
    strSQL = strSQL & "   AND SLIPNO2  =   '" & Bno2 & "' "                        '구분
    strSQL = strSQL & "   AND ITEMCD   =    '" & itemcd1 & "'"                     'ITEMCODE
    strSQL = strSQL & "   AND VERIFY   =  'N'"                                    ' 접수결과에서 VERIFY OK한경우에는 UPDATE하지않음
    
    Result = AdoExecute(strSQL)
    
    If Result = True And Rowindicator > 0 Then
        SSPan = "DB(SUB)에 저장 되었습니다. "
        
        adoConnect.CommitTrans                                                   ' TRANSACTION 종료시에 COMMIT 시킴
    Else
        ErrList.AddItem "    Verify Data       " & RSID
        ErrList.AddItem "    or Update Error   " & itemcd1 & "  " & ResultU
        ErrList.ListIndex = ErrList.ListCount - 1
        
        'file write routine insert
        adoConnect.RollbackTrans                                                 ' TRANSACTION ERROR시 ROLLBACK 시킴
        SSPan = "DB(SUB)에 갱신중 ERROR가 발생하였습니다." & vbCrLf & _
                "VERIFY된 DATA인지 확인하십시요."
    End If
    
    
    If strBiDirect_Trans = True Then
        SS.Row = Val(RPoint)
        SS.Col = Val(CPoint) - 4
'        SS.Text = Temp_Result(RPoint, CPoint)
        SS.Text = ResultU
        
        
'        SSR.Text = Temp_Result(Row, i + 10)
'        Temp_Result(RPoint, 10 + CPoint) = ResultU
        Temp_Result(RPoint, CPoint) = ResultU
        
        SS.Row = Val(RPoint)
        SS.Col = 6
        SS.Text = Val(SS.Text) + 1
        
        Dim Comp_Check
        Comp_Check = SS.Text
        
        SS.Col = 5
        
        If Comp_Check = SS.Text Then
            SS.Col = 6
            SS.BackColor = RGB(0, 255, 0)
            
            'general_sub update
            Call Save_Result_Flag(RSID)
        
        ElseIf Comp_Check > SS.Text Then
            SS.Col = 6
            SS.BackColor = RGB(255, 255, 0)
        End If
    Else
        
        SS.Row = Val(RPoint)
        
        SS.Col = 1
        SS.Text = RSID
        
        SS.Col = Val(CPoint) - 4
        SS.Text = ResultU
    
        SS.Col = 6
        SS.Text = Val(SS.Text) + 1
        
    End If

End Sub


Private Sub Save_Result_Flag(JeobsuPT2)

    Dim Bdt
    Dim Bno1
    Dim Bno2
    
    Bdt = convLabnoToExpand(Mid(JeobsuPT2, 1, 5))
    Bno1 = Mid(JeobsuPT2, 6, 2)
    Bno2 = Mid(JeobsuPT2, 8, 5)
    
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
        SSPan = "DB(GENERAL)에 저장 되었습니다. "
        adoConnect.CommitTrans                                                   ' TRANSACTION 종료시에 COMMIT 시킴
    Else
        ErrList.AddItem "    GENERAL STATUS    " & JeobsuPT2
        ErrList.AddItem "    접수취소 or VERIFY"
        ErrList.ListIndex = ErrList.ListCount - 1
        
        adoConnect.RollbackTrans                                                 ' TRANSACTION ERROR시 ROLLBACK 시킴
        SSPan = " DB(GENERAL)에 갱신중 ERROR가 발생하였습니다." & vbCrLf & _
                " 결과완료된 DATA인지 확인하십시요."
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
    Title = "검사를 종료 확인"                          ' 기본 제목.
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
    
    Timer_Picture.Interval = 0                       'Timer_Picture_Timer End
    timerx = False
    Timer_30Sec.Interval = 0
    
    Image1(5).Visible = False
    
    Call WorkDisplay(0)

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

    Print #hSaveFile, " "
    Print #hSaveFile, " " & "  Data Communication End"
    Print #hSaveFile, " "
    
    
    If FileOpen = True Then
        Close #hSaveFile
    End If
    FileOpen = False
    If MSComm1.PortOpen = True Then
        MSComm1.PortOpen = False
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
            Call SS_INIT(SS, 5, 1)
            Call GotoSpreadSet
            
            lstComm_R.Clear
            lstComm_S.Clear
            ErrList.Clear
            
            Update_Check = False
            
            SSPan = ""
            StrGBER = "R"
            RPoint = 0
            CPoint = 0
            
            For i = 0 To 100
                For j = 0 To 30
                    Temp_Jeobsu(i, j) = ""
                    Temp_Request_Code(i, j) = ""
                    Temp_Result(i, j) = ""
                Next j
            Next i
        
        End If
    Else
        Msg = " 화면을 CLEAR하시겠습니까?"
        Style = vbYesNo + vbQuestion + vbDefaultButton2     ' Define buttons.
        Title = "화면 CLEAR"                                ' 기본 제목.
        Response = MsgBox(Msg, Style, Title)
        If Response = vbYes Then                            ' 사용자가 예를 선택.
            Call SS_INIT(SS, 1, 1)
            Call GotoSpreadSet
            
            lstComm_R.Clear
            lstComm_S.Clear
            ErrList.Clear
            
            SSPan = ""
            StrGBER = "R"
            RPoint = 0
            CPoint = 0
            For i = 0 To 100
                For j = 0 To 30
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
    
    GGJCODE = "7"
    
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
    
    SSPan.FontSize = 10
    
    Update_Check = False
    Update_Check_Force = False
    
    '통신용 Definition Character
    SOH = Chr(1)                '<SOH>
    STX = Chr(2)                '<STX>
    ETX = Chr(3)                '<ETX>
    EOT = Chr(4)                '<EOT>
    ENQ = Chr(5)                '<ENQ>
    ACK = Chr(6)                '<ACK>
    NACK = Chr(21)              '<NACK>
    ETB = Chr(23)               '<ETB>
    
    H_FRAME = "H|\^&||||||||||P|1"
    P_FRAME = "P|||||"
    
    O_FRAME = "O||||^^^|||||||N||||||||||||||Q"
    
    L_FRAME = "L|1"
    O_NO = 0
    P_NO = 0
    MOD_8 = 0
    
    
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
           
'           Frame1.Enabled = False
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
                Image1(5).Visible = True
        Case 2
                Image1(5).Visible = False
                Tcounter = 0
    End Select

End Sub


Sub CodeKy_Search_A()
    
    'Axsym Code Setting Axsym에 검사항목이 변경될 경우 수정
     
    Temp_K(1, 0) = "220112"         'CODEKY
    Temp_K(1, 1) = "PSA_Total"      'YAGEO
    Temp_K(1, 2) = "443"            'GEOMJAN3 (test code)
    
    Temp_K(2, 0) = "220207"
    Temp_K(2, 1) = "PSA_Free"
    Temp_K(2, 2) = "442"
    
    Temp_K(3, 0) = "310131"
    Temp_K(3, 1) = "HBsAg"
    Temp_K(3, 2) = "106"
    
    Temp_K(4, 0) = "310132"
    Temp_K(4, 1) = "Anti-HBs"       '"AUSAB"
    Temp_K(4, 2) = "118"
    
    Temp_K(5, 0) = "310133"
    Temp_K(5, 1) = "HBe"
    Temp_K(5, 2) = "193"
    
    Temp_K(6, 0) = "310134"
    Temp_K(6, 1) = "Anti-HBe"       'version 1.0
    Temp_K(6, 2) = "181"            '192  => 181  2000.04.26
    
'    Temp_K(6, 0) = "310134"
'    Temp_K(6, 1) = "Anti-HBe 2"     ' version 2.0
'    Temp_K(6, 2) = "192"            '192  => 181  2000.04.26
    
    Temp_K(7, 0) = "320135"
    Temp_K(7, 1) = "RubellaG"
    Temp_K(7, 2) = "723"
    
    Temp_K(8, 0) = "320136"
    Temp_K(8, 1) = "RubellaM"
    Temp_K(8, 2) = "754"
    
    Temp_K(9, 0) = "220926"
    Temp_K(9, 1) = "B2-M"
    Temp_K(9, 2) = "474"
    
    Temp_K(10, 0) = "220820"
    Temp_K(10, 1) = "Insulin"
    Temp_K(10, 2) = "370"
    
    Temp_K(11, 0) = "220904"
    Temp_K(11, 1) = "B12"
    Temp_K(11, 2) = "306"
    
    Temp_K(12, 0) = "220905"
    Temp_K(12, 1) = "Folate"
    Temp_K(12, 2) = "340"
    
    Temp_K(13, 0) = "221014"
    Temp_K(13, 1) = "Carb"
    Temp_K(13, 2) = "656"
    
    Temp_K(14, 0) = "221030"
    Temp_K(14, 1) = "Digoxin"
    Temp_K(14, 2) = "601"
    
    Temp_K(15, 0) = "221032"
    Temp_K(15, 1) = "Theo_II"
    Temp_K(15, 2) = "613"
    
    Temp_K(16, 0) = "221017"
    Temp_K(16, 1) = "Pheny"
    Temp_K(16, 2) = "623"
    
    Temp_K(17, 0) = "221019"
    Temp_K(17, 1) = "Valp"
    Temp_K(17, 2) = "689"
    
    Temp_K(18, 0) = "221023"
    Temp_K(18, 1) = "Ethanol"
    Temp_K(18, 2) = "544"
    
    MaxRecordCount = 18                        ' 검사항목 갯수
    SS.MaxCols = 6 + 18                        'Record Count를 check 하여 max columns를 결정한다.
    
End Sub


