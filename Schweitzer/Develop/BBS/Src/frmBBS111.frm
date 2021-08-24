VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frmBBS111 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E8EEEE&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "Accept List"
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6645
   Icon            =   "frmBBS111.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows 기본값
   Begin VB.Frame fraS 
      BackColor       =   &H00E8EEEE&
      BorderStyle     =   0  '없음
      Caption         =   "Frame1"
      Height          =   840
      Index           =   0
      Left            =   720
      TabIndex        =   5
      Top             =   135
      Width           =   3210
      Begin VB.OptionButton optAc 
         BackColor       =   &H00E8EEEE&
         Caption         =   "Acting일"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1170
         TabIndex        =   23
         Top             =   555
         Value           =   -1  'True
         Width           =   1290
      End
      Begin VB.OptionButton optAc 
         BackColor       =   &H00E8EEEE&
         Caption         =   "처방일"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   150
         TabIndex        =   27
         Top             =   555
         Width           =   1185
      End
      Begin VB.TextBox txtDay 
         Alignment       =   2  '가운데 맞춤
         Height          =   285
         Left            =   2130
         TabIndex        =   6
         Top             =   150
         Width           =   390
      End
      Begin VB.Label lblSubMenu 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         Caption         =   "접수번호리스트("
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   480
         Left            =   0
         TabIndex        =   8
         Top             =   135
         Width           =   2355
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "일전)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2550
         TabIndex        =   7
         Top             =   165
         Width           =   750
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H005E957E&
         BorderWidth     =   3
         FillColor       =   &H00A7C9B9&
         FillStyle       =   0  '단색
         Height          =   420
         Index           =   1
         Left            =   60
         Shape           =   4  '둥근 사각형
         Top             =   90
         Width           =   3105
      End
   End
   Begin VB.PictureBox pichook 
      Height          =   555
      Left            =   780
      ScaleHeight     =   495
      ScaleWidth      =   795
      TabIndex        =   24
      Top             =   8460
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSComctlLib.ImageList ImgList 
      Left            =   1785
      Top             =   8310
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBBS111.frx":144A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBBS111.frx":176E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBBS111.frx":1A8A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.OptionButton OptS 
      BackColor       =   &H00E8EEEE&
      Caption         =   "Prep"
      Height          =   225
      Index           =   2
      Left            =   30
      TabIndex        =   14
      Top             =   735
      Width           =   690
   End
   Begin VB.CommandButton cmdRefresh 
      BackColor       =   &H00C8CEDF&
      Caption         =   "&Refresh"
      Height          =   450
      Left            =   4470
      Style           =   1  '그래픽
      TabIndex        =   0
      Top             =   165
      Width           =   1125
   End
   Begin VB.CommandButton cmdUpDown 
      BackColor       =   &H00E7BAB4&
      Caption         =   "▲"
      Height          =   435
      Left            =   4455
      Style           =   1  '그래픽
      TabIndex        =   13
      Tag             =   "0"
      Top             =   615
      Width           =   1140
   End
   Begin VB.Timer Timer1 
      Left            =   1350
      Top             =   8325
   End
   Begin VB.OptionButton OptS 
      BackColor       =   &H00E8EEEE&
      Caption         =   "접수"
      Height          =   225
      Index           =   0
      Left            =   30
      TabIndex        =   4
      Top             =   180
      Value           =   -1  'True
      Width           =   660
   End
   Begin VB.OptionButton OptS 
      BackColor       =   &H00E8EEEE&
      Caption         =   "채혈"
      Height          =   225
      Index           =   1
      Left            =   30
      TabIndex        =   3
      Top             =   450
      Width           =   660
   End
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H00C8CEDF&
      Caption         =   "닫기(&X)"
      Height          =   495
      Left            =   5055
      Style           =   1  '그래픽
      TabIndex        =   2
      Top             =   8310
      Width           =   1455
   End
   Begin FPSpread.vaSpread tblPtList 
      Height          =   7005
      Left            =   30
      TabIndex        =   1
      Top             =   1080
      Width           =   6495
      _Version        =   196608
      _ExtentX        =   11456
      _ExtentY        =   12356
      _StockProps     =   64
      BackColorStyle  =   1
      ColHeaderDisplay=   0
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   15265518
      GridColor       =   16703181
      GridShowVert    =   0   'False
      MaxCols         =   14
      MaxRows         =   10
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      ShadowText      =   0
      SpreadDesigner  =   "frmBBS111.frx":1DAE
      TextTip         =   2
   End
   Begin VB.Frame fraS 
      BackColor       =   &H00E8EEEE&
      BorderStyle     =   0  '없음
      Caption         =   "Frame1"
      Height          =   840
      Index           =   2
      Left            =   735
      TabIndex        =   15
      Top             =   135
      Width           =   3210
      Begin VB.TextBox Text1 
         Alignment       =   2  '가운데 맞춤
         Height          =   285
         Left            =   2130
         TabIndex        =   16
         Top             =   105
         Width           =   390
      End
      Begin MedControls1.LisLabel LisLabel8 
         Height          =   195
         Index           =   1
         Left            =   1425
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   540
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   344
         BackColor       =   255
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel8 
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   540
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   344
         BackColor       =   16711680
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "실처방항목"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   390
         TabIndex        =   22
         Tag             =   "103"
         Top             =   540
         Width           =   975
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Prep 취소항목"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   1665
         TabIndex        =   21
         Tag             =   "103"
         Top             =   540
         Width           =   1230
      End
      Begin VB.Label Label4 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         Caption         =   "혈액Prep리스트("
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   480
         Left            =   0
         TabIndex        =   17
         Top             =   90
         Width           =   2355
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "일전)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2550
         TabIndex        =   18
         Top             =   120
         Width           =   750
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H005E957E&
         BorderWidth     =   3
         FillColor       =   &H00A7C9B9&
         FillStyle       =   0  '단색
         Height          =   420
         Index           =   2
         Left            =   60
         Shape           =   4  '둥근 사각형
         Top             =   45
         Width           =   3105
      End
   End
   Begin VB.Frame fraS 
      BackColor       =   &H00E8EEEE&
      BorderStyle     =   0  '없음
      Height          =   840
      Index           =   1
      Left            =   735
      TabIndex        =   9
      Top             =   135
      Width           =   3210
      Begin VB.OptionButton optAct 
         BackColor       =   &H00E8EEEE&
         Caption         =   "처방일"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   0
         Left            =   15
         TabIndex        =   26
         Top             =   570
         Width           =   1185
      End
      Begin VB.OptionButton optAct 
         BackColor       =   &H00E8EEEE&
         Caption         =   "Acting 일"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   1
         Left            =   1290
         TabIndex        =   25
         Top             =   570
         Value           =   -1  'True
         Width           =   1320
      End
      Begin VB.TextBox txtCollect 
         Alignment       =   2  '가운데 맞춤
         Height          =   285
         Left            =   2055
         TabIndex        =   11
         Top             =   135
         Width           =   390
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "일전)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2475
         TabIndex        =   12
         Top             =   150
         Width           =   750
      End
      Begin VB.Label Label2 
         Alignment       =   2  '가운데 맞춤
         BackStyle       =   0  '투명
         Caption         =   "채혈 항목리스트("
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   480
         Left            =   -120
         TabIndex        =   10
         Top             =   165
         Width           =   2355
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  '투명하지 않음
         BorderColor     =   &H005E957E&
         BorderWidth     =   3
         FillColor       =   &H00A7C9B9&
         FillStyle       =   0  '단색
         Height          =   420
         Index           =   0
         Left            =   30
         Shape           =   4  '둥근 사각형
         Top             =   75
         Width           =   3105
      End
   End
   Begin VB.Image imgSound 
      Height          =   480
      Index           =   1
      Left            =   195
      Picture         =   "frmBBS111.frx":248F
      Top             =   8505
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmBBS111"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'Option Explicit
'
'Public Event LastFormUnload()
'Public Event ThisFormUnload()
'Public Event ListSelected(ByVal SelPtId As String, ByVal SelFrDt As String, ByVal SelToDt As String)
'
'Private Type NOTIFYICONDATA
'    cbSize As Long
'    hwnd As Long
'    uId As Long
'    uFlags As Long
'    ucallbackMessage As Long
'    hIcon As Long
'    szTip As String * 64
'End Type
'
'Private Const NIM_ADD = &H0
'Private Const NIM_MODIFY = &H1
'Private Const NIM_DELETE = &H2
'Private Const NIF_MESSAGE = &H1
'Private Const NIF_ICON = &H2
'Private Const NIF_TIP = &H4
'
'Private Const WM_LBUTTONDBLCLK = &H203
'Private Const WM_LBUTTONDOWN = &H201
'Private Const WM_LBUTTONUP = &H202
'Private Const WM_MBUTTONDBLCLK = &H209
'Private Const WM_MBUTTONDOWN = &H207
'Private Const WM_MBUTTONUP = &H208
'Private Const WM_RBUTTONDBLCLK = &H206
'Private Const WM_RBUTTONDOWN = &H204
'Private Const WM_RBUTTONUP = &H205
'
'Private blnStopFg As Boolean
'Private blnSound As Boolean
'Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
'Private TrayI As NOTIFYICONDATA
'Private blnForce As Boolean
'
'
'Private Sub cmdExit_Click()
'    Unload Me
'    Set frmBBS111 = Nothing
'End Sub
'
'Private Sub cmdRefresh_Click()
'    Me.MousePointer = 11
'
'    blnStopFg = True
''    mmSound.Command = "Stop"
''    mmSound.Command = "Close"
'
'    DoEvents
'    If OptS(0).value = True Then
'        Call Query
'    ElseIf OptS(1).value = True Then
'        Call CollectQuery
'        Timer1.Enabled = True
'    Else
'        Call PrepQuery
'    End If
'    Me.MousePointer = 0
'End Sub
'
'Private Sub cmdUpDown_Click()
'    If cmdUpDown.tag = "0" Then
'        cmdUpDown.tag = "1"
'        cmdUpDown.Caption = "▼"
'        Me.Height = 1450
'        blnForce = True
'    Else
'        cmdUpDown.tag = "0"
'        cmdUpDown.Caption = "▲"
'        Me.Height = 9300
'        blnForce = False
'    End If
'End Sub
'
'Private Sub Form_Load()
'
'    blnStopFg = False
'    blnSound = True
'
'    fraS(0).Visible = True
'    fraS(1).Visible = False
'    fraS(2).Visible = False
'    OptS(0).value = True
'
'    mmSound.Notify = False
'    mmSound.Wait = True
'    mmSound.Shareable = False
'    mmSound.DeviceType = "WaveAudio"
'    mmSound.FileName = gBloodRequestMusic
'    mmSound.Enabled = True
'    mmSound.Command = "Open"
'
'    Timer1.Enabled = False
'    txtCollect.Text = "1"
'    txtDay.Text = "3"
'    Text1.Text = "0"
'    Query
'    Me.Top = 800
'    Me.Left = 9000
'    Me.Show
'    Call medAlwaysOn(frmBBS111, 1)
'
'    TrayI.cbSize = Len(TrayI)
'    TrayI.hwnd = pichook.hwnd 'Link the trayicon to this picturebox
'    TrayI.uId = 1&
'    TrayI.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
'    TrayI.ucallbackMessage = WM_LBUTTONDOWN
'    TrayI.hIcon = ImgList.ListImages(1).Picture
'    TrayI.szTip = "수혈요청 리스트" & Chr$(0)
'    'Create the icon
'    Shell_NotifyIcon NIM_ADD, TrayI
'End Sub
'Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
'    If blnForce Then Exit Sub
'    If cmdUpDown.tag = "1" Then
'        cmdUpDown.tag = "0"
'        cmdUpDown.Caption = "▲"
'        Me.Height = 9300
'    End If
'
'End Sub
'Private Sub Form_Unload(Cancel As Integer)
'    Set frmBBS111 = Nothing
'End Sub
'
'Private Sub optAc_Click(Index As Integer)
'    If Index = 1 Then
'        tblPtList.Row = 0: tblPtList.Col = 14: tblPtList.value = "Acting일"
'    Else
'        tblPtList.Row = 0: tblPtList.Col = 14: tblPtList.value = "처방일"
'    End If
'
'End Sub
'
'Private Sub optAct_Click(Index As Integer)
'    If Index = 1 Then
'        tblPtList.Row = 0: tblPtList.Col = 14: tblPtList.value = "Acting일"
'    Else
'        tblPtList.Row = 0: tblPtList.Col = 14: tblPtList.value = "처방일"
'    End If
'End Sub
'
'Private Sub OptS_Click(Index As Integer)
'
'    If Index = 0 Then
'        fraS(0).Visible = True
'        fraS(1).Visible = False
'        fraS(2).Visible = False
'        DoEvents
'        Timer1.Enabled = False
'        Me.Caption = "접수번호 리스트..."
'        Call Query
'    ElseIf Index = 1 Then
'
'        Timer1.Interval = 1000
'        Timer1.Enabled = True
'        fraS(1).Visible = True
'        fraS(0).Visible = False
'        fraS(2).Visible = False
'        DoEvents
'        Call CollectQuery
'    Else
'        fraS(0).Visible = False
'        fraS(1).Visible = False
'        fraS(2).Visible = True
'        DoEvents
'        Timer1.Enabled = False
'        Me.Caption = "혈액 Prep리스트..."
'        Call PrepQuery
'    End If
'
'End Sub
'
'Private Sub tblPtList_DblClick(ByVal Col As Long, ByVal Row As Long)
'    If Row = 0 Then Exit Sub
'    If Col = 1 Then Exit Sub
'    If Row > tblPtList.DataRowCnt Then Exit Sub
'    If OptS(0).value Then
'        Unload frmBBS102
'        DoEvents
'        frmBBS201_B.Show
'        DoEvents
'        tblPtList.Row = Row
''        tblPtList.Col = 8: frmBBS201.txtSpcNO.Text = tblPtList.value
''        tblPtList.Col = 3: frmBBS201.lblPtNm.Caption = tblPtList.value
'        tblPtList.Col = 2: frmBBS201_B.txtPtId.Text = tblPtList.value
'         tblPtList.ForeColor = DCM_LightBlue
'        DoEvents
'        frmBBS201_B.ClickQueryButton
'    ElseIf OptS(1).value Then
'        Unload frmBBS201
'        DoEvents
'        frmBBS102.Show
'        DoEvents
'        tblPtList.Row = Row
'        tblPtList.Col = 2: frmBBS102.txtPtId.Text = tblPtList.value
'
'
'        tblPtList.Col = 3: frmBBS102.lblPtNm.Caption = tblPtList.value
'        DoEvents
'
''        frmBBS102.ClickQueryButton
'    End If
'    If cmdUpDown.tag = "0" Then
'        cmdUpDown.tag = "1"
'        cmdUpDown.Caption = "▼"
'        Me.Height = 1450
'    End If
'End Sub
'Private Function GetTestInformation(ByVal sPtid As String) As String
'    Dim objSql As New clsCrossMatching
'    Dim RS     As Recordset
'    Dim strTmp As String
'    Dim SSQL   As String
'    Dim ii     As Integer
'
'    SSQL = objSql.TestResultXM(sPtid)
'    If SSQL <> "" Then
'    Set RS = New Recordset
'    RS.Open SSQL, DBConn
'        If Not RS.EOF Then
'             Do Until RS.EOF
'                 strTmp = strTmp & RS.Fields("workarea").value & "" & "-" & _
'                          RS.Fields("accdt").value & "" & "-" & _
'                          RS.Fields("accseq").value & "" & _
'                          "    " & RS.Fields("abbrnm10").value & "" & " : " & _
'                          RS.Fields("rstcd").value & "" & vbNewLine & "       "
'                RS.MoveNext
'            Loop
'        End If
'        Set RS = Nothing
'    End If
'
'    If strTmp <> "" Then
'        strTmp = "  ★ 관련검사 ★ " & vbNewLine & "       " & strTmp
'        GetTestInformation = strTmp
'    End If
'
'    Set objSql = Nothing
'End Function
'
'Private Sub tblPtList_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
'    Dim strtip As String
'    Dim strPtid As String
'    Dim strTmp  As String
'    Dim sICSStr As String
'
'
'    If Row = 0 Then Exit Sub
'    If tblPtList.DataRowCnt < Row Then Exit Sub
'    If OptS(2).value Then Exit Sub
'    With tblPtList
'        Call .SetTextTipAppearance("굴림체", 9, False, False, &HEEFDF2, vbBlack)
'        .Row = Row
'
'        .Col = 2: strPtid = Trim(.value)
'        sICSStr = ICSPatientString(strPtid, enICSNum.BBS_ALL)
'
'        .Col = 3: strtip = vbCrLf & " [ " & .value & sICSStr & " ]  "
'        .Col = 4: strtip = strtip & .value & vbCrLf & vbCrLf
'        .Col = 5: strtip = strtip & " 제재 : " & .value
'        .Col = 9: strtip = strtip & " [ " & .value & " ]" & vbCrLf
'        .Col = 7: strtip = strtip & " 검체번호 : " & .value & vbCrLf
'        .Col = 8: strtip = strtip & " 접수번호 : " & .value & vbCrLf
'        .Col = 9: strtip = strtip & " 검체위치 : " & .value & vbCrLf
'        .Col = 14: strtip = strtip & " 요청일자 : " & .value & vbCrLf
'        .Col = 12:
'        If .value <> "" Then strtip = strtip & " 수혈사유 : " & .value & vbCrLf
'        .Col = 10:
'        If .value <> "" Then strtip = strtip & " 병    동 : " & .value & vbCrLf
'        .Col = 13
'        If .value <> "" Then strtip = strtip & " 응급여부 : 응급검사" & vbCrLf
'    End With
'    strTmp = GetTestInformation(strPtid)
'    If strTmp <> "" Then strtip = strtip & vbNewLine & strTmp
'
'    TipWidth = 5000
'    MultiLine = 1
'    TipText = strtip
'    ShowTip = True
'
'End Sub
'Private Sub CollectQuery()
'    Dim i           As Long
'    Dim j           As Long
'
'    Dim DrRS        As Recordset
'    Dim QueryOrder  As clsQueryOrder
'    Dim ObjABO      As clsABO
'
'    Dim accno       As String
'    Dim reason      As String
'    Dim status      As String
'    Dim spcno       As String
'    Dim storeleg    As String
'    Dim storerow    As String
'    Dim storecol    As String
'    Dim center      As String
'
'    Dim strLeg      As String
'    Dim strRow      As String
'    Dim strCol      As String
'    Dim MaxRowCnt   As Long
'    Dim PreCnt      As Long
'
'    Dim objPrgBar   As clsProgress
'
'    Dim lngDay      As Long
'
'    lngDay = "-" & Val(txtCollect.Text)
'
'    Set QueryOrder = New clsQueryOrder
'
'
'    If optAct(1).value = True Then
'        tblPtList.Row = 0: tblPtList.Col = 14: tblPtList.value = "Acting일"
'        'tblPtList.Row = 0: tblPtList.Col = 14: tblPtList.value = "처방일"
'        QueryOrder.ActFG = "Acting"
'    Else
'        tblPtList.Row = 0: tblPtList.Col = 14: tblPtList.value = "처방일"
'        QueryOrder.ActFG = ""
'    End If
'
'
'    Set DrRS = New Recordset
'    DrRS.Open QueryOrder.CollectionList(Format(DateAdd("d", lngDay, Now), PRESENTDATE_FORMAT), Format(Now, PRESENTDATE_FORMAT)), DBConn
'
'
'    If DrRS Is Nothing Then
'        Set DrRS = Nothing
'        Set QueryOrder = Nothing
'        Exit Sub
'    End If
'
'
'
'    Set ObjABO = New clsABO
'
'    Set objPrgBar = New clsProgress
''    Set objPrgBar.StatusBar = medMain.stsBar
'    objPrgBar.Container = MainFrm.stsBar
'
'    objPrgBar.Min = 1
'    objPrgBar.Max = DrRS.RecordCount
'
'    PreCnt = tblPtList.DataRowCnt
'
'    tblPtList.MaxRows = 0
'    With tblPtList
'        .Row = 0: .Col = 8: .value = "비고"
'        .ReDraw = False
'        For i = 1 To DrRS.RecordCount
'
'            objPrgBar.value = i
'
'            MaxRowCnt = MaxRowCnt + 1
'            .MaxRows = MaxRowCnt
'            .Row = MaxRowCnt
'
'            '==========================
'            '보관장소및 검체 번호구하기
'            '==========================
'            Call QueryOrder.GetSpcNoAndStore(DrRS.Fields("ptid").value & "", spcno, storeleg, storerow, storecol, center)
'            '================
'            '접수번호확인작업
'            '================
'
'            accno = Trim(DrRS.Fields("accdt").value & "") & "-" & Val(Trim(DrRS.Fields("accseq").value & ""))
'            If accno = "-0" Then accno = "" 'accno = "미접수"
'
'            '==================
'            '수혈사유 구하기...
'            '==================
'
'            reason = QueryOrder.GetTransReason(DrRS.Fields("ptid").value & "", DrRS.Fields("orddt").value & "", DrRS.Fields("ordno").value & "")
'            If reason = "" Then reason = "(없음)"
'
'            .Col = 2:   .value = DrRS.Fields("ptid").value & ""
'            .Col = 3:   .value = DrRS.Fields("ptnm").value & ""
'
'            '================
'            '혈액형을 구한다.
'            '================
'
'            ObjABO.PtId = DrRS.Fields("ptid").value & ""
'            If ObjABO.GetABO = False Then
'                .Col = 4:    .value = ""
'            Else
'                .Col = 4:    .value = ObjABO.ABO & ObjABO.Rh
'            End If
'
'            .Col = 5:   .value = DrRS.Fields("testnm").value & ""
'            .Col = 6:   .value = DrRS.Fields("unitqty").value & ""
'
'            .Col = 8:
'                Select Case DrRS.Fields("bussdiv").value & ""
'                    Case "1"
'                        .Col = 8: .value = DrRS.Fields("deptcd").value & ""
'                    Case "2"
'                        .Col = 8: .value = DrRS.Fields("wardid").value & ""
'                        If .value <> "" Then
'                            If DrRS.Fields("hosilid").value & "" <> "" Then
'                                .value = .value & "-" & DrRS.Fields("hosilid").value & ""
'                            End If
'                        End If
'                End Select
'            .Col = 10:   .value = DrRS.Fields("wardid").value & ""
'            .Col = 11:  .value = DrRS.Fields("hosilid").value & ""
'            .Col = 12:  .value = Trim(Trim0(reason))
'            .Col = 13:  .value = IIf(DrRS.Fields("statfg").value & "" = "1", "Y", "")
'                        .ForeColor = vbRed
'                        .FontBold = True
'            .Col = 14:
'
'
'            If optAc(1).value Then
'                .value = Format("" & DrRS.Fields("orddt").value & "", CS_DateLongMask) & " " & _
'                                     Format(Mid("" & DrRS.Fields("ordtm").value, 1, 4), "0#:##")
'            Else
'                .value = Format("" & DrRS.Fields("entdt").value, CS_DateLongMask) & " " & _
'                                     Format(Mid("" & DrRS.Fields("enttm").value, 1, 4), "0#:##")
'            End If
'
'
'            '--------------------------
'            '검체번호와 보관장소 구하기
'            '--------------------------
'            If storerow = "0" Then storerow = ""
'            If storecol = "0" Then storecol = ""
'
'            .Col = 9:   .value = storeleg & ";" & storerow & ";" & storecol
'
'            .Col = 7:   .value = spcno
'
'            If spcno = "" Then
'                .Col = 7:   .value = "" '.value = "미채혈"
'            Else
'                If storeleg = "" Then
'                    .Col = 9:    .value = ""
'                Else
'                    .Col = 9:    .value = storeleg & "(" & storerow & "," & storecol & ")"
'                End If
'            End If
'Skip:
'            DrRS.MoveNext
'        Next i
'        .ReDraw = True
'    End With
'
'    If tblPtList.MaxRows < 33 Then tblPtList.MaxRows = 33
'
'    Set DrRS = Nothing
'    Set ObjABO = Nothing
'    Set objPrgBar = Nothing
'    Set QueryOrder = Nothing
'
'    If tblPtList.DataRowCnt <> PreCnt Then
'    'If tblPtList.MaxRows > 0 Then
'        'Me.Show
'        Me.WindowState = 0
'
'        blnStopFg = False
'
'        mmSound.Command = "Stop"
'        mmSound.Command = "Close"
'        mmSound.FileName = gBloodRequestMusic
'        mmSound.Enabled = True
'        mmSound.Command = "Open"
'        mmSound.StopVisible = True
'        mmSound.StopEnabled = True
'        mmSound.Command = "Prev"
'        If blnSound Then mmSound.Command = "Play"
'    End If
'
'End Sub
'Private Sub Query()
'    Dim i           As Long
'    Dim j           As Long
'
'    Dim DrRS        As Recordset
'    Dim QueryOrder  As clsQueryOrder
'    Dim ObjABO      As clsABO
'
'    Dim accno       As String
'    Dim reason      As String
'    Dim status      As String
'    Dim spcno       As String
'    Dim storeleg    As String
'    Dim storerow    As String
'    Dim storecol    As String
'    Dim center      As String
'
'    Dim strLeg      As String
'    Dim strRow      As String
'    Dim strCol      As String
'    Dim MaxRowCnt   As Long
'    Dim blnComplete As Boolean
'
'    Dim objPrgBar   As clsProgress
'
'
'    '윗줄과 같은내용이면 글자를 감추기 위한변수들
'
'
'    Dim lngDay      As Long
'
'    lngDay = "-" & Val(txtDay.Text)
'
'
'    Set QueryOrder = New clsQueryOrder
'    If optAc(1).value Then
'        tblPtList.Row = 0: tblPtList.Col = 14: tblPtList.value = "Acting일"
'
'        QueryOrder.ActFG = "Acting"
'    Else
'        tblPtList.Row = 0: tblPtList.Col = 14: tblPtList.value = "처방일"
'        QueryOrder.ActFG = ""
'    End If
'        Set DrRS = New Recordset
'        DrRS.Open QueryOrder.QueryAccdt(Format(DateAdd("d", lngDay, Now), PRESENTDATE_FORMAT), Format(Now, PRESENTDATE_FORMAT)), DBConn
'
'
'    If DrRS Is Nothing Then
'        Set DrRS = Nothing
'        Set QueryOrder = Nothing
'        Exit Sub
'    End If
'
'
'
'    Set ObjABO = New clsABO
'
'    Set objPrgBar = New clsProgress
''    Set objPrgBar.StatusBar = medMain.stsBar
'    objPrgBar.Container = MainFrm.stsBar
'
'    objPrgBar.Min = 1
'    objPrgBar.Max = DrRS.RecordCount
'
'    tblPtList.MaxRows = 0
'    With tblPtList
'        .Row = 0: .Col = 8: .value = "접수번호"
'        .ReDraw = False
'        For i = 1 To DrRS.RecordCount
'
'            objPrgBar.value = i
'
'            '=====================================
'            '처방이 완료 되었는지 확인작업
'            'Unitqty=assigncnt-cancelcnt 이면 완료
'            '=====================================
'            If (Val(DrRS.Fields("assigncnt").value & "") - _
'                Val(DrRS.Fields("assigncancelcnt").value & "") - _
'                Val(DrRS.Fields("retcnt").value & "") - _
'                Val(DrRS.Fields("expcnt").value & "")) = _
'                Val(DrRS.Fields("unitqty").value & "") Then
'                '(어사인-취소-반환-폐기)=처방수량
'                GoTo Skip
'
'            End If
'
''            blnComplete = CompleteOrderChk(DrRS.Fields("accdt").value & "", DrRS.Fields("accseq").value & "", DrRS.Fields("unitqty").value & "")
''            If blnComplete Then GoTo Skip
'            MaxRowCnt = MaxRowCnt + 1
'            .MaxRows = MaxRowCnt
'            .Row = MaxRowCnt
'
'
'            '==========================
'            '보관장소및 검체 번호구하기
'            '==========================
'            Call QueryOrder.GetSpcNoAndStore(DrRS.Fields("ptid").value & "", spcno, storeleg, storerow, storecol, center)
'            If medGetP(center, 1, vbTab) <> ObjSysInfo.BuildingCd Then GoTo Skip
'
'
'            '================
'            '접수번호확인작업
'            '================
'
'            accno = Trim(DrRS.Fields("accdt").value & "") & "-" & Val(Trim(DrRS.Fields("accseq").value & ""))
'            If accno = "-0" Then accno = "" 'accno = "미접수"
'
'            '==================
'            '수혈사유 구하기...
'            '==================
'
'            reason = QueryOrder.GetTransReason(DrRS.Fields("ptid").value & "", DrRS.Fields("orddt").value & "", DrRS.Fields("ordno").value & "")
'            If reason = "" Then reason = "(없음)"
'
'            .Col = 2:   .value = DrRS.Fields("ptid").value & ""
'            .Col = 3:   .value = DrRS.Fields("ptnm").value & ""
'
'            '================
'            '혈액형을 구한다.
'            '================
'
'            ObjABO.PtId = DrRS.Fields("ptid").value & ""
'            If ObjABO.GetABO = False Then
'                .Col = 4:    .value = ""
'            Else
'                .Col = 4:    .value = ObjABO.ABO & ObjABO.Rh
'            End If
'
'            .Col = 5:   .value = DrRS.Fields("testnm").value & ""
'            .Col = 6:   .value = DrRS.Fields("unitqty").value & ""
'
'            .Col = 8:   .value = accno: .ForeColor = vbRed
'
'
'            .Col = 10:   .value = DrRS.Fields("wardid").value & ""
'            .Col = 11:  .value = DrRS.Fields("hosilid").value & ""
'            .Col = 12:  .value = Trim(Trim0(reason))
'            .Col = 13:  .value = IIf(DrRS.Fields("statfg").value = "1", "Y", "")
'                        .ForeColor = vbRed
'                        .FontBold = True
'            .Col = 14:
'
'
'            If optAc(1).value Then
'                .value = Format("" & DrRS.Fields("orddt").value, CS_DateLongMask) & " " & _
'                                     Format(Mid("" & DrRS.Fields("ordtm").value, 1, 4), "0#:##")
'            Else
'                .value = Format("" & DrRS.Fields("entdt").value, CS_DateLongMask) & " " & _
'                                     Format(Mid("" & DrRS.Fields("enttm").value, 1, 4), "0#:##")
'            End If
'
'
'            '--------------------------
'            '검체번호와 보관장소 구하기
'            '--------------------------
'            If storerow = "0" Then storerow = ""
'            If storecol = "0" Then storecol = ""
'
'            .Col = 9:   .value = storeleg & ";" & storerow & ";" & storecol
'
'            .Col = 7:   .value = spcno
'
'            If spcno = "" Then
'                .Col = 7:   .value = "" '.value = "미채혈"
'            Else
'                If storeleg = "" Then
'                    .Col = 9:    .value = ""
'                Else
'                    .Col = 9:    .value = storeleg & "(" & storerow & "," & storecol & ")"
'                End If
'            End If
'Skip:
'            DrRS.MoveNext
'        Next i
'        .ReDraw = True
'    End With
'
'    If tblPtList.MaxRows < 33 Then tblPtList.MaxRows = 33
'
'
'    Set DrRS = Nothing
'    Set ObjABO = Nothing
'    Set objPrgBar = Nothing
'    Set QueryOrder = Nothing
'
'
'
'
'End Sub
'
'
'
'Private Function CompleteOrderChk(ByVal accdt As String, ByVal accseq As String, ByVal unitqty As Long) As Boolean
'    Dim objXM As New clsCrossMatching
'    Dim A_Cnt As Long   'Assign수량
'    Dim C_Cnt As Long   'Assign Cancel 수량
'    Dim O_Cnt As Long   '출고수량
'    Dim R_Cnt As Long   '반환수량
'    Dim X_Cnt As Long   '폐기수량
'    Dim T_Cnt As Long   '총Assign 수량
'
'
'    'CompleteOrderChk=True이면 완결처방
'    'CompleteOrderChk=미완결처방
'    CompleteOrderChk = False
'    If accdt <> "" Then
'
'        With objXM
'            .Assign_Cnt accdt, Val(accseq)
'            A_Cnt = .AssignCnt
'            C_Cnt = .CancelCnt
'            O_Cnt = .OutCnt
'            R_Cnt = .RetCnt
'            X_Cnt = .ExpCnt
'        End With
'
'        T_Cnt = A_Cnt - C_Cnt ' - R_Cnt - X_Cnt
'
'        If unitqty = T_Cnt Then
'            CompleteOrderChk = True
'        End If
'    End If
'    Set objXM = Nothing
'
'End Function
'
'Private Sub Timer1_Timer()
'    Static TimeCount As Long
'    Static ImgCount As Integer
'
'    Me.Icon = ImgList.ListImages(TimeCount Mod 3 + 1).Picture
'    TrayI.hIcon = ImgList.ListImages(TimeCount Mod 3 + 1).Picture
'
'    Shell_NotifyIcon NIM_MODIFY, TrayI
'
'    TrayI.szTip = Me.Caption & Chr$(0)
'
'    TimeCount = TimeCount + 1
'
'    If Timer1.Enabled Then
'        Me.Caption = "채혈항목 리스트..(RUN)..." & TimeCount
'    Else
'        If OptS(0).value Then
'            Me.Caption = "접수번호 리스트..."
'        ElseIf OptS(2).value Then
'            Me.Caption = "혈액 Prep 리스트..."
'        End If
'    End If
'
'    If TimeCount = 300 Then
'        mmSound.Command = "Stop"
'        mmSound.Command = "Close"
'        DoEvents
'        Call CollectQuery
'        TimeCount = 0
'    End If
'
'End Sub
'
'
'Private Sub PrepQuery()
'    Dim i           As Long
'    Dim j           As Long
'
'    Dim DrRS        As Recordset
'    Dim ObjABO      As clsABO
'
'    Dim accno       As String
'    Dim reason      As String
'    Dim status      As String
'    Dim spcno       As String
'    Dim storeleg    As String
'    Dim storerow    As String
'    Dim storecol    As String
'    Dim center      As String
'
'    Dim strLeg      As String
'    Dim strRow      As String
'    Dim strCol      As String
'    Dim MaxRowCnt   As Long
'
'    Dim objPrgBar   As clsProgress
'
'    Dim lngDay      As Long
'
'    lngDay = "-" & Val(Text1.Text)
'
'
'
'    Set DrRS = New Recordset
'    DrRS.Open PrepList(Format(DateAdd("d", lngDay, Now), PRESENTDATE_FORMAT), Format(Now, PRESENTDATE_FORMAT)), DBConn
'
'
''    If DrRS Is Nothing Then
''        Set DrRS = Nothing
''        Exit Sub
''    End If
'
'
'
'    Set ObjABO = New clsABO
'
'    Set objPrgBar = New clsProgress
''    Set objPrgBar.StatusBar = medMain.stsBar
'    objPrgBar.Container = MainFrm.stsBar
'
'    objPrgBar.Min = 1
'    objPrgBar.Max = DrRS.RecordCount
'
'    tblPtList.MaxRows = 0
'    With tblPtList
'        .Row = 0: .Col = 8: .value = "비고"
'
'        .ReDraw = False
'        For i = 1 To DrRS.RecordCount
'
'            objPrgBar.value = i
'
'            MaxRowCnt = MaxRowCnt + 1
'            .MaxRows = MaxRowCnt
'            .Row = MaxRowCnt
'
'
'
'            .Col = 2:   .value = DrRS.Fields("ptid").value & ""
'            .Col = 3:   .value = DrRS.Fields("ptnm").value & ""
'                        If PrepOrder(DrRS.Fields("serial").value & "") = True Then
'                            .ForeColor = DCM_LightBlue: .FontBold = True
'
'                        End If
'            '================
'            '혈액형을 구한다.
'            '================
'
'            ObjABO.PtId = DrRS.Fields("ptid").value & ""
'            If ObjABO.GetABO = False Then
'                .Col = 4:    .value = ""
'            Else
'                .Col = 4:    .value = ObjABO.ABO & ObjABO.Rh
'            End If
'
'            .Col = 5:   .value = DrRS.Fields("testnm").value & ""
'
'            Select Case DrRS.Fields("bussdiv").value & ""
'                Case "1"
'                    .Col = 8: .value = DrRS.Fields("deptcd").value & ""
'                Case "2"
'                    .Col = 8: .value = DrRS.Fields("wardid").value & ""
'                    If .value <> "" Then
'                        If DrRS.Fields("hosilid").value & "" <> "" Then
'                            .value = .value & "-" & DrRS.Fields("hosilid").value & ""
'                        End If
'                    End If
'            End Select
'
'
'
'            If DrRS.Fields("dcfg").value & "" = "1" Then
'                .ForeColor = DCM_LightRed
'            End If
'            .Col = 9: .value = DrRS.Fields("serial").value & ""
'
'            .Col = 10: .value = IIf(DrRS.Fields("dcfg").value & "" = "1", "Y", "")
'
'            .Col = 14: .value = Format(DrRS.Fields("orddt").value & "", "####-0#-0#")
'
'            DrRS.MoveNext
'        Next i
'        .ReDraw = True
'    End With
'
'    If tblPtList.MaxRows < 33 Then tblPtList.MaxRows = 33
'
'    Set DrRS = Nothing
'    Set ObjABO = Nothing
'    Set objPrgBar = Nothing
'
'End Sub
'
'Public Function PrepOrder(ByVal Serial As String) As Boolean
'
'    Dim RS    As Recordset
'    Dim SSQL  As String
'
'    SSQL = " select * from " & T_LAB102 & " where " & DBW("ocsordno", Serial, 2)
'    Set RS = New Recordset
'    RS.Open SSQL, DBConn
'    If Not RS.EOF Then
'        PrepOrder = True
'    End If
'    Set RS = Nothing
'
'End Function
'
'
'Public Function PrepList(ByVal FrDt As String, ByVal ToDt As String) As String
'    Dim SSQL As String
'
'    SSQL = "select distinct a.ptid,a.statfg,a.serial,a.wardid,a.hosilid,a.deptcd,a.dcfg,a.orddt," & _
'           "                a.bussdiv, b." & F_PTNM & " as ptnm ,a.testcd,c.abbrnm10 as testnm"
'
'    SSQL = SSQL & " from s2bbs001 c," & T_HIS001 & " b,s2bbs_prep a"
'    SSQL = SSQL & " where " & DBW("a.orddt>=", FrDt) & " and " & DBW("a.orddt<=", ToDt)
'    SSQL = SSQL & " and a.ptid=b." & F_PTID & " and a.testcd=c.testcd"
'    SSQL = SSQL & " order by ptid,orddt"
'
'    PrepList = SSQL
'End Function
'
''-- Sound 추가작업 By M.G.Choi 2002.09.28/////////////////////////////////////////////////////////////
'Private Sub imgSound_Click(Index As Integer)
'
'    imgSound(Index).Visible = False
'    imgSound((Index + 1) Mod 2).Visible = True
'    blnSound = Choose(Index + 1, False, True)
'    If Not blnSound Then mmSound.Command = "Stop"
'
'End Sub
'
'Private Sub mmSound_Done(NotifyCode As Integer)
'
'    If Not blnStopFg Then
'        mmSound.Command = "Stop"
'        mmSound.Command = "Close"
'        mmSound.FileName = gBloodRequestMusic
'        mmSound.Enabled = True
'        mmSound.Command = "Open"
'        If blnSound Then mmSound.Command = "Play"
'        mmSound.StopVisible = True
'        mmSound.StopEnabled = True
'    End If
'
'End Sub
'
'Private Sub mmSound_PauseClick(Cancel As Integer)
'    mmSound.Command = "Stop"
'    'mmSound.StopVisible = False
'    mmSound.StopEnabled = False
'
'    blnStopFg = Not blnStopFg
'    Timer1.Enabled = Not Timer1.Enabled
'    If Timer1.Enabled Then
'        Me.Caption = "접수번호리스트..(RUN)"
'    Else
'        Me.Caption = "접수번호리스트..(STOP)"
'    End If
'End Sub
'
'Private Sub mmSound_StopClick(Cancel As Integer)
'    blnStopFg = True
''    mmSound.StopVisible = False
'    mmSound.StopEnabled = False
'End Sub
''/////////////////////////////////////////////////////////////////////////////////////////////////////
