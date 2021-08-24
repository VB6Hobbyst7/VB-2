VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmResult 
   BorderStyle     =   1  '¥‹¿œ ∞Ì¡§
   Caption         =   "DumpTEST"
   ClientHeight    =   5040
   ClientLeft      =   2100
   ClientTop       =   2340
   ClientWidth     =   8835
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   8835
   Begin VB.Timer tmrSample 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7515
      Top             =   3915
   End
   Begin VB.CheckBox chkSendgbn 
      Caption         =   "¡ÔΩ√∫∏≥ª±‚"
      BeginProperty Font 
         Name            =   "±º∏≤√º"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   7470
      TabIndex        =   16
      Top             =   945
      Width           =   1275
   End
   Begin VB.TextBox txtSend 
      BeginProperty Font 
         Name            =   "±º∏≤"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   135
      TabIndex        =   13
      Top             =   900
      Width           =   7260
   End
   Begin VB.TextBox txtResult 
      BeginProperty Font 
         Name            =   "±º∏≤"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2850
      Left            =   135
      MultiLine       =   -1  'True
      ScrollBars      =   2  'ºˆ¡˜
      TabIndex        =   7
      Top             =   1260
      Width           =   7260
   End
   Begin VB.TextBox txtRecve 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "±º∏≤"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   135
      TabIndex        =   6
      Top             =   4140
      Width           =   7260
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   4230
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      ParityReplace   =   64
      RThreshold      =   1
      SThreshold      =   1
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   690
      Left            =   7605
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   45
      Width           =   1200
      _Version        =   65536
      _ExtentX        =   2117
      _ExtentY        =   1217
      _StockProps     =   78
      Caption         =   "¡æ∑·ESC"
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±º∏≤√º"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "REUSLT.frx":0000
   End
   Begin Threed.SSCommand cmdAck 
      Height          =   465
      Left            =   8130
      TabIndex        =   1
      Top             =   1260
      Width           =   570
      _Version        =   65536
      _ExtentX        =   1005
      _ExtentY        =   820
      _StockProps     =   78
      Caption         =   "ACK"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±º∏≤"
         Size            =   10.5
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
      Picture         =   "REUSLT.frx":08DA
   End
   Begin Threed.SSCommand cmdEtx 
      Height          =   465
      Left            =   7515
      TabIndex        =   2
      Top             =   1800
      Width           =   570
      _Version        =   65536
      _ExtentX        =   1005
      _ExtentY        =   820
      _StockProps     =   78
      Caption         =   "ETX"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±º∏≤"
         Size            =   10.5
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
      Picture         =   "REUSLT.frx":08F6
   End
   Begin Threed.SSCommand cmdEot 
      Height          =   465
      Left            =   7515
      TabIndex        =   3
      Top             =   2340
      Width           =   570
      _Version        =   65536
      _ExtentX        =   1005
      _ExtentY        =   820
      _StockProps     =   78
      Caption         =   "EOT"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±º∏≤"
         Size            =   10.5
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
      Picture         =   "REUSLT.frx":0912
   End
   Begin Threed.SSCommand cmdStx 
      Height          =   465
      Left            =   7515
      TabIndex        =   4
      Top             =   1260
      Width           =   570
      _Version        =   65536
      _ExtentX        =   1005
      _ExtentY        =   820
      _StockProps     =   78
      Caption         =   "STX"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±º∏≤"
         Size            =   10.5
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
      Picture         =   "REUSLT.frx":092E
   End
   Begin Threed.SSCommand cmdNak 
      Height          =   465
      Left            =   8100
      TabIndex        =   5
      Top             =   1800
      Width           =   570
      _Version        =   65536
      _ExtentX        =   1005
      _ExtentY        =   820
      _StockProps     =   78
      Caption         =   "NAK"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±º∏≤"
         Size            =   10.5
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
      Picture         =   "REUSLT.frx":094A
   End
   Begin Threed.SSPanel Panel3D11 
      Height          =   735
      Left            =   135
      TabIndex        =   8
      Top             =   90
      Width           =   4110
      _Version        =   65536
      _ExtentX        =   7250
      _ExtentY        =   1296
      _StockProps     =   15
      ForeColor       =   8454143
      BackColor       =   -2147483644
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelOuter      =   1
      BevelInner      =   2
      Begin Threed.SSPanel SSPanel5 
         Height          =   375
         Left            =   225
         TabIndex        =   9
         Top             =   180
         Width           =   1530
         _Version        =   65536
         _ExtentX        =   2699
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "∆ƒ¿œ∏Ì"
         ForeColor       =   8454143
         BackColor       =   12582912
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±º∏≤"
            Size            =   11.24
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         BorderWidth     =   1
      End
      Begin VB.Label lblFilename 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  '¥‹¿œ ∞Ì¡§
         BeginProperty Font 
            Name            =   "±º∏≤"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1845
         TabIndex        =   10
         Top             =   225
         Width           =   1815
      End
   End
   Begin Threed.SSCommand cmdEnq 
      Height          =   465
      Left            =   8100
      TabIndex        =   11
      Top             =   2340
      Width           =   570
      _Version        =   65536
      _ExtentX        =   1005
      _ExtentY        =   820
      _StockProps     =   78
      Caption         =   "ENQ"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±º∏≤"
         Size            =   10.5
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
      Picture         =   "REUSLT.frx":0966
   End
   Begin Threed.SSCommand cmdClear 
      Height          =   690
      Left            =   5175
      TabIndex        =   12
      Top             =   45
      Width           =   1200
      _Version        =   65536
      _ExtentX        =   2117
      _ExtentY        =   1217
      _StockProps     =   78
      Caption         =   "Clear"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±º∏≤√º"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "REUSLT.frx":0982
   End
   Begin Threed.SSCommand cmdSend 
      Height          =   690
      Left            =   6390
      TabIndex        =   14
      Top             =   45
      Width           =   1200
      _Version        =   65536
      _ExtentX        =   2117
      _ExtentY        =   1217
      _StockProps     =   78
      Caption         =   "Send"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±º∏≤"
         Size            =   10.5
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "REUSLT.frx":099E
   End
   Begin Threed.SSCommand cmdFS 
      Height          =   465
      Left            =   8100
      TabIndex        =   15
      Top             =   2880
      Width           =   570
      _Version        =   65536
      _ExtentX        =   1005
      _ExtentY        =   820
      _StockProps     =   78
      Caption         =   "FS"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±º∏≤"
         Size            =   10.5
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
      Picture         =   "REUSLT.frx":09BA
   End
   Begin Threed.SSCommand cmdCr 
      Height          =   465
      Left            =   7515
      TabIndex        =   17
      Top             =   2880
      Width           =   570
      _Version        =   65536
      _ExtentX        =   1005
      _ExtentY        =   820
      _StockProps     =   78
      Caption         =   "CR"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±º∏≤"
         Size            =   10.5
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
      Picture         =   "REUSLT.frx":09D6
   End
   Begin Threed.SSCommand cmdLf 
      Height          =   465
      Left            =   7515
      TabIndex        =   18
      Top             =   3375
      Width           =   570
      _Version        =   65536
      _ExtentX        =   1005
      _ExtentY        =   820
      _StockProps     =   78
      Caption         =   "LF"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±º∏≤"
         Size            =   10.5
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
      Picture         =   "REUSLT.frx":09F2
   End
   Begin Threed.SSCommand cmdCnv 
      Height          =   465
      Left            =   8100
      TabIndex        =   19
      Top             =   3375
      Width           =   570
      _Version        =   65536
      _ExtentX        =   1005
      _ExtentY        =   820
      _StockProps     =   78
      Caption         =   "con"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±º∏≤"
         Size            =   10.5
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
      Picture         =   "REUSLT.frx":0A0E
   End
   Begin Threed.SSCommand cmdSoh 
      Height          =   465
      Left            =   8100
      TabIndex        =   20
      Top             =   3870
      Width           =   570
      _Version        =   65536
      _ExtentX        =   1005
      _ExtentY        =   820
      _StockProps     =   78
      Caption         =   "SOH"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±º∏≤"
         Size            =   10.5
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
      Picture         =   "REUSLT.frx":0A2A
   End
End
Attribute VB_Name = "frmResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
'*                                                              *
'*  SLBI_30F  = BIOMIC ∞·∞˙πﬁ±‚                          *
'*                                                              *
'*  System    : Ω≈√Ãºº∫Í∂ıΩ∫∫¥ø¯ Ω√Ω∫≈€                         *
'*  Subsystem : ¿”ªÛ∫¥∏Æ ∞¸∏Æ Ω√Ω∫≈€                            *
'*                                                              *
'*  Designed  : 1997-08-30                                      *
'*  Coded     : 1997-08-30 ¿Ø¿∫¿⁄                               *
'*  Modified  :                                                 *
'*                                                              *
'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
Option Explicit

Dim f_strBuffer As String
Dim f_strDta()  As String, f_intCnt As Integer
Dim f_intIdx    As Integer



Private Sub cmdClear_Click()

    txtSend.Text = ""
    txtResult.Text = ""
    txtRecve.Text = ""
End Sub

Private Sub cmdCnv_Click()
    
'    Dim wkbuf   As String
    Dim sTmp    As String
    Dim iCnt    As Integer
    Dim wkbuf   As Byte

    Open App.Path + "\" + "genius_1.dat" For Append As #3
    Open "a:\Giorn.ris" For Input As #4

'    wkbuf = "": iCnt = 0
    Do While Not EOF(4)
        wkbuf = Input(1, #4)

'        If 60 <= Asc(wkbuf) And Asc(wkbuf) <= 69 Then MsgBox "B0"

'        If Asc(wkbuf) = &HB0B0 Then wkbuf = "  "
'        If 60 <= Asc(wkbuf) And Asc(wkbuf) <= 69 Then wkbuf = " " + Chr(Asc(wkbuf) - 12)

            Print #3, wkbuf;
    Loop

    Close #3
    Close #4
    MsgBox "¿€æ˜øœ∑·"
    
    

End Sub

Private Sub cmdCr_Click()

    If chkSendgbn.Value = vbChecked Then
        MSComm1.Output = Chr$(13)
    Else
        txtSend.Text = txtSend.Text + Chr$(13)
        txtSend.SetFocus
    End If

End Sub

Private Sub cmdEnq_Click()

    If chkSendgbn.Value = vbChecked Then
        MSComm1.Output = Chr$(5)
    Else
        txtSend.Text = txtSend.Text + Chr$(5)
        txtSend.SetFocus
    End If

End Sub

Private Sub cmdExit_Click()

    Unload Me

End Sub

Private Sub cmdFS_Click()

    If chkSendgbn.Value = vbChecked Then
        MSComm1.Output = Chr$(28)
    Else
        txtSend.Text = txtSend.Text + Chr$(28)
        txtSend.SetFocus
    End If
    
End Sub

Private Sub cmdLf_Click()

    If chkSendgbn.Value = vbChecked Then
        MSComm1.Output = Chr$(10)
    Else
        txtSend.Text = txtSend.Text + Chr$(10)
        txtSend.SetFocus
    End If

End Sub


Private Sub cmdSend_Click()

    ReDim f_strDta(1 To 10) As String
    
    f_intCnt = 8
    f_strDta(1) = "[ 0,702,01,100603, 92610,53509,RE, 1, 2,RO,#########,SE,0610002ER  ,                    ,                         ,                         ,                  ,               , ,            ,                  ,      ,    ,                    ,   ,5,      ,M,                         ,#######,####,####,######,  2,30A ,31A ]2B" + Chr(13) + Chr(10)
    f_strDta(2) = "[ 0,702,03,100603, 93625,53509,     5646, 1, 2,0610002ER  ,30A ,9UN,302260,63, 1,       26,#########,2,0, 9,NA,NR,NA,0,NA,25.594528,         ,NO,NO,NO,NO,NO,NO,NO,NO,NO,NO,NO,NO,NO,NO,NO,NO,1.0000,NR,####################]99" + Chr(13) + Chr(10)
    f_strDta(3) = "[ 0,702,03,100603, 93641,53509,     6806, 1, 2,0610002ER  ,31A ,030,302259,64, 1,       31,#########,2,0, 9,NA,NR,NA,0,NA,31.242405,         ,NO,NO,NO,NO,NO,NO,NO,NO,NO,NO,NO,NO,NO,NO,NO,NO,1.0000,NR,####################]F2" + Chr(13) + Chr(10)
    f_strDta(4) = "[ 0,702,05,100603, 93641,53509,10002      , 1, 2]ED" + Chr(13) + Chr(10)
    
    f_strDta(5) = "[ 0,702,01,100603, 92610,53509,RE, 1, 3,RO,#########,SE,0610003ER  ,                    ,                         ,                         ,                  ,               , ,            ,                  ,      ,    ,                    ,   ,5,      ,M,                         ,#######,####,####,######,  2,30A ,31A ]2B" + Chr(13) + Chr(10)
    f_strDta(6) = "[ 0,702,03,100603, 93625,53509,     5646, 1, 2,0610003ER  ,30A ,9UN,302260,63, 1,      190,#########,2,0, 9,NA,NR,NA,0,NA,25.594528,         ,NO,NO,NO,NO,NO,NO,NO,NO,NO,NO,NO,NO,NO,NO,NO,NO,1.0000,NR,####################]99" + Chr(13) + Chr(10)
    f_strDta(7) = "[ 0,702,03,100603, 93641,53509,     6806, 1, 2,0610003ER  ,31A ,030,302259,64, 1,       50,#########,2,0, 9,NA,NR,NA,0,NA,31.242405,         ,NO,NO,NO,NO,NO,NO,NO,NO,NO,NO,NO,NO,NO,NO,NO,NO,1.0000,NR,####################]F2" + Chr(13) + Chr(10)
    f_strDta(8) = "[ 0,702,05,100603, 93641,53509,10002      , 1, 2]ED" + Chr(13) + Chr(10)
    
    f_intIdx = 0:   tmrSample.Enabled = False
    
End Sub

Private Sub cmdSoh_Click()

    MSComm1.Output = Chr$(1)

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyEscape:    Call cmdExit_Click
    End Select

End Sub

Private Sub Form_Load()

    lblFilename.Caption = "k4500_rst.log"
    
    Dim sPath   As String

    With MSComm1
        .CommPort = 2
        'º”µµ,∆‰∏Æ∆º,≈◊¿Ã≈∏∫Ò∆Æ,stop bit
        .Settings = "9600,n,8,1"
        .PortOpen = True
        .RTSEnable = True
        .RThreshold = 1
    End With

    sPath = App.Path
    Open sPath + "\" + "k4500_rst.log" For Append As #1
    
    tmrSample.Enabled = False

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If MSComm1.PortOpen Then MSComm1.PortOpen = False
 
    Close #1

End Sub

Private Sub MSComm1_OnComm()
    
    Me.MousePointer = vbHourglass

    Dim wkbuf As String
    Dim wkdat As String
    Dim ix1 As Integer, cnt As Integer

    Dim temp   As Variant
    Dim stemp  As String
    
    Select Case MSComm1.CommEvent
        ' Events
        Case MSCOMM_EV_SEND     ' There are SThreshold number of
           
                                ' character in the transmit buffer.
        Case MSCOMM_EV_RECEIVE  ' Received RThreshold # of chars.
            wkbuf = MSComm1.Input

            Print #1, wkbuf;

            Select Case Asc(wkbuf)
            Case 1  '-- SOH
            Case 2  '-- STX
            Case 3  '-- ETX
                    f_intIdx = f_intIdx + 1
                    tmrSample.Enabled = True

            Case 6  '-- ACK
                    f_intIdx = f_intIdx + 1
                    tmrSample.Enabled = True
                    
            Case 10 '-- LF
                    
            Case 21 '-- NAK
                    tmrSample.Enabled = True
            End Select

            txtRecve.Text = wkbuf
            txtResult.Text = txtResult.Text + wkbuf

        Case MSCOMM_EV_CTS      ' Change in the CTSj
        Case MSCOMM_EV_DSR      ' Change in the DSR line.
        Case MSCOMM_EV_CD       ' Change in the CD line.
        Case MSCOMM_EV_RING     ' Change in the Ring Indicator.
        ' Errors
        Case MSCOMM_ER_BREAK    ' A Break was received.
        ' Code to handle a BREAK goes here, and so on.
        Case MSCOMM_ER_CTSTO    ' CTS Timeout.
        Case MSCOMM_ER_DSRTO    ' DSR Timeout.
        Case MSCOMM_ER_FRAME    ' Framing Error.
        Case MSCOMM_ER_OVERRUN  ' Data Lost.
        Case MSCOMM_ER_CDTO     ' CD (RLSD) Timeout.
        Case MSCOMM_ER_RXOVER   ' Receive buffer overflow.
        Case MSCOMM_ER_RXPARITY ' Parity Error.
        Case MSCOMM_ER_TXFULL   ' Transmit buffer full.
    End Select
    
    Me.MousePointer = vbDefault

End Sub

Private Sub cmdAck_Click()

    If chkSendgbn.Value = vbChecked Then
        MSComm1.Output = Chr$(6)
    Else
        txtSend.Text = txtSend.Text + Chr$(6)
        txtSend.SetFocus
    End If
    
End Sub

Private Sub cmdEtx_Click()

    If chkSendgbn.Value = vbChecked Then
        MSComm1.Output = Chr$(3)
    Else
        txtSend.Text = txtSend.Text + Chr$(3)
        txtSend.SetFocus
    End If

End Sub

Private Sub cmdEot_Click()

    If chkSendgbn.Value = vbChecked Then
        MSComm1.Output = Chr$(4)
    Else
        txtSend.Text = txtSend.Text + Chr$(4)
        txtSend.SetFocus
    End If
   
End Sub


Private Sub cmdStx_Click()
    
    If chkSendgbn.Value = vbChecked Then
        MSComm1.Output = Chr$(2)
    Else
        txtSend.Text = txtSend.Text + Chr$(2)
        txtSend.SetFocus
    End If
    
End Sub

Private Sub cmdNak_Click()

    If chkSendgbn.Value = vbChecked Then
        MSComm1.Output = Chr$(21)
    Else
        txtSend.Text = txtSend.Text + Chr$(21)
        txtSend.SetFocus
    End If

End Sub

Private Sub tmrSample_Timer()

    If f_intIdx > f_intCnt Then Exit Sub
    
    MSComm1.Output = f_strDta(f_intIdx)
    tmrSample.Enabled = False
    
End Sub


