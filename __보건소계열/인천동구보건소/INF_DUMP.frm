VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmDump 
   Caption         =   "DumpTEST"
   ClientHeight    =   5460
   ClientLeft      =   2115
   ClientTop       =   2355
   ClientWidth     =   5970
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5460
   ScaleWidth      =   5970
   Begin VB.CheckBox chkSendgbn 
      Caption         =   "¡ÔΩ√∫∏≥ª±‚"
      Height          =   225
      Left            =   4365
      TabIndex        =   19
      Top             =   1125
      Value           =   1  '»Æ¿Œ
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
      TabIndex        =   16
      Top             =   900
      Width           =   4110
   End
   Begin VB.TextBox txtResult 
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
      ForeColor       =   &H00FFFF80&
      Height          =   2850
      Left            =   135
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   1260
      Width           =   4110
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
      Width           =   4110
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   4185
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      RThreshold      =   1
      SThreshold      =   1
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   915
      Left            =   4410
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   90
      Width           =   1200
      _Version        =   65536
      _ExtentX        =   2117
      _ExtentY        =   1614
      _StockProps     =   78
      Caption         =   "¡æ∑·ESC"
      ForeColor       =   16711680
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
      Picture         =   "INF_DUMP.frx":0000
   End
   Begin Threed.SSCommand cmdAck 
      Height          =   465
      Left            =   5085
      TabIndex        =   1
      Top             =   1485
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
      Picture         =   "INF_DUMP.frx":08DA
   End
   Begin Threed.SSCommand cmdEtx 
      Height          =   465
      Left            =   4500
      TabIndex        =   2
      Top             =   2025
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
      Picture         =   "INF_DUMP.frx":08F6
   End
   Begin Threed.SSCommand cmdEot 
      Height          =   465
      Left            =   4500
      TabIndex        =   3
      Top             =   2565
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
      Picture         =   "INF_DUMP.frx":0912
   End
   Begin Threed.SSCommand cmdStx 
      Height          =   465
      Left            =   4500
      TabIndex        =   4
      Top             =   1485
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
      Picture         =   "INF_DUMP.frx":092E
   End
   Begin Threed.SSCommand cmdNak 
      Height          =   465
      Left            =   5085
      TabIndex        =   5
      Top             =   2025
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
      Picture         =   "INF_DUMP.frx":094A
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
            Size            =   11.25
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
   Begin Threed.SSPanel SSPanel1 
      Height          =   690
      Left            =   135
      TabIndex        =   11
      Top             =   4545
      Width           =   4110
      _Version        =   65536
      _ExtentX        =   7250
      _ExtentY        =   1217
      _StockProps     =   15
      ForeColor       =   8454143
      BackColor       =   16761024
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
      Begin Threed.SSPanel SSPanel2 
         Height          =   375
         Left            =   180
         TabIndex        =   12
         Top             =   135
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "Sample ∞πºˆ"
         ForeColor       =   12640511
         BackColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±º∏≤"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         BorderWidth     =   1
      End
      Begin VB.Label lblCount 
         Alignment       =   2  '∞°øÓµ• ∏¬√„
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
         TabIndex        =   13
         Top             =   180
         Width           =   1815
      End
   End
   Begin Threed.SSCommand cmdEnq 
      Height          =   465
      Left            =   5085
      TabIndex        =   14
      Top             =   2565
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
      Picture         =   "INF_DUMP.frx":0966
   End
   Begin Threed.SSCommand cmdClear 
      Height          =   465
      Left            =   4545
      TabIndex        =   15
      Top             =   4815
      Width           =   1110
      _Version        =   65536
      _ExtentX        =   1958
      _ExtentY        =   820
      _StockProps     =   78
      Caption         =   "Clear"
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
      Picture         =   "INF_DUMP.frx":0982
   End
   Begin Threed.SSCommand cmdSend 
      Height          =   465
      Left            =   4545
      TabIndex        =   17
      Top             =   4275
      Width           =   1110
      _Version        =   65536
      _ExtentX        =   1958
      _ExtentY        =   820
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
      BevelWidth      =   3
      Picture         =   "INF_DUMP.frx":099E
   End
   Begin Threed.SSCommand cmdFS 
      Height          =   465
      Left            =   5085
      TabIndex        =   18
      Top             =   3105
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
      Picture         =   "INF_DUMP.frx":09BA
   End
   Begin Threed.SSCommand cmdCr 
      Height          =   465
      Left            =   4500
      TabIndex        =   20
      Top             =   3105
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
      Picture         =   "INF_DUMP.frx":09D6
   End
   Begin Threed.SSCommand cmdLf 
      Height          =   465
      Left            =   4500
      TabIndex        =   21
      Top             =   3600
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
      Picture         =   "INF_DUMP.frx":09F2
   End
End
Attribute VB_Name = "frmDump"
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

Dim RecvBuffer  As String
Dim Sample_Cnt As Integer


Sub Test()
    
    Dim rcvBuffer As String
    
    rcvBuffer = "01 FE"

    rcvBuffer = rcvBuffer & "19/07/2001"
    
    rcvBuffer = rcvBuffer & "000001"
    
    rcvBuffer = rcvBuffer & Space$(30)
    
    rcvBuffer = rcvBuffer & Space$(30)
    
    rcvBuffer = rcvBuffer & Space$(3)
    
    rcvBuffer = rcvBuffer & "1"
    
    rcvBuffer = rcvBuffer & "1:"
    
    rcvBuffer = rcvBuffer & Space$(46)
    
    rcvBuffer = rcvBuffer & ""
    
    MSComm1.Output = rcvBuffer

End Sub


Private Sub cmdClear_Click()

    txtSend.Text = ""
    txtResult.Text = ""
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

    MSComm1.Output = txtSend.Text
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyEscape:    Call cmdExit_Click
    End Select

End Sub

Private Sub Form_Load()

    Sample_Cnt = 0
    lblFilename.Caption = "dump.dat"
    
    Dim sPath   As String

    With MSComm1
        .CommPort = 1
        'º”µµ,∆‰∏Æ∆º,≈◊¿Ã≈∏∫Ò∆Æ,stop bit
        .Settings = "9600,n,8,1"
        .PortOpen = True
        .RTSEnable = True
        .RThreshold = 1
    End With

    sPath = App.Path
    Open sPath + "\" + "dump.log" For Append As #1
    Open sPath + "\" + "dump.dat" For Output As #2

'    Test

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If MSComm1.PortOpen Then MSComm1.PortOpen = False
 
    Close #1:   Close #2

End Sub

Private Sub MSComm1_OnComm()
    
    Me.MousePointer = vbHourglass

    Dim wkbuf As String
    Dim wkdat As String
    Dim ix1 As Integer, cnt As Integer

    Dim temp   As Variant
    
    Select Case MSComm1.CommEvent
        ' Events
        Case MSCOMM_EV_SEND     ' There are SThreshold number of
           
                                ' character in the transmit buffer.
        Case MSCOMM_EV_RECEIVE  ' Received RThreshold # of chars.
            wkbuf = MSComm1.Input
            
            Print #1, wkbuf;
            Print #2, wkbuf;
            
            txtRecve.Text = wkbuf
            txtResult.Text = txtResult.Text + wkbuf
            
            For ix1 = 1 To Len(wkbuf)
                wkdat = Mid$(wkbuf, ix1, 1)
'                Select Case Asc(wkdat)
'                    Case 3 '/* ETX ?
'                        'MSComm1.Output = Chr$(6)
'                        'wkdat = ""
'
'                End Select
            Next
            
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




Private Sub txtSend_GotFocus()

    txtSend.SelStart = Len(txtSend.Text)
    
End Sub


