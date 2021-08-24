VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmDump 
   Caption         =   "DumpTEST"
   ClientHeight    =   7665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11115
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7665
   ScaleWidth      =   11115
   StartUpPosition =   3  'Windows ±‚∫ª∞™
   Begin FPSpread.vaSpread spdTest 
      Height          =   6225
      Left            =   5760
      OleObjectBlob   =   "dumpData.frx":0000
      TabIndex        =   23
      Top             =   630
      Width           =   5010
   End
   Begin VB.CheckBox chkSendgbn 
      Caption         =   "¡ÔΩ√∫∏≥ª±‚"
      Height          =   225
      Left            =   4365
      TabIndex        =   19
      Top             =   1125
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
      Picture         =   "dumpData.frx":093C
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
      Picture         =   "dumpData.frx":1216
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
      Picture         =   "dumpData.frx":1232
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
      Picture         =   "dumpData.frx":124E
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
      Picture         =   "dumpData.frx":126A
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
      Picture         =   "dumpData.frx":1286
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
      Picture         =   "dumpData.frx":12A2
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
      Picture         =   "dumpData.frx":12BE
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
      Picture         =   "dumpData.frx":12DA
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
      Picture         =   "dumpData.frx":12F6
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
      Picture         =   "dumpData.frx":1312
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
      Picture         =   "dumpData.frx":132E
   End
   Begin Threed.SSCommand cmdCheckSum 
      Height          =   510
      Left            =   5670
      TabIndex        =   22
      Top             =   90
      Width           =   1740
      _Version        =   65536
      _ExtentX        =   3069
      _ExtentY        =   900
      _StockProps     =   78
      Caption         =   "Check Sum"
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
      Picture         =   "dumpData.frx":134A
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

Private transcfg        As commset

Private Sub f_subCheckSum(ByVal icol As Integer)

    Dim iChkSum As Integer
    Dim iIdx    As Integer
    Dim vTmp    As Variant
    Dim sBuffer As String
    
    sBuffer = Chr(2)
    With spdTest
        For iIdx = 2 To .MaxRows
            .GetText icol, iIdx, vTmp
            
            Select Case icol
                Case 2:     If iIdx > 15 And Trim$(vTmp) = "" Then Exit For
                Case 3, 4:  If iIdx = 3 Then Exit For
                Case 6:     If iIdx = 7 Then Exit For
            End Select
            sBuffer = sBuffer + Trim$(vTmp) + Chr(28)
        Next
    End With
    
    iChkSum = 0
    For iIdx = 2 To Len(sBuffer)
        iChkSum = iChkSum + (0 Xor Asc(Mid(sBuffer, iIdx, 1)))
    Next

    txtSend.Text = sBuffer + Right$(CStr(Hex(iChkSum)), 2) + Chr(3)

End Sub


Private Sub cmdCheckSum_Click()

    Dim iChkSum As Integer
    Dim iIdx    As Integer
    Dim vTmp    As Variant
    Dim sBuffer As String
    
    If txtSend.Text = "" Then
        sBuffer = Chr(2)
        With spdTest
            For iIdx = 2 To .MaxRows
        
                .GetText 2, iIdx, vTmp
                
                If iIdx > 15 And Trim$(vTmp) = "" Then Exit For
                
                sBuffer = sBuffer + Trim$(vTmp) + Chr(28)
                
            Next
        End With
        
        iChkSum = 0
        For iIdx = 2 To Len(sBuffer)
            iChkSum = iChkSum + (0 Xor Asc(Mid(sBuffer, iIdx, 1)))
            
            If iChkSum >= 255 Then Exit For
        Next
    
        txtSend.Text = sBuffer + Right$(CStr(Hex(iChkSum)), 2) + Chr(3)
    Else
        iChkSum = 0
        For iIdx = 1 To Len(txtSend.Text)
            iChkSum = iChkSum + (0 Xor Asc(Mid(txtSend.Text, iIdx, 1)))
        Next
    
        txtSend.Text = txtSend.Text + Right$(CStr(Hex(iChkSum)), 2) + Chr(3)
    End If
    
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

    MSComm1.Output = txtSend.Text + "7d" + Chr(3)
    
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

'/  ≈ÎΩ≈»Ø∞Ê º≥¡§ π◊ Open
    Set dbcomm = OpenDatabase(filename & commstr)
    Set tbcomm = dbcomm.OpenRecordset("cfgcomm")

    tbcomm.MoveFirst
        
    With transcfg
        .Port = tbcomm!Port
        .data_bit = tbcomm!data_bit
        .stop_bit = tbcomm!stop_bit
        .baud_rate = tbcomm!baud_rate
        .parity = tbcomm!parity
        .blocksize = tbcomm!blocksize
    End With
    
    tbcomm.Close
    dbcomm.Close
    
    With MSComm1
        .CommPort = transcfg.Port
        .Settings = transcfg.baud_rate & "," & transcfg.parity & "," & transcfg.data_bit & "," & transcfg.stop_bit
        .PortOpen = True
        .RTSEnable = True
        .RThreshold = 1
    End With

    sPath = App.Path
'    sPath = "d:\∫∏∞«º“Ω√Ω∫≈€\«¡∑Œ±◊∑•\interface\" + Trim$(Title)
    Open sPath + "\" + machinit + "dump.log" For Append As #1
    Open sPath + "\" + machinit + "dump.dat" For Output As #2

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If MSComm1.PortOpen Then MSComm1.PortOpen = False
 
    Close #1:   Close #2

End Sub

Private Sub MSComm1_OnComm()
    
    Screen.MousePointer = vbHourglass

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
                Select Case Asc(wkdat)
'                    Case 5 '/* ENQ ?
'                        MSComm1.Output = Chr$(6)
''                    Case 2 '/* STX ?
'                        wkdat = ""
                    Case 3 '/* ETX ?
                        MSComm1.Output = Chr$(6)
                        
'                        If InStr(wkdat, "P") > 0 Then
'                            Call cmdCheckSum_Click
                            Call cmdSend_Click
'                        End If
                        
                        wkdat = ""
'                    Case 6
'                        RecvBuffer = RecvBuffer + wkdat
'                        txtResult.Text = txtResult + wkdat
'
'                        temp = InputBox("º±≈√«œººø‰(1:No Result, 2:wait)")
'                        If temp = "1" Then
'                            Call f_subCheckSum(3)
'                            Call cmdSend_Click
'                        Else
'                            Call f_subCheckSum(4)
'                            Call cmdSend_Click
'                        End If
'                    Case Else

'                    Case 10
'                        txtResult.Text = RecvBuffer
'                        MSComm1.Output = Chr$(6)
'                    Case Else:  RecvBuffer = RecvBuffer + wkdat
                End Select
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
    
    Screen.MousePointer = vbDefault

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


Private Sub spdTest_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

    Call f_subCheckSum(Col)
    Call cmdSend_Click
    
End Sub


Private Sub txtSend_GotFocus()

    txtSend.SelStart = Len(txtSend.Text)
    
End Sub


