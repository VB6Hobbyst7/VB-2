VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5775
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18810
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   18810
   StartUpPosition =   3  'Windows 기본값
   Begin VB.TextBox Text5 
      Height          =   405
      Left            =   9030
      TabIndex        =   8
      Text            =   "5150"
      Top             =   270
      Width           =   1485
   End
   Begin VB.CommandButton Command3 
      Caption         =   "연결"
      Height          =   615
      Left            =   11340
      TabIndex        =   7
      Top             =   270
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Height          =   555
      Left            =   390
      TabIndex        =   6
      Top             =   3480
      Width           =   5115
   End
   Begin VB.CommandButton Command2 
      Caption         =   "연결"
      Height          =   615
      Left            =   5550
      TabIndex        =   5
      Top             =   150
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Left            =   5250
      Top             =   3360
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   5940
      Top             =   3360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text3 
      Height          =   405
      Left            =   3930
      TabIndex        =   3
      Text            =   "1504"
      Top             =   150
      Width           =   1485
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   360
      TabIndex        =   2
      Text            =   "192.168.0.8"
      Top             =   150
      Width           =   3315
   End
   Begin VB.CommandButton Command1 
      Caption         =   "전송"
      Height          =   2205
      Left            =   5490
      TabIndex        =   1
      Top             =   1110
      Width           =   1485
   End
   Begin VB.TextBox Text1 
      Height          =   2205
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   1110
      Width           =   5055
   End
   Begin VB.Label lblStatus 
      Height          =   435
      Left            =   390
      TabIndex        =   4
      Top             =   630
      Width           =   5025
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
    Call Winsock1.SendData(Text1.Text & vbCr)

End Sub


Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)

    Winsock1.Close
    Winsock1.Listen
End Sub





Private Sub Command2_Click()
    Winsock1.Close
    
    Winsock1.Protocol = sckTCPProtocol
    Winsock1.RemoteHost = Text2.Text    '192.168.11.127
    Winsock1.RemotePort = Text3.Text    '192.168.11.127
    
'    Winsock1.LocalPort = Text3.Text
    Winsock1.Connect
    
    If Winsock1.State = sckConnected Then
        lblStatus.Caption = Winsock1.RemoteHost & ":" & Winsock1.RemotePort & " 에 연결되었습니다"
    ElseIf Winsock1.State = sckConnecting Then
        lblStatus.Caption = Winsock1.RemoteHost & ":" & Winsock1.RemotePort & " 에 연결중.."
    End If
    
    Timer1.Interval = 1000
    Timer1.Enabled = True

End Sub

Private Sub Command3_Click()
    Winsock1.Close

    Winsock1.LocalPort = Text3.Text
    Winsock1.Listen
    
    If Winsock1.State = sckConnected Then
        lblStatus.Caption = Winsock1.RemoteHost & ":" & Winsock1.RemotePort & " 에 연결되었습니다"
    ElseIf Winsock1.State = sckConnecting Then
        lblStatus.Caption = Winsock1.RemoteHost & ":" & Winsock1.RemotePort & " 에 연결중.."
    End If
    
    Timer1.Interval = 1000
    Timer1.Enabled = True

End Sub

Private Sub Timer1_Timer()
    
    If Winsock1.State = sckConnected Then
        lblStatus.Caption = Winsock1.RemoteHost & ":" & Winsock1.RemotePort & " 에 연결되었습니다"
        Timer1.Enabled = False
    ElseIf Winsock1.State = sckClosed Then
        lblStatus.Caption = Winsock1.RemoteHost & ":" & Winsock1.RemotePort & " 에 연결중.."
        Winsock1.Protocol = sckTCPProtocol
        Winsock1.RemoteHost = Text2.Text
        Winsock1.RemotePort = Text3.Text
        Winsock1.Connect
    End If
    
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim strRcvData As String
    
    Winsock1.GetData strRcvData
    Text4.Text = strRcvData


    Dim strRcvBuffer As String
    Dim strSndBuffer As String
   

    
    Winsock1.GetData strRcvBuffer
    Debug.Print strRcvBuffer


    strSndBuffer = "ORDER"
    Winsock1.SendData (strSndBuffer)


End Sub



Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    lblStatus.Caption = Number & ":" & Description
    Winsock1.Close

End Sub
