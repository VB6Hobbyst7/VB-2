VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmChat 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "채팅 서버 프로그램"
   ClientHeight    =   6660
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command1 
      Caption         =   "연결"
      Height          =   465
      Left            =   3060
      TabIndex        =   3
      Top             =   4920
      Width           =   1125
   End
   Begin VB.TextBox txtPort 
      Height          =   405
      Left            =   720
      TabIndex        =   2
      Top             =   4950
      Width           =   1965
   End
   Begin VB.TextBox txtChat 
      Height          =   735
      Left            =   0
      ScrollBars      =   2  '수직
      TabIndex        =   1
      Top             =   5880
      Width           =   4695
   End
   Begin MSWinsockLib.Winsock ctrSocket 
      Left            =   2040
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox rtfView 
      Height          =   4635
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   8176
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmChat.frx":0000
   End
   Begin VB.Label Label2 
      Caption         =   "Port"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   90
      TabIndex        =   4
      Top             =   4980
      Width           =   585
   End
   Begin VB.Menu mnuconnection 
      Caption         =   "접속"
      Begin VB.Menu mnuexit 
         Caption         =   "접속 종료"
      End
      Begin VB.Menu mnuseparation 
         Caption         =   "-"
      End
      Begin VB.Menu mnuquit 
         Caption         =   "프로그램 종료"
      End
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'클라이언트 코드
'구성요소 추카 1. rich text box control 6.0
'              2. winsock control 6.0
'rich text box의 속성 수정 'Locked'체크표시

Private Sub Command1_Click()
    
    ctrSocket.LocalPort = CInt(txtPort.Text) '5150
    ctrSocket.Listen

End Sub

Private Sub ctrsocket_connectionRequest(ByVal requestID As Long)
    If ctrSocket.State <> sckClosed Then
        ctrSocket.Close
        
        ctrSocket.Accept requestID
        MsgBox "클라이언트에 접속되었습니다."
    End If
End Sub

Private Sub ctrsocket_dataarrival(ByVal bytestotal As Long)
    Dim strText As String
    Dim strtmp As String
    
    ctrSocket.GetData strText
    strtmp = rtfView.Text + Chr(13) + Chr(10) + strText
    rtfView.Text = strtmp
End Sub

Private Sub ctrsocket_error(ByVal number As Integer, description As String, ByVal scode As Long, ByVal source As String, ByVal helpfile As String, ByVal helpcontext As Long, canceldisplay As Boolean)
    MsgBox description, vbOKOnly, "오류"
End Sub

Private Sub Form_Load()
'    ctrSocket.LocalPort = CInt(5150) ' Text1.Text '50003
'    ctrSocket.Listen
End Sub

Private Sub mnuexit_Click()
    ctrsocket_close
End Sub

Private Sub mnuquit_Click()
End
End Sub

Private Sub txtChat_keyPress(keyAscii As Integer)
    Dim strText As String
        
    If keyAscii = vbKeyReturn Then
        strText = "서버:" + txtChat.Text
        ctrSocket.SendData strText
        rtfView.Text = rtfView.Text + Chr(13) + Chr(10) + strText
        txtChat.Text = ""
    End If
End Sub

Private Sub ctrsocket_close()
    MsgBox "클라이언트와의 접속이 끊어졌습니다."
    ctrSocket.Close
End Sub
