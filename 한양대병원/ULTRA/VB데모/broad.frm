VERSION 5.00
Begin VB.Form Form8 
   Caption         =   "Broadcasting"
   ClientHeight    =   2775
   ClientLeft      =   5235
   ClientTop       =   4575
   ClientWidth     =   6585
   LinkTopic       =   "Form8"
   ScaleHeight     =   2775
   ScaleWidth      =   6585
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   5880
      Top             =   1200
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   3
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Broadcast"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   2400
      Width           =   6735
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   6375
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim ret As Long
    Dim tp_err_no As Integer
    Dim recvlen As tuxbuf
    Dim recvptr As tuxbuf
    Dim lret As Long
    Dim slen As Long
    Dim errptr As Long
    Dim ErrMsg As String
    Dim temp1 As String
    Dim temp2 As String
    Dim err_ret As Integer

    
Text2.Text = "서비스가 수행중입니다."

'**********************************************
' 메모리 할당
'**********************************************
strbuf.bufptr = tpalloc("STRING", "", 100)
        
If strbuf.bufptr = 0 Then
        TuxError ("TPALLOC 실패. 에러번호 : ")
        Text2.Text = "TPALLOC 실패."
        Exit Sub
End If
    
'********************************************
' Message Broadcasting
'********************************************
temp1$ = Space$(100)
temp1$ = Text1.Text
lret = lstrcpy(ByVal strbuf.bufptr, ByVal temp1$)
    
ret = tpbroadcast("", "", "", ByVal strbuf.bufptr, ByVal 0&, TPNOBLOCK)
    
If ret = -1 Then
       err_ret = ErrorMsg(ByVal strbuf.bufptr&, "TPBROADCAST", 0)
       ret = tpabort(0)
       ret = tpfree(strbuf.bufptr)
       Text2.Text = "TPBROADCAST 실패."
       Exit Sub
End If
    
Text2.Text = "Message를 정상적으로 Broadcasting 하였습니다."
ret = tpfree(strbuf.bufptr)


End Sub

Private Sub Form_Load()
    
    Dim ret As Long
    Dim tp_err_no As Integer
    Dim initbuf As tuxbuf
    Dim errptr As Long
    Dim ErrMsg As String
    
    If SetPrivateUnsolMsg(1) = -1 Then
        tp_err_no = gettperrno()
        errptr = tpstrerror(tp_err_no)
        ret = lstrcpy(ByVal ErrMsg$, ByVal errptr&)
        MsgBox "SetUnSol 실패. 에러번호 : " + Str$(tp_err_no) + ErrMsg
    End If
    
End Sub

Private Sub Timer1_Timer()
    tpchkunsol
End Sub
