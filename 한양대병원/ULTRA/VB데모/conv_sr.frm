VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Conersation_SR"
   ClientHeight    =   7050
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9375
   LinkTopic       =   "Form6"
   ScaleHeight     =   7050
   ScaleWidth      =   9375
   StartUpPosition =   3  'Windows ±âº»°ª
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      TabIndex        =   6
      Top             =   5520
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SEND"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      TabIndex        =   5
      Top             =   5520
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   6600
      Width           =   9375
   End
   Begin VB.ListBox List2 
      Height          =   4020
      Left            =   6000
      TabIndex        =   3
      Top             =   720
      Width           =   2535
   End
   Begin VB.Frame Frame2 
      Caption         =   " Receive Data "
      Height          =   4935
      Left            =   5760
      TabIndex        =   2
      Top             =   240
      Width           =   3015
   End
   Begin VB.ListBox List1 
      Height          =   4020
      Left            =   840
      TabIndex        =   1
      Top             =   720
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Caption         =   " Send Data "
      Height          =   4935
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "<===>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   7
      Top             =   2520
      Width           =   1575
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    
Dim strbuf As tuxbuf
Dim conv_handle As Integer

Private Sub Command1_Click()
    Dim lret As Long
    Dim Fbfrp_size As Long
    Dim revent As Long
    
    Dim ret As Long
    Dim tinput As String
    Dim tresult As String
    

    tinput = List1.Text
    
    ret = PUTSTR(ByVal strbuf.bufptr&, F_STRING, 0, tinput)
    
    ret = tpsend(conv_handle%, strbuf.bufptr&, 0&, TPRECVONLY, revent&)
    If ret = -1 Then
        If gettperrno() = TPEEVENT Then
            Select Case revent
                Case TPEV_DISCONIMM
                Case TPEV_SVCERR
                Case TPEV_SVCFAIL
                    ret = ErrorMsg(ByVal strbuf.bufptr&, "TPSEND", 0)
                    ret = tpdiscon(conv_handle)
                    Exit Sub
            End Select
        End If
    End If
                
    ret = tprecv(conv_handle%, strbuf, 0&, TPNOCHANGE, revent&)
    If ret = -1 Then
        If gettperrno() = TPEEVENT Then
            Select Case revent
                Case TPEV_DISCONIMM
                Case TPEV_SVCERR
                Case TPEV_SVCFAIL
                    ret = ErrorMsg(ByVal strbuf.bufptr&, "TPRECV", 0)
                    ret = tpdiscon(conv_handle)
                    Exit Sub
            End Select
        End If
    End If
    
    ret = GETSTR(ByVal strbuf.bufptr&, F_STRING, 0, tresult)
    
    List2.AddItem tresult
    

End Sub

Private Sub Command2_Click()
    
    Dim ret As Long
    
    ret = tpdiscon(conv_handle)
    Me.Cls

End Sub

Private Sub Form_Load()
    Dim lret As Long
    Dim Fbfrp_size As Long
    
    Dim ret As Long
    Dim entry As String
    
    
'------------------------------------------------------
' FML Memory Allocation
'------------------------------------------------------
    strbuf.bufptr = tpalloc("FML32", "", 1024)
            
    If strbuf.bufptr = 0 Then
        err_ret = ErrorMsg(ByVal strbuf.bufptr&, "Tpalloc", 0)
        ret = tpfree(strbuf.bufptr)
        Exit Sub
    End If
    
    Fbfrp_size = Fsizeof32(ByVal strbuf.bufptr&)

'------------------------------------------------------
' FML Buffer Initialize
'------------------------------------------------------
    ret = Finit32(ByVal strbuf.bufptr&, Fbfrp_size)
    If ret = -1 Then
        err_ret = ErrorMsg(ByVal strbuf.bufptr&, "Finit", 1)
        ret = tpfree(strbuf.bufptr)
        Exit Sub
    End If
    
'------------------------------------------------------
' Connection Initialize to Service
'------------------------------------------------------
    conv_handle = tpconnect("FCONV_SR", ByVal strbuf.bufptr&, 0, TPSENDONLY)
    If conv_handle = -1 Then
        err_ret = ErrorMsg(ByVal strbuf.bufptr&, "TPCONNECT", 0)
        ret = tpfree(strbuf.bufptr)
        Exit Sub
    End If
    
    For i = 1 To 100
        entry = "Entry String #" & i
        List1.AddItem entry
    Next i
        
End Sub
