VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "Conversational Recieve Only"
   ClientHeight    =   8415
   ClientLeft      =   4830
   ClientTop       =   1590
   ClientWidth     =   6645
   LinkTopic       =   "Form7"
   ScaleHeight     =   8415
   ScaleWidth      =   6645
   Begin VB.TextBox Text2 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1920
      TabIndex        =   4
      Top             =   6120
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   8040
      Width           =   6735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3720
      TabIndex        =   2
      Top             =   6960
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "START"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   720
      TabIndex        =   1
      Top             =   6960
      Width           =   2175
   End
   Begin VB.ListBox List1 
      Height          =   5460
      Left            =   1200
      TabIndex        =   0
      Top             =   360
      Width           =   4095
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    
Dim strbuf As tuxbuf
Dim conv_handle As Integer

Private Sub Command1_Click()
    
    Dim lret As Long
    Dim Fbfrp_size As Long
    
    Dim ret As Long
    Dim entry1 As String
    Dim entry2 As String
    Dim tresult As String
    Dim i As Integer
    Dim rlen As Long
    
'------------------------------------------------------
' FML Memory Allocation
'------------------------------------------------------
    strbuf.bufptr = tpalloc("FML32", "", 2048)
            
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
        err_ret = ErrorMsg(ByVal strbuf.bufptr&, "Finit32", 1)
        ret = tpfree(strbuf.bufptr)
        Exit Sub
    End If
    
'------------------------------------------------------
' Connection Initialize to Service
'------------------------------------------------------
    conv_handle = tpconnect("FCONV", ByVal strbuf.bufptr&, 0, TPRECVONLY)
    If conv_handle = -1 Then
        err_ret = ErrorMsg(ByVal strbuf.bufptr&, "TPCONNECT", 0)
        ret = tpfree(strbuf.bufptr)
        Exit Sub
    End If
    
    ret = tprecv(conv_handle%, strbuf, rlen&, TPNOCHANGE, revent&)
    If ret = -1 Then
        If gettperrno() = TPEEVENT Then
            Select Case revent
                Case TPEV_DISCONIMM
                Case TPEV_SVCERR
                Case TPEV_SVCFAIL
                    ret = ErrorMsg(ByVal strbuf.bufptr&, "TPRECV SVCFAIL", 0)
                    ret = tpdiscon(conv_handle)
                    ret = tpfree(strbuf.bufptr)
                    Exit Sub
            End Select
        End If
    End If
    
    ret = GETLONG(ByVal strbuf.bufptr&, F_LONG, 0, tresult)
    
    For i = 1 To CInt(tresult) Step 1
    
        ret = tprecv(conv_handle%, strbuf, rlen&, TPNOCHANGE, revent&)
        If ret = -1 Then
            If gettperrno() = TPEEVENT Then
                Select Case revent
                    Case TPEV_DISCONIMM
                    Case TPEV_SVCERR
                    Case TPEV_SVCFAIL
                        ret = ErrorMsg(ByVal strbuf.bufptr&, "TPRECVL SVCFAIL", 0)
                        ret = tpdiscon(conv_handle)
                        ret = tpfree(strbuf.bufptr)
                        Exit Sub
                End Select
            End If
        End If
        
        If revent = TPEV_SVCSUCC Then
            Exit For
        End If

        ret = GETLONG(ByVal strbuf.bufptr&, F_LONG, 0, entry1)
        ret = GETSTR(ByVal strbuf.bufptr&, F_STRING, 0, entry2)
        List1.AddItem entry2
        Text2.Text = entry1 + " / " + tresult
    
    Next i
    
    ret = tpfree(strbuf.bufptr)
    
End Sub

