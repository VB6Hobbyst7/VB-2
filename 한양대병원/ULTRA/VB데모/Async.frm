VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Async Call"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10275
   LinkTopic       =   "Form5"
   ScaleHeight     =   6375
   ScaleWidth      =   10275
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text6 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   495
      Left            =   0
      TabIndex        =   17
      Top             =   5880
      Width           =   10335
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
      Left            =   5640
      TabIndex        =   6
      Top             =   4440
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "처리"
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
      Left            =   2160
      TabIndex        =   5
      Top             =   4440
      Width           =   2535
   End
   Begin VB.TextBox Text5 
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
      Left            =   1920
      TabIndex        =   4
      Top             =   3480
      Width           =   3375
   End
   Begin VB.TextBox Text4 
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
      Left            =   1920
      TabIndex        =   3
      Top             =   2760
      Width           =   3375
   End
   Begin VB.TextBox Text3 
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
      Left            =   1920
      TabIndex        =   2
      Top             =   2040
      Width           =   3375
   End
   Begin VB.TextBox Text2 
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
      Left            =   1920
      TabIndex        =   1
      Top             =   1320
      Width           =   3375
   End
   Begin VB.TextBox Text1 
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
      Left            =   1920
      TabIndex        =   0
      Top             =   600
      Width           =   3375
   End
   Begin VB.Label Label15 
      Caption         =   "ACALL 5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   22
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label14 
      Caption         =   "ACALL 4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   21
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label13 
      Caption         =   "ACALL 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   20
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label12 
      Caption         =   "ACALL 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   19
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label11 
      Caption         =   "ACALL 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   18
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label10 
      Caption         =   "===>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   5520
      TabIndex        =   16
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "===>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   5520
      TabIndex        =   15
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "===>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   5520
      TabIndex        =   14
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "===>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   5520
      TabIndex        =   13
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "===>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   5520
      TabIndex        =   12
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
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
      Left            =   6840
      TabIndex        =   11
      Top             =   3480
      Width           =   3015
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
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
      Left            =   6840
      TabIndex        =   10
      Top             =   2760
      Width           =   3015
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
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
      Left            =   6840
      TabIndex        =   9
      Top             =   2040
      Width           =   3015
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
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
      Left            =   6840
      TabIndex        =   8
      Top             =   1320
      Width           =   3015
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
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
      Left            =   6840
      TabIndex        =   7
      Top             =   600
      Width           =   3015
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
    Dim strbuf As tuxbuf
    Dim recvptr As tuxbuf
   
    Dim lret As Long
    Dim Fbfrp_size As Long
    Dim i As Integer
    Dim svcnm As String
    Dim ret_cd As Long
    Dim start_cd As Long

' Buffer
    
    Dim ret As Long
    Dim tinput As String
    Dim tresult As String
    
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
' Async Service를 요구
'------------------------------------------------------
    For i = 1 To 5 Step 1
        Select Case i
            Case 1
                tinput = Text1.Text
                svcnm = "FACALL1"
            Case 2
                tinput = Text2.Text
                svcnm = "FACALL2"
            Case 3
                tinput = Text3.Text
                svcnm = "FACALL3"
            Case 4
                tinput = Text4.Text
                svcnm = "FACALL4"
            Case 5
                tinput = Text5.Text
                svcnm = "FACALL5"
        End Select
        
        ret = Finit32(ByVal strbuf.bufptr&, Fbfrp_size)
        ret = PUTSTR(ByVal strbuf.bufptr&, F_STRING, 0, tinput)
        
        ret_cd = tpacall(svcnm, ByVal strbuf.bufptr&, 0&, 0&)
        If ret_cd < 0 Then
            err_ret = ErrorMsg(ByVal strbuf.bufptr&, "TPACALL", 0)
            ret = tpfree(strbuf.bufptr)
            Exit Sub
        End If
        If i = 1 Then
            start_cd = ret_cd
        End If
    Next i
    
'------------------------------------------------------
' Async Response 받음
'------------------------------------------------------
    For i = 1 To 5 Step 1
        
        ret = Finit32(ByVal strbuf.bufptr&, Fbfrp_size)
        
        ret = tpgetrply(ret_cd&, strbuf, 0&, TPGETANY)
        If ret < 0 Then
            err_ret = ErrorMsg(ByVal strbuf.bufptr&, "TPGETRPLY", 0)
            ret = tpfree(strbuf.bufptr)
            Exit Sub
        End If
        
        ret = GETSTR(ByVal strbuf.bufptr&, F_STRING, 0, tresult)
        Select Case ret_cd
            Case start_cd
                Label1.Caption = tresult
            Case start_cd + 1
                Label2.Caption = tresult
            Case start_cd + 2
                Label3.Caption = tresult
            Case start_cd + 3
                Label4.Caption = tresult
            Case start_cd + 4
                Label5.Caption = tresult
        End Select
        
    Next i
    
    ret = tpfree(strbuf.bufptr)

End Sub

