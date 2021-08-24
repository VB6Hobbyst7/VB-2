VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "FML32"
   ClientHeight    =   5490
   ClientLeft      =   3180
   ClientTop       =   1545
   ClientWidth     =   10080
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5490
   ScaleWidth      =   10080
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      TabIndex        =   19
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      TabIndex        =   18
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "-------->"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      TabIndex        =   13
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "-------->"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      TabIndex        =   12
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "-------->"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      TabIndex        =   11
      Top             =   1440
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H80000001&
      ForeColor       =   &H80000018&
      Height          =   495
      Left            =   0
      TabIndex        =   10
      Top             =   5040
      Width           =   10095
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1800
      TabIndex        =   9
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1800
      TabIndex        =   8
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1800
      TabIndex        =   7
      Top             =   1440
      Width           =   3015
   End
   Begin VB.CommandButton Command3 
      Caption         =   "-------->"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      TabIndex        =   3
      Top             =   600
      Width           =   1695
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
      Height          =   615
      Left            =   1800
      TabIndex        =   2
      Top             =   600
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Closed"
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
      Left            =   6720
      TabIndex        =   0
      Top             =   4080
      Width           =   3015
   End
   Begin VB.Label Label10 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   21
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label6 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   20
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Label9 
      Caption         =   "DOUBLE:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   17
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label Label8 
      Caption         =   "LONG    :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   16
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "CARRAY:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "STRING :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   720
      Width           =   1575
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
      Height          =   615
      Left            =   6720
      TabIndex        =   6
      Top             =   3060
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
      Height          =   615
      Left            =   6720
      TabIndex        =   5
      Top             =   2280
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
      Height          =   615
      Left            =   6720
      TabIndex        =   4
      Top             =   1440
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
      Height          =   615
      Left            =   6720
      TabIndex        =   1
      Top             =   600
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command1_Click()
    
    Dim strbuf As tuxbuf
    Dim recvptr As tuxbuf
   
    Dim lret As Long
    Dim Fbfrp_size As Long

' Buffer
    
    Dim ret As Long
    Dim tinput As String
    Dim tresult1 As String
    
    
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
' FSecurity Service를 요구
'------------------------------------------------------
    tinput = Text2.Text
    
    ret = PUTCARR(ByVal strbuf.bufptr&, F_CARRAY, 0, tinput, Len(tinput))
    
    ret = tpcall("CARRAYS", ByVal strbuf.bufptr&, 0&, strbuf, 0&, 0&)
    If ret = -1 Then
        err_ret = ErrorMsg(ByVal strbuf.bufptr&, "TPCALL", 0)
        ret = tpfree(strbuf.bufptr)
        Exit Sub
    End If
    
    ret = GETCARR(ByVal strbuf.bufptr&, F_CARRAY, 0, tresult1, Len(tinput))
    Label2.Caption = tresult1
    
    ret = tpfree(strbuf.bufptr)

End Sub

Private Sub Command2_Click()
    Me.Cls
End Sub

Private Sub Command3_Click()
    
    Dim strbuf As tuxbuf
    Dim recvptr As tuxbuf
   
    Dim lret As Long
    Dim Fbfrp_size As Long

' Buffer
    
    Dim ret As Long
    Dim tinput As String
    Dim tresult1 As String
    Dim tresult2 As String
    Dim tresult3 As String
    Dim tresult4 As String
    
    
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
' FSecurity Service를 요구
'------------------------------------------------------
    tinput = Text1.Text
    
    ret = PUTSTR(ByVal strbuf.bufptr&, F_STRING, 0, tinput)
    
    ret = tpcall("FTOUPPER", ByVal strbuf.bufptr&, 0&, strbuf, 0&, 0&)
    If ret = -1 Then
        err_ret = ErrorMsg(ByVal strbuf.bufptr&, "TPCALL", 0)
        ret = tpfree(strbuf.bufptr)
        Exit Sub
    End If
    
    ret = GETSTR(ByVal strbuf.bufptr&, F_STRING, 0, tresult1)
    Label1.Caption = tresult1
    
    ret = tpfree(strbuf.bufptr)

End Sub


Private Sub Command4_Click()
    
    Dim strbuf As tuxbuf
    Dim recvptr As tuxbuf
   
    Dim lret As Long
    Dim Fbfrp_size As Long

' Buffer
    
    Dim ret As Long
    Dim tinput1 As String
    Dim tinput2 As String
    Dim tresult1 As String
        
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
' FSecurity Service를 요구
'------------------------------------------------------
    tinput1 = Text3.Text
    ret = PUTLONG(ByVal strbuf.bufptr&, F_LONG, 0, tinput1)
    
    tinput2 = Text6.Text
    ret = PUTLONG(ByVal strbuf.bufptr&, F_LONG, 1, tinput2)
    
    ret = tpcall("LONGS", ByVal strbuf.bufptr&, 0&, strbuf, 0&, 0&)
    If ret = -1 Then
        err_ret = ErrorMsg(ByVal strbuf.bufptr&, "TPCALL", 0)
        ret = tpfree(strbuf.bufptr)
        Exit Sub
    End If
    
    ret = GETLONG(ByVal strbuf.bufptr&, F_LONG, 0, tresult1)
    Label3.Caption = tresult1
    
    ret = tpfree(strbuf.bufptr)

End Sub

Private Sub Command5_Click()
    
    Dim strbuf As tuxbuf
    Dim recvptr As tuxbuf
   
    Dim lret As Long
    Dim Fbfrp_size As Long

' Buffer
    
    Dim ret As Long
    Dim tinput1 As String
    Dim tinput2 As String
    Dim tresult1 As String
        
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
' FSecurity Service를 요구
'------------------------------------------------------
    tinput1 = Text4.Text
    ret = PUTDBL(ByVal strbuf.bufptr&, F_DOUBLE, 0, tinput1)
    
    tinput2 = Text7.Text
    ret = PUTDBL(ByVal strbuf.bufptr&, F_DOUBLE, 1, tinput2)
    
    ret = tpcall("DOUBLES", ByVal strbuf.bufptr&, 0&, strbuf, 0&, 0&)
    If ret = -1 Then
        err_ret = ErrorMsg(ByVal strbuf.bufptr&, "TPCALL", 0)
        ret = tpfree(strbuf.bufptr)
        Exit Sub
    End If
    
    ret = GETDBL(ByVal strbuf.bufptr&, F_DOUBLE, 0, tresult1)
    Label4.Caption = tresult1
    
    ret = tpfree(strbuf.bufptr)

End Sub
