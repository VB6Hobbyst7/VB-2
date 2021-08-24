VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form4 
   Caption         =   "Muti Row Sync Call"
   ClientHeight    =   6570
   ClientLeft      =   3165
   ClientTop       =   2415
   ClientWidth     =   9645
   LinkTopic       =   "Form4"
   ScaleHeight     =   6570
   ScaleWidth      =   9645
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4815
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   8493
      _Version        =   393216
      Cols            =   6
   End
   Begin VB.CommandButton Command3 
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
      Left            =   7440
      TabIndex        =   3
      Top             =   5520
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
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
      Left            =   5520
      TabIndex        =   2
      Top             =   5520
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Query"
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
      Left            =   3480
      TabIndex        =   1
      Top             =   5520
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   6240
      Width           =   9735
   End
End
Attribute VB_Name = "Form4"
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
    Dim i As Long
    Dim j As Long
    Dim index As Long
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
' Multi Row Query Service¸¦ ¿ä±¸
'------------------------------------------------------
    ret = tpcall("DATAWIN", ByVal strbuf.bufptr&, 0&, strbuf, 0&, 0&)
    If ret = -1 Then
        err_ret = ErrorMsg(ByVal strbuf.bufptr&, "TPCALL", 0)
        ret = tpfree(strbuf.bufptr)
        Exit Sub
    End If
    
    index = Foccur32(strbuf.bufptr&, F_NAME)
    
    MSFlexGrid1.Rows = index + 1
        
    For i = 0 To index - 1 Step 1
        
        MSFlexGrid1.Row = i + 1
    
        MSFlexGrid1.Col = 0
        MSFlexGrid1.Text = CStr(i + 1)
        
        ret = GETSTR(ByVal strbuf.bufptr&, F_NAME, i, tresult)
        MSFlexGrid1.Col = 1
        MSFlexGrid1.Text = tresult
        ret = GETSTR(ByVal strbuf.bufptr&, F_SEX, i, tresult)
        MSFlexGrid1.Col = 2
        MSFlexGrid1.Text = tresult
        ret = GETLONG(ByVal strbuf.bufptr&, F_AGE, i, tresult)
        MSFlexGrid1.Col = 3
        MSFlexGrid1.Text = tresult
        ret = GETDBL(ByVal strbuf.bufptr&, F_HEIGHT, i, tresult)
        MSFlexGrid1.Col = 4
        MSFlexGrid1.Text = tresult
        ret = GETSTR(ByVal strbuf.bufptr&, F_TELNO, i, tresult)
        MSFlexGrid1.Col = 5
        MSFlexGrid1.Text = tresult
    
    Next i
'    ret = F_FPRINT32(strbuf.bufptr, "D:\tuxedo\ulog\aa.log")
    ret = tpfree(strbuf.bufptr)

End Sub

Private Sub Command3_Click()
    Me.Cls
End Sub

Private Sub Form_Load()
    Dim i As Integer
        MSFlexGrid1.Row = 0
        MSFlexGrid1.Col = 1
        MSFlexGrid1.Text = "NAME"
        MSFlexGrid1.Col = 2
        MSFlexGrid1.Text = "SEX"
        MSFlexGrid1.Col = 3
        MSFlexGrid1.Text = "AGE"
        MSFlexGrid1.Col = 4
        MSFlexGrid1.Text = "HEIGHT"
        MSFlexGrid1.Col = 5
        MSFlexGrid1.Text = "TEL_NO"
End Sub
