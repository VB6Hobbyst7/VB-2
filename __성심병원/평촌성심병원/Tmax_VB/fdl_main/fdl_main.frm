VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8025
   LinkTopic       =   "Form1"
   ScaleHeight     =   7545
   ScaleWidth      =   8025
   StartUpPosition =   3  'Windows ±âº»°ª
   Begin VB.Frame Frame2 
      Caption         =   "°á°ú"
      Height          =   5295
      Left            =   240
      TabIndex        =   4
      Top             =   1920
      Width           =   7455
      Begin VB.TextBox txtDeptno 
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   21
         Top             =   4560
         Width           =   2775
      End
      Begin VB.TextBox txtDate 
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   20
         Top             =   2760
         Width           =   2775
      End
      Begin VB.TextBox txtSal 
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   19
         Top             =   3360
         Width           =   2775
      End
      Begin VB.TextBox txtComm 
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   18
         Top             =   3960
         Width           =   2775
      End
      Begin VB.TextBox txtMgr 
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   17
         Top             =   2160
         Width           =   2775
      End
      Begin VB.TextBox txtJob 
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   9
         Top             =   1560
         Width           =   2775
      End
      Begin VB.TextBox txtEname 
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   7
         Top             =   960
         Width           =   2775
      End
      Begin VB.TextBox txtEmpno 
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   6
         Top             =   360
         Width           =   2775
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "reset"
         Height          =   495
         Left            =   5400
         TabIndex        =   5
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblDeptno 
         Caption         =   "deptno"
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   16
         Top             =   4680
         Width           =   855
      End
      Begin VB.Label lblComm 
         Caption         =   "comm"
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   720
         TabIndex        =   15
         Top             =   4080
         Width           =   855
      End
      Begin VB.Label lblSal 
         Caption         =   "sal"
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   720
         TabIndex        =   14
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label lblDate 
         Caption         =   "date"
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   13
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label lblMgr 
         Caption         =   "mgr"
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   12
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label lblJob 
         Caption         =   "job"
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   11
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label lblEname 
         Caption         =   "ename"
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   10
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblEmpno 
         Caption         =   "empno"
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   720
         TabIndex        =   8
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Á¶È¸"
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7455
      Begin VB.TextBox txtSalEmpno 
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   2
         Text            =   "7788"
         Top             =   360
         Width           =   2775
      End
      Begin VB.CommandButton cmdTpcall 
         Caption         =   "tpcall"
         Height          =   495
         Left            =   5280
         TabIndex        =   1
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lblSelEmpno 
         Caption         =   "empno"
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   3
         Top             =   480
         Width           =   855
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private initcheck As Boolean

Private Sub cmdReset_Click()
    txtEmpno.text = ""
    txtEname.text = ""
    txtSal.text = ""
    txtComm.text = ""
    txtDate.text = ""
    txtDeptno.text = ""
    txtJob.text = ""
    txtMgr.text = ""
    
End Sub

Private Sub cmdTpcall_Click()
    Dim lsndbuf As Long
    Dim lrcvbuf As Long
    Dim lrbuflen As Long
    Dim lret As Long
    Dim text As String
    Dim sal_empno As Long
    Dim empno As Long
    Dim ename As String
    Dim job As String
    Dim mgr As Long
    Dim date1 As String
    Dim sal As Single
    Dim comm As Single
    Dim deptno As Long
                                                 
    If initcheck = False Then
        MsgBox "TMAX¿¬°áÀÌ ½ÇÆÐÇÑ »óÅÂÀÔ´Ï´Ù. ÁøÇàÇÒ ¼ö ¾ø½À´Ï´Ù"
        Exit Sub
    End If
    
    lsndbuf = tpalloc("FIELD", "", 1024)
    lrcvbuf = tpalloc("FIELD", "", 1024)
    
    If lsndbuf = 0 Or lrcvbuf = 0 Then
        MsgBox "buffer allocation error" & gettperrno()
        Exit Sub
    End If
    
    sal_empno = CLng(txtSalEmpno.text)
    
    If fbput(ByVal lsndbuf, ByVal fbget_fldkey("EMPNO"), sal_empno, 0) = -1 Then
       MsgBox "fbput error"
       GoTo memoryfree
    End If
    
    
    'or---------------------------------
    'sal_empno = CLng(txtSalEmpno.text)
    'If PUTLONG(ByVal lsndbuf, "EMPNO", 0, sal_empno) = -1 Then
    '    MsgBox "puntint error"
    '    GoTo memoryfree
    'End If
    '-----------------------------------
    
    If tpcall(ByVal "FDLSEL", ByVal lsndbuf, ByVal 0, lrcvbuf, lrbuflen, ByVal 0) = -1 Then
        MsgBox "tpcall error " & gettperrno()
        GoTo memoryfree
    End If

    MsgBox fbkeyoccur(lrcvbuf, fbget_fldkey("ENAME"))
    
    txtEmpno.text = sal_empno
    
    ename = String$(1024, Chr$(0))
    If fbget(ByVal lrcvbuf&, ByVal fbget_fldkey("ENAME"), ByVal ename, 0) = -1 Then
        MsgBox "ename_fbget error"
        GoTo memoryfree
    End If
    txtEname.text = ename
    
    'or---------------------
    'If GETVAR(ByVal lrcvbuf, "ENAME", 0, ename) = -1 Then
    '    MsgBox "getcar erroro" & gettperrno()
    '    GoTo memoryfree
    'End If
    'txtEname.text = ename
    '-----------------------
    
    job = String$(1024, Chr$(0))
    If fbget(ByVal lrcvbuf&, ByVal fbget_fldkey("JOB"), ByVal job, 0) = -1 Then
        MsgBox "job_fbget error"
        GoTo memoryfree
    End If
    txtJob.text = job
    
    date1 = String$(1024, Chr$(0))
    If fbget(ByVal lrcvbuf&, ByVal fbget_fldkey("DATE"), ByVal date1, 0) = -1 Then
        MsgBox "date_fbget error"
        GoTo memoryfree
    End If
    txtDate.text = date1
    
    If fbget(ByVal lrcvbuf&, ByVal fbget_fldkey("SAL"), sal, 0) = -1 Then
        MsgBox "sal_fbget error"
        GoTo memoryfree
    End If
    txtSal.text = sal
    
    If fbget(ByVal lrcvbuf&, ByVal fbget_fldkey("COMM"), comm, 0) = -1 Then
        MsgBox "comm_fbget error"
        GoTo memoryfree
    End If
    txtComm.text = comm
    
    If fbget(ByVal lrcvbuf&, ByVal fbget_fldkey("MGR"), mgr, 0) = -1 Then
        MsgBox "mgr_fbget error"
        GoTo memoryfree
    End If
    txtMgr.text = mgr
    
    If fbget(ByVal lrcvbuf&, ByVal fbget_fldkey("DEPTNO"), deptno, 0) = -1 Then
        MsgBox "deptno_fbget error"
        GoTo memoryfree
    End If
    txtDeptno.text = deptno
    
    
    'or -------------------------
    
    'If GETLONG(ByVal lrcvbuf, "MGR", 0, mgr) = -1 Then
    '    MsgBox "mgr_getlong erroro" & gettperrno()
    '    GoTo memoryfree
    'End If
    'txtMgr.text = CStr(mgr)
    
    
    'If GETLONG(ByVal lrcvbuf, "DEPTNO", 0, deptno) = -1 Then
    '    MsgBox "deptno_getlong erroro" & gettperrno()
    '    GoTo memoryfree
    'End If
    'txtDeptno.text = CStr(deptno)
    
    '-----------------------------
    
memoryfree:
    Call tpfree(ByVal lsndbuf)
    Call tpfree(ByVal lrcvbuf)
    
End Sub

Private Sub Form_Load()
                         
    Call tmaxinit
    
End Sub


Private Sub tmaxinit()
    Dim ret As Integer
    Dim lsndbuf As Long
    Dim tpinfo As tpstart_t
    
    ' (1) tmaxreadenv()¸¦ »ç¿ëÇÏ¿© È¯°æÈ­ÀÏ·ÎºÎÅÍ È¯°æº¯¼ö¸¦ ÀÐ¾î¿È
    If tmaxreadenv(ByVal "..\env\tmax.env", "TMAX") = -1 Then
        MsgBox ("read env file error") & "," & gettperrno()
        Exit Sub
    End If
        
    lsndbuf = tpalloc("TPSTART", "", 0)
    If lsndbuf = 0 Then
        MsgBox "send buffer allocation error" & gettperrno()
        Exit Sub
    End If
       
    tpinfo.cltname = "Tmax" + Chr(0)
    tpinfo.dompwd = "tmax123" + Chr(0)
    tpinfo.usrname = "Tmax" + Chr(0)
    tpinfo.usrpwd = "" + Chr(0)
    
    ret = FilltpstartBuf(lsndbuf, tpinfo)
    
    ret = tpstart(ByVal lsndbuf)
    If ret = -1 Then
        MsgBox ("tp start error") & "," & gettperrno()
        initcheck = False
        End
    End If
    
    Call tpfree(ByVal lsndbuf)
    
    initcheck = True

    
End Sub

Private Function tmaxexit() As Boolean
            
    If tpend() = -1 Then
        MsgBox ("tp end error") & "," & gettperrno()
        tmaxexit = False
        Exit Function
    End If

    tmaxexit = True
    
End Function

Private Sub Form_Unload(Cancel As Integer)

    Call tmaxexit
        
End Sub

