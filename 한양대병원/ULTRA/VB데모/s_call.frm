VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "string call"
   ClientHeight    =   4845
   ClientLeft      =   4035
   ClientTop       =   3090
   ClientWidth     =   6510
   LinkTopic       =   "Form2"
   ScaleHeight     =   4845
   ScaleWidth      =   6510
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   0
      TabIndex        =   7
      Top             =   4320
      Width           =   6495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   6
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   5
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CALL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   4
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   1560
      Width           =   3735
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
      Left            =   2040
      TabIndex        =   0
      Top             =   600
      Width           =   3735
   End
   Begin VB.Label Label2 
      Caption         =   "OUTPUT :"
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
      Left            =   480
      TabIndex        =   3
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "INPUT :"
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
      Left            =   480
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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

    
Text3.Text = "���񽺰� �������Դϴ�."

'**********************************************
' �޸� �Ҵ�
'**********************************************
strbuf.bufptr = tpalloc("STRING", "", 80)
        
If strbuf.bufptr = 0 Then
        TuxError ("TPALLOC ����. ������ȣ : ")
        Text3.Text = "TPALLOC ����."
        Exit Sub
End If
    
'    ret = AEWISBLOCKED()
'    If ret = 1 Then
'        MsgBox "AEWISBLOCKED ���� : Application Busy"
'        Exit Sub
'    End If

'********************************************
' transaction ����
'********************************************
'ret = tpbegin(30, 0)
'If ret = -1 Then
'        TuxError ("TPBEGIN ����. ������ȣ : ")
'        Text3.Text = "TPBEGIN ����."
'        Exit Sub
'End If
    
'********************************************
' ���� CALL
'********************************************
temp1$ = Space$(100)
temp1$ = Text1.Text
lret = lstrcpy(ByVal strbuf.bufptr, ByVal temp1$)
    
ret = tpcall("TOUPPER", ByVal strbuf.bufptr, ByVal 0&, strbuf, recvlen, ByVal 0&)
    
If ret = -1 Then
       err_ret = ErrorMsg(ByVal strbuf.bufptr&, "TPCALL", 0)
       ret = tpabort(0)
       ret = tpfree(strbuf.bufptr)
       Text3.Text = "TPCALL ����."
       Exit Sub
End If
    
'********************************************
' Transaction Commit
'********************************************
'ret = tpcommit(0)
'If ret = -1 Then
'        TuxError ("TPCOMMIT ����. ������ȣ : ")
'        Text3.Text = "TPCOMMIT ����."
'        Exit Sub
'End If
        
temp2$ = Space$(recvlen.bufptr)
lret = lstrcpy(ByVal temp2$, ByVal strbuf.bufptr)
Text2.Text = temp2$
Text3.Text = "���񽺰� �������� ����Ǿ����ϴ�."
ret = tpfree(strbuf.bufptr)

End Sub

Private Sub Command3_Click()
    Dim ret As Long
    Close
End Sub

