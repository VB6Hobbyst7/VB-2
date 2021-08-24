VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   7110
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10215
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnu_file 
      Caption         =   "&File"
      Begin VB.Menu mnu_exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnu_buf 
      Caption         =   "&Buffer_Type"
      Begin VB.Menu mnu_str 
         Caption         =   "Sting"
      End
      Begin VB.Menu mnu_fml 
         Caption         =   "FML"
      End
   End
   Begin VB.Menu mnu_com 
      Caption         =   "Com_Type"
      Begin VB.Menu mnu_sync 
         Caption         =   "Sync"
      End
      Begin VB.Menu mnu_async 
         Caption         =   "Async"
      End
      Begin VB.Menu mnu_conv 
         Caption         =   "Conversation"
      End
      Begin VB.Menu mnu_convr 
         Caption         =   "Conversation_r"
      End
      Begin VB.Menu mnu_broad 
         Caption         =   "Broadcasting"
      End
   End
   Begin VB.Menu mnu_help 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
    
Dim ret As Long
Dim tp_err_no As Integer
Dim recvlen As tuxbuf
Dim recvptr As tuxbuf
Dim lret As Long
Dim slen As Long
Dim initbuf As tuxbuf
Dim errptr As Long
Dim ErrMsg As String
    
Dim tpinFop As tpinfobuf

    Form3.Cls

    '**********************************************
    ' 메모리 할당
    '**********************************************
'    initbuf.bufptr = tpalloc("TPINIT", "", 136 + 8)
'
'    If initbuf.bufptr = 0 Then
'        tp_err_no = gettperrno()
'        MsgBox "INIT Buffer TPALLOC 실패. 에러번호 : " + Str$(tp_err_no)
'        Exit Sub
'    End If
'
'    tpinFop.usrname = Space$(32)
'    tpinFop.cltname = Space$(32)
'    tpinFop.passwd = Space$(32)
'    tpinFop.grpname = Space$(32)
'    tpinFop.flags = 0
'    tpinFop.datalen = 0
'    tpinFop.data = 0
'
'
'    tpinFop.usrname = "joony"
'    tpinFop.cltname = "batman"
'
'    Call FillTpinitBuF(initbuf.bufptr&, tpinFop, "password")
'
'    If tpinit(ByVal initbuf.bufptr&) = -1 Then
'            tp_err_no = gettperrno()
'            MsgBox "TPINIT 실패. 에러번호 : " + Str$(tp_err_no)
'        End
'    Else
'          MsgBox "TPINIT  성공"
'    End If

' ATMI Function 사용
 If tpinit(ByVal 0&) = -1 Then
        tp_err_no = gettperrno()
        errptr = tpstrerror(tp_err_no)
        ret = lstrcpy(ByVal ErrMsg$, ByVal errptr&)
        MsgBox "TPINIT 실패. 에러번호 : " + Str$(tp_err_no) + ErrMsg
  End If
    

End Sub

Private Sub mnu_async_Click()
    Form5.Show
End Sub

Private Sub mnu_broad_Click()
    Form8.Show
End Sub

Private Sub mnu_conv_Click()
    Form6.Show
End Sub

Private Sub mnu_convr_Click()
    Form7.Show
End Sub

Private Sub mnu_exit_Click()
    End
End Sub

Private Sub mnu_fml_Click()
    Form1.Show
End Sub

Private Sub mnu_str_Click()
    Form2.Show
End Sub

Private Sub mnu_sync_Click()
    Form4.Show
End Sub
