VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows 기본값
   Begin VB.Menu mnu 
      Caption         =   "파일"
      Index           =   0
      Begin VB.Menu mnu_1 
         Caption         =   "종료"
         Index           =   0
      End
   End
   Begin VB.Menu mnu 
      Caption         =   "Communication"
      Index           =   1
      Begin VB.Menu mnu_2 
         Caption         =   "Job List"
         Index           =   0
      End
      Begin VB.Menu mnu_2 
         Caption         =   "Result"
         Index           =   1
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub mnu_1_Click(Index As Integer)

    Select Case intdx
        Case 0: Unload Me
    End Select
    
End Sub

Private Sub mnu_2_Click(Index As Integer)

    Select Case Index
        Case 0: frmJobList.Show
        Case 1: frmReult.Show
    End Select
    
End Sub


Private Sub mnu_Click(Index As Integer)

End Sub


