VERSION 5.00
Begin VB.MDIForm mdiMain 
   BackColor       =   &H8000000C&
   Caption         =   "임상병리통계Report"
   ClientHeight    =   5745
   ClientLeft      =   1545
   ClientTop       =   2610
   ClientWidth     =   9795
   Icon            =   "mdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  '최대화
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
   Begin VB.Menu mnuJob 
      Caption         =   "Job"
      Begin VB.Menu mnuYear 
         Caption         =   "검사건수(년통계)"
      End
      Begin VB.Menu mnuMonth 
         Caption         =   "검사건수(월통계)"
      End
      Begin VB.Menu mnuItemDept 
         Caption         =   "검사항목(과별) 통계"
      End
      Begin VB.Menu mnuBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTime 
         Caption         =   "항목,시간대별 통계"
      End
      Begin VB.Menu mnuDept 
         Caption         =   "진료과별"
      End
      Begin VB.Menu mnuEx 
         Caption         =   "외부검사통계"
      End
      Begin VB.Menu mnuEr 
         Caption         =   "응급검사통계"
      End
      Begin VB.Menu mnuWeek 
         Caption         =   "요일별 통계"
      End
   End
   Begin VB.Menu mnuPart 
      Caption         =   "Part통계"
      Begin VB.Menu mnuPart1 
         Caption         =   "일일통계"
      End
      Begin VB.Menu mnuPartWeek 
         Caption         =   "주간통계"
      End
      Begin VB.Menu mnuPartMonth 
         Caption         =   "월간통계"
      End
      Begin VB.Menu mnuPartYear 
         Caption         =   "년간통계"
      End
   End
   Begin VB.Menu mnuCal 
      Caption         =   "달력통계"
   End
   Begin VB.Menu mnuWhonet 
      Caption         =   "WhoNet"
      Begin VB.Menu mnuWhonet0 
         Caption         =   "일자별 균주 백분율(Gram-neg)"
      End
      Begin VB.Menu mnuWhonet1 
         Caption         =   "일자별 균주 백분율(sau & efa)"
      End
      Begin VB.Menu mnuWhonet2 
         Caption         =   "항균제별 감수성 백분율"
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()
    Call adoDbConnect("TW_MIS_EXAM", "HOSPITAL", "V2MTS")
    
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call adoDbDisconnect
    
End Sub

Private Sub mnuCal_Click()
    frmCalData.Show
    
End Sub

Private Sub mnuDept_Click()
    frmDept.Show
    
End Sub

Private Sub mnuEr_Click()
    frmEr.Show
    
End Sub

Private Sub mnuEx_Click()
    frmEx.Show
    
End Sub

Private Sub mnuExit_Click()
    If vbYes = MsgBox("프로그램을 종료하시겠습니까?", vbYesNo + vbQuestion, "종료확인Box") Then
        End
    End If
    
End Sub

Private Sub mnuItemDept_Click()
    frmTongitem.Show
    
End Sub

Private Sub mnuMonth_Click()
    frmTongMonth.Show
    
End Sub

Private Sub mnuPart1_Click()
    frmPartDay.Show
    
End Sub

Private Sub mnuPartMonth_Click()
    frmPartMonth.Show
    
End Sub

Private Sub mnuPartWeek_Click()
    frmPartWeek.Show
    
End Sub

Private Sub mnuPartYear_Click()
    frmPartYear.Show
    
End Sub

Private Sub mnuRet_Click()

'    frmReport.Show
    
End Sub

Private Sub mnuTime_Click()
    frmTime.Show
    
End Sub

Private Sub mnuWeek_Click()
    frmWeek.Show
    
End Sub

Private Sub mnuWhonet0_Click()
    
    frmWhonet0.Show
    
End Sub

Private Sub mnuWhonet1_Click()
    frmWhonet2.Show
    
    
End Sub

Private Sub mnuWhonet2_Click()
    frmWhonet1.Show    '항균제별 Sens 백분율
    
End Sub

Private Sub mnuYear_Click()
    frmTongYear.Show
    
End Sub
