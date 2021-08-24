VERSION 5.00
Begin VB.MDIForm FSMMain 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows 기본값
   Begin VB.Menu mnuB00 
      Caption         =   "☞ 기초코드"
      Begin VB.Menu mnuB 
         Caption         =   "▷ SLIP"
         Index           =   1
      End
      Begin VB.Menu mnuB 
         Caption         =   "▷ SPECIMEN"
         Index           =   2
      End
      Begin VB.Menu mnuB 
         Caption         =   "▷ TESTITEM"
         Index           =   3
      End
      Begin VB.Menu mnuB 
         Caption         =   "▷ ROUTINE"
         Index           =   4
      End
      Begin VB.Menu mnuB 
         Caption         =   "▷ DEPT"
         Index           =   5
      End
      Begin VB.Menu mnuB 
         Caption         =   "▷ USER"
         Index           =   6
      End
      Begin VB.Menu mnuB 
         Caption         =   "▷ COMMENT"
         Index           =   7
      End
      Begin VB.Menu mnuB 
         Caption         =   "▷ MACHINE"
         Index           =   8
      End
      Begin VB.Menu mnuB 
         Caption         =   "▷ CONFIG"
         Index           =   9
      End
   End
   Begin VB.Menu mnuJR00 
      Caption         =   "☞ 샘플 접수와 결과"
      Begin VB.Menu mnuJ 
         Caption         =   "▶ 샘플 접수"
         Index           =   1
      End
      Begin VB.Menu mnuR 
         Caption         =   "▶ 샘플별 결과등록"
         Index           =   1
      End
   End
   Begin VB.Menu mnuO00 
      Caption         =   "☞ 자료 출력"
      Begin VB.Menu mnuO 
         Caption         =   "▶ 검사보고서 출력"
         Index           =   1
      End
      Begin VB.Menu mnuO 
         Caption         =   "▶ 결과대장 출력"
         Index           =   2
      End
      Begin VB.Menu mnuO 
         Caption         =   "▶ WorkSheet 출력"
         Index           =   3
      End
   End
   Begin VB.Menu mnuS00 
      Caption         =   "☞ 결과 조회"
      Begin VB.Menu mnuS 
         Caption         =   "▶ 날짜구간별 조회"
         Index           =   1
      End
      Begin VB.Menu mnuS 
         Caption         =   "▶ 환자 HISTORY"
         Index           =   2
      End
      Begin VB.Menu mnuS 
         Caption         =   "▶ 이상자 체크"
         Index           =   3
      End
      Begin VB.Menu mnuS 
         Caption         =   "▶ DELTA 체크"
         Index           =   4
      End
   End
   Begin VB.Menu mnuT00 
      Caption         =   "☞ 통계"
      Begin VB.Menu mnuT 
         Caption         =   "▶ 일월년 검사건수"
         Index           =   1
      End
   End
   Begin VB.Menu mnuI00 
      Caption         =   "☞ 인터페이스"
      Begin VB.Menu mnuI 
         Caption         =   "▷ Selectra II"
         Index           =   1
      End
      Begin VB.Menu mnuI 
         Caption         =   "▷ Miditron"
         Index           =   2
      End
      Begin VB.Menu mnuI 
         Caption         =   "▷ Genius"
         Index           =   3
      End
   End
   Begin VB.Menu mnuE00 
      Caption         =   "☞ 마치기"
      Begin VB.Menu mnuE01 
         Caption         =   "▶ 종  료"
      End
      Begin VB.Menu mnuE02 
         Caption         =   "▶ 사용자 재 로그인"
      End
   End
End
Attribute VB_Name = "FSMMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
