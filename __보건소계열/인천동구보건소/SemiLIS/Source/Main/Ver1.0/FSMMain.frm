VERSION 5.00
Begin VB.MDIForm FSMMain 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.Menu mnuB00 
      Caption         =   "�� �����ڵ�"
      Begin VB.Menu mnuB 
         Caption         =   "�� SLIP"
         Index           =   1
      End
      Begin VB.Menu mnuB 
         Caption         =   "�� SPECIMEN"
         Index           =   2
      End
      Begin VB.Menu mnuB 
         Caption         =   "�� TESTITEM"
         Index           =   3
      End
      Begin VB.Menu mnuB 
         Caption         =   "�� ROUTINE"
         Index           =   4
      End
      Begin VB.Menu mnuB 
         Caption         =   "�� DEPT"
         Index           =   5
      End
      Begin VB.Menu mnuB 
         Caption         =   "�� USER"
         Index           =   6
      End
      Begin VB.Menu mnuB 
         Caption         =   "�� COMMENT"
         Index           =   7
      End
      Begin VB.Menu mnuB 
         Caption         =   "�� MACHINE"
         Index           =   8
      End
      Begin VB.Menu mnuB 
         Caption         =   "�� CONFIG"
         Index           =   9
      End
   End
   Begin VB.Menu mnuJR00 
      Caption         =   "�� ���� ������ ���"
      Begin VB.Menu mnuJ 
         Caption         =   "�� ���� ����"
         Index           =   1
      End
      Begin VB.Menu mnuR 
         Caption         =   "�� ���ú� ������"
         Index           =   1
      End
   End
   Begin VB.Menu mnuO00 
      Caption         =   "�� �ڷ� ���"
      Begin VB.Menu mnuO 
         Caption         =   "�� �˻纸�� ���"
         Index           =   1
      End
      Begin VB.Menu mnuO 
         Caption         =   "�� ������� ���"
         Index           =   2
      End
      Begin VB.Menu mnuO 
         Caption         =   "�� WorkSheet ���"
         Index           =   3
      End
   End
   Begin VB.Menu mnuS00 
      Caption         =   "�� ��� ��ȸ"
      Begin VB.Menu mnuS 
         Caption         =   "�� ��¥������ ��ȸ"
         Index           =   1
      End
      Begin VB.Menu mnuS 
         Caption         =   "�� ȯ�� HISTORY"
         Index           =   2
      End
      Begin VB.Menu mnuS 
         Caption         =   "�� �̻��� üũ"
         Index           =   3
      End
      Begin VB.Menu mnuS 
         Caption         =   "�� DELTA üũ"
         Index           =   4
      End
   End
   Begin VB.Menu mnuT00 
      Caption         =   "�� ���"
      Begin VB.Menu mnuT 
         Caption         =   "�� �Ͽ��� �˻�Ǽ�"
         Index           =   1
      End
   End
   Begin VB.Menu mnuI00 
      Caption         =   "�� �������̽�"
      Begin VB.Menu mnuI 
         Caption         =   "�� Selectra II"
         Index           =   1
      End
      Begin VB.Menu mnuI 
         Caption         =   "�� Miditron"
         Index           =   2
      End
      Begin VB.Menu mnuI 
         Caption         =   "�� Genius"
         Index           =   3
      End
   End
   Begin VB.Menu mnuE00 
      Caption         =   "�� ��ġ��"
      Begin VB.Menu mnuE01 
         Caption         =   "�� ��  ��"
      End
      Begin VB.Menu mnuE02 
         Caption         =   "�� ����� �� �α���"
      End
   End
End
Attribute VB_Name = "FSMMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
