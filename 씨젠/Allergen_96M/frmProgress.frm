VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProgress 
   Appearance      =   0  '���
   BackColor       =   &H80000005&
   BorderStyle     =   0  '����
   Caption         =   "Form1"
   ClientHeight    =   465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8580
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   465
   ScaleWidth      =   8580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '������ ���
   Begin MSComctlLib.ProgressBar Xprog 
      Height          =   465
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8580
      _ExtentX        =   15134
      _ExtentY        =   820
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------'
'   ���ϸ�  : frmLogin.frm
'   �ۼ���  : ������
'   ��  ��  : ���α׷����� ��
'   �ۼ���  : 2015-04-29
'   ��  ��  : 1.0.0
'-----------------------------------------------------------------------------'

Option Explicit

Private Sub Form_Load()

    Screen.MousePointer = 11
    Xprog.Min = 1
    DoEvents
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Screen.MousePointer = 0
    DoEvents

End Sub
